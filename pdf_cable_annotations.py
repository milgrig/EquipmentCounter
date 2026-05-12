"""
pdf_cable_annotations.py — извлечение аннотаций к кабельным трассам вида:

    Гр.1-Гр.8, Гр.15-Гр.31, Гр.34, Гр.35
    на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200

Каждая такая аннотация рядом с цветной (синей/красной) полилинией означает,
что графически нарисованная трасса физически проходит через N отметок,
поэтому реальная длина кабеля = горизонтальная_длина × N + вертикальные_стояки.

Используется детектором pdf_cables_by_method для коррекции длин.
"""

from __future__ import annotations

import re
import math
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber


# ────────────────────────────────────────────────────────────────────────────
# Регулярки
# ────────────────────────────────────────────────────────────────────────────

# Список отметок: "0.000, +9.000, +13.800, +18.600, +23.400, +28.200"
ELEV_LIST_RE = re.compile(
    r"на\s+отм\.?\s*"
    r"((?:[+-]?\d+[.,]\d+\s*[,;\s]+\s*)*[+-]?\d+[.,]\d+)",
    re.IGNORECASE,
)

# Одна отметка: "0.000", "+4.200", "-1.500"
ELEV_RE = re.compile(r"[+-]?\d+[.,]\d+")

# Группа: "Гр.1", "Гр.34А", "Гр.1-Гр.8" (диапазон)
GROUP_TOKEN_RE = re.compile(
    r"Гр\.?\s*(\d+)([А-Яа-я]?)"
    r"(?:\s*[-–—]\s*Гр\.?\s*(\d+)([А-Яа-я]?))?",
    re.IGNORECASE,
)

# Цвет для классификации — соответствует pdf_cables_by_method
COLOR_RED = "red"
COLOR_BLUE = "blue"


@dataclass
class CableAnnotation:
    """Распознанная аннотация рядом с трассой."""
    text: str                       # исходный текст
    x: float                        # центр текста, x (top-left coord)
    y: float                        # центр текста, y
    elevations: list[str]           # ["0.000", "+9.000", ...]
    groups: list[str]               # ["Гр.1", "Гр.2", ..., "Гр.8", "Гр.34"]
    color: Optional[str] = None     # red/blue по цвету шрифта (если удалось)


@dataclass
class AnnotationMatch:
    """Привязка аннотации к ближайшей полилинии."""
    annotation: CableAnnotation
    polyline_idx: int               # индекс полилинии в исходном списке
    distance_pt: float              # расстояние от центра аннотации до линии
    multiplier: int                 # = len(elevations)
    n_groups: int                   # число групп (для информации)


# ────────────────────────────────────────────────────────────────────────────
# Парсинг текста
# ────────────────────────────────────────────────────────────────────────────

def _normalize_elev(raw: str) -> str:
    """'+4.200' / '4,200' / '0.000' → нормализованный вид."""
    raw = raw.replace(",", ".").strip()
    try:
        v = float(raw)
    except ValueError:
        return raw
    if abs(v) < 1e-6:
        return "0.000"
    return f"{v:+.3f}"


def parse_elevations(text: str) -> list[str]:
    """Из текста "на отм. 0.000, +9.000, ..." вернуть нормализованный список."""
    m = ELEV_LIST_RE.search(text)
    if not m:
        return []
    chunk = m.group(1)
    out: list[str] = []
    seen: set[str] = set()
    for em in ELEV_RE.finditer(chunk):
        norm = _normalize_elev(em.group(0))
        if norm not in seen:
            seen.add(norm)
            out.append(norm)
    return out


def parse_groups(text: str) -> list[str]:
    """Из 'Гр.1-Гр.8, Гр.15-Гр.31, Гр.34' → ['Гр.1', ..., 'Гр.8', 'Гр.15', ..., 'Гр.31', 'Гр.34']."""
    out: list[str] = []
    for m in GROUP_TOKEN_RE.finditer(text):
        n1 = int(m.group(1))
        suf1 = (m.group(2) or "").upper()
        n2 = m.group(3)
        suf2 = (m.group(4) or "").upper()
        if n2 is None:
            out.append(f"Гр.{n1}{suf1}")
        else:
            n2 = int(n2)
            # диапазон Гр.1-Гр.8 (с одинаковыми суффиксами)
            sfx = suf1 if suf1 else suf2
            for k in range(min(n1, n2), max(n1, n2) + 1):
                out.append(f"Гр.{k}{sfx}")
    # дедуп с сохранением порядка
    seen: set[str] = set()
    uniq: list[str] = []
    for g in out:
        if g not in seen:
            seen.add(g)
            uniq.append(g)
    return uniq


# ────────────────────────────────────────────────────────────────────────────
# Извлечение аннотаций со страницы PDF
# ────────────────────────────────────────────────────────────────────────────

def _word_color(word: dict) -> Optional[str]:
    """Определить red/blue по non-stroking color слова pdfplumber."""
    nsc = word.get("non_stroking_color") or word.get("stroking_color")
    if nsc is None:
        return None
    if isinstance(nsc, (int, float)):
        return None
    if isinstance(nsc, (tuple, list)):
        if len(nsc) >= 3:
            r, g, b = float(nsc[0]), float(nsc[1]), float(nsc[2])
            if r > 0.7 and g < 0.3 and b < 0.3:
                return COLOR_RED
            if b > 0.7 and r < 0.3 and g < 0.3:
                return COLOR_BLUE
    return None


def _cluster_words_into_blocks(
    words: list[dict],
    max_dx: float = 80.0,
    max_dy: float = 16.0,
) -> list[list[dict]]:
    """Слепить слова в текстовые блоки (надписи) по близости.
    Слова сортируются сверху-вниз, слева-направо; смежные сливаются."""
    if not words:
        return []
    sw = sorted(words, key=lambda w: (round(w["top"] / max_dy), w["x0"]))

    blocks: list[list[dict]] = []
    for w in sw:
        placed = False
        for blk in blocks:
            # Проверяем близость к любому слову блока
            for bw in blk:
                dx = min(abs(w["x0"] - bw["x1"]),
                         abs(w["x1"] - bw["x0"]),
                         abs((w["x0"] + w["x1"]) / 2 - (bw["x0"] + bw["x1"]) / 2))
                dy = abs((w["top"] + w["bottom"]) / 2 - (bw["top"] + bw["bottom"]) / 2)
                if dx < max_dx and dy < max_dy:
                    blk.append(w)
                    placed = True
                    break
            if placed:
                break
        if not placed:
            blocks.append([w])
    return blocks


def extract_annotations_from_page(page) -> list[CableAnnotation]:
    """Извлечь все аннотации '... на отм. ...' со страницы pdfplumber.Page."""
    try:
        words = page.extract_words(extra_attrs=["non_stroking_color", "stroking_color"])
    except Exception:
        return []
    if not words:
        return []

    blocks = _cluster_words_into_blocks(words)
    out: list[CableAnnotation] = []
    for blk in blocks:
        text = " ".join(w["text"] for w in blk)
        if "отм" not in text.lower():
            continue
        elevs = parse_elevations(text)
        if not elevs:
            continue
        groups = parse_groups(text)
        # Геометрический центр блока
        x0 = min(w["x0"] for w in blk)
        x1 = max(w["x1"] for w in blk)
        y0 = min(w["top"] for w in blk)
        y1 = max(w["bottom"] for w in blk)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2

        # Доминирующий цвет шрифта в блоке
        color_votes = {}
        for w in blk:
            c = _word_color(w)
            if c:
                color_votes[c] = color_votes.get(c, 0) + 1
        color = None
        if color_votes:
            color = max(color_votes, key=color_votes.get)

        out.append(CableAnnotation(
            text=text, x=cx, y=cy,
            elevations=elevs, groups=groups, color=color,
        ))
    return out


# ────────────────────────────────────────────────────────────────────────────
# Привязка аннотаций к полилиниям
# ────────────────────────────────────────────────────────────────────────────

def _point_to_segment_dist(px: float, py: float,
                            ax: float, ay: float,
                            bx: float, by: float) -> float:
    """Минимальное расстояние от точки P до отрезка AB."""
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    if L2 < 1e-9:
        return math.hypot(px - ax, py - ay)
    t = ((px - ax) * dx + (py - ay) * dy) / L2
    t = max(0.0, min(1.0, t))
    fx, fy = ax + t * dx, ay + t * dy
    return math.hypot(px - fx, py - fy)


def match_annotations_to_polylines(
    annotations: list[CableAnnotation],
    polylines: list,            # объекты с .segments=[(x1,y1,x2,y2),...] и .color
    *,
    max_distance_pt: float = 80.0,
    require_color_match: bool = True,
) -> list[AnnotationMatch]:
    """Привязать каждую аннотацию к ближайшей полилинии того же цвета."""
    matches: list[AnnotationMatch] = []
    for ann in annotations:
        best_idx = -1
        best_dist = float("inf")
        for i, pl in enumerate(polylines):
            pl_color = getattr(pl, "color", None)
            if require_color_match and ann.color and pl_color and ann.color != pl_color:
                continue
            segs = getattr(pl, "segments", None) or []
            for s in segs:
                if len(s) < 4:
                    continue
                ax, ay, bx, by = s[0], s[1], s[2], s[3]
                d = _point_to_segment_dist(ann.x, ann.y, ax, ay, bx, by)
                if d < best_dist:
                    best_dist = d
                    best_idx = i
        if best_idx >= 0 and best_dist <= max_distance_pt:
            matches.append(AnnotationMatch(
                annotation=ann,
                polyline_idx=best_idx,
                distance_pt=best_dist,
                multiplier=max(1, len(ann.elevations)),
                n_groups=max(1, len(ann.groups)),
            ))
    return matches


# ────────────────────────────────────────────────────────────────────────────
# Высокоуровневый помощник
# ────────────────────────────────────────────────────────────────────────────

def collect_page_annotations(pdf_path: str) -> dict[int, list[CableAnnotation]]:
    """Извлечь аннотации со всех страниц, вернуть {page_num_1based: [ann, ...]}."""
    out: dict[int, list[CableAnnotation]] = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            anns = extract_annotations_from_page(page)
            if anns:
                out[i] = anns
    return out


def vertical_riser_length_m(
    elevations: list[str],
    *,
    n_groups: int = 1,
) -> float:
    """Длина вертикальных стояков между N отметками для n_groups трасс.

    Между отметками 0.000, +4.200, +9.000 → 9.000 м перепада на группу.
    Возвращает суммарный метраж стояков для всех групп.
    """
    vals: list[float] = []
    for e in elevations:
        try:
            vals.append(float(e.replace(",", ".")))
        except ValueError:
            continue
    if len(vals) < 2:
        return 0.0
    vals.sort()
    return (vals[-1] - vals[0]) * max(1, n_groups)
