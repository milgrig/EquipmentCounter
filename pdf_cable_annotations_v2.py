"""
pdf_cable_annotations_v2.py — извлечение аннотаций к кабельным трассам через fitz.

Аннотация:
    Гр.1-Гр.8, Гр.15-Гр.31, Гр.34, Гр.35
    на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200

Каждая такая аннотация рядом с цветной полилинией означает, что физически
трасса проходит через N отметок, поэтому реальная длина =
горизонтальная_длина × N (на каждой отметке) + вертикальные_стояки.

Текстовый слой берётся через PyMuPDF (fitz) — pdfplumber на CAD-PDF
часто разбивает строки на отдельные слова.
"""

from __future__ import annotations

import re
import math
from dataclasses import dataclass
from typing import Optional

import fitz  # PyMuPDF


# ────────────────────────────────────────────────────────────────────────────
# Регулярки парсинга
# ────────────────────────────────────────────────────────────────────────────

ELEV_LIST_RE = re.compile(
    r"на\s+отм\.?\s*"
    r"((?:[+-]?\d+[.,]\d+\s*[,;\s]+\s*)*[+-]?\d+[.,]\d+)",
    re.IGNORECASE,
)
ELEV_RE = re.compile(r"[+-]?\d+[.,]\d+")
GROUP_TOKEN_RE = re.compile(
    r"Гр\.?\s*(\d+)([А-Яа-я]?)"
    r"(?:\s*[-–—]\s*Гр\.?\s*(\d+)([А-Яа-я]?))?",
    re.IGNORECASE,
)

COLOR_RED = "red"
COLOR_BLUE = "blue"


@dataclass
class CableAnnotation:
    text: str
    x: float
    y: float
    page_num: int
    elevations: list[str]
    groups: list[str]
    color: Optional[str] = None


@dataclass
class AnnotationMatch:
    annotation: CableAnnotation
    polyline_idx: int
    distance_pt: float
    multiplier: int
    n_groups: int


# ────────────────────────────────────────────────────────────────────────────
# Парсинг текста
# ────────────────────────────────────────────────────────────────────────────

def _normalize_elev(raw: str) -> str:
    raw = raw.replace(",", ".").strip()
    try:
        v = float(raw)
    except ValueError:
        return raw
    if abs(v) < 1e-6:
        return "0.000"
    return f"{v:+.3f}"


def parse_elevations(text: str) -> list[str]:
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
            sfx = suf1 if suf1 else suf2
            for k in range(min(n1, n2), max(n1, n2) + 1):
                out.append(f"Гр.{k}{sfx}")
    seen: set[str] = set()
    uniq: list[str] = []
    for g in out:
        if g not in seen:
            seen.add(g)
            uniq.append(g)
    return uniq


# ────────────────────────────────────────────────────────────────────────────
# Извлечение строк через fitz
# ────────────────────────────────────────────────────────────────────────────

def _color_from_int(rgb_int: int) -> Optional[str]:
    r = (rgb_int >> 16) & 0xFF
    g = (rgb_int >> 8) & 0xFF
    b = rgb_int & 0xFF
    if r > 180 and g < 80 and b < 80:
        return COLOR_RED
    if b > 180 and r < 80 and g < 80:
        return COLOR_BLUE
    return None


def _collect_text_lines(page) -> list[dict]:
    """Собрать text-lines с bbox и доминирующим цветом."""
    raw_lines: list[dict] = []
    d = page.get_text("dict")
    for blk in d.get("blocks", []):
        if "lines" not in blk:
            continue
        for line in blk["lines"]:
            spans = line.get("spans", [])
            if not spans:
                continue
            full_text = "".join(s.get("text", "") for s in spans).strip()
            if not full_text:
                continue
            x0, y0, x1, y1 = line.get("bbox", (0, 0, 0, 0))
            color_votes: dict[str, int] = {}
            for s in spans:
                c = _color_from_int(int(s.get("color", 0)))
                if c:
                    color_votes[c] = color_votes.get(c, 0) + len(s.get("text", ""))
            color = max(color_votes, key=color_votes.get) if color_votes else None
            raw_lines.append({
                "text": full_text,
                "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                "color": color,
            })
    return raw_lines


def _merge_multiline_blocks(lines: list[dict],
                             max_dx: float = 60.0,
                             max_dy: float = 22.0) -> list[dict]:
    """Объединить близко расположенные строки в один блок."""
    if not lines:
        return []
    s_lines = sorted(lines, key=lambda L: (L["y0"], L["x0"]))
    used = [False] * len(s_lines)
    out: list[dict] = []
    for i, L in enumerate(s_lines):
        if used[i]:
            continue
        block = [L]
        used[i] = True
        for j in range(i + 1, len(s_lines)):
            if used[j]:
                continue
            M = s_lines[j]
            close = False
            for B in block:
                dy = M["y0"] - B["y1"]
                if dy < 0:
                    dy = abs(M["y0"] - B["y0"])
                if 0 <= dy <= max_dy:
                    bx0, bx1 = B["x0"], B["x1"]
                    mx0, mx1 = M["x0"], M["x1"]
                    if not (mx1 < bx0 - max_dx or mx0 > bx1 + max_dx):
                        close = True
                        break
            if close:
                block.append(M)
                used[j] = True
        text = " ".join(b["text"] for b in block)
        x0 = min(b["x0"] for b in block)
        x1 = max(b["x1"] for b in block)
        y0 = min(b["y0"] for b in block)
        y1 = max(b["y1"] for b in block)
        votes: dict[str, int] = {}
        for b in block:
            if b["color"]:
                votes[b["color"]] = votes.get(b["color"], 0) + len(b["text"])
        col = max(votes, key=votes.get) if votes else None
        out.append({"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1, "color": col})
    return out


def extract_annotations_from_page(page, page_num: int = 1) -> list[CableAnnotation]:
    """Найти на странице все аннотации '... на отм. ...'."""
    raw = _collect_text_lines(page)
    blocks = _merge_multiline_blocks(raw)
    out: list[CableAnnotation] = []
    for blk in blocks:
        text = blk["text"]
        low = text.lower()
        if "отм" not in low:
            continue
        # фильтр: рамки/штампы вида "План освещения на отм. 0.000" с одной отметкой
        # — это не аннотация трассы, а заголовок листа
        if low.startswith("план") and len(text) < 60:
            elevs_count = len(parse_elevations(text))
            if elevs_count <= 2:
                continue
        elevs = parse_elevations(text)
        if not elevs:
            continue
        groups = parse_groups(text)
        cx = (blk["x0"] + blk["x1"]) / 2
        cy = (blk["y0"] + blk["y1"]) / 2
        out.append(CableAnnotation(
            text=text, x=cx, y=cy, page_num=page_num,
            elevations=elevs, groups=groups, color=blk["color"],
        ))
    return out


# ────────────────────────────────────────────────────────────────────────────
# Геометрия
# ────────────────────────────────────────────────────────────────────────────

def _point_to_segment_dist(px: float, py: float,
                            ax: float, ay: float,
                            bx: float, by: float) -> float:
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
    polylines: list,
    *,
    max_distance_pt: float = 100.0,
    require_color_match: bool = True,
) -> list[AnnotationMatch]:
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
                d = _point_to_segment_dist(ann.x, ann.y, s[0], s[1], s[2], s[3])
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
# Высокоуровневое API
# ────────────────────────────────────────────────────────────────────────────

def collect_page_annotations(pdf_path: str) -> dict[int, list[CableAnnotation]]:
    out: dict[int, list[CableAnnotation]] = {}
    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc, start=1):
            anns = extract_annotations_from_page(page, page_num=i)
            if anns:
                out[i] = anns
    finally:
        doc.close()
    return out


def vertical_riser_length_m(elevations: list[str], *, n_groups: int = 1) -> float:
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


def page_summary(annotations: list[CableAnnotation]) -> dict:
    """Сводка по странице — какой средний/макс мультипликатор, сколько групп всего."""
    if not annotations:
        return {"avg_multiplier": 1.0, "max_multiplier": 1, "n_anns": 0,
                "total_groups": 0, "elevations_used": []}
    mults = [max(1, len(a.elevations)) for a in annotations]
    grps = [max(1, len(a.groups)) for a in annotations]
    all_elevs: set[str] = set()
    for a in annotations:
        all_elevs.update(a.elevations)
    return {
        "avg_multiplier": sum(mults) / len(mults),
        "max_multiplier": max(mults),
        "total_groups": sum(grps),
        "n_anns": len(annotations),
        "elevations_used": sorted(all_elevs),
    }
