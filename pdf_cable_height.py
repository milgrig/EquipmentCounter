"""
pdf_cable_height.py — извлечение аннотаций «Гр.X на отм. ...» через PyMuPDF
+ расчёт ОБЩЕЙ ВЫСОТЫ КАБЕЛЕЙ (вертикальных стояков).

Каждая аннотация рядом с цветной трассой означает:
  • трасса физически проходит через N отметок
  • реальная длина = горизонталь × N + вертикальные стояки между крайними отметками

Например, аннотация:
    Гр.1-Гр.8 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200
означает 8 групп трасс, каждая с горизонталью на 6 этажах + стояк 28.2 м.
Полная высота: 8 × 28.2 = 225.6 м только стояков.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import fitz  # PyMuPDF


# ────────────────────────────────────────────────────────────────────────────
# Регулярки
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
    source_pdf: str = ""

    @property
    def n_elevations(self) -> int:
        return len(self.elevations)

    @property
    def n_groups(self) -> int:
        return len(self.groups)

    @property
    def riser_height_m(self) -> float:
        """Высота стояка для одной группы (метры)."""
        vals = []
        for e in self.elevations:
            try:
                vals.append(float(e.replace(",", ".")))
            except ValueError:
                continue
        if len(vals) < 2:
            return 0.0
        return max(vals) - min(vals)

    @property
    def total_riser_length_m(self) -> float:
        """Полный метраж стояков для всех групп."""
        return self.riser_height_m * max(1, self.n_groups)


@dataclass
class CableHeightSummary:
    total_horizontal_m: float = 0.0
    total_vertical_m: float = 0.0
    total_m: float = 0.0
    n_annotations: int = 0
    by_color: dict[str, float] = field(default_factory=dict)
    by_elevation_pair: dict[str, float] = field(default_factory=dict)
    annotations: list[CableAnnotation] = field(default_factory=list)


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
    if r > 180 and g < 90 and b < 90:
        return COLOR_RED
    if b > 180 and r < 90 and g < 130:
        return COLOR_BLUE
    return None


def _collect_text_lines(page) -> list[dict]:
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
                             max_dx: float = 80.0,
                             max_dy: float = 25.0) -> list[dict]:
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
        changed = True
        while changed:
            changed = False
            for j in range(len(s_lines)):
                if used[j]:
                    continue
                M = s_lines[j]
                close = False
                for B in block:
                    dy = abs(M["y0"] - B["y1"])
                    if M["y0"] >= B["y1"]:
                        dy = M["y0"] - B["y1"]
                    if 0 <= dy <= max_dy:
                        bx0, bx1 = B["x0"], B["x1"]
                        mx0, mx1 = M["x0"], M["x1"]
                        if not (mx1 < bx0 - max_dx or mx0 > bx1 + max_dx):
                            close = True
                            break
                if close:
                    block.append(M)
                    used[j] = True
                    changed = True
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


_TITLE_RE = re.compile(
    r"^\s*план\s+(освещения|привязк|кабеленесущ)",
    re.IGNORECASE,
)


def extract_annotations_from_page(page, page_num: int = 1,
                                   source_pdf: str = "") -> list[CableAnnotation]:
    raw = _collect_text_lines(page)
    blocks = _merge_multiline_blocks(raw)
    out: list[CableAnnotation] = []
    for blk in blocks:
        text = blk["text"]
        low = text.lower()
        if "отм" not in low:
            continue
        if _TITLE_RE.search(text):
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
            source_pdf=source_pdf,
        ))
    return out


# ────────────────────────────────────────────────────────────────────────────
# Высокоуровневые помощники
# ────────────────────────────────────────────────────────────────────────────

def collect_page_annotations(pdf_path: str) -> dict[int, list[CableAnnotation]]:
    out: dict[int, list[CableAnnotation]] = {}
    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc, start=1):
            anns = extract_annotations_from_page(page, page_num=i,
                                                  source_pdf=pdf_path)
            if anns:
                out[i] = anns
    finally:
        doc.close()
    return out


def summarize_cable_heights(pdf_paths: list[str],
                              dedupe_global: bool = True) -> CableHeightSummary:
    """Пройти по списку PDF, собрать все аннотации, посчитать высоту кабелей.

    dedupe_global: если True, аннотация с одинаковыми группами и отметками,
    встретившаяся на разных листах (этажах освещения/привязки/лотков),
    учитывается ОДИН раз, поскольку описывает один и тот же физический стояк.
    """
    s = CableHeightSummary()
    seen_keys: set[tuple] = set()
    for p in pdf_paths:
        try:
            anns_by_page = collect_page_annotations(p)
        except Exception:
            continue
        for page_anns in anns_by_page.values():
            for a in page_anns:
                if a.n_elevations < 2:
                    continue
                if dedupe_global:
                    key = (
                        tuple(a.groups),
                        tuple(a.elevations),
                        a.color,
                    )
                    if key in seen_keys:
                        continue
                    seen_keys.add(key)
                s.annotations.append(a)
                s.n_annotations += 1
                v = a.total_riser_length_m
                s.total_vertical_m += v
                if a.color:
                    s.by_color[a.color] = s.by_color.get(a.color, 0.0) + v
                vals = sorted(float(e.replace(",", "."))
                              for e in a.elevations)
                for i in range(1, len(vals)):
                    delta = vals[i] - vals[i - 1]
                    key2 = f"{vals[i-1]:+.3f} → {vals[i]:+.3f}"
                    s.by_elevation_pair[key2] = (s.by_elevation_pair.get(key2, 0.0)
                                                  + delta * max(1, a.n_groups))
    s.total_m = s.total_horizontal_m + s.total_vertical_m
    return s


def format_summary(s: CableHeightSummary) -> str:
    lines = []
    lines.append("Кабельная сводка (вертикальные стояки):")
    lines.append(f"  Уникальных аннотаций: {s.n_annotations}")
    lines.append(f"  Высота кабелей всего: {s.total_vertical_m:.1f} м")
    if s.total_horizontal_m:
        lines.append(f"  Горизонталь:           {s.total_horizontal_m:.1f} м")
        lines.append(f"  Полный метраж:         {s.total_m:.1f} м")
    if s.by_color:
        lines.append("  По цвету:")
        for col, v in sorted(s.by_color.items()):
            lines.append(f"    {col:5s}: {v:.1f} м")
    if s.by_elevation_pair:
        lines.append("  По участкам стояков (отметка → отметка):")
        for key in sorted(s.by_elevation_pair.keys()):
            lines.append(f"    {key:30s}: {s.by_elevation_pair[key]:.1f} м")
    if s.annotations:
        lines.append("  Топ-аннотаций по высоте:")
        top = sorted(s.annotations,
                     key=lambda a: a.total_riser_length_m, reverse=True)[:8]
        for a in top:
            grp_s = (f"{len(a.groups)} гр." if a.groups else "1 трасса")
            lines.append(f"    {a.total_riser_length_m:6.1f} м  "
                         f"({a.riser_height_m:.1f}м × {grp_s}, "
                         f"{a.n_elevations} отм., {a.color or '?'})")
    return "\n".join(lines)
