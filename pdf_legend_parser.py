"""
pdf_legend_parser.py — Core legend extraction module.

Detects, parses, and extracts legend (условные обозначения) tables from
Russian electrical engineering PDF drawings using pdfplumber.

Uses PDF line/rect objects for table boundary detection instead of
hardcoded offsets. Captures FULL descriptions including model names
and parameters. Handles multi-line descriptions and items without symbols.

Usage:
    python pdf_legend_parser.py <path.pdf>
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class LegendItem:
    """One row from the legend table."""
    symbol: str            # e.g. "1", "1А", "3АЭ", "ВЫХОД", "" for non-symbol items
    description: str       # FULL text: model names, parameters, etc.
    category: str = ""     # auto-detected: "светильник", "выключатель", "кабельная трасса", etc.
    bbox: tuple[float, float, float, float] = (0, 0, 0, 0)  # (x0, y0, x1, y1)
    color: str = ""        # detected dominant color: "red", "blue", "black", "grey", or ""


@dataclass
class LegendResult:
    """Complete parsing result for one legend table."""
    items: list[LegendItem] = field(default_factory=list)
    legend_bbox: tuple[float, float, float, float] = (0, 0, 0, 0)  # overall table bbox
    page_index: int = 0
    columns_detected: int = 1  # 1 = single column, 2 = two-column (GPC-style)


# ---------------------------------------------------------------------------
# Constants / Regex
# ---------------------------------------------------------------------------

# Legend header patterns (case-insensitive search)
LEGEND_HEADER_PATTERNS = [
    re.compile(r"Условные", re.IGNORECASE),
    re.compile(r"Легенда", re.IGNORECASE),
    re.compile(r"Обозначени", re.IGNORECASE),       # "Обозначения" or truncated
    re.compile(r"Спецификац", re.IGNORECASE),        # "Спецификация"
    re.compile(r"Перечень", re.IGNORECASE),           # "Перечень оборудования"
    re.compile(r"Экспликац", re.IGNORECASE),          # "Экспликация"
]

# Symbol regex — 1-2 digits + optional up to 3 Cyrillic letters (1, 1А, 3АЭ, 10А, ...)
SYMBOL_RE = re.compile(r"^\d{1,2}[А-Яа-яЁё]{0,3}$")

# Text-based symbol identifiers that are NOT numeric (e.g. "ВЫХОД")
TEXT_SYMBOL_RE = re.compile(r"^[A-ZА-ЯЁ]{3,}$")

# Emergency circuit variant (e.g. 2А → base 2 + "А")
CIRCUIT_VARIANT_RE = re.compile(r"^(\d+)(А|АЭ)$")

# Column header keywords
COL_HEADERS = {"Обозначение", "Наименование", "Примечание"}

# Category detection keywords (checked in order — first match wins)
CATEGORY_KEYWORDS = {
    "эвакуационный": ["эвакуационный", "ВЫХОД"],
    "светильник": ["Светильник", "светодиодный"],
    "выключатель": ["Выключатель"],
    "розетка": ["Розетка"],
    "щит": ["Щит"],
    "кабельная трасса": ["Кабельная трасса"],
    "проводка": ["Проводка"],
}

# Minimum length for a horizontal/vertical line to be considered a table border
# Lowered from 80 to 50: small legend tables with few rows have short vertical
# lines (e.g. 78pt in pos_8_2_em). 50pt still distinguishes from tick marks.
MIN_TABLE_LINE_LEN = 50

# Title block keywords that indicate we hit the title block, not legend
TITLE_BLOCK_KEYWORDS = {"Лист", "док.", "Подп.", "Дата", "Разраб.", "Пров.",
                         "Н.контр.", "ГИП", "Изм.", "Стадия"}

# Additional title block patterns — catches concatenated text like "ЛистДатаПодп."
# where pdfplumber merges adjacent title block words without spaces.
TITLE_BLOCK_CONCAT_PATTERNS = [
    "Формат", "ЛистДата", "Подп.", "N dok", "Kol.uch", "Izm.",
    "Нач.отд", "ГИП", "Разраб.", "Пров.", "Н.контр",
]

# Construction note phrases that are NOT equipment names.
# These appear when note text from the drawing leaks into descriptions.
# Ordered longest-first so that more specific phrases match before
# shorter sub-phrases (e.g. "Монтаж крайних" before "крайних на линии").
CONSTRUCTION_NOTE_PHRASES = [
    "Коробки распаячные установить",
    "Монтаж крайних на линии",
    "Монтаж крайних",
    "крайних на линии",
    "Расчет количества",
    "Для соединения",
    "установить на",
    "согласно проекту",
]

# Maximum sane description length — beyond this we look for concatenation issues
MAX_DESCRIPTION_LEN = 150


# ---------------------------------------------------------------------------
# Color detection helpers
# ---------------------------------------------------------------------------

def _normalize_color(c) -> Optional[tuple]:
    """Normalize a pdfplumber color value to an RGB tuple of floats, or None."""
    if c is None:
        return None
    if isinstance(c, (int, float)):
        v = float(c)
        return (v, v, v)
    if isinstance(c, (tuple, list)):
        if len(c) == 3:
            return tuple(round(float(x), 4) for x in c)
        if len(c) == 4:
            # CMYK -> RGB approximation
            cc, m, y, k = [float(x) for x in c]
            r = (1 - cc) * (1 - k)
            g = (1 - m) * (1 - k)
            b = (1 - y) * (1 - k)
            return (round(r, 4), round(g, 4), round(b, 4))
        if len(c) == 1:
            v = float(c[0])
            return (v, v, v)
    return None


def _classify_rgb(rgb: tuple) -> str:
    """Classify an RGB tuple (0-1 floats) into a color label.

    Returns "red", "blue", "grey", "black", or "".
    """
    if rgb is None:
        return ""
    r, g, b = rgb[0], rgb[1], rgb[2]

    # Black: all channels near 0
    if r < 0.15 and g < 0.15 and b < 0.15:
        return "black"

    # Grey: all channels similar, mid-range
    if abs(r - g) < 0.1 and abs(g - b) < 0.1 and abs(r - b) < 0.1:
        if r > 0.15 and r < 0.85:
            return "grey"
        if r >= 0.85:
            return ""  # white / near-white, skip
        return "black"

    # Red: high R, low G and B
    if r > 0.5 and g < 0.3 and b < 0.3:
        return "red"

    # Blue: high B, low R and G
    if b > 0.5 and r < 0.3 and g < 0.3:
        return "blue"

    # Less strict red (e.g. (0.8, 0.1, 0.1))
    if r > 0.4 and r > g * 2 and r > b * 2:
        return "red"

    # Less strict blue (e.g. (0.1, 0.1, 0.8))
    if b > 0.4 and b > r * 2 and b > g * 2:
        return "blue"

    return ""


def _detect_row_color(
    lines: list[dict],
    rects: list[dict],
    row_bbox: tuple[float, float, float, float],
    sym_x_boundary: float,
) -> str:
    """Detect the dominant color of graphical elements in the symbol cell area.

    Looks at stroking_color of lines and rects that fall within the symbol
    column of this row (left of sym_x_boundary).

    Returns "red", "blue", "grey", "black", or "".
    """
    x0, y0, x1, y1 = row_bbox
    # Focus on symbol column area (left part of the row)
    sym_x1 = min(x1, sym_x_boundary + 10)

    color_counts: dict[str, int] = {}

    for ln in lines:
        # Check if line is within the symbol cell area
        lx0 = min(ln["x0"], ln["x1"])
        lx1 = max(ln["x0"], ln["x1"])
        ly0 = min(ln["top"], ln["bottom"])
        ly1 = max(ln["top"], ln["bottom"])

        # Line must overlap with the row's symbol cell
        if lx1 < x0 - 5 or lx0 > sym_x1 + 5:
            continue
        if ly1 < y0 - 3 or ly0 > y1 + 3:
            continue

        # Skip very long lines (table borders)
        line_len = max(abs(ln["x1"] - ln["x0"]), abs(ln["bottom"] - ln["top"]))
        if line_len > (y1 - y0) * 2 and line_len > 100:
            continue

        rgb = _normalize_color(ln.get("stroking_color"))
        if rgb is None:
            continue
        label = _classify_rgb(rgb)
        if label and label != "black":
            color_counts[label] = color_counts.get(label, 0) + 1

    for rect in rects:
        rx0 = rect["x0"]
        ry0 = rect["top"]
        rx1 = rect["x1"]
        ry1 = rect["bottom"]

        if rx1 < x0 - 5 or rx0 > sym_x1 + 5:
            continue
        if ry1 < y0 - 3 or ry0 > y1 + 3:
            continue

        # Skip large rects (table cell borders)
        rect_w = rx1 - rx0
        rect_h = ry1 - ry0
        if rect_w > sym_x1 - x0 + 20 and rect_h > y1 - y0 + 10:
            continue

        for color_key in ("stroking_color", "non_stroking_color"):
            rgb = _normalize_color(rect.get(color_key))
            if rgb is None:
                continue
            label = _classify_rgb(rgb)
            if label and label != "black":
                color_counts[label] = color_counts.get(label, 0) + 1

    if not color_counts:
        return ""

    # Return the most common non-black color
    best = max(color_counts, key=color_counts.get)  # type: ignore[arg-type]
    return best


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _y_group(wlist: list[dict], tol: float = 5) -> list[tuple[float, list[dict]]]:
    """Group words by Y coordinate (top) with tolerance."""
    if not wlist:
        return []
    wlist = sorted(wlist, key=lambda w: (w["top"], w["x0"]))
    groups: list[tuple[float, list[dict]]] = []
    cur: list[dict] = [wlist[0]]
    for w in wlist[1:]:
        if abs(w["top"] - cur[0]["top"]) <= tol:
            cur.append(w)
        else:
            groups.append((cur[0]["top"], cur))
            cur = [w]
    groups.append((cur[0]["top"], cur))
    return groups


def _detect_category(text: str) -> str:
    """Auto-detect category from description text."""
    for category, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in text.lower():
                return category
    return ""


def _join_words(words: list[dict]) -> str:
    """Join word dicts into a single string, sorted by x0."""
    return " ".join(
        w["text"] for w in sorted(words, key=lambda w: w["x0"])
    ).strip()


def _is_title_block_text(text: str) -> bool:
    """Check if text looks like title block content (not legend data).

    Checks both space-separated keywords and substring patterns to catch
    concatenated title block text like 'ЛистДатаПодп.' that pdfplumber
    sometimes produces when adjacent words are merged.
    """
    words_set = set(text.split())
    if words_set & TITLE_BLOCK_KEYWORDS:
        return True
    # Date pattern (dd.mm.yy or dd.mm.yyyy)
    if re.search(r"\d{2}\.\d{2}\.\d{2,4}", text):
        return True
    # Concatenated title block patterns (substring match)
    for pattern in TITLE_BLOCK_CONCAT_PATTERNS:
        if pattern in text:
            return True
    return False


def _strip_construction_notes(text: str) -> str:
    """Remove construction note phrases that leak into equipment descriptions.

    These phrases come from drawing notes, not from the legend table itself.
    If the ENTIRE text is a construction note, returns empty string.
    Otherwise strips the note portion.

    After stripping, if the remaining text is too short (<5 chars) to be
    a meaningful equipment description, returns empty string.
    """
    text_lower = text.lower()
    for phrase in CONSTRUCTION_NOTE_PHRASES:
        idx = text_lower.find(phrase.lower())
        if idx != -1:
            # If phrase starts at beginning, the whole text may be a note
            if idx == 0:
                return ""
            # Trim the note portion from the description
            text = text[:idx].rstrip(" .,;:")
            text_lower = text.lower()  # update for next iteration
    # After stripping, if remaining text is too short, treat as garbage
    if text and len(text.strip()) < 5:
        return ""
    return text


def _split_on_duplicate_words(text: str) -> str:
    """Split on repeated Cyrillic words (>=5 chars).

    If the same word appears twice, the second occurrence likely starts
    a new concatenated equipment name.  Return text up to (but not
    including) the second occurrence.
    """
    cyrillic_words = re.findall(r'[А-ЯЁа-яё]{5,}', text)
    seen: dict[str, int] = {}
    for w in cyrillic_words:
        wl = w.lower()
        if wl in seen:
            idx = text.lower().find(wl, seen[wl] + 1)
            if idx > 10:
                first_part = text[:idx].rstrip(" .,;:")
                if len(first_part) >= 10:
                    return first_part
            break
        seen[wl] = text.lower().find(wl)
    return text


def _sanitize_long_description(text: str) -> str:
    """Check if a long description (>MAX_DESCRIPTION_LEN) contains concatenated
    equipment names and try to extract just the first one.

    Common delimiters that signal concatenation:
    - Multiple model numbers (e.g. 'OPTIMA.OPL 236 HF ... Светильник ...')
    - Sentence-ending punctuation followed by new sentence
    - Repeated equipment name keywords (e.g. 'Консоль ... Консоль ...')

    Returns cleaned description.
    """
    if len(text) <= MAX_DESCRIPTION_LEN:
        return text

    # Try splitting at sentence boundaries — period followed by space and
    # uppercase letter (Cyrillic or Latin) signals a new sentence / item.
    parts = re.split(r'(?<=\.)\s+(?=[A-ZА-ЯЁ])', text)
    if len(parts) > 1:
        # Keep only the first meaningful part
        text = parts[0].strip()
        if len(text) <= MAX_DESCRIPTION_LEN:
            # Still try duplicate-word split (first sentence part may
            # itself contain concatenated items)
            text = _split_on_duplicate_words(text)
            return text

    # Try splitting on repeated equipment-name keywords.
    text = _split_on_duplicate_words(text)
    if len(text) <= MAX_DESCRIPTION_LEN:
        return text

    # Truncate at MAX_DESCRIPTION_LEN as last resort, break at word boundary
    truncated = text[:MAX_DESCRIPTION_LEN]
    last_space = truncated.rfind(" ")
    if last_space > MAX_DESCRIPTION_LEN // 2:
        truncated = truncated[:last_space]
    return truncated.rstrip(" .,;:")


# ---------------------------------------------------------------------------
# Table boundary detection using PDF lines/rects
# ---------------------------------------------------------------------------

@dataclass
class TableBounds:
    """Detected table boundaries from PDF line objects."""
    x0: float          # left edge
    x1: float          # right edge
    y0: float          # top edge
    y1: float          # bottom edge
    col_xs: list[float]  # internal vertical column dividers (sorted)
    row_ys: list[float]  # horizontal row divider Y positions (sorted)


def _find_table_bounds(
    lines: list[dict],
    header_y: float,
    header_x: float,
    page_width: float,
) -> Optional[TableBounds]:
    """
    Detect legend table boundaries using PDF line objects near the legend header.

    Strategy:
    1. Find long horizontal lines near/below the header Y.
    2. Filter out title block (shtamp) lines that have different X positions.
    3. Find long vertical lines that span the same Y range.
    4. Determine table extents (x0, y0, x1, y1) and internal column dividers.

    Title block filtering: After initial candidate selection, only keep
    horizontal lines whose X range overlaps with the header position +/-200pt.
    Title block lines are usually at completely different X positions.
    """
    # Gather horizontal lines (nearly horizontal: dy < 2, length > MIN)
    h_lines = [
        l for l in lines
        if abs(l["top"] - l["bottom"]) < 2
        and abs(l["x1"] - l["x0"]) > MIN_TABLE_LINE_LEN
    ]

    # Gather vertical lines (nearly vertical: dx < 2, length > MIN)
    v_lines = [
        l for l in lines
        if abs(l["x0"] - l["x1"]) < 2
        and abs(l["bottom"] - l["top"]) > MIN_TABLE_LINE_LEN
    ]

    # Filter horizontal lines near/below the legend header, within
    # a reasonable X range (header_x +/- generous bounds for left side)
    search_x_min = header_x - 300
    search_x_max = header_x + 800
    search_y_min = header_y - 30
    search_y_max = header_y + 1000  # generous: legend can be tall

    candidate_h = [
        l for l in h_lines
        if search_y_min < l["top"] < search_y_max
        and l["x0"] < search_x_max
        and l["x1"] > search_x_min
    ]

    if not candidate_h:
        return None

    # --- Title block line filtering ---
    # Only keep lines whose X range overlaps with the header X +/-200pt.
    # Title block lines at far-away X positions are excluded.
    overlap_margin = 200
    header_x_min = header_x - overlap_margin
    header_x_max = header_x + overlap_margin
    filtered_h = [
        l for l in candidate_h
        if l["x1"] > header_x_min and l["x0"] < header_x_max
    ]
    # If filtering removed everything, fall back to unfiltered
    if filtered_h:
        candidate_h = filtered_h

    # --- Group lines by (x0, x1) to identify distinct tables ---
    # Multiple tables may be nearby (legend table + title block).  We need
    # the legend table, which starts closest below the header, not just the
    # table with the most lines.
    _XTOL = 5  # tolerance for grouping x0/x1 values

    # Build groups: key = (rounded_x0, rounded_x1) -> list of lines
    line_groups: dict[tuple[float, float], list[dict]] = {}
    for l in candidate_h:
        rx0 = round(l["x0"], 0)
        rx1 = round(l["x1"], 0)
        matched_key = None
        for gk in line_groups:
            if abs(gk[0] - rx0) < _XTOL and abs(gk[1] - rx1) < _XTOL:
                matched_key = gk
                break
        if matched_key is None:
            matched_key = (rx0, rx1)
        line_groups.setdefault(matched_key, []).append(l)

    # Score each group: prefer the one whose first line is closest to the
    # header.  Require at least 2 lines and reasonable width (40..800 pt).
    best_group_key = None
    best_score = float("inf")

    for gk, g_lines in line_groups.items():
        if len(g_lines) < 2:
            continue
        gx0, gx1 = gk
        width = gx1 - gx0
        if width < 40 or width > 800:
            continue
        min_dist = min(abs(l["top"] - header_y) for l in g_lines)
        if min_dist < best_score:
            best_score = min_dist
            best_group_key = gk

    if best_group_key is None:
        # Fallback: pick the most-common x0/x1 (old behaviour)
        x0_counts: dict[float, int] = {}
        x1_counts: dict[float, int] = {}
        for l in candidate_h:
            rx0 = round(l["x0"], 0)
            rx1 = round(l["x1"], 0)
            x0_counts[rx0] = x0_counts.get(rx0, 0) + 1
            x1_counts[rx1] = x1_counts.get(rx1, 0) + 1
        table_x0 = max(x0_counts, key=x0_counts.get)  # type: ignore[arg-type]
        table_x1 = max(x1_counts, key=x1_counts.get)  # type: ignore[arg-type]
    else:
        table_x0, table_x1 = best_group_key

    # Filter to lines matching these x bounds (tolerance +/-5)
    matched_h = [
        l for l in candidate_h
        if abs(round(l["x0"], 0) - table_x0) < _XTOL
        and abs(round(l["x1"], 0) - table_x1) < _XTOL
    ]

    if not matched_h:
        return None

    row_ys = sorted(set(round(l["top"], 1) for l in matched_h))
    y0 = min(row_ys)
    y1 = max(row_ys)

    # Find vertical lines within the table X/Y range
    matched_v = [
        l for l in v_lines
        if table_x0 - 5 < l["x0"] < table_x1 + 5
        and l["top"] < y0 + 20
        and l["bottom"] > y1 - 20
    ]

    col_xs = sorted(set(round(l["x0"], 1) for l in matched_v))

    return TableBounds(
        x0=table_x0,
        x1=table_x1,
        y0=y0,
        y1=y1,
        col_xs=col_xs,
        row_ys=row_ys,
    )


# ---------------------------------------------------------------------------
# Second-column (GPC-style) detection
# ---------------------------------------------------------------------------

def _find_second_table(
    lines: list[dict],
    words: list[dict],
    first_bounds: TableBounds,
    page_width: float,
) -> Optional[TableBounds]:
    """
    Look for a second legend table to the right of the first one.
    GPC-style drawings often have a second table with different items.
    """
    # Look for column headers to the right of the first table
    right_headers = [
        w for w in words
        if w["x0"] > first_bounds.x1 + 50
        and w["text"] in ("Наименование", "Примечание")
        and abs(w["top"] - first_bounds.y0) < 50
    ]

    if not right_headers:
        return None

    # Use the rightmost header area as a starting point
    header_x = min(w["x0"] for w in right_headers) - 100
    header_y = min(w["top"] for w in right_headers) - 20

    bounds = _find_table_bounds(lines, header_y, header_x, page_width)
    if bounds is None:
        return None

    # Validate: the second table should have actual legend-like content,
    # not just a title block. Check if it has reasonable number of rows
    # and that its content includes legend-like items (not just metadata).
    if len(bounds.row_ys) < 3:
        return None

    return bounds


# ---------------------------------------------------------------------------
# Core legend extraction
# ---------------------------------------------------------------------------

def _find_legend_header(words: list[dict]) -> Optional[tuple[float, float, int]]:
    """
    Find the legend header on any page.
    Returns (y_top, x0, page_index) or None.

    Uses two-pass search:
    Pass 1: Only "Условные" and "Легенда" (highly reliable, no false positives).
    Pass 2: Other patterns with guards (only if pass 1 found nothing).
    """
    # --- Pass 1: primary patterns (no false positives) ---
    primary_patterns = [p for p in LEGEND_HEADER_PATTERNS
                        if p.pattern in (r"Условные", r"Легенда")]
    for w in words:
        for pattern in primary_patterns:
            if pattern.search(w["text"]):
                return (w["top"], w["x0"], w.get("page_index", 0))

    # --- Pass 2: secondary patterns with guards ---
    # Build a set of Y positions where column/table-header words appear.
    # These indicate the word is part of a table header row, not a legend title.
    col_header_words = {"Наименование", "Примечание", "Примечания",
                        "Поставщик", "характеристика"}
    col_header_ys: set[float] = set()
    for w in words:
        if w["text"] in col_header_words:
            col_header_ys.add(round(w["top"], 0))

    secondary_patterns = [p for p in LEGEND_HEADER_PATTERNS
                          if p.pattern not in (r"Условные", r"Легенда")]
    for w in words:
        for pattern in secondary_patterns:
            if pattern.search(w["text"]):
                # Guard: skip "Обозначени" when it's a column header
                if pattern.pattern == r"Обозначени":
                    w_y_rounded = round(w["top"], 0)
                    # Use 15pt tolerance: column headers may span 2 lines
                    is_col_header = any(
                        abs(w_y_rounded - cy) <= 15 for cy in col_header_ys
                    )
                    if is_col_header:
                        continue
                    # Also skip if the word is in genitive/other case forms
                    # that appear in running text (e.g. "обозначений")
                    if w["text"].lower() not in ("обозначения", "обозначение"):
                        continue
                return (w["top"], w["x0"], w.get("page_index", 0))
    return None

def _find_row_slot(y: float, row_ys: list[float]) -> int:
    """
    Find which table row slot a given Y coordinate belongs to.
    Returns the index of the row (between row_ys[i] and row_ys[i+1]).
    """
    for i in range(len(row_ys) - 1):
        if row_ys[i] - 3 <= y < row_ys[i + 1] - 3:
            return i
    # After last row divider
    if row_ys and y >= row_ys[-1] - 3:
        return len(row_ys) - 1
    return 0



def _is_char_split(text: str) -> bool:
    """Check if text appears to be character-split (mostly single-char words).

    Example: 'Щ и т с и ло в о й' is character-split.
    """
    words = text.split()
    if len(words) < 3:
        return False
    single_char = sum(1 for w in words if len(w) <= 2)
    return single_char / len(words) > 0.5

def _extract_items_from_table(
    words: list[dict],
    bounds: TableBounds,
    y_tol: float = 5,
    page_lines: Optional[list[dict]] = None,
    page_rects: Optional[list[dict]] = None,
) -> list[LegendItem]:
    """
    Extract legend items from within detected table boundaries.

    Uses table row lines (row_ys) to determine cell boundaries.
    All text within the same table cell (between consecutive row_ys)
    is merged into one item. This correctly handles multi-line
    descriptions within a single table row.

    If page_lines/page_rects are provided, detects the dominant color
    of graphical elements in each row's symbol cell.
    """
    if len(bounds.col_xs) < 2:
        sym_x_boundary = bounds.x0 + 75
    else:
        # Second column divider separates symbol from description
        sym_x_boundary = bounds.col_xs[1] if len(bounds.col_xs) > 1 else bounds.x0 + 75

    # Filter words inside the table (below first row line, above bottom)
    # Skip the header area (first 1-2 row_ys are title + column headers)
    if len(bounds.row_ys) < 2:
        return []

    # The first data row starts after the second horizontal line (header row)
    first_data_y = bounds.row_ys[1]
    area_words = [
        w for w in words
        if first_data_y - 2 < w["top"] < bounds.y1 + 5
        and bounds.x0 - 5 < w["x0"] < bounds.x1 + 5
    ]

    if not area_words:
        return []

    # Determine data row boundaries from row_ys
    # data_row_ys = row_ys starting from the first data row
    data_row_ys = [y for y in bounds.row_ys if y >= first_data_y - 1]

    # Group words into table cells (rows) based on horizontal line positions
    # Each cell is between data_row_ys[i] and data_row_ys[i+1]
    cells: dict[int, list[dict]] = {}  # row_index → words
    for w in area_words:
        slot = _find_row_slot(w["top"], data_row_ys)
        if slot not in cells:
            cells[slot] = []
        cells[slot].append(w)

    items: list[LegendItem] = []

    for slot_idx in sorted(cells.keys()):
        cell_words = cells[slot_idx]

        # Separate symbol-area words from description-area words
        sym_parts = [w for w in cell_words if w["x0"] < sym_x_boundary]
        desc_parts = [w for w in cell_words if w["x0"] >= sym_x_boundary and w["x1"] <= bounds.x1 + 5]

        # Extract symbol text(s)
        numeric_syms = [w for w in sym_parts if SYMBOL_RE.match(w["text"])]
        text_syms = [w for w in sym_parts if TEXT_SYMBOL_RE.match(w["text"])]

        # Prefer numeric symbol (1, 1А, 9А, 10А) over text symbol (ВЫХОД)
        # If both exist, use numeric as the item symbol
        if numeric_syms:
            sym_text = numeric_syms[0]["text"]
        elif text_syms:
            sym_text = text_syms[0]["text"]
        else:
            sym_text = ""

        # Build full description: group by text lines, join all
        desc_lines = _y_group(desc_parts, tol=y_tol)
        desc_text = " ".join(_join_words(ws) for _, ws in desc_lines).strip()

        # Skip column header row
        if any(kw in desc_text for kw in COL_HEADERS):
            continue

        # Skip title block content
        if _is_title_block_text(desc_text):
            continue

        # Strip construction note phrases that leaked into description
        desc_text = _strip_construction_notes(desc_text)

        # Sanitize overly long descriptions (concatenated equipment names)
        desc_text = _sanitize_long_description(desc_text)

        # Skip empty rows
        if not sym_text and not desc_text:
            continue

        # Compute row bbox
        all_cell_words = sym_parts + desc_parts
        if all_cell_words:
            row_bbox = (
                min(w["x0"] for w in all_cell_words),
                min(w["top"] for w in all_cell_words),
                max(w["x1"] for w in all_cell_words),
                max(w["bottom"] for w in all_cell_words),
            )
        else:
            row_bbox = (bounds.x0, data_row_ys[slot_idx] if slot_idx < len(data_row_ys) else bounds.y0,
                       bounds.x1, bounds.y1)

        category = _detect_category(desc_text)

        # Detect dominant color of graphical elements in the symbol cell
        row_color = ""
        if page_lines is not None or page_rects is not None:
            row_color = _detect_row_color(
                page_lines or [],
                page_rects or [],
                row_bbox,
                sym_x_boundary,
            )

        item = LegendItem(
            symbol=sym_text,
            description=desc_text,
            category=category,
            bbox=row_bbox,
            color=row_color,
        )
        items.append(item)

    return items



# ---------------------------------------------------------------------------
# Adaptive x_tolerance helpers
# ---------------------------------------------------------------------------


def _repair_char_split_descriptions(
    items: list[LegendItem],
    page,
    bounds: TableBounds,
) -> list[LegendItem]:
    """Repair character-split descriptions using page.crop().extract_text().

    When pdfplumber extract_words() splits Cyrillic text into individual
    characters (e.g. 'Щ и т с и ло в о й'), this function falls back to
    extract_text() on the cropped row area to get proper word grouping.
    """
    if not items:
        return items

    # Check if any descriptions are character-split
    has_splits = any(_is_char_split(it.description) for it in items if it.description)
    if not has_splits:
        return items

    # Determine the description column boundaries
    if len(bounds.col_xs) >= 2:
        desc_x0 = bounds.col_xs[1]  # after symbol column
    else:
        desc_x0 = bounds.x0 + 75
    desc_x1 = bounds.x1

    repaired = []
    for item in items:
        if not _is_char_split(item.description):
            repaired.append(item)
            continue

        # Crop the row area and extract text
        row_y0 = item.bbox[1] - 2
        row_y1 = item.bbox[3] + 2

        try:
            cropped = page.crop((desc_x0, row_y0, desc_x1, row_y1))
            text = cropped.extract_text() or ''
            text = text.strip()
            # Remove newlines, collapse whitespace
            text = ' '.join(text.split())
            if text and len(text) > 3:
                item = LegendItem(
                    symbol=item.symbol,
                    description=text,
                    category=_detect_category(text),
                    bbox=item.bbox,
                    color=item.color,
                )
        except Exception:
            pass  # keep original if crop fails

        repaired.append(item)

    return repaired

def _has_good_descriptions(items: list[LegendItem], threshold: float = 0.5) -> bool:
    """Check if enough items have non-empty descriptions.

    Returns True if at least *threshold* fraction of items with symbols
    also have descriptions longer than 3 characters.
    """
    sym_items = [it for it in items if it.symbol]
    if not sym_items:
        return False
    good = sum(1 for it in sym_items if len(it.description.strip()) > 3)
    return good / len(sym_items) >= threshold


def _extract_with_tolerance(
    page,
    page_idx: int,
    x_tol: int,
) -> tuple[list[dict], Optional[tuple[float, float, int]]]:
    """Extract words from *page* with given x_tolerance and find legend header.

    Returns (words_raw, header_info) where header_info may be None.
    """
    words_raw = page.extract_words(x_tolerance=x_tol, y_tolerance=3) or []
    for w in words_raw:
        w["page_index"] = page_idx
    header_info = _find_legend_header(words_raw)
    return words_raw, header_info

# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def parse_legend(pdf_path: str) -> LegendResult:
    """
    Parse legend table(s) from a PDF file.

    Searches ALL pages for a legend header. If a header is found but no items
    are extracted, continues searching subsequent pages. If no header is found
    on any page, falls back to content-based detection (looking for tables
    with equipment keywords).

    Uses adaptive x_tolerance: tries x_tolerance=3 first (precise), and
    if the extracted descriptions are mostly empty, retries with x_tolerance=7
    to handle PDFs where pdfplumber splits Cyrillic text into individual
    characters at low tolerance.
    """
    X_TOLERANCES = [3, 7]

    # Track pages where no header was found, for content-based fallback
    no_header_pages: list[tuple[int, object]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            lines = page.lines or []
            rects = page.rects or []

            best_result: Optional[LegendResult] = None
            header_found_on_page = False

            for x_tol in X_TOLERANCES:
                words_raw, header_info = _extract_with_tolerance(
                    page, page_idx, x_tol,
                )

                # 1) DETECT legend header
                if header_info is None:
                    continue

                header_found_on_page = True
                header_y, header_x, _ = header_info

                # 2) DETERMINE table boundaries using PDF lines
                bounds = _find_table_bounds(lines, header_y, header_x, page.width)

                if bounds is None:
                    # Fallback: use word-based heuristic
                    bounds = _fallback_bounds(words_raw, header_y, header_x, page.width)
                    if bounds is None:
                        continue

                # 3) EXTRACT items from the main table (with color detection)
                items = _extract_items_from_table(
                    words_raw, bounds,
                    page_lines=lines, page_rects=rects,
                )

                # 3b) Repair character-split descriptions using extract_text fallback
                items = _repair_char_split_descriptions(items, page, bounds)

                # 4) Check for a second (GPC-style) table column
                columns_detected = 1
                second_bounds = _find_second_table(lines, words_raw, bounds, page.width)
                if second_bounds is not None:
                    second_items = _extract_items_from_table(
                        words_raw, second_bounds,
                        page_lines=lines, page_rects=rects,
                    )
                    # Filter out title-block-like items from the second table
                    second_items = [
                        it for it in second_items
                        if not _is_title_block_text(it.description)
                        and not _is_title_block_text(it.symbol)
                    ]
                    if second_items:
                        items.extend(second_items)
                        columns_detected = 2

                # 5) Compute overall legend_bbox
                if items:
                    overall_x0 = bounds.x0
                    overall_y0 = bounds.y0
                    overall_x1 = bounds.x1
                    overall_y1 = bounds.y1
                    if second_bounds and columns_detected == 2:
                        overall_x1 = max(overall_x1, second_bounds.x1)
                        overall_y1 = max(overall_y1, second_bounds.y1)
                    legend_bbox = (overall_x0, overall_y0, overall_x1, overall_y1)
                else:
                    legend_bbox = (bounds.x0, bounds.y0, bounds.x1, bounds.y1)

                result = LegendResult(
                    items=items,
                    legend_bbox=legend_bbox,
                    page_index=page_idx,
                    columns_detected=columns_detected,
                )

                # If descriptions look good, accept immediately
                if _has_good_descriptions(items):
                    return result

                # Otherwise, keep as best candidate and try next tolerance
                if best_result is None or len(items) > len(best_result.items):
                    best_result = result

            # If we tried all tolerances on this page and got good items, return
            if best_result is not None and best_result.items:
                return best_result

            # Track pages for content-based fallback: either no header found,
            # or header found but no items extracted (secondary pattern false positive)
            if not header_found_on_page or best_result is None or not best_result.items:
                no_header_pages.append((page_idx, page))

        # --- Content-based fallback ---
        # No legend found via headers on any page.
        # Try content-based detection on pages where no header was found.
        for page_idx, page in no_header_pages:
            result = _content_based_legend_search(page, page_idx)
            if result is not None and result.items:
                return result

    # No legend found on any page
    return LegendResult()



def _fallback_bounds(
    words: list[dict],
    header_y: float,
    header_x: float,
    page_width: float,
) -> Optional[TableBounds]:
    """
    Fallback table boundary detection using word positions when PDF lines
    don't provide enough structure.
    """
    # Find column header words near the legend header
    oboz_x = None
    naim_x = None
    prim_x = None

    for w in words:
        if abs(w["top"] - header_y) > 50:
            continue
        if w["top"] <= header_y:
            continue
        if "Обозначение" in w["text"]:
            oboz_x = w["x0"]
        elif "Наименование" in w["text"]:
            naim_x = w["x0"]
        elif "Примечание" in w["text"]:
            prim_x = w["x0"]

    if oboz_x is not None:
        # Determine table extent from words
        table_x0 = oboz_x - 20
        sym_divider = naim_x - 10 if naim_x else oboz_x + 70
        notes_divider = prim_x - 10 if prim_x else sym_divider + 500
        table_x1 = notes_divider + 100

        # Find bottom by scanning words
        area_words = [
            w for w in words
            if w["top"] > header_y + 15
            and table_x0 - 10 < w["x0"] < table_x1 + 10
        ]
        if area_words:
            y_bottom = max(w["bottom"] for w in area_words)

            # Build approximate row Y positions from text lines
            rows = _y_group(area_words, tol=5)
            row_ys = [header_y] + [y for y, _ in rows]

            return TableBounds(
                x0=table_x0,
                x1=table_x1,
                y0=header_y,
                y1=y_bottom + 5,
                col_xs=[table_x0, sym_divider, notes_divider, table_x1],
                row_ys=sorted(row_ys),
            )

    # --- Strategy 2: Equipment keyword cluster ---
    # Look for words below the header that contain equipment description keywords.
    # If we find a cluster, infer table bounds from text positions.
    equip_keywords = [
        "светильник", "розетка", "выключатель", "кабель", "щит",
        "провод", "лампа", "датчик", "извещатель", "оповещатель",
    ]

    below_header_words = [
        w for w in words
        if header_y < w["top"] < header_y + 500
        and abs(w["x0"] - header_x) < 400
    ]
    if not below_header_words:
        return None

    # Count how many equipment keyword matches exist
    equip_matches = []
    for w in below_header_words:
        w_lower = w["text"].lower()
        for kw in equip_keywords:
            if kw in w_lower:
                equip_matches.append(w)
                break

    if len(equip_matches) < 2:
        return None

    # Infer table bounds from all words in the equipment cluster area
    all_x0 = min(w["x0"] for w in below_header_words)
    all_x1 = max(w["x1"] for w in below_header_words)
    all_y_bottom = max(w["bottom"] for w in below_header_words)

    # Try to detect a narrow symbol column on the left (50-100px wide)
    # by looking for short numeric/symbol words
    sym_words = [w for w in below_header_words if SYMBOL_RE.match(w["text"])]
    if sym_words:
        sym_x_max = max(w["x1"] for w in sym_words) + 10
        sym_divider_val = min(sym_x_max, all_x0 + 100)
    else:
        sym_divider_val = all_x0 + 75

    table_x0 = all_x0 - 10
    table_x1 = all_x1 + 10

    rows = _y_group(below_header_words, tol=5)
    row_ys = [header_y] + [y for y, _ in rows]

    return TableBounds(
        x0=table_x0,
        x1=table_x1,
        y0=header_y,
        y1=all_y_bottom + 5,
        col_xs=[table_x0, sym_divider_val, table_x1],
        row_ys=sorted(row_ys),
    )


# Equipment keywords for content-based legend search (no header found)
_EQUIPMENT_KEYWORDS_LOWER = [
    "светильник", "розетка", "выключатель", "кабель", "щит",
    "провод", "лампа", "датчик", "извещатель", "оповещатель",
    "трасса", "автомат", "узо", "дифавтомат",
]


def _content_based_legend_search(
    page,
    page_idx: int,
) -> Optional[LegendResult]:
    """
    Search for a legend table on a page using content-based heuristics
    when no legend header was found.

    Looks for small tables (detected via PDF lines) that contain equipment
    description keywords. This handles PDFs where the legend header text
    is missing or uses an unknown term.

    Returns a LegendResult if a likely legend table is found, else None.
    """
    from collections import Counter

    words_raw = page.extract_words(x_tolerance=3, y_tolerance=3) or []
    for w in words_raw:
        w["page_index"] = page_idx

    lines = page.lines or []
    rects = page.rects or []

    # Gather horizontal lines
    h_lines = [
        l for l in lines
        if abs(l["top"] - l["bottom"]) < 2
        and abs(l["x1"] - l["x0"]) > MIN_TABLE_LINE_LEN
    ]

    # Gather vertical lines
    v_lines = [
        l for l in lines
        if abs(l["x0"] - l["x1"]) < 2
        and abs(l["bottom"] - l["top"]) > MIN_TABLE_LINE_LEN
    ]

    if len(h_lines) < 3 or len(v_lines) < 2:
        return None

    # Group horizontal lines by x0/x1 to find potential tables
    # A table is a group of 3+ horizontal lines with the same x0 and x1
    x_pairs: Counter = Counter()
    for l in h_lines:
        key = (round(l["x0"], 0), round(l["x1"], 0))
        x_pairs[key] += 1

    # Find the most frequent x0/x1 pair with at least 3 lines (= 2+ rows)
    candidates = [(pair, count) for pair, count in x_pairs.items() if count >= 3]
    if not candidates:
        return None

    # Sort by count descending, then prefer narrower tables (more likely legend)
    candidates.sort(key=lambda pc: (-pc[1], pc[0][1] - pc[0][0]))

    for (rx0, rx1), _count in candidates:
        table_width = rx1 - rx0
        # Legend tables are typically 100-600pt wide
        if table_width < 80 or table_width > 700:
            continue

        # Get the matching horizontal lines
        matched_h = [
            l for l in h_lines
            if abs(round(l["x0"], 0) - rx0) < 5
            and abs(round(l["x1"], 0) - rx1) < 5
        ]
        row_ys = sorted(set(round(l["top"], 1) for l in matched_h))
        y0 = min(row_ys)
        y1 = max(row_ys)

        # Check for vertical lines in this range
        matched_v = [
            l for l in v_lines
            if rx0 - 5 < l["x0"] < rx1 + 5
            and l["top"] < y0 + 20
            and l["bottom"] > y1 - 20
        ]
        col_xs = sorted(set(round(l["x0"], 1) for l in matched_v))

        # Need at least one internal column divider for a legend table
        if len(col_xs) < 2:
            continue

        # Check content: words inside this table area
        table_words = [
            w for w in words_raw
            if y0 - 5 < w["top"] < y1 + 5
            and rx0 - 5 < w["x0"] < rx1 + 5
        ]

        # Count equipment keywords
        equip_count = 0
        for w in table_words:
            w_lower = w["text"].lower()
            for kw in _EQUIPMENT_KEYWORDS_LOWER:
                if kw in w_lower:
                    equip_count += 1
                    break

        # Need at least 2 equipment keyword matches to be confident
        if equip_count < 2:
            continue

        # This looks like a legend table
        bounds = TableBounds(
            x0=rx0, x1=rx1, y0=y0, y1=y1,
            col_xs=col_xs, row_ys=row_ys,
        )

        items = _extract_items_from_table(
            words_raw, bounds,
            page_lines=lines, page_rects=rects,
        )

        if items:
            legend_bbox = (bounds.x0, bounds.y0, bounds.x1, bounds.y1)
            return LegendResult(
                items=items,
                legend_bbox=legend_bbox,
                page_index=page_idx,
                columns_detected=1,
            )

    return None


# ---------------------------------------------------------------------------
# CLI interface
# ---------------------------------------------------------------------------

def main() -> None:
    """CLI entry point for testing."""
    if len(sys.argv) < 2:
        print("Usage: python pdf_legend_parser.py <path.pdf>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    result = parse_legend(pdf_path)

    if not result.items:
        print(f"No legend found in: {pdf_path}")
        sys.exit(0)

    print(f"Legend found on page {result.page_index + 1}")
    print(f"Legend bbox: ({result.legend_bbox[0]:.1f}, {result.legend_bbox[1]:.1f}, "
          f"{result.legend_bbox[2]:.1f}, {result.legend_bbox[3]:.1f})")
    print(f"Columns detected: {result.columns_detected}")
    print(f"Total items: {len(result.items)}")
    print("-" * 80)

    for i, item in enumerate(result.items, 1):
        sym_display = item.symbol if item.symbol else "(no symbol)"
        cat_display = f" [{item.category}]" if item.category else ""
        color_display = f" <{item.color}>" if item.color else ""
        print(f"  {i:3d}. {sym_display:>10s}  {item.description}{cat_display}{color_display}")
        print(f"       bbox: ({item.bbox[0]:.1f}, {item.bbox[1]:.1f}, "
              f"{item.bbox[2]:.1f}, {item.bbox[3]:.1f})")

    print("-" * 80)


if __name__ == "__main__":
    main()
