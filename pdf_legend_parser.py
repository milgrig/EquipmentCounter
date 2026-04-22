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
    # S1.5 (T020): 0-based PDF page this item was extracted from.  When a
    # LegendResult is produced by multi-page aggregation, individual items
    # may come from different pages.  For single-page parses this matches
    # LegendResult.page_index.
    page_index: int = 0


@dataclass
class LegendResult:
    """Complete parsing result for one legend table."""
    items: list[LegendItem] = field(default_factory=list)
    legend_bbox: tuple[float, float, float, float] = (0, 0, 0, 0)  # overall table bbox
    page_index: int = 0
    columns_detected: int = 1  # 1 = single column, 2 = two-column (GPC-style)
    # S1.9 (T024): which detection path produced this candidate.
    # One of: "header", "content", "spec", "density", "" (empty).
    # Optional "reversed:" prefix (S1.8 / T023) when the page required
    # reversed-char preprocessing before the fallback fired, e.g.
    # "reversed:spec" or "reversed:header".
    detection_method: str = ""


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

# S1.3 (T018): Legend row identifier pattern used by
# `_content_based_legend_search` when line-based detection fails.
# Accepts: a bare 1–2-digit number (1..99) OR a letter-prefixed identifier
# such as "5A", "2Б", "10А".  This is deliberately broader than SYMBOL_RE
# because we only need to recognise that a row STARTS with an identifier,
# not that the whole token is a canonical symbol.
ROW_IDENT_RE = re.compile(
    r"^(?:\d{1,2}|\d{1,2}[A-Za-zА-Яа-яЁё]{1,3}|[A-Za-zА-Яа-яЁё]\d{1,2})$"
)

# Text-based symbol identifiers that are NOT numeric (e.g. "ВЫХОД")
TEXT_SYMBOL_RE = re.compile(r"^[A-ZА-ЯЁ]{3,}$")

# Emergency circuit variant (e.g. 2А → base 2 + "А")
CIRCUIT_VARIANT_RE = re.compile(r"^(\d+)(А|АЭ)$")

# Column header keywords
COL_HEADERS = {"Обозначение", "Наименование", "Примечание"}

# S1.8 (T023): Reversed-text detection.
# Some CAD-exported PDFs embed Cyrillic text with characters in reversed
# (right-to-left) order within each content-stream operator.  pdfplumber
# returns those tokens literally — e.g. "Защита" becomes "атищаЗ".  We
# detect this by checking whether a short, common stem-set matches more
# often on reversed tokens than on forward tokens.  When confirmed we
# flip the `text` field of every word before the usual legend parsing
# sees it (bbox coordinates are preserved — only glyph order flips).
REVERSED_TEXT_STEMS = {
    # nouns frequently present in Russian electrical/engineering legends
    "защита", "наименование", "обозначение", "примечание",
    "автомат", "кабель", "провод", "проводка",
    "выключатель", "розетка", "светильник",
    "щит", "щиток", "группа", "тип", "номер",
    # column headers from spec-style tables
    "поз", "кол",
    # common construction/equipment words
    "этаж", "план", "схема", "система", "линия",
    "питание", "электрика", "освещение",
}

_CYR_WORD_RE = re.compile(r"^[А-Яа-яЁё]{4,}$")


def _is_text_reversed(words: list[dict]) -> bool:
    """S1.8 (T023): detect if a page's extracted text is char-reversed.

    Heuristic: for every Cyrillic-only token of length >= 4, check
    whether its lowercased form STARTS with (or equals) any stem in
    ``REVERSED_TEXT_STEMS``, and whether its REVERSED lowercased form
    does.  If reversed-hits clearly outnumber forward-hits, the page is
    flagged as reversed.

    Threshold: reversed_hits >= 3 (absolute) AND
               reversed_hits >= 2 * forward_hits.

    Returns ``True`` when the page should be pre-reversed before legend
    detection, ``False`` otherwise (which is the common case — most
    PDFs are correctly oriented).
    """
    if not words:
        return False

    forward_hits = 0
    reversed_hits = 0

    for w in words:
        text = (w.get("text") or "").strip()
        if not _CYR_WORD_RE.match(text):
            continue
        low = text.lower()
        low_rev = low[::-1]
        # Forward match: any stem is a prefix of the token, OR the token
        # equals the stem.  Using `startswith` catches inflected forms
        # (e.g. "защиты", "автоматов") without a full morphology table.
        if any(low == s or low.startswith(s) for s in REVERSED_TEXT_STEMS):
            forward_hits += 1
        if any(low_rev == s or low_rev.startswith(s) for s in REVERSED_TEXT_STEMS):
            reversed_hits += 1

    if reversed_hits < 3:
        return False
    if reversed_hits < 2 * max(forward_hits, 1):
        return False
    return True


def _reverse_cyrillic_words(words: list[dict]) -> list[dict]:
    """S1.8 (T023): return a shallow-copied word list with each text
    reversed character-by-character.

    Only the ``text`` field is flipped.  Geometry (x0/x1/top/bottom)
    and any other fields are preserved unchanged — glyph-order reversal
    does not move the bounding box.  Non-string or missing ``text``
    entries are passed through untouched.
    """
    out: list[dict] = []
    for w in words:
        if not isinstance(w, dict):
            out.append(w)
            continue
        new_w = dict(w)
        txt = new_w.get("text")
        if isinstance(txt, str) and txt:
            new_w["text"] = txt[::-1]
        out.append(new_w)
    return out

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

# Minimum length for a horizontal/vertical line to be considered a table border.
# Calibration history:
#   - Original: 80pt (rejected small legend tables entirely)
#   - S1.1: 50pt — admitted pos_8_2_em (78pt vertical) but still rejected
#           compact GPC-style tables.
#   - S1.2 (T017): 25pt — required for small GPC tables (abk_em, abk_eg).
#     These drawings use 2-row legends whose vertical borders are ~30pt
#     tall and horizontal row separators as short as 26–40pt.
#     25pt is still comfortably above typical tick / hatch strokes
#     (normally < 10pt) so false-positives stay rare; additional
#     filtering (line grouping, X-range overlap) keeps real tables.
MIN_TABLE_LINE_LEN = 25

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


def _collect_cell_colors(
    lines: list[dict],
    rects: list[dict],
    row_bbox: tuple[float, float, float, float],
    sym_x_boundary: float,
) -> set[str]:
    """Collect the set of distinct non-black color labels in the symbol cell.

    Uses the same spatial filtering as _detect_row_color but returns ALL
    colors present rather than the most frequent one.  Used to detect
    dual-color rows (e.g. blue working + red emergency) that need splitting.
    """
    x0, y0, x1, y1 = row_bbox
    sym_x1 = min(x1, sym_x_boundary + 10)
    colors: set[str] = set()

    for ln in lines:
        lx0 = min(ln["x0"], ln["x1"])
        lx1 = max(ln["x0"], ln["x1"])
        ly0 = min(ln["top"], ln["bottom"])
        ly1 = max(ln["top"], ln["bottom"])
        if lx1 < x0 - 5 or lx0 > sym_x1 + 5:
            continue
        if ly1 < y0 - 3 or ly0 > y1 + 3:
            continue
        line_len = max(abs(ln["x1"] - ln["x0"]), abs(ln["bottom"] - ln["top"]))
        if line_len > (y1 - y0) * 2 and line_len > 100:
            continue
        rgb = _normalize_color(ln.get("stroking_color"))
        if rgb is None:
            continue
        label = _classify_rgb(rgb)
        if label and label != "black":
            colors.add(label)

    for rect in rects:
        rx0, ry0 = rect["x0"], rect["top"]
        rx1, ry1 = rect["x1"], rect["bottom"]
        if rx1 < x0 - 5 or rx0 > sym_x1 + 5:
            continue
        if ry1 < y0 - 3 or ry0 > y1 + 3:
            continue
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
                colors.add(label)

    return colors


def _detect_sub_row_color(
    lines: list[dict],
    rects: list[dict],
    row_bbox: tuple[float, float, float, float],
    sym_x_boundary: float,
    y_min: float,
    y_max: float,
) -> str:
    """Detect dominant color in a vertical sub-range of the symbol cell.

    Like _detect_row_color but restricts the Y range to [y_min, y_max]
    instead of the full row_bbox height.  Used after splitting a
    dual-color row to assign a color to each sub-item.
    """
    x0, _, x1, _ = row_bbox
    sym_x1 = min(x1, sym_x_boundary + 10)
    color_counts: dict[str, int] = {}

    for ln in lines:
        lx0 = min(ln["x0"], ln["x1"])
        lx1 = max(ln["x0"], ln["x1"])
        ly0 = min(ln["top"], ln["bottom"])
        ly1 = max(ln["top"], ln["bottom"])
        if lx1 < x0 - 5 or lx0 > sym_x1 + 5:
            continue
        if ly1 < y_min - 3 or ly0 > y_max + 3:
            continue
        line_len = max(abs(ln["x1"] - ln["x0"]), abs(ln["bottom"] - ln["top"]))
        if line_len > (y_max - y_min) * 2 and line_len > 100:
            continue
        rgb = _normalize_color(ln.get("stroking_color"))
        if rgb is None:
            continue
        label = _classify_rgb(rgb)
        if label and label != "black":
            color_counts[label] = color_counts.get(label, 0) + 1

    for rect in rects:
        rx0, ry0 = rect["x0"], rect["top"]
        rx1, ry1 = rect["x1"], rect["bottom"]
        if rx1 < x0 - 5 or rx0 > sym_x1 + 5:
            continue
        if ry1 < y_min - 3 or ry0 > y_max + 3:
            continue
        rect_w = rx1 - rx0
        rect_h = ry1 - ry0
        if rect_w > sym_x1 - x0 + 20 and rect_h > (y_max - y_min) + 10:
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
    return max(color_counts, key=color_counts.get)  # type: ignore[arg-type]


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
    # S1.2 (T017): tolerate partial vertical frames.  Some GPC-style small
    # legends (abk_em, abk_eg) draw only horizontal row separators and omit
    # the outer vertical borders entirely, so matched_v may be empty.  In
    # that case we fall back to treating the table edges (table_x0, table_x1)
    # as implicit column bounds.  We also widen the Y tolerance slightly —
    # compact tables sometimes have verticals that stop a few points short
    # of the first/last horizontal separator.
    matched_v = [
        l for l in v_lines
        if table_x0 - 5 < l["x0"] < table_x1 + 5
        and l["top"] < y0 + 20
        and l["bottom"] > y1 - 20
    ]

    if not matched_v:
        # Relaxed pass: accept vertical lines that merely overlap the
        # table's Y range rather than spanning it top-to-bottom.  This
        # catches internal column dividers when the outer frame is absent.
        matched_v = [
            l for l in v_lines
            if table_x0 - 5 < l["x0"] < table_x1 + 5
            and l["bottom"] > y0 - 5
            and l["top"] < y1 + 5
        ]

    col_xs = sorted(set(round(l["x0"], 1) for l in matched_v))

    # Horizontal-only table fallback: if no verticals were found at all,
    # synthesise column boundaries from the table edges so downstream
    # row/column extraction can still run.  This is the abk_em / abk_eg
    # shape: the legend is just a stack of rows delimited by horizontal
    # rules without any visible vertical frame.
    if not col_xs:
        col_xs = [round(table_x0, 1), round(table_x1, 1)]

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

        # Build full description early (shared by all sub-items if split)
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

        # Compute row bbox (needed for both split and normal paths)
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

        # --- Dual-color split detection ---
        # On ГПК-style drawings, a single legend row may contain two symbols
        # stacked vertically: e.g. blue "4" (working) above red "4АЭ" (emergency).
        # Detect this pattern and split into separate LegendItems.
        if (len(numeric_syms) >= 2
                and (page_lines is not None or page_rects is not None)):
            sym_y_groups = _y_group(numeric_syms, tol=3)

            if len(sym_y_groups) >= 2:
                colors_in_cell = _collect_cell_colors(
                    page_lines or [], page_rects or [],
                    row_bbox, sym_x_boundary,
                )

                if len(colors_in_cell) >= 2:
                    # SPLIT: create one LegendItem per Y-group
                    sym_y_groups.sort(key=lambda g: g[0])
                    category = _detect_category(desc_text)

                    # Compute Y boundaries between consecutive symbol groups.
                    # Use bottom of previous group / top of next group to find
                    # the gap between the two stacked symbols.  This is more
                    # accurate than using tops, because the symbol TEXT sits
                    # above the graphical symbol and may overlap in Y range
                    # with the graphics below it.
                    y_boundaries: list[float] = []
                    for g_idx in range(len(sym_y_groups)):
                        if g_idx == 0:
                            y_boundaries.append(row_bbox[1])
                        else:
                            # bottom of previous group's words
                            prev_words = sym_y_groups[g_idx - 1][1]
                            prev_bottom = max(w["bottom"] for w in prev_words)
                            # top of current group's words
                            curr_top = sym_y_groups[g_idx][0]
                            y_boundaries.append((prev_bottom + curr_top) / 2)
                    y_boundaries.append(row_bbox[3])

                    for g_idx, (g_y, g_words) in enumerate(sym_y_groups):
                        sub_sym_text = g_words[0]["text"]
                        sub_y_min = y_boundaries[g_idx]
                        sub_y_max = y_boundaries[g_idx + 1]
                        sub_bbox = (row_bbox[0], sub_y_min, row_bbox[2], sub_y_max)

                        sub_color = _detect_sub_row_color(
                            page_lines or [], page_rects or [],
                            row_bbox, sym_x_boundary,
                            sub_y_min, sub_y_max,
                        )

                        if not sub_sym_text and not desc_text:
                            continue

                        items.append(LegendItem(
                            symbol=sub_sym_text,
                            description=desc_text,
                            category=category,
                            bbox=sub_bbox,
                            color=sub_color,
                        ))

                    continue  # skip normal single-item creation below

        # --- Normal single-item path ---
        # Prefer numeric symbol (1, 1А, 9А, 10А) over text symbol (ВЫХОД)
        if numeric_syms:
            sym_text = numeric_syms[0]["text"]
        elif text_syms:
            sym_text = text_syms[0]["text"]
        else:
            sym_text = ""

        # Skip empty rows
        if not sym_text and not desc_text:
            continue

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

    S1.8 (T023): if the page's Cyrillic text is character-reversed
    (common on CAD-exported PDFs), flip every word's `text` before
    returning.  The caller doesn't need to know — geometry is
    preserved.
    """
    words_raw = page.extract_words(x_tolerance=x_tol, y_tolerance=3) or []
    if _is_text_reversed(words_raw):
        words_raw = _reverse_cyrillic_words(words_raw)
        # Tag the page so parse_legend can annotate detection_method
        # with a "reversed:" prefix.
        try:
            setattr(page, "_legend_reversed", True)
        except Exception:
            pass
    for w in words_raw:
        w["page_index"] = page_idx
    header_info = _find_legend_header(words_raw)
    return words_raw, header_info

# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def _tag_method(base: str, page) -> str:
    """S1.9 (T024): return `base` with a "reversed:" prefix when the
    page was flagged as char-reversed by `_extract_with_tolerance`.

    `base` is one of "header", "content", "spec", "density".
    """
    try:
        if getattr(page, "_legend_reversed", False):
            return f"reversed:{base}"
    except Exception:
        pass
    return base


def _stamp_page_index(result: LegendResult, page_idx: int) -> LegendResult:
    """S1.5 (T020): ensure every item in `result` carries `page_idx`.

    `LegendItem.page_index` defaults to 0, so callers that built items
    without explicitly setting it (the norm today) get the correct page
    stamped here.  Returns the same `result` for chaining convenience.
    """
    for it in result.items:
        it.page_index = page_idx
    return result


def _merge_legend_candidates(
    candidates: list[LegendResult],
) -> LegendResult:
    """
    S1.5 (T020): merge multiple per-page LegendResults into a single
    aggregate result with symbol-based de-duplication.

    Strategy:
      * The densest candidate becomes the base (its bbox / page_index /
        columns_detected are kept on the returned result so that
        downstream consumers that rely on those fields still see a
        meaningful primary page).
      * Remaining candidates are walked in density-descending order.
        For each item, if its symbol is empty or NOT already present
        in the aggregate, it is appended.  If a symbol collides, the
        item with the longer description wins (the other is discarded).

    Items with empty symbols are kept verbatim (they're usually unique
    "text-only" legend rows like "ВЫХОД" — dedup by description instead).
    """
    if not candidates:
        return LegendResult()

    # Sort candidates by density desc; earlier page wins on tie.
    ordered = sorted(
        candidates,
        key=lambda r: (-_legend_density_score(r), r.page_index),
    )
    base = ordered[0]

    # Normalise a symbol for dedup comparison.
    def _key(it: LegendItem) -> str:
        if it.symbol:
            return ("S", it.symbol.strip().lower())
        # Items without a symbol are deduped on description.
        return ("D", (it.description or "").strip().lower())

    by_key: dict = {}
    order: list = []
    for it in base.items:
        k = _key(it)
        if k not in by_key:
            by_key[k] = it
            order.append(k)

    for cand in ordered[1:]:
        for it in cand.items:
            k = _key(it)
            if k in by_key:
                # Collision: keep the item with the longer description.
                existing = by_key[k]
                if len(it.description) > len(existing.description):
                    by_key[k] = it
                continue
            by_key[k] = it
            order.append(k)

    merged_items = [by_key[k] for k in order]

    # Aggregate bbox across all contributing candidates whose items
    # actually made it into the merge (cheap: union of all bboxes).
    x0s = [c.legend_bbox[0] for c in ordered if c.legend_bbox[2] > 0]
    y0s = [c.legend_bbox[1] for c in ordered if c.legend_bbox[3] > 0]
    x1s = [c.legend_bbox[2] for c in ordered if c.legend_bbox[2] > 0]
    y1s = [c.legend_bbox[3] for c in ordered if c.legend_bbox[3] > 0]
    if x0s:
        agg_bbox = (min(x0s), min(y0s), max(x1s), max(y1s))
    else:
        agg_bbox = base.legend_bbox

    return LegendResult(
        items=merged_items,
        legend_bbox=agg_bbox,
        page_index=base.page_index,   # primary page
        columns_detected=base.columns_detected,
        detection_method=base.detection_method,
    )


def _should_aggregate(base: LegendResult, others: list[LegendResult]) -> bool:
    """
    Decide whether multi-page aggregation is worthwhile.  We only
    aggregate when another page contributes at least one NEW symbol that
    isn't already present in the densest (base) candidate.  This keeps
    legacy single-legend PDFs untouched.
    """
    base_syms = {it.symbol.strip().lower() for it in base.items if it.symbol}
    for other in others:
        for it in other.items:
            if not it.symbol:
                continue
            if it.symbol.strip().lower() not in base_syms:
                return True
    return False


def _legend_density_score(result: LegendResult) -> float:
    """
    S1.4 (T019): score a candidate LegendResult for multi-page selection.

    Density heuristic:
      * Primary: number of items with a non-empty symbol (the "symbol
        density" requested by the task).
      * Tie-break: total item count (fuller tables preferred).
      * Secondary tie-break: number of items with a non-empty description
        (guards against pure-symbol columns without meaningful text).

    Higher score = better candidate.
    """
    if not result.items:
        return 0.0
    sym_count = sum(1 for it in result.items if it.symbol)
    desc_count = sum(1 for it in result.items if it.description)
    total = len(result.items)
    # Weighted sum — sym count dominates, total is secondary, desc is third.
    return sym_count * 1000.0 + total * 10.0 + desc_count


def parse_legend(pdf_path: str) -> LegendResult:
    """
    Parse legend table(s) from a PDF file.

    S1.4 (T019): scans ALL pages, collects EVERY plausible legend
    candidate (both header-based and content-based), and returns the one
    with the highest symbol density (see `_legend_density_score`).  This
    ensures legends that live on page 2+ are preferred when earlier pages
    either have no legend or contain only a spurious small match.

    Backward compatibility: the return type is still a single
    ``LegendResult`` and consumers that relied on ``page_index`` /
    ``legend_bbox`` continue to work unchanged.  Aggregation across
    multiple pages is deferred to S002-05 (T020).

    Uses adaptive x_tolerance: tries x_tolerance=3 first (precise), and
    if the extracted descriptions are mostly empty, retries with
    x_tolerance=7 to handle PDFs where pdfplumber splits Cyrillic text
    into individual characters at low tolerance.
    """
    X_TOLERANCES = [3, 7]

    # Track pages where no header was found, for content-based fallback.
    no_header_pages: list[tuple[int, object]] = []

    # Collect every plausible legend candidate across ALL pages.
    # A candidate is kept even if descriptions look imperfect — the final
    # selection is done by density score after the sweep completes.
    all_candidates: list[LegendResult] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            lines = page.lines or []
            rects = page.rects or []

            best_page_result: Optional[LegendResult] = None
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
                    detection_method=_tag_method("header", page),
                )

                # Keep the best header-based result for THIS page across
                # x_tolerances.  Once the best for this page is chosen we
                # add it to the cross-page candidate pool.
                if best_page_result is None or \
                        _legend_density_score(result) > _legend_density_score(best_page_result):
                    best_page_result = result

                # Early-accept on this page only if the descriptions look
                # great — this avoids re-trying x_tol=7 unnecessarily but
                # does NOT short-circuit the multi-page sweep.
                if _has_good_descriptions(items):
                    break

            if best_page_result is not None and best_page_result.items:
                _stamp_page_index(best_page_result, page_idx)
                all_candidates.append(best_page_result)

            # Track pages for content-based fallback: either no header
            # found, or header found but no items extracted (secondary
            # pattern false positive).
            if not header_found_on_page or best_page_result is None \
                    or not best_page_result.items:
                no_header_pages.append((page_idx, page))

        # --- Content-based fallback on pages without a confident header ---
        # Even when header-based candidates exist elsewhere, the
        # content-based pass may discover denser legends on other pages,
        # so we always add any results it produces to the pool instead
        # of returning the first hit.
        for page_idx, page in no_header_pages:
            result = _content_based_legend_search(page, page_idx)
            if result is not None and result.items:
                _stamp_page_index(result, page_idx)
                all_candidates.append(result)

        # --- S1.7 (T022) Spec-table-as-legend fallback -----------------
        # Many DXF-derived PDFs (abk_em-like) carry their "legend" as a
        # conventional spec table with columns Поз. / Наименование /
        # Кол. (no legend header word).  Scan every page for such a
        # table and feed results into the candidate pool.  Running on
        # every page is safe — it's strictly additive because the
        # density-score selector will pick the densest candidate.
        for page_idx, page in enumerate(pdf.pages):
            spec_result = _spec_table_as_legend(page, page_idx)
            if spec_result is not None and spec_result.items:
                _stamp_page_index(spec_result, page_idx)
                all_candidates.append(spec_result)

        # --- S1.6 (T021) Last-resort symbol-density fallback -----------
        # If after header-based + content-based + spec-table passes we
        # still have zero candidates across the whole PDF, try the
        # purely geometric symbol-density detector on every page.  It's
        # best-effort, so we only activate it when nothing else
        # produced a legend.
        if not all_candidates:
            for page_idx, page in enumerate(pdf.pages):
                dens_result = _detect_legend_by_symbol_density(page, page_idx)
                if dens_result is not None and dens_result.items:
                    _stamp_page_index(dens_result, page_idx)
                    all_candidates.append(dens_result)

    # --- Select / aggregate candidates across all pages ---
    if not all_candidates:
        return LegendResult()

    # Prefer the highest density.  On ties, prefer the earliest page so
    # behaviour remains deterministic and close to the legacy semantics
    # when density is identical (original code returned the first match).
    best = max(
        all_candidates,
        key=lambda r: (_legend_density_score(r), -r.page_index),
    )

    # S1.5 (T020): if additional candidates from OTHER pages contribute
    # at least one fresh symbol, merge them all into one aggregate
    # LegendResult.  Otherwise return `best` unchanged — this preserves
    # byte-for-byte legacy behaviour on PDFs whose legend fits on a
    # single page.
    other_candidates = [
        c for c in all_candidates if c.page_index != best.page_index
    ]
    if other_candidates and _should_aggregate(best, other_candidates):
        return _merge_legend_candidates([best] + other_candidates)

    return best



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


def _find_identifier_cluster(
    words: list[dict],
    min_rows: int = 4,
    y_tol: float = 3.0,
) -> Optional[tuple[float, float, float, float, list[float]]]:
    """
    S1.3 (T018): locate a legend-like cluster of rows based on word content
    alone.  Returns (x0, y0, x1, y1, row_tops) for a run of >= ``min_rows``
    vertically-stacked rows whose LEFTMOST token looks like a legend
    identifier (``ROW_IDENT_RE``).

    Strategy:
      1. Group all words into rows by rounded "top" Y coordinate.
      2. For each row, pick the leftmost word as the row's identifier
         candidate.
      3. Find runs of consecutive rows whose identifier candidates match
         ``ROW_IDENT_RE`` and whose leftmost X values are close to each
         other (same vertical column of identifiers).
      4. Keep runs of at least ``min_rows`` rows and return the tightest
         bounding box around the full run.

    Returns ``None`` when no suitable cluster is found.
    """
    if not words:
        return None

    # --- 1. Group words into rows by rounded top Y -----------------------
    rows_by_y: dict[float, list[dict]] = {}
    for w in words:
        # Use a small tolerance bucket so that words that are a fraction of
        # a point apart still collapse into the same row.
        ky = round(w["top"] / y_tol) * y_tol
        rows_by_y.setdefault(ky, []).append(w)

    # Sort rows top-to-bottom.
    sorted_rows = sorted(rows_by_y.items(), key=lambda kv: kv[0])

    # --- 2. Build a list of (top, leftmost_word) pairs -------------------
    row_entries: list[tuple[float, dict, list[dict]]] = []
    for ky, row_words in sorted_rows:
        leftmost = min(row_words, key=lambda w: w["x0"])
        row_entries.append((ky, leftmost, row_words))

    # --- 3. Scan for runs of rows whose leftmost token matches -----------
    best_run: list[tuple[float, dict, list[dict]]] = []
    current_run: list[tuple[float, dict, list[dict]]] = []
    _IDENT_X_TOL = 12.0  # pt — identifiers must be roughly aligned
    _ROW_GAP_MAX = 40.0  # pt — tolerate modest vertical gaps between rows

    def _flush():
        nonlocal best_run
        if len(current_run) > len(best_run):
            best_run = list(current_run)

    for entry in row_entries:
        _ky, lw, _row_words = entry
        token = lw["text"]
        if not ROW_IDENT_RE.match(token):
            _flush()
            current_run = []
            continue

        if not current_run:
            current_run = [entry]
            continue

        prev_ky, prev_lw, _ = current_run[-1]
        same_column = abs(lw["x0"] - prev_lw["x0"]) <= _IDENT_X_TOL
        close_row = (entry[0] - prev_ky) <= _ROW_GAP_MAX

        if same_column and close_row:
            current_run.append(entry)
        else:
            _flush()
            current_run = [entry]

    _flush()

    if len(best_run) < min_rows:
        return None

    # --- 4. Compute the tightest bbox around the run ---------------------
    all_row_words: list[dict] = []
    for _ky, _lw, row_words in best_run:
        all_row_words.extend(row_words)

    x0 = min(w["x0"] for w in all_row_words)
    x1 = max(w["x1"] for w in all_row_words)
    y0 = min(w["top"] for w in all_row_words)
    y1 = max(w["bottom"] for w in all_row_words)
    row_tops = [ky for ky, _lw, _rw in best_run]
    return (x0, y0, x1, y1, row_tops)


_SPEC_HEADER_TOKENS = {
    # Position/identifier column (primary)
    "Поз.", "Поз", "Обозначение",
    # Name column (primary)
    "Наименование",
    # Quantity column (used only to anchor the header row when present)
    "Кол.", "Кол",
    # Notes column (optional)
    "Примечание",
}


def _spec_table_as_legend(
    page,
    page_idx: int,
    header_y_tol: float = 12.0,
) -> Optional[LegendResult]:
    """
    S1.7 (T022): detect a spec table on a drawing (DXF-derived PDFs)
    and treat it as a legend.

    Typical spec-table header shapes observed in abk_em drawings:
        "Поз. | Наименование | Кол. | Примечание"
        "Обозначение | Наименование | Примечание"
        "Поз. | Обозначение | Наименование | Кол."

    Strategy:
      1. Scan words for those matching ``_SPEC_HEADER_TOKENS``.
      2. Group header hits by Y (tol=``header_y_tol``).  A valid header
         row has >= 2 distinct header-kind tokens present; at minimum it
         must include one identifier-style column (``Поз.`` /
         ``Обозначение``) AND ``Наименование``.
      3. The header row's X positions define column bands:
           * id column  = from header.x0 - 20 until Наименование.x0 - 2
           * name column = from Наименование.x0 until next header.x0 - 2
      4. Collect words BELOW the header row; group by Y (tol=3pt) into
         data rows.  Stop at the first large Y-gap (> 60pt).
      5. For each data row: the leftmost word whose x0 falls inside the
         id column becomes ``LegendItem.symbol``; concatenated words in
         the name column become ``LegendItem.description``.  Rows with
         an empty symbol OR empty description are skipped.
      6. Return a ``LegendResult`` if at least 3 data rows were
         extracted.

    This fallback targets a different shape than the existing legend-
    header path: here the header cell reads "Поз." rather than any of
    the legend words in ``LEGEND_HEADER_PATTERNS``.
    """
    words_raw = page.extract_words(x_tolerance=3, y_tolerance=3) or []
    if not words_raw:
        return None
    for w in words_raw:
        w.setdefault("page_index", page_idx)

    # --- 1. header-token hits -------------------------------------------
    hits = [w for w in words_raw if w.get("text", "") in _SPEC_HEADER_TOKENS]
    if len(hits) < 2:
        return None

    # --- 2. group by Y band ---------------------------------------------
    by_y: dict[float, list[dict]] = {}
    for h in hits:
        ky = round(h["top"] / header_y_tol) * header_y_tol
        by_y.setdefault(ky, []).append(h)

    # Filter to candidate header rows: must include BOTH an identifier-
    # column token AND "Наименование".
    #
    # NB: we deliberately restrict id tokens to "Поз."/"Поз".  The
    # alternative "Обозначение" also matches LEGEND_HEADER_PATTERNS and
    # is already handled by the primary header-based path; more
    # importantly, in "Обозначение"-keyed tables the id column holds
    # graphical symbols (no text words) which would confuse the
    # leftmost-word-as-symbol heuristic below.
    ID_TOKENS = {"Поз.", "Поз"}
    candidate_rows: list[tuple[float, list[dict]]] = []
    for ky, row_hits in by_y.items():
        texts = {h["text"] for h in row_hits}
        has_id = bool(texts & ID_TOKENS)
        has_name = "Наименование" in texts
        if has_id and has_name:
            candidate_rows.append((ky, row_hits))

    if not candidate_rows:
        return None

    # Sort header candidates top-to-bottom and try EACH one; keep the
    # result with the most items.  A single page may carry multiple spec
    # tables (020-Узлы has three), and the topmost is not always the
    # most fruitful.
    candidate_rows.sort(key=lambda kv: kv[0])

    def _extract_one(header_row: list[dict]) -> Optional[LegendResult]:
        id_hdr = next(
            (h for h in header_row if h["text"] in {"Поз.", "Поз"}),
            None,
        )
        name_hdr = next(
            (h for h in header_row if h["text"] == "Наименование"),
            None,
        )
        if id_hdr is None or name_hdr is None:
            return None

        hdr_sorted = sorted(header_row, key=lambda w: w["x0"])
        name_idx = hdr_sorted.index(name_hdr)
        right_neighbour = (
            hdr_sorted[name_idx + 1] if name_idx + 1 < len(hdr_sorted) else None
        )

        id_x0 = id_hdr["x0"] - 20.0
        id_x1 = name_hdr["x0"] - 2.0
        name_x0 = name_hdr["x0"]
        name_x1 = (
            right_neighbour["x0"] - 2.0
            if right_neighbour is not None
            else name_hdr["x1"] + 200.0
        )

        header_bottom = max(h["bottom"] for h in header_row)

        data_words = [
            w for w in words_raw
            if w["top"] > header_bottom + 1
            and w["x0"] < name_x1 + 5
            and w["x1"] > id_x0 - 5
        ]
        if not data_words:
            return None

        row_bucket: dict[float, list[dict]] = {}
        for w in data_words:
            ky = round(w["top"] / 3.0) * 3.0
            row_bucket.setdefault(ky, []).append(w)
        sorted_row_keys = sorted(row_bucket)

        # Stop at the first large Y-gap (> 60pt) — ends the table.
        cut_idx = len(sorted_row_keys)
        for i in range(1, len(sorted_row_keys)):
            if sorted_row_keys[i] - sorted_row_keys[i - 1] > 60.0:
                cut_idx = i
                break
        data_row_keys = sorted_row_keys[:cut_idx]

        items: list[LegendItem] = []
        min_row_x = float("inf")
        max_row_x = float("-inf")
        max_row_y = None

        for ky in data_row_keys:
            row = row_bucket[ky]
            id_cell = [w for w in row if id_x0 - 1 <= w["x0"] < id_x1]
            name_cell = [w for w in row if name_x0 - 1 <= w["x0"] < name_x1]
            if not id_cell or not name_cell:
                continue
            id_cell.sort(key=lambda w: w["x0"])
            name_cell.sort(key=lambda w: w["x0"])
            symbol = id_cell[0]["text"].strip()
            description = " ".join(w["text"].strip() for w in name_cell).strip()
            if not symbol or not description:
                continue

            row_x0 = min(w["x0"] for w in row)
            row_x1 = max(w["x1"] for w in row)
            row_y0 = min(w["top"] for w in row)
            row_y1 = max(w["bottom"] for w in row)
            min_row_x = min(min_row_x, row_x0)
            max_row_x = max(max_row_x, row_x1)
            max_row_y = row_y1 if max_row_y is None else max(max_row_y, row_y1)

            items.append(LegendItem(
                symbol=symbol,
                description=description,
                bbox=(row_x0, row_y0, row_x1, row_y1),
                page_index=page_idx,
            ))

        if len(items) < 3:
            return None

        legend_x0 = min(id_hdr["x0"], min_row_x)
        legend_y0 = min(h["top"] for h in header_row)
        legend_x1 = max(name_x1, max_row_x)
        legend_y1 = max_row_y

        return LegendResult(
            items=items,
            legend_bbox=(legend_x0, legend_y0, legend_x1, legend_y1),
            page_index=page_idx,
            columns_detected=1,
            detection_method=_tag_method("spec", page),
        )

    best_result: Optional[LegendResult] = None
    for _ky, header_row in candidate_rows:
        r = _extract_one(header_row)
        if r is None:
            continue
        if best_result is None or len(r.items) > len(best_result.items):
            best_result = r

    return best_result


def _detect_legend_by_symbol_density(
    page,
    page_idx: int,
    min_markers: int = 4,
    row_y_tol: float = 3.0,
) -> Optional[LegendResult]:
    """
    S1.6 (T021): last-resort fallback that discovers legend-like regions
    purely from word geometry, with NO requirement for equipment
    vocabulary, table frames, or legend headers.

    Algorithm:
      1. Scan all words on the page; collect candidate "markers" — tokens
         matching ``ROW_IDENT_RE`` (e.g. ``1``, ``2``, ``3A``, ``5B``).
      2. For each marker, require adjacent text to the right on the same
         Y-row: at least one other word whose top is within ``row_y_tol``
         of the marker and whose x0 lies within a bounded X-window
         (marker.x1 .. marker.x1 + 400pt).  Markers that are isolated
         (no right-hand text) are discarded — this guards against grid
         axis labels, tick numbers and stray digits.
      3. Group surviving markers by X-column (tolerance 12pt) — a legend
         keeps all its identifiers in a single vertical column.
      4. Within each X-column, find the longest Y-contiguous run of
         markers (gap <= 40pt).  The run must have at least
         ``min_markers`` markers.
      5. The densest run wins; its bbox plus the adjacent right-hand
         text forms the returned region.

    Returns a ``LegendResult`` (or ``None``) ready to be fed into the
    usual extractor; this function synthesises ``TableBounds`` in the
    same shape that ``_content_based_legend_search`` produces.
    """
    words_raw = page.extract_words(x_tolerance=3, y_tolerance=3) or []
    for w in words_raw:
        w["page_index"] = page_idx

    if not words_raw:
        return None

    # --- 1. Candidate markers --------------------------------------------
    markers = [w for w in words_raw if ROW_IDENT_RE.match(w.get("text", ""))]
    if len(markers) < min_markers:
        return None

    # --- 2. Markers with adjacent right-hand text ------------------------
    _X_WINDOW = 400.0
    _RIGHT_MIN_GAP = 1.0
    anchored: list[dict] = []
    for m in markers:
        m_top = m["top"]
        m_x1 = m["x1"]
        for w in words_raw:
            if w is m:
                continue
            if abs(w["top"] - m_top) > row_y_tol:
                continue
            dx = w["x0"] - m_x1
            if _RIGHT_MIN_GAP <= dx <= _X_WINDOW:
                anchored.append(m)
                break

    if len(anchored) < min_markers:
        return None

    # --- 3. Group by X-column --------------------------------------------
    _COL_X_TOL = 12.0
    columns: dict[float, list[dict]] = {}
    for m in anchored:
        mx = m["x0"]
        placed = None
        for key in columns:
            if abs(key - mx) <= _COL_X_TOL:
                placed = key
                break
        if placed is None:
            placed = mx
        columns.setdefault(placed, []).append(m)

    # --- 4. Longest Y-contiguous run per column --------------------------
    _ROW_GAP_MAX = 40.0
    best_run: list[dict] = []
    for _col_x, col_markers in columns.items():
        col_markers.sort(key=lambda w: w["top"])
        current: list[dict] = []
        for m in col_markers:
            if not current:
                current = [m]
                continue
            if m["top"] - current[-1]["top"] <= _ROW_GAP_MAX:
                current.append(m)
            else:
                if len(current) > len(best_run):
                    best_run = list(current)
                current = [m]
        if len(current) > len(best_run):
            best_run = list(current)

    if len(best_run) < min_markers:
        return None

    # --- 5. Build a bbox that includes the markers AND their right-hand text
    run_tops = [m["top"] for m in best_run]
    y0 = min(run_tops)
    y1 = max(m["bottom"] for m in best_run)
    # Extend X right to include the widest right-hand text we can find
    # for any marker in the run.
    max_right = max(m["x1"] for m in best_run)
    for m in best_run:
        for w in words_raw:
            if w is m:
                continue
            if abs(w["top"] - m["top"]) > row_y_tol:
                continue
            dx = w["x0"] - m["x1"]
            if _RIGHT_MIN_GAP <= dx <= _X_WINDOW:
                if w["x1"] > max_right:
                    max_right = w["x1"]
    x0 = min(m["x0"] for m in best_run)
    x1 = max_right

    # --- 6. Synthesise TableBounds and extract items ---------------------
    # row_ys: one line 2pt above each marker's top + one line below the
    # last row, matching the shape expected by _extract_items_from_table.
    synth_row_ys = sorted(set(round(float(t) - 2.0, 1) for t in run_tops))
    synth_row_ys.append(round(float(y1) + 2.0, 1))

    # One column divider just to the right of the identifier column.
    ident_right = max(m["x1"] for m in best_run) + 2.0
    if ident_right >= x1:
        ident_right = x0 + (x1 - x0) * 0.25
    col_xs = [round(x0, 1), round(ident_right, 1), round(x1, 1)]

    bounds = TableBounds(
        x0=float(x0), x1=float(x1),
        y0=float(y0), y1=float(y1),
        col_xs=col_xs, row_ys=synth_row_ys,
    )

    items = _extract_items_from_table(
        words_raw, bounds,
        page_lines=page.lines or [],
        page_rects=page.rects or [],
    )

    if not items:
        return None

    # Guard against the grid-axis false positive: if NO item has any
    # description text at all, this is almost certainly a column of
    # standalone digits rather than a real legend.
    if not any(it.description.strip() for it in items):
        return None

    legend_bbox = (bounds.x0, bounds.y0, bounds.x1, bounds.y1)
    return LegendResult(
        items=items,
        legend_bbox=legend_bbox,
        page_index=page_idx,
        columns_detected=1,
        detection_method=_tag_method("density", page),
    )


def _content_based_legend_search(
    page,
    page_idx: int,
) -> Optional[LegendResult]:
    """
    Search for a legend table on a page using content-based heuristics
    when no legend header was found.

    Two-pass approach (S1.3 / T018):
      1. **Line-based pass** — detect rectangular tables via PDF line
         primitives.  The previous hard guard
         ``len(h_lines) < 3 or len(v_lines) < 2`` is removed; instead, if
         the line-based pass does not yield any candidate, we fall through
         to the content-cluster pass below.
      2. **Identifier-cluster pass** — locate a block of >= 4 stacked rows
         whose leftmost token matches ``ROW_IDENT_RE`` (e.g. ``1``, ``5A``,
         ``2Б``).  This recovers legends that lack a visible table frame —
         very common on GPC-style drawings where the legend is a typeset
         list rather than a drawn table.

    Returns a LegendResult if a likely legend is found, else None.
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

    # NOTE (T018): the hard early-return was removed here.  Even if the
    # page has very few detected lines we still attempt the line-based
    # pass (it will simply find no candidates and return None from its
    # own branch) and then fall back to the content-cluster pass.

    # Group horizontal lines by x0/x1 to find potential tables
    # A table is a group of 3+ horizontal lines with the same x0 and x1
    x_pairs: Counter = Counter()
    for l in h_lines:
        key = (round(l["x0"], 0), round(l["x1"], 0))
        x_pairs[key] += 1

    # Find the most frequent x0/x1 pair with at least 3 lines (= 2+ rows).
    # (S1.3) If there are no such candidates we simply skip the line-based
    # pass — we no longer bail out of the whole function here.
    candidates = [(pair, count) for pair, count in x_pairs.items() if count >= 3]

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
                detection_method=_tag_method("content", page),
            )

    # ------------------------------------------------------------------
    # S1.3 (T018) — Identifier-cluster fallback
    # ------------------------------------------------------------------
    # The line-based pass did not produce a usable legend.  Try to locate
    # a cluster of >= 4 vertically-stacked rows whose leftmost token is a
    # legend identifier (ROW_IDENT_RE).  This pass does not rely on any
    # PDF line primitives at all.
    cluster = _find_identifier_cluster(words_raw, min_rows=4)
    if cluster is None:
        return None

    cx0, cy0, cx1, cy1, row_tops = cluster

    # Sanity-check: require at least one equipment keyword in the cluster
    # so we don't accidentally match a drawing's position-numbering grid.
    cluster_words = [
        w for w in words_raw
        if cy0 - 2 < w["top"] < cy1 + 2
        and cx0 - 2 < w["x0"] < cx1 + 2
    ]
    equip_hits = 0
    for w in cluster_words:
        w_lower = w["text"].lower()
        for kw in _EQUIPMENT_KEYWORDS_LOWER:
            if kw in w_lower:
                equip_hits += 1
                break
    if equip_hits < 1:
        # No equipment vocabulary inside the cluster — probably a grid or
        # axis-label column, not a legend.
        return None

    # Build synthetic TableBounds for downstream extraction.
    # row_ys are synthesised as (row_top - 2) above each row plus a final
    # line just below the last row, so _extract_items_from_table sees a
    # well-formed row grid.
    synth_row_ys = sorted(set(round(float(t) - 2.0, 1) for t in row_tops))
    synth_row_ys.append(round(float(cy1) + 2.0, 1))
    # Synthesise a single internal column divider placed just to the
    # right of the identifier column.  Use the max x1 of any leftmost
    # identifier as an approximation.
    ident_right_edges = []
    _IDENT_X_TOL = 12.0
    # Identifiers are in a column aligned with x0 ≈ cx0; find their x1s.
    for w in cluster_words:
        if abs(w["x0"] - cx0) <= _IDENT_X_TOL and ROW_IDENT_RE.match(w["text"]):
            ident_right_edges.append(w["x1"])
    if ident_right_edges:
        col_div = max(ident_right_edges) + 2
    else:
        col_div = cx0 + 40
    col_xs = [round(cx0, 1), round(min(col_div, cx1 - 1), 1), round(cx1, 1)]

    bounds = TableBounds(
        x0=float(cx0), x1=float(cx1),
        y0=float(cy0), y1=float(cy1),
        col_xs=col_xs, row_ys=synth_row_ys,
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
            detection_method=_tag_method("content", page),
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
