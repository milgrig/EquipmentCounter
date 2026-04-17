"""
pdf_count_text.py — Count equipment by text markers on PDF drawings.

On ЭО (lighting) plans, each piece of equipment has a text marker next
to it: '5', '5А', '3', '1А', etc.  These are the same symbols from the
legend table.  This module counts occurrences of each symbol in the
drawing area, excluding the legend table, title block, and other
non-equipment text.

Approach:
  1. Accept pdf_path + legend result (from pdf_legend_parser)
  2. Extract all pdfplumber words with coordinates + color/font metadata
  3. Define exclusion zones (legend, title block, grid axes, spec tables, notes)
  4. For each legend symbol: find standalone word matches in drawing area
  5. Merge split text: '5' + 'А' nearby → '5А'
  6. Filter false positives (model name fragments, ЛК values, cables,
     wrong color, wrong size, text-dense areas)
  7. Return counts and positions

Usage:
    python pdf_count_text.py <path.pdf>
"""

from __future__ import annotations

import io
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

from pdf_legend_parser import parse_legend, LegendResult, LegendItem, _normalize_color, _classify_rgb

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class SymbolPosition:
    """A single found marker position."""
    symbol: str
    x: float
    y: float
    merged: bool = False  # True if merged from split text ('5' + 'А')


@dataclass
class CountResult:
    """Result of counting equipment markers."""
    counts: dict[str, int] = field(default_factory=dict)       # symbol → count
    positions: list[SymbolPosition] = field(default_factory=list)
    page_index: int = 0
    exclusion_zones: list[tuple[str, tuple[float, float, float, float]]] = \
        field(default_factory=list)  # [(name, (x0, y0, x1, y1)), ...]
    symbols_searched: list[str] = field(default_factory=list)
    total_words_on_page: int = 0
    words_in_drawing_area: int = 0


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Symbol regex — standalone equipment marker text
SYMBOL_RE = re.compile(r"^\d{1,2}[А-Яа-яЁё]{0,3}$")

# Cyrillic suffix for split-text merging
CYRILLIC_SUFFIX_RE = re.compile(r"^[А-Яа-яЁё]{1,3}$")

# Patterns to EXCLUDE from matching (not equipment markers)
# ЛК (lux) values: 300ЛК, 500лк
LK_RE = re.compile(r"\d+\s*[Лл][Кк]")
# Cable sections: 1х40, 3x2.5, 1х2 (both Cyrillic х and Latin x)
CABLE_RE = re.compile(r"\d+[хxХX]\d")
# Panel references: ЩАО1, ЩРО2, ЩСП-1
PANEL_RE = re.compile(r"^Щ[А-Яа-я]")
# Room/area numbers: typically 3+ digit numbers
ROOM_NUMBER_RE = re.compile(r"^\d{3,}$")
# Dimension/spec text that contains Latin chars nearby
LATIN_CHARS_RE = re.compile(r"[A-Za-z]")

# Max distance (pt) for merging split text (digit + Cyrillic suffix)
MERGE_MAX_DX = 12    # horizontal gap
MERGE_MAX_DY = 5     # vertical tolerance

# Grid axis margin (pt from page edge)
GRID_AXIS_MARGIN = 60

# Title block detection: minimum number of short H-lines in bottom-right
TITLE_BLOCK_MIN_LINES = 6

# Text-density filter: how many words within radius makes an area "text-dense"
TEXT_DENSE_RADIUS = 30      # pt — search radius around the digit
TEXT_DENSE_THRESHOLD = 8    # if ≥ this many words nearby, it's a text block

# Minimum text height fraction — standalone digits shorter than
# min_height_factor * median_ref_height are likely annotation artifacts
MIN_HEIGHT_FACTOR = 0.5


# ---------------------------------------------------------------------------
# Color helpers
# ---------------------------------------------------------------------------

def _word_color_label(word: dict) -> str:
    """
    Classify a word's fill color into a label ("red", "blue", "black", "grey", "").

    Uses non_stroking_color (text fill) from pdfplumber extra_attrs.
    Falls back to stroking_color if non_stroking_color is missing.
    """
    raw = word.get("non_stroking_color") or word.get("stroking_color")
    if raw is None:
        return ""
    rgb = _normalize_color(raw)
    if rgb is None:
        return ""
    return _classify_rgb(rgb)


def _build_symbol_color_map(legend_result: LegendResult) -> dict[str, str]:
    """
    Build a mapping from text symbol to its expected marker color.

    Returns e.g. {"1": "blue", "1А": "red", "5": "blue", "5А": "red"}.
    """
    sym_color: dict[str, str] = {}
    for item in legend_result.items:
        if item.symbol and item.color:
            sym_color[item.symbol] = item.color
    return sym_color


# ---------------------------------------------------------------------------
# Exclusion zone detection
# ---------------------------------------------------------------------------

def _detect_title_block(
    page,
    lines: list[dict],
) -> Optional[tuple[float, float, float, float]]:
    """
    Detect the title block (штамп) area from PDF lines.

    The title block is in the bottom-right corner, characterized by a dense
    grid of horizontal and vertical lines.

    Returns (x0, y0, x1, y1) or None.
    """
    pw, ph = page.width, page.height

    # Horizontal lines in the bottom-right quadrant
    h_lines = [
        l for l in lines
        if abs(l["top"] - l["bottom"]) < 2
        and abs(l["x1"] - l["x0"]) > 50
        and l["x0"] > pw * 0.4
        and l["top"] > ph * 0.6
    ]

    if len(h_lines) < TITLE_BLOCK_MIN_LINES:
        return None

    # Find the most common x0 among these lines → left edge of title block
    x0_counter: dict[float, int] = {}
    for l in h_lines:
        rx = round(l["x0"], 0)
        x0_counter[rx] = x0_counter.get(rx, 0) + 1

    # Take the x0 with the most lines
    best_x0 = max(x0_counter, key=x0_counter.get)  # type: ignore[arg-type]

    # Also find secondary x0 groups (title block may have multiple sections)
    all_x0s = sorted(x0_counter.keys())

    # Title block left edge = leftmost x0 that has ≥3 lines
    tb_x0 = pw
    for x in all_x0s:
        if x0_counter[x] >= 3 and x < tb_x0:
            tb_x0 = x

    # Right/bottom edges
    tb_x1 = max(l["x1"] for l in h_lines)
    tb_y0 = min(l["top"] for l in h_lines if abs(round(l["x0"], 0) - tb_x0) < 10
                or abs(round(l["x0"], 0) - best_x0) < 10)
    tb_y1 = max(l["top"] for l in h_lines)

    return (tb_x0, tb_y0, tb_x1, tb_y1)


def _detect_grid_axis_zones(
    page,
) -> list[tuple[str, tuple[float, float, float, float]]]:
    """
    Define exclusion zones for grid axis labels at page edges.

    Grid axes are typically:
    - Numbers in circles along the left/right edges
    - Letters in circles along the top/bottom edges
    """
    pw, ph = page.width, page.height
    m = GRID_AXIS_MARGIN

    zones = [
        ("grid_left", (0, 0, m, ph)),
        ("grid_right", (pw - m, 0, pw, ph)),
        ("grid_top", (0, 0, pw, m)),
        ("grid_bottom", (0, ph - m, pw, ph)),
    ]
    return zones


def _detect_spec_table_zones(
    page,
    words: list[dict],
    lines: list[dict],
    legend_bbox: tuple[float, float, float, float],
) -> list[tuple[str, tuple[float, float, float, float]]]:
    """
    Detect specification / schedule tables that are NOT the legend table.

    These tables contain quantities, model names, etc. whose digits should
    not be counted as equipment markers.  They typically appear near the
    legend table or in the bottom-right area of the drawing.

    Heuristic: look for dense clusters of horizontal lines that form
    table-like structures, excluding the already-detected legend and
    title block zones.
    """
    pw, ph = page.width, page.height
    zones: list[tuple[str, tuple[float, float, float, float]]] = []

    # Look for column-header keywords that indicate a specification table
    spec_keywords = {"Кол.", "Кол-во", "кол-во", "Масса", "масса",
                     "Примечание", "Количество", "Ед.", "Поз.",
                     "Марка", "Сечение", "Длина"}

    for w in words:
        if w.get("text", "") not in spec_keywords:
            continue
        # If this keyword is inside the legend bbox, skip
        if legend_bbox != (0, 0, 0, 0):
            if (legend_bbox[0] - 20 <= w["x0"] <= legend_bbox[2] + 20
                    and legend_bbox[1] - 20 <= w["top"] <= legend_bbox[3] + 20):
                continue

        # Found a spec-table keyword outside the legend.
        # Build a zone around it: look for nearby horizontal lines
        kw_x, kw_y = w["x0"], w["top"]
        nearby_h_lines = [
            l for l in lines
            if abs(l["top"] - l["bottom"]) < 2
            and abs(l["x1"] - l["x0"]) > 50
            and abs(l["top"] - kw_y) < 200
            and l["x0"] < kw_x + 300
            and l["x1"] > kw_x - 50
        ]
        if len(nearby_h_lines) >= 3:
            zone_x0 = min(l["x0"] for l in nearby_h_lines) - 5
            zone_y0 = min(l["top"] for l in nearby_h_lines) - 5
            zone_x1 = max(l["x1"] for l in nearby_h_lines) + 5
            zone_y1 = max(l["top"] for l in nearby_h_lines) + 5
            zones.append(("spec_table", (zone_x0, zone_y0, zone_x1, zone_y1)))

    return zones


def _detect_note_block_zones(
    page,
    words: list[dict],
) -> list[tuple[str, tuple[float, float, float, float]]]:
    """
    Detect note/annotation blocks (text-dense areas with sequential
    numbering or paragraph text).

    These typically appear in the bottom area of the drawing and contain
    text like "Примечание:", "Примечания:", numbered notes (1., 2., ...),
    or dense paragraphs that include digits in cross-references.
    """
    pw, ph = page.width, page.height
    zones: list[tuple[str, tuple[float, float, float, float]]] = []

    # Look for note header keywords
    note_headers = {"Примечание", "Примечания", "Примечания:", "Примечание:",
                    "ПРИМЕЧАНИЕ", "ПРИМЕЧАНИЯ", "Общие", "указания"}

    for w in words:
        if w.get("text", "") not in note_headers:
            continue
        # Found a note header.  Build a zone covering the note block area.
        # Notes are typically a vertical block of text below the header.
        hdr_x, hdr_y = w["x0"], w["top"]

        # Find all words below and nearby horizontally
        note_words = [
            nw for nw in words
            if nw["top"] >= hdr_y - 5
            and abs(nw["x0"] - hdr_x) < 400
            and nw["top"] < hdr_y + 300  # max 300pt height for a note block
        ]
        if len(note_words) >= 5:
            zone_x0 = min(nw["x0"] for nw in note_words) - 10
            zone_y0 = hdr_y - 10
            zone_x1 = max(nw["x1"] for nw in note_words) + 10
            zone_y1 = max(nw["bottom"] for nw in note_words) + 10
            zones.append(("note_block", (zone_x0, zone_y0, zone_x1, zone_y1)))

    return zones


# ---------------------------------------------------------------------------
# False positive filtering
# ---------------------------------------------------------------------------

def _is_near_latin_text(
    word: dict,
    all_words: list[dict],
    radius: float = 20,
) -> bool:
    """
    Check if a word is near Latin-character words (likely part of a model name).

    Equipment markers on Russian electrical drawings are standalone or near
    Cyrillic text. If surrounded by Latin chars, it's likely a model name
    fragment (e.g., '5' in 'SLICK.PRS LED 50').
    """
    wx, wy = word["x0"], word["top"]
    for w in all_words:
        if w is word:
            continue
        if abs(w["x0"] - wx) > radius and abs(w["x1"] - wx) > radius:
            continue
        if abs(w["top"] - wy) > 8:
            continue
        # Check if nearby word contains Latin chars
        if LATIN_CHARS_RE.search(w["text"]):
            return True
    return False


def _is_near_panel_ref(
    word: dict,
    all_words: list[dict],
    radius: float = 30,
) -> bool:
    """Check if digit is part of a panel reference (e.g., ЩАО1-Гр.5)."""
    wx, wy = word["x0"], word["top"]
    for w in all_words:
        if w is word:
            continue
        if abs(w["top"] - wy) > 8:
            continue
        if abs(w["x0"] - wx) > radius and abs(w["x1"] - wx) > radius:
            continue
        txt = w["text"]
        if txt.startswith("Щ") or "Гр." in txt or "Гр" == txt:
            return True
    return False


def _is_part_of_multi_digit(
    word: dict,
    all_words: list[dict],
) -> bool:
    """
    Check if a standalone digit is part of a multi-digit number that was
    split by pdfplumber (e.g., '1' '0' '0' for room number "100",
    or '2' '6' for "26").

    True markers are isolated digits. If another digit is immediately
    adjacent (within 5pt horizontally, same Y), this is likely a split
    multi-digit number.
    """
    wx1, wy = word["x1"], word["top"]
    wx0 = word["x0"]
    for w in all_words:
        if w is word:
            continue
        if abs(w["top"] - wy) > 4:
            continue
        # Another digit right next to this one (either side)
        if re.match(r"^\d$", w["text"]):
            # Immediately to the right
            if 0 <= (w["x0"] - wx1) < 5:
                return True
            # Immediately to the left
            if 0 <= (wx0 - w["x1"]) < 5:
                return True
    return False


def _is_lk_context(
    word: dict,
    all_words: list[dict],
) -> bool:
    """
    Check if digit word is part of an illumination annotation (e.g., '300 ЛК').

    Handles both combined ('300ЛК') and split ('3' '0' '0' 'Л' 'К') forms.
    All checks require the candidate word to be within 50pt horizontally.
    """
    wx1, wy = word["x1"], word["top"]
    wx0 = word["x0"]
    for w in all_words:
        if abs(w["top"] - wy) > 5:
            continue
        # Skip words that are too far horizontally
        if w["x0"] > wx1 + 50 or w["x1"] < wx0 - 50:
            continue
        # ЛК word nearby
        dx = w["x0"] - wx1
        if -5 < dx < 40 and re.match(r"^[Лл][Кк]?$", w["text"]):
            return True
        # Combined like '300ЛК' — only if this digit's word IS the combined word
        # or is immediately adjacent to it
        if LK_RE.search(w["text"]):
            return True
        # Split 'Л' then 'К' — check for standalone 'Л' nearby to the right
        if w["text"] == "Л" and 0 < (w["x0"] - wx0) < 50:
            return True
    return False


def _is_cable_context(word: dict, all_words: list[dict]) -> bool:
    """Check if digit is part of a cable cross-section (e.g., '3x2.5')."""
    wx, wy = word["x0"], word["top"]
    for w in all_words:
        if abs(w["top"] - wy) > 5:
            continue
        if abs(w["x0"] - wx) > 30:
            continue
        if CABLE_RE.search(w["text"]):
            return True
    return False


def _in_zone(
    word: dict,
    zone: tuple[float, float, float, float],
    margin: float = 2,
) -> bool:
    """Check if a word falls within a bounding box zone."""
    return (
        word["x0"] >= zone[0] - margin
        and word["x0"] <= zone[2] + margin
        and word["top"] >= zone[1] - margin
        and word["top"] <= zone[3] + margin
    )


def _is_in_text_dense_area(
    word: dict,
    all_words: list[dict],
    radius: float = TEXT_DENSE_RADIUS,
    threshold: int = TEXT_DENSE_THRESHOLD,
) -> bool:
    """
    Check if a digit is surrounded by many text words (notes, specs, annotations).

    Equipment markers are typically isolated — one digit next to a graphical
    symbol.  If a digit sits among 8+ other text words within 30pt, it is
    likely part of a paragraph or table cell, not an equipment marker.
    """
    wx, wy = word["x0"], word["top"]
    count = 0
    for w in all_words:
        if w is word:
            continue
        if abs(w["x0"] - wx) > radius and abs(w["x1"] - wx) > radius:
            continue
        if abs(w["top"] - wy) > radius:
            continue
        count += 1
        if count >= threshold:
            return True
    return False


def _is_near_cyrillic_text(
    word: dict,
    all_words: list[dict],
    radius: float = 15,
) -> bool:
    """
    Check if a digit is next to Cyrillic text on the same line.

    Equipment digit markers are standalone or near other digit markers.
    If a digit is adjacent to Cyrillic words (not a known suffix like 'А'),
    it's likely part of a text phrase like 'см. лист 5' or 'этаж 2'.
    """
    wx, wy = word["x0"], word["top"]
    wx1 = word["x1"]
    for w in all_words:
        if w is word:
            continue
        if abs(w["top"] - wy) > 6:
            continue
        # Word must be close horizontally
        if w["x0"] > wx1 + radius:
            continue
        if w["x1"] < wx - radius:
            continue
        text = w["text"]
        # Skip short Cyrillic suffixes that could be part of a valid compound
        if len(text) <= 3 and re.match(r"^[А-Яа-яЁё]{1,3}$", text):
            continue
        # If nearby word has 2+ Cyrillic chars, it's a text context
        if re.search(r"[А-Яа-яЁё]{2,}", text):
            return True
    return False


# ---------------------------------------------------------------------------
# Core counting logic
# ---------------------------------------------------------------------------

def count_symbols(
    pdf_path: str,
    legend_result: Optional[LegendResult] = None,
    equipment_zones: Optional[dict[str, list[tuple[float, float, float, float]]]] = None,
) -> CountResult:
    """
    Count equipment text markers on a PDF drawing page.

    Args:
        pdf_path: Path to the PDF file.
        legend_result: Pre-parsed legend result. If None, will parse automatically.
        equipment_zones: Optional per-color equipment cluster bboxes in PDF pt
            (from ``build_equipment_cluster_bboxes``).
            Format: ``{"red": [(x0,y0,x1,y1), ...], "blue": [...]}``.
            When provided, standalone digit markers are only accepted if
            they fall near an equipment cluster of the matching color.

    Returns:
        CountResult with counts per symbol and positions found.
    """
    # Parse legend if not provided
    if legend_result is None:
        legend_result = parse_legend(pdf_path)

    if not legend_result.items:
        return CountResult()

    # Collect text symbols from legend
    text_symbols: list[str] = []
    for item in legend_result.items:
        if item.symbol:
            text_symbols.append(item.symbol)

    if not text_symbols:
        # No text symbols in legend (e.g., ЭМ drawings — all graphical)
        return CountResult(symbols_searched=[], page_index=legend_result.page_index)

    # Build lookup: which compound symbols can be formed from digit + suffix
    # e.g., '5' + 'А' → '5А', '5' + 'АЭ' → '5АЭ'
    split_map: dict[str, dict[str, str]] = {}  # digit → {suffix → compound}
    for sym in text_symbols:
        m = re.match(r"^(\d{1,2})([А-Яа-яЁё]+)$", sym)
        if m:
            digit, suffix = m.group(1), m.group(2)
            if digit not in split_map:
                split_map[digit] = {}
            split_map[digit][suffix] = sym

    # Also build a set of pure-digit symbols (e.g., '1', '5') that we DO want
    pure_digit_symbols = {s for s in text_symbols if re.match(r"^\d{1,2}$", s)}

    # Build symbol → expected color mapping from legend
    symbol_color_map = _build_symbol_color_map(legend_result)

    page_idx = legend_result.page_index

    with pdfplumber.open(pdf_path) as pdf:
        if page_idx >= len(pdf.pages):
            return CountResult()

        page = pdf.pages[page_idx]

        # Extract words WITH extra attributes for color/font/size filtering.
        # extra_attrs makes pdfplumber group characters by their visual
        # attributes, so words of different colors stay separate.
        words_raw = page.extract_words(
            x_tolerance=3, y_tolerance=3,
            extra_attrs=["fontname", "size", "stroking_color",
                         "non_stroking_color"],
        ) or []
        lines = page.lines or []

        total_words = len(words_raw)

        # --- Build exclusion zones ---
        zones: list[tuple[str, tuple[float, float, float, float]]] = []

        # 1. Legend table
        lb = legend_result.legend_bbox
        if lb != (0, 0, 0, 0):
            # Add generous margin around legend
            zones.append(("legend", (lb[0] - 10, lb[1] - 30, lb[2] + 10, lb[3] + 10)))

        # 2. Title block
        tb = _detect_title_block(page, lines)
        if tb:
            zones.append(("title_block", tb))

        # 3. Grid axis margins
        zones.extend(_detect_grid_axis_zones(page))

        # 4. Specification tables (outside legend)
        zones.extend(_detect_spec_table_zones(
            page, words_raw, lines, legend_result.legend_bbox))

        # 5. Note/annotation blocks
        zones.extend(_detect_note_block_zones(page, words_raw))

        # --- Filter words to drawing area ---
        drawing_words: list[dict] = []
        for w in words_raw:
            excluded = False
            for _, zone_bbox in zones:
                if _in_zone(w, zone_bbox):
                    excluded = True
                    break
            if not excluded:
                drawing_words.append(w)

        # --- Find matches ---
        found_positions: list[SymbolPosition] = []
        used_word_indices: set[int] = set()  # track merged suffix words

        # Index words for fast neighbor lookup
        word_index = drawing_words  # simple list for now

        # Pass 1: Find exact compound symbol matches (e.g., '5А', '10А', '5АЭ')
        # These are highest confidence — no false positive risk
        compound_symbols = {s for s in text_symbols if re.search(r"[А-Яа-яЁё]", s)}
        for idx, w in enumerate(drawing_words):
            if w["text"] in compound_symbols:
                # Verify it's not a cable section or panel reference
                if CABLE_RE.search(w["text"]):
                    continue
                if PANEL_RE.search(w["text"]):
                    continue
                found_positions.append(SymbolPosition(
                    symbol=w["text"],
                    x=w["x0"],
                    y=w["top"],
                    merged=False,
                ))
                used_word_indices.add(idx)

        # Pass 2: Merge split text — digit word + nearby Cyrillic suffix
        digit_indices = [
            idx for idx, w in enumerate(drawing_words)
            if re.match(r"^\d{1,2}$", w["text"])
            and idx not in used_word_indices
        ]
        suffix_indices = [
            idx for idx, w in enumerate(drawing_words)
            if CYRILLIC_SUFFIX_RE.match(w["text"])
            and len(w["text"]) <= 3
        ]

        merged_digit_indices: set[int] = set()
        merged_suffix_indices: set[int] = set()

        for di in digit_indices:
            dw = drawing_words[di]
            digit_text = dw["text"]

            if digit_text not in split_map:
                continue

            possible_suffixes = split_map[digit_text]

            for si in suffix_indices:
                if si in merged_suffix_indices:
                    continue
                sw = drawing_words[si]

                # Suffix must be to the right and close
                dx = sw["x0"] - dw["x1"]
                dy = abs(sw["top"] - dw["top"])

                if 0 < dx < MERGE_MAX_DX and dy < MERGE_MAX_DY:
                    suffix_text = sw["text"]
                    if suffix_text in possible_suffixes:
                        compound = possible_suffixes[suffix_text]
                        found_positions.append(SymbolPosition(
                            symbol=compound,
                            x=dw["x0"],
                            y=dw["top"],
                            merged=True,
                        ))
                        merged_digit_indices.add(di)
                        merged_suffix_indices.add(si)
                        break  # one merge per digit word

        used_word_indices.update(merged_digit_indices)

        # --- Determine reference marker height from compound symbols ---
        # Compound symbols (e.g., '5А') are high-confidence matches.
        # Their text height = the expected height for equipment markers.
        ref_heights: list[float] = []
        for idx, w in enumerate(drawing_words):
            if idx in used_word_indices and w["text"] in compound_symbols:
                ref_heights.append(round(w["bottom"] - w["top"], 1))
        # For merged pairs, use the digit word height
        for di in merged_digit_indices:
            dw = drawing_words[di]
            ref_heights.append(round(dw["bottom"] - dw["top"], 1))

        # Determine acceptable height range for standalone digits
        if ref_heights:
            median_h = sorted(ref_heights)[len(ref_heights) // 2]
            min_marker_height = median_h * MIN_HEIGHT_FACTOR
            max_marker_height = median_h * 1.5  # allow 50% tolerance
        else:
            min_marker_height = 3.0  # fallback
            max_marker_height = 8.0  # fallback

        # --- Determine if color filtering is available ---
        # Color filtering is the strongest signal for standalone digits.
        # If legend items have color information, use it to reject
        # digits of the wrong color.
        has_color_info = bool(symbol_color_map)

        # Pass 3: Count standalone digit symbols (e.g., '1', '5')
        # These need careful false-positive filtering
        for di in digit_indices:
            if di in used_word_indices:
                continue  # already merged
            dw = drawing_words[di]
            digit_text = dw["text"]

            if digit_text not in pure_digit_symbols:
                continue

            # --- Height filters ---
            word_height = dw["bottom"] - dw["top"]

            # Filter: skip if text is too tall (room numbers, section labels)
            if word_height > max_marker_height:
                continue

            # Filter: skip if text is too short (tiny annotation artifacts)
            if word_height < min_marker_height:
                continue

            # --- Color filter (strongest signal) ---
            # If this digit symbol has a known color from the legend,
            # only accept words of that same color.
            if has_color_info and digit_text in symbol_color_map:
                expected_color = symbol_color_map[digit_text]
                actual_color = _word_color_label(dw)
                if actual_color and expected_color and actual_color != expected_color:
                    continue

            # --- Grey / very light text filter ---
            # Grey text (grid numbers, dimensions) is never an equipment marker.
            word_color = _word_color_label(dw)
            if word_color == "grey":
                continue

            # --- Equipment zone proximity filter ---
            # If we have colour-based equipment zones, check whether this
            # digit is near an equipment cluster of the matching colour.
            # A colour-matched digit near equipment is high-confidence,
            # so we can relax some of the softer contextual filters.
            near_color_equip = False
            if equipment_zones and digit_text in symbol_color_map:
                expected_clr = symbol_color_map[digit_text]
                zones_for_color = equipment_zones.get(expected_clr, [])
                if zones_for_color:
                    wx = float(dw["x0"])
                    wy = float(dw["top"])
                    for zx0, zy0, zx1, zy1 in zones_for_color:
                        if zx0 <= wx <= zx1 and zy0 <= wy <= zy1:
                            near_color_equip = True
                            break
                    if not near_color_equip:
                        continue

            # --- Contextual filters ---

            # Filter: skip if near Latin text (model name fragment)
            # Relax near coloured equipment — markers like "2" next to
            # "CD LED 27" are legitimate equipment labels.
            if not near_color_equip and _is_near_latin_text(dw, word_index):
                continue

            # Filter: skip if part of ЛК value
            if _is_lk_context(dw, word_index):
                continue

            # Filter: skip if part of cable section
            if _is_cable_context(dw, word_index):
                continue

            # Filter: skip room numbers (3+ digit standalone)
            if ROOM_NUMBER_RE.match(dw["text"]):
                continue

            # Filter: skip if near panel reference (ЩАО1-Гр.5)
            if _is_near_panel_ref(dw, word_index):
                continue

            # Filter: skip if part of a split multi-digit number
            if _is_part_of_multi_digit(dw, word_index):
                continue

            # Filter: skip if in a text-dense area (notes, specs, paragraphs)
            # Relax this filter when digit is near coloured equipment —
            # equipment labels often sit in annotation-dense areas.
            if not near_color_equip and _is_in_text_dense_area(dw, word_index):
                continue

            # Filter: skip if adjacent to Cyrillic text (sentence context)
            # Also relaxed near coloured equipment clusters.
            if not near_color_equip and _is_near_cyrillic_text(dw, word_index):
                continue

            found_positions.append(SymbolPosition(
                symbol=digit_text,
                x=dw["x0"],
                y=dw["top"],
                merged=False,
            ))

        # --- Build counts ---
        counts: dict[str, int] = {}
        for sym in text_symbols:
            counts[sym] = sum(1 for p in found_positions if p.symbol == sym)

        return CountResult(
            counts=counts,
            positions=found_positions,
            page_index=page_idx,
            exclusion_zones=zones,
            symbols_searched=text_symbols,
            total_words_on_page=total_words,
            words_in_drawing_area=len(drawing_words),
        )


# ---------------------------------------------------------------------------
# CLI interface
# ---------------------------------------------------------------------------

def main() -> None:
    """CLI entry point for testing."""
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print("Usage: python pdf_count_text.py <path.pdf>")
        sys.exit(1)

    pdf_path = args[0]
    print(f"Counting equipment markers: {pdf_path}")
    print()

    # Parse legend first
    legend = parse_legend(pdf_path)
    if not legend.items:
        print("No legend found — cannot count.")
        sys.exit(0)

    print(f"Legend: {len(legend.items)} items on page {legend.page_index + 1}")
    text_syms = [it.symbol for it in legend.items if it.symbol]
    print(f"Text symbols to count: {text_syms}")
    print()

    # Count
    result = count_symbols(pdf_path, legend)

    print(f"Page {result.page_index + 1}: {result.total_words_on_page} words total, "
          f"{result.words_in_drawing_area} in drawing area")
    print(f"Exclusion zones: {len(result.exclusion_zones)}")
    for name, bbox in result.exclusion_zones:
        print(f"  {name}: ({bbox[0]:.0f}, {bbox[1]:.0f}) — ({bbox[2]:.0f}, {bbox[3]:.0f})")
    print()

    # Results table
    print(f"{'Symbol':<10s} {'Count':>6s}   {'Merged':>6s}")
    print("-" * 30)
    total = 0
    for sym in result.symbols_searched:
        cnt = result.counts.get(sym, 0)
        merged = sum(1 for p in result.positions if p.symbol == sym and p.merged)
        direct = cnt - merged
        total += cnt
        merge_info = f"({direct}+{merged}m)" if merged else ""
        print(f"  {sym:<8s} {cnt:>6d}   {merge_info}")

    print("-" * 30)
    print(f"  {'TOTAL':<8s} {total:>6d}")
    print()

    # Show positions for each symbol
    if "--positions" in args or "-p" in args:
        print("Positions:")
        for sym in result.symbols_searched:
            sym_pos = [p for p in result.positions if p.symbol == sym]
            if sym_pos:
                print(f"\n  '{sym}' ({len(sym_pos)} found):")
                for p in sym_pos:
                    m = " [merged]" if p.merged else ""
                    print(f"    ({p.x:.1f}, {p.y:.1f}){m}")


if __name__ == "__main__":
    main()
