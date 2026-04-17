"""
pdf_count_cables.py — Extract cable annotation data from PDF electrical drawings.

On ЭО (lighting) and ГПК plans, cable annotations appear as clusters
of text near cable routes.  Each annotation typically contains:
  - Cross-section: '3х1,5', '1х40', '5х2,5' (conductor_count х size)
  - Group label: 'ЩО1-Гр.7', 'ЩАО1-Гр.15А', 'ЦСАО3-Гр.34АЭ'
  - Cable brand: 'ППГнг(А)-HF', 'ВБШвнг(А)-LS'
  - Cable mm² / length: nearby numeric values

Two distinct annotation styles:
  A. Floor plan annotations: small text clusters near cable routes
  B. Schema columns: vertical stacks with reversed text (bottom-to-top)

This module finds cross-section words, searches nearby context, groups
by panel, and builds a cable schedule.

Usage:
    python pdf_count_cables.py <path.pdf> [--page N] [--all-pages]
"""

from __future__ import annotations

import io
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

from pdf_legend_parser import parse_legend, LegendResult

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class CableRun:
    """A single cable run annotation found on a drawing."""
    panel: str = ""                # Panel name, e.g., 'ЩО1', 'ЩАО1', 'ЦСАО3'
    group: str = ""                # Group label, e.g., 'Гр.7', 'Гр.15А'
    group_full: str = ""           # Full group label, e.g., 'ЩО1-Гр.7'
    cross_section: str = ""        # Cable cross-section, e.g., '3х1,5'
    length_m: Optional[float] = None  # Length in meters (if found)
    cable_type: str = ""           # Cable brand, e.g., 'ППГнг(А)-HF'
    position: tuple[float, float] = (0.0, 0.0)  # (x, y) position on page
    color: str = ""                # Wire color code (if found)
    page_index: int = 0            # Page where this annotation was found
    is_reversed: bool = False      # True if text was reversed (schema page)


@dataclass
class CableResult:
    """Result of cable annotation extraction."""
    runs: list[CableRun] = field(default_factory=list)
    # Summary
    total_runs: int = 0
    panels: dict[str, list[CableRun]] = field(default_factory=dict)  # panel → runs
    cable_schedule: list[dict] = field(default_factory=list)  # grouped summary
    # Metadata
    pages_scanned: list[int] = field(default_factory=list)
    total_cross_sections_found: int = 0
    total_group_labels_found: int = 0
    exclusion_zones: list[tuple[str, tuple[float, float, float, float]]] = \
        field(default_factory=list)


# ---------------------------------------------------------------------------
# Constants & patterns
# ---------------------------------------------------------------------------

# Cross-section pattern: NхM or NхM,D or NхM.D
# Uses Cyrillic х (U+0445) and Latin x, case-insensitive
CROSS_SECTION_RE = re.compile(
    r"^(\d{1,2})[хxХX](\d{1,4}(?:[,\.]\d{1,2})?)$"
)

# Reversed cross-section (schema pages): e.g., '5,1х3' for '3х1,5'
REVERSED_CROSS_SECTION_RE = re.compile(
    r"^(\d{1,4}(?:[,\.]\d{1,2})?)[хxХX](\d{1,2})$"
)

# Group label pattern: ЩXX-Гр.N or ЩXX-Гр.NА  (or ЦСАО3-Гр.34АЭ)
# Panel prefix: starts with Щ or Ц, followed by letters/digits
GROUP_FULL_RE = re.compile(
    r"^([ЩЦ][А-Яа-я0-9\.]+)[-–]Гр\.?(\d{1,3}[А-Яа-яЁё0-9]*)$"
)

# Split group: just 'Гр.N' or 'Гр.NA' without panel prefix
GROUP_PART_RE = re.compile(r"^Гр\.?(\d{1,3}[А-Яа-яЁё0-9]*)$")

# Panel name pattern (standalone)
PANEL_RE = re.compile(r"^([ЩЦ][А-Яа-я]*\d[\d\.]*)$")

# Cable brand patterns
CABLE_BRANDS = [
    "ВБШвнг",
    "ВБШв",
    "ППГнг",
    "ППГ",
    "АВВГ",
    "ВВГнг",
    "ВВГ",
    "КГ",
    "NYM",
]
# Full brand regex — may include suffix like (А)-HF, (А)-LS, (А)-FRHF, (А)-FRLS
CABLE_BRAND_RE = re.compile(
    r"^(" + "|".join(re.escape(b) for b in CABLE_BRANDS) + r")"
)

# Reversed brand patterns for schema pages (text is character-reversed)
# e.g., 'FH-)А(гнГПП' → 'ППГнг(А)-HF'
REVERSED_BRAND_TOKENS = [
    "гнГПП",   # ППГнг reversed
    "гнвШБВ",  # ВБШвнг reversed
    "гнГВВ",   # ВВГнг reversed
    "ГВВA",    # АВВГ reversed
]
REVERSED_BRAND_RE = re.compile(
    r"(" + "|".join(re.escape(t) for t in REVERSED_BRAND_TOKENS) + r")"
)

# Length pattern in reversed text: мNN=L → L=NNм
REVERSED_LENGTH_RE = re.compile(r"^м(\d+(?:[,\.]\d+)?)=L$")
# Normal length pattern: L=NNм or L=NN,Nм
NORMAL_LENGTH_RE = re.compile(r"^L=(\d+(?:[,\.]\d+)?)м$")

# Numeric value (potential mm² or length): small float like 2.5, 3,5, 12.5
NUMERIC_VALUE_RE = re.compile(r"^(\d{1,3}(?:[,\.]\d{1,2})?)$")

# Cable tray dimension filter: both sides > 50 → NOT a cable cross-section
# (e.g., 600х800, 100х50х3000)
TRAY_DIMENSION_RE = re.compile(
    r"^\d+[хxХX]\d+([хxХX]\d+)?$"
)

# Paper format filter: A3x3, A4x3
PAPER_FORMAT_RE = re.compile(r"^[AА]\d[хxХX]\d$", re.IGNORECASE)

# Grid axis margin
GRID_AXIS_MARGIN = 60

# Title block detection
TITLE_BLOCK_MIN_LINES = 6

# Nearby search radii (pt)
NEARBY_GROUP_RADIUS = 120     # group labels can be farther away
NEARBY_BRAND_RADIUS = 60      # brand is usually close
NEARBY_LENGTH_RADIUS = 50     # length value close
NEARBY_MM2_RADIUS = 15        # mm² value immediately adjacent
NEARBY_SCHEMA_RADIUS = 60     # schema column vertical spacing
NEARBY_LINE_COLOR_RADIUS = 50  # radius for line color voting

# Cable color codes (from line stroking_color / text non_stroking_color)
COLOR_RED = (1.0, 0.0, 0.0)     # RED  → аварийное (emergency) lighting
COLOR_BLUE = (0.0, 0.0, 1.0)    # BLUE → рабочее (working) lighting
COLOR_TOLERANCE = 0.05           # tolerance for float comparison


# ---------------------------------------------------------------------------
# Exclusion zone detection (shared with pdf_count_text.py)
# ---------------------------------------------------------------------------

def _detect_title_block(
    page,
    lines: list[dict],
) -> Optional[tuple[float, float, float, float]]:
    """
    Detect the title block (штамп) area from PDF lines.
    Returns (x0, y0, x1, y1) or None.
    """
    pw, ph = page.width, page.height

    h_lines = [
        ln for ln in lines
        if abs(ln["top"] - ln["bottom"]) < 2
        and abs(ln["x1"] - ln["x0"]) > 50
        and ln["x0"] > pw * 0.4
        and ln["top"] > ph * 0.6
    ]

    if len(h_lines) < TITLE_BLOCK_MIN_LINES:
        return None

    x0_counter: dict[float, int] = {}
    for ln in h_lines:
        rx = round(ln["x0"], 0)
        x0_counter[rx] = x0_counter.get(rx, 0) + 1

    best_x0 = max(x0_counter, key=x0_counter.get)  # type: ignore[arg-type]
    all_x0s = sorted(x0_counter.keys())

    tb_x0 = pw
    for x in all_x0s:
        if x0_counter[x] >= 3 and x < tb_x0:
            tb_x0 = x

    tb_x1 = max(ln["x1"] for ln in h_lines)
    tb_y0 = min(
        ln["top"] for ln in h_lines
        if abs(round(ln["x0"], 0) - tb_x0) < 10
        or abs(round(ln["x0"], 0) - best_x0) < 10
    )
    tb_y1 = max(ln["top"] for ln in h_lines)

    return (tb_x0, tb_y0, tb_x1, tb_y1)


def _build_exclusion_zones(
    page,
    lines: list[dict],
    legend_bbox: Optional[tuple[float, float, float, float]] = None,
) -> list[tuple[str, tuple[float, float, float, float]]]:
    """Build exclusion zones: title block, grid margins, legend."""
    pw, ph = page.width, page.height
    m = GRID_AXIS_MARGIN
    zones: list[tuple[str, tuple[float, float, float, float]]] = []

    # Legend
    if legend_bbox and legend_bbox != (0, 0, 0, 0):
        zones.append((
            "legend",
            (legend_bbox[0] - 10, legend_bbox[1] - 30,
             legend_bbox[2] + 10, legend_bbox[3] + 10),
        ))

    # Title block
    tb = _detect_title_block(page, lines)
    if tb:
        zones.append(("title_block", tb))

    # Grid axes
    zones.extend([
        ("grid_left", (0, 0, m, ph)),
        ("grid_right", (pw - m, 0, pw, ph)),
        ("grid_top", (0, 0, pw, m)),
        ("grid_bottom", (0, ph - m, pw, ph)),
    ])

    return zones


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


def _word_excluded(
    word: dict,
    zones: list[tuple[str, tuple[float, float, float, float]]],
) -> bool:
    """Check if word is in any exclusion zone."""
    for _, zb in zones:
        if _in_zone(word, zb):
            return True
    return False


# ---------------------------------------------------------------------------
# Cross-section detection & filtering
# ---------------------------------------------------------------------------

def _is_valid_cross_section(text: str) -> bool:
    """
    Check if a word is a valid cable cross-section (not a tray dim, paper format, etc).
    """
    # Must match basic pattern
    if not re.search(r"\d+[хxХX]\d+", text):
        return False

    # Filter: paper format (A3x3, A4x3)
    if PAPER_FORMAT_RE.match(text):
        return False

    # Filter: contains parentheses → equipment spec like (595x595)
    if "(" in text or ")" in text:
        return False

    # Filter: 3-part dimensions → tray spec like 100х50х3000
    if len(re.findall(r"[хxХX]", text)) >= 2:
        return False

    # Parse the two numbers
    m = re.match(r"^(\d+)[хxХX](\d+(?:[,\.]\d+)?)$", text)
    if not m:
        return False

    n1 = float(m.group(1).replace(",", "."))
    n2_str = m.group(2).replace(",", ".")
    n2 = float(n2_str)

    # Cable cross-sections: first number = conductor count (1-7 typically),
    # second number = wire gauge (0.5, 0.75, 1, 1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95)
    # Tray dimensions: both sides typically ≥ 30 (e.g., 50х50, 100х50, 600х800)

    # Filter: conductor count too large (> 19) → likely tray dimension
    if n1 > 19:
        return False

    # Filter: wire gauge too large (> 95mm²) → likely tray dimension
    if n2 > 95:
        return False

    # Filter: both sides ≥ 30 → likely tray dimension (e.g., 50х50, 50х30)
    if n1 >= 30 and n2 >= 30:
        return False

    return True


def _normalize_cross_section(text: str) -> str:
    """Normalize cross-section to standard form with Cyrillic х."""
    # Replace Latin x/X with Cyrillic х
    result = re.sub(r"[xX]", "х", text)
    return result


def _reverse_text(text: str) -> str:
    """Reverse a string (for schema page reversed text)."""
    return text[::-1]


def _try_parse_reversed_cross_section(text: str) -> Optional[str]:
    """
    Try to parse reversed cross-section: '5,1х3' → '3х1,5'.
    Returns normalized form or None.
    """
    m = REVERSED_CROSS_SECTION_RE.match(text)
    if not m:
        return None

    size_part = m.group(1)  # e.g., '5,1'
    count_part = m.group(2)  # e.g., '3'

    # In reversed form, the count and size are swapped
    # '5,1х3' means original was '3х1,5'
    return f"{count_part}х{size_part}"


# ---------------------------------------------------------------------------
# Nearby word search
# ---------------------------------------------------------------------------

def _find_nearby_words(
    anchor: dict,
    all_words: list[dict],
    radius: float,
    exclude_indices: Optional[set[int]] = None,
) -> list[tuple[int, dict, float]]:
    """
    Find words within radius of anchor word.
    Returns list of (index, word, distance) sorted by distance.
    """
    ax = (anchor["x0"] + anchor.get("x1", anchor["x0"])) / 2
    ay = (anchor["top"] + anchor.get("bottom", anchor["top"])) / 2

    results = []
    for idx, w in enumerate(all_words):
        if exclude_indices and idx in exclude_indices:
            continue
        if w is anchor:
            continue

        wx = (w["x0"] + w.get("x1", w["x0"])) / 2
        wy = (w["top"] + w.get("bottom", w["top"])) / 2

        dist = ((ax - wx) ** 2 + (ay - wy) ** 2) ** 0.5
        if dist <= radius:
            results.append((idx, w, dist))

    results.sort(key=lambda t: t[2])
    return results


def _find_nearest_group(
    anchor: dict,
    all_words: list[dict],
    radius: float,
    exclude_indices: Optional[set[int]] = None,
) -> Optional[tuple[str, str, str, dict]]:
    """
    Find nearest group label near anchor word.
    Returns (panel, group_num, full_label, matched_word) or None.
    The matched_word is the pdfplumber word dict for color detection.
    """
    nearby = _find_nearby_words(anchor, all_words, radius, exclude_indices)

    for _, w, _ in nearby:
        text = w["text"]

        # Full group label: ЩО1-Гр.7
        m = GROUP_FULL_RE.match(text)
        if m:
            panel = m.group(1)
            group = m.group(2)
            return (panel, f"Гр.{group}", text, w)

        # Compound with hyphen — might be split differently
        # Handle merged panel-group: 'ЩО1-Гр.7' can appear as-is
        if re.match(r"^[ЩЦ]", text) and "Гр" in text:
            # Try to extract panel and group
            parts = re.split(r"[-–]", text, maxsplit=1)
            if len(parts) == 2:
                panel = parts[0]
                gm = re.match(r"Гр\.?(\d+[А-Яа-яЁё0-9]*)", parts[1])
                if gm:
                    return (panel, f"Гр.{gm.group(1)}", text, w)

    # Try two-word pattern: panel word + group word nearby
    # First find panel words
    for _, pw, pdist in nearby:
        panel_m = PANEL_RE.match(pw["text"])
        if not panel_m:
            continue
        panel_name = panel_m.group(1)

        # Look for group word near this panel word
        for _, gw, gdist in nearby:
            if gw is pw:
                continue
            gm = GROUP_PART_RE.match(gw["text"])
            if gm:
                group_num = gm.group(1)
                full = f"{panel_name}-Гр.{group_num}"
                return (panel_name, f"Гр.{group_num}", full, pw)

    return None


def _find_nearest_brand(
    anchor: dict,
    all_words: list[dict],
    radius: float,
    is_reversed: bool = False,
    exclude_indices: Optional[set[int]] = None,
) -> Optional[str]:
    """
    Find nearest cable brand near anchor word.
    Returns brand name or None.
    """
    nearby = _find_nearby_words(anchor, all_words, radius, exclude_indices)

    for _, w, _ in nearby:
        text = w["text"]

        # Normal brand
        if CABLE_BRAND_RE.match(text):
            return text

        # Reversed brand (schema pages)
        if is_reversed and REVERSED_BRAND_RE.search(text):
            return _reverse_text(text)

    return None


def _find_nearest_length(
    anchor: dict,
    all_words: list[dict],
    radius: float,
    is_reversed: bool = False,
    exclude_indices: Optional[set[int]] = None,
) -> Optional[float]:
    """
    Find nearest length value near anchor word.
    Returns length in meters or None.
    """
    nearby = _find_nearby_words(anchor, all_words, radius, exclude_indices)

    for _, w, _ in nearby:
        text = w["text"]

        # Reversed length: мNN=L → L=NNм
        if is_reversed:
            m = REVERSED_LENGTH_RE.match(text)
            if m:
                val = m.group(1).replace(",", ".")
                try:
                    return float(val)
                except ValueError:
                    pass

        # Normal length: L=NNм
        m = NORMAL_LENGTH_RE.match(text)
        if m:
            val = m.group(1).replace(",", ".")
            try:
                return float(val)
            except ValueError:
                pass

    return None


def _find_nearest_mm2(
    anchor: dict,
    all_words: list[dict],
    radius: float = NEARBY_MM2_RADIUS,
    exclude_indices: Optional[set[int]] = None,
) -> Optional[float]:
    """
    Find nearest mm² value (small float) near cross-section word.
    On floor plans, this appears as 2.5, 3,5, 1.5 immediately adjacent.
    """
    nearby = _find_nearby_words(anchor, all_words, radius, exclude_indices)

    for _, w, _ in nearby:
        text = w["text"]
        m = NUMERIC_VALUE_RE.match(text)
        if m:
            val_str = m.group(1).replace(",", ".")
            try:
                val = float(val_str)
                # mm² values are typically 0.5 to 50
                if 0.5 <= val <= 50:
                    return val
            except ValueError:
                pass

    return None


# ---------------------------------------------------------------------------
# Cable color detection
# ---------------------------------------------------------------------------

def _color_matches(color, target: tuple, tol: float = COLOR_TOLERANCE) -> bool:
    """Check if a color tuple matches a target within tolerance."""
    if not isinstance(color, tuple) or len(color) != len(target):
        return False
    return all(abs(c - t) < tol for c, t in zip(color, target))


def _detect_color_from_chars(
    anchor: dict,
    chars: list[dict],
    radius: float = 15,
) -> str:
    """
    Detect cable color from nearby character non_stroking_color.

    Group label chars (ЩО1-Гр.N) carry the cable circuit color:
    - RED (1,0,0) = аварийное (emergency) lighting
    - BLUE (0,0,1) = рабочее (working) lighting

    Returns 'red', 'blue', or '' if no color detected.
    """
    ax = anchor["x0"]
    ay = anchor["top"]

    for ch in chars:
        if abs(ch["x0"] - ax) > radius or abs(ch["top"] - ay) > 5:
            continue
        nsc = ch.get("non_stroking_color")
        if nsc is None:
            continue
        if _color_matches(nsc, COLOR_RED):
            return "red"
        if _color_matches(nsc, COLOR_BLUE):
            return "blue"

    return ""


def _detect_color_from_lines(
    anchor: dict,
    lines: list[dict],
    radius: float = NEARBY_LINE_COLOR_RADIUS,
) -> str:
    """
    Detect cable color by voting among nearby colored line segments.

    Cable route lines are drawn in RED or BLUE. We count colored lines
    within a radius of the annotation position and return the dominant color.

    Returns 'red', 'blue', or '' if no clear winner.
    """
    ax = (anchor["x0"] + anchor.get("x1", anchor["x0"])) / 2
    ay = (anchor["top"] + anchor.get("bottom", anchor["top"])) / 2

    red_count = 0
    blue_count = 0

    for ln in lines:
        sc = ln.get("stroking_color")
        if not isinstance(sc, tuple) or len(sc) != 3:
            continue

        # Check distance from line midpoint to anchor
        mx = (ln["x0"] + ln["x1"]) / 2
        my = (ln["top"] + ln["bottom"]) / 2
        if abs(mx - ax) > radius or abs(my - ay) > radius:
            continue

        if _color_matches(sc, COLOR_RED):
            red_count += 1
        elif _color_matches(sc, COLOR_BLUE):
            blue_count += 1

    if red_count > blue_count and red_count > 0:
        return "red"
    if blue_count > red_count and blue_count > 0:
        return "blue"
    return ""


def _detect_cable_color(
    cs_word: dict,
    group_word: Optional[dict],
    chars: list[dict],
    lines: list[dict],
) -> str:
    """
    Detect cable color using a dual strategy:
    1. Primary: text color of the group label chars (most reliable)
    2. Fallback: nearby line color voting

    Returns 'red', 'blue', or ''.
    """
    # Strategy 1: group label text color
    if group_word is not None:
        color = _detect_color_from_chars(group_word, chars)
        if color:
            return color

    # Strategy 2: nearby line colors around cross-section word
    color = _detect_color_from_lines(cs_word, lines)
    if color:
        return color

    return ""


# ---------------------------------------------------------------------------
# Page type detection
# ---------------------------------------------------------------------------

def _is_schema_page(words: list[dict]) -> bool:
    """
    Detect if this page is a schema (single-line diagram) page.
    Schema pages have many reversed text words and vertical text columns.

    Heuristics:
    - Contains reversed brand tokens (FH-)А(гнГПП, SL-)А(гнвШБВ)
    - Contains reversed length pattern (мNN=L)
    """
    reversed_count = 0
    for w in words:
        text = w["text"]
        if REVERSED_BRAND_RE.search(text):
            reversed_count += 1
        if REVERSED_LENGTH_RE.match(text):
            reversed_count += 1

    return reversed_count >= 3


# ---------------------------------------------------------------------------
# Core extraction logic
# ---------------------------------------------------------------------------

def extract_cables(
    pdf_path: str,
    legend_result: Optional[LegendResult] = None,
    pages: Optional[list[int]] = None,
) -> CableResult:
    """
    Extract cable annotation data from a PDF drawing.

    Args:
        pdf_path: Path to the PDF file.
        legend_result: Pre-parsed legend result (for exclusion zones).
        pages: Specific page indices to scan. If None, scans all pages.

    Returns:
        CableResult with cable runs, panel grouping, and schedule.
    """
    # Parse legend for exclusion zones if not provided
    if legend_result is None:
        try:
            legend_result = parse_legend(pdf_path)
        except Exception:
            legend_result = None

    legend_bbox = None
    legend_page = -1
    if legend_result and legend_result.items:
        legend_bbox = legend_result.legend_bbox
        legend_page = legend_result.page_index

    all_runs: list[CableRun] = []
    all_zones: list[tuple[str, tuple[float, float, float, float]]] = []
    pages_scanned: list[int] = []
    total_cs = 0
    total_gl = 0

    with pdfplumber.open(pdf_path) as pdf:
        scan_pages = pages if pages is not None else list(range(len(pdf.pages)))

        for page_idx in scan_pages:
            if page_idx >= len(pdf.pages):
                continue

            page = pdf.pages[page_idx]
            words_raw = page.extract_words(x_tolerance=3, y_tolerance=3) or []
            pdf_lines = page.lines or []
            page_chars = page.chars or []

            if not words_raw:
                continue

            pages_scanned.append(page_idx)

            # Build exclusion zones
            lb = legend_bbox if page_idx == legend_page else None
            zones = _build_exclusion_zones(page, pdf_lines, lb)

            # Store zones from legend page (most useful) or first page
            if page_idx == legend_page or not all_zones:
                all_zones = zones

            # Filter words to drawing area
            drawing_words = [
                w for w in words_raw
                if not _word_excluded(w, zones)
            ]

            if not drawing_words:
                continue

            # Detect page type (schema vs floor plan)
            is_schema = _is_schema_page(drawing_words)

            # --- Find all cross-section words ---
            cs_indices: list[int] = []
            used_indices: set[int] = set()

            for idx, w in enumerate(drawing_words):
                text = w["text"]

                if not _is_valid_cross_section(text):
                    continue

                cs_indices.append(idx)

            total_cs += len(cs_indices)

            # --- Count group labels (for stats) ---
            for w in drawing_words:
                if GROUP_FULL_RE.match(w["text"]):
                    total_gl += 1
                elif GROUP_PART_RE.match(w["text"]):
                    total_gl += 1

            # --- For each cross-section, build a CableRun ---
            for cs_idx in cs_indices:
                cs_word = drawing_words[cs_idx]
                cs_text = cs_word["text"]

                # Determine if this is a reversed cross-section
                cs_reversed = False
                cs_normalized = _normalize_cross_section(cs_text)

                if is_schema:
                    # Try reversed interpretation
                    rev = _try_parse_reversed_cross_section(cs_text)
                    if rev:
                        cs_normalized = rev
                        cs_reversed = True

                used_indices.add(cs_idx)

                # Search nearby for context
                search_radius = NEARBY_SCHEMA_RADIUS if is_schema else NEARBY_GROUP_RADIUS

                group_info = _find_nearest_group(
                    cs_word, drawing_words, search_radius, used_indices
                )
                brand = _find_nearest_brand(
                    cs_word, drawing_words,
                    NEARBY_SCHEMA_RADIUS if is_schema else NEARBY_BRAND_RADIUS,
                    is_reversed=is_schema,
                    exclude_indices=used_indices,
                )
                length = _find_nearest_length(
                    cs_word, drawing_words,
                    NEARBY_SCHEMA_RADIUS if is_schema else NEARBY_LENGTH_RADIUS,
                    is_reversed=is_schema,
                    exclude_indices=used_indices,
                )

                # mm² value (only on floor plans)
                mm2_val = None
                if not is_schema:
                    mm2_val = _find_nearest_mm2(
                        cs_word, drawing_words,
                        exclude_indices=used_indices,
                    )

                # Detect cable color (RED=emergency, BLUE=working)
                group_word = group_info[3] if group_info else None
                cable_color = _detect_cable_color(
                    cs_word, group_word, page_chars, pdf_lines,
                )

                # Build CableRun
                run = CableRun(
                    cross_section=cs_normalized,
                    position=(round(cs_word["x0"], 1), round(cs_word["top"], 1)),
                    page_index=page_idx,
                    is_reversed=cs_reversed,
                    color=cable_color,
                )

                if group_info:
                    run.panel = group_info[0]
                    run.group = group_info[1]
                    run.group_full = group_info[2]

                if brand:
                    run.cable_type = brand

                if length is not None:
                    run.length_m = length

                # Use mm² to enrich cross-section info if available
                if mm2_val is not None and "," not in cs_normalized and "." not in cs_normalized:
                    # Cross-section like '3х1' might actually be '3х1,5'
                    # If mm² nearby is 1.5, the full spec is '3х1,5'
                    # But only if the mm2_val makes sense as a decimal part
                    pass  # Keep as-is for now, mm² is supplementary data

                all_runs.append(run)

    # --- Group by panel ---
    panels: dict[str, list[CableRun]] = defaultdict(list)
    for run in all_runs:
        panel_key = run.panel if run.panel else "(без панели)"
        panels[panel_key].append(run)

    # --- Build cable schedule ---
    schedule = _build_cable_schedule(all_runs)

    return CableResult(
        runs=all_runs,
        total_runs=len(all_runs),
        panels=dict(panels),
        cable_schedule=schedule,
        pages_scanned=pages_scanned,
        total_cross_sections_found=total_cs,
        total_group_labels_found=total_gl,
        exclusion_zones=all_zones,
    )


# ---------------------------------------------------------------------------
# Cable schedule builder
# ---------------------------------------------------------------------------

def _build_cable_schedule(runs: list[CableRun]) -> list[dict]:
    """
    Build a cable schedule (summary table) from extracted runs.

    Groups by panel + group, aggregates cross-sections and lengths.
    """
    # Group by (panel, group_full)
    groups: dict[str, list[CableRun]] = defaultdict(list)
    for run in runs:
        key = run.group_full if run.group_full else f"{run.panel or '?'}-{run.cross_section}"
        groups[key].append(run)

    schedule: list[dict] = []
    for group_key, group_runs in sorted(groups.items()):
        # Aggregate
        cross_sections = list({r.cross_section for r in group_runs})
        brands = list({r.cable_type for r in group_runs if r.cable_type})
        lengths = [r.length_m for r in group_runs if r.length_m is not None]
        total_length = sum(lengths) if lengths else None

        colors = sorted({r.color for r in group_runs if r.color})

        entry = {
            "group": group_key,
            "panel": group_runs[0].panel,
            "cross_sections": cross_sections,
            "cable_types": brands,
            "run_count": len(group_runs),
            "total_length_m": round(total_length, 1) if total_length else None,
            "pages": sorted({r.page_index for r in group_runs}),
            "colors": colors,
        }
        schedule.append(entry)

    return schedule


# ---------------------------------------------------------------------------
# CLI interface
# ---------------------------------------------------------------------------

def main() -> None:
    """CLI entry point for testing."""
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", errors="replace"
        )

    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print("Usage: python pdf_count_cables.py <path.pdf> [--page N] [--all-pages]")
        print()
        print("Options:")
        print("  --page N      Scan specific page (0-based)")
        print("  --all-pages   Scan all pages (default: legend page only)")
        print("  --detail      Show position details for each run")
        sys.exit(1)

    pdf_path = args[0]
    scan_pages = None
    show_detail = "--detail" in args or "-d" in args

    # Parse --page N
    for i, arg in enumerate(args[1:], 1):
        if arg == "--page" and i + 1 < len(args):
            try:
                scan_pages = [int(args[i + 1])]
            except ValueError:
                print(f"Invalid page number: {args[i + 1]}")
                sys.exit(1)

    # Default: scan all pages if --all-pages, else legend page
    all_pages_mode = "--all-pages" in args

    print(f"Cable extraction: {pdf_path}")
    print()

    # Parse legend
    legend = None
    try:
        legend = parse_legend(pdf_path)
        if legend.items:
            print(f"Legend: {len(legend.items)} items on page {legend.page_index + 1}")
        else:
            print("No legend found")
    except Exception as e:
        print(f"Legend parse failed: {e}")

    print()

    # Determine pages to scan
    if scan_pages is None:
        if all_pages_mode:
            scan_pages = None  # all pages
        elif legend and legend.items:
            scan_pages = [legend.page_index]
        else:
            scan_pages = None  # all pages as fallback

    # Extract cables
    result = extract_cables(pdf_path, legend, scan_pages)

    print(f"Pages scanned: {[p + 1 for p in result.pages_scanned]}")
    print(f"Cross-sections found: {result.total_cross_sections_found}")
    print(f"Group labels found: {result.total_group_labels_found}")
    print(f"Cable runs extracted: {result.total_runs}")
    print()

    if result.exclusion_zones:
        print("Exclusion zones:")
        for name, bbox in result.exclusion_zones:
            print(f"  {name}: ({bbox[0]:.0f}, {bbox[1]:.0f}) — ({bbox[2]:.0f}, {bbox[3]:.0f})")
        print()

    # Panel summary
    if result.panels:
        print("=== Panels ===")
        for panel, runs in sorted(result.panels.items()):
            print(f"\n  {panel}: {len(runs)} cable runs")
            # Group by group label within panel
            by_group: dict[str, list[CableRun]] = defaultdict(list)
            for r in runs:
                gk = r.group_full if r.group_full else "(без группы)"
                by_group[gk].append(r)

            for gk, gruns in sorted(by_group.items()):
                cs_list = ", ".join(sorted({r.cross_section for r in gruns}))
                brands = ", ".join(sorted({r.cable_type for r in gruns if r.cable_type})) or "—"
                lengths = [r.length_m for r in gruns if r.length_m is not None]
                length_str = f"{sum(lengths):.1f}м" if lengths else "—"
                pages_str = ", ".join(str(p + 1) for p in sorted({r.page_index for r in gruns}))
                colors = sorted({r.color for r in gruns if r.color})
                color_str = f" [{'/'.join(colors)}]" if colors else ""
                print(f"    {gk}: {cs_list} | {brands} | {length_str}"
                      f" | стр. {pages_str}{color_str}")

    print()

    # Cable schedule
    if result.cable_schedule:
        print("=== Cable Schedule ===")
        print(f"{'Group':<30s} {'Panel':<12s} {'Cross-sect':<15s} "
              f"{'Brand':<20s} {'Runs':>5s} {'Length':>8s}")
        print("-" * 95)
        for entry in result.cable_schedule:
            cs_str = ", ".join(entry["cross_sections"])
            brand_str = ", ".join(entry["cable_types"]) or "—"
            length_str = f"{entry['total_length_m']:.1f}м" if entry["total_length_m"] else "—"
            print(f"  {entry['group']:<28s} {entry['panel']:<12s} {cs_str:<15s} "
                  f"{brand_str:<20s} {entry['run_count']:>5d} {length_str:>8s}")

    # Summary
    print()
    total_length = sum(
        r.length_m for r in result.runs if r.length_m is not None
    )
    brand_counts: dict[str, int] = defaultdict(int)
    for r in result.runs:
        if r.cable_type:
            brand_counts[r.cable_type] += 1

    print(f"=== Summary ===")
    print(f"  Total cable runs: {result.total_runs}")
    print(f"  Total length: {total_length:.1f}м" if total_length else "  Total length: —")
    print(f"  Unique panels: {len(result.panels)}")
    if brand_counts:
        print(f"  Cable types:")
        for brand, cnt in sorted(brand_counts.items(), key=lambda x: -x[1]):
            print(f"    {brand}: {cnt} runs")

    # Color summary
    color_counts: dict[str, int] = defaultdict(int)
    for r in result.runs:
        if r.color:
            color_counts[r.color] += 1
    if color_counts:
        color_labels = {"red": "аварийное", "blue": "рабочее"}
        print(f"  Cable colors:")
        for col, cnt in sorted(color_counts.items()):
            label = color_labels.get(col, col)
            print(f"    {col} ({label}): {cnt} runs")

    # Detailed positions
    if show_detail and result.runs:
        print()
        print("=== Detailed Positions ===")
        for i, run in enumerate(result.runs, 1):
            rev = " [reversed]" if run.is_reversed else ""
            length = f" L={run.length_m}м" if run.length_m else ""
            brand = f" [{run.cable_type}]" if run.cable_type else ""
            color = f" {run.color}" if run.color else ""
            print(f"  {i:3d}. {run.cross_section:<12s} {run.group_full:<25s}"
                  f" ({run.position[0]:.0f}, {run.position[1]:.0f})"
                  f" p.{run.page_index + 1}{rev}{length}{brand}{color}")


if __name__ == "__main__":
    main()
