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
MIN_TABLE_LINE_LEN = 80

# Title block keywords that indicate we hit the title block, not legend
TITLE_BLOCK_KEYWORDS = {"Лист", "док.", "Подп.", "Дата", "Разраб.", "Пров.",
                         "Н.контр.", "ГИП", "Изм.", "Стадия"}


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
    """Check if text looks like title block content (not legend data)."""
    words_set = set(text.split())
    if words_set & TITLE_BLOCK_KEYWORDS:
        return True
    # Date pattern (dd.mm.yy or dd.mm.yyyy)
    if re.search(r"\d{2}\.\d{2}\.\d{2,4}", text):
        return True
    return False


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
    2. Find long vertical lines that span the same Y range.
    3. Determine table extents (x0, y0, x1, y1) and internal column dividers.
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
    # a reasonable X range (header_x ± 500 for left side)
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

    # Find the most common x0 among these horizontal lines → table left edge
    x0_counts: dict[float, int] = {}
    x1_counts: dict[float, int] = {}
    for l in candidate_h:
        rx0 = round(l["x0"], 0)
        rx1 = round(l["x1"], 0)
        x0_counts[rx0] = x0_counts.get(rx0, 0) + 1
        x1_counts[rx1] = x1_counts.get(rx1, 0) + 1

    table_x0 = max(x0_counts, key=x0_counts.get)  # type: ignore[arg-type]
    table_x1 = max(x1_counts, key=x1_counts.get)  # type: ignore[arg-type]

    # Filter to lines matching these x bounds (tolerance ±5)
    matched_h = [
        l for l in candidate_h
        if abs(round(l["x0"], 0) - table_x0) < 5
        and abs(round(l["x1"], 0) - table_x1) < 5
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
    Find the legend header ('Условные обозначения' or 'Легенда') on any page.
    Returns (y_top, x0, page_index) or None.
    """
    for w in words:
        for pattern in LEGEND_HEADER_PATTERNS:
            if pattern.search(w["text"]):
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


def _extract_items_from_table(
    words: list[dict],
    bounds: TableBounds,
    y_tol: float = 5,
) -> list[LegendItem]:
    """
    Extract legend items from within detected table boundaries.

    Uses table row lines (row_ys) to determine cell boundaries.
    All text within the same table cell (between consecutive row_ys)
    is merged into one item. This correctly handles multi-line
    descriptions within a single table row.
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
        desc_parts = [w for w in cell_words if w["x0"] >= sym_x_boundary]

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

        item = LegendItem(
            symbol=sym_text,
            description=desc_text,
            category=category,
            bbox=row_bbox,
        )
        items.append(item)

    return items


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def parse_legend(pdf_path: str) -> LegendResult:
    """
    Parse legend table(s) from a PDF file.

    Returns a LegendResult with all extracted items and the overall
    legend bounding box.
    """
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            words_raw = page.extract_words(x_tolerance=3, y_tolerance=3) or []
            # Annotate each word with page index
            for w in words_raw:
                w["page_index"] = page_idx

            # 1) DETECT legend header
            header_info = _find_legend_header(words_raw)
            if header_info is None:
                continue

            header_y, header_x, _ = header_info
            lines = page.lines or []

            # 2) DETERMINE table boundaries using PDF lines
            bounds = _find_table_bounds(lines, header_y, header_x, page.width)

            if bounds is None:
                # Fallback: use word-based heuristic
                bounds = _fallback_bounds(words_raw, header_y, header_x, page.width)
                if bounds is None:
                    continue

            # 3) EXTRACT items from the main table
            items = _extract_items_from_table(words_raw, bounds)

            # 4) Check for a second (GPC-style) table column
            columns_detected = 1
            second_bounds = _find_second_table(lines, words_raw, bounds, page.width)
            if second_bounds is not None:
                second_items = _extract_items_from_table(words_raw, second_bounds)
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

            return LegendResult(
                items=items,
                legend_bbox=legend_bbox,
                page_index=page_idx,
                columns_detected=columns_detected,
            )

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

    if oboz_x is None:
        return None

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
    if not area_words:
        return None

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
        print(f"  {i:3d}. {sym_display:>10s}  {item.description}{cat_display}")
        print(f"       bbox: ({item.bbox[0]:.1f}, {item.bbox[1]:.1f}, "
              f"{item.bbox[2]:.1f}, {item.bbox[3]:.1f})")

    print("-" * 80)


if __name__ == "__main__":
    main()
