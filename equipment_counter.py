#!/usr/bin/env python3
"""
Equipment Counter for Engineering PDF/DXF Schematics.

Reads engineering drawings (PDF or DXF), extracts the legend table
(Условные обозначения), counts each equipment symbol in the
drawing body, and outputs an equipment list with quantities.

Usage:
    python equipment_counter.py 1.pdf
    python equipment_counter.py 1.dxf
    python equipment_counter.py .                   # all PDF+DXF in folder
    python equipment_counter.py 1.dxf --csv out.csv
"""

import argparse
import logging
import csv
import json
import re
import sys
from collections import OrderedDict, Counter
from dataclasses import dataclass, asdict
from pathlib import Path

try:
    import pdfplumber

    _HAS_PDF = True
except ImportError:
    _HAS_PDF = False

try:
    import ezdxf
    from ezdxf.entities.acad_table import read_acad_table_content as _read_acad_table

    _HAS_DXF = True
except ImportError:
    _HAS_DXF = False

if not _HAS_PDF and not _HAS_DXF:
    sys.exit("Install at least one backend:\n  pip install pdfplumber   # for PDF\n  pip install ezdxf        # for DXF")

SYMBOL_RE = re.compile(r"^\d{1,2}[А-Яа-яЁё]{0,3}$")
GRID_LINE_RE = re.compile(r"^\d+(\s+\d+){3,}$")
SKIP_LINE_PATTERNS = [re.compile(r"^\d{4,}(\s+\d{4,})*$")]
CIRCUIT_VARIANT_RE = re.compile(r"^(\d+)(А|АЭ)$")

SUPPORTED_EXT = set()
if _HAS_PDF:
    SUPPORTED_EXT.add(".pdf")
if _HAS_DXF:
    SUPPORTED_EXT.update((".dxf", ".DXF"))

ELEVATION_RE = re.compile(
    r"на отм[-._ ]*([+-]?\d+[-.,]\d+)", re.IGNORECASE
)


# Encoding pairs for recovering garbled Cyrillic filenames.
# Each tuple is (misinterpreted_encoding, actual_encoding).
_ENCODING_RECOVERY_PAIRS = [
    ("cp857", "cp866"),     # Turkish OEM -> Russian OEM (superset of cp850)
    ("cp850", "cp866"),     # Western OEM -> Russian OEM
    ("cp1252", "cp1251"),   # Western Windows -> Russian Windows
    ("latin-1", "cp1251"),  # Latin-1 -> Russian Windows
    ("latin-1", "cp866"),   # Latin-1 -> Russian OEM
    ("iso-8859-15", "cp1251"),
    ("cp437", "cp866"),     # US OEM -> Russian OEM
]

_logger = logging.getLogger(__name__)


def _try_recover_cyrillic(text: str) -> str:
    """Try to recover garbled Cyrillic text by re-encoding through common pairs.

    When filenames are created on systems with different codepages,
    Cyrillic characters get mangled.  This function attempts to reverse
    the mangling by trying known encoding misinterpretation pairs.

    Returns the original text unchanged if no Cyrillic recovery succeeds.
    """
    # Already contains Cyrillic -- no recovery needed
    if any("Ѐ" <= ch <= "ӿ" for ch in text):
        return text
    # Only try recovery if text has chars in the high-byte range
    if not any(0x80 <= ord(ch) <= 0xFF for ch in text):
        return text
    for src_enc, dst_enc in _ENCODING_RECOVERY_PAIRS:
        try:
            raw = text.encode(src_enc)
            recovered = raw.decode(dst_enc)
            if any("Ѐ" <= ch <= "ӿ" for ch in recovered):
                _logger.debug("Recovered Cyrillic: %r -> %r (%s->%s)",
                              text, recovered, src_enc, dst_enc)
                return recovered
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return text


def _classify_filename(fl: str) -> str:
    """Core classification logic on a lowercased filename string."""
    if "розеточн" in fl:
        return "розетки"
    if "расположени" in fl or "расстановк" in fl:
        return "расположение"
    if "привязк" in fl:
        return "привязка"
    if "освещени" in fl:
        return "освещение"
    if "кабеленесущ" in fl:
        return "кабеленесущие"
    # T076: "схем" catches all schema variants including:
    #   Схемы ЩО, ЩАО (abk_eo), Схемы АБК РД (abk_em),
    #   Схемы РД (sklad), Схема ВРУ (kpp_30)
    if "схем" in fl:
        return "схема"
    if "общие данные" in fl:
        return "общие"
    if "опросн" in fl:
        return "опросные"
    if fl.startswith("со") and len(fl) < 10:
        return "спецификация"
    if "спецификац" in fl:
        return "спецификация"
    return "другое"


def classify_plan(filename: str) -> str:
    """Classify a DXF/DWG file by plan type based on its filename.

    First tries direct Cyrillic matching on the filename.  If that yields
    "другое" (unrecognised), attempts to recover garbled Cyrillic text
    by re-encoding through common Windows codepage misinterpretation pairs
    (e.g. cp850-displayed cp866 bytes) and re-classifies.

    Note: encoding recovery runs on the ORIGINAL filename (before lowering)
    because .lower() changes Latin-supplement byte values and breaks the
    codepage re-interpretation.
    """
    fl = filename.lower()
    result = _classify_filename(fl)
    # T076: debug-log every classification to help diagnose misclassified files
    _logger.debug("classify_plan: %r -> %s", filename, result)
    if result != "другое":
        return result
    # Try recovering garbled Cyrillic on the ORIGINAL (pre-lower) filename
    recovered = _try_recover_cyrillic(filename)
    if recovered != filename:
        result = _classify_filename(recovered.lower())
        _logger.debug("classify_plan: %r -> recovered %r -> %s",
                      filename, recovered, result)
    return result


def extract_elevation_str(filename: str) -> str | None:
    """Extract elevation string from filename (e.g. '+7.800')."""
    m = ELEVATION_RE.search(filename)
    if m:
        return m.group(1).replace(",", ".")
    if "кровл" in filename.lower():
        return "roof"
    return None


def extract_elevation_float(filename: str) -> float | None:
    """Extract elevation as float from filename (e.g. 7.8)."""
    m = ELEVATION_RE.search(filename)
    if m:
        raw = m.group(1).replace(",", ".").replace("-", ".")
        if raw.startswith("+"):
            raw = raw[1:]
        try:
            return float(raw)
        except ValueError:
            pass
    return None


@dataclass
class EquipmentItem:
    symbol: str
    name: str
    count: int = 0
    count_ae: int = 0


@dataclass
class CableItem:
    cable_type: str
    count: int = 0
    total_length_m: int = 0


# ===================================================================
#  DXF processing  (ezdxf)
# ===================================================================

def _strip_mtext_codes(s: str) -> str:
    r"""Remove DXF MTEXT formatting codes, keeping readable text.

    T080: Handles complex formatting found in real DXF schemas:
      - {\fISOCPEUR|b0|i1|c0;text} -- font codes (keep text)
      - {\C1;text} -- color codes (keep text)
      - {\C256;\c0;HF} -- multi-code braced groups (keep text after last ;)
      - \pxi-11.208,...; -- paragraph formatting codes
      - \pi0,l0,...; -- inline paragraph codes
      - \P -- paragraph break (replaced with newline)
      - \W, \H, \A, \L, \O, \T, etc. -- other formatting codes
      - Nested braces like {\C1;{\fFont;text}}
    """
    # Step 1: Handle braced font/color groups -- extract text content.
    # Repeat to handle nesting like {\C1;{\fFont;text}}.
    for _ in range(3):
        # {\fFont|b0|i1|c0;text} -> text
        s = re.sub(r"\{\\f[^;]*;([^}]*)\}", r"\1", s)
        # {\C###;text} -> text  (color code with braces)
        s = re.sub(r"\{\\C\d+;([^}]*)\}", r"\1", s)
        # Generic braced code: {\X...;text} -> text
        s = re.sub(r"\{\\[A-Za-z][^;]*;([^}]*)\}", r"\1", s)

    # Step 2: Handle standalone (non-braced) formatting codes
    # \P = paragraph break -> newline
    s = re.sub(r"\\P", "\n", s)
    # \fFont|...; -- font change (without braces)
    s = re.sub(r"\\f[^;]*;", "", s)
    # \pxi...; and \pi...; -- paragraph/tab codes
    s = re.sub(r"\\p[^;]*;", "", s)
    # \W, \H, \C, \A, etc. -- other single-letter codes with ;
    s = re.sub(r"\\[A-Za-z][^;]*;", "", s)
    # Remaining \L (underline toggle, no semicolon)
    s = s.replace("\\L", "")

    # Step 3: Remove remaining braces
    s = s.replace("{", "").replace("}", "")
    return s


def _clean_mtext(raw: str) -> str:
    """Strip DXF MTEXT formatting codes, keep readable text."""
    return _strip_mtext_codes(raw).strip()


def _get_mtext_entries(dxf_path: str) -> list[tuple[str, float, float]]:
    """Return [(cleaned_text, x, y), ...] for every MTEXT in modelspace."""
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    entries = []
    for e in msp.query("MTEXT"):
        raw = e.text
        clean = _clean_mtext(raw)
        if not clean:
            continue
        x, y = e.dxf.insert.x, e.dxf.insert.y
        entries.append((clean, x, y))
    return entries


def _get_text_entries(msp) -> list[tuple[str, float, float]]:
    """Return [(text, x, y), ...] for every TEXT entity in modelspace.

    TEXT entities are simpler than MTEXT — they contain plain strings
    without formatting codes.  Many Revit/AutoCAD exports use TEXT for
    symbol labels even when longer descriptions use MTEXT.
    """
    entries: list[tuple[str, float, float]] = []
    for e in msp.query("TEXT"):
        raw = e.dxf.get("text", "")
        if not raw or not raw.strip():
            continue
        x, y = e.dxf.insert.x, e.dxf.insert.y
        entries.append((raw.strip(), x, y))
    return entries


_LEGEND_SKIP = {
    "Условные обозначения", "Обозначение", "Наименование",
    "Примечание", "Обозначение на планах сети",
}

_LEGEND_NOISE_RE = re.compile(
    r"^(?:Щ[А-ЯЁ]*\d.*Гр\.|ЩР\d|РП\d|На отм|^\d+[хx×]\d)"
)

_LEGEND_DESC_MIN_LEN = 8

_NON_COUNTABLE_PREFIXES = (
    "Прокладка", "Проводка", "Кабельная трасса",
    "Групповая сеть", "Кабель ", "Распаечная",
    "Противопожарная кабельная",
    "Щит аварийного", "Щит рабочего",
)

_PANEL_PREFIXES = ("Щит",)

_REVIT_CIRCLE_RADIUS = 50
_REVIT_CIRCLE_RADIUS_TOL = 25      # P-002 fix: widened from ±2 to ±25 for real radii 30–80
_REVIT_CIRCLE_RADIUS_RANGE = (25, 85)  # P-002 fix: accept circles with radius in this range
_REVIT_ANNO_LAYER_KW = "ANNO"

# ── Grid axis detection ──────────────────────────────────────────────

_GRID_BLOCK_KW = ("ось", "оси", "сетк", "grid", "axis", "bubble", "head",
                  "m_grid", "координ")
_GRID_LAYER_KW = ("ОСИ", "ОСЕЙ", "GRID", "AXIS", "СЕТК", "S-GRID", "S_GRID",
                  "КООРДИН", "A-ANNO-GRID")


def _detect_grid_labels(msp) -> set[tuple[int, int]]:
    """Detect grid axis labels (1,2,3… / А,Б,В…) to exclude from counting.

    Returns a set of (round(x), round(y)) positions that are grid MTEXT.

    Three detection methods, applied in order:
      1. Layer keywords — MTEXT on layers named like 'Оси', 'Grid', etc.
      2. INSERT proximity — short MTEXT near axis-bubble blocks
      3. Spatial pairing — same digit at same X (or Y), far apart
         (grid labels always appear at both ends of a grid line)
    """
    grid_pos: set[tuple[int, int]] = set()

    # ── Find grid INSERT blocks (axis bubbles) ──
    grid_inserts: list[tuple[float, float]] = []
    for e in msp.query("INSERT"):
        bl = e.dxf.name.lower()
        if any(kw in bl for kw in _GRID_BLOCK_KW):
            grid_inserts.append((e.dxf.insert.x, e.dxf.insert.y))

    # ── Scan all MTEXT ──
    short_labels: list[tuple[str, float, float]] = []

    for e in msp.query("MTEXT"):
        clean = _clean_mtext(e.text)
        if not clean:
            continue
        fl = clean.split("\n")[0].strip()
        if len(fl) > 4:
            continue
        x, y = e.dxf.insert.x, e.dxf.insert.y

        # Method 1: layer keywords
        layer_up = e.dxf.layer.upper()
        if any(kw in layer_up for kw in _GRID_LAYER_KW):
            grid_pos.add((round(x), round(y)))
            continue

        # Method 2: near a grid INSERT block
        near_grid_block = False
        for gx, gy in grid_inserts:
            if abs(x - gx) < 2000 and abs(y - gy) < 2000:
                grid_pos.add((round(x), round(y)))
                near_grid_block = True
                break
        if near_grid_block:
            continue

        is_pure_digit = fl.isdecimal() and fl.isascii() and 1 <= int(fl) <= 50
        is_single_letter = len(fl) == 1 and fl.isalpha()
        if is_pure_digit or is_single_letter:
            short_labels.append((fl, x, y))

    # ── Method 3: spatial pairing ──
    # Grid axes come in pairs at opposite ends of each grid line:
    # same number at same X but very different Y  (vertical grid line)
    # same number at same Y but very different X  (horizontal grid line)
    from collections import defaultdict
    by_text: dict[str, list[tuple[float, float]]] = defaultdict(list)
    for fl, x, y in short_labels:
        by_text[fl].append((x, y))

    paired: set[tuple[int, int]] = set()
    for txt, pts in by_text.items():
        if len(pts) < 2:
            continue
        for i in range(len(pts)):
            for j in range(i + 1, len(pts)):
                x1, y1 = pts[i]
                x2, y2 = pts[j]
                same_x = abs(x1 - x2) < 1000 and abs(y1 - y2) > 5000
                same_y = abs(y1 - y2) < 1000 and abs(x1 - x2) > 5000
                if same_x or same_y:
                    paired.add((round(x1), round(y1)))
                    paired.add((round(x2), round(y2)))
    grid_pos |= paired

    # If we found paired labels, also catch any remaining unpaired labels
    # that are aligned with the paired ones (e.g. middle-of-line repeats)
    if paired:
        paired_xs = {p[0] for p in paired}
        paired_ys = {p[1] for p in paired}
        for fl, x, y in short_labels:
            rx, ry = round(x), round(y)
            if (rx, ry) in grid_pos:
                continue
            if any(abs(rx - px) < 500 for px in paired_xs):
                if any(abs(ry - py) < 500 for py in paired_ys):
                    grid_pos.add((rx, ry))

    if grid_pos:
        print(f"  [grid] Detected {len(grid_pos)} grid axis labels (excluded)")

    return grid_pos

_GEOMETRIC_SKIP_DESC = (
    "Прокладка", "Проводка", "Кабельная трасса",
    "Групповая сеть", "Кабель ",
)


_LEGEND_HEADER_VARIANTS = (
    "Условные обозначения",
    "УСЛОВНЫЕ ОБОЗНАЧЕНИЯ",
    "Усл. обозначения",
    "Условные  обозначения",   # double space variant
    "Условные\nобозначения",   # multiline variant
)


def _is_legend_header(text: str) -> bool:
    """Check if text contains a legend header in any known variant."""
    for variant in _LEGEND_HEADER_VARIANTS:
        if variant in text:
            return True
    # Normalized check: collapse whitespace and compare case-insensitively
    normalized = " ".join(text.split()).lower()
    if "условные обозначения" in normalized:
        return True
    return False


def _find_all_legend_bboxes(
    entries: list[tuple[str, float, float]],
) -> list[tuple[float, float, float, float]]:
    """Find ALL legend bounding boxes (one per 'Условные обозначения' header)."""
    headers: list[tuple[float, float]] = []
    for text, x, y in entries:
        fl = text.split("\n")[0].strip()
        if _is_legend_header(text) or _is_legend_header(fl):
            if not any(abs(x - hx) < 5000 and abs(y - hy) < 5000 for hx, hy in headers):
                headers.append((x, y))
    bboxes: list[tuple[float, float, float, float]] = []
    for hx, hy in headers:
        nearby = [(ex, ey) for _, ex, ey in entries
                  if abs(ex - hx) < 15000 and (hy - ey) < 15000 and (hy - ey) > -2000]
        if not nearby:
            continue
        xs = [p[0] for p in nearby]
        ys = [p[1] for p in nearby]
        bboxes.append((min(xs) - 2000, min(ys) - 2000, max(xs) + 5000, max(ys) + 2000))
    return bboxes


def _find_legend_bbox_generic(
    entries: list[tuple[str, float, float]],
) -> tuple[float, float, float, float] | None:
    """Find first legend bounding box (backward compat)."""
    bboxes = _find_all_legend_bboxes(entries)
    return bboxes[0] if bboxes else None


def _parse_dxf_legend_numbered(
    entries: list[tuple[str, float, float]],
    bbox: tuple[float, float, float, float],
) -> OrderedDict:
    """Build {symbol: description} for drawings with numbered symbols (1, 2, 1А, 4АЭ).
    Used primarily for ЭО (lighting) plans."""
    xmin, ymin, xmax, ymax = bbox

    all_desc: list[tuple[str, float, float]] = []
    sym_texts: list[tuple[str, float, float]] = []

    for text, x, y in entries:
        if not (xmin < x < xmax and ymin < y < ymax):
            continue
        fl = text.split("\n")[0].strip()
        if fl in _LEGEND_SKIP:
            continue
        if _LEGEND_NOISE_RE.match(fl):
            continue
        if SYMBOL_RE.match(fl) and len(fl) <= 4:
            sym_texts.append((fl, x, y))
        elif len(fl) >= _LEGEND_DESC_MIN_LEN:
            all_desc.append((fl, x, y))

    if sym_texts:
        sym_x_median = sorted(x for _, x, _ in sym_texts)[len(sym_texts) // 2]
        desc_texts = [
            (t, x, y) for t, x, y in all_desc
            if not any(t.startswith(p) for p in _NON_COUNTABLE_PREFIXES)
            and x > sym_x_median
        ]
    else:
        desc_texts = [
            (t, x, y) for t, x, y in all_desc
            if not any(t.startswith(p) for p in _NON_COUNTABLE_PREFIXES)
        ]
    if not desc_texts:
        desc_texts = all_desc

    if len(sym_texts) >= 3:
        sym_xs = sorted(x for _, x, _ in sym_texts)
        sx_median = sym_xs[len(sym_xs) // 2]
        sym_texts = [(t, x, y) for t, x, y in sym_texts if abs(x - sx_median) < 2000]

    if not sym_texts or not desc_texts:
        return OrderedDict()

    sym_ys = [y for _, _, y in sym_texts]
    if max(sym_ys) - min(sym_ys) < 300 and len(sym_texts) > 3:
        return OrderedDict()

    pairs = []
    for si, (_, _, sy) in enumerate(sym_texts):
        for di, (_, _, dy) in enumerate(desc_texts):
            pairs.append((abs(sy - dy), si, di))
    pairs.sort()

    matched_syms: set[int] = set()
    matched_descs: set[int] = set()
    sym_desc: dict[int, int] = {}

    if len(sym_texts) >= 3:
        sym_ys_sorted = sorted(y for _, _, y in sym_texts)
        gaps = [abs(sym_ys_sorted[i + 1] - sym_ys_sorted[i])
                for i in range(len(sym_ys_sorted) - 1)]
        gaps = [g for g in gaps if g > 20]
        if gaps:
            typical_gap = sorted(gaps)[len(gaps) // 2]
            max_match_dist = max(500, typical_gap * 1.2)
        else:
            max_match_dist = 1000
    else:
        max_match_dist = 1000

    for dist, si, di in pairs:
        if dist > max_match_dist:
            break
        if si in matched_syms or di in matched_descs:
            continue
        sym_desc[si] = di
        matched_syms.add(si)
        matched_descs.add(di)

    if len(matched_syms) < len(sym_texts) and len(matched_descs) < len(desc_texts):
        extended_dist = max_match_dist * 3
        for dist, si, di in pairs:
            if dist > extended_dist:
                break
            if si in matched_syms or di in matched_descs:
                continue
            sym_desc[si] = di
            matched_syms.add(si)
            matched_descs.add(di)

    legend = OrderedDict()
    for si in sorted(range(len(sym_texts)), key=lambda i: -sym_texts[i][2]):
        if si not in sym_desc:
            continue
        sym = sym_texts[si][0]
        legend[sym] = desc_texts[sym_desc[si]][0]

    return legend


def _parse_dxf_legend_blocks(
    msp,
    entries: list[tuple[str, float, float]],
    bbox: tuple[float, float, float, float],
) -> tuple[OrderedDict, dict[str, str]]:
    """Build {block_name: description} for drawings with graphical block symbols (ЭМ/ЭС).

    Returns (legend_ordered_dict, block_to_sym_map).
    legend has {display_label: description}, block_to_sym_map has {block_name: display_label}.
    """
    xmin, ymin, xmax, ymax = bbox

    desc_texts: list[tuple[str, float, float]] = []
    for text, x, y in entries:
        if not (xmin < x < xmax and ymin < y < ymax):
            continue
        fl = text.split("\n")[0].strip()
        if fl in _LEGEND_SKIP or len(fl) < _LEGEND_DESC_MIN_LEN:
            continue
        desc_texts.append((fl, x, y))

    block_inserts: list[tuple[str, float, float]] = []
    for e in msp.query("INSERT"):
        x, y = e.dxf.insert.x, e.dxf.insert.y
        if xmin < x < xmax and ymin < y < ymax:
            bname = e.dxf.name.lower()
            if "ось" in bname or "сетки" in bname or "помещен" in bname:
                continue
            block_inserts.append((e.dxf.name, x, y))

    if not block_inserts or not desc_texts:
        return OrderedDict(), {}

    pairs = []
    for bi, (_, _, by) in enumerate(block_inserts):
        for di, (_, _, dy) in enumerate(desc_texts):
            pairs.append((abs(by - dy), bi, di))
    pairs.sort()

    matched_blocks: set[int] = set()
    matched_descs: set[int] = set()
    block_desc: dict[int, int] = {}

    for dist, bi, di in pairs:
        if bi in matched_blocks or di in matched_descs:
            continue
        if dist > 3000:
            break
        block_desc[bi] = di
        matched_blocks.add(bi)
        matched_descs.add(di)

    legend = OrderedDict()
    block_map: dict[str, str] = {}
    label_idx = 1

    for bi in sorted(range(len(block_inserts)), key=lambda i: -block_inserts[i][2]):
        if bi not in block_desc:
            continue
        bname = block_inserts[bi][0]
        desc = desc_texts[block_desc[bi]][0]
        label = f"B{label_idx}"
        label_idx += 1
        legend[label] = desc
        block_map[bname] = label

    return legend, block_map


def _parse_dxf_legend_descriptive(
    entries: list[tuple[str, float, float]],
    bbox: tuple[float, float, float, float],
) -> list[tuple[str, float, float]]:
    """Extract pure description entries from legend when no symbols/blocks match.

    Returns [(description, x, y), ...] sorted by descending Y.
    """
    xmin, ymin, xmax, ymax = bbox
    descs: list[tuple[str, float, float]] = []
    for text, x, y in entries:
        if not (xmin < x < xmax and ymin < y < ymax):
            continue
        fl = text.split("\n")[0].strip()
        if fl in _LEGEND_SKIP:
            continue
        if SYMBOL_RE.match(fl) and len(fl) <= 4:
            continue
        if _LEGEND_NOISE_RE.match(fl):
            continue
        if len(fl) >= _LEGEND_DESC_MIN_LEN:
            descs.append((fl, x, y))
    return sorted(descs, key=lambda t: -t[2])


def _is_annotation_circle(circle) -> bool:
    """Check if a CIRCLE entity is an annotation circle (equipment marker).

    P-002 fix: Uses wider radius range and relaxed layer check.
    Accepts circles with radius 25–85 (real-world annotation sizes vary).
    Also accepts circles on non-ANNO layers if radius is in the typical range.
    """
    r = circle.dxf.radius
    rmin, rmax = _REVIT_CIRCLE_RADIUS_RANGE
    if not (rmin <= r <= rmax):
        return False
    layer_up = circle.dxf.layer.upper()
    # Primary: ANNO layer keyword
    if _REVIT_ANNO_LAYER_KW in layer_up:
        return True
    # Secondary: accept circles in annotation-like radius range on any layer
    # (many exports don't use ANNO layer naming)
    if 30 <= r <= 70:
        return True
    return False


def _count_plan_circles_by_color(
    msp,
    legend_bbox: tuple[float, float, float, float] | None,
) -> dict[int, int]:
    """Count annotation circles by DXF color, excluding legend.

    P-002 fix: Uses wider radius tolerance and accepts multiple colors.
    Typical Revit ЭМ color mapping:
      color 1 (red)  → Щиты
      color 5 (blue) → розетки / кабельные выводы / оборудование
    """
    xmin, ymin, xmax, ymax = legend_bbox if legend_bbox else (0, 0, 0, 0)
    counts: Counter = Counter()
    for c in msp.query("CIRCLE"):
        if not _is_annotation_circle(c):
            continue
        cx, cy = c.dxf.center.x, c.dxf.center.y
        if legend_bbox and xmin < cx < xmax and ymin < cy < ymax:
            continue
        color = c.dxf.get("color", 256)
        counts[color] += 1
    return dict(counts)


def _find_legend_bbox(entries: list[tuple[str, float, float]]) -> tuple[float, float, float, float] | None:
    """Return (xmin, ymin, xmax, ymax) of the legend area."""
    return _find_legend_bbox_generic(entries)


_SPECIAL_LABELS = {"ВЫХОД": 'Пиктограмма "Выход/Exit"'}


def _count_dxf_symbols(
    entries: list[tuple[str, float, float]],
    symbols: list[str],
    legend_bboxes: list[tuple[float, float, float, float]],
    grid_positions: set[tuple[int, int]] | None = None,
) -> tuple[dict[str, int], dict[str, int], dict[str, int], dict[str, int]]:
    """Count symbol MTEXT labels on the drawing, excluding legend areas and grid axes.
    Returns (exact_counts, ae_counts, unlisted_counts, special_counts)."""
    sym_set = set(symbols)
    exact = Counter()
    ae = Counter()
    unlisted = Counter()
    special = Counter()

    for text, x, y in entries:
        first_line = text.split("\n")[0].strip()

        if any(xmin < x < xmax and ymin < y < ymax
               for xmin, ymin, xmax, ymax in legend_bboxes):
            continue

        if grid_positions and (round(x), round(y)) in grid_positions:
            continue

        if first_line in _SPECIAL_LABELS:
            special[first_line] += 1
            continue

        if not SYMBOL_RE.match(first_line):
            continue
        if len(first_line) > 4:
            continue

        if first_line in sym_set:
            exact[first_line] += 1
        else:
            m = CIRCUIT_VARIANT_RE.match(first_line)
            if m and m.group(1) in sym_set:
                ae[m.group(1)] += 1
            else:
                unlisted[first_line] += 1

    return dict(exact), dict(ae), dict(unlisted), dict(special)


_PLAN_ANNOTATION_RE = re.compile(r"^(\d{1,3})\s*[-–—]\s*(.+)", re.MULTILINE)


def _enrich_annotations_with_legend(
    annotation_items: list[EquipmentItem],
    legend: OrderedDict,
) -> list[EquipmentItem]:
    """Match annotation short names to full legend descriptions, use annotation counts.

    Annotations carry authoritative designer counts (e.g. "133 - SLICK.PRS LED 30 5000K")
    while the legend has full descriptions. This merges both: full name + correct count.

    Matching constraints:
      - All 2+ digit numbers from annotation must appear in legend description
      - Text-word overlap must be >= 60%
      - Annotation must have >= 3 words for a match (too-short names stay unmatched)
    """
    legend_entries = list(legend.items())
    used_legend: set[int] = set()
    result = []

    for ann in annotation_items:
        ann_tokens = re.findall(r"[a-zA-Z]+|\d+", ann.name.lower())
        ann_words = set(ann_tokens)
        ann_numbers = {w for w in ann_words if w.isdigit() and len(w) >= 2}

        if len(ann_words) < 3:
            result.append(ann)
            continue

        best_idx = -1
        best_score = 0.0

        for li, (sym, desc) in enumerate(legend_entries):
            if li in used_legend:
                continue
            desc_words = set(re.findall(r"[a-zA-Z]+|\d+", desc.lower()))
            if ann_numbers and not ann_numbers.issubset(desc_words):
                continue
            overlap = len(ann_words & desc_words) / len(ann_words)
            if overlap > best_score:
                best_score = overlap
                best_idx = li

        if best_idx >= 0 and best_score >= 0.6:
            sym, desc = legend_entries[best_idx]
            used_legend.add(best_idx)
            result.append(EquipmentItem(
                symbol=sym, name=desc, count=ann.count, count_ae=0,
            ))
        else:
            result.append(ann)

    for li, (sym, desc) in enumerate(legend_entries):
        if li not in used_legend:
            if not any(desc.startswith(p) for p in _NON_COUNTABLE_PREFIXES):
                result.append(EquipmentItem(
                    symbol=sym, name=desc, count=0, count_ae=0,
                ))

    return result


def _extract_plan_annotations(
    entries: list[tuple[str, float, float]],
    legend_bboxes: list[tuple[float, float, float, float]],
) -> list[EquipmentItem]:
    """Extract equipment counts from Revit plan annotations like '5-INSEL LB_S LED G3...'.

    Deduplicates: same annotation text on different views → counted once.
    """
    seen_texts: dict[str, int] = {}

    for text, x, y in entries:
        in_legend = any(
            xmin < x < xmax and ymin < y < ymax
            for xmin, ymin, xmax, ymax in legend_bboxes
        )
        if in_legend:
            continue

        fl = text.split("\n")[0].strip()
        m = _PLAN_ANNOTATION_RE.match(fl)
        if not m:
            continue
        count = int(m.group(1))
        desc = m.group(2).strip()
        if count < 1 or count > 500 or len(desc) < 10:
            continue
        if desc.startswith(("На отм", "Щ", "ВРУ", "Гр.")):
            continue
        if desc not in seen_texts:
            seen_texts[desc] = count
        else:
            seen_texts[desc] = max(seen_texts[desc], count)

    items = []
    for i, (desc, count) in enumerate(seen_texts.items(), 1):
        items.append(EquipmentItem(symbol=f"P{i}", name=desc, count=count))
    return items


_CABLE_TYPE_RE = re.compile(
    r"((?:ВБШвнг|ВБбШвнг|ВВГнг|ППГнг|АВВГнг|КГнг|АПвПу|ПвПу)"
    r"(?:\([А-Яа-яA-Za-z]+\))?-[A-Z]+\s+\d+[хx×]\d+[\.,]?\d*"
    r"|"
    r"(?:ПуВВнг|ПуВВ|ПВС|ПВ[13]|ШВВП|ППВ|АПВ)"
    r"(?:\([А-Яа-яA-Za-z]+\))?(?:-[A-Z]+)?\s+\d+[хx×]\d+[\.,]?\d*)"
)
_CABLE_LENGTH_RE = re.compile(r"L\s*=\s*(\d+)")
# T065: Extended length regex — also matches L-N (dash) and L=Nм (with units)
_CABLE_LENGTH_EXT_RE = re.compile(r"L\s*[=\-]\s*(\d+)")


def _extract_cables_mtext(
    entries: list[tuple[str, float, float]],
) -> dict[str, CableItem]:
    """Extract cable lengths from standalone MTEXT entries."""
    result: dict[str, CableItem] = {}
    for text, _, _ in entries:
        cm = _CABLE_TYPE_RE.search(text)
        lm = _CABLE_LENGTH_EXT_RE.search(text)
        if cm and lm:
            ct = cm.group(1).replace("х", "×").replace("x", "×")
            length = int(lm.group(1))
            if ct not in result:
                result[ct] = CableItem(cable_type=ct)
            result[ct].count += 1
            result[ct].total_length_m += length
    return result


def _extract_cables_raw_dxf(dxf_path: str) -> dict[str, CableItem]:
    """Extract cable lengths from raw DXF text (ENTITIES-only).

    T079: Scans ONLY the ENTITIES section of the DXF file to avoid inflated
    counts from serialized table duplicates in OBJECTS section.
    If ENTITIES yields 0 cables, falls back to BLOCKS section.
    NEVER scans OBJECTS section.

    P-005 fix: Uses a two-pass strategy:
      Pass 1: Same-line scan (cable type + L= on one line) — original approach
      Pass 2: Multi-line scan with sliding window — catches updated DXF files
              where cable type and L= are on separate lines (wrapped MTEXT,
              TABLE entities, different DXF versions).

    Deduplicates by (cable_type, length) pair to avoid double-counting from
    multiple DXF representations of the same table cell.
    """
    def _scan_section(dxf_path: str, target_section: str) -> dict[str, CableItem]:
        """Scan a specific DXF section (ENTITIES or BLOCKS) for cable data.

        Args:
            dxf_path: Path to DXF file.
            target_section: Section name to scan — "ENTITIES" or "BLOCKS".
                            NEVER pass "OBJECTS".
        """
        result: dict[str, CableItem] = {}
        seen_pairs: set[tuple[str, int]] = set()

        def _add_cable(ct_raw: str, length: int) -> None:
            ct = ct_raw.replace("х", "×").replace("x", "×")
            pair = (ct, length)
            if pair in seen_pairs:
                return
            seen_pairs.add(pair)
            if ct not in result:
                result[ct] = CableItem(cable_type=ct)
            result[ct].count += 1
            result[ct].total_length_m += length

        try:
            in_section = False
            prev_line = ""
            # P-005 fix: Keep a sliding window of recent lines for multi-line matching
            recent_lines: list[str] = []
            _WINDOW_SIZE = 5  # look back up to 5 lines

            with open(dxf_path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    ls = line.strip()
                    # Detect section start: group code 2 followed by section name
                    if ls == target_section and prev_line.strip() == "2":
                        in_section = True
                        prev_line = line
                        recent_lines.clear()
                        continue
                    # Detect section end: group code 0 followed by ENDSEC
                    if in_section and ls == "ENDSEC" and prev_line.strip() == "0":
                        in_section = False
                        prev_line = line
                        continue
                    prev_line = line

                    if not in_section:
                        continue

                    # T080: Strip MTEXT formatting codes before matching
                    clean_line = _strip_mtext_codes(line)
                    recent_lines.append(clean_line)
                    if len(recent_lines) > _WINDOW_SIZE:
                        recent_lines.pop(0)

                    # Pass 1: Same-line match (original approach)
                    if ("L=" in clean_line or "L-" in clean_line):
                        cm = _CABLE_TYPE_RE.search(clean_line)
                        lm = _CABLE_LENGTH_EXT_RE.search(clean_line)
                        if cm and lm:
                            _add_cable(cm.group(1), int(lm.group(1)))
                            continue

                    # Pass 2: Multi-line match — if this line has L=,
                    # look back in the window for a cable type
                    if ("L=" in clean_line or "L-" in clean_line):
                        lm = _CABLE_LENGTH_EXT_RE.search(clean_line)
                        if lm:
                            length = int(lm.group(1))
                            # Search recent lines backward for cable type
                            for prev in reversed(recent_lines[:-1]):
                                cm = _CABLE_TYPE_RE.search(prev)
                                if cm:
                                    _add_cable(cm.group(1), length)
                                    break

                    # Pass 2b: If this line has a cable type,
                    # look back for L= (covers reversed order)
                    cm = _CABLE_TYPE_RE.search(clean_line)
                    if cm:
                        for prev in reversed(recent_lines[:-1]):
                            if ("L=" in prev or "L-" in prev):
                                lm = _CABLE_LENGTH_EXT_RE.search(prev)
                                if lm:
                                    _add_cable(cm.group(1), int(lm.group(1)))
                                    break

        except Exception:
            pass
        return result

    # T079: ENTITIES-only scan — preferred, avoids OBJECTS inflation
    entities_result = _scan_section(dxf_path, target_section="ENTITIES")
    ent_total = sum(c.total_length_m for c in entities_result.values())

    if ent_total > 0:
        return entities_result

    # Fallback: try BLOCKS section if ENTITIES yielded nothing
    blocks_result = _scan_section(dxf_path, target_section="BLOCKS")
    return blocks_result


def _extract_cables_mtext_multiline(
    entries: list[tuple[str, float, float]],
) -> dict[str, CableItem]:
    """P-005 fix: Extract cables from MTEXT entries with multi-line content.

    Updated DXF files may wrap cable type and L= into multi-line MTEXT.
    This function handles both same-line and cross-line matches within
    a single MTEXT entity.
    """
    result: dict[str, CableItem] = {}
    seen_pairs: set[tuple[str, int]] = set()

    for text, _, _ in entries:
        # First try same-line (original behavior)
        cm = _CABLE_TYPE_RE.search(text)
        lm = _CABLE_LENGTH_EXT_RE.search(text)
        if cm and lm:
            ct = cm.group(1).replace("х", "×").replace("x", "×")
            length = int(lm.group(1))
            pair = (ct, length)
            if pair not in seen_pairs:
                seen_pairs.add(pair)
                if ct not in result:
                    result[ct] = CableItem(cable_type=ct)
                result[ct].count += 1
                result[ct].total_length_m += length
            continue

        # Multi-line: search all lines in MTEXT for cable type,
        # then search for L= in the same or neighboring lines
        if "\n" not in text and "\\P" not in text:
            continue
        lines = text.replace("\\P", "\n").split("\n")
        found_types: list[str] = []
        found_lengths: list[int] = []
        for ln in lines:
            cm2 = _CABLE_TYPE_RE.search(ln)
            if cm2:
                found_types.append(cm2.group(1))
            lm2 = _CABLE_LENGTH_EXT_RE.search(ln)
            if lm2:
                found_lengths.append(int(lm2.group(1)))

        # If we found types and lengths separately, pair them
        if found_types and found_lengths:
            for i, ct_raw in enumerate(found_types):
                ct = ct_raw.replace("х", "×").replace("x", "×")
                length = found_lengths[i] if i < len(found_lengths) else found_lengths[-1]
                pair = (ct, length)
                if pair not in seen_pairs:
                    seen_pairs.add(pair)
                    if ct not in result:
                        result[ct] = CableItem(cable_type=ct)
                    result[ct].count += 1
                    result[ct].total_length_m += length

    return result


def _extract_cables_mtext_table(dxf_path: str) -> dict[str, CableItem]:
    """T065: Extract cables from MTEXT in ACAD_TABLE geometry blocks (*T blocks).

    Multi-sheet DWG schemas store cable journal data as ACAD_TABLE entities
    whose geometry lives in *T blocks.  Each *T block contains MTEXT entries
    arranged in a grid (rows by Y, columns by X).  Cable journal rows
    typically contain cable_type + cross_section + length.

    Handles two formats:
      - Inline: cable type and L= in the same MTEXT
        e.g. 'ППГнг(А)-HF 3х2,5 в гофре ΔU=0,18% L=13м'
      - Multi-MTEXT: cable type in one MTEXT, length in an adjacent MTEXT
        at the same Y coordinate (same table row)
        e.g. MTEXT1='ППГнг(А)-HF 3х2,5' MTEXT2='L=13'

    Deduplicates by (cable_type, length) pair.
    """
    result: dict[str, CableItem] = {}
    seen_pairs: set[tuple[str, int]] = set()

    def _add_cable(ct_raw: str, length: int) -> None:
        ct = ct_raw.replace("х", "×").replace("x", "×")
        pair = (ct, length)
        if pair in seen_pairs:
            return
        seen_pairs.add(pair)
        if ct not in result:
            result[ct] = CableItem(cable_type=ct)
        result[ct].count += 1
        result[ct].total_length_m += length

    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception:
        return result

    for block in doc.blocks:
        if not block.name.startswith("*T"):
            continue

        # Collect all MTEXT entries from this table block
        entries: list[tuple[str, float, float]] = []
        for ent in block:
            if ent.dxftype() != "MTEXT":
                continue
            try:
                raw = ent.text
                clean = _clean_mtext(raw)
                if not clean.strip():
                    continue
                x = float(ent.dxf.insert.x)
                y = float(ent.dxf.insert.y)
                entries.append((clean.strip(), x, y))
            except Exception:
                continue

        if not entries:
            continue

        # Check if this block has any cable-related content
        has_cable = any(_CABLE_TYPE_RE.search(t) for t, _, _ in entries)
        has_length = any(_CABLE_LENGTH_EXT_RE.search(t) for t, _, _ in entries)
        if not has_cable:
            continue

        # -- Inline extraction: cable type + L= in same MTEXT --
        for text, x, y in entries:
            cm = _CABLE_TYPE_RE.search(text)
            lm = _CABLE_LENGTH_EXT_RE.search(text)
            if cm and lm:
                _add_cable(cm.group(1), int(lm.group(1)))

        # -- Multi-MTEXT extraction: cable type and length in adjacent cells --
        if not has_length:
            continue

        # Group entries by Y coordinate (rows) with tolerance
        row_map: dict[int, list[tuple[str, float]]] = {}
        y_tolerance = 50  # Y tolerance for grouping into same row
        y_keys: list[float] = sorted(set(y for _, _, y in entries))

        # Build Y-cluster mapping
        y_cluster: dict[float, int] = {}
        cluster_id = 0
        for i, yk in enumerate(y_keys):
            if i == 0 or abs(yk - y_keys[i - 1]) > y_tolerance:
                cluster_id = i
            y_cluster[yk] = cluster_id

        for text, x, y in entries:
            cid = y_cluster[y]
            if cid not in row_map:
                row_map[cid] = []
            row_map[cid].append((text, x))

        # For each row, try to pair cable types with lengths
        for cid, cells in row_map.items():
            type_cells: list[tuple[str, float]] = []
            length_cells: list[tuple[int, float]] = []
            for text, x in cells:
                cm = _CABLE_TYPE_RE.search(text)
                lm = _CABLE_LENGTH_EXT_RE.search(text)
                if cm and not lm:
                    type_cells.append((cm.group(1), x))
                elif lm and not cm:
                    length_cells.append((int(lm.group(1)), x))

            if not type_cells or not length_cells:
                continue

            # Pair by nearest X coordinate
            type_cells.sort(key=lambda t: t[1])
            length_cells.sort(key=lambda t: t[1])

            for ct_raw, tx in type_cells:
                # Find closest length cell by X
                best_len = None
                best_dist = float("inf")
                for ln, lx in length_cells:
                    dist = abs(tx - lx)
                    if dist < best_dist:
                        best_dist = dist
                        best_len = ln
                if best_len is not None:
                    _add_cable(ct_raw, best_len)

    return result


def _extract_cables_all_blocks(dxf_path: str) -> dict[str, CableItem]:
    """T074: Extract cables from ALL block definitions, not just *T blocks.

    Multi-sheet DWG files store cable data in various block types:
      - *T blocks (ACAD_TABLE geometry) — already handled by _extract_cables_mtext_table
      - Named blocks (A$C..., etc.) — schema diagram blocks with cable labels
      - Layout blocks — paper space content

    This function scans every block definition for MTEXT/TEXT entities
    matching cable patterns with lengths.  It handles:
      - Inline: cable type + L= in the same MTEXT/TEXT entity
      - Multi-entity: cable type in one entity, L= in a nearby entity
        within the same block (matched by Y-proximity for table rows)

    Deduplicates by (cable_type, length) pair to avoid double-counting
    from multiple DXF representations.
    """
    result: dict[str, CableItem] = {}
    seen_pairs: set[tuple[str, int]] = set()

    def _add_cable(ct_raw: str, length: int) -> None:
        ct = ct_raw.replace("х", "×").replace("x", "×")
        pair = (ct, length)
        if pair in seen_pairs:
            return
        seen_pairs.add(pair)
        if ct not in result:
            result[ct] = CableItem(cable_type=ct)
        result[ct].count += 1
        result[ct].total_length_m += length

    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception:
        return result

    _SKIP_BLOCK_NAMES = {"*Model_Space", "*Paper_Space", "*Paper_Space0"}
    for block in doc.blocks:
        # Skip modelspace / paper space (handled by _get_mtext_entries)
        # Skip *T blocks (handled by _extract_cables_mtext_table)
        if block.name in _SKIP_BLOCK_NAMES or block.name.startswith("*T"):
            continue

        # Collect all MTEXT/TEXT entries from this block
        entries: list[tuple[str, float, float]] = []
        for ent in block:
            if ent.dxftype() == "MTEXT":
                try:
                    clean = _clean_mtext(ent.text)
                    if not clean.strip():
                        continue
                    x = float(ent.dxf.insert.x)
                    y = float(ent.dxf.insert.y)
                    entries.append((clean.strip(), x, y))
                except Exception:
                    continue
            elif ent.dxftype() == "TEXT":
                try:
                    text = ent.dxf.text
                    if not text or not text.strip():
                        continue
                    x = float(ent.dxf.insert[0])
                    y = float(ent.dxf.insert[1])
                    entries.append((text.strip(), x, y))
                except Exception:
                    continue

        if not entries:
            continue

        # Check if this block has any cable-related content
        has_cable = any(_CABLE_TYPE_RE.search(t) for t, _, _ in entries)
        if not has_cable:
            continue

        # -- Pass 1: Inline extraction (cable type + L= in same entity) --
        for text, x, y in entries:
            cm = _CABLE_TYPE_RE.search(text)
            lm = _CABLE_LENGTH_EXT_RE.search(text)
            if cm and lm:
                _add_cable(cm.group(1), int(lm.group(1)))

        # -- Pass 2: Multi-entity extraction (cable + length in nearby entities) --
        has_length = any(_CABLE_LENGTH_EXT_RE.search(t) for t, _, _ in entries)
        if not has_length:
            continue

        # Group entries by Y coordinate (rows) with tolerance
        y_tolerance = 50
        y_keys: list[float] = sorted(set(y for _, _, y in entries))

        y_cluster: dict[float, int] = {}
        cluster_id = 0
        for i, yk in enumerate(y_keys):
            if i == 0 or abs(yk - y_keys[i - 1]) > y_tolerance:
                cluster_id = i
            y_cluster[yk] = cluster_id

        row_map: dict[int, list[tuple[str, float]]] = {}
        for text, x, y in entries:
            cid = y_cluster[y]
            if cid not in row_map:
                row_map[cid] = []
            row_map[cid].append((text, x))

        for cid, cells in row_map.items():
            type_cells: list[tuple[str, float]] = []
            length_cells: list[tuple[int, float]] = []
            for text, x in cells:
                cm = _CABLE_TYPE_RE.search(text)
                lm = _CABLE_LENGTH_EXT_RE.search(text)
                if cm and not lm:
                    type_cells.append((cm.group(1), x))
                elif lm and not cm:
                    length_cells.append((int(lm.group(1)), x))

            if not type_cells or not length_cells:
                continue

            type_cells.sort(key=lambda t: t[1])
            length_cells.sort(key=lambda t: t[1])

            for ct_raw, tx in type_cells:
                best_len = None
                best_dist = float("inf")
                for ln, lx in length_cells:
                    dist = abs(tx - lx)
                    if dist < best_dist:
                        best_dist = dist
                        best_len = ln
                if best_len is not None:
                    _add_cable(ct_raw, best_len)

    return result




def _extract_cables_ezdxf_structured(dxf_path: str) -> dict[str, CableItem]:
    """T082: Unified ezdxf structured cable extraction.

    Single ezdxf.readfile() call, then extracts cables from ALL structured
    sources with global dedup by (cable_type, length) tuple:

      1. Modelspace MTEXT entities -- strip formatting, apply cable regex
      2. Modelspace TEXT entities -- apply cable regex directly
      3. ACAD_TABLE entities -- use read_acad_table_content() for cell text,
         strip formatting, apply cable regex
      4. Blocks INSERTed in modelspace -- scan MTEXT/TEXT in block defs
         (only blocks actually referenced by INSERT in modelspace)

    Returns dict keyed by normalized cable_type string.
    """
    result: dict[str, CableItem] = {}
    seen_pairs: set[tuple[str, int]] = set()

    def _add_cable(ct_raw: str, length: int) -> None:
        ct = ct_raw.replace("х", "×").replace("x", "×")
        pair = (ct, length)
        if pair in seen_pairs:
            return
        seen_pairs.add(pair)
        if ct not in result:
            result[ct] = CableItem(cable_type=ct)
        result[ct].count += 1
        result[ct].total_length_m += length

    def _scan_text(text: str) -> None:
        """Apply cable regex to cleaned text, handling both single-line
        and multi-line content."""
        # Single-line: cable type + L= on same line
        cm = _CABLE_TYPE_RE.search(text)
        lm = _CABLE_LENGTH_EXT_RE.search(text)
        if cm and lm:
            _add_cable(cm.group(1), int(lm.group(1)))
            return

        # Multi-line: cable type and L= on different lines within same entity
        if "\n" not in text:
            return
        text_lines = text.split("\n")
        found_types: list[str] = []
        found_lengths: list[int] = []
        for ln in text_lines:
            cm2 = _CABLE_TYPE_RE.search(ln)
            if cm2:
                found_types.append(cm2.group(1))
            lm2 = _CABLE_LENGTH_EXT_RE.search(ln)
            if lm2:
                found_lengths.append(int(lm2.group(1)))
        if found_types and found_lengths:
            for i, ct_raw in enumerate(found_types):
                length = found_lengths[i] if i < len(found_lengths) else found_lengths[-1]
                _add_cable(ct_raw, length)

    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception:
        return result

    msp = doc.modelspace()

    # -- Source 1: Modelspace MTEXT --
    for ent in msp.query("MTEXT"):
        try:
            clean = _clean_mtext(ent.text)
            if clean:
                _scan_text(clean)
        except Exception:
            continue

    # -- Source 2: Modelspace TEXT --
    for ent in msp.query("TEXT"):
        try:
            text = ent.dxf.get("text", "")
            if text and text.strip():
                _scan_text(text.strip())
        except Exception:
            continue

    # -- Source 3: ACAD_TABLE entities --
    for ent in msp.query("ACAD_TABLE"):
        try:
            rows = _read_acad_table(ent)
            for row in rows:
                for cell in row:
                    if not cell or not cell.strip():
                        continue
                    clean = _clean_mtext(cell)
                    if clean:
                        _scan_text(clean)
        except Exception:
            continue

    # -- Source 4: Blocks INSERTed in modelspace --
    # Collect all block names referenced by INSERT in modelspace
    inserted_names: set[str] = set()
    for ent in msp.query("INSERT"):
        try:
            inserted_names.add(ent.dxf.name)
        except Exception:
            continue

    _SKIP_BLOCKS = {"*Model_Space", "*Paper_Space", "*Paper_Space0"}
    for block in doc.blocks:
        if block.name in _SKIP_BLOCKS:
            continue
        # Only scan blocks that are INSERTed in modelspace
        if block.name not in inserted_names:
            continue
        for ent in block:
            try:
                if ent.dxftype() == "MTEXT":
                    clean = _clean_mtext(ent.text)
                    if clean:
                        _scan_text(clean)
                elif ent.dxftype() == "TEXT":
                    text = ent.dxf.get("text", "")
                    if text and text.strip():
                        _scan_text(text.strip())
            except Exception:
                continue

    return result


def extract_cables_dxf(dxf_path: str) -> list[CableItem]:
    """Extract cable data from a DXF file.

    T082: ezdxf structured parsing is the PRIMARY method.
    Uses a single ezdxf.readfile() call to query:
      - Modelspace MTEXT/TEXT entities
      - ACAD_TABLE cell content (via read_acad_table_content)
      - MTEXT/TEXT in blocks INSERTed in modelspace
    Global dedup by (cable_type, length) tuple.

    Falls back to raw DXF line scan + legacy ezdxf helpers
    only when the structured method finds nothing.
    """
    # -- PRIMARY: ezdxf structured extraction (T082) --
    cables_structured = _extract_cables_ezdxf_structured(dxf_path)
    struct_total = sum(c.total_length_m for c in cables_structured.values())

    # -- LEGACY: *T block scanner + all-blocks scanner --
    # These catch cables in *T blocks (ACAD_TABLE geometry) and
    # non-INSERTed blocks that the structured method skips.
    cables_legacy: dict[str, CableItem] = {}
    cables_table = _extract_cables_mtext_table(dxf_path)
    for ct, item in cables_table.items():
        if ct not in cables_legacy:
            cables_legacy[ct] = item
        elif item.total_length_m > cables_legacy[ct].total_length_m:
            cables_legacy[ct] = item

    cables_blocks = _extract_cables_all_blocks(dxf_path)
    for ct, item in cables_blocks.items():
        if ct not in cables_legacy:
            cables_legacy[ct] = item
        elif item.total_length_m > cables_legacy[ct].total_length_m:
            cables_legacy[ct] = item

    # Merge: structured is primary, supplement from legacy
    cables_ezdxf = dict(cables_structured)
    for ct, item in cables_legacy.items():
        if ct not in cables_ezdxf:
            cables_ezdxf[ct] = item
        elif item.total_length_m > cables_ezdxf[ct].total_length_m:
            cables_ezdxf[ct] = item

    ezdxf_total = sum(c.total_length_m for c in cables_ezdxf.values())

    # -- FALLBACK: raw DXF scan --
    cables_raw = _extract_cables_raw_dxf(dxf_path)
    raw_total = sum(c.total_length_m for c in cables_raw.values())

    if ezdxf_total == 0 and raw_total == 0:
        return []

    # ezdxf is primary -- use it when it found data
    if ezdxf_total > 0:
        cables = cables_ezdxf
        # Supplement with raw results for cable types ezdxf missed
        for ct, item in cables_raw.items():
            if ct not in cables:
                cables[ct] = item
    else:
        # ezdxf found nothing -- use raw scan as fallback
        cables = cables_raw

    return sorted(cables.values(), key=lambda c: -c.total_length_m)


# ── Specification table parser (ACAD_TABLE in СО.dxf) ────────────────

_MTEXT_CONTENT_RE = re.compile(r"\{\\f[^;]+;(.+?)(?:\}|$)")
_TABLE_SEPARATOR_VALUES = {"44", "172", "176", "91", "145", "290", "11",
                           "100", "62", "70", "40", "90"}


def _strip_mtext_formatting(raw: str) -> str:
    r"""Strip MTEXT formatting codes and return plain text (single-line).

    T080: Uses _strip_mtext_codes() for robust stripping, then converts
    newlines to spaces (for spec table cells where single-line output
    is needed).
    """
    s = _strip_mtext_codes(raw)
    # Spec table cells need single-line output: \P → space
    s = s.replace("\n", " ")
    return s.strip()


@dataclass
class SpecItem:
    """One row from the equipment specification table."""
    position: str
    description: str
    model: str
    catalog_code: str
    supplier: str
    unit: str
    quantity: int


def _extract_table_cells_ezdxf(
    dxf_path: str, log=print,
) -> list[tuple[float, float, str]]:
    """Extract table cell texts from *T blocks (ACAD_TABLE data) via ezdxf."""
    cells: list[tuple[float, float, str]] = []
    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception as e:
        log(f"    Spec ezdxf open error: {e}")
        return cells
    for block in doc.blocks:
        if not block.name.startswith("*T"):
            continue
        for ent in block:
            if ent.dxftype() != "MTEXT":
                continue
            try:
                ins = ent.dxf.insert
                x, y = float(ins[0]), float(ins[1])
            except Exception:
                continue
            text = ent.plain_text().strip()
            if text:
                cells.append((x, y, text))
    return cells


def _extract_table_cells_raw(
    dxf_path: str, log=print,
) -> list[tuple[float, float, str]]:
    """Fallback: extract cells from raw DXF scanning codes 1, 3, 10, 20."""
    cells: list[tuple[float, float, str]] = []
    try:
        with open(dxf_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
    except Exception as e:
        log(f"    Spec parse error (read): {e}")
        return cells

    i = 0
    cur_x: float | None = None
    cur_y: float | None = None
    text_parts: list[str] = []

    while i < len(lines) - 1:
        code = lines[i].strip()
        value = lines[i + 1].strip()
        if code == "0":
            # Flush accumulated text from previous entity
            if text_parts and cur_x is not None and cur_y is not None:
                full = _strip_mtext_formatting("".join(text_parts))
                if full:
                    cells.append((cur_x, cur_y, full))
            text_parts = []
            cur_x = None
            cur_y = None
        elif code == "10":
            try:
                cur_x = float(value)
            except ValueError:
                pass
        elif code == "20":
            try:
                cur_y = float(value)
            except ValueError:
                pass
        elif code == "3":
            # MTEXT continuation chunk (comes BEFORE code 1)
            text_parts.append(value)
        elif code == "1":
            text_parts.append(value)
        i += 2

    # Flush last entity
    if text_parts and cur_x is not None and cur_y is not None:
        full = _strip_mtext_formatting("".join(text_parts))
        if full:
            cells.append((cur_x, cur_y, full))
    return cells


def parse_spec_dxf(dxf_path: str, log=print) -> list[SpecItem]:
    """Parse equipment specification from ACAD_TABLE in a СО.dxf file.

    Uses ezdxf to extract MTEXT cell contents from embedded *T table
    blocks, reconstructs the grid layout, and extracts equipment rows
    with names and quantities.  Falls back to raw DXF scanning when
    ezdxf yields no cells.

    T063: Improved cable extraction — separator filtering is only applied
    for the raw-DXF fallback (ezdxf cells are clean and don’t contain
    DXF group-code noise).  Float quantities supported (rounded to int).
    """
    # Primary: ezdxf-based extraction from *T blocks
    cells = _extract_table_cells_ezdxf(dxf_path, log)
    _use_sep_filter = False
    if not cells:
        # Fallback: raw DXF scan with code-1 + code-3 support
        cells = _extract_table_cells_raw(dxf_path, log)
        _use_sep_filter = True  # raw extraction may include DXF group codes

    if not cells:
        return []

    cells.sort(key=lambda c: (-c[1], c[0]))

    rows: list[list[str]] = []
    current_row: list[tuple[float, str]] = []
    current_y_val: float | None = None
    for x, y, text in cells:
        if current_y_val is None or abs(y - current_y_val) > 2.0:
            if current_row:
                current_row.sort(key=lambda c: c[0])
                rows.append([t for _, t in current_row])
            current_row = [(x, text)]
            current_y_val = y
        else:
            current_row.append((x, text))
    if current_row:
        current_row.sort(key=lambda c: c[0])
        rows.append([t for _, t in current_row])

    _UNIT_VALUES = {"шт", "м", "м.", "м.п.", "м.п",
                    "компл", "комплект", "кг", "м²", "м³",
                    "м2", "м3", "компл.", "маш/час", "маш/ча", "каб.",
                    "измерени", "изм."}

    items: list[SpecItem] = []
    for row in rows:
        # T063: Only filter separator noise for raw-DXF fallback.
        # ezdxf extraction yields clean cell text; filtering drops
        # legitimate quantity values like 91, 44, 40, 70, etc.
        if _use_sep_filter:
            data = [c for c in row if c not in _TABLE_SEPARATOR_VALUES]
        else:
            data = list(row)
        if len(data) < 3:
            continue
        pos = data[0]
        if not pos.isdigit() or int(pos) < 1 or int(pos) > 999:
            continue

        unit_idx = None
        for i, cell in enumerate(data):
            if cell.lower().strip() in _UNIT_VALUES:
                unit_idx = i
                break

        if unit_idx is None or unit_idx + 1 >= len(data):
            continue

        unit = data[unit_idx]
        qty_str = data[unit_idx + 1]
        # T063: Support float quantities (e.g. 156.5) — round to int.
        try:
            qty = int(qty_str)
        except (ValueError, TypeError):
            try:
                qty = round(float(qty_str.replace(",", ".")))
            except (ValueError, TypeError):
                continue
        if qty < 1:
            continue

        desc_parts = data[1:unit_idx]
        desc_parts = [p for p in desc_parts
                      if len(p) > 3 and not p.isdigit()]
        if not desc_parts:
            continue

        _SUPPLIER_HINTS = ("systeme electric", "ostec", "ооо ", "оао ",
                           "«dkc»", "световые технологии",
                           "мгк ", "электрокабель")
        _CATALOG_RE = re.compile(
            r"^[A-Za-z0-9]{5,}[-]?[A-Za-z0-9]*$"
        )

        model = ""
        supplier = ""
        catalog = ""
        clean_parts: list[str] = []
        for part in desc_parts:
            pl = part.lower().strip(' "«»')
            if any(sh in pl for sh in _SUPPLIER_HINTS):
                supplier = part
                continue
            if _CATALOG_RE.match(part.strip()) and not any(
                kw in part for kw in ["LED", "OPL", "ECO", "ARCTIC",
                                       "SLICK", "MARS", "LUNA", "STAR",
                                       "INSEL", "NERO", "MERCURY",
                                       "ВБШвнг", "ППГнг", "ВВГнг",
                                       "FRHF", "ПуГВнг",
                                       "ПуВВнг", "ПуВВ", "ПВС",
                                       "ПВ1", "ПВ3", "ШВВП", "ППВ", "АПВ"]
            ):
                catalog = part
                continue
            if not model and any(kw in part for kw in [
                "LED", "OPL", "ECO", "ARCTIC", "SLICK", "MARS", "LUNA",
                "STAR", "INSEL", "NERO", "MERCURY", "CONVERSION",
                "SIRAH", "ATLASDESIGN", "ETUDE",
            ]):
                model = part
            clean_parts.append(part)

        desc = " ".join(clean_parts) if clean_parts else ""
        if not desc:
            continue

        items.append(SpecItem(
            position=pos, description=desc,
            model=model, catalog_code=catalog,
            supplier=supplier, unit=unit, quantity=qty,
        ))

    if items:
        log(f"    [spec] Parsed {len(items)} items from equipment specification")
        for it in items:
            log(f"      [{it.position:>3}] {it.quantity:>4} {it.unit:>3}"
                f"  {it.description[:70]}")

    return items


def _collect_legend_descriptions(
    entries: list[tuple[str, float, float]],
    bbox: tuple[float, float, float, float],
) -> OrderedDict:
    """Collect all meaningful descriptions from the legend bounding box.

    Used when numbered-symbol matching is incomplete (e.g. Revit ЭМ exports
    where the legend has graphical symbols but no numbered text labels).
    """
    xmin, ymin, xmax, ymax = bbox
    descs: list[tuple[str, float]] = []
    for text, x, y in entries:
        if not (xmin < x < xmax and ymin < y < ymax):
            continue
        fl = text.split("\n")[0].strip()
        if fl in _LEGEND_SKIP:
            continue
        if SYMBOL_RE.match(fl) and len(fl) <= 4:
            continue
        if len(fl) < _LEGEND_DESC_MIN_LEN:
            continue
        descs.append((fl, y))

    descs.sort(key=lambda t: -t[1])
    result = OrderedDict()
    idx = 1
    for desc, _ in descs:
        result[str(idx)] = desc
        idx += 1
    return result


def _count_equipment_by_geometry(
    msp,
    legend: OrderedDict,
    legend_bbox: tuple[float, float, float, float] | None,
    entries: list[tuple[str, float, float]] | None = None,
) -> list[EquipmentItem]:
    """Count equipment by geometric entities for Revit-exported ЭМ drawings.

    When the numbered legend has very few entries (misidentified grid numbers),
    re-collects all descriptions from the legend bbox.

    Circles (color-based):
      color 1 (red)           → Щиты
      color 5 (blue) / others → розетки, кабельные выводы, оборудование
    INSERT blocks matching 'Подключение' → cable outlets.
    """
    if entries and legend_bbox and len(legend) <= 2:
        full_legend = _collect_legend_descriptions(entries, legend_bbox)
        if len(full_legend) > len(legend):
            legend = full_legend

    circle_colors = _count_plan_circles_by_color(msp, legend_bbox)
    total = sum(circle_colors.values())
    if total < 5:
        return []

    xmin, ymin, xmax, ymax = legend_bbox if legend_bbox else (0, 0, 0, 0)

    red_seen: set[tuple[int, int]] = set()
    blue_seen: set[tuple[int, int]] = set()
    for c in msp.query("CIRCLE"):
        if not _is_annotation_circle(c):
            continue
        cx, cy = c.dxf.center.x, c.dxf.center.y
        if legend_bbox and xmin < cx < xmax and ymin < cy < ymax:
            continue
        key = (round(cx), round(cy))
        color = c.dxf.get("color", 256)
        if color == 1:
            red_seen.add(key)
        else:
            blue_seen.add(key)

    red_total = len(red_seen)
    blue_total = len(blue_seen)

    cable_outlet_count = 0
    for e in msp.query("INSERT"):
        if "Подключение" not in e.dxf.name:
            continue
        x, y = e.dxf.insert.x, e.dxf.insert.y
        if legend_bbox and xmin < x < xmax and ymin < y < ymax:
            continue
        cable_outlet_count += 1

    # Count INSERT blocks that match equipment keywords from legend.
    # This is more accurate than blue circles when blocks represent actual
    # equipment instances (circles may be cable route markers).
    # Match against base block name (strip viewport suffix like "-План ...").
    _EQUIP_INSERT_KW = [
        "розетк", "выключател", "датчик", "светильник",
        "family", "пост управлен", "блок аварийн",
    ]
    _INSERT_SKIP_KW = [
        "подключение", "трасс", "трубе", "последовательность",
        "ось сетки", "помещени", "чертеж", "галерея",
    ]
    equip_insert_count = 0
    for e in msp.query("INSERT"):
        bname = e.dxf.name
        # Strip viewport suffix: "Block-ID-План ..." → "Block-ID"
        plan_idx = bname.find("-План ")
        base_name = bname[:plan_idx] if plan_idx > 0 else bname
        base_lower = base_name.lower()
        if any(sk in base_lower for sk in _INSERT_SKIP_KW):
            continue
        if not any(kw in base_lower for kw in _EQUIP_INSERT_KW):
            continue
        x, y = e.dxf.insert.x, e.dxf.insert.y
        if legend_bbox and xmin < x < xmax and ymin < y < ymax:
            continue
        equip_insert_count += 1

    shield_descs: list[tuple[str, str]] = []
    equip_descs: list[tuple[str, str]] = []
    cable_out_descs: list[tuple[str, str]] = []
    non_countable: list[tuple[str, str]] = []

    for sym, desc in legend.items():
        dl = desc.lower()
        if any(dl.startswith(p.lower()) for p in _NON_COUNTABLE_PREFIXES):
            non_countable.append((sym, desc))
        elif any(dl.startswith(p.lower()) for p in _PANEL_PREFIXES):
            shield_descs.append((sym, desc))
        elif "кабельный вывод" in dl:
            cable_out_descs.append((sym, desc))
        else:
            equip_descs.append((sym, desc))

    # Prefer INSERT block count over blue circles when INSERT blocks are
    # found and circle count is suspiciously high (circles may be cable
    # route junction markers, not equipment symbols).
    use_blue = blue_total
    if equip_insert_count > 0 and blue_total > equip_insert_count * 3:
        print(f"  [geometry mode] INSERT blocks={equip_insert_count} << "
              f"circles={blue_total} — using INSERT count")
        use_blue = equip_insert_count

    items: list[EquipmentItem] = []

    if shield_descs:
        for i, (sym, desc) in enumerate(shield_descs):
            items.append(EquipmentItem(
                symbol=sym, name=desc,
                count=red_total if i == 0 else 0,
            ))
    elif red_total:
        items.append(EquipmentItem(symbol="Щ", name="Щит (по цвету)", count=red_total))

    if len(equip_descs) == 1:
        sym, desc = equip_descs[0]
        items.append(EquipmentItem(symbol=sym, name=desc, count=use_blue))
    elif equip_descs:
        for sym, desc in equip_descs:
            items.append(EquipmentItem(symbol=sym, name=desc, count=0))
        items.append(EquipmentItem(
            symbol="⚙",
            name=f"Точечное оборудование на плане ({len(equip_descs)} видов)",
            count=use_blue,
        ))

    for sym, desc in cable_out_descs:
        items.append(EquipmentItem(
            symbol=sym, name=desc,
            count=cable_outlet_count if cable_outlet_count else 0,
        ))

    for sym, desc in non_countable:
        items.append(EquipmentItem(symbol=sym, name=desc, count=0))

    if red_total or use_blue:
        print(f"  [geometry mode] red(щиты)={red_total}  blue(оборуд.)={use_blue}")
    if cable_outlet_count:
        print(f"  [geometry mode] Cable outlets (Подключение): {cable_outlet_count}")

    return items


def _count_insert_symbols(
    msp,
    symbols: list[str],
    legend_bboxes: list[tuple[float, float, float, float]],
    grid_positions: set[tuple[int, int]] | None = None,
) -> dict[str, int]:
    """Count equipment by scanning INSERT block references whose names
    contain or match known symbol identifiers.

    Many Revit exports place equipment as INSERT block references rather
    than MTEXT labels.  This function counts those INSERTs outside
    the legend area and grid positions.
    """
    counts: Counter = Counter()
    sym_set = set(symbols)

    for e in msp.query("INSERT"):
        bname = e.dxf.name
        x, y = e.dxf.insert.x, e.dxf.insert.y

        # Skip entities inside any legend bbox
        if any(bx0 < x < bx1 and by0 < y < by1
               for bx0, by0, bx1, by1 in legend_bboxes):
            continue

        # Skip grid positions
        if grid_positions and (round(x), round(y)) in grid_positions:
            continue

        # Check if block name contains a known symbol
        # Block names may look like "SYMBOL_1", "1_SLICK", etc.
        bname_clean = bname.strip()
        for sym in sym_set:
            # Exact match or block name starts/ends with the symbol
            if bname_clean == sym:
                counts[sym] += 1
                break
            # Check ATTRIB values inside INSERT for symbol text
        else:
            # Also check block attributes for symbol text
            try:
                for attrib in e.attribs:
                    val = attrib.dxf.get("text", "").strip()
                    if val in sym_set:
                        counts[val] += 1
                        break
            except Exception:
                pass

    return dict(counts)


def process_dxf(dxf_path: str) -> list[EquipmentItem]:
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    entries: list[tuple[str, float, float]] = []
    for e in msp.query("MTEXT"):
        raw = e.text
        clean = _clean_mtext(raw)
        if not clean:
            continue
        x, y = e.dxf.insert.x, e.dxf.insert.y
        entries.append((clean, x, y))

    # P-001 fix: Also scan TEXT entities (not just MTEXT)
    text_entries = _get_text_entries(msp)
    if text_entries:
        entries.extend(text_entries)

    if not entries:
        print(f"  WARNING: No MTEXT/TEXT found in {dxf_path}")
        return []

    grid_positions = _detect_grid_labels(msp)

    all_bboxes = _find_all_legend_bboxes(entries)
    legend_bbox = all_bboxes[0] if all_bboxes else None

    # --- Multi-legend: collect descriptions from ALL legends ---
    all_legend_descs: OrderedDict = OrderedDict()
    for bbox in all_bboxes:
        descs = _parse_dxf_legend_descriptive(entries, bbox)
        for desc_text, _, _ in descs:
            if desc_text not in all_legend_descs.values():
                key = f"L{len(all_legend_descs) + 1}"
                all_legend_descs[key] = desc_text

    # --- Try plan annotations (Revit "N-description" pattern) ---
    annotation_items: list[EquipmentItem] = []
    if all_bboxes:
        annotation_items = _extract_plan_annotations(entries, all_bboxes)
        if annotation_items:
            print(f"  [annotations] Found {len(annotation_items)} equipment types from plan text")
            for it in annotation_items:
                print(f"    {it.symbol}: {it.count} × {it.name[:60]}")

    # --- Try numbered-symbol mode (ЭО plans) ---
    legend: OrderedDict = OrderedDict()
    block_map: dict[str, str] = {}

    for bbox in all_bboxes:
        partial = _parse_dxf_legend_numbered(entries, bbox)
        for sym, desc in partial.items():
            if sym not in legend:
                legend[sym] = desc

    if legend:
        exact, ae, unlisted, special = _count_dxf_symbols(
            entries, list(legend.keys()), all_bboxes, grid_positions
        )
        total_on_plan = sum(exact.values()) + sum(ae.values())

        # P-002 fix: If MTEXT symbol counting found zero/very few,
        # try INSERT block attribute counting as supplement
        if total_on_plan == 0:
            insert_counts = _count_insert_symbols(
                msp, list(legend.keys()), all_bboxes, grid_positions
            )
            if insert_counts:
                total_insert = sum(insert_counts.values())
                print(f"  [insert-count] MTEXT count=0, INSERT blocks found {total_insert} items")
                for sym in insert_counts:
                    exact[sym] = exact.get(sym, 0) + insert_counts[sym]
                total_on_plan = sum(exact.values()) + sum(ae.values())

        if annotation_items:
            ann_total = sum(it.count for it in annotation_items)
            if ann_total > total_on_plan * 3 and ann_total > 20:
                enriched = _enrich_annotations_with_legend(annotation_items, legend)
                # P-002 fix: For legend items that ended up with count=0 in the
                # enriched list, use their actual MTEXT/INSERT counts instead.
                # This prevents zero-count for items like switches and panels
                # that exist on the plan but weren't matched to annotations.
                for it in enriched:
                    if it.count == 0 and it.symbol in exact:
                        it.count = exact[it.symbol]
                    if it.count_ae == 0 and it.symbol in ae:
                        it.count_ae = ae[it.symbol]
                enriched_total = sum(it.count for it in enriched)
                print(f"  [merge] Annotations ({ann_total}) >> symbols ({total_on_plan})"
                      f" — using annotation counts ({enriched_total} items)")
                return enriched

        if total_on_plan <= len(legend) * 2:
            full_descs = _parse_dxf_legend_descriptive(entries, legend_bbox)
            if full_descs:
                full_legend = OrderedDict()
                for i, (desc, _, _) in enumerate(full_descs):
                    full_legend[str(i + 1)] = desc
                geo_items = _count_equipment_by_geometry(msp, full_legend, legend_bbox, entries)
            else:
                geo_items = _count_equipment_by_geometry(msp, legend, legend_bbox, entries)
            if geo_items:
                return geo_items

        items = []
        for sym, name in legend.items():
            items.append(EquipmentItem(
                symbol=sym, name=name,
                count=exact.get(sym, 0), count_ae=ae.get(sym, 0),
            ))
        for label, cnt in special.items():
            if cnt > 0:
                items.append(EquipmentItem(
                    symbol=label,
                    name=_SPECIAL_LABELS.get(label, label),
                    count=cnt,
                ))
        if unlisted:
            print("\n  Unlisted symbols in drawing body (not in legend):")
            for tok, cnt in sorted(unlisted.items(), key=lambda x: -x[1]):
                print(f"    {tok:>5} × {cnt}")
                items.append(EquipmentItem(
                    symbol=tok, name=f"[? {tok}]", count=cnt,
                ))
        return items

    # --- Try block-based mode (ЭМ/ЭС drawings) ---
    if legend_bbox:
        legend, block_map = _parse_dxf_legend_blocks(msp, entries, legend_bbox)

    if legend and block_map:
        reverse_map: dict[str, str] = {v: k for k, v in block_map.items()}
        block_counts: Counter = Counter()
        for e in msp.query("INSERT"):
            bname = e.dxf.name
            if bname not in block_map:
                continue
            x, y = e.dxf.insert.x, e.dxf.insert.y
            if any(bx0 < x < bx1 and by0 < y < by1
                   for bx0, by0, bx1, by1 in all_bboxes):
                continue
            block_counts[block_map[bname]] += 1

        total_block_count = sum(block_counts.values())
        n_circles = sum(1 for c in msp.query("CIRCLE"))
        if total_block_count < 5 and n_circles > 20:
            print(f"  Block mode found only {total_block_count} instances but {n_circles} circles — switching to geometric mode")
        else:
            items = []
            for label, desc in legend.items():
                items.append(EquipmentItem(
                    symbol=label,
                    name=desc,
                    count=block_counts.get(label, 0),
                ))
            # Merge plan annotation items for equipment not found by blocks
            if all_bboxes:
                annotation_items = _extract_plan_annotations(entries, all_bboxes)
                found_names = {it.name.lower() for it in items if it.count > 0}
                for ait in annotation_items:
                    if ait.name.lower() not in found_names:
                        items.append(ait)
            return items

    # --- Try geometric circle counting (Revit ЭМ розетка plans) ---
    if legend_bbox:
        descs = _parse_dxf_legend_descriptive(entries, legend_bbox)
        if descs:
            desc_legend = OrderedDict(
                (f"D{i+1}", d) for i, (d, _, _) in enumerate(descs)
                if not any(d.startswith(p) for p in _NON_COUNTABLE_PREFIXES)
            )
            if desc_legend:
                geo_items = _count_equipment_by_geometry(msp, desc_legend, legend_bbox)
                if geo_items:
                    geo_total = sum(it.count for it in geo_items)
                    ann_total = sum(it.count for it in annotation_items)
                    if ann_total > 0 and geo_total > ann_total * 3:
                        print(f"  [geometry skip] geo_total={geo_total} >> "
                              f"annotation_total={ann_total} — preferring annotations")
                    else:
                        print("  Using geometric circle counting mode")
                        return geo_items

    # --- Fallback: plan annotations + auto-detect ---
    if all_bboxes:
        if not annotation_items:
            annotation_items = _extract_plan_annotations(entries, all_bboxes)
        if annotation_items:
            print("  Using plan annotation mode (N-description)")
            return annotation_items

    print("  WARNING: Legend not found, auto-detecting symbols...")
    all_syms = set()
    for text, x, y in entries:
        if grid_positions and (round(x), round(y)) in grid_positions:
            continue
        fl = text.split("\n")[0].strip()
        if SYMBOL_RE.match(fl) and len(fl) <= 4:
            all_syms.add(fl)
    legend = OrderedDict((s, f"[Auto {s}]") for s in sorted(all_syms))

    exact, ae, unlisted, special = _count_dxf_symbols(entries, list(legend.keys()), all_bboxes, grid_positions)
    items = []
    for sym, name in legend.items():
        items.append(EquipmentItem(
            symbol=sym, name=name,
            count=exact.get(sym, 0), count_ae=ae.get(sym, 0),
        ))
    for label, cnt in special.items():
        if cnt > 0:
            items.append(EquipmentItem(
                symbol=label, name=_SPECIAL_LABELS.get(label, label), count=cnt,
            ))
    return items


# ===================================================================
#  PDF processing  (pdfplumber)
# ===================================================================

def extract_text(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        pages = []
        for page in pdf.pages:
            t = page.extract_text(x_tolerance=3, y_tolerance=3)
            if t:
                pages.append(t)
    return "\n".join(pages)


def extract_words(pdf_path: str) -> list[dict]:
    with pdfplumber.open(pdf_path) as pdf:
        words = []
        for page in pdf.pages:
            for w in (page.extract_words(x_tolerance=3, y_tolerance=3) or []):
                words.append(w)
    return words


def find_legend_y(words: list[dict]) -> float | None:
    for w in words:
        if "Условные" in w["text"]:
            return w["top"]
    return None


def _find_oboz_x(words: list[dict], legend_y: float) -> float:
    for w in words:
        if abs(w["top"] - legend_y) > 40:
            continue
        if "Обозначение" in w["text"] and w["top"] > legend_y:
            return w["x0"]
    return 700.0


def _y_group(wlist: list[dict], tol: float = 5) -> list[tuple[float, list[dict]]]:
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


def parse_legend_coords(pdf_path: str) -> OrderedDict:
    words = extract_words(pdf_path)
    legend_y = find_legend_y(words)
    if legend_y is None:
        return OrderedDict()

    oboz_x = _find_oboz_x(words, legend_y)
    sym_x_boundary = oboz_x + 60
    legend_bottom = legend_y + 400

    area_words = [
        w for w in words
        if legend_y + 20 < w["top"] < legend_bottom
        and oboz_x - 10 < w["x0"] < 1200
    ]

    rows = _y_group(area_words, tol=5)

    sym_rows: list[tuple[float, str]] = []
    desc_rows: list[tuple[int, float, str]] = []
    sym_desc_rows: list[tuple[float, str, str]] = []

    for idx, (y, row_words) in enumerate(rows):
        sym_parts = [
            w for w in row_words
            if w["x0"] < sym_x_boundary and SYMBOL_RE.match(w["text"])
        ]
        desc_parts = [w for w in row_words if w["x0"] >= sym_x_boundary]

        sym_text = sym_parts[0]["text"] if sym_parts else None
        desc_text = " ".join(
            w["text"] for w in sorted(desc_parts, key=lambda w: w["x0"])
        ).strip()

        if not sym_text and ("Наименование" in desc_text or "Примечание" in desc_text):
            continue

        if sym_text and desc_text:
            sym_desc_rows.append((y, sym_text, desc_text))
        elif sym_text:
            sym_rows.append((y, sym_text))
        elif desc_text and len(desc_text) > 5:
            desc_rows.append((idx, y, desc_text))

    legend = OrderedDict()
    for _, sym, desc in sym_desc_rows:
        legend[sym] = desc

    unmatched_syms: list[tuple[float, str]] = []
    for y, sym in sym_rows:
        m = CIRCUIT_VARIANT_RE.match(sym)
        if m and m.group(2) == "А" and m.group(1) in legend:
            legend[sym] = legend[m.group(1)] + " [авар. цепь]"
        else:
            unmatched_syms.append((y, sym))

    used_desc: set[int] = set()
    assigned_syms: set[int] = set()
    remaining_sym_list = list(enumerate(unmatched_syms))

    while True:
        best_si, best_di, best_dist = -1, -1, 999.0
        for si, (y_sym, _sym) in remaining_sym_list:
            if si in assigned_syms:
                continue
            for di, (_, y_d, _) in enumerate(desc_rows):
                if di in used_desc:
                    continue
                dist = abs(y_d - y_sym)
                if dist < best_dist:
                    best_dist = dist
                    best_si, best_di = si, di
        if best_si < 0 or best_dist > 150:
            break
        _, (_, sym) = remaining_sym_list[best_si]
        desc_text = desc_rows[best_di][2]
        base_y = desc_rows[best_di][1]
        for j in range(best_di + 1, len(desc_rows)):
            if j in used_desc:
                continue
            if abs(desc_rows[j][1] - base_y) < 20:
                desc_text += " " + desc_rows[j][2]
                used_desc.add(j)
            else:
                break
        legend[sym] = desc_text
        assigned_syms.add(best_si)
        used_desc.add(best_di)

    for si, (_, sym) in remaining_sym_list:
        if si not in assigned_syms:
            legend[sym] = f"[Обозначение {sym}]"

    all_y = {s: y for y, s, _ in sym_desc_rows}
    all_y.update({s: y for y, s in sym_rows})
    return OrderedDict(sorted(legend.items(), key=lambda kv: all_y.get(kv[0], 9999)))


def find_legend_start_text(text: str) -> int:
    for marker in ["Условные обозначения", "УСЛОВНЫЕ ОБОЗНАЧЕНИЯ"]:
        pos = text.find(marker)
        if pos != -1:
            return pos
    return -1


def parse_legend_text(text: str) -> OrderedDict:
    legend = OrderedDict()
    pos = find_legend_start_text(text)
    if pos < 0:
        return legend
    section = text[pos:]
    for end in ["Групповую сеть", "Примечание:", "Изм."]:
        idx = section.find(end)
        if idx > 0:
            section = section[:idx]
            break
    lines = [l.strip() for l in section.split("\n") if l.strip()]
    for line in lines:
        if "\t" not in line:
            continue
        parts = [p.strip() for p in line.split("\t")]
        if len(parts) >= 2 and SYMBOL_RE.match(parts[0]):
            legend[parts[0]] = parts[1]
    if legend:
        return legend
    symbols, descriptions = [], []
    for line in lines:
        if "Условные" in line or "Обозначение" in line:
            continue
        if SYMBOL_RE.match(line):
            symbols.append(line)
        elif len(line) > 20:
            descriptions.append(line)
    if symbols and descriptions and len(symbols) == len(descriptions):
        for s, d in zip(symbols, descriptions):
            legend[s] = d
    elif symbols:
        for s in symbols:
            legend[s] = f"[Equipment {s}]"
    return legend


def parse_legend(pdf_path: str, text: str) -> OrderedDict:
    legend = parse_legend_coords(pdf_path)
    if legend:
        return legend
    return parse_legend_text(text)


def count_symbols_pdf(
    text: str,
    symbols: list[str],
    legend_pos: int,
) -> tuple[dict[str, int], dict[str, int], dict[str, int]]:
    body = text[:legend_pos] if legend_pos > 0 else text
    lines = body.split("\n")
    filtered = []
    for line in lines:
        s = line.strip()
        if GRID_LINE_RE.match(s):
            continue
        if any(p.match(s) for p in SKIP_LINE_PATTERNS):
            continue
        filtered.append(line)
    body = "\n".join(filtered)

    exact, ae = {}, {}
    sym_set = set(symbols)
    for sym in sorted(symbols, key=len, reverse=True):
        pat = r"(?<!\S)" + re.escape(sym) + r"(?!\S)"
        exact[sym] = len(re.findall(pat, body))
        if not sym.endswith("АЭ"):
            ae_sym = sym + "АЭ"
            if ae_sym not in sym_set:
                pat_ae = r"(?<!\S)" + re.escape(ae_sym) + r"(?!\S)"
                ae[sym] = len(re.findall(pat_ae, body))

    all_tokens = set(re.findall(r"(?<!\S)(\d{1,2}[А-Яа-яЁё]{0,3})(?!\S)", body))
    unlisted = {}
    for tok in sorted(all_tokens):
        if tok in sym_set:
            continue
        m_var = CIRCUIT_VARIANT_RE.match(tok)
        if m_var and m_var.group(1) in sym_set:
            continue
        if tok in ("0", "1х") or tok.endswith("ЛК"):
            continue
        pat = r"(?<!\S)" + re.escape(tok) + r"(?!\S)"
        cnt = len(re.findall(pat, body))
        if cnt > 0:
            unlisted[tok] = cnt
    return exact, ae, unlisted


def process_pdf(pdf_path: str) -> list[EquipmentItem]:
    text = extract_text(pdf_path)
    if not text.strip():
        print(f"  WARNING: No text extracted from {pdf_path}")
        return []

    legend = parse_legend(pdf_path, text)
    if not legend:
        print("  WARNING: Legend (Условные обозначения) not found!")
        all_tokens = set(re.findall(r"(?<!\S)(\d{1,2}[А-Яа-яЁё]{0,3})(?!\S)", text))
        legend = OrderedDict(
            (s, f"[Auto-detected {s}]") for s in sorted(all_tokens) if len(s) <= 4
        )

    legend_pos = find_legend_start_text(text)
    exact, ae, unlisted = count_symbols_pdf(text, list(legend.keys()), legend_pos)

    items = []
    for sym, name in legend.items():
        items.append(EquipmentItem(
            symbol=sym, name=name,
            count=exact.get(sym, 0), count_ae=ae.get(sym, 0),
        ))

    if unlisted:
        print("\n  Unlisted symbols in drawing body (not in legend):")
        for tok, cnt in sorted(unlisted.items(), key=lambda x: -x[1]):
            print(f"    {tok:>5} × {cnt}")
    return items


# ===================================================================
#  Shared output
# ===================================================================

def print_table(items: list[EquipmentItem], file_name: str) -> None:
    if not items:
        print("  No equipment found.\n")
        return

    max_sym = max(len(it.symbol) for it in items)
    max_name = min(max(len(it.name) for it in items), 85)
    has_ae = any(it.count_ae > 0 for it in items)

    if has_ae:
        header = (
            f"  {'#':<4} {'Обозн.':<{max_sym+2}} {'Наименование':<{max_name+2}}"
            f" {'Кол-во':>6} {'АЭ':>5} {'Всего':>6}"
        )
    else:
        header = (
            f"  {'#':<4} {'Обозн.':<{max_sym+2}} {'Наименование':<{max_name+2}}"
            f" {'Кол-во':>6}"
        )
    sep = "  " + "-" * (len(header) - 2)

    print(f"\n  === {file_name} ===")
    print(header)
    print(sep)

    total, total_ae = 0, 0
    for i, it in enumerate(items, 1):
        name = it.name[:max_name]
        row_total = it.count + it.count_ae
        total += it.count
        total_ae += it.count_ae
        if has_ae:
            ae_str = str(it.count_ae) if it.count_ae else ""
            print(
                f"  {i:<4} {it.symbol:<{max_sym+2}} {name:<{max_name+2}}"
                f" {it.count:>6} {ae_str:>5} {row_total:>6}"
            )
        else:
            print(
                f"  {i:<4} {it.symbol:<{max_sym+2}} {name:<{max_name+2}}"
                f" {it.count:>6}"
            )

    print(sep)
    grand = total + total_ae
    if has_ae:
        print(
            f"  {'':4} {'ИТОГО':<{max_sym+2}} {'':<{max_name+2}}"
            f" {total:>6} {total_ae:>5} {grand:>6}"
        )
    else:
        print(
            f"  {'':4} {'ИТОГО':<{max_sym+2}} {'':<{max_name+2}}"
            f" {total:>6}"
        )
    print()


def export_csv(items: list[EquipmentItem], path: str) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Обозначение", "Наименование", "Кол-во", "АЭ-вариант", "Всего"])
        for it in items:
            w.writerow([it.symbol, it.name, it.count, it.count_ae, it.count + it.count_ae])
    print(f"  CSV saved: {path}")


def export_json(items: list[EquipmentItem], path: str) -> None:
    data = [asdict(it) for it in items]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  JSON saved: {path}")


# ===================================================================
#  CLI
# ===================================================================

def _process_file(path: Path) -> list[EquipmentItem]:
    ext = path.suffix.lower()
    if ext == ".dxf":
        if not _HAS_DXF:
            print(f"  SKIP {path.name}: ezdxf not installed")
            return []
        return process_dxf(str(path))
    elif ext == ".pdf":
        if not _HAS_PDF:
            print(f"  SKIP {path.name}: pdfplumber not installed")
            return []
        return process_pdf(str(path))
    return []


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Count equipment in engineering PDF/DXF schematics"
    )
    parser.add_argument("input", help="PDF/DXF file or directory")
    parser.add_argument("--csv", help="Export to CSV file")
    parser.add_argument("--json", help="Export to JSON file")
    parser.add_argument(
        "--png", nargs="?", const="auto", default=None,
        help="Render DXF to PNG with equipment markers (optionally specify output path)",
    )
    args = parser.parse_args()

    inp = Path(args.input)
    if inp.is_file() and inp.suffix.lower() in SUPPORTED_EXT:
        files = [inp]
    elif inp.is_dir():
        files = sorted(
            f for f in inp.iterdir()
            if f.suffix.lower() in SUPPORTED_EXT and not f.name.startswith("_")
        )
        if not files:
            sys.exit(f"No PDF/DXF files in {inp}")
    else:
        sys.exit(f"Invalid input: {args.input} (supported: {', '.join(SUPPORTED_EXT)})")

    all_items: list[EquipmentItem] = []
    for f in files:
        items = _process_file(f)
        print_table(items, f.name)
        all_items.extend(items)

        if args.png is not None and f.suffix.lower() == ".dxf":
            try:
                from dxf_visualizer import visualize_dxf
                png_out = None if args.png == "auto" else args.png
                visualize_dxf(str(f), png_out, items=items)
            except ImportError as e:
                print(f"  PNG skipped (missing dependency): {e}")
            except Exception as e:
                print(f"  PNG error: {e}")

    if args.csv:
        export_csv(all_items, args.csv)
    if args.json:
        export_json(all_items, args.json)


if __name__ == "__main__":
    main()
