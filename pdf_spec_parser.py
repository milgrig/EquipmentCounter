#!/usr/bin/env python3
"""pdf_spec_parser.py -- Parse specification tables from СО PDF files.

Extracts equipment specifications from PDF files that follow the standard
Russian electrical specification table format (Спецификация оборудования,
изделий и материалов).

These files typically have names like:
  СО-СО_01.pdf, СО-СО_02.pdf, etc.

The table columns are (column indices vary between files!):
  Поз. | Наименование и техническая характеристика | Тип, марка |
  Код продукции | Поставщик | Ед. измерения | Количество | Масса | Примечания

Usage:
    from pdf_spec_parser import parse_spec_pdf, is_spec_pdf
    if is_spec_pdf("СО-СО_01.pdf"):
        items = parse_spec_pdf("СО-СО_01.pdf")
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pdfplumber

from equipment_counter import SpecItem

_log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Header column names (partial match, case-insensitive)
_HDR_POS = "поз"
_HDR_NAME = "наименование"
_HDR_TYPE = "тип"
_HDR_CODE = "код"
_HDR_SUPPLIER = "поставщик"
_HDR_UNIT = "изм"
_HDR_QTY = "колич"
_HDR_MASS = "масса"
_HDR_NOTE = "примеч"

# Known units
_KNOWN_UNITS = {"шт", "м", "м.п.", "кг", "компл", "компл.", "комп", "каб.",
                "маш/час", "изм.", "шт /кг", "шт/кг"}

# Section header keywords (these rows have only a description, no position/qty)
_SECTION_HEADER_KEYWORDS = (
    "щитовое оборудование",
    "светотехническое оборудование",
    "электроустановочные",
    "кабельная продукция",
    "кабеленесущие",
    "пвх изделия",
    "монтажные изделия",
    "пусконаладочные",
    "заземление",
    "молниезащит",
)

# Cable description prefix rows (multi-line cable family descriptions)
_CABLE_DESC_PREFIXES = (
    "кабель силовой",
    "с медной жилой",
    "с медными жилами",
    "пониженной пожарной",
    "не распространяющий",
    "оболочкой из пвх",
    "оболочкий из пвх",
)

# Model keywords for extraction
_MODEL_KEYWORDS_RE = re.compile(
    r"\b(SLICK|ARCTIC|INSEL|MERCURY|ATOM|MARS|LUNA|NERO|STAR|"
    r"CD\s+LED|OPL|ECO\s+LED|CONVERSION\s+KIT|SIRAH|PRS\s+LED|"
    r"ВБШвнг|ППГнг|ВВГнг|КГнг|ПуВВнг|FRHF|FRLS|"
    r"ПКУ|ЭТЮД|Этюд|FSB|ГМЛ|ТТК|DKC|Ostec)\b",
    re.IGNORECASE,
)

# Supplier hints
_SUPPLIER_HINTS = (
    "systeme electric", "световые технологии", "мгк", "ostec",
    "oao", "оао", "ооо", "dkc", "квт", "экоэл",
    "электрокабель",
)


# ---------------------------------------------------------------------------
# File identification
# ---------------------------------------------------------------------------

def is_spec_pdf(filename: str) -> bool:
    """Check if a filename looks like a spec (СО) PDF file.

    Matches patterns like: СО-СО_01.pdf, СО_01.pdf, CO-CO_01.pdf
    """
    name = Path(filename).stem.lower()
    return bool(re.match(r"^[сc][оo][-_]", name, re.IGNORECASE))


# ---------------------------------------------------------------------------
# Column mapping
# ---------------------------------------------------------------------------

@dataclass
class _ColumnMap:
    """Maps logical column names to physical column indices."""
    pos: int = -1       # Поз. (position number)
    name: int = -1      # Наименование
    type_brand: int = -1  # Тип, марка
    code: int = -1      # Код продукции
    supplier: int = -1  # Поставщик
    unit: int = -1      # Ед. измерения
    qty: int = -1       # Количество
    mass: int = -1      # Масса
    note: int = -1      # Примечания


def _detect_columns(header_row: list[str | None]) -> Optional[_ColumnMap]:
    """Detect column mapping from a header row.

    Returns None if this doesn't look like a header row.
    """
    cm = _ColumnMap()
    found = 0

    for ci, cell in enumerate(header_row):
        if not cell:
            continue
        lower = cell.strip().lower().replace("\n", " ")

        if _HDR_POS in lower and cm.pos == -1:
            cm.pos = ci
            found += 1
        elif _HDR_NAME in lower and cm.name == -1:
            cm.name = ci
            found += 1
        elif _HDR_TYPE in lower and "тип" in lower and cm.type_brand == -1:
            cm.type_brand = ci
            found += 1
        elif _HDR_CODE in lower and cm.code == -1:
            cm.code = ci
            found += 1
        elif _HDR_SUPPLIER in lower and cm.supplier == -1:
            cm.supplier = ci
            found += 1
        elif _HDR_UNIT in lower and cm.unit == -1:
            cm.unit = ci
            found += 1
        elif _HDR_QTY in lower and cm.qty == -1:
            cm.qty = ci
            found += 1
        elif _HDR_MASS in lower and cm.mass == -1:
            cm.mass = ci
            found += 1
        elif _HDR_NOTE in lower and cm.note == -1:
            cm.note = ci
            found += 1

    # Need at least position + name + unit + qty to be valid
    if found >= 4 and cm.pos >= 0 and cm.name >= 0 and cm.unit >= 0 and cm.qty >= 0:
        return cm
    return None


# ---------------------------------------------------------------------------
# Quantity parsing
# ---------------------------------------------------------------------------

def _parse_quantity(raw: str | None) -> Optional[int]:
    """Parse a quantity string that may contain spaces, newlines, or fractions.

    Examples:
        "414" -> 414
        "2 781" -> 2781
        "33\n560" -> 33560
        "27\n759" -> 27759
        "1176/\n745" -> 1176  (take first number before '/')
        "шт" -> None (not a number)
    """
    if not raw:
        return None

    txt = raw.strip()
    if not txt:
        return None

    # Handle "X/ Y" or "X/\nY" format (e.g., "1176/ 745" = qty/mass)
    if "/" in txt:
        parts = txt.split("/")
        txt = parts[0].strip()

    # Remove all whitespace (spaces, newlines, tabs)
    txt = re.sub(r"\s+", "", txt)

    # Must be a number
    if not re.match(r"^\d+$", txt):
        return None

    try:
        return int(txt)
    except ValueError:
        return None


def _parse_unit(raw: str | None) -> str:
    """Parse unit string, handling multiline and composite units.

    Examples:
        "шт" -> "шт"
        "Ед. \\n измерения" -> "" (this is a header, not data)
        "шт /кг" -> "шт"
        "м" -> "м"
    """
    if not raw:
        return ""
    txt = raw.strip().replace("\n", " ").strip()

    # Handle composite units like "шт /кг" -> take first
    if "/" in txt:
        txt = txt.split("/")[0].strip()

    return txt


# ---------------------------------------------------------------------------
# Model extraction
# ---------------------------------------------------------------------------

def _extract_model(description: str, type_brand: str = "") -> str:
    """Extract model identifier from description and type/brand columns."""
    combined = f"{description} {type_brand}".strip()

    # Try to find known model keywords
    match = _MODEL_KEYWORDS_RE.search(combined)
    if match:
        # Get a broader context around the match
        start = max(0, match.start())
        # Find the end of the model string (until end or next comma/semicolon)
        rest = combined[start:]
        end_match = re.search(r"[,;]", rest)
        if end_match:
            return rest[:end_match.start()].strip()
        return rest[:80].strip()

    # For cable items, extract brand+section
    cable_match = re.search(
        r"(ВБШвнг|ППГнг|ВВГнг|КГнг|ПуВВнг|ПВС|NYM)"
        r"[^\s]*\s*\d+[xхХ×]\d+",
        combined, re.IGNORECASE,
    )
    if cable_match:
        return cable_match.group(0).strip()

    return type_brand.strip()[:80] if type_brand else ""


def _extract_supplier(row: list[str | None], cm: _ColumnMap) -> str:
    """Extract supplier name from the appropriate column."""
    if cm.supplier >= 0 and cm.supplier < len(row):
        val = row[cm.supplier]
        if val:
            return val.strip().replace("\n", " ")
    return ""


def _extract_catalog_code(row: list[str | None], cm: _ColumnMap) -> str:
    """Extract catalog/product code."""
    if cm.code >= 0 and cm.code < len(row):
        val = row[cm.code]
        if val:
            return val.strip().replace("\n", "")
    return ""


# ---------------------------------------------------------------------------
# Section detection
# ---------------------------------------------------------------------------

def _is_section_header(row: list[str | None], cm: _ColumnMap) -> Optional[str]:
    """Check if this row is a section header (only description filled).

    Returns section name or None.
    """
    # Get description cell
    desc = ""
    if cm.name >= 0 and cm.name < len(row):
        desc = (row[cm.name] or "").strip()

    if not desc:
        return None

    # Check that position and quantity are empty
    pos_val = (row[cm.pos] if cm.pos >= 0 and cm.pos < len(row) else None) or ""
    qty_val = (row[cm.qty] if cm.qty >= 0 and cm.qty < len(row) else None) or ""
    unit_val = (row[cm.unit] if cm.unit >= 0 and cm.unit < len(row) else None) or ""

    if pos_val.strip() or qty_val.strip() or unit_val.strip():
        return None

    # Check against known section keywords
    lower = desc.lower()
    for kw in _SECTION_HEADER_KEYWORDS:
        if kw in lower:
            return desc

    return None


def _is_cable_continuation(row: list[str | None], cm: _ColumnMap) -> bool:
    """Check if this row is a continuation of a cable family description.

    These are multi-line rows like:
      "Кабель силовой бронированный лентами,"
      "с медной жилой, изоляцией и защитным шлангом из ПВХ"
      "пониженной пожарной опасности количеством жил и сечением:"
    """
    desc = ""
    if cm.name >= 0 and cm.name < len(row):
        desc = (row[cm.name] or "").strip()

    if not desc:
        return False

    # No position number → might be continuation
    pos_val = (row[cm.pos] if cm.pos >= 0 and cm.pos < len(row) else None) or ""
    qty_val = (row[cm.qty] if cm.qty >= 0 and cm.qty < len(row) else None) or ""

    if pos_val.strip() or qty_val.strip():
        return False

    lower = desc.lower()
    for prefix in _CABLE_DESC_PREFIXES:
        if lower.startswith(prefix):
            return True

    return False


def _is_footer_row(row: list[str | None]) -> bool:
    """Check if this row belongs to the document footer/stamp block."""
    # Footer rows contain patterns like "Изм.", "Кол.уч", "Лист", "Подп."
    text = " ".join((c or "") for c in row).lower()
    footer_signals = 0
    for kw in ("изм.", "кол.уч", "подп.", "дата", "разработал", "проверил",
               "нормоконтр", "гип", "зам.", "лист", "листов", "стадия"):
        if kw in text:
            footer_signals += 1
    return footer_signals >= 2


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_spec_pdf(pdf_path: str) -> list[SpecItem]:
    """Parse specification tables from a СО PDF file.

    Parameters
    ----------
    pdf_path : str
        Path to the specification PDF file.

    Returns
    -------
    list[SpecItem]
        List of specification items extracted from the PDF.
    """
    path = Path(pdf_path)
    if not path.exists():
        _log.warning("Spec PDF not found: %s", pdf_path)
        return []

    _log.info("Parsing spec PDF: %s", path.name)

    items: list[SpecItem] = []
    current_section = ""
    cable_family_desc = ""  # Accumulated cable family description
    cable_family_brand = ""  # Brand from cable family header (e.g., ГОСТ ref)
    cm: Optional[_ColumnMap] = None

    try:
        with pdfplumber.open(str(path)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                if not tables:
                    continue

                for table in tables:
                    for row in table:
                        if not row:
                            continue

                        # Skip if all cells are empty
                        if not any((c or "").strip() for c in row):
                            continue

                        # Detect column mapping from header row (once per file)
                        if cm is None:
                            detected = _detect_columns(row)
                            if detected:
                                cm = detected
                                _log.debug("Column map detected: pos=%d name=%d unit=%d qty=%d",
                                          cm.pos, cm.name, cm.unit, cm.qty)
                                continue
                            else:
                                continue  # Skip rows before header

                        # Skip footer rows
                        if _is_footer_row(row):
                            continue

                        # Check for section header
                        section = _is_section_header(row, cm)
                        if section:
                            current_section = section
                            cable_family_desc = ""
                            cable_family_brand = ""
                            _log.debug("Section: %s", current_section)
                            continue

                        # Check for cable family continuation
                        if _is_cable_continuation(row, cm):
                            desc = (row[cm.name] or "").strip()
                            cable_family_desc += " " + desc
                            # Check type/brand column for ГОСТ reference
                            if cm.type_brand >= 0 and cm.type_brand < len(row):
                                brand = (row[cm.type_brand] or "").strip()
                                if brand:
                                    cable_family_brand = brand
                            continue

                        # Try to parse as data row
                        pos_raw = (row[cm.pos] if cm.pos < len(row) else None) or ""
                        pos_raw = pos_raw.strip().replace("\n", "")

                        # Skip non-data rows (no valid position number)
                        if not re.match(r"^\d{1,3}$", pos_raw):
                            # Could be a garbled side-stamp cell
                            continue

                        # Extract fields
                        desc_raw = (row[cm.name] if cm.name < len(row) else None) or ""
                        desc = desc_raw.strip().replace("\n", " ")

                        type_brand_raw = ""
                        if cm.type_brand >= 0 and cm.type_brand < len(row):
                            type_brand_raw = (row[cm.type_brand] or "").strip().replace("\n", " ")

                        unit_raw = (row[cm.unit] if cm.unit < len(row) else None) or ""
                        unit = _parse_unit(unit_raw)

                        qty_raw = (row[cm.qty] if cm.qty < len(row) else None) or ""
                        qty = _parse_quantity(qty_raw)

                        if not desc or qty is None:
                            continue

                        # For cable items with short descriptions (e.g., "3х1,5мм-0,66кВ"),
                        # prepend the cable family description
                        full_desc = desc
                        is_short_cable = (
                            len(desc) < 30
                            and re.search(r"\d+[xхХ×]\d", desc)
                        )
                        if is_short_cable:
                            # Short desc like "3х1,5мм-0,66кВ" → this is a cable sub-item
                            # Build full description: "Кабель силовой ... 3х1,5мм-0,66кВ"
                            brand = type_brand_raw or cable_family_brand
                            if brand:
                                full_desc = f"Кабель силовой с медными жилами {brand} {desc}"
                            elif cable_family_desc:
                                full_desc = f"{cable_family_desc.strip()} {desc}"
                            else:
                                full_desc = f"Кабель {desc}"

                        model = _extract_model(full_desc, type_brand_raw)
                        supplier = _extract_supplier(row, cm)
                        catalog = _extract_catalog_code(row, cm)

                        item = SpecItem(
                            position=pos_raw,
                            description=full_desc,
                            model=model,
                            catalog_code=catalog,
                            supplier=supplier,
                            unit=unit,
                            quantity=qty,
                        )
                        items.append(item)
                        _log.info("  [spec] pos=%s qty=%d %s desc=%s",
                                 pos_raw, qty, unit, full_desc[:70])

    except Exception as e:
        _log.error("Error parsing spec PDF %s: %s", pdf_path, e)
        return items

    _log.info("Parsed %d spec items from %s", len(items), path.name)
    return items


# ---------------------------------------------------------------------------
# Convenience: parse all spec PDFs in a folder
# ---------------------------------------------------------------------------

def parse_all_specs_in_folder(folder: str | Path) -> list[SpecItem]:
    """Find and parse all СО spec PDFs in a folder.

    Returns combined list of SpecItem from all spec files.
    """
    folder_path = Path(folder)
    if not folder_path.is_dir():
        return []

    all_items: list[SpecItem] = []
    spec_files = sorted(
        f for f in folder_path.glob("*.pdf")
        if is_spec_pdf(f.name)
    )

    if not spec_files:
        _log.info("No spec (СО) PDF files found in %s", folder)
        return []

    _log.info("Found %d spec PDF files: %s", len(spec_files),
             ", ".join(f.name for f in spec_files))

    for spec_file in spec_files:
        items = parse_spec_pdf(str(spec_file))
        all_items.extend(items)

    return all_items


# ---------------------------------------------------------------------------
# CLI / debug
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    if len(sys.argv) < 2:
        print("Usage: python pdf_spec_parser.py <folder_or_pdf_path>")
        print("Example: python pdf_spec_parser.py Data/...")
        sys.exit(1)

    target = Path(sys.argv[1])

    if target.is_file():
        items = parse_spec_pdf(str(target))
    elif target.is_dir():
        items = parse_all_specs_in_folder(target)
    else:
        print(f"Not found: {target}")
        sys.exit(1)

    print(f"\n{'='*80}")
    print(f"TOTAL: {len(items)} spec items")
    print(f"{'='*80}")
    print(f"{'Pos':>4s} | {'Qty':>7s} | {'Unit':>5s} | {'Description':<70s} | Model")
    print("-" * 160)
    for it in items:
        print(f"{it.position:>4s} | {it.quantity:>7d} | {it.unit:>5s} | "
              f"{it.description[:70]:<70s} | {it.model[:40]}")
