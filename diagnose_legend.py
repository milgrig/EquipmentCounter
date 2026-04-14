#!/usr/bin/env python3
"""diagnose_legend.py -- Diagnostic: run the legend parser on all 10 test cases.

For each test case and each PDF inside it:
  1. Call parse_legend() from pdf_legend_parser.py (new S017 parser)
  2. Call old parse_legend() from equipment_counter.py for comparison
  3. Record detailed per-item diagnostics
  4. For the worst cases, dump raw pdfplumber data for manual inspection

Outputs:
  - diagnose_legend_report.json  -- full structured report
  - stdout summary
"""

from __future__ import annotations

import json
import re
import sys
import time
import traceback
from collections import OrderedDict
from dataclasses import asdict, dataclass, field
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

import pdfplumber

# ── Import S017 (new) parser ─────────────────────────────────────────
from pdf_legend_parser import parse_legend as new_parse_legend, LegendResult

# ── Import old parser from equipment_counter ─────────────────────────
from equipment_counter import (
    extract_text as old_extract_text,
    parse_legend as old_parse_legend,
    process_pdf as old_process_pdf,
    EquipmentItem,
)

# ── Test case definitions (copied from test_vor_accuracy.py) ─────────
DATA_DIR = PROJECT_ROOT / "Data"
DBT_DIR = DATA_DIR / "ДБТ разделы для ИИ"
_ABK = DBT_DIR / "01_АБК" / "Обновленные файлы"

TEST_CASES: list[dict[str, str]] = [
    {
        "name": "abk_eo",
        "ref_vor": str(_ABK / "ВОР_" / "ВОР АБК ЭО.xlsx"),
        "pdf_dir": str(_ABK / "PDF_" / "ЭО"),
    },
    {
        "name": "abk_em",
        "ref_vor": str(_ABK / "ВОР_" / "ВОР АБК ЭМ.xlsx"),
        "pdf_dir": str(_ABK / "PDF_" / "ЭМ"),
    },
    {
        "name": "abk_eg",
        "ref_vor": str(_ABK / "ВОР_" / "ВОР АБК ЭГ.xlsx"),
        "pdf_dir": str(_ABK / "PDF_" / "ЭГ"),
    },
    {
        "name": "pos_12_3",
        "ref_vor": str(DBT_DIR / "12.3 Насосная станция поверхностных стоков"
                       / "ВОР поз. 12.3.xlsx"),
        "pdf_dir": str(DBT_DIR / "12.3 Насосная станция поверхностных стоков"
                       / "02_PDF"),
    },
    {
        "name": "pos_16_2",
        "ref_vor": str(DBT_DIR / "16.2_ЖД КПП" / "Обновленные файлы 1"
                       / "ВОР поз. 16.2.xlsx"),
        "pdf_dir": str(DBT_DIR / "16.2_ЖД КПП" / "Обновленные файлы 1"
                       / "02_PDF"),
    },
    {
        "name": "pos_27",
        "ref_vor": str(DBT_DIR / "27_Склад вспомогательных материалов "
                       "с участком погрузки крытых вагонов"
                       / "Обновленные файлы" / "ВОР поз.27.xlsx"),
        "pdf_dir": str(DBT_DIR / "27_Склад вспомогательных материалов "
                       "с участком погрузки крытых вагонов"
                       / "Обновленные файлы" / "03_PDF"),
    },
    {
        "name": "pos_28",
        "ref_vor": str(DBT_DIR / "28. Автовесы" / "ВОР поз.28.xlsx"),
        "pdf_dir": str(DBT_DIR / "28. Автовесы" / "02_PDF"),
    },
    {
        "name": "pos_30",
        "ref_vor": str(DBT_DIR / "30. КПП" / "ВОР поз.30.xlsx"),
        "pdf_dir": str(DBT_DIR / "30. КПП" / "03_PDF"),
    },
    {
        "name": "pos_8_2_em",
        "ref_vor": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭМ" / "ВОР ЭМ поз. 8.2.xlsx"),
        "pdf_dir": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭМ" / "02_PDF"),
    },
    {
        "name": "pos_8_2_eo",
        "ref_vor": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭО" / "ВОР поз. 8.2.xlsx"),
        "pdf_dir": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭО" / "02_PDF"),
    },
]


# =====================================================================
#  Item quality classification
# =====================================================================

PLACEHOLDER_RE = re.compile(r"^\[.*\]$")  # [Equipment X], [Обозначение 3], etc.


def _is_placeholder(desc: str) -> bool:
    """Check if description is a placeholder like [Equipment X]."""
    return bool(PLACEHOLDER_RE.match(desc.strip()))


def _is_suspicious(desc: str) -> bool:
    """Check if description looks suspicious (too long, non-equipment text)."""
    if len(desc) > 150:
        return True
    # Check for non-equipment text patterns
    non_equip_patterns = [
        r"Лист\b", r"Подп\.", r"Дата\b", r"Разраб\.", r"Пров\.",
        r"Н\.контр", r"ГИП\b", r"Изм\.", r"Стадия\b",
        r"масштаб", r"формат",
    ]
    for pat in non_equip_patterns:
        if re.search(pat, desc, re.IGNORECASE):
            return True
    return False


# =====================================================================
#  Diagnose one PDF with the NEW parser
# =====================================================================

def diagnose_new_parser(pdf_path: str) -> dict:
    """Run new parse_legend and return diagnostic info."""
    result: dict = {
        "pdf": Path(pdf_path).name,
        "parser": "new_s017",
        "legend_found": False,
        "item_count": 0,
        "page_index": -1,
        "columns_detected": 0,
        "legend_bbox": None,
        "items": [],
        "error": None,
    }

    try:
        legend = new_parse_legend(pdf_path)
        result["legend_found"] = bool(legend.items)
        result["item_count"] = len(legend.items)
        result["page_index"] = legend.page_index
        result["columns_detected"] = legend.columns_detected
        result["legend_bbox"] = list(legend.legend_bbox)

        for item in legend.items:
            desc = item.description or ""
            item_info = {
                "symbol": item.symbol,
                "description": desc[:200],
                "description_length": len(desc),
                "has_description": bool(desc) and not _is_placeholder(desc),
                "is_placeholder": _is_placeholder(desc),
                "suspicious": _is_suspicious(desc),
                "category": item.category,
                "color": item.color,
                "bbox": list(item.bbox),
            }
            result["items"].append(item_info)

    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"

    return result


# =====================================================================
#  Diagnose one PDF with the OLD parser
# =====================================================================

def diagnose_old_parser(pdf_path: str) -> dict:
    """Run old parse_legend from equipment_counter.py and return diagnostic info."""
    result: dict = {
        "pdf": Path(pdf_path).name,
        "parser": "old_equipment_counter",
        "legend_found": False,
        "item_count": 0,
        "items": [],
        "error": None,
    }

    try:
        text = old_extract_text(pdf_path)
        if not text.strip():
            result["error"] = "No text extracted"
            return result

        legend: OrderedDict = old_parse_legend(pdf_path, text)
        result["legend_found"] = bool(legend)
        result["item_count"] = len(legend)

        for sym, desc in legend.items():
            desc_str = str(desc)
            item_info = {
                "symbol": sym,
                "description": desc_str[:200],
                "description_length": len(desc_str),
                "has_description": bool(desc_str) and not _is_placeholder(desc_str),
                "is_placeholder": _is_placeholder(desc_str),
                "suspicious": _is_suspicious(desc_str),
            }
            result["items"].append(item_info)

    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"

    return result


# =====================================================================
#  Deep inspection of a PDF (raw pdfplumber data)
# =====================================================================

def deep_inspect_pdf(pdf_path: str) -> dict:
    """Extract raw pdfplumber data for manual inspection of worst cases.

    Returns per-page info: text, lines, rects, words relevant to legend area.
    """
    inspection: dict = {
        "pdf": Path(pdf_path).name,
        "pages": [],
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                page_info: dict = {
                    "page_index": page_idx,
                    "width": float(page.width),
                    "height": float(page.height),
                    "has_legend_header": False,
                    "legend_header_text": "",
                    "legend_header_pos": None,
                    "full_text_excerpt": "",
                    "word_count": 0,
                    "line_count": 0,
                    "rect_count": 0,
                    "horizontal_lines": [],
                    "vertical_lines": [],
                    "legend_area_words": [],
                }

                words = page.extract_words(x_tolerance=3, y_tolerance=3) or []
                lines = page.lines or []
                rects = page.rects or []

                page_info["word_count"] = len(words)
                page_info["line_count"] = len(lines)
                page_info["rect_count"] = len(rects)

                # Full text excerpt
                full_text = page.extract_text() or ""
                page_info["full_text_excerpt"] = full_text[:2000]

                # Look for legend header
                header_y = None
                header_x = None
                for w in words:
                    if re.search(r"Условные|Легенда", w["text"], re.IGNORECASE):
                        page_info["has_legend_header"] = True
                        page_info["legend_header_text"] = w["text"]
                        page_info["legend_header_pos"] = {
                            "x0": round(w["x0"], 1),
                            "top": round(w["top"], 1),
                            "x1": round(w["x1"], 1),
                            "bottom": round(w["bottom"], 1),
                        }
                        header_y = w["top"]
                        header_x = w["x0"]
                        break

                if header_y is not None:
                    # Collect long horizontal lines near legend
                    h_lines = [
                        {
                            "x0": round(l["x0"], 1),
                            "x1": round(l["x1"], 1),
                            "y": round(l["top"], 1),
                            "length": round(abs(l["x1"] - l["x0"]), 1),
                        }
                        for l in lines
                        if abs(l["top"] - l["bottom"]) < 2
                        and abs(l["x1"] - l["x0"]) > 50
                        and header_y - 30 < l["top"] < header_y + 600
                    ]
                    page_info["horizontal_lines"] = sorted(h_lines, key=lambda x: x["y"])[:30]

                    # Collect long vertical lines near legend
                    v_lines = [
                        {
                            "x": round(l["x0"], 1),
                            "y0": round(l["top"], 1),
                            "y1": round(l["bottom"], 1),
                            "length": round(abs(l["bottom"] - l["top"]), 1),
                        }
                        for l in lines
                        if abs(l["x0"] - l["x1"]) < 2
                        and abs(l["bottom"] - l["top"]) > 50
                        and header_x - 300 < l["x0"] < header_x + 800
                        and l["top"] < header_y + 600
                    ]
                    page_info["vertical_lines"] = sorted(v_lines, key=lambda x: x["x"])[:20]

                    # All words in the legend area
                    area_words = [
                        {
                            "text": w["text"],
                            "x0": round(w["x0"], 1),
                            "top": round(w["top"], 1),
                            "x1": round(w["x1"], 1),
                            "bottom": round(w["bottom"], 1),
                        }
                        for w in words
                        if header_y - 10 < w["top"] < header_y + 600
                        and header_x - 300 < w["x0"] < header_x + 800
                    ]
                    page_info["legend_area_words"] = sorted(
                        area_words, key=lambda w: (w["top"], w["x0"])
                    )[:100]

                inspection["pages"].append(page_info)

    except Exception as exc:
        inspection["error"] = f"{type(exc).__name__}: {exc}"

    return inspection


# =====================================================================
#  Main diagnostic loop
# =====================================================================

def main():
    print("=" * 80)
    print("  LEGEND PARSER DIAGNOSTIC -- All 10 Test Cases")
    print("=" * 80)
    print()

    report: dict = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "summary": {},
        "per_case": [],
        "worst_cases_inspection": [],
    }

    # Global counters
    total_pdfs = 0
    total_new_legend_found = 0
    total_new_legend_not_found = 0
    total_new_items = 0
    total_new_good_desc = 0
    total_new_placeholder = 0
    total_new_suspicious = 0
    total_new_no_desc = 0
    total_new_errors = 0

    total_old_legend_found = 0
    total_old_items = 0
    total_old_good_desc = 0
    total_old_placeholder = 0
    total_old_errors = 0

    # Track worst cases for deep inspection
    worst_cases: list[tuple[str, str, str]] = []  # (case_name, pdf_path, reason)

    for case_idx, case in enumerate(TEST_CASES):
        case_name = case["name"]
        pdf_dir = Path(case["pdf_dir"])

        print(f"\n{'_' * 70}")
        print(f"[{case_idx + 1}/{len(TEST_CASES)}] Case: {case_name}")
        print(f"  PDF dir: {pdf_dir}")

        if not pdf_dir.exists():
            print(f"  SKIP: PDF directory not found")
            report["per_case"].append({
                "name": case_name,
                "status": "skip",
                "reason": "PDF directory not found",
            })
            continue

        pdf_files = sorted(pdf_dir.glob("*.pdf"))
        if not pdf_files:
            print(f"  SKIP: No PDF files in directory")
            report["per_case"].append({
                "name": case_name,
                "status": "skip",
                "reason": "No PDF files",
            })
            continue

        print(f"  PDF files: {len(pdf_files)}")

        case_report: dict = {
            "name": case_name,
            "status": "ok",
            "pdf_dir": str(pdf_dir),
            "pdf_count": len(pdf_files),
            "pdfs": [],
            "new_parser_summary": {},
            "old_parser_summary": {},
        }

        # Per-case counters (new)
        case_new_found = 0
        case_new_not_found = 0
        case_new_items = 0
        case_new_good = 0
        case_new_placeholder = 0
        case_new_suspicious = 0
        case_new_no_desc = 0
        case_new_errors = 0

        # Per-case counters (old)
        case_old_found = 0
        case_old_items = 0
        case_old_good = 0
        case_old_placeholder = 0
        case_old_errors = 0

        for pdf_path in pdf_files:
            total_pdfs += 1
            pdf_str = str(pdf_path)

            # ── New parser ──
            new_diag = diagnose_new_parser(pdf_str)

            if new_diag["error"]:
                case_new_errors += 1
                total_new_errors += 1
                status_new = f"ERROR: {new_diag['error'][:80]}"
            elif new_diag["legend_found"]:
                case_new_found += 1
                total_new_legend_found += 1
                status_new = f"OK: {new_diag['item_count']} items"
            else:
                case_new_not_found += 1
                total_new_legend_not_found += 1
                status_new = "NO LEGEND"

            # Count item quality (new)
            for it in new_diag["items"]:
                case_new_items += 1
                total_new_items += 1
                if it["is_placeholder"]:
                    case_new_placeholder += 1
                    total_new_placeholder += 1
                elif it["suspicious"]:
                    case_new_suspicious += 1
                    total_new_suspicious += 1
                elif it["has_description"]:
                    case_new_good += 1
                    total_new_good_desc += 1
                else:
                    case_new_no_desc += 1
                    total_new_no_desc += 1

            # ── Old parser ──
            old_diag = diagnose_old_parser(pdf_str)

            if old_diag["error"]:
                case_old_errors += 1
                total_old_errors += 1
                status_old = f"ERROR: {old_diag['error'][:80]}"
            elif old_diag["legend_found"]:
                case_old_found += 1
                total_old_legend_found += 1
                status_old = f"OK: {old_diag['item_count']} items"
            else:
                status_old = "NO LEGEND"

            for it in old_diag["items"]:
                case_old_items += 1
                total_old_items += 1
                if it["is_placeholder"]:
                    case_old_placeholder += 1
                    total_old_placeholder += 1
                elif it["has_description"]:
                    case_old_good += 1
                    total_old_good_desc += 1

            # Track worst cases -- focus on REGRESSIONS (old found, new didn't)
            # and cases where new parser finds legend but all items are placeholders
            if new_diag["legend_found"] and new_diag["item_count"] > 0:
                placeholders = sum(1 for it in new_diag["items"] if it["is_placeholder"])
                if placeholders > 0 and placeholders == new_diag["item_count"]:
                    worst_cases.append((case_name, pdf_str, "all_placeholder"))
            if not new_diag["legend_found"] and old_diag["legend_found"]:
                # REGRESSION: old parser found legend, new didn't
                worst_cases.append((case_name, pdf_str, "REGRESSION_old_found_new_not"))

            # Print per-PDF status
            print(f"    {pdf_path.name:50s}  new={status_new:30s}  old={status_old}")

            case_report["pdfs"].append({
                "pdf": pdf_path.name,
                "new": new_diag,
                "old": old_diag,
            })

        # Per-case summary
        case_report["new_parser_summary"] = {
            "legend_found": case_new_found,
            "legend_not_found": case_new_not_found,
            "errors": case_new_errors,
            "total_items": case_new_items,
            "good_descriptions": case_new_good,
            "placeholder_descriptions": case_new_placeholder,
            "suspicious_descriptions": case_new_suspicious,
            "no_description": case_new_no_desc,
        }
        case_report["old_parser_summary"] = {
            "legend_found": case_old_found,
            "errors": case_old_errors,
            "total_items": case_old_items,
            "good_descriptions": case_old_good,
            "placeholder_descriptions": case_old_placeholder,
        }

        print(f"\n  NEW parser summary:")
        print(f"    Legend found:     {case_new_found}/{len(pdf_files)}")
        print(f"    Legend NOT found: {case_new_not_found}/{len(pdf_files)}")
        print(f"    Errors:          {case_new_errors}")
        print(f"    Total items:     {case_new_items}")
        print(f"    Good descs:      {case_new_good}")
        print(f"    Placeholders:    {case_new_placeholder}")
        print(f"    Suspicious:      {case_new_suspicious}")
        print(f"    No description:  {case_new_no_desc}")

        print(f"  OLD parser summary:")
        print(f"    Legend found:     {case_old_found}/{len(pdf_files)}")
        print(f"    Errors:          {case_old_errors}")
        print(f"    Total items:     {case_old_items}")
        print(f"    Good descs:      {case_old_good}")
        print(f"    Placeholders:    {case_old_placeholder}")

        report["per_case"].append(case_report)

    # ── Overall summary ──
    print(f"\n{'=' * 80}")
    print("  OVERALL SUMMARY")
    print(f"{'=' * 80}")

    print(f"\n  Total PDFs processed: {total_pdfs}")
    print(f"\n  --- NEW parser (pdf_legend_parser.py) ---")
    print(f"  Legend found:         {total_new_legend_found}/{total_pdfs}")
    print(f"  Legend NOT found:     {total_new_legend_not_found}/{total_pdfs}")
    print(f"  Errors:               {total_new_errors}")
    print(f"  Total items:          {total_new_items}")
    print(f"  Good descriptions:    {total_new_good_desc}")
    print(f"  Placeholder descs:    {total_new_placeholder}")
    print(f"  Suspicious descs:     {total_new_suspicious}")
    print(f"  No description:       {total_new_no_desc}")

    print(f"\n  --- OLD parser (equipment_counter.py) ---")
    print(f"  Legend found:         {total_old_legend_found}/{total_pdfs}")
    print(f"  Errors:               {total_old_errors}")
    print(f"  Total items:          {total_old_items}")
    print(f"  Good descriptions:    {total_old_good_desc}")
    print(f"  Placeholder descs:    {total_old_placeholder}")

    report["summary"] = {
        "total_pdfs": total_pdfs,
        "new_parser": {
            "legend_found": total_new_legend_found,
            "legend_not_found": total_new_legend_not_found,
            "errors": total_new_errors,
            "total_items": total_new_items,
            "good_descriptions": total_new_good_desc,
            "placeholder_descriptions": total_new_placeholder,
            "suspicious_descriptions": total_new_suspicious,
            "no_description": total_new_no_desc,
        },
        "old_parser": {
            "legend_found": total_old_legend_found,
            "errors": total_old_errors,
            "total_items": total_old_items,
            "good_descriptions": total_old_good_desc,
            "placeholder_descriptions": total_old_placeholder,
        },
    }

    # ── Deep inspection of worst cases ──
    print(f"\n{'=' * 80}")
    print("  DEEP INSPECTION OF WORST CASES")
    print(f"{'=' * 80}")

    # Limit deep inspection to first 8 worst cases (prioritize regressions)
    regressions = [wc for wc in worst_cases if wc[2] == "REGRESSION_old_found_new_not"]
    others = [wc for wc in worst_cases if wc[2] != "REGRESSION_old_found_new_not"]
    inspect_cases = (regressions + others)[:8]
    if not inspect_cases:
        print("  No worst cases detected (all PDFs have legends with items).")
    else:
        for case_name, pdf_path, reason in inspect_cases:
            print(f"\n  --- {case_name}: {Path(pdf_path).name} (reason: {reason}) ---")
            inspection = deep_inspect_pdf(pdf_path)
            report["worst_cases_inspection"].append({
                "case": case_name,
                "reason": reason,
                "inspection": inspection,
            })

            for pg in inspection.get("pages", []):
                if pg.get("has_legend_header"):
                    print(f"  Page {pg['page_index']}: LEGEND HEADER FOUND")
                    print(f"    Header text: {pg['legend_header_text']}")
                    print(f"    Header pos: {pg['legend_header_pos']}")
                    print(f"    Words in area: {len(pg.get('legend_area_words', []))}")
                    print(f"    H-lines near legend: {len(pg.get('horizontal_lines', []))}")
                    print(f"    V-lines near legend: {len(pg.get('vertical_lines', []))}")

                    # Print horizontal lines for debugging table detection
                    if pg.get("horizontal_lines"):
                        print(f"    Horizontal lines (y, x0-x1, length):")
                        for hl in pg["horizontal_lines"][:15]:
                            print(f"      y={hl['y']:7.1f}  x=[{hl['x0']:7.1f} - {hl['x1']:7.1f}]  len={hl['length']:6.1f}")

                    if pg.get("vertical_lines"):
                        print(f"    Vertical lines (x, y0-y1, length):")
                        for vl in pg["vertical_lines"][:10]:
                            print(f"      x={vl['x']:7.1f}  y=[{vl['y0']:7.1f} - {vl['y1']:7.1f}]  len={vl['length']:6.1f}")

                    # Print words in the legend area
                    if pg.get("legend_area_words"):
                        print(f"    Words in legend area (first 40):")
                        for aw in pg["legend_area_words"][:40]:
                            print(f"      [{aw['x0']:7.1f}, {aw['top']:7.1f}]  \"{aw['text']}\"")

                    # Also print full text excerpt around legend
                    full_text = pg.get("full_text_excerpt", "")
                    legend_pos = full_text.lower().find("условн")
                    if legend_pos >= 0:
                        start = max(0, legend_pos - 100)
                        end = min(len(full_text), legend_pos + 800)
                        print(f"    Full text around legend:")
                        for line in full_text[start:end].split("\n"):
                            print(f"      | {line}")
                else:
                    print(f"  Page {pg['page_index']}: no legend header")
                    print(f"    Words: {pg['word_count']}, Lines: {pg['line_count']}, Rects: {pg['rect_count']}")

    # ── Per-case quality breakdown ──
    print(f"\n{'=' * 80}")
    print("  PER-CASE QUALITY BREAKDOWN (NEW PARSER)")
    print(f"{'=' * 80}")
    print(f"  {'Case':<16s}  {'PDFs':>5s}  {'Found':>5s}  {'!Found':>6s}  "
          f"{'Items':>5s}  {'Good':>5s}  {'Plchd':>5s}  {'Susp':>5s}  {'Empty':>5s}")
    print(f"  {'-'*16}  {'-'*5}  {'-'*5}  {'-'*6}  "
          f"{'-'*5}  {'-'*5}  {'-'*5}  {'-'*5}  {'-'*5}")

    for cr in report["per_case"]:
        if cr.get("status") == "skip":
            print(f"  {cr['name']:<16s}  {'SKIP':>5s}")
            continue
        ns = cr["new_parser_summary"]
        print(f"  {cr['name']:<16s}  {cr['pdf_count']:>5d}  {ns['legend_found']:>5d}"
              f"  {ns['legend_not_found']:>6d}  {ns['total_items']:>5d}"
              f"  {ns['good_descriptions']:>5d}  {ns['placeholder_descriptions']:>5d}"
              f"  {ns['suspicious_descriptions']:>5d}  {ns['no_description']:>5d}")

    # ── Items with problems (detailed listing) ──
    print(f"\n{'=' * 80}")
    print("  PROBLEM ITEMS (placeholders, suspicious, no description)")
    print(f"{'=' * 80}")

    for cr in report["per_case"]:
        if cr.get("status") == "skip":
            continue
        case_problems = []
        for pdf_entry in cr.get("pdfs", []):
            for it in pdf_entry["new"]["items"]:
                if it["is_placeholder"] or it["suspicious"] or not it["has_description"]:
                    case_problems.append({
                        "pdf": pdf_entry["pdf"],
                        **it,
                    })
        if case_problems:
            print(f"\n  [{cr['name']}] -- {len(case_problems)} problem items:")
            for p in case_problems[:20]:
                kind = ("PLACEHOLDER" if p["is_placeholder"]
                        else "SUSPICIOUS" if p["suspicious"]
                        else "NO_DESC")
                sym_display = p["symbol"] if p["symbol"] else "(none)"
                print(f"    {p['pdf']:40s}  sym={sym_display:>6s}  {kind:12s}  "
                      f"\"{p['description'][:80]}\"")

    # ── Save report ──
    report_path = PROJECT_ROOT / "diagnose_legend_report.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"\n  Report saved: {report_path}")
    print("=" * 80)


if __name__ == "__main__":
    main()
