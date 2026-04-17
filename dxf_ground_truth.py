#!/usr/bin/env python3
"""dxf_ground_truth.py -- Use DXF parser as ground truth to evaluate PDF parser accuracy.

For test cases that have both DXF and PDF versions of the same drawing,
runs both parsers and compares results. DXF parser output is treated as
the "correct answer" since it has access to structured CAD data.

Produces:
  - Per-file DXF vs PDF comparison (items found, name matches, count deltas)
  - Aggregate statistics across all matched pairs
  - JSON report: dxf_ground_truth_report.json

Usage:
    python dxf_ground_truth.py
"""

from __future__ import annotations

import json
import re
import sys
import time
import difflib
from dataclasses import dataclass, field, asdict
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

from equipment_counter import process_dxf, EquipmentItem  # noqa: E402
from pdf_legend_parser import parse_legend  # noqa: E402
from pdf_count_text import count_symbols  # noqa: E402

DATA_DIR = PROJECT_ROOT / "Data"
DBT_DIR = DATA_DIR / "ДБТ разделы для ИИ"

# ── File matching: DXF ↔ PDF pairs ──────────────────────────────────
# Each entry: (case_name, section, dxf_dir, pdf_dir, pairs)
# pairs: list of (dxf_filename, pdf_filename, drawing_number)

_ABK = DBT_DIR / "01_АБК"
_ABK_UPD = _ABK / "Обновленные файлы"


def _find_dxf_files(base_dir: Path) -> list[Path]:
    """Find all .dxf files in a directory (non-recursive)."""
    if not base_dir.exists():
        return []
    return sorted(base_dir.glob("*.dxf"))


def _find_pdf_files(base_dir: Path) -> list[Path]:
    """Find all .pdf files in a directory (non-recursive)."""
    if not base_dir.exists():
        return []
    return sorted(base_dir.glob("*.pdf"))


def _extract_drawing_number(filename: str) -> str:
    """Extract the leading drawing number from a filename.

    Examples:
        '006 - План освещения на отм- 0-000.dxf' -> '006'
        '006-План освещения на отм. 0.000.pdf'   -> '006'
        '013 - План расположения ...'             -> '013'
    """
    m = re.match(r"^(\d{3})", filename)
    if m:
        return m.group(1)
    return ""


def _match_dxf_pdf_pairs(
    dxf_dir: Path, pdf_dir: Path
) -> list[tuple[Path, Path, str]]:
    """Match DXF files to PDF files by drawing number prefix.

    Returns list of (dxf_path, pdf_path, drawing_number).
    """
    dxf_files = _find_dxf_files(dxf_dir)
    pdf_files = _find_pdf_files(pdf_dir)

    # Build lookup: drawing_number -> pdf_path
    pdf_by_num: dict[str, Path] = {}
    for p in pdf_files:
        num = _extract_drawing_number(p.name)
        if num and num not in pdf_by_num:
            pdf_by_num[num] = p

    pairs = []
    for dxf in dxf_files:
        num = _extract_drawing_number(dxf.name)
        if num and num in pdf_by_num:
            pairs.append((dxf, pdf_by_num[num], num))

    return pairs


# ── Test case definitions ────────────────────────────────────────────

MATCH_CASES: list[dict] = [
    {
        "name": "abk_eo",
        "section": "ЭО",
        "dxf_dir": str(_ABK_UPD / "DWG" / "ЭО" / "_converted_dxf"),
        "pdf_dir": str(_ABK_UPD / "PDF_" / "ЭО"),
    },
    {
        "name": "abk_em",
        "section": "ЭМ",
        "dxf_dir": str(_ABK_UPD / "DWG" / "ЭМ" / "_converted_dxf"),
        "pdf_dir": str(_ABK_UPD / "PDF_" / "ЭМ"),
    },
    {
        "name": "abk_eg",
        "section": "ЭГ",
        "dxf_dir": str(_ABK_UPD / "DWG" / "ЭГ" / "_converted_dxf"),
        "pdf_dir": str(_ABK_UPD / "PDF_" / "ЭГ"),
    },
    {
        "name": "pos_12_3",
        "section": "ЭО",
        "dxf_dir": str(DBT_DIR / "12.3 Насосная станция поверхностных стоков"
                       / "01_DWG" / "_converted_dxf"),
        "pdf_dir": str(DBT_DIR / "12.3 Насосная станция поверхностных стоков"
                       / "02_PDF"),
    },
    {
        "name": "pos_27",
        "section": "ЭОМ",
        "dxf_dir": str(DBT_DIR / "27_Склад вспомогательных материалов "
                       "с участком погрузки крытых вагонов"
                       / "Обновленные файлы" / "02_DWG" / "_converted_dxf"),
        "pdf_dir": str(DBT_DIR / "27_Склад вспомогательных материалов "
                       "с участком погрузки крытых вагонов"
                       / "Обновленные файлы" / "03_PDF"),
    },
    {
        "name": "pos_28",
        "section": "ЭО",
        "dxf_dir": str(DBT_DIR / "28. Автовесы" / "01_DWG" / "_converted_dxf"),
        "pdf_dir": str(DBT_DIR / "28. Автовесы" / "02_PDF"),
    },
    {
        "name": "pos_30",
        "section": "ЭО",
        "dxf_dir": str(DBT_DIR / "30. КПП" / "02_DWG" / "_converted_dxf"),
        "pdf_dir": str(DBT_DIR / "30. КПП" / "03_PDF"),
    },
    {
        "name": "pos_8_2_eo",
        "section": "ЭО",
        "dxf_dir": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭО" / "01_DWG" / "_converted_dxf"),
        "pdf_dir": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭО" / "02_PDF"),
    },
    {
        "name": "pos_8_2_em",
        "section": "ЭМ",
        "dxf_dir": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭМ" / "Обновленные файлы" / "01_DWG" / "_converted_dxf"),
        "pdf_dir": str(DBT_DIR / "8.2 Участок хранения сульфата аммония"
                       / "ЭМ" / "02_PDF"),
    },
    {
        "name": "gpk_z3",
        "section": "ЭО",
        "dxf_dir": str(DBT_DIR / "03_ГПК_" / "3-я захватка" / "_converted_dxf" / "01_DWG"),
        "pdf_dir": str(DBT_DIR / "03_ГПК_" / "3-я захватка" / "02_PDF"),
    },
]


# ── Comparison logic ─────────────────────────────────────────────────

def _fuzzy_match(a: str, b: str, threshold: float = 0.50) -> float:
    """Fuzzy match ratio between two equipment name strings."""
    a_clean = re.sub(r"[\s\-_.,;:()]+", " ", a.lower()).strip()
    b_clean = re.sub(r"[\s\-_.,;:()]+", " ", b.lower()).strip()
    return difflib.SequenceMatcher(None, a_clean, b_clean).ratio()


def _run_pdf_pipeline(pdf_path: str) -> list[dict]:
    """Run PDF legend + count pipeline, return list of {symbol, name, count}."""
    legend = parse_legend(pdf_path)
    if not legend or not legend.items:
        return []

    count_result = count_symbols(pdf_path, legend)
    counts = count_result.counts if count_result else {}

    results = []
    for item in legend.items:
        sym = item.symbol
        name = item.description
        cnt = counts.get(sym, 0)
        if sym or name:
            results.append({
                "symbol": sym,
                "name": name,
                "count": cnt,
                "color": item.color,
            })
    return results


def _run_dxf_pipeline(dxf_path: str) -> list[dict]:
    """Run DXF parser, return list of {symbol, name, count}."""
    items = process_dxf(dxf_path)
    results = []
    for item in items:
        total = item.count + item.count_ae
        results.append({
            "symbol": item.symbol,
            "name": item.name,
            "count": total,
        })
    return results


def compare_results(
    dxf_items: list[dict],
    pdf_items: list[dict],
) -> dict:
    """Compare DXF (ground truth) vs PDF results.

    Returns comparison dict with matches, misses, extras, and deltas.
    """
    # Build name-based matching between DXF and PDF items
    # Strategy: for each DXF item, find best fuzzy match in PDF items
    dxf_matched = set()
    pdf_matched = set()
    matches = []

    for i, dxf_it in enumerate(dxf_items):
        dxf_name = dxf_it["name"]
        if not dxf_name or dxf_name.startswith("[?"):
            continue

        best_ratio = 0.0
        best_j = -1
        for j, pdf_it in enumerate(pdf_items):
            if j in pdf_matched:
                continue
            pdf_name = pdf_it["name"]
            if not pdf_name:
                continue

            ratio = _fuzzy_match(dxf_name, pdf_name)
            if ratio > best_ratio:
                best_ratio = ratio
                best_j = j

        if best_ratio >= 0.45 and best_j >= 0:
            dxf_matched.add(i)
            pdf_matched.add(best_j)
            pdf_it = pdf_items[best_j]
            matches.append({
                "dxf_symbol": dxf_it["symbol"],
                "pdf_symbol": pdf_it["symbol"],
                "dxf_name": dxf_name,
                "pdf_name": pdf_it["name"],
                "dxf_count": dxf_it["count"],
                "pdf_count": pdf_it["count"],
                "name_ratio": round(best_ratio, 3),
                "count_match": dxf_it["count"] == pdf_it["count"],
                "count_delta": pdf_it["count"] - dxf_it["count"],
            })

    # Items in DXF but not matched in PDF
    missing = []
    for i, dxf_it in enumerate(dxf_items):
        if i not in dxf_matched and dxf_it["name"] and not dxf_it["name"].startswith("[?"):
            missing.append({
                "symbol": dxf_it["symbol"],
                "name": dxf_it["name"],
                "count": dxf_it["count"],
            })

    # Items in PDF but not matched in DXF
    extra = []
    for j, pdf_it in enumerate(pdf_items):
        if j not in pdf_matched and pdf_it["name"]:
            extra.append({
                "symbol": pdf_it["symbol"],
                "name": pdf_it["name"],
                "count": pdf_it["count"],
            })

    name_match_count = len(matches)
    exact_count = sum(1 for m in matches if m["count_match"])
    total_dxf = len([d for d in dxf_items if d["name"] and not d["name"].startswith("[?")])

    return {
        "dxf_item_count": total_dxf,
        "pdf_item_count": len(pdf_items),
        "name_matches": name_match_count,
        "exact_matches": exact_count,
        "name_match_rate": round(name_match_count / total_dxf * 100, 1) if total_dxf else 0,
        "exact_match_rate": round(exact_count / total_dxf * 100, 1) if total_dxf else 0,
        "missing_in_pdf": len(missing),
        "extra_in_pdf": len(extra),
        "matched_items": matches,
        "missing_items": missing[:10],
        "extra_items": extra[:10],
    }


# ── Main ─────────────────────────────────────────────────────────────

def main():
    print("=" * 80)
    print("  DXF GROUND TRUTH: DXF vs PDF Parser Comparison")
    print("=" * 80)
    print()

    all_results = []
    totals = {
        "pairs": 0, "dxf_items": 0, "pdf_items": 0,
        "name_matches": 0, "exact_matches": 0,
        "missing": 0, "extra": 0,
    }

    for case in MATCH_CASES:
        case_name = case["name"]
        dxf_dir = Path(case["dxf_dir"])
        pdf_dir = Path(case["pdf_dir"])

        if not dxf_dir.exists():
            continue

        pairs = _match_dxf_pdf_pairs(dxf_dir, pdf_dir)
        if not pairs:
            continue

        print(f"{'_' * 70}")
        print(f"[{case_name}] {case['section']}  ({len(pairs)} matched pairs)")
        print(f"  DXF: {dxf_dir}")
        print(f"  PDF: {pdf_dir}")

        case_result = {
            "name": case_name,
            "section": case["section"],
            "pairs": [],
        }

        for dxf_path, pdf_path, drawing_num in pairs:
            print(f"\n  --- Drawing {drawing_num}: {dxf_path.stem}")

            # Run DXF parser
            try:
                dxf_items = _run_dxf_pipeline(str(dxf_path))
            except Exception as e:
                print(f"    DXF ERROR: {e}")
                dxf_items = []

            # Run PDF parser
            try:
                pdf_items = _run_pdf_pipeline(str(pdf_path))
            except Exception as e:
                print(f"    PDF ERROR: {e}")
                pdf_items = []

            # Compare
            comparison = compare_results(dxf_items, pdf_items)

            print(f"    DXF: {comparison['dxf_item_count']} items  |  "
                  f"PDF: {comparison['pdf_item_count']} items")
            print(f"    Name matches: {comparison['name_matches']}  |  "
                  f"Exact (name+count): {comparison['exact_matches']}  |  "
                  f"Rate: {comparison['name_match_rate']}%")

            if comparison["matched_items"]:
                for m in comparison["matched_items"][:5]:
                    delta_str = f"delta={m['count_delta']:+d}" if m['count_delta'] else "="
                    print(f"      ✓ [{m['dxf_symbol']:4s}→{m['pdf_symbol']:4s}] "
                          f"ratio={m['name_ratio']:.0%}  "
                          f"dxf={m['dxf_count']}  pdf={m['pdf_count']}  {delta_str}")
                    print(f"        DXF: {m['dxf_name'][:55]}")
                    print(f"        PDF: {m['pdf_name'][:55]}")

            if comparison["missing_items"]:
                print(f"    Missing in PDF ({comparison['missing_in_pdf']}):")
                for m in comparison["missing_items"][:3]:
                    print(f"      ✗ [{m['symbol']}] {m['name'][:50]}  (count={m['count']})")

            pair_data = {
                "drawing_num": drawing_num,
                "dxf_file": dxf_path.name,
                "pdf_file": pdf_path.name,
                "dxf_items": dxf_items,
                "pdf_items": pdf_items,
                "comparison": comparison,
            }
            case_result["pairs"].append(pair_data)

            totals["pairs"] += 1
            totals["dxf_items"] += comparison["dxf_item_count"]
            totals["pdf_items"] += comparison["pdf_item_count"]
            totals["name_matches"] += comparison["name_matches"]
            totals["exact_matches"] += comparison["exact_matches"]
            totals["missing"] += comparison["missing_in_pdf"]
            totals["extra"] += comparison["extra_in_pdf"]

        all_results.append(case_result)
        print()

    # ── Overall summary ──
    print("=" * 80)
    print("  OVERALL SUMMARY")
    print("=" * 80)
    print(f"  Total matched DXF↔PDF pairs: {totals['pairs']}")
    print(f"  Total DXF items (ground truth): {totals['dxf_items']}")
    print(f"  Total PDF items (generated):    {totals['pdf_items']}")
    print()
    nm_rate = (totals["name_matches"] / totals["dxf_items"] * 100
               if totals["dxf_items"] else 0)
    ex_rate = (totals["exact_matches"] / totals["dxf_items"] * 100
               if totals["dxf_items"] else 0)
    print(f"  Name match rate:  {totals['name_matches']}/{totals['dxf_items']}"
          f"  = {nm_rate:.1f}%")
    print(f"  Exact match rate: {totals['exact_matches']}/{totals['dxf_items']}"
          f"  = {ex_rate:.1f}%")
    print(f"  Missing in PDF:   {totals['missing']}")
    print(f"  Extra in PDF:     {totals['extra']}")

    # ── Per-case summary table ──
    print()
    print(f"  {'Case':<14} {'Pairs':>5} {'DXF':>5} {'PDF':>5} "
          f"{'NmMt':>5} {'ExMt':>5} {'NmRt%':>6} {'ExRt%':>6}")
    print(f"  {'-'*14} {'-'*5} {'-'*5} {'-'*5} {'-'*5} {'-'*5} {'-'*6} {'-'*6}")
    for cr in all_results:
        c_pairs = len(cr["pairs"])
        c_dxf = sum(p["comparison"]["dxf_item_count"] for p in cr["pairs"])
        c_pdf = sum(p["comparison"]["pdf_item_count"] for p in cr["pairs"])
        c_nm = sum(p["comparison"]["name_matches"] for p in cr["pairs"])
        c_ex = sum(p["comparison"]["exact_matches"] for p in cr["pairs"])
        c_nmr = c_nm / c_dxf * 100 if c_dxf else 0
        c_exr = c_ex / c_dxf * 100 if c_dxf else 0
        print(f"  {cr['name']:<14} {c_pairs:>5} {c_dxf:>5} {c_pdf:>5} "
              f"{c_nm:>5} {c_ex:>5} {c_nmr:>5.1f}% {c_exr:>5.1f}%")

    # ── Save report ──
    report = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "summary": totals,
        "summary_rates": {
            "name_match_rate_pct": round(nm_rate, 1),
            "exact_match_rate_pct": round(ex_rate, 1),
        },
        "cases": all_results,
    }

    report_path = PROJECT_ROOT / "dxf_ground_truth_report.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  Report saved: {report_path}")
    print("=" * 80)


if __name__ == "__main__":
    main()
