#!/usr/bin/env python3
"""
run_T122_gpk3_count_only.py — S011-02 executor script.

Запускает текущий пайплайн EquipmentCounter.process_pdf() на всех PDF
из 'Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/02_PDF/' и формирует:

  • equipment_report.json (сводный отчёт по всем 32 PDF)
  • gpk3_software_count_only.xlsx (collapse по высоте, как T049/T121)
  • T122_run.log (полный лог с commit hash и timestamps)

Collapse-правила (identical normalization rules to T121):
  • Убрать из имени маркеры строительной высоты: «отм. 0.000», «+4.200»,
    «отм. +13.800», «на отм. ...».
  • Убрать висячие пробелы/запятые после нормализации.
  • Свернуть строки с одинаковым (name_normalized, unit), суммируя count.
  • Сохранить список original_names как list для трассировки.

Запуск:
    python run_T122_gpk3_count_only.py
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import time
import traceback
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

# UTF-8 stdout per .tayfa language policy (KB-006)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from openpyxl import Workbook  # noqa: E402

# Импорт публичного API; не модифицируем equipment_counter.py.
import equipment_counter as ec  # noqa: E402


# ────────────────────────────────────────────────────────────────────────────
# Пути
# ────────────────────────────────────────────────────────────────────────────

REPO = Path(__file__).resolve().parent
PDF_DIR = REPO / "Data" / "ДБТ разделы для ИИ" / "03_ГПК_" / "3-я захватка" / "02_PDF"
OUT_DIR = REPO / "Data" / "ДБТ разделы для ИИ" / "03_ГПК_" / "3-я захватка"

REPORT_JSON = OUT_DIR / "equipment_report.json"
OUT_XLSX = OUT_DIR / "gpk3_software_count_only.xlsx"
RUN_LOG = OUT_DIR / "T122_run.log"


# ────────────────────────────────────────────────────────────────────────────
# Нормализация имён (collapse по высоте)
# ────────────────────────────────────────────────────────────────────────────

# Маркеры строительной высоты в имени файла / описания
# Примеры: «отм. 0.000», «отм. +4.200», «на отм. +13.800», «отм +28.200»
_HEIGHT_PATTERNS = [
    re.compile(r"\s*на\s*отм\.?\s*[+-]?\d+[.,]\d+\s*", re.IGNORECASE),
    re.compile(r"\s*отм\.?\s*[+-]?\d+[.,]\d+\s*", re.IGNORECASE),
    re.compile(r"\s*[+-]?\d+[.,]\d{2,3}\s*", re.IGNORECASE),  # bare elevation
]
_DRAWING_PATTERNS = [
    re.compile(r"^\s*\d{3,4}\s*[-—]\s*", re.IGNORECASE),  # ведущий «005 - »
    re.compile(r"\.pdf\s*$", re.IGNORECASE),
    re.compile(r"планы?\s+(освещения|привязк[аи]|кабеленесущ\w*)", re.IGNORECASE),
    re.compile(r"гпк|щитово[ея]", re.IGNORECASE),
]


def normalize_name(name: str) -> str:
    """Срезает маркеры высоты и нормализует пробелы.

    Дополнительно: trim trailing ", ", " -", скобок-пустышек.
    """
    s = name or ""
    for rx in _HEIGHT_PATTERNS:
        s = rx.sub(" ", s)
    # Зачистка типовых хвостов и разделителей вокруг высоты
    s = re.sub(r"[,;]\s*[,;]+", ",", s)
    s = re.sub(r"\(\s*\)", "", s)
    s = re.sub(r"\s*[-—]\s*$", "", s)
    s = re.sub(r"\s+", " ", s).strip(" ,.-—–")
    return s


def normalize_drawing_token(s: str) -> str:
    """Нормализация имени PDF-листа для логов."""
    out = s
    for rx in _DRAWING_PATTERNS:
        out = rx.sub(" ", out)
    return re.sub(r"\s+", " ", out).strip(" -—,.")


# ────────────────────────────────────────────────────────────────────────────
# Pipeline meta
# ────────────────────────────────────────────────────────────────────────────

def _git_info() -> dict:
    def _run(args: list[str]) -> str:
        try:
            return subprocess.check_output(args, cwd=str(REPO),
                                            stderr=subprocess.DEVNULL).decode().strip()
        except Exception:
            return ""
    return {
        "commit": _run(["git", "rev-parse", "HEAD"]),
        "short":  _run(["git", "rev-parse", "--short", "HEAD"]),
        "branch": _run(["git", "rev-parse", "--abbrev-ref", "HEAD"]),
        "status_porcelain": _run(["git", "status", "--porcelain"]).splitlines()[:10],
    }


# ────────────────────────────────────────────────────────────────────────────
# Запуск пайплайна
# ────────────────────────────────────────────────────────────────────────────

def run_all(pdf_paths: list[Path], log) -> dict:
    """Возвращает структурированный отчёт по всем PDF."""
    per_file: list[dict] = []
    t0 = time.time()
    for i, pdf in enumerate(sorted(pdf_paths), start=1):
        bname = pdf.name
        log(f"[{i:>2}/{len(pdf_paths)}] {bname}")
        f0 = time.time()
        try:
            items = ec.process_pdf(str(pdf))
        except Exception as exc:
            log(f"    ERROR: {type(exc).__name__}: {exc}")
            log("    " + traceback.format_exc().replace("\n", "\n    "))
            per_file.append({
                "pdf": bname,
                "error": f"{type(exc).__name__}: {exc}",
                "items": [],
                "n_items": 0,
                "elapsed_s": round(time.time() - f0, 2),
            })
            continue
        items_dicts = [
            {
                "symbol": it.symbol,
                "name": it.name,
                "count": int(it.count),
                "count_ae": int(it.count_ae),
            }
            for it in items
        ]
        per_file.append({
            "pdf": bname,
            "items": items_dicts,
            "n_items": len(items_dicts),
            "elapsed_s": round(time.time() - f0, 2),
        })
        log(f"    items={len(items_dicts):>3}  "
            f"sum_count={sum(d['count'] for d in items_dicts):>4}  "
            f"sum_ae={sum(d['count_ae'] for d in items_dicts):>4}  "
            f"t={time.time()-f0:.1f}s")

    return {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "pdf_dir": str(PDF_DIR),
        "n_pdfs": len(pdf_paths),
        "elapsed_total_s": round(time.time() - t0, 2),
        "pipeline_version": _git_info(),
        "per_file": per_file,
    }


# ────────────────────────────────────────────────────────────────────────────
# Collapse + XLSX
# ────────────────────────────────────────────────────────────────────────────

def collapse_by_normalized_name(report: dict) -> list[dict]:
    """Свернуть позиции, отличающиеся только высотой.

    Ключ свертки: (name_normalized, unit). unit для оборудования EquipmentItem
    всегда «шт» (штук). Кабельная продукция в этой задаче (count-only) не
    собирается отдельно — process_pdf возвращает только оборудование.
    """
    buckets: "OrderedDict[tuple[str, str], dict]" = OrderedDict()
    for pf in report["per_file"]:
        src_pdf = pf["pdf"]
        for it in pf.get("items", []):
            raw_name = it["name"] or ""
            symbol = it["symbol"] or ""
            n = normalize_name(raw_name)
            if not n:
                n = symbol or "(empty)"
            # Также удаляем условный маркер высоты, который может попасть в name
            unit = "шт"
            key = (n.lower(), unit)
            bucket = buckets.setdefault(key, {
                "item_name_normalized": n,
                "unit": unit,
                "original_names": [],
                "original_names_set": set(),
                "total_qty": 0,
                "total_ae": 0,
                "sources": [],
            })
            qty = int(it["count"]) + int(it["count_ae"])
            bucket["total_qty"] += qty
            bucket["total_ae"] += int(it["count_ae"])
            if raw_name and raw_name not in bucket["original_names_set"]:
                bucket["original_names_set"].add(raw_name)
                bucket["original_names"].append(raw_name)
            bucket["sources"].append({
                "pdf": src_pdf, "symbol": symbol,
                "count": int(it["count"]), "count_ae": int(it["count_ae"]),
            })
    out = []
    for b in buckets.values():
        b.pop("original_names_set", None)
        out.append(b)
    # Сортируем по total_qty desc для удобства анализа
    out.sort(key=lambda r: r["total_qty"], reverse=True)
    return out


def write_xlsx(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "count_only"
    ws.append(["item_name_normalized", "original_names", "unit", "total_qty"])
    for r in rows:
        ws.append([
            r["item_name_normalized"],
            " | ".join(r["original_names"]),
            r["unit"],
            r["total_qty"],
        ])
    # Ширины
    widths = [60, 80, 10, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(ord("A") + i - 1)].width = w
    wb.save(str(path))


# ────────────────────────────────────────────────────────────────────────────
# Main
# ────────────────────────────────────────────────────────────────────────────

def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    log_lines: list[str] = []

    def log(msg: str = "") -> None:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        print(line)
        log_lines.append(line)

    git = _git_info()
    log("T122 run start")
    log(f"  branch={git['branch']}  commit={git['short']}  full={git['commit']}")
    log(f"  PDF_DIR={PDF_DIR}")
    log(f"  OUT_DIR={OUT_DIR}")

    if not PDF_DIR.is_dir():
        log(f"FATAL: PDF_DIR does not exist")
        return 2
    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    log(f"  found {len(pdfs)} PDFs")
    if not pdfs:
        log("FATAL: no PDFs to process")
        return 3

    report = run_all(pdfs, log)
    REPORT_JSON.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"WROTE {REPORT_JSON.name} "
        f"({REPORT_JSON.stat().st_size} bytes)")

    rows = collapse_by_normalized_name(report)
    log(f"Collapsed to {len(rows)} unique items (was "
        f"{sum(pf['n_items'] for pf in report['per_file'])} raw items)")
    log(f"Top-10 by total_qty:")
    for r in rows[:10]:
        log(f"  {r['total_qty']:>5}  {r['item_name_normalized'][:80]}")

    write_xlsx(rows, OUT_XLSX)
    log(f"WROTE {OUT_XLSX.name}")

    # Pretty stats
    n_with_zero = sum(1 for r in rows if r["total_qty"] == 0)
    log(f"  items_total={len(rows)}  items_with_zero_qty={n_with_zero}  "
        f"sum_total_qty={sum(r['total_qty'] for r in rows)}")
    err_files = [pf["pdf"] for pf in report["per_file"] if "error" in pf]
    if err_files:
        log(f"  PDFs with errors: {len(err_files)}")
        for f in err_files:
            log(f"    ! {f}")
    else:
        log("  All PDFs processed without exceptions")

    log("T122 run done")

    RUN_LOG.write_text("\n".join(log_lines) + "\n", encoding="utf-8")
    print(f"\nLog saved to: {RUN_LOG}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
