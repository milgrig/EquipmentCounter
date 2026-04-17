#!/usr/bin/env python3
"""
Batch Equipment Counter — batch-processes engineering drawings.

  1. Scans a folder for DWG / DXF / PDF files
  2. Converts all DWG → DXF via ODA File Converter
  3. Parses every DXF/PDF through equipment_counter
  4. Saves a consolidated JSON report

Handles Cyrillic paths and special characters in file/folder names.

Usage (CLI):
    python batch_equipment.py "C:\\Projects\\007 - План освещения"
    python batch_equipment.py /mnt/c/Projects/drawings --output report.json

Usage (GUI):
    python equipment_gui.py
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections import Counter, defaultdict
from dataclasses import asdict
from pathlib import Path

from equipment_counter import (
    EquipmentItem,
    CableItem,
    process_dxf,
    process_pdf,
    extract_cables_dxf,
    print_table,
    classify_plan,
    extract_elevation_str,
    ELEVATION_RE,
    _HAS_DXF,
    _HAS_PDF,
)

_PLAN_DUPES = {
    "освещение": "light",
    "привязка": "light",
    "розетки": "power",
    "расположение": "power",
}
_PLAN_PRIORITY = {
    "привязка": 10,
    "розетки": 10,
    "освещение": 5,
    "расположение": 5,
    "кабеленесущие": 1,
    "схема": 0,
    "опросные": 0,
    "другое": 1,
}


def _classify_plan(filename: str) -> str:
    return classify_plan(filename)


def _extract_elevation(filename: str) -> str | None:
    return extract_elevation_str(filename)

ODA_EXE_CANDIDATES_WSL = [
    r"/mnt/c/Program Files/ODA/ODAFileConverter 27.1.0/ODAFileConverter.exe",
    r"/mnt/c/Program Files/ODA/ODAFileConverter/ODAFileConverter.exe",
    r"/mnt/c/Program Files (x86)/ODA/ODAFileConverter/ODAFileConverter.exe",
]

ODA_EXE_CANDIDATES_WIN = [
    r"C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe",
    r"C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe",
    r"C:\Program Files (x86)\ODA\ODAFileConverter\ODAFileConverter.exe",
]

DWG_EXTS = {".dwg", ".DWG"}
DXF_EXTS = {".dxf", ".DXF"}
PDF_EXTS = {".pdf", ".PDF"}


# ---------------------------------------------------------------------------
#  Utilities
# ---------------------------------------------------------------------------

def _is_wsl() -> bool:
    try:
        with open("/proc/version", "r") as f:
            return "microsoft" in f.read().lower()
    except OSError:
        return False


def _wsl_to_win(path: str) -> str:
    try:
        return subprocess.check_output(
            ["wslpath", "-w", path], text=True, timeout=5
        ).strip()
    except (FileNotFoundError, subprocess.SubprocessError):
        return path


def find_oda() -> str | None:
    candidates = ODA_EXE_CANDIDATES_WSL if _is_wsl() else ODA_EXE_CANDIDATES_WIN
    for p in candidates:
        if os.path.isfile(p):
            return p
    return shutil.which("ODAFileConverter")


def scan_files(folder: Path) -> tuple[list[Path], list[Path], list[Path]]:
    """Scan folder recursively for DWG, DXF, PDF files."""
    dwg, dxf, pdf = [], [], []
    for f in sorted(folder.rglob("*")):
        if not f.is_file():
            continue
        rel_parts = f.relative_to(folder).parts
        if any(p.startswith(".") for p in rel_parts):
            continue
        if any(p.startswith("_") and p != "_converted_dxf" for p in rel_parts):
            continue
        ext = f.suffix
        if ext in DWG_EXTS:
            dwg.append(f)
        elif ext in DXF_EXTS:
            dxf.append(f)
        elif ext in PDF_EXTS:
            pdf.append(f)
    return dwg, dxf, pdf


def convert_dwg_files(
    oda_exe: str,
    dwg_files: list[Path],
    output_root: Path,
    log=print,
) -> dict[Path, Path]:
    """Convert DWG files → DXF via ODA. Calls ODA once per source directory.

    Returns {dwg_path: converted_dxf_path}.
    """
    # Group DWG files by parent directory (ODA works per-folder)
    by_dir: dict[Path, list[Path]] = {}
    for f in dwg_files:
        by_dir.setdefault(f.parent, []).append(f)

    result_map: dict[Path, Path] = {}
    t0 = time.time()

    for src_dir, files in by_dir.items():
        rel = src_dir.relative_to(output_root.parent) if src_dir != output_root.parent else Path(".")
        out_dir = output_root / rel
        out_dir.mkdir(parents=True, exist_ok=True)

        if _is_wsl():
            win_in = _wsl_to_win(str(src_dir))
            win_out = _wsl_to_win(str(out_dir))
        else:
            win_in = str(src_dir)
            win_out = str(out_dir)

        cmd = [oda_exe, win_in, win_out, "ACAD2018", "DXF", "0", "1", "*.DWG"]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

        if r.returncode != 0:
            log(f"  ⚠ ODA ошибка в {src_dir.name}: код {r.returncode}")

        for dwg in files:
            expected_dxf = out_dir / (dwg.stem + ".dxf")
            if expected_dxf.exists():
                result_map[dwg] = expected_dxf
                sz_mb = expected_dxf.stat().st_size / 1024 / 1024
                log(f"    ✓ {dwg.name} → DXF ({sz_mb:.1f} MB)")
            else:
                log(f"    ⚠ {dwg.name} — конвертация не удалась")

    elapsed = time.time() - t0
    log(f"  Итого: {len(result_map)}/{len(dwg_files)} файлов за {elapsed:.1f}с")
    return result_map


def _try_pdf_overlay(
    dxf_path: Path,
    items: list[EquipmentItem],
    dpi: int,
    source_folder: Path,
    log=print,
) -> None:
    """Look for a matching PDF and generate overlay PNGs if found."""
    candidates = [
        dxf_path.with_suffix(".pdf"),
        dxf_path.parent.parent / (dxf_path.stem + ".pdf"),
    ]
    for child in source_folder.iterdir():
        if child.suffix.lower() == ".pdf" and child.stem == dxf_path.stem:
            candidates.append(child)

    pdf_found = None
    for c in candidates:
        if c.exists():
            pdf_found = c
            break

    if pdf_found is None:
        return

    try:
        from pdf_overlay import visualize_on_pdf
        log(f"    [pdf-overlay] Найден PDF: {pdf_found.name}")
        visualize_on_pdf(str(pdf_found), str(dxf_path),
                         items=items, dpi=dpi, log=log)
    except ImportError as exc:
        log(f"    ⚠ PDF overlay пропущен (зависимость): {exc}")
    except Exception as exc:
        log(f"    ⚠ PDF overlay ошибка: {exc}")


def process_file(path: Path, log=print) -> list[EquipmentItem]:
    ext = path.suffix.lower()
    try:
        if ext == ".dxf" and _HAS_DXF:
            return process_dxf(str(path))
        elif ext == ".pdf" and _HAS_PDF:
            return process_pdf(str(path))
    except Exception as e:
        if "DXFStructureError" in type(e).__name__ or "ENDSEC" in str(e):
            log(f"  ⚠ DXF ошибка чтения, повторная попытка через 2с...")
            time.sleep(2)
            try:
                return process_dxf(str(path))
            except Exception as e2:
                log(f"  ⚠ Повторная ошибка: {path.name}: {e2}")
        else:
            log(f"  ⚠ Ошибка: {path.name}: {e}")
    return []


def _postprocess(
    results: dict[str, list[EquipmentItem]],
    cable_results: dict[str, list[CableItem]],
    log=print,
) -> tuple[dict[str, list[EquipmentItem]], dict[str, list[CableItem]]]:
    """Apply cross-file improvements to parsed results.

    1. Merge legends: resolve [Auto X] names using SAME-ELEVATION files' legends
       P-003 fix: only merge legends within the same elevation to prevent
       symbol 1 on floor A overwriting symbol 1 on floor B.
    2. Deduplicate floors: keep best file per (elevation, plan_type)
    3. Deduplicate cables: skip Опросные листы, merge across files
    """
    # ── 1. Cross-file legend merging (elevation-scoped) ──
    # P-003 fix: Build per-elevation legend maps instead of a global one.
    # Only resolve [Auto X] names using legends from the SAME elevation.
    by_elevation: dict[str | None, dict[str, str]] = defaultdict(dict)
    for fname, items in results.items():
        elev = _extract_elevation(fname)
        for it in items:
            if it.name.startswith("[Auto") or it.name.startswith("[?"):
                continue
            if it.symbol not in by_elevation[elev]:
                by_elevation[elev][it.symbol] = it.name

    resolved = 0
    for fname, items in results.items():
        elev = _extract_elevation(fname)
        elev_legend = by_elevation.get(elev, {})
        for it in items:
            if it.name.startswith("[Auto") and it.symbol in elev_legend:
                it.name = elev_legend[it.symbol]
                resolved += 1
    if resolved:
        log(f"  [пост] Объединение легенд (по этажам): {resolved} символов получили имена")

    # ── 2. Floor deduplication ──
    groups: dict[tuple[str, str | None], list[tuple[str, int, str]]] = defaultdict(list)
    for fname in results:
        ptype = _classify_plan(fname)
        elev = _extract_elevation(fname)
        category = _PLAN_DUPES.get(ptype)
        if category and elev:
            total = sum(it.count + it.count_ae for it in results[fname])
            priority = _PLAN_PRIORITY.get(ptype, 0)
            groups[(category, elev)].append((fname, total, ptype))

    dropped: set[str] = set()
    for key, files in groups.items():
        if len(files) <= 1:
            continue
        files.sort(key=lambda x: (_PLAN_PRIORITY.get(x[2], 0), x[1]), reverse=True)
        best = files[0]
        for fname, total, ptype in files[1:]:
            dropped.add(fname)
            log(f"  [пост] Дубль этажа: {fname} → убран (приоритет у {best[0]})")

    out_results = {k: v for k, v in results.items() if k not in dropped}

    # ── 3. Cable deduplication ──
    cable_out: dict[str, list[CableItem]] = {}
    global_seen: set[tuple[str, int]] = set()

    for fname in sorted(cable_results.keys()):
        ptype = _classify_plan(fname)
        if ptype == "опросные":
            log(f"  [пост] Кабели: пропущен {fname} (дубль опросных)")
            continue
        if fname in dropped:
            continue
        deduped: list[CableItem] = []
        for c in cable_results[fname]:
            key = (c.cable_type, c.total_length_m)
            if key not in global_seen:
                global_seen.add(key)
                deduped.append(c)
        if deduped:
            cable_out[fname] = deduped

    orig_cable_m = sum(sum(c.total_length_m for c in v) for v in cable_results.values())
    new_cable_m = sum(sum(c.total_length_m for c in v) for v in cable_out.values())
    if orig_cable_m != new_cable_m:
        log(f"  [пост] Кабели: {orig_cable_m}м → {new_cable_m}м (дедупликация)")

    return out_results, cable_out


def build_report(
    results: dict[str, list[EquipmentItem]],
    cable_results: dict[str, list[CableItem]] | None = None,
) -> dict:
    """Build the final JSON-serializable report."""
    report = {"files": {}, "generated": time.strftime("%Y-%m-%d %H:%M:%S")}

    all_fnames = set(results.keys())
    if cable_results:
        all_fnames |= set(cable_results.keys())

    for fname in sorted(all_fnames):
        items = results.get(fname, [])
        total = sum(it.count + it.count_ae for it in items)
        file_entry: dict = {
            "total_count": total,
            "equipment": [
                {
                    "symbol": it.symbol,
                    "name": it.name,
                    "count": it.count,
                    "count_ae": it.count_ae,
                    "total": it.count + it.count_ae,
                }
                for it in items
            ],
        }
        if cable_results and fname in cable_results:
            cables = cable_results[fname]
            file_entry["cables"] = [
                {
                    "cable_type": c.cable_type,
                    "count": c.count,
                    "total_length_m": c.total_length_m,
                    **({"length_by_laying": c.length_by_laying}
                       if c.length_by_laying else {}),
                }
                for c in cables
            ]
            file_entry["total_cable_length_m"] = sum(c.total_length_m for c in cables)
        report["files"][fname] = file_entry
    return report


# ---------------------------------------------------------------------------
#  Core pipeline — used by both CLI and GUI
# ---------------------------------------------------------------------------

def run_pipeline(
    folder: Path,
    *,
    convert_dwg: bool = True,
    keep_converted: bool = False,
    parse_dxf: bool = True,
    parse_pdf: bool = True,
    generate_png: bool = False,
    png_dpi: int = 200,
    output_json: str = "equipment_report.json",
    log=print,
) -> dict:
    """Run the full processing pipeline. Returns the report dict."""
    log(f"{'=' * 50}")
    log(f"  Обработка: {folder.name}")
    log(f"{'=' * 50}")

    # ── Scan ──
    log(f"\n[0] Сканирование папки")
    dwg_files, dxf_files, pdf_files = scan_files(folder)
    total = len(dwg_files) + len(dxf_files) + len(pdf_files)
    # Count unique subdirectories
    all_dirs = set()
    for f in dwg_files + dxf_files + pdf_files:
        all_dirs.add(f.parent)
    dirs_info = f" в {len(all_dirs)} папках" if len(all_dirs) > 1 else ""
    log(f"  DWG: {len(dwg_files)}  |  DXF: {len(dxf_files)}  |  PDF: {len(pdf_files)}  ({total} файлов{dirs_info})")

    if total == 0:
        log("  Файлы чертежей не найдены.")
        return {}

    # ── Convert DWG → DXF ──
    converted_dir: Path | None = None
    dwg_to_dxf: dict[Path, Path] = {}

    if dwg_files and convert_dwg:
        log(f"\n[1] Конвертация DWG → DXF ({len(dwg_files)} файлов)")
        oda = find_oda()
        if not oda:
            log("  ⚠ ODA File Converter не найден!")
            log("    Скачайте: https://www.opendesign.com/guestfiles/oda_file_converter")
        else:
            log(f"  ODA: {Path(oda).name}")
            converted_dir = folder / "_converted_dxf"
            dwg_to_dxf = convert_dwg_files(oda, dwg_files, converted_dir, log=log)
    elif dwg_files:
        log(f"\n  Пропущено: {len(dwg_files)} DWG файлов")

    def _rel(p: Path) -> str:
        try:
            return str(p.relative_to(folder))
        except ValueError:
            return p.name

    # ── Build file list ──
    all_parseable: list[tuple[str, Path]] = []
    if parse_dxf:
        for f in dxf_files:
            all_parseable.append((_rel(f), f))
    for dwg_path, dxf_path in dwg_to_dxf.items():
        all_parseable.append((_rel(dwg_path), dxf_path))
    if parse_pdf:
        for f in pdf_files:
            all_parseable.append((_rel(f), f))

    if not all_parseable:
        log("\n  Нет файлов для парсинга.")
        return {}

    # ── Parse ──
    log(f"\n[2] Парсинг чертежей ({len(all_parseable)} файлов)")
    results: dict[str, list[EquipmentItem]] = {}
    cable_results: dict[str, list[CableItem]] = {}
    for display_name, fpath in all_parseable:
        log(f"\n  → {display_name}")
        items = process_file(fpath, log=log)
        if items:
            results[display_name] = items
            for it in items:
                total = it.count + it.count_ae
                ae_s = f" +{it.count_ae}АЭ" if it.count_ae else ""
                log(f"    {it.symbol:>4}  {it.count}{ae_s} = {total}  {it.name[:60]}")
        else:
            log("    (оборудование не найдено)")

        if fpath.suffix.lower() == ".dxf" and _HAS_DXF:
            try:
                cables = extract_cables_dxf(str(fpath))
                if cables:
                    cable_results[display_name] = cables
                    total_m = sum(c.total_length_m for c in cables)
                    log(f"    📏 Кабели: {len(cables)} типов, {total_m}м")
            except Exception as e:
                log(f"    ⚠ Кабели: ошибка — {e}")

            if generate_png:
                # V-001: Prefer PDF→PNG (90% quality) over DXF→PNG (40%)
                pdf_overlay_ok = False
                _try_pdf_overlay(fpath, items, png_dpi, folder, log)
                # Check if PDF overlay produced files
                _base = fpath.with_suffix("")
                pdf_pngs = list(_base.parent.glob(f"{_base.name}__pdf_*.png"))
                if pdf_pngs:
                    pdf_overlay_ok = True
                    log(f"    [viz] PDF→PNG предпочтён ({len(pdf_pngs)} файлов)")

                # V-005: Always generate DXF→PNG as fallback/diagnostic
                # (removed gate: no longer requires items or fsize_mb > 5)
                if not pdf_overlay_ok:
                    try:
                        from dxf_visualizer import visualize_dxf
                        png_out = str(fpath.with_suffix(".png"))
                        visualize_dxf(str(fpath), png_out, items=items,
                                      dpi=png_dpi, log=log)
                    except ImportError as exc:
                        log(f"    ⚠ PNG пропущен (нет matplotlib): {exc}")
                    except Exception as exc:
                        log(f"    ⚠ PNG ошибка: {exc}")

    # ── Post-process ──
    if results:
        log(f"\n[3] Пост-обработка")
        results, cable_results = _postprocess(results, cable_results, log=log)

    # ── Save JSON ──
    report = {}
    if results:
        if os.sep not in output_json and "/" not in output_json:
            output_path = str(folder / output_json)
        else:
            output_path = str(Path(output_json).resolve())

        log(f"\n[4] Сохранение отчёта")
        report = build_report(results, cable_results)
        report["source_folder"] = str(folder)
        report["files_processed"] = len(results)
        report["files_total"] = len(all_parseable)

        with open(output_path, "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)
        log(f"  ✓ {output_path}")
    else:
        log("\n  Результатов нет — JSON не создан.")

    # ── Cleanup ──
    if converted_dir and converted_dir.exists() and not keep_converted:
        shutil.rmtree(converted_dir, ignore_errors=True)
        log(f"\n  Временные файлы удалены.")

    # ── Summary ──
    total_equip = sum(
        sum(it.count + it.count_ae for it in items)
        for items in results.values()
    )
    total_cable_m = sum(
        sum(c.total_length_m for c in cables)
        for cables in cable_results.values()
    )
    log(f"\n{'=' * 50}")
    log(f"  Файлов обработано: {len(results)} / {len(all_parseable)}")
    log(f"  Единиц оборудования: {total_equip}")
    if total_cable_m > 0:
        log(f"  Кабелей: {total_cable_m}м ({len(cable_results)} файлов)")
    log(f"{'=' * 50}")

    return report


# ---------------------------------------------------------------------------
#  CLI entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Пакетный подсчёт оборудования из инженерных чертежей (DWG/DXF/PDF)"
    )
    parser.add_argument(
        "folder",
        help="Папка с чертежами (поддерживает кириллицу и спецсимволы в пути)",
    )
    parser.add_argument("--output", "-o", default="equipment_report.json")
    parser.add_argument("--no-convert", action="store_true")
    parser.add_argument("--keep-converted", action="store_true")
    parser.add_argument("--no-pdf", action="store_true")
    parser.add_argument("--no-dxf", action="store_true")
    parser.add_argument("--png", action="store_true",
                        help="Generate PNG images with equipment markers")
    parser.add_argument("--png-dpi", type=int, default=200,
                        help="PNG resolution (default 200)")
    parser.add_argument("--vor", nargs="?", const="auto", default=None,
                        help="Generate VOR .docx (optionally specify output path)")
    args = parser.parse_args()

    folder = Path(args.folder).resolve()
    if not folder.is_dir():
        sys.exit(f"Папка не найдена: {folder}")

    run_pipeline(
        folder,
        convert_dwg=not args.no_convert,
        keep_converted=args.keep_converted,
        parse_dxf=not args.no_dxf,
        parse_pdf=not args.no_pdf,
        generate_png=args.png,
        png_dpi=args.png_dpi,
        output_json=args.output,
    )

    if args.vor is not None:
        try:
            from vor_generator import generate_vor
            vor_out = None if args.vor == "auto" else args.vor
            generate_vor(folder, output_path=vor_out)
        except ImportError as e:
            print(f"  VOR skipped (missing dependency): {e}")
        except Exception as e:
            print(f"  VOR error: {e}")


if __name__ == "__main__":
    main()
