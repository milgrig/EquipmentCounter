#!/usr/bin/env python3
"""
Batch Equipment Counter — пакетный подсчёт оборудования.

1. Пользователь указывает папку с инженерными чертежами
2. DWG файлы конвертируются в DXF через ODA File Converter
3. Все DXF + PDF парсятся (извлекается легенда и подсчёт символов)
4. Результат → единый JSON (по каждому файлу: наименование + количество)

Поддерживает кириллические пути и спецсимволы в именах файлов/папок.

Usage:
    python batch_processor.py                       # интерактивный режим
    python batch_processor.py "C:\\Проект\\Чертежи"  # указать папку
    python batch_processor.py ./drawings -o report.json
"""

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import asdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from equipment_counter import (
    EquipmentItem,
    process_dxf,
    process_pdf,
    _HAS_DXF,
    _HAS_PDF,
)

ODA_SEARCH_PATHS = [
    r"C:\Program Files\ODA",
    r"C:\Program Files (x86)\ODA",
    os.path.expanduser("~/ODAFileConverter"),
    "/opt/ODAFileConverter",
    "/usr/bin",
]

SUPPORTED_DWG = {".dwg", ".DWG"}
SUPPORTED_DXF = {".dxf", ".DXF"}
SUPPORTED_PDF = {".pdf", ".PDF"}

_IS_WSL = "microsoft" in platform.uname().release.lower()


# ── helpers ──────────────────────────────────────────────────────────

def _wsl_to_win(path: str) -> str:
    return subprocess.check_output(
        ["wslpath", "-w", path], text=True
    ).strip()


def _find_oda_exe() -> str | None:
    """Locate ODAFileConverter executable."""
    if _IS_WSL:
        for base in ODA_SEARCH_PATHS:
            mnt = base.replace("C:\\", "/mnt/c/").replace("\\", "/")
            if not os.path.isdir(mnt):
                continue
            for entry in sorted(os.listdir(mnt), reverse=True):
                candidate = os.path.join(mnt, entry, "ODAFileConverter.exe")
                if os.path.isfile(candidate):
                    return candidate
            candidate = os.path.join(mnt, "ODAFileConverter.exe")
            if os.path.isfile(candidate):
                return candidate
    else:
        exe_name = "ODAFileConverter.exe" if os.name == "nt" else "ODAFileConverter"
        found = shutil.which(exe_name)
        if found:
            return found
        for base in ODA_SEARCH_PATHS:
            if not os.path.isdir(base):
                continue
            for entry in sorted(os.listdir(base), reverse=True):
                candidate = os.path.join(base, entry, exe_name)
                if os.path.isfile(candidate):
                    return candidate
    return None


def _printc(msg: str, color: str = "") -> None:
    """Print with optional ANSI color."""
    codes = {"green": "\033[92m", "yellow": "\033[93m", "red": "\033[91m",
             "cyan": "\033[96m", "bold": "\033[1m", "dim": "\033[2m"}
    reset = "\033[0m"
    prefix = codes.get(color, "")
    print(f"{prefix}{msg}{reset}" if prefix else msg)


def _progress(current: int, total: int, label: str) -> None:
    bar_len = 30
    filled = int(bar_len * current / total) if total else 0
    bar = "█" * filled + "░" * (bar_len - filled)
    pct = int(100 * current / total) if total else 0
    print(f"\r  [{bar}] {pct:>3}%  {label:<60}", end="", flush=True)


# ── DWG → DXF conversion ────────────────────────────────────────────

def convert_dwg_folder(
    src_folder: Path,
    dwg_files: list[Path],
    oda_exe: str,
) -> list[Path]:
    """Convert DWG files to DXF using ODA File Converter.
    Returns list of converted DXF paths."""
    if not dwg_files:
        return []

    out_dir = src_folder / "_converted_dxf"
    out_dir.mkdir(exist_ok=True)

    if _IS_WSL:
        win_src = _wsl_to_win(str(src_folder))
        win_dst = _wsl_to_win(str(out_dir))
    else:
        win_src = str(src_folder)
        win_dst = str(out_dir)

    _printc(f"\n  Конвертация {len(dwg_files)} DWG → DXF ...", "cyan")
    t0 = time.time()
    try:
        subprocess.run(
            [oda_exe, win_src, win_dst, "ACAD2018", "DXF", "0", "1", "*.DWG"],
            capture_output=True, timeout=300,
        )
    except subprocess.TimeoutExpired:
        _printc("  ОШИБКА: Таймаут ODA File Converter (5 мин)", "red")
        return []
    except FileNotFoundError:
        _printc(f"  ОШИБКА: ODA не найден: {oda_exe}", "red")
        return []

    converted = sorted(out_dir.glob("*.dxf"))
    dt = time.time() - t0
    _printc(f"  Сконвертировано: {len(converted)} файлов за {dt:.1f}с", "green")
    return converted


# ── batch parse ──────────────────────────────────────────────────────

def parse_file(path: Path) -> list[EquipmentItem]:
    ext = path.suffix.lower()
    if ext == ".dxf" and _HAS_DXF:
        return process_dxf(str(path))
    if ext == ".pdf" and _HAS_PDF:
        return process_pdf(str(path))
    return []


def build_report(
    folder: Path,
    parsed: dict[str, list[EquipmentItem]],
) -> dict:
    """Build consolidated JSON report."""
    report = {
        "source_folder": str(folder),
        "files_processed": len(parsed),
        "files": {},
    }
    grand_total = 0
    for fname, items in parsed.items():
        file_total = sum(it.count + it.count_ae for it in items)
        grand_total += file_total
        report["files"][fname] = {
            "equipment_count": len(items),
            "total_quantity": file_total,
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
    report["grand_total"] = grand_total
    return report


# ── main ─────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Пакетный подсчёт оборудования из инженерных чертежей"
    )
    parser.add_argument(
        "folder", nargs="?", default=None,
        help="Папка с чертежами (DWG/DXF/PDF)",
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="Путь для JSON-отчёта (по умолчанию: <папка>/equipment_report.json)",
    )
    parser.add_argument(
        "--no-convert", action="store_true",
        help="Пропустить конвертацию DWG → DXF",
    )
    args = parser.parse_args()

    _printc("╔══════════════════════════════════════════╗", "bold")
    _printc("║   Подсчёт оборудования — Batch Processor ║", "bold")
    _printc("╚══════════════════════════════════════════╝", "bold")

    # 1. Get folder
    folder_str = args.folder
    if not folder_str:
        _printc("\n  Введите путь к папке с чертежами:", "cyan")
        folder_str = input("  > ").strip().strip('"').strip("'")

    folder = Path(folder_str).resolve()
    if not folder.is_dir():
        _printc(f"\n  ОШИБКА: Папка не найдена: {folder}", "red")
        sys.exit(1)

    _printc(f"\n  Папка: {folder}", "dim")

    # 2. Scan files
    dwg_files = sorted(
        f for f in folder.iterdir()
        if f.suffix in SUPPORTED_DWG and not f.name.startswith("_")
    )
    dxf_files = sorted(
        f for f in folder.iterdir()
        if f.suffix in SUPPORTED_DXF and not f.name.startswith("_")
    )
    pdf_files = sorted(
        f for f in folder.iterdir()
        if f.suffix in SUPPORTED_PDF and not f.name.startswith("_")
    )

    _printc(f"\n  Найдено: {len(dwg_files)} DWG, {len(dxf_files)} DXF, {len(pdf_files)} PDF", "cyan")

    if not dwg_files and not dxf_files and not pdf_files:
        _printc("  Нет файлов для обработки.", "yellow")
        sys.exit(0)

    # 3. Convert DWG → DXF
    converted_dxf: list[Path] = []
    if dwg_files and not args.no_convert:
        oda_exe = _find_oda_exe()
        if oda_exe:
            _printc(f"  ODA File Converter: {oda_exe}", "dim")
            converted_dxf = convert_dwg_folder(folder, dwg_files, oda_exe)
        else:
            _printc("  ВНИМАНИЕ: ODA File Converter не найден — DWG пропущены", "yellow")
            _printc("  Скачать: https://www.opendesign.com/guestfiles/oda_file_converter", "dim")

    # 4. Build file list for parsing (prefer user-provided DXF over converted)
    parse_queue: list[tuple[str, Path]] = []

    for f in dxf_files:
        parse_queue.append((f.name, f))

    for f in converted_dxf:
        original_name = f.stem + ".dwg"
        has_user_dxf = any(
            uf.stem.lower() == f.stem.lower() for uf in dxf_files
        )
        if has_user_dxf:
            _printc(f"  Пропуск конвертированного {f.name} (есть пользовательский DXF)", "dim")
            continue
        parse_queue.append((original_name, f))

    for f in pdf_files:
        parse_queue.append((f.name, f))

    if not parse_queue:
        _printc("  Нет файлов для парсинга.", "yellow")
        sys.exit(0)

    # 5. Parse each file
    _printc(f"\n  Парсинг {len(parse_queue)} файлов...\n", "cyan")
    parsed: dict[str, list[EquipmentItem]] = {}
    errors: list[tuple[str, str]] = []

    for idx, (display_name, fpath) in enumerate(parse_queue, 1):
        _progress(idx, len(parse_queue), display_name)
        try:
            items = parse_file(fpath)
            parsed[display_name] = items
        except Exception as exc:
            errors.append((display_name, str(exc)))
            parsed[display_name] = []

    print()

    # 6. Print results per file
    from equipment_counter import print_table
    for display_name, items in parsed.items():
        if items:
            print_table(items, display_name)

    # 7. Summary
    total_files = len(parsed)
    total_items = sum(
        sum(it.count + it.count_ae for it in items)
        for items in parsed.values()
    )
    _printc(f"\n  ═══ ИТОГО ═══", "bold")
    _printc(f"  Файлов обработано: {total_files}", "cyan")
    _printc(f"  Общее количество оборудования: {total_items}", "cyan")

    if errors:
        _printc(f"\n  Ошибки ({len(errors)}):", "red")
        for name, err in errors:
            _printc(f"    {name}: {err}", "red")

    # 8. Save JSON
    out_path = args.output or str(folder / "equipment_report.json")
    report = build_report(folder, parsed)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    _printc(f"\n  JSON-отчёт: {out_path}", "green")


if __name__ == "__main__":
    main()
