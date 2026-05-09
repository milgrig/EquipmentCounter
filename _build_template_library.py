"""Build per-PDF template library for the 7 lighting plans (005..011)."""
from __future__ import annotations
import io, sys
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF_DIR = ROOT / r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF"
PDFS = [
    "005-Планы освещения-отм. 0.000.pdf",
    "006-Планы освещения-отм. +4.200.pdf",
    "007-Планы освещения-отм. +7.800 +9.000.pdf",
    "008-Планы освещения-отм. +13.800.pdf",
    "009-Планы освещения-отм. +18.600.pdf",
    "010-Планы освещения-отм. +23.400.pdf",
    "011-Планы освещения-отм. +28.200.pdf",
]
SYMBOLS = ["5АЭ", "6АЭ", "7АЭ"]
LIB_DIR = ROOT / "_template_lib"

sys.path.insert(0, str(ROOT))
from template_library import build_library

paths = [PDF_DIR / p for p in PDFS if (PDF_DIR / p).exists()]
print(f"Building library for {len(paths)} PDFs into {LIB_DIR}")
build_library(paths, SYMBOLS, lib_dir=LIB_DIR, verbose=True)
