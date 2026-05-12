"""Build template library for KPP plan-007 with all visible legend symbols."""
from __future__ import annotations
import io, sys
from pathlib import Path
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF = ROOT / r"Data\ДБТ разделы для ИИ\30. КПП\03_PDF\007-План освещения.pdf"
LIB = ROOT / "_template_lib_kpp"
SYMBOLS = ["1", "1А", "2", "3А", "ВЫХОД"]

sys.path.insert(0, str(ROOT))
from template_library import build_library
build_library([PDF], SYMBOLS, lib_dir=LIB, verbose=True)
