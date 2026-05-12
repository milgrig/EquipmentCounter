#!/usr/bin/env python3
"""
Test script for generating VOR from GPK3 3rd capture PDFs and comparing with reference.

Usage:
    python test_gpk3_vor.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8")

# Import the web app processing function
import web_app
from vor_elevation_grouper import regroup_aggregate_by_elevation
from vor_height_injector import inject_vertical_cable_rows
from vor_docx_renderer import render_vor_docx

# Paths
PROJECT_ROOT = Path(__file__).resolve().parent
PDF_DIR = PROJECT_ROOT / "Data" / "ДБТ разделы для ИИ" / "03_ГПК_" / "3-я захватка" / "02_PDF"
OUTPUT_DIR = PROJECT_ROOT / "Data" / "ДБТ разделы для ИИ" / "03_ГПК_" / "3-я захватка"
OUTPUT_DOCX = OUTPUT_DIR / "VOR_GENERATED_TEST.docx"
REF_DOCX = OUTPUT_DIR / "ВОР ЭО, Захватка 3_ГПК.docx"

print("=" * 80)
print("ГЕНЕРАЦИЯ ВОР ИЗ PDF (ГПК3, 3-я захватка)")
print("=" * 80)
print()

if not PDF_DIR.exists():
    sys.exit(f"❌ Папка с PDF не найдена: {PDF_DIR}")

print(f"📁 Папка с PDF: {PDF_DIR}")
print(f"📄 Эталонный ВОР: {REF_DOCX}")
print(f"📝 Выходной файл: {OUTPUT_DOCX}")
print()

# Scan PDF files
pdf_files = sorted(PDF_DIR.glob("*.pdf"))
pdf_files = [f for f in pdf_files if not f.name.startswith("ВОР") and not f.name.startswith("VOR")]
print(f"🔍 Найдено PDF файлов: {len(pdf_files)}")
print()

# Process each PDF
t0 = time.time()
all_results = {}

for i, pdf_path in enumerate(pdf_files, 1):
    filename = pdf_path.name
    print(f"[{i}/{len(pdf_files)}] Обработка: {filename}")

    try:
        # Use web_app's equipment counting function
        items = web_app._count_equipment_in_pdf(str(pdf_path))
        all_results[filename] = items
        print(f"  ✅ Найдено элементов: {len(items)}")

    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        all_results[filename] = []

elapsed = time.time() - t0
print()
print(f"⏱️  Обработка завершена за {elapsed:.1f}с")
print()

# Aggregate equipment using web_app's aggregation function
print("📊 Агрегация оборудования...")

agg_list = web_app._aggregate_equipment(all_results)
print(f"  Уникальных позиций: {len(agg_list)}")
print()

# Regroup by elevation
print("🏗️  Перегруппировка по строительным отметкам...")
try:
    agg_list = regroup_aggregate_by_elevation(agg_list)
    print("  ✅ Перегруппировка выполнена")
except Exception as e:
    print(f"  ⚠️  Предупреждение: ошибка перегруппировки - {e}")
    import traceback
    traceback.print_exc()

# Inject vertical cable rows
print("🔌 Добавление вертикальных стояков...")
pdf_paths_str = [str(p) for p in pdf_files]
try:
    agg_list = inject_vertical_cable_rows(agg_list, pdf_paths_str, mode="split")
    print("  ✅ Стояки добавлены")
except Exception as e:
    print(f"  ⚠️  Предупреждение: ошибка добавления стояков - {e}")
    import traceback
    traceback.print_exc()

print()
print(f"📋 Итого позиций в ВОР: {len(agg_list)}")
print()

# Generate DOCX
print("📝 Генерация DOCX...")
try:
    docx_bytes = render_vor_docx(
        aggregated=agg_list,
        rel_folder="ДБТ разделы для ИИ/03_ГПК_/3-я захватка/02_PDF",
        project_name="«Комплекс по глубокой переработке зерна для производства аминокислот, расположенный по адресу: Ростовская область, г Волгодонск, улица 2-я Заводская, 3»",
        object_name="«Главный производственный корпус, поз. 3 по ГП»",
        section_basis="Основание_Электроосвещение, 3 захватка; 1Д-24-3-3-ЭО изм.2",
        composer_name="Кочерган",
        checker_name="Гончаров",
    )

    # Save to file
    OUTPUT_DOCX.write_bytes(docx_bytes)
    print(f"  ✅ Файл сохранен: {OUTPUT_DOCX}")
    print(f"  📏 Размер: {len(docx_bytes) / 1024:.1f} KB")
except Exception as e:
    print(f"  ❌ Ошибка генерации DOCX: {e}")
    import traceback
    traceback.print_exc()

print()
print("=" * 80)
print("✅ ГЕНЕРАЦИЯ ЗАВЕРШЕНА")
print("=" * 80)
print()
print(f"Результат: {OUTPUT_DOCX}")
print(f"Эталон:    {REF_DOCX}")
print()
print("Следующий шаг: сравнение с эталоном")
print(f"  python vor_compare_excel.py \"{REF_DOCX}\" \"{OUTPUT_DOCX}\"")
