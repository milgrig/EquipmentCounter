"""
Read-only probe: по каждому DWG-файлу (=плану этажа) показать распределение
светильников и их соответствие категории высоты (по потолку этажа).

Цель: понять, какие именно этажи дают «перекос» в категорию 13-20 м.
"""
import json
import sys
import io
from pathlib import Path
from collections import defaultdict

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from height_mapping import (
    extract_elevations_from_filename,
    build_floor_to_ceiling_map,
    height_to_category,
)

REPORT = Path(r"Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/equipment_report.json")

data = json.loads(REPORT.read_text(encoding="utf-8"))
files = data["files"]

# 1. Собираем этажи из имён файлов
all_floors = []
file_to_floor = {}
for fname in files.keys():
    elevs = extract_elevations_from_filename(fname)
    if elevs:
        floor = min(elevs)  # минимальная отметка в имени = пол этажа
        file_to_floor[fname] = floor
        all_floors.append(floor)

print("Найдено этажей (уникальных):")
for f in sorted(set(all_floors)):
    print(f"  {f:+.3f}")

# 2. Потолок для каждого этажа
f2c = build_floor_to_ceiling_map(sorted(set(all_floors)), last_floor_height=4.8)
print("\nFloor → Ceiling → Category:")
for f in sorted(f2c.keys()):
    c = f2c[f]
    cat = height_to_category(c)
    print(f"  пол={f:+.3f}  потолок={c:+.3f}  -> {cat}")

# 3. Суммы по файлу (total_count - all symbols)
print(f"\n{'Файл':<60} {'этаж':>7} {'потолок':>8} {'категория':<20} {'total':>6}")
print("-" * 110)
total_per_cat = defaultdict(int)
for fname, content in files.items():
    if fname not in file_to_floor:
        print(f"  SKIP {fname}  (no floor in name)")
        continue
    floor = file_to_floor[fname]
    ceiling = f2c[floor]
    cat = height_to_category(ceiling)
    total = content.get("total_count", 0)
    total_per_cat[cat] += total
    short = fname[-55:] if len(fname) > 55 else fname
    print(f"{short:<60} {floor:>+7.3f} {ceiling:>+8.3f} {cat:<20} {total:>6}")

print("\nСумма total_count по категориям высоты:")
for cat, n in total_per_cat.items():
    print(f"  {cat:<20}: {n}")
print(f"\n  ВСЕГО: {sum(total_per_cat.values())}")

print("\nЭталон:  117 / 182 / 81 / 206  (шпильки)  — итого 586")
