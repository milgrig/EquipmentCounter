"""
Read-only probe: посмотреть распределение светильников по этажам (file -> count)
в equipment_report.json для Захватки 3_ГПК, и соотнести с потолком (height zone).
"""
import json
import sys
import io
from pathlib import Path
from collections import defaultdict, Counter

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from height_mapping import (
    extract_elevations_from_filename,
    build_floor_to_ceiling_map,
    height_to_category,
    collect_floors_from_filenames,
)

REPORT = Path(r"Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/equipment_report.json")

data = json.loads(REPORT.read_text(encoding="utf-8"))

# print top keys
print("Top keys:", list(data.keys())[:20])

# Try a few shapes
per_file = data["files"]
print(f"type(files) = {type(per_file).__name__}, entries: {len(per_file)}")

# sample
if isinstance(per_file, list):
    sample = per_file[0]
    print("sample keys:", list(sample.keys()) if isinstance(sample, dict) else type(sample).__name__)
    print(json.dumps(sample, ensure_ascii=False, indent=2)[:2000])
elif isinstance(per_file, dict):
    k = next(iter(per_file))
    print("sample key:", k)
    print(json.dumps(per_file[k], ensure_ascii=False, indent=2)[:2000])
