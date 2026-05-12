"""
Read-only probe: разбивка equipment_report.json по символам для каждого этажа.
Цель: понять, какие symbol'ы — это светильники, а какие — группы/обозначения.
"""
import json
import sys
import io
from pathlib import Path
from collections import defaultdict

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

REPORT = Path(r"Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/equipment_report.json")
data = json.loads(REPORT.read_text(encoding="utf-8"))

# Покажем топ-символы для каждого плана этажа
for fname, content in data["files"].items():
    if "привязки светильников" not in fname and "План освещения" not in fname:
        continue
    eq = content.get("equipment", [])
    total = content.get("total_count", 0)
    print(f"\n=== {fname} (total={total}) ===")
    # отсортируем по count убыв.
    eq_sorted = sorted(eq, key=lambda x: -x.get("count", 0))
    for item in eq_sorted[:15]:
        sym = item.get("symbol", "?")
        name = item.get("name", "")
        cnt = item.get("count", 0)
        print(f"  [{sym:>6}] x{cnt:<4}  {name}")

# Агрегат по символам по всем этажам освещения
print("\n\n=== Агрегат по ВСЕМ планам освещения ===")
global_by_sym = defaultdict(int)
for fname, content in data["files"].items():
    if "привязки светильников" not in fname and "План освещения" not in fname:
        continue
    for item in content.get("equipment", []):
        global_by_sym[item.get("symbol", "?")] += item.get("count", 0)
for sym, cnt in sorted(global_by_sym.items(), key=lambda x: -x[1]):
    print(f"  [{sym:>6}] x{cnt}")
