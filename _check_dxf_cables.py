"""Check what extract_cables_dxf gives us as ground truth."""
import sys, io
from pathlib import Path
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, r"C:\Cursor\TayfaProject\EquipmentCounter")
from equipment_counter import extract_cables_dxf

DXFS = [
    r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\30. КПП\02_DWG\_converted_dxf\007 - Планы освещения на отм- 0-000, +2-900.dxf",
    r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\_converted_dxf\01_DWG\005 - План  освещения на отм- 0-000.dxf",
    r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\_converted_dxf\01_DWG\019 - План кабеленесущих систем на отм- 0-000.dxf",
]

for dxf in DXFS:
    name = Path(dxf).name
    print(f"\n========== {name[:60]} ==========")
    try:
        cables = extract_cables_dxf(dxf)
    except Exception as e:
        print(f"  ERROR: {e}")
        continue
    if not cables:
        print(f"  No cables extracted")
        continue
    print(f"  Total cable items: {len(cables)}")
    for c in cables[:30]:
        # CableItem dataclass — print whatever attributes it has
        attrs = {}
        for a in ("designation","brand","cross_section","total_length","laying_methods","laying"):
            if hasattr(c, a):
                v = getattr(c, a)
                if v: attrs[a] = v
        print(f"  - {attrs}")
