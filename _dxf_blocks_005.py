import sys, io, ezdxf
from collections import Counter
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
DXF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\_converted_dxf\01_DWG\005 - План  освещения на отм- 0-000.dxf")
doc = ezdxf.readfile(str(DXF)); msp = doc.modelspace()
c = Counter()
for e in msp:
    if e.dxftype() == "INSERT":
        c[(e.dxf.layer, e.dxf.name)] += 1
for (lay, name), n in sorted(c.items(), key=lambda x: -x[1]):
    print(f"  {n:>4d}  layer={lay}  block={name}")
