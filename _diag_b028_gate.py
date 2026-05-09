"""B028 gate verification: 005 PDF must emit at least one item mapped
to legend wiring categories (indexes 13/14/15 / names matching
'Проводка прокладываемая ...')."""
from __future__ import annotations
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

PDF = Path(r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")

import subprocess
# Invoke via subprocess to avoid test_vor_cross_format import issue in diag.
code = f'''
import sys
sys.path.insert(0, ".")
from test_vor_cross_format import _count_equipment_in_pdf
items = _count_equipment_in_pdf(r"{PDF}")
wiring = [i for i in items
          if "Проводка" in i.get("name", "")
          or i.get("source") == "wiring_derivation"]
print(f"TOTAL ITEMS: {{len(items)}}")
print(f"WIRING ITEMS: {{len(wiring)}}")
for w in wiring:
    print(f"  name={{w.get('name')!r}} total={{w.get('total')}} unit={{w.get('unit')}} legend_index={{w.get('legend_index')}} source={{w.get('source')}}")
print("GATE:", "PASS" if len(wiring) >= 1 else "FAIL")
'''
r = subprocess.run([sys.executable, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8")
print(r.stdout)
if r.returncode != 0:
    print("STDERR:", r.stderr[-2000:])
