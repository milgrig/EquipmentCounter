"""Cross-NCC matrix between all per-PDF library templates to confirm they really are identical."""
from __future__ import annotations
import io, sys, json
from pathlib import Path
import numpy as np

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

LIB = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_template_lib")
index = json.loads((LIB/"index.json").read_text(encoding="utf-8"))

records = []
for stem, syms in sorted(index.items()):
    for sym, rec in sorted(syms.items()):
        tpl = np.load(str(LIB / rec["tpl_path"])).astype(np.float32)
        records.append((f"{stem}/{sym}", tpl))

def ncc(a, b):
    a=a-a.mean(); b=b-b.mean()
    na=np.linalg.norm(a); nb=np.linalg.norm(b)
    if na<1e-6 or nb<1e-6: return 0.0
    return float(np.sum(a*b)/(na*nb))

print(f"Templates: {len(records)}")
print()
print(f"{'':<12}", end="")
for name,_ in records:
    print(f"{name:<12}", end="")
print()
for n1, t1 in records:
    print(f"{n1:<12}", end="")
    for n2, t2 in records:
        print(f"{ncc(t1,t2):>7.3f}     ", end="")
    print()
