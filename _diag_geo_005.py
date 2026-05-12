"""Inspect classification: print all 19 candidate clusters with features and chosen label."""
import io, sys, json
from pathlib import Path
import numpy as np
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
exec(open(r"C:\Cursor\TayfaProject\EquipmentCounter\_geo_features_005.py", encoding="utf-8").read().replace("if __name__", "if False and __name__"))
# vars already in scope: cluster_n, F_all, F_z, P, proto_labels, assignments, cluster_bboxes
print("\n=== All candidate clusters (n>=20), sorted by n ===")
print("idx   n    W      H    aspect  horiz  sx    sy    -> label  (dist)")
order = np.argsort(-cluster_n)
for i in order:
    if cluster_n[i] < 20: continue
    f = F_all[i]
    a = assignments[i]
    bb = cluster_bboxes[i]
    cx = (bb[0]+bb[2])/2; cy = (bb[1]+bb[3])/2
    if a is None:
        ls = "(none)"
    else:
        ls = f"{a[0]} d={a[1]:.2f}"
    print(f"{i:>4d} {int(cluster_n[i]):>3d}  {f[0]:>5.1f}  {f[1]:>5.1f}  {f[2]:>5.2f}   {f[8]:>4.2f}  {f[6]:>4.1f}  {f[7]:>4.1f}  -> {ls:<20s}  cy={cy:.0f}")
