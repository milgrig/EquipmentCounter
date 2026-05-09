import io, sys
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
exec(open(r"C:\Cursor\TayfaProject\EquipmentCounter\_geo_multi_005.py", encoding="utf-8").read())
import numpy as np
print()
print("=== Big clusters (n>=100) and assignments ===")
for i in np.argsort(-cluster_n):
    if cluster_n[i] < 100: continue
    a = assignments[i]; f = F_all[i]
    bb = cluster_bboxes[i]; cx=(bb[0]+bb[2])/2; cy=(bb[1]+bb[3])/2
    print(f"  cl#{i} n={int(cluster_n[i])} W={f[0]:.1f} H={f[1]:.1f} aspect={f[2]:.2f} horiz={f[8]:.2f} cy={cy:.0f}  -> {a}")
