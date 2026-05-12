"""
Why does the detector miss 5АЭ on plans 007/008/009/010?

For each plan, list ALL red clusters whose AR is in 5АЭ canonical OR horizontal
window (regardless of n_parts). Print W/H/AR/n.
"""
from __future__ import annotations
import io, sys, json
from collections import defaultdict
from pathlib import Path

import fitz, numpy as np, cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF_DIR = ROOT / r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF"
TPL_DIR = ROOT / "_shape_005_out"
MANIFEST = TPL_DIR / "templates_curated_v2.json"

PLANS = [
    "007-Планы освещения-отм. +7.800 +9.000.pdf",
    "008-Планы освещения-отм. +13.800.pdf",
    "009-Планы освещения-отм. +18.600.pdf",
    "010-Планы освещения-отм. +23.400.pdf",
]
LINK_DIST = 5.0
MAX_PRIM = 60.0

manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
info_v = manifest["5АЭ"]
info_h = manifest["5АЭ_h"]
print(f"5АЭ vert  : AR {info_v['ar_lo']}–{info_v['ar_hi']}, n {info_v['n_min']}–{info_v['n_max']}")
print(f"5АЭ horiz : AR {info_h['ar_lo']}–{info_h['ar_hi']}, n {info_h['n_min']}–{info_h['n_max']}")
print()

def color_red(c):
    if c is None: return False
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        return r > 0.6 and g < 0.4 and b < 0.4
    return False

def dbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]; xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
        elif it[0] in ("l", "m", "c"):
            for p in it[1:]:
                if hasattr(p, "x"): xs.append(p.x); ys.append(p.y)
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None

def cluster(prims, link):
    n = len(prims); par = list(range(n))
    def f(a):
        while par[a] != a: par[a] = par[par[a]]; a = par[a]
        return a
    def u(a, b):
        ra, rb = f(a), f(b)
        if ra != rb: par[ra] = rb
    bins = defaultdict(list)
    for i, p in enumerate(prims):
        bins[(int(p["cx"] // link), int(p["cy"] // link))].append(i)
    for (bx, by), idxs in bins.items():
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                ne = bins.get((bx+dx, by+dy), [])
                for i in idxs:
                    pi = prims[i]
                    for j in ne:
                        if j <= i: continue
                        pj = prims[j]
                        if abs(pi["cx"]-pj["cx"]) <= link and abs(pi["cy"]-pj["cy"]) <= link:
                            u(i, j)
    g = defaultdict(list)
    for i in range(n): g[f(i)].append(i)
    return list(g.values())

sys.path.insert(0, str(ROOT))
from pdf_legend_parser import parse_legend

for pname in PLANS:
    pdf_path = PDF_DIR / pname
    legend = parse_legend(str(pdf_path))
    LX0, LY0, LX1, LY1 = legend.legend_bbox

    doc = fitz.open(str(pdf_path)); mp = doc[legend.page_index]
    prims = []
    for d in mp.get_drawings():
        bb = dbox(d)
        if not bb: continue
        w, h = bb[2]-bb[0], bb[3]-bb[1]
        if max(w, h) > MAX_PRIM: continue
        col = color_red(d.get("fill")) or color_red(d.get("color"))
        if not col: continue
        cx, cy = (bb[0]+bb[2])/2, (bb[1]+bb[3])/2
        if LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2: continue
        prims.append({"bbox": bb, "cx": cx, "cy": cy})
    cls = cluster(prims, LINK_DIST)

    print(f"\n=== {pname[:3]}  (red prims={len(prims)}, clusters={len(cls)}) ===")
    candidates = []
    for cl in cls:
        if len(cl) < 10: continue
        xs = [prims[i]["bbox"][0] for i in cl] + [prims[i]["bbox"][2] for i in cl]
        ys = [prims[i]["bbox"][1] for i in cl] + [prims[i]["bbox"][3] for i in cl]
        bw = max(xs)-min(xs); bh = max(ys)-min(ys)
        if max(bw, bh) < 5.0: continue
        ar = bh/max(bw, 0.01)
        cx = (min(xs)+max(xs))/2; cy = (min(ys)+max(ys))/2
        # is AR in any 5АЭ window?
        in_v = (info_v["ar_lo"] <= ar <= info_v["ar_hi"]) or (info_v["ar_lo_rotated"] <= ar <= info_v["ar_hi_rotated"])
        in_h = (info_h["ar_lo"] <= ar <= info_h["ar_hi"]) or (info_h["ar_lo_rotated"] <= ar <= info_h["ar_hi_rotated"])
        if not (in_v or in_h): continue
        candidates.append((cx, cy, bw, bh, ar, len(cl), in_v, in_h))
    candidates.sort(key=lambda c: (-c[5]))  # by n desc
    print(f"  candidates with 5АЭ-like AR: {len(candidates)}")
    for cx, cy, bw, bh, ar, n, iv, ih in candidates[:20]:
        flags = ("V" if iv else " ") + ("H" if ih else " ")
        gate_v = info_v["n_min"] <= n <= info_v["n_max"]
        gate_h = info_h["n_min"] <= n <= info_h["n_max"]
        passes = ("v✓" if (iv and gate_v) else "v✗") + " " + ("h✓" if (ih and gate_h) else "h✗")
        print(f"    cx={cx:>5.0f} cy={cy:>5.0f}  W={bw:>5.1f} H={bh:>5.1f} AR={ar:>4.2f}  n={n:>4d}  [{flags}] gate:{passes}")
    doc.close()
