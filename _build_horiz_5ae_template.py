"""
Build a separate template for the HORIZONTAL 5AE variant ("ВЫХОД-arrow").

Strategy:
  * Re-scan plan 005 red primitives, cluster at LINK_DIST=5pt.
  * Keep clusters whose AR is in horiz window (0.35..0.55) AND n_parts in 100..200.
  * Drop clusters that fall inside legend bbox.
  * Save averaged 64x64 template + write a new entry "5АЭ_h" into
    templates_curated_v2.json (if не существует — добавляется).
"""
from __future__ import annotations
import io, sys, json
from collections import defaultdict
from pathlib import Path

import fitz, numpy as np, cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF = ROOT / r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf"
TPL_DIR = ROOT / "_shape_005_out"
MANIFEST = TPL_DIR / "templates_curated_v2.json"

LINK_DIST = 5.0
MAX_PRIM = 60.0
TPL_SIZE = 64
AR_LO_H, AR_HI_H = 0.35, 0.55  # horiz only
N_MIN_H, N_MAX_H = 100, 220

sys.path.insert(0, str(ROOT))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox

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

doc = fitz.open(str(PDF)); mp = doc[legend.page_index]
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
print(f"Red plan primitives: {len(prims)}")

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

def rasterise(idxs, tpl_size=TPL_SIZE, stroke=2):
    if not idxs: return None
    xs0 = [prims[i]["bbox"][0] for i in idxs]; ys0 = [prims[i]["bbox"][1] for i in idxs]
    xs1 = [prims[i]["bbox"][2] for i in idxs]; ys1 = [prims[i]["bbox"][3] for i in idxs]
    bb = (min(xs0), min(ys0), max(xs1), max(ys1))
    bw, bh = bb[2]-bb[0], bb[3]-bb[1]
    if bw <= 0 or bh <= 0: return None
    s = 8.0; W = max(1, int(round(bw*s))); H = max(1, int(round(bh*s)))
    img = np.zeros((H, W), dtype=np.uint8)
    for i in idxs:
        p = prims[i]
        x0 = (p["bbox"][0]-bb[0])*s; y0 = (p["bbox"][1]-bb[1])*s
        x1 = (p["bbox"][2]-bb[0])*s; y1 = (p["bbox"][3]-bb[1])*s
        if (x1-x0) < 1 and (y1-y0) < 1:
            cv2.circle(img, (int((x0+x1)/2), int((y0+y1)/2)), max(1, stroke//2), 255, -1)
        elif (x1-x0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x0), int(y1)), 255, stroke)
        elif (y1-y0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x1), int(y0)), 255, stroke)
        else:
            cv2.rectangle(img, (int(x0), int(y0)), (int(x1), int(y1)), 255, stroke)
    long = max(W, H)
    nW = max(1, int(round(W*tpl_size/long))); nH = max(1, int(round(H*tpl_size/long)))
    re_ = cv2.resize(img, (nW, nH), interpolation=cv2.INTER_AREA)
    cnv = np.zeros((tpl_size, tpl_size), dtype=np.uint8)
    cnv[(tpl_size-nH)//2:(tpl_size-nH)//2+nH, (tpl_size-nW)//2:(tpl_size-nW)//2+nW] = re_
    return cnv.astype(np.float32)/255.0, bb, (bw, bh)

cls = cluster(prims, LINK_DIST)
print(f"Clusters: {len(cls)}")

candidates = []
for cl in cls:
    if not (N_MIN_H <= len(cl) <= N_MAX_H): continue
    out = rasterise(cl)
    if out is None: continue
    img, bb, (bw, bh) = out
    if max(bw, bh) < 5.0: continue
    ar = bh/max(bw, 0.01)
    if not (AR_LO_H <= ar <= AR_HI_H): continue
    candidates.append({"img": img, "bb": bb, "bw": bw, "bh": bh, "n": len(cl), "ar": ar})

print(f"Horiz 5АЭ candidates: {len(candidates)}")
for c in candidates:
    print(f"  cx={(c['bb'][0]+c['bb'][2])/2:.0f} cy={(c['bb'][1]+c['bb'][3])/2:.0f} "
          f"W={c['bw']:.1f} H={c['bh']:.1f} AR={c['ar']:.2f} n={c['n']}")

if not candidates:
    print("NO horizontal 5АЭ candidates found — aborting"); sys.exit(0)

# Average the templates
imgs = np.stack([c["img"] for c in candidates], axis=0)
tpl = imgs.mean(axis=0)
np.save(str(TPL_DIR / "tpl_curated_5АЭ_h.npy"), tpl)
# also save a viewable PNG
cv2.imwrite(str(TPL_DIR / "tpl_curated_5АЭ_h.png"), (tpl*255).astype(np.uint8))

med_w = float(np.median([c["bw"] for c in candidates]))
med_h = float(np.median([c["bh"] for c in candidates]))
med_n = int(np.median([c["n"] for c in candidates]))
med_ar = float(np.median([c["ar"] for c in candidates]))

manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
manifest["5АЭ_h"] = {
    "n_used": len(candidates),
    "ar_lo": round(AR_LO_H - 0.05, 2),
    "ar_hi": round(AR_HI_H + 0.05, 2),
    "ar_lo_rotated": round(1.0/(AR_HI_H + 0.05), 2),
    "ar_hi_rotated": round(1.0/(AR_LO_H - 0.05), 2),
    "median_w": round(med_w, 2),
    "median_h": round(med_h, 2),
    "median_n": med_n,
    "median_ar": round(med_ar, 2),
    "n_min": int(N_MIN_H * 0.85),
    "n_max": int(N_MAX_H * 1.10),
    "_note": "Horizontal ВЫХОД-arrow variant of 5АЭ. Counted toward 5АЭ.",
    "_alias_for": "5АЭ"
}
# also TIGHTEN canonical 5АЭ window so it stops accepting horiz form (avoid double-count)
manifest["5АЭ"]["ar_lo"] = 0.75
manifest["5АЭ"]["ar_hi"] = 1.40
manifest["5АЭ"]["ar_lo_rotated"] = 0.75
manifest["5АЭ"]["ar_hi_rotated"] = 1.40
manifest["5АЭ"]["_note"] = "Vertical 5АЭ (square). Horizontal variant -> 5АЭ_h."

MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"\nSaved tpl_curated_5АЭ_h.npy ({len(candidates)} samples, med n={med_n}, AR={med_ar:.2f})")
print(f"Updated manifest: {MANIFEST}")
doc.close()
