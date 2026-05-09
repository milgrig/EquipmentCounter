"""
Build per-class bitmap shape templates from text-anchored plan examples.

For each label in TARGET (5АЭ, 6АЭ, 7АЭ), this script:
  1. Finds the text label positions on the plan (outside the legend).
  2. Around each label, renders a tight crop of the page at high DPI,
     limited to RED drawing primitives only (so glyph strokes do not
     contaminate the template).
  3. Computes the cluster bbox (region growing on red primitives) and
     extracts the corresponding sub-image, then normalises it to
     TPL_SIZE x TPL_SIZE while preserving aspect ratio (longest side
     fits, the rest is padded with zero / background).
  4. Averages the normalised crops across instances -> the class
     template (a 2-D probability map of "ink").

Outputs (NEW directory _shape_005_out/):
  - tpl_<label>.png          : visualisation of the average template
  - tpl_<label>.npy          : raw float32 template (TPL_SIZE x TPL_SIZE)
  - tpl_<label>_inst_<i>.png : individual normalised crops used
  - templates.json           : metadata (size, source instances)
"""
from __future__ import annotations
import io, sys, json, re
from collections import Counter, defaultdict
from pathlib import Path

import fitz
import numpy as np
import cv2
import pdfplumber

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_shape_005_out")
OUT.mkdir(exist_ok=True)

DPI = 600
TPL_SIZE = 64                 # square size for the normalised template
SEED_RADIUS = 6.0             # initial neighbourhood for region growing
LINK_DIST = 4.0               # connectivity link distance
GROW_LIMIT_PT = 40.0          # bound the cluster
MAX_PRIM = 60.0
TARGET = ["5АЭ", "6АЭ", "7АЭ"]


def color_class(c):
    if c is None: return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
        if r < 0.4 and g < 0.4 and b > 0.6: return "blue"
        if r < 0.15 and g < 0.15 and b < 0.15: return "black"
    return None


def dbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]; xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
        elif it[0] in ("l", "m", "c"):
            for p in it[1:]:
                if hasattr(p, "x"):
                    xs.append(p.x); ys.append(p.y)
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None


# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox
print(f"Legend bbox: ({LX0:.1f},{LY0:.1f})-({LX1:.1f},{LY1:.1f})")

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]
zoom = DPI / 72.0

# Collect coloured small primitives outside the legend.
prims = []
for d in mp.get_drawings():
    bb = dbox(d)
    if bb is None: continue
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    if max(w, h) > MAX_PRIM: continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col != "red": continue
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    if LX0 - 2 <= cx <= LX1 + 2 and LY0 - 2 <= cy <= LY1 + 2: continue
    prims.append({"bbox": bb, "cx": cx, "cy": cy, "w": w, "h": h, "color": col})
print(f"Indexed red plan primitives: {len(prims)}")

# spatial grid for neighbour query
grid = defaultdict(list)
for i, p in enumerate(prims):
    grid[(int(p["cx"] // LINK_DIST), int(p["cy"] // LINK_DIST))].append(i)


def near(cx, cy, radius):
    out = []
    cx_lo = int((cx - radius) // LINK_DIST); cx_hi = int((cx + radius) // LINK_DIST)
    cy_lo = int((cy - radius) // LINK_DIST); cy_hi = int((cy + radius) // LINK_DIST)
    for bx in range(cx_lo, cx_hi + 1):
        for by in range(cy_lo, cy_hi + 1):
            for idx in grid.get((bx, by), []):
                p = prims[idx]
                if abs(p["cx"] - cx) <= radius and abs(p["cy"] - cy) <= radius:
                    out.append(idx)
    return out


def grow(seed, exclude_box=None):
    if not seed: return []
    cluster_idx = set(seed)
    frontier = list(seed)
    bb = _bb_of(cluster_idx)
    while frontier:
        new_front = []
        for i in frontier:
            p = prims[i]
            for j in near(p["cx"], p["cy"], LINK_DIST):
                if j in cluster_idx: continue
                pj = prims[j]
                if exclude_box:
                    ex0, ey0, ex1, ey1 = exclude_box
                    if ex0 - 0.5 <= pj["cx"] <= ex1 + 0.5 and ey0 - 0.5 <= pj["cy"] <= ey1 + 0.5:
                        continue
                nb0 = (min(bb[0], pj["bbox"][0]), min(bb[1], pj["bbox"][1]),
                       max(bb[2], pj["bbox"][2]), max(bb[3], pj["bbox"][3]))
                if (nb0[2] - nb0[0]) > GROW_LIMIT_PT or (nb0[3] - nb0[1]) > GROW_LIMIT_PT:
                    continue
                cluster_idx.add(j); new_front.append(j); bb = nb0
        frontier = new_front
    return list(cluster_idx)


def _bb_of(idxs):
    xs0 = [prims[i]["bbox"][0] for i in idxs]
    ys0 = [prims[i]["bbox"][1] for i in idxs]
    xs1 = [prims[i]["bbox"][2] for i in idxs]
    ys1 = [prims[i]["bbox"][3] for i in idxs]
    return (min(xs0), min(ys0), max(xs1), max(ys1))


# ---------------------------------------------------------------------------
# Render bbox using ONLY listed primitive indices (manual rasteriser):
# stroke each primitive with a 2-px line on a black canvas, returning
# a binary float32 image normalised to TPL_SIZE x TPL_SIZE preserving AR.
# ---------------------------------------------------------------------------
def rasterise(idxs, tpl_size=TPL_SIZE, stroke=2):
    if not idxs:
        return np.zeros((tpl_size, tpl_size), dtype=np.float32)
    bb = _bb_of(idxs)
    bw, bh = bb[2] - bb[0], bb[3] - bb[1]
    if bw <= 0 or bh <= 0:
        return np.zeros((tpl_size, tpl_size), dtype=np.float32)
    # render at high resolution then resize, to keep thin strokes
    scale = 8.0
    W = max(1, int(round(bw * scale)))
    H = max(1, int(round(bh * scale)))
    img = np.zeros((H, W), dtype=np.uint8)
    for i in idxs:
        p = prims[i]
        x0 = (p["bbox"][0] - bb[0]) * scale
        y0 = (p["bbox"][1] - bb[1]) * scale
        x1 = (p["bbox"][2] - bb[0]) * scale
        y1 = (p["bbox"][3] - bb[1]) * scale
        # treat each primitive bbox as a stroked rectangle (or line if degenerate)
        if (x1 - x0) < 1 and (y1 - y0) < 1:
            cv2.circle(img, (int((x0 + x1) / 2), int((y0 + y1) / 2)), max(1, stroke // 2), 255, -1)
        elif (x1 - x0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x0), int(y1)), 255, stroke)
        elif (y1 - y0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x1), int(y0)), 255, stroke)
        else:
            cv2.rectangle(img, (int(x0), int(y0)), (int(x1), int(y1)), 255, stroke)
    # fit into TPL_SIZE preserving aspect ratio
    long = max(W, H)
    nW = max(1, int(round(W * tpl_size / long)))
    nH = max(1, int(round(H * tpl_size / long)))
    resized = cv2.resize(img, (nW, nH), interpolation=cv2.INTER_AREA)
    canvas = np.zeros((tpl_size, tpl_size), dtype=np.uint8)
    off_x = (tpl_size - nW) // 2
    off_y = (tpl_size - nH) // 2
    canvas[off_y:off_y + nH, off_x:off_x + nW] = resized
    return canvas.astype(np.float32) / 255.0


# ---------------------------------------------------------------------------
# Collect text positions
# ---------------------------------------------------------------------------
label_pts = defaultdict(list)
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    for w in page.extract_words() or []:
        t = (w.get("text") or "").strip()
        if t not in TARGET: continue
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        if LX0 - 2 <= cx <= LX1 + 2 and LY0 - 2 <= cy <= LY1 + 2: continue
        label_pts[t].append((cx, cy, w["x0"], w["top"], w["x1"], w["bottom"]))

print()
for k in TARGET:
    print(f"  {k}: {len(label_pts.get(k, []))} text instances")


# ---------------------------------------------------------------------------
# Build templates
# ---------------------------------------------------------------------------
manifest = {}

for label in TARGET:
    pts = label_pts.get(label, [])
    if not pts:
        print(f"  {label}: no labels"); continue
    safe = re.sub(r"[^\w]", "_", label)

    inst_imgs = []
    inst_meta = []
    for k, (cx, cy, tx0, ty0, tx1, ty1) in enumerate(pts):
        # find a seed: nearest red primitive in 6pt of label centre
        seeds = near(cx, cy, SEED_RADIUS)
        seeds = [i for i in seeds
                 if not (tx0 - 0.5 <= prims[i]["cx"] <= tx1 + 0.5
                         and ty0 - 0.5 <= prims[i]["cy"] <= ty1 + 0.5)]
        if not seeds:
            continue
        # but we want the cluster of the actual icon, not text strokes:
        # take the seed nearest to cx,cy, expand to its full connected
        # component (not bounded by label centre), then collect.
        cluster_idx = grow(seeds, exclude_box=(tx0, ty0, tx1, ty1))
        bb = _bb_of(cluster_idx)
        bw, bh = bb[2] - bb[0], bb[3] - bb[1]
        # icon must be larger than 4pt on at least one side
        if max(bw, bh) < 4.0:
            continue
        img = rasterise(cluster_idx)
        inst_imgs.append(img)
        inst_meta.append({"idx": k, "bbox": bb, "n": len(cluster_idx),
                          "w": round(bw, 2), "h": round(bh, 2)})
        # per-instance preview
        cv2.imwrite(str(OUT / f"tpl_{safe}_inst_{k}.png"), (img * 255).astype(np.uint8))

    if not inst_imgs:
        print(f"  {label}: no usable instances"); continue

    avg = np.mean(np.stack(inst_imgs, axis=0), axis=0).astype(np.float32)
    np.save(str(OUT / f"tpl_{safe}.npy"), avg)
    cv2.imwrite(str(OUT / f"tpl_{safe}.png"), (avg * 255).astype(np.uint8))
    manifest[label] = {
        "tpl_size": TPL_SIZE,
        "n_instances_used": len(inst_imgs),
        "instances": inst_meta,
    }
    # diagnostic: average aspect ratio
    ars = [m["h"] / max(m["w"], 0.01) for m in inst_meta]
    print(f"  {label}: built from {len(inst_imgs)} instances, "
          f"AR(h/w)={np.mean(ars):.2f}  W~{np.mean([m['w'] for m in inst_meta]):.1f}pt  "
          f"H~{np.mean([m['h'] for m in inst_meta]):.1f}pt")

(OUT / "templates.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2),
                                     encoding="utf-8")
doc.close()
print()
print(f"Saved templates to {OUT}")
