"""
Build shape fingerprints (bitmap + aspect) per class.

For each pictogram class with reliable text labels (5АЭ, 6АЭ, 7АЭ), we:
  1. Find every text-label location on the plan.
  2. Take a fixed search window around each label (W=40pt x H=40pt below).
  3. Render the window at high DPI (mask of red drawing primitives only).
  4. Crop the mask to its content bbox -> shape silhouette.
  5. Resize to a canonical 32x32 fingerprint.
  6. Aggregate per-class via mean -> "prototype mask" + size stats.

Output: _shape_005_out/{class}_proto.png + shapes.json
"""
from __future__ import annotations
import io, sys, json
from collections import defaultdict
from pathlib import Path
from statistics import median

import fitz
import numpy as np
import cv2
import pdfplumber

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_shape_005_out")
OUT.mkdir(exist_ok=True)

DPI = 300
ZOOM = DPI / 72.0
PROTO = 32                        # canonical fingerprint size
SEARCH_HALF = 20.0                # search window half-size (pt) around label

# Where to look around the text label for the pictogram (label is usually
# above-left of the icon body). We sweep label center +/- SEARCH_HALF in x,y.
TARGETS = ["5АЭ", "6АЭ", "7АЭ"]
MAX_PRIM = 60.0


def color_class(c):
    if c is None: return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
        if r < 0.4 and g < 0.4 and b > 0.6: return "blue"
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


sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

# Index red small primitives outside legend.
prims = []
for d in mp.get_drawings():
    bb = dbox(d)
    if bb is None: continue
    w, h = bb[2]-bb[0], bb[3]-bb[1]
    if max(w, h) > MAX_PRIM: continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col != "red": continue
    cx, cy = (bb[0]+bb[2])/2, (bb[1]+bb[3])/2
    if LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2: continue
    prims.append({"bbox": bb, "cx": cx, "cy": cy, "w": w, "h": h})
print(f"red primitives: {len(prims)}")

# Spatial bins for fast windowed lookup
GRID = 10.0
grid = defaultdict(list)
for i, p in enumerate(prims):
    grid[(int(p["cx"]//GRID), int(p["cy"]//GRID))].append(i)


def in_window(x0, y0, x1, y1):
    out = []
    for bx in range(int(x0//GRID)-1, int(x1//GRID)+2):
        for by in range(int(y0//GRID)-1, int(y1//GRID)+2):
            for idx in grid.get((bx, by), []):
                p = prims[idx]
                if x0 <= p["cx"] <= x1 and y0 <= p["cy"] <= y1:
                    out.append(idx)
    return out


def shape_from_indices(idx_list):
    """Build a binary mask of selected primitives, crop to content, resize."""
    if not idx_list:
        return None, None
    xs0=[prims[i]["bbox"][0] for i in idx_list]
    ys0=[prims[i]["bbox"][1] for i in idx_list]
    xs1=[prims[i]["bbox"][2] for i in idx_list]
    ys1=[prims[i]["bbox"][3] for i in idx_list]
    bx0,by0,bx1,by1=min(xs0),min(ys0),max(xs1),max(ys1)
    w_pt, h_pt = bx1-bx0, by1-by0
    if w_pt < 0.5 or h_pt < 0.5:
        return None, None
    # rasterise into a binary mask at ZOOM
    pw = max(4, int(round(w_pt*ZOOM)))
    ph = max(4, int(round(h_pt*ZOOM)))
    mask = np.zeros((ph, pw), dtype=np.uint8)
    for i in idx_list:
        b = prims[i]["bbox"]
        x0 = int(round((b[0]-bx0)*ZOOM)); y0 = int(round((b[1]-by0)*ZOOM))
        x1 = max(x0+1, int(round((b[2]-bx0)*ZOOM)))
        y1 = max(y0+1, int(round((b[3]-by0)*ZOOM)))
        cv2.rectangle(mask, (x0,y0), (x1-1,y1-1), 255, thickness=-1)
    proto = cv2.resize(mask, (PROTO, PROTO), interpolation=cv2.INTER_AREA)
    _, proto = cv2.threshold(proto, 32, 255, cv2.THRESH_BINARY)
    return proto, (w_pt, h_pt)


# Collect labels
label_pts = defaultdict(list)
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    for w in page.extract_words() or []:
        t = (w.get("text") or "").strip()
        if t not in TARGETS: continue
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        if LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2: continue
        label_pts[t].append((cx, cy, w["x0"], w["top"], w["x1"], w["bottom"]))

print()
print("Building per-class shape fingerprints:")

shapes = {}
for lab in TARGETS:
    pts = label_pts.get(lab, [])
    if not pts: continue
    instances = []
    for k, (cx, cy, tx0, ty0, tx1, ty1) in enumerate(pts):
        # Window around label: search +/- SEARCH_HALF
        wx0, wy0 = cx-SEARCH_HALF, cy-SEARCH_HALF
        wx1, wy1 = cx+SEARCH_HALF, cy+SEARCH_HALF
        idxs = in_window(wx0, wy0, wx1, wy1)
        # exclude primitives within text bbox
        idxs = [i for i in idxs
                if not (tx0-0.5<=prims[i]["cx"]<=tx1+0.5 and ty0-0.5<=prims[i]["cy"]<=ty1+0.5)]
        if not idxs: continue

        # cluster: keep only the largest connected component (link=4pt)
        # Union-Find on this small set
        n=len(idxs); par=list(range(n))
        def f(a):
            while par[a]!=a: par[a]=par[par[a]]; a=par[a]
            return a
        def u(a,b):
            ra,rb=f(a),f(b)
            if ra!=rb: par[ra]=rb
        LK = 4.0
        for i in range(n):
            for j in range(i+1, n):
                if abs(prims[idxs[i]]["cx"]-prims[idxs[j]]["cx"])<=LK and abs(prims[idxs[i]]["cy"]-prims[idxs[j]]["cy"])<=LK:
                    u(i,j)
        comps = defaultdict(list)
        for i in range(n): comps[f(i)].append(idxs[i])
        # pick comp with most primitives
        best = max(comps.values(), key=len)
        proto, sz = shape_from_indices(best)
        if proto is None: continue
        instances.append({"proto": proto, "w_pt": sz[0], "h_pt": sz[1], "n": len(best)})

    if not instances:
        continue
    # average prototype
    stack = np.stack([i["proto"].astype(np.float32) for i in instances], axis=0)
    proto_mean = stack.mean(axis=0)
    proto_mean = (proto_mean > 96).astype(np.uint8)*255
    cv2.imwrite(str(OUT / f"proto_{lab.encode('ascii','replace').decode().replace('?','x')}.png"), proto_mean)

    ws = [i["w_pt"] for i in instances]
    hs = [i["h_pt"] for i in instances]
    ns = [i["n"] for i in instances]

    shapes[lab] = {
        "n_instances": len(instances),
        "median_w_pt": round(median(ws), 2),
        "median_h_pt": round(median(hs), 2),
        "median_aspect_h_w": round(median([h/max(0.01,w) for w,h in zip(ws,hs)]), 3),
        "min_w_pt": round(min(ws),2), "max_w_pt": round(max(ws),2),
        "min_h_pt": round(min(hs),2), "max_h_pt": round(max(hs),2),
        "median_n": int(median(ns)),
        "proto_path": f"proto_{lab.encode('ascii','replace').decode().replace('?','x')}.png",
        "proto_density": round(float((proto_mean>0).mean()), 3),
    }
    print(f"  {lab}: n_inst={len(instances)}  "
          f"W={shapes[lab]['median_w_pt']}  H={shapes[lab]['median_h_pt']}  "
          f"aspect(H/W)={shapes[lab]['median_aspect_h_w']}  "
          f"density={shapes[lab]['proto_density']}  parts={shapes[lab]['median_n']}")

# Save prototypes as a single side-by-side panel for quick viewing
panel_imgs = []
for lab in TARGETS:
    if lab not in shapes: continue
    img = cv2.imread(str(OUT / shapes[lab]["proto_path"]), cv2.IMREAD_GRAYSCALE)
    img_big = cv2.resize(img, (128, 128), interpolation=cv2.INTER_NEAREST)
    img_color = cv2.cvtColor(img_big, cv2.COLOR_GRAY2BGR)
    cv2.putText(img_color, lab.encode('ascii','replace').decode().replace('?','x'),
                (4, 18), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 1, cv2.LINE_AA)
    panel_imgs.append(img_color)
if panel_imgs:
    panel = np.hstack(panel_imgs)
    cv2.imwrite(str(OUT / "all_protos.png"), panel)

(OUT / "shapes.json").write_text(json.dumps(shapes, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"\nSaved: {OUT}")
doc.close()
