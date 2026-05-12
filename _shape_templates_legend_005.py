"""
Build shape templates DIRECTLY from the legend rows.

The legend parser gives each row a bbox and a name. Inside the row's
left half (pictogram column) we have stroke primitives that form the
icon. We rasterise them once per row -> the canonical template.
"""
from __future__ import annotations
import io, sys, json, re
from collections import Counter, defaultdict
from pathlib import Path

import fitz
import numpy as np
import cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_shape_005_out")
OUT.mkdir(exist_ok=True)

TPL_SIZE = 64
MAX_PRIM = 60.0
TARGET_NAMES = {"5АЭ", "6АЭ", "7АЭ"}

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

sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

# index ALL stroke/fill primitives (any colour), with bbox center
prims = []
for d in mp.get_drawings():
    bb = dbox(d)
    if bb is None: continue
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    if max(w, h) > MAX_PRIM: continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col is None: continue
    prims.append({"bbox": bb, "color": col,
                  "cx": (bb[0] + bb[2]) / 2, "cy": (bb[1] + bb[3]) / 2})

def rasterise(idxs, tpl_size=TPL_SIZE, stroke=2):
    if not idxs:
        return np.zeros((tpl_size, tpl_size), dtype=np.float32)
    xs0 = [prims[i]["bbox"][0] for i in idxs]; ys0 = [prims[i]["bbox"][1] for i in idxs]
    xs1 = [prims[i]["bbox"][2] for i in idxs]; ys1 = [prims[i]["bbox"][3] for i in idxs]
    bb = (min(xs0), min(ys0), max(xs1), max(ys1))
    bw, bh = bb[2] - bb[0], bb[3] - bb[1]
    if bw <= 0 or bh <= 0:
        return np.zeros((tpl_size, tpl_size), dtype=np.float32)
    scale = 8.0
    W = max(1, int(round(bw * scale))); H = max(1, int(round(bh * scale)))
    img = np.zeros((H, W), dtype=np.uint8)
    for i in idxs:
        p = prims[i]
        x0 = (p["bbox"][0] - bb[0]) * scale
        y0 = (p["bbox"][1] - bb[1]) * scale
        x1 = (p["bbox"][2] - bb[0]) * scale
        y1 = (p["bbox"][3] - bb[1]) * scale
        if (x1 - x0) < 1 and (y1 - y0) < 1:
            cv2.circle(img, (int((x0 + x1) / 2), int((y0 + y1) / 2)), max(1, stroke // 2), 255, -1)
        elif (x1 - x0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x0), int(y1)), 255, stroke)
        elif (y1 - y0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x1), int(y0)), 255, stroke)
        else:
            cv2.rectangle(img, (int(x0), int(y0)), (int(x1), int(y1)), 255, stroke)
    long = max(W, H)
    nW = max(1, int(round(W * tpl_size / long)))
    nH = max(1, int(round(H * tpl_size / long)))
    resized = cv2.resize(img, (nW, nH), interpolation=cv2.INTER_AREA)
    canvas = np.zeros((tpl_size, tpl_size), dtype=np.uint8)
    off_x = (tpl_size - nW) // 2; off_y = (tpl_size - nH) // 2
    canvas[off_y:off_y + nH, off_x:off_x + nW] = resized
    return canvas.astype(np.float32) / 255.0, bb, (bw, bh)

manifest = {}
print(f"Legend rows: {len(legend.items)}")
for it in legend.items:
    name = (it.symbol or "").strip()
    if name not in TARGET_NAMES:
        continue
    rx0, ry0, rx1, ry1 = it.bbox
    # pictogram column: left half of the row
    px1 = rx0 + (rx1 - rx0) * 0.45
    # collect coloured primitives whose centre lies in the pictogram column
    idxs = []
    for k, p in enumerate(prims):
        if rx0 - 1 <= p["cx"] <= px1 + 1 and ry0 - 1 <= p["cy"] <= ry1 + 1:
            # only red strokes (icons) - skip black text glyphs
            if p["color"] == "red":
                idxs.append(k)
    if not idxs:
        # fallback: any colour stroke (some icons may be black)
        for k, p in enumerate(prims):
            if rx0 - 1 <= p["cx"] <= px1 + 1 and ry0 - 1 <= p["cy"] <= ry1 + 1:
                if p["color"] in ("red", "blue", "black"):
                    idxs.append(k)
    if not idxs:
        print(f"  {name}: no primitives in row")
        continue
    img, bb, (bw, bh) = rasterise(idxs)
    safe = re.sub(r"[^\w]", "_", name)
    np.save(str(OUT / f"tpl_legend_{safe}.npy"), img)
    cv2.imwrite(str(OUT / f"tpl_legend_{safe}.png"), (img * 255).astype(np.uint8))
    manifest[name] = {"row_bbox": list(it.bbox),
                      "icon_bbox": list(bb),
                      "n_prims": len(idxs),
                      "w_pt": round(bw, 2), "h_pt": round(bh, 2),
                      "ar_h_over_w": round(bh / max(bw, 0.01), 2)}
    print(f"  {name}: n={len(idxs)}  W={bw:.1f}pt  H={bh:.1f}pt  AR={bh/max(bw,0.01):.2f}")

(OUT / "templates_legend.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
doc.close()
