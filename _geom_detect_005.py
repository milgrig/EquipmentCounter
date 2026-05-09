"""
Geometric pictogram detector — text-free.

Pipeline:
  1. Read legend bbox via parse_legend (use legend_bbox + per-item bboxes).
  2. For each legend item: collect vector drawings inside its bbox; produce
     a descriptor (W, H, fill, stroke) AND a reference rendered crop.
  3. On the rest of the page: find ALL pure rectangles outside legend_bbox.
  4. Classify each plan rect against legend descriptors:
        - size/color filter (cheap)
        - normalized-crop pixel similarity (precise)
  5. (Control only) compare totals to DXF ground truth and to text labels.

Not modifying any existing module.
"""
from __future__ import annotations
import io
import sys
import re
from collections import Counter, defaultdict
from pathlib import Path
from statistics import median

import pdfplumber
import fitz
import numpy as np
import cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_geom_005_out")
OUT.mkdir(exist_ok=True)

DPI = 300
SIZE_TOL = 0.20         # ±20% allowed on width/height
SIM_THRESH = 0.55       # min normalized-pixel similarity
RECT_PAD = 1.0          # pt padding around rect for crop

DXF_GT = {"1":33, "2":15, "3":7, "4":4, "5АЭ":4, "6АЭ":6, "7АЭ":7}

# ---- helpers ---------------------------------------------------------------
def fmt_color(c):
    if c is None: return "—"
    if isinstance(c, (tuple, list)):
        return "(" + ",".join(f"{v:.2f}" for v in c) + ")"
    return str(c)

def color_class(c):
    """Coarse color category that's stable to PDF render quirks."""
    if c is None: return "none"
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.7 and g < 0.3 and b < 0.3: return "red"
        if r < 0.3 and g < 0.3 and b > 0.7: return "blue"
        if r > 0.9 and g > 0.9 and b > 0.9: return "white"
        if r < 0.1 and g < 0.1 and b < 0.1: return "black"
        return f"rgb({r:.1f},{g:.1f},{b:.1f})"
    return str(c)

def in_bbox(cx, cy, bbox, pad=0.0):
    x0, y0, x1, y1 = bbox
    return (x0 - pad) <= cx <= (x1 + pad) and (y0 - pad) <= cy <= (y1 + pad)

# ---- 1. Parse legend -------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
print(f"Legend: page={legend.page_index}, items={len(legend.items)}")
print(f"Legend bbox: {legend.legend_bbox}")
print()

# ---- 2. Read all drawings (rectangles) on page -----------------------------
doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

all_rects = []
for d in mp.get_drawings():
    items = d.get("items", [])
    if len(items) == 1 and items[0][0] == "re":
        r = items[0][1]
        all_rects.append({
            "bbox": (r.x0, r.y0, r.x1, r.y1),
            "cx": (r.x0 + r.x1) / 2.0,
            "cy": (r.y0 + r.y1) / 2.0,
            "w": r.x1 - r.x0,
            "h": r.y1 - r.y0,
            "fill": d.get("fill"),
            "color": d.get("color"),
        })
print(f"Pure rectangles total: {len(all_rects)}")

# Render whole page once (high DPI) to extract reference & candidate crops
zoom = DPI / 72.0
mat = fitz.Matrix(zoom, zoom)
pix = mp.get_pixmap(matrix=mat, alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4:
    img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
print(f"Page rendered: {img.shape} at {DPI} DPI")
print()

def render_crop(bbox, pad=RECT_PAD):
    x0, y0, x1, y1 = bbox
    px0 = max(0, int((x0 - pad) * zoom))
    py0 = max(0, int((y0 - pad) * zoom))
    px1 = min(img.shape[1], int((x1 + pad) * zoom))
    py1 = min(img.shape[0], int((y1 + pad) * zoom))
    if px1 <= px0 or py1 <= py0: return None
    return img[py0:py1, px0:px1].copy()

def normalize(crop, target=(80, 40)):
    if crop is None or crop.size == 0: return None
    return cv2.resize(crop, target, interpolation=cv2.INTER_AREA)

def similarity(a, b):
    """Normalized cross-correlation on grayscale."""
    if a is None or b is None or a.shape != b.shape:
        return 0.0
    ag = cv2.cvtColor(a, cv2.COLOR_RGB2GRAY).astype(np.float32)
    bg = cv2.cvtColor(b, cv2.COLOR_RGB2GRAY).astype(np.float32)
    ag = (ag - ag.mean()) / (ag.std() + 1e-6)
    bg = (bg - bg.mean()) / (bg.std() + 1e-6)
    return float(np.clip((ag * bg).mean(), 0.0, 1.0))

# ---- 3. Build legend descriptors -------------------------------------------
legend_rects_used = set()  # ids of all_rects taken by legend
descriptors = []   # list of {symbol, desc_text, w, h, fill_class, sample_crop, norm_ref}

# Pictograms sit to the LEFT of each row.  Extend search left to the start of legend_bbox.
LEG_X0 = legend.legend_bbox[0] if legend.legend_bbox and legend.legend_bbox[2] > 0 else None

print("=== Legend descriptors ===")
print(f"{'sym':>6s}{'fill':>6s}  {'W':>6s} x {'H':>6s}   desc")
for item in legend.items:
    bx0, by0, bx1, by1 = item.bbox
    # extend search bbox LEFT to legend_bbox.x0 (pictogram column)
    sx0 = LEG_X0 if LEG_X0 is not None and LEG_X0 < bx0 else bx0
    # also slightly extend Y to capture rects whose center is just above/below the row
    pad_y = (by1 - by0) * 0.3
    sy0, sy1 = by0 - pad_y, by1 + pad_y

    inside = []
    for i, r in enumerate(all_rects):
        # only consider rects where the WHOLE rect is left of bx0 (in the pictogram column)
        # OR the rect intersects the row band
        if (sx0 <= r["cx"] <= bx1) and (sy0 <= r["cy"] <= sy1):
            inside.append((i, r))
    if not inside:
        continue
    # pick rects whose width*height is "fixture-like" (>5 sq.pt) — drop tiny noise dots
    inside = [(i, r) for i, r in inside if r["w"] * r["h"] >= 5.0]
    if not inside:
        continue
    # group by fill color (one symbol can have BLUE working + RED emergency variants)
    by_fill = defaultdict(list)
    for i, r in inside:
        by_fill[color_class(r["fill"])].append((i, r))

    for fill_cls, irs in by_fill.items():
        # within same color, take the LARGEST rect (pictogram body)
        irs.sort(key=lambda ir: ir[1]["w"] * ir[1]["h"], reverse=True)
        best_idx, best_rect = irs[0]
        legend_rects_used.add(best_idx)

        crop = render_crop(best_rect["bbox"])
        norm_ref = normalize(crop)

        safe_sym = re.sub(r"[^\w]", "_", (item.symbol or f"item{len(descriptors)}") + "_" + fill_cls)
        if crop is not None:
            cv2.imwrite(str(OUT / f"legend_{safe_sym}.png"), cv2.cvtColor(crop, cv2.COLOR_RGB2BGR))

        descriptors.append({
            "symbol": item.symbol,
            "desc": (item.description or "")[:50],
            "w": best_rect["w"],
            "h": best_rect["h"],
            "fill": fill_cls,
            "norm_ref": norm_ref,
            "bbox": best_rect["bbox"],
        })
        print(f"{item.symbol or '—':>6s}{fill_cls:>6s}  {best_rect['w']:>6.2f} x {best_rect['h']:>6.2f}   {(item.description or '')[:50]}")
print()

# ---- 4. Plan rectangles (outside legend) -----------------------------------
plan_rects = []
lb = legend.legend_bbox
for i, r in enumerate(all_rects):
    if i in legend_rects_used: continue
    # also skip anything that falls inside the overall legend bbox
    if lb and lb[2] > 0 and in_bbox(r["cx"], r["cy"], lb, pad=2.0):
        continue
    plan_rects.append(r)
print(f"Plan rectangles (excluding legend): {len(plan_rects)}")

# ---- 5. Classify plan rects -------------------------------------------------
def best_descriptor(r):
    """Return (descriptor, score) — highest scoring legend match, or None."""
    cand = []
    plan_fill = color_class(r["fill"])
    crop = render_crop(r["bbox"])
    norm_cand = normalize(crop)
    for desc in descriptors:
        # 1) color must match (strict)
        if desc["fill"] != plan_fill:
            continue
        # 2) size ratio within tol (allow 90° rotation: try swapped W/H too)
        def size_ok(dw, dh):
            return (abs(r["w"] - dw) / max(dw, 1e-6) <= SIZE_TOL and
                    abs(r["h"] - dh) / max(dh, 1e-6) <= SIZE_TOL)
        rotated = False
        if size_ok(desc["w"], desc["h"]):
            pass
        elif size_ok(desc["h"], desc["w"]):
            rotated = True
        else:
            continue
        # 3) pixel similarity
        ref = desc["norm_ref"]
        if rotated and norm_cand is not None:
            n = cv2.rotate(norm_cand, cv2.ROTATE_90_CLOCKWISE)
            n = cv2.resize(n, (ref.shape[1], ref.shape[0])) if ref is not None else n
            sim = similarity(ref, n)
        else:
            sim = similarity(ref, norm_cand)
        cand.append((desc, sim))
    if not cand: return None, 0.0
    cand.sort(key=lambda x: x[1], reverse=True)
    return cand[0]

assignments = []   # (rect_idx, descriptor, sim)
for r in plan_rects:
    desc, sim = best_descriptor(r)
    if desc and sim >= SIM_THRESH:
        assignments.append((r, desc, sim))

print(f"Classified plan rects: {len(assignments)} / {len(plan_rects)}")
print()

# ---- 6. Aggregate per symbol ----------------------------------------------
counts = Counter()
sims_per_sym = defaultdict(list)
for r, desc, sim in assignments:
    counts[desc["symbol"]] += 1
    sims_per_sym[desc["symbol"]].append(sim)

print("=== Geometric counts vs DXF ground truth ===")
print(f"{'symbol':>8s} {'GEO':>5s} {'DXF':>5s} {'avg_sim':>8s}  description")
for desc in descriptors:
    s = desc["symbol"]
    if not s: continue
    geo = counts.get(s, 0)
    gt  = DXF_GT.get(s, "—")
    sims = sims_per_sym.get(s, [])
    avg = sum(sims)/len(sims) if sims else 0.0
    print(f"{s:>8s} {geo:>5d} {str(gt):>5s} {avg:>8.3f}  {desc['desc']}")
print()

# ---- 7. (Control) cross-check vs text labels -------------------------------
LABEL_RE = re.compile(r"^[1-9](?:[АA][ЭE]?)?$")
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    words = page.extract_words() or []

text_labels = []
for w in words:
    t = (w.get("text") or "").strip()
    if not LABEL_RE.match(t): continue
    cx = (w["x0"] + w["x1"]) / 2.0
    cy = (w["top"] + w["bottom"]) / 2.0
    if lb and lb[2] > 0 and in_bbox(cx, cy, lb, pad=2.0):
        continue
    text_labels.append({"text": t, "cx": cx, "cy": cy})

text_counts = Counter(l["text"] for l in text_labels)

# For DXF items like '1' = '1А' + '1АЭ' (working + emergency)
print("=== Text-control vs DXF ===")
print(f"{'sym':>5s}  text(1А+1АЭ)  geo  DXF")
for sym, gt in DXF_GT.items():
    if sym.endswith("АЭ"):  # 5АЭ, 6АЭ, 7АЭ — direct
        text_total = text_counts.get(sym, 0)
    else:
        text_total = text_counts.get(f"{sym}А", 0) + text_counts.get(f"{sym}АЭ", 0)
    geo = counts.get(sym, 0)
    print(f"{sym:>5s}  {text_total:>11d}  {geo:>3d}  {gt:>3d}")
print()

# ---- 8. Save annotated overlay ---------------------------------------------
overlay = img.copy()
sym_color = {
    "1":(255,0,255), "2":(0,255,255), "3":(255,128,0), "4":(128,0,255),
    "5АЭ":(0,255,0), "6АЭ":(0,200,200), "7АЭ":(200,0,200),
}
for r, desc, sim in assignments:
    x0, y0, x1, y1 = r["bbox"]
    color = sym_color.get(desc["symbol"] or "", (255,255,255))
    cv2.rectangle(overlay,
                  (int(x0*zoom)-3, int(y0*zoom)-3),
                  (int(x1*zoom)+3, int(y1*zoom)+3),
                  color, 2)

cv2.imwrite(str(OUT / "overlay.png"), cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))
print(f"Overlay saved: {OUT / 'overlay.png'}")
print(f"Legend reference crops: {OUT}")

doc.close()
