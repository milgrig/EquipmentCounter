"""
Visual deep-dive: render high-DPI crops of legend rows AND a sample plan
pictogram, with overlay of every drawing primitive (with its bbox + colour).

Output: PNGs in _inspect_005_out/.
"""
from __future__ import annotations
import io, sys, re
from collections import Counter, defaultdict
from pathlib import Path
import fitz
import numpy as np
import cv2
import pdfplumber

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_inspect_005_out")
OUT.mkdir(exist_ok=True)

DPI = 600  # very high for clear inspection

def color_name(c):
    if c is None: return "—"
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b = c[0],c[1],c[2]
        if r>0.6 and g<0.4 and b<0.4: return "RED"
        if r<0.4 and g<0.4 and b>0.6: return "BLUE"
        if r>0.85 and g>0.85 and b>0.85: return "white"
        if r<0.15 and g<0.15 and b<0.15: return "black"
        return f"({r:.2f},{g:.2f},{b:.2f})"
    return str(c)

def color_bgr(c):
    """OpenCV BGR for overlay."""
    if c is None: return (180,180,180)
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b = c[0],c[1],c[2]
        if r>0.6 and g<0.4 and b<0.4: return (0,0,255)        # red
        if r<0.4 and g<0.4 and b>0.6: return (255,0,0)        # blue
        if r<0.15 and g<0.15 and b<0.15: return (50,50,50)    # black
        return (180,180,180)
    return (180,180,180)

def dbox(d):
    xs,ys=[],[]
    for it in d.get("items",[]):
        if it[0]=="re":
            r=it[1]; xs+=[r.x0,r.x1]; ys+=[r.y0,r.y1]
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None

sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LEG_X0, LEG_Y0, LEG_X1, LEG_Y1 = legend.legend_bbox

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]
zoom = DPI/72.0

# ----------------------------------------------------------------------------
# Helper: render a crop and overlay all drawings with their bboxes
# ----------------------------------------------------------------------------
def render_region_overlay(region, save_name, label="", desc=""):
    """region = (x0,y0,x1,y1) in PDF pt."""
    rx0, ry0, rx1, ry1 = region
    clip = fitz.Rect(rx0, ry0, rx1, ry1)
    pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4: img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    over = img.copy()

    drawings_in_region = 0
    by_color = Counter()
    for d in mp.get_drawings():
        bb = dbox(d)
        if bb is None: continue
        # intersection with region?
        if bb[2] < rx0 or bb[0] > rx1 or bb[3] < ry0 or bb[1] > ry1:
            continue
        col = d.get("fill") or d.get("color")
        cn = color_name(col)
        by_color[cn] += 1
        drawings_in_region += 1
        bgr = color_bgr(col)
        # convert PDF pt to crop pixels
        px0 = int((bb[0] - rx0) * zoom)
        py0 = int((bb[1] - ry0) * zoom)
        px1 = int((bb[2] - rx0) * zoom)
        py1 = int((bb[3] - ry0) * zoom)
        cv2.rectangle(over, (px0,py0), (px1,py1), bgr, 1)

    # save raw + overlay side by side
    h = img.shape[0]; w = img.shape[1]
    canvas = np.full((h, w*2 + 30, 3), 255, dtype=np.uint8)
    canvas[:h, :w] = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    canvas[:h, w+30:w+30+w] = cv2.cvtColor(over, cv2.COLOR_RGB2BGR)
    # text
    cv2.putText(canvas, f"{label} | drawings:{drawings_in_region}", (5,20),
                cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0,0,0), 1)
    cv2.imwrite(str(OUT / save_name), canvas)
    print(f"  Saved {save_name} ({w}x{h}px) drawings={drawings_in_region} colors={by_color.most_common(5)}")
    print(f"    region in pt: ({rx0:.1f}, {ry0:.1f}) — ({rx1:.1f}, {ry1:.1f}) | desc={desc}")

# ----------------------------------------------------------------------------
# 1) Render each legend row's pictogram column with overlay
# ----------------------------------------------------------------------------
print("=== Legend rows: pictogram column ===")
for item in legend.items:
    if not item.symbol: continue
    if item.symbol not in {"1","1А","2","2А","3","3А","4","4АЭ","5АЭ","6АЭ","7АЭ"}: continue
    bx0, by0, bx1, by1 = item.bbox
    pad_y = 5.0
    # pictogram region: from LEG_X0 to bx0 (text starts), full row Y ±pad
    region = (LEG_X0 - 2, by0 - pad_y, bx0 + 2, by1 + pad_y)
    safe = re.sub(r"[^\w]", "_", item.symbol)
    render_region_overlay(region, f"legend_row_{safe}.png",
                          label=f"legend '{item.symbol}'",
                          desc=(item.description or "")[:60])

# ----------------------------------------------------------------------------
# 2) Render ALL legend pictogram column (one big strip)
# ----------------------------------------------------------------------------
print()
print("=== Whole legend pictogram strip ===")
# find the leftmost text-bbox start (defines pictogram column right edge)
text_left = min((it.bbox[0] for it in legend.items if it.symbol), default=LEG_X1)
strip = (LEG_X0 - 2, LEG_Y0 - 2, text_left + 5, LEG_Y1 + 2)
render_region_overlay(strip, "legend_strip.png", label="legend strip", desc="all rows")

# ----------------------------------------------------------------------------
# 3) Find one '1А' label on the plan and render its surroundings
# ----------------------------------------------------------------------------
print()
print("=== Sample plan pictograms near text labels ===")
LABEL_RE = re.compile(r"^([1-7])(А|АЭ)$")
SAMPLES = {"1А": 3, "1АЭ": 3, "2А": 2, "5АЭ": 2, "6АЭ": 2}
collected = defaultdict(list)
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    for w in page.extract_words() or []:
        t = (w.get("text") or "").strip()
        if t not in SAMPLES: continue
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        # outside legend
        if LEG_X0-2 <= cx <= LEG_X1+2 and LEG_Y0-2 <= cy <= LEG_Y1+2: continue
        if len(collected[t]) < SAMPLES[t]:
            collected[t].append((cx, cy))

R = 12.0  # half-width of crop in pt
for label, pts in collected.items():
    for i, (cx, cy) in enumerate(pts):
        region = (cx - R, cy - R, cx + R, cy + R)
        safe = re.sub(r"[^\w]", "_", label)
        render_region_overlay(region, f"plan_{safe}_{i}.png",
                              label=f"plan '{label}' #{i} @ ({cx:.0f},{cy:.0f})",
                              desc=f"text label center")

doc.close()
print()
print(f"Output: {OUT}")
