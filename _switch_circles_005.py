"""
Switch detection v0: find 'circle blobs' on the plan.

A circle (switch contact) in this PDF is a tight cluster of >=50 micro line
segments whose bbox fits within ~3pt x 3pt. We index all blue line segments
(stroke color (0,0,1)) and find such hot-spots.
"""
import io, sys, json
from collections import defaultdict, Counter
from pathlib import Path
import fitz, numpy as np, cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_switch_005_out")
OUT.mkdir(exist_ok=True)

sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox
doc = fitz.open(str(PDF)); mp = doc[legend.page_index]

# Index ALL blue tiny line segments
prims = []
for d in mp.get_drawings():
    col = d.get("color")
    if not col or len(col) < 3: continue
    if not (col[0] < 0.2 and col[1] < 0.2 and col[2] > 0.8):  # blue stroke
        continue
    items = d.get("items", [])
    if not items: continue
    # only tiny segments (the circle is made of these)
    for it in items:
        if it[0] != "l": continue
        if len(it) < 3: continue
        a, b = it[1], it[2]
        if not (hasattr(a, "x") and hasattr(b, "x")): continue
        x0, y0 = a.x, a.y
        x1, y1 = b.x, b.y
        seg_len = ((x1-x0)**2 + (y1-y0)**2) ** 0.5
        if seg_len > 1.5: continue   # only micro-segments
        cx = (x0 + x1) / 2; cy = (y0 + y1) / 2
        # exclude legend area? keep both for diag, mark via flag
        in_legend = LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2
        prims.append({"cx": cx, "cy": cy, "in_legend": in_legend})

print(f"Tiny blue segments total: {len(prims)}  (in_legend: {sum(1 for p in prims if p['in_legend'])})")

# Bin by 1pt grid; a circle contact = cell with very high count (> 50)
GRID = 1.0
bins = defaultdict(list)
for i, p in enumerate(prims):
    bins[(int(p["cx"] // GRID), int(p["cy"] // GRID))].append(i)

print(f"Total bins: {len(bins)}")
heavy = [(k, idxs) for k, idxs in bins.items() if len(idxs) >= 30]
print(f"Bins with >= 30 micro-segments: {len(heavy)}")

# For each heavy bin, examine the surrounding 3x3 mass to refine
def mass(bx, by, R=1):
    total = 0
    for dx in range(-R, R+1):
        for dy in range(-R, R+1):
            total += len(bins.get((bx+dx, by+dy), []))
    return total

contacts = []
seen = set()
for (bx, by), idxs in sorted(heavy, key=lambda x: -mass(x[0][0], x[0][1])):
    if (bx, by) in seen: continue
    m = mass(bx, by, R=2)
    if m < 50: continue
    # mark a 5x5 neighbourhood as visited so we don't double-count
    for dx in range(-2, 3):
        for dy in range(-2, 3):
            seen.add((bx+dx, by+dy))
    cx_mean = (bx + 0.5) * GRID
    cy_mean = (by + 0.5) * GRID
    in_leg = LX0-2 <= cx_mean <= LX1+2 and LY0-2 <= cy_mean <= LY1+2
    contacts.append({"cx": cx_mean, "cy": cy_mean, "mass": m, "in_legend": in_leg})

print()
print(f"Detected circle contacts (mass>=50): {len(contacts)}")
print(f"  outside legend: {sum(1 for c in contacts if not c['in_legend'])}")
print(f"  inside legend:  {sum(1 for c in contacts if c['in_legend'])}")

# Save overlay at 220 DPI marking each contact
zoom = 220 / 72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4: img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
canvas = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
cv2.rectangle(canvas, (int(LX0*zoom), int(LY0*zoom)), (int(LX1*zoom), int(LY1*zoom)), (160,160,160), 1)
for c in contacts:
    x = int(c["cx"]*zoom); y = int(c["cy"]*zoom)
    col = (200,200,200) if c["in_legend"] else (0, 200, 0)
    cv2.circle(canvas, (x, y), 12, col, 2)
cv2.imwrite(str(OUT/"contacts_overlay.png"), canvas)

# Save raw json
out_json = {"GRID_pt": GRID, "n_contacts": len(contacts),
            "contacts": [{"cx":round(c["cx"],2),"cy":round(c["cy"],2),
                          "mass":c["mass"],"in_legend":c["in_legend"]} for c in contacts]}
(OUT/"contacts.json").write_text(json.dumps(out_json, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"Saved: {OUT/'contacts.json'}, {OUT/'contacts_overlay.png'}")
doc.close()
