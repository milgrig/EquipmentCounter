"""
Build shape templates for the THREE switch icons in the legend (rows 16/17/18)
and prove they can be told apart by NCC + circle-count + segment-count.

Output (NEW dir _switch_005_out/):
  - tpl_switch_<row>.png/.npy   (rasterised templates)
  - switch_signatures.json      (numeric signature: AR, n_segments, n_curve_blobs)
  - cross_ncc.txt               (NCC matrix between the three templates)
"""
import io, sys, json
from collections import Counter, defaultdict
from pathlib import Path
import fitz, numpy as np, cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_switch_005_out")
OUT.mkdir(exist_ok=True)
TPL_SIZE = 64

sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0 = legend.legend_bbox[0]
doc = fitz.open(str(PDF)); mp = doc[legend.page_index]

# Index ALL blue stroke segments (no size cap) and all blue fills.
segments = []           # tiny line bits (the ink)
all_drawings_in_region = []
for d in mp.get_drawings():
    col = d.get("color") or d.get("fill")
    if not col or len(col) < 3: continue
    if not (col[0] < 0.2 and col[1] < 0.2 and col[2] > 0.8): continue
    items = d.get("items", [])
    if not items: continue
    for it in items:
        if it[0] == "l" and len(it) >= 3 and hasattr(it[1], "x") and hasattr(it[2], "x"):
            x0, y0 = it[1].x, it[1].y; x1, y1 = it[2].x, it[2].y
            segments.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1,
                             "cx": (x0 + x1) / 2, "cy": (y0 + y1) / 2,
                             "len": ((x1-x0)**2 + (y1-y0)**2) ** 0.5})
        elif it[0] == "re":
            r = it[1]
            segments.append({"x0": r.x0, "y0": r.y0, "x1": r.x1, "y1": r.y1,
                             "cx": (r.x0 + r.x1) / 2, "cy": (r.y0 + r.y1) / 2,
                             "len": max(r.x1 - r.x0, r.y1 - r.y0), "rect": True})

print(f"Indexed blue segments: {len(segments)}")

def collect_in(rx0, ry0, rx1, ry1):
    return [s for s in segments if rx0 - 1 <= s["cx"] <= rx1 + 1 and ry0 - 1 <= s["cy"] <= ry1 + 1]

def count_circle_blobs(segs):
    """A circle blob = bin (1pt grid) holding >= 30 micro-segments."""
    bins = defaultdict(int)
    for s in segs:
        if s.get("len", 99) > 1.5: continue
        bins[(int(s["cx"]), int(s["cy"]))] += 1
    # merge adjacent heavy bins
    heavy = [(k, v) for k, v in bins.items() if v >= 25]
    heavy.sort(key=lambda x: -x[1])
    seen = set(); blobs = 0
    for (bx, by), _ in heavy:
        if (bx, by) in seen: continue
        # mark 5x5 around
        for dx in range(-2, 3):
            for dy in range(-2, 3):
                seen.add((bx+dx, by+dy))
        blobs += 1
    return blobs

def rasterise(segs, tpl_size=TPL_SIZE, stroke=2):
    if not segs: return np.zeros((tpl_size, tpl_size), dtype=np.float32), None, (0,0)
    xs0 = [s["x0"] for s in segs]; ys0 = [s["y0"] for s in segs]
    xs1 = [s["x1"] for s in segs]; ys1 = [s["y1"] for s in segs]
    bb = (min(xs0+xs1), min(ys0+ys1), max(xs0+xs1), max(ys0+ys1))
    bw, bh = bb[2]-bb[0], bb[3]-bb[1]
    if bw <= 0 or bh <= 0: return np.zeros((tpl_size, tpl_size), dtype=np.float32), bb, (bw, bh)
    scale = 8.0
    W = max(1, int(round(bw*scale))); H = max(1, int(round(bh*scale)))
    img = np.zeros((H, W), dtype=np.uint8)
    for s in segs:
        x0 = (s["x0"]-bb[0])*scale; y0 = (s["y0"]-bb[1])*scale
        x1 = (s["x1"]-bb[0])*scale; y1 = (s["y1"]-bb[1])*scale
        cv2.line(img, (int(x0), int(y0)), (int(x1), int(y1)), 255, stroke)
    long = max(W, H)
    nW = max(1, int(round(W*tpl_size/long))); nH = max(1, int(round(H*tpl_size/long)))
    re_ = cv2.resize(img, (nW, nH), interpolation=cv2.INTER_AREA)
    cnv = np.zeros((tpl_size, tpl_size), dtype=np.uint8)
    cnv[(tpl_size-nH)//2:(tpl_size-nH)//2+nH, (tpl_size-nW)//2:(tpl_size-nW)//2+nW] = re_
    return cnv.astype(np.float32) / 255.0, bb, (bw, bh)

def ncc(a, b):
    a = a - a.mean(); b = b - b.mean()
    na = np.linalg.norm(a); nb = np.linalg.norm(b)
    if na < 1e-6 or nb < 1e-6: return 0.0
    return float(np.sum(a*b) / (na*nb))

# Process rows 16, 17, 18
TARGET = [
    (16, "switch_simple",    "Одноклавишный (одиночный)"),
    (17, "switch_pereklyuch", "Переключатель"),
    (18, "control_post",     "Пост управления"),
]

manifest = {}
templates = {}
for ridx, key, hr_name in TARGET:
    it = legend.items[ridx]
    rx0, ry0, rx1, ry1 = it.bbox
    # icon column: from legend left edge up to description x0
    sx0, sx1 = LX0, rx0 + 5
    sy0, sy1 = ry0 - 8, ry1 + 4
    segs = collect_in(sx0, sy0, sx1, sy1)
    blobs = count_circle_blobs(segs)
    img, bb, (bw, bh) = rasterise(segs)
    if bb is None:
        print(f"  row {ridx}: empty"); continue
    ar = bh / max(bw, 0.01)
    np.save(str(OUT / f"tpl_{key}.npy"), img)
    cv2.imwrite(str(OUT / f"tpl_{key}.png"), (img*255).astype(np.uint8))
    templates[key] = img
    manifest[key] = {
        "row_idx": ridx, "human_name": hr_name,
        "bbox_pt": list(bb), "w_pt": round(bw, 2), "h_pt": round(bh, 2),
        "ar_h_over_w": round(ar, 3),
        "n_segments": len(segs),
        "n_circle_blobs": blobs,
    }
    print(f"  row {ridx} '{hr_name}':  W={bw:.1f}pt H={bh:.1f}pt AR={ar:.2f}  "
          f"segs={len(segs)}  circles={blobs}")

# Cross-NCC matrix
keys = list(templates.keys())
print("\nCross-NCC matrix:")
print("            " + "  ".join(f"{k:<18s}" for k in keys))
for k1 in keys:
    row = []
    for k2 in keys:
        v = ncc(templates[k1], templates[k2])
        row.append(f"{v:>+.3f}{'':<14s}")
    print(f"  {k1:<10s}  " + "  ".join(row))

(OUT / "switch_signatures.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"\nSaved templates and signatures to {OUT}")
doc.close()
