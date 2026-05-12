"""
Shape-aware text-free detector v2.

Improvements over v1:
  * Each template is matched in 4 orientations (0/90/180/270) to handle
    fixtures rotated on the plan. Score = max NCC across rotations.
  * AR gate is widened so a horizontal variant of a "vertical" icon is
    not rejected. We test ar against BOTH (ar_lo, ar_hi) AND its inverse
    window when the rotated template is squareish.
  * NCC_MIN lowered to 0.45 to admit slightly noisy clusters.
  * For each cluster, also compute a "winner margin" = best - second_best
    to suppress ambiguous hits.
"""
from __future__ import annotations
import io, sys, json
from collections import defaultdict
from pathlib import Path

import fitz
import numpy as np
import cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
TPL_DIR = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_shape_005_out")
TPL_MANIFEST = TPL_DIR / "templates_curated.json"
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_shape_det_v2_005_out")
OUT.mkdir(exist_ok=True)

DPI_OVER = 220
TPL_SIZE = 64
LINK_DIST = 5.0
MAX_PRIM = 60.0
NCC_MIN = 0.45
MARGIN_MIN = 0.10        # winner margin: best - second_best
MIN_PARTS = 18
DXF_GT = {"5АЭ": 4, "6АЭ": 6, "7АЭ": 7}
ASCII_MAP = {"5АЭ": "5AE", "6АЭ": "6AE", "7АЭ": "7AE"}
PALETTE = {"5АЭ": (0, 200, 0), "6АЭ": (255, 0, 255), "7АЭ": (0, 165, 255)}


def color_class(c):
    if c is None: return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
    return None

def dbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]; xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None


sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0,LY0,LX1,LY1 = legend.legend_bbox

manifest = json.loads(TPL_MANIFEST.read_text(encoding="utf-8"))

# Load templates and prepare 4 rotations each
TEMPLATES = {}      # lab -> list of 4 rotated tpls
AR_INFO = {}        # lab -> (ar_lo, ar_hi) of the canonical template
for lab in manifest:
    npy = TPL_DIR / f"tpl_curated_{lab}.npy"
    if not npy.exists():
        print(f"WARN: template missing for {lab}"); continue
    base = np.load(str(npy)).astype(np.float32)
    rots = [base,
            np.rot90(base, 1),
            np.rot90(base, 2),
            np.rot90(base, 3)]
    TEMPLATES[lab] = rots
    AR_INFO[lab] = (manifest[lab]["ar_lo"], manifest[lab]["ar_hi"])
print(f"Templates loaded: {list(TEMPLATES.keys())}")
for lab,info in manifest.items():
    print(f"  {lab}: AR window {info['ar_lo']}..{info['ar_hi']}, "
          f"median W={info['median_w']}pt H={info['median_h']}pt")


# ---------------------------------------------------------------------------
doc = fitz.open(str(PDF)); mp = doc[legend.page_index]

prims=[]
for d in mp.get_drawings():
    bb=dbox(d)
    if not bb: continue
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if max(w,h)>MAX_PRIM: continue
    col=color_class(d.get("fill")) or color_class(d.get("color"))
    if col!="red": continue
    cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
    if LX0-2<=cx<=LX1+2 and LY0-2<=cy<=LY1+2: continue
    prims.append({"bbox":bb,"cx":cx,"cy":cy})
print(f"Red plan primitives: {len(prims)}")

def cluster(prims,link):
    n=len(prims); par=list(range(n))
    def f(a):
        while par[a]!=a: par[a]=par[par[a]]; a=par[a]
        return a
    def u(a,b):
        ra,rb=f(a),f(b)
        if ra!=rb: par[ra]=rb
    bins=defaultdict(list)
    for i,p in enumerate(prims):
        bins[(int(p["cx"]//link),int(p["cy"]//link))].append(i)
    for (bx,by),idxs in bins.items():
        for dx in (-1,0,1):
            for dy in (-1,0,1):
                ne=bins.get((bx+dx,by+dy),[])
                for i in idxs:
                    pi=prims[i]
                    for j in ne:
                        if j<=i: continue
                        pj=prims[j]
                        if abs(pi["cx"]-pj["cx"])<=link and abs(pi["cy"]-pj["cy"])<=link:
                            u(i,j)
    g=defaultdict(list)
    for i in range(n): g[f(i)].append(i)
    return list(g.values())

clusters_idx = cluster(prims, LINK_DIST)
print(f"Clusters at link={LINK_DIST}pt: {len(clusters_idx)}")

def rasterise(idxs, tpl_size=TPL_SIZE, stroke=2):
    if not idxs: return None
    xs0=[prims[i]["bbox"][0] for i in idxs]; ys0=[prims[i]["bbox"][1] for i in idxs]
    xs1=[prims[i]["bbox"][2] for i in idxs]; ys1=[prims[i]["bbox"][3] for i in idxs]
    bb=(min(xs0),min(ys0),max(xs1),max(ys1))
    bw,bh=bb[2]-bb[0],bb[3]-bb[1]
    if bw<=0 or bh<=0: return None
    s=8.0; W=max(1,int(round(bw*s))); H=max(1,int(round(bh*s)))
    img=np.zeros((H,W),dtype=np.uint8)
    for i in idxs:
        p=prims[i]
        x0=(p["bbox"][0]-bb[0])*s; y0=(p["bbox"][1]-bb[1])*s
        x1=(p["bbox"][2]-bb[0])*s; y1=(p["bbox"][3]-bb[1])*s
        if (x1-x0)<1 and (y1-y0)<1:
            cv2.circle(img,(int((x0+x1)/2),int((y0+y1)/2)),max(1,stroke//2),255,-1)
        elif (x1-x0)<1:
            cv2.line(img,(int(x0),int(y0)),(int(x0),int(y1)),255,stroke)
        elif (y1-y0)<1:
            cv2.line(img,(int(x0),int(y0)),(int(x1),int(y0)),255,stroke)
        else:
            cv2.rectangle(img,(int(x0),int(y0)),(int(x1),int(y1)),255,stroke)
    long=max(W,H); nW=max(1,int(round(W*tpl_size/long))); nH=max(1,int(round(H*tpl_size/long)))
    re_=cv2.resize(img,(nW,nH),interpolation=cv2.INTER_AREA)
    cnv=np.zeros((tpl_size,tpl_size),dtype=np.uint8)
    cnv[(tpl_size-nH)//2:(tpl_size-nH)//2+nH,(tpl_size-nW)//2:(tpl_size-nW)//2+nW]=re_
    return cnv.astype(np.float32)/255.0,bb,(bw,bh)

def ncc(a,b):
    a=a-a.mean(); b=b-b.mean()
    na=np.linalg.norm(a); nb=np.linalg.norm(b)
    if na<1e-6 or nb<1e-6: return 0.0
    return float(np.sum(a*b)/(na*nb))

def ar_in_window(ar, lo, hi):
    """AR matches if it fits the canonical window OR its rotated inverse."""
    if lo <= ar <= hi:
        return True
    # rotated 90deg => ar' = 1/ar
    if ar > 1e-3 and lo <= 1.0/ar <= hi:
        return True
    return False

# Detection
hits = defaultdict(list)
debug_rows = []
for cl in clusters_idx:
    if len(cl) < MIN_PARTS: continue
    out = rasterise(cl)
    if out is None: continue
    img, bb, (bw, bh) = out
    if max(bw, bh) < 5.0: continue
    ar = bh/max(bw, 0.01)

    scores = {}     # lab -> best NCC across 4 rotations
    rot_used = {}
    for lab, rots in TEMPLATES.items():
        lo, hi = AR_INFO[lab]
        if not ar_in_window(ar, lo, hi):
            scores[lab] = -1.0; rot_used[lab] = -1
            continue
        best = -2.0; best_r = 0
        for ri, tpl in enumerate(rots):
            v = ncc(img, tpl)
            if v > best:
                best = v; best_r = ri
        scores[lab] = best
        rot_used[lab] = best_r

    # Winner-takes-all + margin check
    sorted_lab = sorted(scores.items(), key=lambda kv: -kv[1])
    best_lab, best_sc = sorted_lab[0]
    second_sc = sorted_lab[1][1] if len(sorted_lab)>1 else -1.0
    margin = best_sc - second_sc
    if best_sc >= NCC_MIN and margin >= MARGIN_MIN:
        hits[best_lab].append({
            "bbox": bb, "w": round(bw,2), "h": round(bh,2),
            "n": len(cl), "ar": round(ar,2),
            "ncc": round(best_sc,3), "margin": round(margin,3),
            "rot": rot_used[best_lab]*90,
        })
        debug_rows.append((best_lab, bb, len(cl), ar, best_sc, margin))

print()
print("=== Detection v2 (shape + 4 rotations) ===")
total_ok=0
for lab in ("5АЭ","6АЭ","7АЭ"):
    n = len(hits.get(lab,[]))
    exp = DXF_GT[lab]
    diff = n-exp
    mark = "OK" if diff==0 else f"{diff:+d}"
    if diff==0: total_ok+=1
    print(f"  {lab}: detected {n:>3d}  expected {exp}  ({mark})")
print(f"Exact matches: {total_ok}/3")

# JSON
out_obj = {lab: {"detected":len(hits.get(lab,[])), "expected":DXF_GT[lab],
                 "hits":hits.get(lab,[])} for lab in ("5АЭ","6АЭ","7АЭ")}
out_obj["_thresholds"]={"NCC_MIN":NCC_MIN,"MARGIN_MIN":MARGIN_MIN,"MIN_PARTS":MIN_PARTS,"LINK_DIST":LINK_DIST}
(OUT/"results.json").write_text(json.dumps(out_obj,ensure_ascii=False,indent=2),encoding="utf-8")

# overlay
zoom = DPI_OVER/72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom,zoom),alpha=False)
img_np = np.frombuffer(pix.samples,dtype=np.uint8).reshape(pix.height,pix.width,pix.n)
if pix.n==4: img_np=cv2.cvtColor(img_np,cv2.COLOR_RGBA2RGB)
canvas = cv2.cvtColor(img_np,cv2.COLOR_RGB2BGR)
cv2.rectangle(canvas,(int(LX0*zoom),int(LY0*zoom)),(int(LX1*zoom),int(LY1*zoom)),(160,160,160),1)
for lab,hs in hits.items():
    col = PALETTE.get(lab,(0,200,0)); asc = ASCII_MAP.get(lab,lab)
    for h in hs:
        bb=h["bbox"]
        x0=int(bb[0]*zoom); y0=int(bb[1]*zoom)
        x1=int(bb[2]*zoom); y1=int(bb[3]*zoom)
        cv2.rectangle(canvas,(x0-3,y0-3),(x1+3,y1+3),col,2)
        cv2.putText(canvas,f"{asc} {h['ncc']:.2f} r{h['rot']}",(x0,y0-4),
                    cv2.FONT_HERSHEY_SIMPLEX,0.40,col,1,cv2.LINE_AA)
cv2.imwrite(str(OUT/"overlay.png"),canvas)
print(f"\nSaved: {OUT/'overlay.png'}  ({canvas.shape[1]}x{canvas.shape[0]})")

doc.close()
