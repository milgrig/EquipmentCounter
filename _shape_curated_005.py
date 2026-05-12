"""
Curated shape templates: for each text label gather candidate clusters,
then pick those whose aspect ratio matches the EXPECTED visual signature.
This avoids contamination when a label sits closer to a different icon.

Expected visual aspect ratio (AR = h/w):
   5АЭ   ≈ 0.85..1.20  (square-ish, large)
   6АЭ   ≈ 1.80..2.40  (tall, narrow)
   7АЭ   ≈ 0.70..1.20  (compact)
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

TPL_SIZE = 64
SEARCH_RADIUS = 30.0      # search around label for ALL nearby cluster candidates
LINK_DIST = 4.0
GROW_LIMIT_PT = 40.0
MAX_PRIM = 60.0

# Per-class AR window (h / w) and minimum n_parts.
EXPECTED = {
    "5АЭ": {"ar_lo": 0.75, "ar_hi": 1.40, "n_min": 80},
    "6АЭ": {"ar_lo": 1.60, "ar_hi": 3.20, "n_min": 25},
    "7АЭ": {"ar_lo": 0.70, "ar_hi": 1.30, "n_min": 60},
}


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
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None

sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0,LY0,LX1,LY1 = legend.legend_bbox
doc = fitz.open(str(PDF)); mp = doc[legend.page_index]

prims=[]
for d in mp.get_drawings():
    bb = dbox(d)
    if not bb: continue
    w,h = bb[2]-bb[0], bb[3]-bb[1]
    if max(w,h) > MAX_PRIM: continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col != "red": continue
    cx,cy = (bb[0]+bb[2])/2,(bb[1]+bb[3])/2
    if LX0-2<=cx<=LX1+2 and LY0-2<=cy<=LY1+2: continue
    prims.append({"bbox":bb,"cx":cx,"cy":cy,"color":col})

grid=defaultdict(list)
for i,p in enumerate(prims):
    grid[(int(p["cx"]//LINK_DIST),int(p["cy"]//LINK_DIST))].append(i)

def near(cx,cy,r):
    out=[]
    for bx in range(int((cx-r)//LINK_DIST), int((cx+r)//LINK_DIST)+1):
        for by in range(int((cy-r)//LINK_DIST), int((cy+r)//LINK_DIST)+1):
            for idx in grid.get((bx,by),[]):
                p=prims[idx]
                if abs(p["cx"]-cx)<=r and abs(p["cy"]-cy)<=r:
                    out.append(idx)
    return out

def grow_from(seed, exclude_box=None):
    if not seed: return []
    cl=set(seed); fr=list(seed); bb=_bb(cl)
    while fr:
        nf=[]
        for i in fr:
            p=prims[i]
            for j in near(p["cx"],p["cy"],LINK_DIST):
                if j in cl: continue
                pj=prims[j]
                if exclude_box:
                    ex0,ey0,ex1,ey1 = exclude_box
                    if ex0-0.5<=pj["cx"]<=ex1+0.5 and ey0-0.5<=pj["cy"]<=ey1+0.5: continue
                nb=(min(bb[0],pj["bbox"][0]),min(bb[1],pj["bbox"][1]),
                    max(bb[2],pj["bbox"][2]),max(bb[3],pj["bbox"][3]))
                if (nb[2]-nb[0])>GROW_LIMIT_PT or (nb[3]-nb[1])>GROW_LIMIT_PT: continue
                cl.add(j); nf.append(j); bb=nb
        fr=nf
    return list(cl)

def _bb(idxs):
    xs0=[prims[i]["bbox"][0] for i in idxs]; ys0=[prims[i]["bbox"][1] for i in idxs]
    xs1=[prims[i]["bbox"][2] for i in idxs]; ys1=[prims[i]["bbox"][3] for i in idxs]
    return (min(xs0),min(ys0),max(xs1),max(ys1))

def rasterise_to(idxs, tpl_size=TPL_SIZE, stroke=2):
    bb=_bb(idxs); bw,bh=bb[2]-bb[0],bb[3]-bb[1]
    if bw<=0 or bh<=0: return np.zeros((tpl_size,tpl_size),dtype=np.float32),bb,(bw,bh)
    scale=8.0
    W=max(1,int(round(bw*scale))); H=max(1,int(round(bh*scale)))
    img=np.zeros((H,W),dtype=np.uint8)
    for i in idxs:
        p=prims[i]
        x0=(p["bbox"][0]-bb[0])*scale; y0=(p["bbox"][1]-bb[1])*scale
        x1=(p["bbox"][2]-bb[0])*scale; y1=(p["bbox"][3]-bb[1])*scale
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
    cnv[(tpl_size-nH)//2:(tpl_size-nH)//2+nH, (tpl_size-nW)//2:(tpl_size-nW)//2+nW] = re_
    return cnv.astype(np.float32)/255.0, bb, (bw,bh)

# label texts on the plan
label_pts=defaultdict(list)
with pdfplumber.open(str(PDF)) as pdf:
    page=pdf.pages[legend.page_index]
    for w in page.extract_words() or []:
        t=(w.get("text") or "").strip()
        if t not in EXPECTED: continue
        cx=(w["x0"]+w["x1"])/2; cy=(w["top"]+w["bottom"])/2
        if LX0-2<=cx<=LX1+2 and LY0-2<=cy<=LY1+2: continue
        label_pts[t].append((cx,cy,w["x0"],w["top"],w["x1"],w["bottom"]))

manifest={}
for label,exp in EXPECTED.items():
    pts=label_pts.get(label,[])
    if not pts:
        print(f"  {label}: no labels"); continue

    safe=re.sub(r"[^\w]","_",label)
    # Each text label may be near several seeds belonging to different icons.
    # Try each seed, grow it, evaluate AR, keep clusters that pass the window.
    candidates_by_inst=[]
    for k,(cx,cy,tx0,ty0,tx1,ty1) in enumerate(pts):
        nearby=near(cx,cy,SEARCH_RADIUS)
        nearby=[i for i in nearby
                if not (tx0-0.5<=prims[i]["cx"]<=tx1+0.5 and ty0-0.5<=prims[i]["cy"]<=ty1+0.5)]
        if not nearby: continue
        # group seeds by connected component first, by trying region grow
        # from each not-yet-visited seed
        seen=set(); comps=[]
        for s in nearby:
            if s in seen: continue
            comp=grow_from([s], exclude_box=(tx0,ty0,tx1,ty1))
            seen.update(comp); comps.append(comp)
        # evaluate each component: AR & n_parts
        for comp in comps:
            bb=_bb(comp); bw,bh=bb[2]-bb[0],bb[3]-bb[1]
            if max(bw,bh)<4.0: continue
            ar=bh/max(bw,0.01)
            if not (exp["ar_lo"]<=ar<=exp["ar_hi"]): continue
            if len(comp)<exp["n_min"]: continue
            # distance label-> cluster centre
            ccx=(bb[0]+bb[2])/2; ccy=(bb[1]+bb[3])/2
            dist=((ccx-cx)**2+(ccy-cy)**2)**0.5
            candidates_by_inst.append((dist, comp, bb, (bw,bh), ar, k))

    if not candidates_by_inst:
        print(f"  {label}: no AR-matching candidates"); continue

    # for each text instance, pick the closest matching component
    best_per_label={}  # k -> (dist, comp, bb, dims, ar)
    for dist,comp,bb,dims,ar,k in candidates_by_inst:
        if k not in best_per_label or dist<best_per_label[k][0]:
            best_per_label[k]=(dist,comp,bb,dims,ar)

    imgs=[]; meta=[]
    for k,(dist,comp,bb,dims,ar) in best_per_label.items():
        img,_,_ = rasterise_to(comp)
        imgs.append(img)
        meta.append({"idx":k,"dist":round(dist,2),"bbox":bb,
                     "n":len(comp),"w":round(dims[0],2),"h":round(dims[1],2),"ar":round(ar,2)})
        cv2.imwrite(str(OUT/f"tpl_curated_{safe}_inst_{k}.png"),(img*255).astype(np.uint8))

    avg=np.mean(np.stack(imgs,axis=0),axis=0).astype(np.float32)
    np.save(str(OUT/f"tpl_curated_{safe}.npy"),avg)
    cv2.imwrite(str(OUT/f"tpl_curated_{safe}.png"),(avg*255).astype(np.uint8))
    manifest[label]={"n_used":len(imgs),
                     "ar_lo":exp["ar_lo"],"ar_hi":exp["ar_hi"],
                     "median_w":round(np.median([m["w"] for m in meta]),2),
                     "median_h":round(np.median([m["h"] for m in meta]),2),
                     "median_n":int(np.median([m["n"] for m in meta])),
                     "median_ar":round(np.median([m["ar"] for m in meta]),2),
                     "instances":meta}
    print(f"  {label}: kept {len(imgs)} instances  W~{manifest[label]['median_w']}pt  "
          f"H~{manifest[label]['median_h']}pt  AR={manifest[label]['median_ar']}  "
          f"n_parts~{manifest[label]['median_n']}")

(OUT/"templates_curated.json").write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding="utf-8")
doc.close()
print(f"\nSaved curated templates to {OUT}")
