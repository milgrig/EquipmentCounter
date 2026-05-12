"""
Strict legend-aware sweep:
  * For each PDF, parse legend → set of present symbols (subset of {5АЭ, 6АЭ, 7АЭ}).
  * Run shape detector with templates_curated_v2.
  * Reject any hit whose target class is NOT in this legend (no per-page templates,
    just gating).
  * Ambiguous square (5АЭ vs 7АЭ): if exactly one of them is in legend → route there;
    if both → keep raw NCC winner; if neither → drop.
  * 5АЭ_h → 5АЭ only if 5АЭ is in legend.
"""
from __future__ import annotations
import io, os, sys, json, re, subprocess
from collections import defaultdict
from pathlib import Path

import fitz, numpy as np, cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF_DIR = ROOT / r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF"
DXF_DIR = ROOT / r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\_converted_dxf\01_DWG"
TPL_DIR = ROOT / "_shape_005_out"
MANIFEST = TPL_DIR / "templates_curated_v2.json"
OUT = ROOT / "_sweep_legend_strict_out"
OUT.mkdir(exist_ok=True)

PAIRS = [
    ("005-Планы освещения-отм. 0.000.pdf",          "005 - План  освещения на отм- 0-000.dxf"),
    ("006-Планы освещения-отм. +4.200.pdf",         "006 - План освещения на отм- +4-200.dxf"),
    ("007-Планы освещения-отм. +7.800 +9.000.pdf",  "007 - План освещения на отм- +7-800, +9-000.dxf"),
    ("008-Планы освещения-отм. +13.800.pdf",        "008 - План освещения на отм- +13-800.dxf"),
    ("009-Планы освещения-отм. +18.600.pdf",        "009 - План освещения на отм- +18-600.dxf"),
    ("010-Планы освещения-отм. +23.400.pdf",        "010 - План освещения на отм- +23-400.dxf"),
    ("011-Планы освещения-отм. +28.200.pdf",        "011 - План освещения на отм- +28-200.dxf"),
]
CLASSES = ["5АЭ", "6АЭ", "7АЭ"]
AMBIG = {"5АЭ", "7АЭ"}

TPL_SIZE=64; LINK_DIST=5.0; MAX_PRIM=60.0
NCC_MIN=0.25; MARGIN_MIN=0.05; MIN_PARTS=18

GT_RE = re.compile(r"^\s*\d+\s+(\S+)\s+.+?\s+(\d+)(?:\s+(\d+))?(?:\s+(\d+))?\s*$")
def collect_gt(dxf_path: Path) -> dict[str, int]:
    cmd=[sys.executable,str(ROOT/"equipment_counter.py"),str(dxf_path)]
    env=os.environ.copy(); env["PYTHONIOENCODING"]="utf-8"
    try:
        r=subprocess.run(cmd,capture_output=True,text=True,encoding="utf-8",
                         errors="replace",timeout=300,env=env)
    except subprocess.TimeoutExpired:
        return {c:-1 for c in CLASSES}
    counts={c:0 for c in CLASSES}
    for line in r.stdout.splitlines():
        m=GT_RE.match(line)
        if not m: continue
        sym=m.group(1).strip()
        if sym not in counts: continue
        nums=[int(x) for x in m.groups()[1:] if x is not None]
        if nums: counts[sym]+=nums[-1]
    return counts

sys.path.insert(0, str(ROOT))
from pdf_legend_parser import parse_legend  # noqa: E402

manifest=json.loads(MANIFEST.read_text(encoding="utf-8"))
TEMPLATES,INFO={},{}
for lab in manifest:
    npy=TPL_DIR/f"tpl_curated_{lab}.npy"
    if not npy.exists(): continue
    base=np.load(str(npy)).astype(np.float32)
    TEMPLATES[lab]=[base,np.rot90(base,1),np.rot90(base,2),np.rot90(base,3)]
    INFO[lab]=manifest[lab]
print(f"Templates: {list(TEMPLATES.keys())}")

def color_red(c):
    if c is None: return False
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b=c[0],c[1],c[2]; return r>0.6 and g<0.4 and b<0.4
    return False
def dbox(d):
    xs,ys=[],[]
    for it in d.get("items",[]):
        if it[0]=="re":
            r=it[1]; xs+=[r.x0,r.x1]; ys+=[r.y0,r.y1]
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None
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
def rasterise(idxs, prims, tpl_size=TPL_SIZE, stroke=2):
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
def ar_window(ar,info):
    return (info["ar_lo"]<=ar<=info["ar_hi"]) or (info["ar_lo_rotated"]<=ar<=info["ar_hi_rotated"])

def detect_pdf(pdf_path: Path):
    try:
        legend=parse_legend(str(pdf_path))
    except Exception as e:
        print(f"  legend parse failed: {e}")
        return {c:-1 for c in CLASSES}, set()
    LX0,LY0,LX1,LY1=legend.legend_bbox
    page_idx=legend.page_index
    legend_syms={it.symbol for it in legend.items if it.symbol in CLASSES}

    doc=fitz.open(str(pdf_path)); mp=doc[page_idx]
    prims=[]
    for d in mp.get_drawings():
        bb=dbox(d)
        if not bb: continue
        w,h=bb[2]-bb[0],bb[3]-bb[1]
        if max(w,h)>MAX_PRIM: continue
        c=d.get("fill") or d.get("color")
        if not color_red(c): continue
        cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
        if LX0-2<=cx<=LX1+2 and LY0-2<=cy<=LY1+2: continue
        prims.append({"bbox":bb,"cx":cx,"cy":cy})

    cls=cluster(prims,LINK_DIST)
    counts={c:0 for c in CLASSES}
    present_ambig = AMBIG & legend_syms
    ambig_redirect = next(iter(present_ambig)) if len(present_ambig)==1 else None

    for cl in cls:
        if len(cl)<MIN_PARTS: continue
        out_r=rasterise(cl,prims)
        if out_r is None: continue
        img,bb,(bw,bh)=out_r
        if max(bw,bh)<5.0: continue
        ar=bh/max(bw,0.01)
        scores={}
        for lab,rots in TEMPLATES.items():
            info=INFO[lab]
            if not ar_window(ar,info):
                scores[lab]=-1.0; continue
            if not (info["n_min"]<=len(cl)<=info["n_max"]):
                scores[lab]=-1.0; continue
            best=-2.0
            for tpl in rots:
                v=ncc(img,tpl)
                if v>best: best=v
            scores[lab]=best
        sorted_lab=sorted(scores.items(),key=lambda kv:-kv[1])
        best_lab,best_sc=sorted_lab[0]
        second_sc=sorted_lab[1][1] if len(sorted_lab)>1 else -1.0
        if best_sc<NCC_MIN or (best_sc-second_sc)<MARGIN_MIN: continue

        # alias 5АЭ_h -> 5АЭ
        target = INFO[best_lab].get("_alias_for", best_lab)

        # ambiguous-square routing
        if target in AMBIG:
            if ambig_redirect is not None:
                target = ambig_redirect
            elif not present_ambig:
                continue   # neither in legend → drop

        # STRICT: target must be in legend, otherwise drop
        if target not in legend_syms:
            continue

        if target in counts:
            counts[target]+=1
    doc.close()
    return counts, legend_syms

results=[]
print()
print(f"{'File':<5} | {'Legend':<14} | {'Class':<5} | {'GT':>4} | {'DET':>4} | {'diff':>5}")
print("-"*55)
for pdf_name,dxf_name in PAIRS:
    pdf_path=PDF_DIR/pdf_name; dxf_path=DXF_DIR/dxf_name
    tag=pdf_name[:3]
    if not pdf_path.exists() or not dxf_path.exists():
        print(f"  MISSING: {pdf_name}"); continue
    print(f"\n[{tag}] GT…",end=" ",flush=True)
    gt=collect_gt(dxf_path); print(f"GT={gt}",end=" ; ",flush=True)
    print("DET…",end=" ",flush=True)
    det,leg=detect_pdf(pdf_path)
    leg_str=",".join(sorted(leg)) or "(none)"
    print(f"DET={det}  legend={leg_str}")
    for c in CLASSES:
        g,d=gt[c],det[c]
        diff=(d-g) if (g>=0 and d>=0) else None
        results.append({"file":tag,"legend":sorted(leg),"class":c,"gt":g,"det":d,"diff":diff})
        print(f"  {tag} | {leg_str:<14} | {c:<5} | {g:>4} | {d:>4} | "
              f"{('—' if diff is None else f'{diff:+d}'):>5}")

agg={c:{"tp":0,"gt_tot":0,"det_tot":0} for c in CLASSES}
for r in results:
    if r["gt"]<0 or r["det"]<0: continue
    c=r["class"]
    agg[c]["gt_tot"]+=r["gt"]; agg[c]["det_tot"]+=r["det"]
    agg[c]["tp"]+=min(r["gt"],r["det"])
print("\n=== Aggregate (legend-strict) ===")
for c in CLASSES:
    g,d,tp=agg[c]["gt_tot"],agg[c]["det_tot"],agg[c]["tp"]
    P=tp/d if d else 0.0; R=tp/g if g else 0.0
    print(f"  {c}: GT={g:>3} DET={d:>3}  P~{P:.2f}  R~{R:.2f}")

(OUT/"sweep_legend_strict.json").write_text(
    json.dumps({"per_pair":results,"aggregate":agg},ensure_ascii=False,indent=2),
    encoding="utf-8")
print(f"\nSaved: {OUT/'sweep_legend_strict.json'}")
