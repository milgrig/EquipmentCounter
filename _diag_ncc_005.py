"""Show NCC scores per template for the top red clusters."""
from __future__ import annotations
import io, sys, json
from collections import defaultdict
from pathlib import Path
import fitz, numpy as np, cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
PDF=Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
TPL=Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_shape_005_out")
TPL_SIZE=64; LINK=5.0; MAX=60.0
sys.path.insert(0,str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend=parse_legend(str(PDF)); LX0,LY0,LX1,LY1=legend.legend_bbox
doc=fitz.open(str(PDF)); mp=doc[legend.page_index]
def cc(c):
    if not c: return None
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b=c[0],c[1],c[2]
        if r>0.6 and g<0.4 and b<0.4: return "red"
    return None
def dbox(d):
    xs,ys=[],[]
    for it in d.get("items",[]):
        if it[0]=="re":
            r=it[1]; xs+=[r.x0,r.x1]; ys+=[r.y0,r.y1]
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None
prims=[]
for d in mp.get_drawings():
    bb=dbox(d)
    if not bb: continue
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if max(w,h)>MAX: continue
    if cc(d.get("fill")) or cc(d.get("color")):
        cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
        if LX0-2<=cx<=LX1+2 and LY0-2<=cy<=LY1+2: continue
        prims.append({"bbox":bb,"cx":cx,"cy":cy})
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
cls=cluster(prims,LINK)
def rast(idxs,t=TPL_SIZE,stk=2):
    xs0=[prims[i]["bbox"][0] for i in idxs]; ys0=[prims[i]["bbox"][1] for i in idxs]
    xs1=[prims[i]["bbox"][2] for i in idxs]; ys1=[prims[i]["bbox"][3] for i in idxs]
    bb=(min(xs0),min(ys0),max(xs1),max(ys1))
    bw,bh=bb[2]-bb[0],bb[3]-bb[1]
    if bw<=0 or bh<=0: return None,bb,(bw,bh)
    s=8.0; W=max(1,int(round(bw*s))); H=max(1,int(round(bh*s)))
    img=np.zeros((H,W),dtype=np.uint8)
    for i in idxs:
        p=prims[i]
        x0=(p["bbox"][0]-bb[0])*s; y0=(p["bbox"][1]-bb[1])*s
        x1=(p["bbox"][2]-bb[0])*s; y1=(p["bbox"][3]-bb[1])*s
        if (x1-x0)<1 and (y1-y0)<1:
            cv2.circle(img,(int((x0+x1)/2),int((y0+y1)/2)),max(1,stk//2),255,-1)
        elif (x1-x0)<1:
            cv2.line(img,(int(x0),int(y0)),(int(x0),int(y1)),255,stk)
        elif (y1-y0)<1:
            cv2.line(img,(int(x0),int(y0)),(int(x1),int(y0)),255,stk)
        else:
            cv2.rectangle(img,(int(x0),int(y0)),(int(x1),int(y1)),255,stk)
    long=max(W,H); nW=max(1,int(round(W*t/long))); nH=max(1,int(round(H*t/long)))
    re_=cv2.resize(img,(nW,nH),interpolation=cv2.INTER_AREA)
    cnv=np.zeros((t,t),dtype=np.uint8)
    cnv[(t-nH)//2:(t-nH)//2+nH,(t-nW)//2:(t-nW)//2+nW]=re_
    return cnv.astype(np.float32)/255.0,bb,(bw,bh)
def ncc(a,b):
    a=a-a.mean(); b=b-b.mean()
    na=np.linalg.norm(a); nb=np.linalg.norm(b)
    if na<1e-6 or nb<1e-6: return 0.0
    return float(np.sum(a*b)/(na*nb))
TPLS={lab: np.load(str(TPL/f"tpl_curated_{lab}.npy")).astype(np.float32) for lab in ("5АЭ","6АЭ","7АЭ")}
big=[c for c in cls if len(c)>=20]
print(f"Clusters with n>=20: {len(big)}")
for c in sorted(big,key=lambda c:-len(c))[:25]:
    out=rast(c)
    if out[0] is None: continue
    img,bb,(bw,bh)=out
    ar=bh/max(bw,0.01)
    s={lab: round(ncc(img,t),3) for lab,t in TPLS.items()}
    print(f"  n={len(c):>3d}  W={bw:>5.1f} H={bh:>5.1f} AR={ar:>4.2f}  cx={(bb[0]+bb[2])/2:>6.0f} cy={(bb[1]+bb[3])/2:>6.0f}  ncc={s}")
doc.close()
