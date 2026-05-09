"""
Per-PDF template library for shape-based icon detection.

Builds and caches bitmap templates extracted directly from each PDF's legend.
Each template is keyed by (pdf_stem, symbol).

Public API:
    extract_template_from_legend(pdf_path, symbol, legend) -> TemplateRecord | None
    build_library(pdf_paths, symbols) -> dict
    load_library(lib_dir) -> dict
    get_template(pdf_stem, symbol, lib) -> TemplateRecord | None
"""
from __future__ import annotations
import io, sys, json
from collections import defaultdict
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import fitz, numpy as np, cv2

if sys.platform == "win32" and not isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent
DEFAULT_LIB_DIR = ROOT / "_template_lib"

# extraction params
TPL_SIZE = 64
LINK_DIST = 5.0
ICON_ROI_WIDTH = 30.0      # pt; width of the leftmost strip of legend row that holds the icon
N_PRIMS_MIN = 1            # minimum prims to qualify as an icon
N_PRIMS_MAX = 600


@dataclass
class TemplateRecord:
    pdf_stem: str
    symbol: str
    tpl_path: str       # relative to lib root
    n_parts: int
    bw: float
    bh: float
    ar: float
    src_bbox: tuple     # (x0,y0,x1,y1) of the icon ROI we used
    method: str         # "red_clustered" / "any_color_clustered" / "fallback_005"


# ---------- low-level helpers (mirror the v3 detector) ----------

def _is_red(c) -> bool:
    if c is None: return False
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        return r > 0.6 and g < 0.4 and b < 0.4
    return False

def _dbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]; xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
        elif it[0] in ("l", "m", "c"):
            for p in it[1:]:
                if hasattr(p, "x"): xs.append(p.x); ys.append(p.y)
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None

def _cluster(prims, link=LINK_DIST):
    n = len(prims); par = list(range(n))
    def f(a):
        while par[a] != a: par[a] = par[par[a]]; a = par[a]
        return a
    def u(a, b):
        ra, rb = f(a), f(b)
        if ra != rb: par[ra] = rb
    bins = defaultdict(list)
    for i, p in enumerate(prims):
        bins[(int(p["cx"]//link), int(p["cy"]//link))].append(i)
    for (bx, by), idxs in bins.items():
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                ne = bins.get((bx+dx, by+dy), [])
                for i in idxs:
                    pi = prims[i]
                    for j in ne:
                        if j <= i: continue
                        pj = prims[j]
                        if abs(pi["cx"]-pj["cx"]) <= link and abs(pi["cy"]-pj["cy"]) <= link:
                            u(i, j)
    g = defaultdict(list)
    for i in range(n): g[f(i)].append(i)
    return list(g.values())

def _rasterise(idxs, prims, tpl_size=TPL_SIZE, stroke=2):
    if not idxs: return None
    xs0 = [prims[i]["bbox"][0] for i in idxs]; ys0 = [prims[i]["bbox"][1] for i in idxs]
    xs1 = [prims[i]["bbox"][2] for i in idxs]; ys1 = [prims[i]["bbox"][3] for i in idxs]
    bb = (min(xs0), min(ys0), max(xs1), max(ys1))
    bw, bh = bb[2]-bb[0], bb[3]-bb[1]
    if bw <= 0 or bh <= 0: return None
    s = 8.0; W = max(1, int(round(bw*s))); H = max(1, int(round(bh*s)))
    img = np.zeros((H, W), dtype=np.uint8)
    for i in idxs:
        p = prims[i]
        x0 = (p["bbox"][0]-bb[0])*s; y0 = (p["bbox"][1]-bb[1])*s
        x1 = (p["bbox"][2]-bb[0])*s; y1 = (p["bbox"][3]-bb[1])*s
        if (x1-x0) < 1 and (y1-y0) < 1:
            cv2.circle(img, (int((x0+x1)/2), int((y0+y1)/2)), max(1, stroke//2), 255, -1)
        elif (x1-x0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x0), int(y1)), 255, stroke)
        elif (y1-y0) < 1:
            cv2.line(img, (int(x0), int(y0)), (int(x1), int(y0)), 255, stroke)
        else:
            cv2.rectangle(img, (int(x0), int(y0)), (int(x1), int(y1)), 255, stroke)
    long = max(W, H)
    nW = max(1, int(round(W*tpl_size/long))); nH = max(1, int(round(H*tpl_size/long)))
    re_ = cv2.resize(img, (nW, nH), interpolation=cv2.INTER_AREA)
    cnv = np.zeros((tpl_size, tpl_size), dtype=np.uint8)
    cnv[(tpl_size-nH)//2:(tpl_size-nH)//2+nH, (tpl_size-nW)//2:(tpl_size-nW)//2+nW] = re_
    return cnv.astype(np.float32)/255.0, bb, (bw, bh)


# ---------- public API ----------

def extract_template_from_legend(pdf_path: Path, symbol: str, legend) -> Optional[dict]:
    """
    Locate the row for `symbol` in `legend.items`, scan multiple ROI strategies
    (right-of-symbol, left-of-symbol-in-table-cell), cluster small primitives of any
    relevant color, return the densest cluster as a 64x64 template.

    Strategies tried in order:
      A) RIGHT_OF_BBOX: leftmost ICON_ROI_WIDTH points of row.bbox
         (works for legends where bbox = full row including icon, e.g. ГПК).
      B) LEFT_OF_BBOX:  strip [legend_bbox.x0 .. row.bbox.x0] at the row's y
         (works for legends where bbox = description column only, icon in a separate
          cell to the left, e.g. КПП).

    For each strategy we try (red, blue, any) color filters in turn.

    Returns dict with keys: tpl, n_parts, bw, bh, ar, src_bbox, method
    or None if extraction failed.
    """
    rows = [it for it in legend.items if it.symbol == symbol]
    if not rows:
        return None
    row = rows[0]
    rx0, ry0, rx1, ry1 = row.bbox
    LX0, LY0, LX1, LY1 = legend.legend_bbox
    cy_row = (ry0 + ry1) / 2

    rois = [
        ("right_of_bbox", rx0,            rx0 + ICON_ROI_WIDTH, ry0 - 2.0, ry1 + 2.0),
        ("left_of_bbox",  max(LX0, rx0 - 100.0), rx0 - 1.0,     cy_row - 25.0, cy_row + 25.0),
    ]

    doc = fitz.open(str(pdf_path))
    try:
        mp = doc[legend.page_index]
        for roi_name, ix0, ix1, iy0, iy1 in rois:
            if ix1 <= ix0: continue
            for color_mode in ("red", "blue", "any"):
                prims = []
                for d in mp.get_drawings():
                    bb = _dbox(d)
                    if not bb: continue
                    cx, cy = (bb[0]+bb[2])/2, (bb[1]+bb[3])/2
                    if not (ix0 <= cx <= ix1 and iy0 <= cy <= iy1): continue
                    if color_mode != "any":
                        c = d.get("fill") or d.get("color")
                        if color_mode == "red" and not _is_red(c): continue
                        if color_mode == "blue" and not _is_blue(c): continue
                    w, h = bb[2]-bb[0], bb[3]-bb[1]
                    if max(w, h) > 30.0: continue
                    prims.append({"bbox": bb, "cx": cx, "cy": cy})
                if not prims: continue
                cls = _cluster(prims)
                if not cls: continue
                cls.sort(key=lambda x: -len(x))
                best = cls[0]
                if not (N_PRIMS_MIN <= len(best) <= N_PRIMS_MAX): continue
                out = _rasterise(best, prims)
                if out is None: continue
                tpl, bb, (bw, bh) = out
                if max(bw, bh) < 3.0: continue
                return {
                    "tpl": tpl,
                    "n_parts": len(best),
                    "bw": float(bw), "bh": float(bh),
                    "ar": float(bh/max(bw, 0.01)),
                    "src_bbox": tuple(round(v, 2) for v in bb),
                    "method": f"{roi_name}/{color_mode}",
                }
        return None
    finally:
        doc.close()


def _is_blue(c) -> bool:
    if c is None: return False
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        return b > 0.6 and g < 0.4 and r < 0.4
    return False


def build_library(pdf_paths: list[Path], symbols: list[str], lib_dir: Path = DEFAULT_LIB_DIR,
                  parse_legend_fn=None, verbose: bool = True) -> dict:
    """
    Build per-PDF templates and write them to lib_dir. Returns the in-memory index.
    """
    lib_dir.mkdir(exist_ok=True)
    if parse_legend_fn is None:
        sys.path.insert(0, str(ROOT))
        from pdf_legend_parser import parse_legend as parse_legend_fn

    index = {}
    for pdf in pdf_paths:
        stem = pdf.stem.split("-")[0].strip()    # "005-Планы..." -> "005"
        if verbose:
            print(f"\n[{stem}] {pdf.name}")
        try:
            legend = parse_legend_fn(str(pdf))
        except Exception as e:
            print(f"  legend parse failed: {e}")
            continue

        present = sorted({it.symbol for it in legend.items if it.symbol in symbols})
        if verbose:
            print(f"  legend symbols of interest: {present}")
        sub_dir = lib_dir / stem
        sub_dir.mkdir(exist_ok=True)
        per_pdf = {}
        for sym in present:
            res = extract_template_from_legend(pdf, sym, legend)
            if res is None:
                if verbose:
                    print(f"  {sym}: EXTRACTION FAILED")
                continue
            safe_sym = sym.replace("/", "_")
            tpl_file = sub_dir / f"{safe_sym}.npy"
            png_file = sub_dir / f"{safe_sym}.png"
            np.save(str(tpl_file), res["tpl"])
            cv2.imwrite(str(png_file), (res["tpl"]*255).astype(np.uint8))
            rec = {
                "pdf_stem": stem,
                "symbol": sym,
                "tpl_path": str(tpl_file.relative_to(lib_dir)).replace("\\", "/"),
                "n_parts": res["n_parts"],
                "bw": round(res["bw"], 2),
                "bh": round(res["bh"], 2),
                "ar": round(res["ar"], 3),
                "src_bbox": list(res["src_bbox"]),
                "method": res["method"],
            }
            per_pdf[sym] = rec
            if verbose:
                print(f"  {sym}: n={rec['n_parts']:>3} W={rec['bw']:>5.1f} H={rec['bh']:>5.1f} "
                      f"AR={rec['ar']:>4.2f} via {rec['method']}")
        index[stem] = per_pdf
    (lib_dir / "index.json").write_text(json.dumps(index, ensure_ascii=False, indent=2),
                                        encoding="utf-8")
    if verbose:
        print(f"\nLibrary saved: {lib_dir/'index.json'} "
              f"({sum(len(v) for v in index.values())} templates across {len(index)} PDFs)")
    return index


def load_library(lib_dir: Path = DEFAULT_LIB_DIR) -> dict:
    """Load the library index and the actual template arrays into memory."""
    idx_file = lib_dir / "index.json"
    if not idx_file.exists():
        return {}
    index = json.loads(idx_file.read_text(encoding="utf-8"))
    for stem, syms in index.items():
        for sym, rec in syms.items():
            tpl_arr = np.load(str(lib_dir / rec["tpl_path"]))
            rec["tpl"] = tpl_arr.astype(np.float32)
    return index


def get_template(pdf_stem: str, symbol: str, lib: dict) -> Optional[dict]:
    """Lookup with no fallback."""
    return lib.get(pdf_stem, {}).get(symbol)


def get_template_with_fallback(pdf_stem: str, symbol: str, lib: dict,
                                fallback_order: list[str] = None) -> Optional[dict]:
    """Lookup; if missing, try other PDFs in `fallback_order`."""
    rec = get_template(pdf_stem, symbol, lib)
    if rec is not None: return rec
    if fallback_order:
        for fb_stem in fallback_order:
            rec = get_template(fb_stem, symbol, lib)
            if rec is not None: return rec
    # last resort: any PDF that has this symbol
    for stem, syms in lib.items():
        if symbol in syms:
            return syms[symbol]
    return None
