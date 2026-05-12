"""Probe v2: extract cables from TABLECONTENT + MTEXT in all DXFs.

TABLECONTENT stores the cable journal as a sequence of cells; each cell has
code=302 with CONTENT marker + value + GRIDFORMAT markers. By reading them
in order we recover rows (length, marking, section) triples.
"""
from __future__ import annotations

import io
import os
import re
import sys
from collections import defaultdict

if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
    except (AttributeError, ValueError):
        pass


_CABLE_BRAND_RE = re.compile(
    r"(ВБШвнг\(А\)-FRLS|ВБШвнг\(А\)-LS|ППГнг\(А\)-FRHF|ППГнг\(А\)-HF|"
    r"ВБШвнг[-\s]?[А-Яа-яA-Za-z]*|ППГнг[-\s]?[А-Яа-яA-Za-z]*|"
    r"ВВГнг[-\s]?[А-Яа-яA-Za-z]*)",
    re.IGNORECASE,
)
_SEC_RE = re.compile(r"(\d+)[xх×](\d+(?:[.,]\d+)?)")
_LEN_RE = re.compile(r"L\s*[=\-]\s*(\d+)")
_MTEXT_CTRL = re.compile(r"\\[pP]|\\[A-Za-z]\w*|[{}]|\\\\")


def _clean(s: str) -> str:
    s = _MTEXT_CTRL.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()


def iter_entities(dxf_path: str):
    """Yield (entity_type, list_of_(code,value)) pairs."""
    try:
        fh = open(dxf_path, "r", encoding="utf-8", errors="replace")
    except OSError:
        return
    with fh:
        lines = fh.readlines()
    i = 0
    cur_type = None
    cur_pairs = []
    while i < len(lines) - 1:
        code = lines[i].strip()
        val = lines[i + 1].rstrip("\n\r")
        if code == "0":
            if cur_type is not None:
                yield cur_type, cur_pairs
            cur_type = val.strip()
            cur_pairs = []
        else:
            cur_pairs.append((code, val))
        i += 2
    if cur_type is not None:
        yield cur_type, cur_pairs


def parse_tablecontent(pairs):
    """Return list of cable dicts extracted from a TABLECONTENT cell stream.

    Strategy: walk 302-coded values in order. When we see 'L=NNNм' treat it
    as length; right before/after it should be the marking cell (like
    'ВБШвнг(А)-FRLS-3х') and the section cell ('1.5'/'2.5'). The journal
    has a consistent column order, so for each length we look at the
    previous CONTENT value with a marking (brand) and nearest next value
    that matches a section number.
    """
    cells = []  # list of content strings (302 values right after 'CONTENT')
    prev_is_content = False
    for code, val in pairs:
        if code != "302":
            continue
        v = val.strip()
        if v == "CONTENT":
            prev_is_content = True
            continue
        if v in ("GRIDFORMAT", ""):
            prev_is_content = False
            continue
        if prev_is_content:
            cells.append(_clean(v))
        prev_is_content = False

    # Now cells is the flat sequence of data cells in row-major order.
    # Column layout for the cable journal (based on observed content):
    # col0: row number | col1: panel/feeder | col2: marking | col3: section
    # col4: length | col5: voltage drop | col6: laying | col7: notes
    # But columns repeat per row, and row header cells contain text like
    # 'Марка;сечение проводника' / 'Длина участка, м'.
    # A more robust heuristic: for each cell that contains a brand, find
    # next cell with section-number OR combine with following length cell.

    results = []
    n = len(cells)
    for idx, cell in enumerate(cells):
        bm = _CABLE_BRAND_RE.search(cell)
        if not bm:
            continue
        brand_raw = bm.group(1).rstrip("- ")
        # Section may be embedded in this cell (e.g. 'ВБШвнг(А)-FRLS 3х1,5')
        sec_m = _SEC_RE.search(cell)
        section = None
        if sec_m:
            section = f"{sec_m.group(1)}х{sec_m.group(2)}"
        else:
            # Look at next 2 cells for section like '1.5' / '2.5' / '1,5'
            for look in range(idx + 1, min(idx + 4, n)):
                nxt = cells[look]
                snum = re.fullmatch(r"(\d+[.,]\d+|\d+)", nxt.strip())
                if snum:
                    section = f"3х{snum.group(1).replace('.', ',')}"
                    break

        # Find nearest length cell within ±8 cells
        length = None
        for look in range(max(0, idx - 8), min(n, idx + 8)):
            lm = _LEN_RE.search(cells[look])
            if lm:
                length = int(lm.group(1))
                break
        if length is None:
            continue
        if not section:
            section = "3х1,5"  # most common default

        results.append({
            "brand": brand_raw,
            "section": section,
            "length": length,
        })

    return results


def parse_mtext(pairs):
    """Extract cable info from single MTEXT pair stream."""
    text_parts = []
    for code, val in pairs:
        if code in ("1", "3"):
            text_parts.append(val)
    text = _clean("".join(text_parts))
    bm = _CABLE_BRAND_RE.search(text)
    lm = _LEN_RE.search(text)
    if not bm or not lm:
        return []
    brand_raw = bm.group(1).rstrip("- ")
    sec_m = _SEC_RE.search(text)
    section = (f"{sec_m.group(1)}х{sec_m.group(2)}"
               if sec_m else "3х1,5")
    # laying method
    laying = ""
    low = text.lower()
    has_tray = "лотк" in low
    has_gofra = "гофр" in low
    if has_tray and has_gofra:
        laying = "в лотке+гофре"
    elif has_tray:
        laying = "в лотке"
    elif has_gofra:
        laying = "в гофре"
    elif "земле" in low or "траншее" in low:
        laying = "в земле"
    return [{
        "brand": brand_raw,
        "section": section,
        "length": int(lm.group(1)),
        "laying": laying,
        "text": text[:100],
    }]


def analyze_dxf(dxf_path):
    by_source = {"MTEXT": [], "TABLECONTENT": [], "ACAD_TABLE": [], "TEXT": []}
    for ent, pairs in iter_entities(dxf_path):
        if ent == "MTEXT":
            by_source["MTEXT"].extend(parse_mtext(pairs))
        elif ent == "TABLECONTENT":
            by_source["TABLECONTENT"].extend(parse_tablecontent(pairs))
    return by_source


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else \
        "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG"

    grand_total = 0
    grand_count = 0
    by_bs = defaultdict(lambda: [0, 0])

    for f in sorted(os.listdir(folder)):
        if not f.lower().endswith(".dxf"):
            continue
        p = os.path.join(folder, f)
        result = analyze_dxf(p)
        tot_n = 0
        tot_l = 0
        for source, items in result.items():
            n = len(items)
            L = sum(i["length"] for i in items)
            if n:
                print(f"  {f[:50]:50s} [{source:14s}]  n={n:4d}  L={L:6d}")
                tot_n += n
                tot_l += L
                for it in items:
                    k = f"{it['brand']} {it['section']}"
                    by_bs[k][0] += 1
                    by_bs[k][1] += it["length"]
        grand_count += tot_n
        grand_total += tot_l

    print(f"\n--- GRAND TOTAL: {grand_count} каб, {grand_total} м ---")
    print("\nПо маркам × сечениям (без дедупа):")
    for k, (n, L) in sorted(by_bs.items(), key=lambda x: -x[1][1]):
        print(f"  {k:50s}  {n:4d} каб  {L:6d} м")


if __name__ == "__main__":
    main()
