"""Read-only diagnostic: parse all MTEXT blocks in all DXF files of a folder
and extract (cable_type, length, laying_method) triples.

Uses true MTEXT grouping: walks DXF group codes (0/MTEXT ... 0/next_entity),
collects the text (group code 1 with continuations via 3) and insertion
point (codes 10/20/30). Geometric dedup by rounded (x, y) within 0.5 units.
"""
from __future__ import annotations

import io
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
    except (AttributeError, ValueError):
        pass


_CABLE_RE = re.compile(
    r"(ВБШвнг\(А\)-FRLS|ВБШвнг\(А\)-LS|ППГнг\(А\)-FRHF|ППГнг\(А\)-HF|"
    r"ВБШвнг[^\s,;\\]*|ППГнг[^\s,;\\]*|ВВГнг[^\s,;\\]*)"
    r"[^\d]{0,20}"
    r"(\d+[xх×]\d+(?:[.,]\d+)?)",
    re.IGNORECASE,
)
_LEN_RE = re.compile(r"L\s*[=\-]\s*(\d+)")
_DU_RE = re.compile(r"[\u0394\u2206ΔΔ]?U\s*=\s*(\d+[.,]\d+)\s*%")
_LAYING_RE = re.compile(
    r"в\s*(кабель[- ]?канале|гофре|лотке|земле|трубе|траншее)",
    re.IGNORECASE,
)
# Strip MTEXT codes: {\fArial|b0|i0|c204|p34;text}, \P paragraph, \L underline, etc.
_MTEXT_CTRL = re.compile(r"\\[A-Za-z]\w*|\\[pP]|\\~|[{}]|\\\\")


def _strip_mtext(text: str) -> str:
    t = _MTEXT_CTRL.sub(" ", text)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def iter_mtext_blocks(dxf_path: str):
    """Yield (text, x, y, z) for every MTEXT entity in the DXF.

    Low-level parser: reads group code pairs, when entity type is MTEXT
    accumulates code 1/3 (text) and 10/20/30 (insertion point) until next
    entity (code 0).
    """
    try:
        fh = open(dxf_path, "r", encoding="utf-8", errors="replace")
    except OSError:
        return
    with fh:
        cur_type = None
        cur_text_parts: list[str] = []
        cur_x = cur_y = cur_z = 0.0
        waiting_code = True
        code = None
        for line in fh:
            ls = line.rstrip("\n\r")
            stripped = ls.strip()
            if waiting_code:
                try:
                    code = int(stripped)
                except (ValueError, TypeError):
                    code = None
                waiting_code = False
                continue
            # value line
            value = ls
            waiting_code = True

            if code == 0:
                # flush previous
                if cur_type == "MTEXT" and cur_text_parts:
                    full = "".join(cur_text_parts)
                    yield (full, cur_x, cur_y, cur_z)
                cur_type = value.strip()
                cur_text_parts = []
                cur_x = cur_y = cur_z = 0.0
            elif cur_type == "MTEXT":
                if code == 1 or code == 3:
                    cur_text_parts.append(value)
                elif code == 10:
                    try:
                        cur_x = float(value)
                    except ValueError:
                        pass
                elif code == 20:
                    try:
                        cur_y = float(value)
                    except ValueError:
                        pass
                elif code == 30:
                    try:
                        cur_z = float(value)
                    except ValueError:
                        pass
        # last
        if cur_type == "MTEXT" and cur_text_parts:
            yield ("".join(cur_text_parts), cur_x, cur_y, cur_z)


def _classify_laying(text: str) -> str:
    matches = [m.group(1).lower() for m in _LAYING_RE.finditer(text)]
    if not matches:
        return ""
    if any("лотк" in x for x in matches) and any("гофр" in x for x in matches):
        return "в лотке+гофре"
    if any("лотк" in x for x in matches):
        return "в лотке"
    if any("гофр" in x for x in matches):
        return "в гофре"
    if any("кабель" in x or "канал" in x for x in matches):
        return "в кабель-канале"
    if any("земле" in x or "траншее" in x or "трубе" in x for x in matches):
        return "в земле"
    return matches[0]


def analyze_folder(folder: str):
    all_cables = []
    per_file = defaultdict(lambda: {"count": 0, "length": 0})
    for f in sorted(os.listdir(folder)):
        if not f.lower().endswith(".dxf"):
            continue
        p = os.path.join(folder, f)
        file_cables = 0
        file_len = 0
        for text, x, y, z in iter_mtext_blocks(p):
            clean = _strip_mtext(text)
            if not clean:
                continue
            cm = _CABLE_RE.search(clean)
            if not cm:
                continue
            lm = _LEN_RE.search(clean)
            if not lm:
                continue
            brand = cm.group(1)
            sec = cm.group(2).replace("x", "х").replace("×", "х")
            length = int(lm.group(1))
            laying = _classify_laying(clean)
            du_m = _DU_RE.search(clean)
            du = du_m.group(1).replace(",", ".") if du_m else ""
            all_cables.append({
                "file": f,
                "brand": brand,
                "sec": sec,
                "length": length,
                "laying": laying,
                "du": du,
                "x": round(x, 1),
                "y": round(y, 1),
                "z": round(z, 1),
                "text": clean[:80],
            })
            file_cables += 1
            file_len += length
        if file_cables:
            per_file[f] = {"count": file_cables, "length": file_len}
            print(f"  {f[:55]:55s}  n={file_cables:4d}  L={file_len:6d}")

    # Geometric dedup: within same file, (rounded x, y) cluster
    seen = set()
    unique = []
    for c in all_cables:
        key = (c["file"], c["brand"], c["sec"], c["length"],
               round(c["x"], 0), round(c["y"], 0), c["laying"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(c)

    print(f"\n--- TOTAL raw MTEXT matches: {len(all_cables)}")
    print(f"--- After geometric dedup:   {len(unique)}")
    total_len = sum(c["length"] for c in unique)
    print(f"--- Total length unique:     {total_len} м")

    # By brand × section
    print("\nПо маркам × сечениям (после дедупа):")
    by_bs = defaultdict(lambda: [0, 0])
    for c in unique:
        k = f"{c['brand']} {c['sec']}"
        by_bs[k][0] += 1
        by_bs[k][1] += c["length"]
    for k, (n, L) in sorted(by_bs.items(), key=lambda x: -x[1][1]):
        print(f"  {k:40s}  {n:4d} каб  {L:6d} м")

    # By laying method
    print("\nПо способу прокладки:")
    by_lay = defaultdict(lambda: [0, 0])
    for c in unique:
        by_lay[c["laying"] or "(не указан)"][0] += 1
        by_lay[c["laying"] or "(не указан)"][1] += c["length"]
    for k, (n, L) in sorted(by_lay.items(), key=lambda x: -x[1][1]):
        print(f"  {k:30s}  {n:4d} каб  {L:6d} м")

    return unique


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else \
        "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG"
    analyze_folder(folder)
