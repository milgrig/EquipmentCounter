"""Probe v3: TABLECONTENT is the authoritative source (cable journal).
MTEXT cells in the table area are just visual rendering of the same rows,
so we keep TABLECONTENT entries and ignore MTEXT duplicates.

Dedup within TABLECONTENT by (brand, section, length) with a counter, so
several cables with the same params but different endpoints still count
as separate rows (TABLECONTENT preserves them naturally — each row is a
separate cell group).
"""
from __future__ import annotations

import io
import os
import re
import sys
from collections import defaultdict

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

from _cable_probe2 import iter_entities, parse_tablecontent, parse_mtext


def analyze_preferring_tablecontent(dxf_path):
    """Return combined cable list: TABLECONTENT rows are primary;
    MTEXT rows are added only if total per (brand,section,length) triple
    in MTEXT exceeds count from TABLECONTENT."""
    tc = []
    mt = []
    for ent, pairs in iter_entities(dxf_path):
        if ent == "TABLECONTENT":
            tc.extend(parse_tablecontent(pairs))
        elif ent == "MTEXT":
            mt.extend(parse_mtext(pairs))

    # Count triples in TC
    tc_key = defaultdict(int)
    for it in tc:
        k = (it["brand"], it["section"], it["length"])
        tc_key[k] += 1

    # Count triples in MT
    mt_key = defaultdict(int)
    for it in mt:
        k = (it["brand"], it["section"], it["length"])
        mt_key[k] += 1

    combined = list(tc)
    # For each MT triple, add extras beyond what's in TC
    extras_added = 0
    for k, mt_count in mt_key.items():
        extra = mt_count - tc_key.get(k, 0)
        if extra > 0:
            # Find original MT items matching this key and add 'extra' of them
            matched = [it for it in mt if (it["brand"], it["section"],
                                            it["length"]) == k]
            for it in matched[:extra]:
                combined.append(it)
                extras_added += 1

    return combined, len(tc), len(mt), extras_added


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else \
        "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG"

    all_cables = []
    for f in sorted(os.listdir(folder)):
        if not f.lower().endswith(".dxf"):
            continue
        p = os.path.join(folder, f)
        cables, n_tc, n_mt, n_extra = analyze_preferring_tablecontent(p)
        if cables or n_mt or n_tc:
            n = len(cables)
            L = sum(c["length"] for c in cables)
            print(f"  {f[:50]:50s}  tc={n_tc:3d} mt={n_mt:3d} "
                  f"+extra_mt={n_extra:2d}  TOTAL={n:3d}  L={L:5d}")
        all_cables.extend(cables)

    print(f"\n--- TOTAL: {len(all_cables)} каб, "
          f"{sum(c['length'] for c in all_cables)} м ---")

    by_bs = defaultdict(lambda: [0, 0])
    for c in all_cables:
        k = f"{c['brand']} {c['section']}"
        by_bs[k][0] += 1
        by_bs[k][1] += c["length"]
    print("\nПо маркам × сечениям:")
    for k, (n, L) in sorted(by_bs.items(), key=lambda x: -x[1][1]):
        print(f"  {k:40s}  {n:4d} каб  {L:6d} м")

    print("\nЭТАЛОН: 116 каб, 9864 м")


if __name__ == "__main__":
    main()
