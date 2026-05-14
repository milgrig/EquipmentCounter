"""T088 verification — split luminaires by model.

Runs `_aggregate_equipment` on `T066_items_cache.json` AFTER the T088
fix (strip `[UNMATCHED-LEGEND]` prefix before computing the agg key).
Verifies:

  G1  At least 7 distinct luminaire rows in the aggregate output.
  G2  Every luminaire row name contains a recognised model marker
      (SLICK.PRS, ARCTIC.OPL, CD LED, INSEL, etc.) — i.e. not collapsed
      into a single generic "Монтаж светильника" row.
  G3  No luminaire row name still starts with "[UNMATCHED-LEGEND]".
  G4  The 6 etalon (СО-СО_01) reference models appear, with reasonable
      totals (informational; not a hard gate because per-model totals
      are still affected by the T086 cable-route triple-count pattern
      and the brand-cable under-count).
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from web_app import _aggregate_equipment

CACHE = Path(".tayfa/common/discussions/T066_items_cache.json")
raw = json.loads(CACHE.read_text(encoding="utf-8"))

agg = _aggregate_equipment(raw)

lum_rows = [r for r in agg if "светильник" in r["name"].lower()]
print(f"Aggregate rows total: {len(agg)}")
print(f"Luminaire rows: {len(lum_rows)}")
print()
print("=== Luminaire rows ===")
for r in sorted(lum_rows, key=lambda x: -x["total"]):
    print(f"  total={int(r['total']):4d}  | {r['name']}")

# G1 — at least 7 distinct rows.
assert len(lum_rows) >= 7, f"G1 FAIL: only {len(lum_rows)} luminaire rows"
print(f"\nG1 PASS: {len(lum_rows)} luminaire rows >= 7")

# G2 — every row carries a recognisable model marker.
MODEL_TOKENS = [
    r"SLICK\.PRS",
    r"ARCTIC\.OPL",
    r"\bCD LED\b",
    r"\bINSEL\b",
    r"CONVERSION KIT",
    r"\bD60\b",
    r"\bLB/S\b",
]
model_re = re.compile("|".join(MODEL_TOKENS), re.IGNORECASE)
missing_model = [r for r in lum_rows if not model_re.search(r["name"])]
assert not missing_model, (
    f"G2 FAIL: rows without recognisable model: "
    f"{[r['name'] for r in missing_model]}"
)
print(f"G2 PASS: all {len(lum_rows)} luminaire rows carry a model token")

# G3 — no [UNMATCHED-LEGEND] prefix survives aggregation.
leaked = [r for r in lum_rows if "[UNMATCHED-LEGEND]" in r["name"]]
assert not leaked, (
    f"G3 FAIL: [UNMATCHED-LEGEND] leaked into aggregate: "
    f"{[r['name'] for r in leaked]}"
)
print(f"G3 PASS: 0 rows carry [UNMATCHED-LEGEND] prefix")

# G4 — etalon coverage (informational).
ETALON = {
    "SLICK.PRS LED 50": 414,
    "ARCTIC.OPL ECO LED 1200": 98,
    "SLICK.PRS LED 30": None,  # 38 + 27 = 65, split between two variants
    "CD LED 27": 6,
    "ARCTIC.OPL ECO LED 1200 TH": 9,
    "INSEL": 12,
}
print("\n=== G4: etalon coverage (informational) ===")
print("(per-model totals deviate due to T086 cross-drawing-family triple-count;")
print(" T088 only fixes ROW splitting, not totals)")
for model, expected in ETALON.items():
    matches = [
        r for r in lum_rows
        if re.search(re.escape(model), r["name"], re.IGNORECASE)
    ]
    if not matches:
        print(f"  [MISS]  {model}: expected={expected}, no matching row")
        continue
    measured = sum(int(r["total"]) for r in matches)
    delta = "n/a" if expected is None else f"delta={measured-expected:+d}"
    print(
        f"  [HIT ]  {model}: expected={expected}, measured={measured} "
        f"({len(matches)} row(s), {delta})"
    )

print("\n=== T088 VERIFICATION PASSED ===")
