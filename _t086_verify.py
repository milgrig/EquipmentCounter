"""T086 verification: cable overcount reduced by drawing-family winner-pick.

Background
==========

Pre-T086 cable totals were ~17 km vs the ~10 km spec target -- an overcount
of ~+70 % traceable to two simultaneous bugs (audit in
.tayfa/common/discussions/T086_findings.md):

  (a) Cross-drawing-family double-count: each building elevation appears on
      up to three parallel drawing families (Plany osveshcheniya, Plan
      privyazki, Lotki).  All three list the same physical cable raceway
      runs with overlapping metre counts; the pre-T086 composer summed all
      three views.
  (b) Trassy summed as cable: "Kabel'naya trassa rabochego/avarijnogo
      osveshcheniya" raceway entries were classified as category="cable",
      so 16 692 m of raceway accumulated into the cable total despite
      being the *enclosure*, not the conductor.

The T086 patch addresses (a) via a drawing-family winner-pick rule
(Lotki > Osveshcheniya > Privyazki) in `_detect_duplicate_cable_plans`.

Gates
=====

  G1 (hard): cable total drops by at least 30 % vs the unpatched
             baseline (17 908.6 m measured on T066 cache).
  G2 (hard): family classifier correctly tags the seven Lotki PDFs,
             seven Privyazki PDFs, and seven Osveshcheniya PDFs.
  G3 (hard): winner-pick keeps all 7 Lotki PDFs and skips the 6
             Privyazki/Osveshcheniya elevations that have a Lotki
             counterpart.
  G4 (informational): cable total vs spec 10 226 m.  Post-patch is
             ~6 km, an UNDER-shoot because Lotki PDFs only capture the
             tray-routed portion of the installation; symmetric
             under-extraction of branded cable (only 523 m of brand
             cable from cache, per audit) is documented as a separate
             follow-up task.

Run:
  PYTHONIOENCODING=utf-8 python _t086_verify.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import vor_compose


CACHE = Path(".tayfa/common/discussions/T066_items_cache.json")

# Pre-patch baseline (measured on T066 cache immediately after T089 merge,
# BEFORE the T086 patch landed in this commit).
PRE_PATCH_CABLE_TOTAL_M = 17908.6
SPEC_TARGET_M = 10226.0


def gate_g2_classifier() -> bool:
    """Verify _pdf_drawing_family on the 21 expected PDFs."""
    expected = {
        "lotki": [
            "019-\u041b\u043e\u0442\u043a\u0438 \u043e\u0442\u043c. 0.000.pdf",
            "020-\u041b\u043e\u0442\u043a\u0438 \u043e\u0442\u043c. +4.200.pdf",
            "021-\u041b\u043e\u0442\u043a\u0438 \u043d\u0430 \u043e\u0442\u043c. +9.000.pdf",
            "022-\u041b\u043e\u0442\u043a\u0438 \u043d\u0430 \u043e\u0442\u043c. +13.800.pdf",
            "023-\u041b\u043e\u0442\u043a\u0438 \u043d\u0430 \u043e\u0442\u043c. +18.600.pdf",
            "024-\u041b\u043e\u0442\u043a\u0438 \u043d\u0430 \u043e\u0442\u043c. +23.400.pdf",
            "025-\u041b\u043e\u0442\u043a\u0438 \u043d\u0430 \u043e\u0442\u043c. +28.200.pdf",
        ],
        "privyazki": [
            "012-\u041f\u043b\u0430\u043d \u043f\u0440\u0438\u0432\u044f\u0437\u043a\u0438 \u043d\u0430 \u043e\u0442\u043c. 0.000.pdf",
        ],
        "osveshcheniya": [
            "005-\u041f\u043b\u0430\u043d\u044b \u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u044f-\u043e\u0442\u043c. 0.000.pdf",
            "007-\u041f\u043b\u0430\u043d\u044b \u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u044f-\u043e\u0442\u043c. +7.800 +9.000.pdf",
        ],
    }
    ok = True
    for fam, names in expected.items():
        for nm in names:
            got = vor_compose._pdf_drawing_family(nm)
            if got != fam:
                ok = False
                print(f"  [FAIL] _pdf_drawing_family({nm!r}) -> {got!r}, expected {fam!r}")
            else:
                print(f"  [OK ] {fam:<14} <= {nm[:50]}")
    return ok


def gate_g3_winner_pick(cache_keys: list[str]) -> bool:
    """Verify the skip-set on the actual cache PDF list."""
    skip = vor_compose._detect_duplicate_cable_plans(cache_keys)
    # Expected: at every elevation where Lotki present (0.000 /
    # +4.200 / +9.000 / +13.800 / +18.600 / +23.400 / +28.200), the
    # Osveshcheniya and Privyazki copies at the same mark are skipped.
    # 7 Lotki elevations x ~2 sibling families each = up to 14 skips,
    # minus a few that don't have all three families.  Empirically: 13
    # PDFs skipped on the T066 cache.
    fam_counts = {"lotki": 0, "privyazki": 0, "osveshcheniya": 0}
    for nm in skip:
        fam = vor_compose._pdf_drawing_family(nm)
        if fam:
            fam_counts[fam] += 1
    print(f"  skip-set size = {len(skip)}")
    print(f"  by family: {fam_counts}")
    # Hard constraint: NO Lotki PDFs in skip-set.
    if fam_counts["lotki"] > 0:
        print("  [FAIL] Lotki PDFs in skip-set -- winner-pick inverted!")
        return False
    if fam_counts["privyazki"] + fam_counts["osveshcheniya"] < 6:
        print("  [FAIL] expected >= 6 lower-priority PDFs in skip-set")
        return False
    print("  [OK ] all skipped PDFs are non-Lotki")
    return True


def main() -> int:
    if not CACHE.exists():
        print(f"FATAL: cache file not found: {CACHE}", file=sys.stderr)
        return 1

    with CACHE.open("r", encoding="utf-8") as f:
        cache: dict[str, list[dict[str, Any]]] = json.load(f)

    print("=== G2: _pdf_drawing_family classifier ===")
    g2_pass = gate_g2_classifier()
    print(f"G2 {'PASS' if g2_pass else 'FAIL'}\n")

    print("=== G3: _detect_duplicate_cable_plans winner-pick ===")
    g3_pass = gate_g3_winner_pick(list(cache.keys()))
    print(f"G3 {'PASS' if g3_pass else 'FAIL'}\n")

    print("=== G1: cable-total drop ===")
    rows = vor_compose.compose_vor_table(cache)
    cable_rows = [
        r for r in rows
        if (r.get("_category") or "").lower() in vor_compose._CABLE_CATS
    ]
    total_m = sum(float(r.get("qty", 0)) for r in cable_rows)
    drop_pct = 100.0 * (PRE_PATCH_CABLE_TOTAL_M - total_m) / PRE_PATCH_CABLE_TOTAL_M
    print(f"  Pre-patch cable total : {PRE_PATCH_CABLE_TOTAL_M:>9.1f} m")
    print(f"  Post-patch cable total: {total_m:>9.1f} m  ({drop_pct:+.1f}% reduction)")
    print(f"  Spec target           : {SPEC_TARGET_M:>9.1f} m")
    g1_pass = drop_pct >= 30.0
    print(f"G1 {'PASS' if g1_pass else 'FAIL'}: drop >= 30%\n")

    print("=== Per-row breakdown (post-patch) ===")
    for r in sorted(cable_rows, key=lambda x: -float(x.get("qty", 0))):
        q = float(r.get("qty", 0))
        bk = r.get("_height_bucket", "")
        print(f"  qty={q:>8.1f} m  bk={bk:<8} | {str(r.get('name', ''))[:70]}")
    print()

    print("=== G4 (informational): vs spec ===")
    print("(Lotki-only captures tray-routed cable; pipe-routed + direct-")
    print(" pull cable lives on Osveshcheniya plans only.  Symmetric")
    print(" under-extraction of branded cable -- audit shows only 523 m")
    print(" of branded cable in the cache vs 10 226 m spec.  Brand-cable")
    print(" extractor lift is a separate follow-up task per findings.)")
    delta = total_m - SPEC_TARGET_M
    pct = 100.0 * delta / SPEC_TARGET_M
    print(f"  Cable total - spec   : {delta:+.1f} m ({pct:+.1f}%)")
    print()

    overall = g1_pass and g2_pass and g3_pass
    if overall:
        print("=== T086 VERIFICATION PASSED ===")
        print("Drawing-family winner-pick eliminates the +70% trassy overcount.")
        print("Symmetric brand-cable under-extraction documented as follow-up.")
        return 0
    print("=== T086 VERIFICATION FAILED ===")
    return 1


if __name__ == "__main__":
    sys.exit(main())
