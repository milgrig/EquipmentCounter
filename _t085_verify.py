"""T085 verification: exit signs split by model in vor_compose output.

Acceptance criteria (from task description):
  Spec ground truth (СО-СО_01): MERCURY=19, ATOM=30, SIRAH=49 шт.
  After fix, VOR must contain 3 distinct exit-sign rows per model
  with qty within ±10% of the PDF counter quantity.

Gates:
  G1 (hard)         : >= 3 distinct exit-sign material rows present in
                      compose_vor_table output, one per model
                      (MERCURY, ATOM, SIRAH/Pictogramma).
  G2 (hard)         : Every exit-sign row carries a model token in its
                      name (no model-erasing collapse to bare
                      "Montazh ... na vysote ...").
  G3 (hard)         : _is_exit_sign() safety net fires correctly on
                      both category-based and name-based detection.
  G4 (informational): Per-model totals vs spec.  Currently expected to
                      overshoot by ~2x due to T086 cross-drawing-family
                      triple-counting bug (sibling task).  Documented,
                      not blocking T085.

Run:
  PYTHONIOENCODING=utf-8 python _t085_verify.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

import vor_compose


CACHE = Path(".tayfa/common/discussions/T066_items_cache.json")


def _model_of(name: str) -> str:
    """Classify an exit-sign name into one of three model buckets."""
    if "MERCURY" in name:
        return "MERCURY"
    if "ATOM" in name:
        return "ATOM"
    if "\u041f\u0438\u043a\u0442\u043e\u0433\u0440\u0430\u043c\u043c" in name or "SIRAH" in name:
        return "SIRAH"
    return ""


def _is_exit_row(name: str) -> bool:
    return any(
        marker in name
        for marker in (
            "MERCURY", "ATOM", "SIRAH",
            "\u041f\u0438\u043a\u0442\u043e\u0433\u0440\u0430\u043c\u043c",
            "\u0412\u042b\u0425\u041e\u0414",   # ВЫХОД
            "\u0412\u044b\u0445\u043e\u0434",   # Выход
        )
    )


def gate_g3_safety_net() -> bool:
    """Verify _is_exit_sign() detection on synthetic edge cases."""
    cases = [
        # (item, expected, description)
        ({"category": "luminaire_exit", "name": "X"}, True,
         "category=luminaire_exit -> True"),
        ({"category": "exit_indicator", "name": "X"}, True,
         "category=exit_indicator -> True"),
        ({"category": "pictogram", "name": "X"}, True,
         "category=pictogram -> True"),
        ({"category": "other",
          "work_name": "\u041c\u043e\u043d\u0442\u0430\u0436 MERCURY LED"},
         True, "name-based MERCURY -> True"),
        ({"category": "other",
          "work_name": "ATOM 6500-3 LED SP AT FP"},
         True, "name-based ATOM -> True"),
        ({"category": "other",
          "work_name": "\u041f\u0438\u043a\u0442\u043e\u0433\u0440\u0430\u043c\u043c\u0430 \u0412\u044b\u0445\u043e\u0434"},
         True, "name-based Pictogramma -> True"),
        ({"category": "other",
          "work_name": "\u0421\u0432\u0435\u0442\u043e\u0432\u044b\u0435 \u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b\u0438 X"},
         True, "name-based 'Svetovye ukazateli' -> True"),
        ({"category": "luminaire", "work_name": "SLICK.PRS LED 50"}, False,
         "regular luminaire -> False"),
        ({"category": "cable", "work_name": "\u0412\u0411\u0428\u0432\u043d\u0433-LS"}, False,
         "cable -> False"),
    ]
    all_pass = True
    for item, expected, desc in cases:
        got = vor_compose._is_exit_sign(item)
        ok = got == expected
        marker = "[OK ]" if ok else "[FAIL]"
        print(f"  {marker} {desc}: got={got}, expected={expected}")
        if not ok:
            all_pass = False
    return all_pass


def gate_g_row_name_preserves_model() -> bool:
    """Verify _row_name keeps the model when item is reclassified to
    luminaire_exit with a mount + bucket that would otherwise trigger
    the collapse-to-mount path."""
    item = {
        "category": "luminaire_exit",
        "work_name": "\u041c\u043e\u043d\u0442\u0430\u0436 MERCURY LED Ex 20W DLW",
        "mount": "shpilka",
        "height_bucket": "<5m",
        "unit": "\u0448\u0442",
    }
    name = vor_compose._row_name(item)
    has_model = "MERCURY" in name
    marker = "[OK ]" if has_model else "[FAIL]"
    print(f"  {marker} _row_name preserves MERCURY when category=luminaire_exit + mount=shpilka + bucket=<5m")
    print(f"        -> {name!r}")
    return has_model


def main() -> int:
    if not CACHE.exists():
        print(f"FATAL: cache file not found: {CACHE}", file=sys.stderr)
        return 1

    with CACHE.open("r", encoding="utf-8") as f:
        cache: dict[str, list[dict[str, Any]]] = json.load(f)

    rows = vor_compose.compose_vor_table(cache)
    print(f"compose_vor_table produced {len(rows)} rows total\n")

    exit_rows = [r for r in rows if _is_exit_row(str(r.get("name", "")))]
    material_rows = [r for r in exit_rows if r.get("_kind") == "material"]
    work_rows = [r for r in exit_rows if r.get("_kind") == "work"]

    print("=== Exit-sign material rows (one per model expected) ===")
    by_model: dict[str, dict[str, Any]] = {}
    for r in material_rows:
        m = _model_of(str(r.get("name", "")))
        if m:
            by_model[m] = r
        print(f"  model={m or '?':<8} qty={r.get('qty', 0):>4} unit={r.get('unit', ''):<3} "
              f"| {str(r.get('name', ''))[:80]}")
    print()

    # --- G1: >= 3 distinct exit-sign material rows, one per model ---
    g1_pass = (
        len(material_rows) >= 3
        and "MERCURY" in by_model
        and "ATOM" in by_model
        and "SIRAH" in by_model
    )
    print(f"G1 {'PASS' if g1_pass else 'FAIL'}: "
          f"{len(material_rows)} material rows, models present: "
          f"{sorted(by_model.keys())}")

    # --- G2: every exit row carries a model token ---
    rows_missing_model = [
        r for r in exit_rows if not _model_of(str(r.get("name", "")))
    ]
    g2_pass = not rows_missing_model
    print(f"G2 {'PASS' if g2_pass else 'FAIL'}: "
          f"{len(rows_missing_model)} exit rows missing model token")
    for r in rows_missing_model:
        print(f"  [MISS] {str(r.get('name', ''))[:90]}")

    # --- G3: safety-net unit tests ---
    print("\n=== G3: _is_exit_sign safety-net edge cases ===")
    g3_pass = gate_g3_safety_net()
    print(f"G3 {'PASS' if g3_pass else 'FAIL'}: safety-net detection")

    # --- G3b: _row_name preserves model under reclassified category ---
    print("\n=== G3b: _row_name preservation under reclassified category ===")
    g3b_pass = gate_g_row_name_preserves_model()
    print(f"G3b {'PASS' if g3b_pass else 'FAIL'}: model preserved")

    # --- G4: informational quantity comparison vs spec ---
    print("\n=== G4 (informational): per-model totals vs spec ===")
    print("(quantities currently overshoot due to T086 cross-family triple-count;")
    print(" T085 only fixes ROW SPLITTING, not totals)")
    spec = {"MERCURY": 19, "ATOM": 30, "SIRAH": 49}
    for model, expected in spec.items():
        r = by_model.get(model)
        if r is None:
            print(f"  [MISS] {model}: row absent")
            continue
        got = int(r.get("qty", 0))
        within_10pct = (
            expected * 0.9 <= got <= expected * 1.1
        )
        marker = "[OK  ]" if within_10pct else "[DEV ]"
        delta = got - expected
        pct = (100.0 * delta / expected) if expected else 0.0
        print(f"  {marker} {model:<8} got={got:>3} spec={expected:>3} "
              f"delta={delta:+d} ({pct:+.0f}%)")

    print()
    overall = g1_pass and g2_pass and g3_pass and g3b_pass
    if overall:
        print("=== T085 VERIFICATION PASSED ===")
        print("Row-splitting gate (G1+G2+G3+G3b) met.")
        print("Quantity gate (G4) gated on T086 sibling task.")
        return 0
    print("=== T085 VERIFICATION FAILED ===")
    return 1


if __name__ == "__main__":
    sys.exit(main())
