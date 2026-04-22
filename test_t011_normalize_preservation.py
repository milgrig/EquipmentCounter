"""T011: Unit tests for _normalize() — verify critical tokens are preserved.

Task T011 / Sprint S5.1: cable brand, cross-section, diameter, lamp power and
IP-class tokens are discriminating and must survive normalization so that the
fuzzy comparison can distinguish otherwise-similar product lines.

Run with:
    python test_t011_normalize_preservation.py
"""

from __future__ import annotations

import sys

from test_vor_cross_format import _normalize


def _assert_contains(sample: str, must_contain: list[str]) -> tuple[bool, str]:
    out = _normalize(sample)
    missing = [tok for tok in must_contain if tok.lower() not in out]
    if missing:
        return False, f"  FAIL {sample!r}\n        -> {out!r}\n        missing: {missing}"
    return True, f"  OK   {sample!r:55s} -> {out!r}"


def main() -> int:
    cases: list[tuple[str, list[str]]] = [
        # ── Cable brands ──
        ("Кабель VVGng-LS 3x1.5",        ["vvgng", "3x1.5"]),
        ("Кабель ВВГнг(A)-LS 5х2,5",     ["ввгнг", "5х2,5"]),
        ("Провод ПуГВ 1х4",              ["пугв", "1х4"]),
        ("Кабель PPGng 3x2.5",           ["ppgng", "3x2.5"]),
        ("Кабель VBShvng 4x10",          ["vbshvng", "4x10"]),
        ("Кабель KGng 3x1.5",            ["kgng", "3x1.5"]),
        ("PuGV 1x4",                     ["pugv", "1x4"]),
        # ── Cross-sections ──
        ("Кабель 3x1.5",                 ["3x1.5"]),
        ("Кабель 5x2.5",                 ["5x2.5"]),
        # ── Diameters ──
        ("Труба d=20",                   ["d=20"]),
        ("Труба d 16",                   ["d 16"]),
        ("Труба гофрированная d=16",     ["d=16"]),
        # ── Lamp power (was stripped before T011) ──
        ("Светильник OWP EVO 30W OPL",   ["30w"]),
        ("Светильник 40Vt IP44",         ["40vt", "ip44"]),
        ("Светильник 30 Вт IP55",        ["30 вт", "ip55"]),
        # ── IP-class ──
        ("Светильник IP44",              ["ip44"]),
        ("Светильник IP55",              ["ip55"]),
        ("Светильник IP20",              ["ip20"]),
        # ── Combined (full real-world strings) ──
        ("Кабель VVGng-LS 3x1.5 IP55",                 ["vvgng", "3x1.5", "ip55"]),
        ("Светильник OWP EVO 30W OPL 840 WH IP54",     ["30w", "ip54"]),
    ]

    print("=" * 78)
    print("  T011: _normalize() preservation unit test")
    print("=" * 78)

    passed = 0
    failed = 0
    for sample, tokens in cases:
        ok, msg = _assert_contains(sample, tokens)
        print(msg)
        if ok:
            passed += 1
        else:
            failed += 1

    print("-" * 78)
    print(f"  Result: {passed} passed, {failed} failed (total {passed + failed})")
    print("=" * 78)
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
