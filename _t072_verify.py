"""T072 acceptance verifier -- thickness_extractor.

Gates:
  G1: cable cross_section set for >= 80% of items with a marked
      cable_label.
  G2: every lotok item has width_mm + height_mm.
  G3: every gofra item has diameter_mm.
  G4: imports clean.
"""

import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]


def _build_items() -> list[dict]:
    """Synthetic item set mirroring real vor_map_items output."""
    return [
        # 5 cable items with cable_label; 4 must yield cross_section.
        # 1 deliberately has nonsense label to test the unparseable path.
        {"name": "Prokladka kabelya VVGng 3x1,5", "unit": "\u043c",
         "category": "cable", "cable_label": "VVGng 3x1,5"},
        {"name": "Prokladka kabelya VBSHv 5x16", "unit": "\u043c",
         "category": "cable", "cable_label": "VBShv 5x16"},
        {"name": "Prokladka kabelya 4\u04452,5 NYM", "unit": "\u043c",
         "category": "cable_run", "cable_label": "NYM 4x2,5"},
        # KB-007 Unicode MULTIPLICATION SIGN U+00D7
        {"name": "Prokladka kabelya 3\u00d72,5 PPGng", "unit": "\u043c",
         "category": "cable", "cable_label": "PPGng 3\u00d72,5"},
        {"name": "Prokladka kabelya 1\u044516", "unit": "\u043c",
         "category": "cable", "cable_label": "1\u044516"},
        # Unparseable label -- should NOT be counted as "with cs".
        {"name": "Prokladka kabelya unknown", "unit": "\u043c",
         "category": "cable", "cable_label": "ZZZZZ"},
        # Cable without label -- excluded from G1 denominator.
        {"name": "Prokladka kabelya 6x4 raw", "unit": "\u043c",
         "category": "cable"},
        # 3 lotok items -- all must have width/height.
        {"name": "Montazh lotka 100x80",   "unit": "\u043c",
         "category": "cable_tray"},
        {"name": "Montazh lotka 200\u0445100", "unit": "\u043c",
         "category": "cable_tray"},
        {"name": "Lotok 300\u00d7100", "unit": "\u043c",
         "category": "lotok"},
        # 3 gofra items -- all must have diameter.
        {"name": "Montazh gofry D16", "unit": "\u043c",
         "category": "conduit_corrugated"},
        {"name": "Gofra PVH d.20",    "unit": "\u043c",
         "category": "conduit_corrugated"},
        {"name": "\u0413\u043e\u0444\u0440\u0430 PVH 25 \u043c\u043c",
         "unit": "\u043c", "category": "gofra"},
        # Non-cable / non-tray / non-gofra item -- must be ignored.
        {"name": "Svetilnik LED 40W", "unit": "\u0448\u0442",
         "category": "luminaire"},
    ]


def g1_cable_cs() -> bool:
    import thickness_extractor as tx
    items = _build_items()
    sm = tx.attribute_items(items)
    print("G1 cable_total=%d with_label=%d with_cs=%d share=%.1f%%"
          % (sm["cable_total"], sm["cable_with_label"],
             sm["cable_with_cs"], sm["cable_cs_share"] * 100))
    ok = True
    if sm["cable_with_label"] < 5:
        print("G1 FAIL: cable_with_label %d, expected >= 5"
              % sm["cable_with_label"])
        ok = False
    if sm["cable_cs_share"] < 0.80 - 1e-9:
        print("G1 FAIL: share %.1f%% < 80%% threshold"
              % (sm["cable_cs_share"] * 100))
        ok = False
    # Direct check: each labeled item with parseable label has cs.
    expected = {
        "VVGng 3x1,5":  "3\u04451,5",
        "VBShv 5x16":   "5\u044516",
        "NYM 4x2,5":    "4\u04452,5",
        "PPGng 3\u00d72,5": "3\u04452,5",
        "1\u044516":    "1\u044516",
    }
    for it in items:
        lbl = it.get("cable_label")
        if lbl in expected:
            got = (it.get("cross_section") or "")
            if got != expected[lbl]:
                print("G1 FAIL: label=%r expected cs=%r got=%r"
                      % (lbl, expected[lbl], got))
                ok = False
    print("G1 %s" % ("PASS" if ok else "FAIL"))
    return ok


def g2_lotok_dims() -> bool:
    import thickness_extractor as tx
    items = _build_items()
    sm = tx.attribute_items(items)
    print("G2 lotok_total=%d with_dims=%d"
          % (sm["lotok_total"], sm["lotok_with_dims"]))
    ok = True
    if sm["lotok_total"] < 3:
        print("G2 FAIL: lotok_total %d < 3" % sm["lotok_total"])
        ok = False
    if sm["lotok_with_dims"] != sm["lotok_total"]:
        print("G2 FAIL: %d/%d lotok items have width+height"
              % (sm["lotok_with_dims"], sm["lotok_total"]))
        ok = False
    expected = {
        "Montazh lotka 100x80": (100, 80),
        "Montazh lotka 200\u0445100": (200, 100),
        "Lotok 300\u00d7100": (300, 100),
    }
    for it in items:
        nm = it.get("name")
        if nm in expected:
            got = (it.get("width_mm"), it.get("height_mm"))
            if got != expected[nm]:
                print("G2 FAIL: %r expected=%s got=%s"
                      % (nm, expected[nm], got))
                ok = False
    print("G2 %s" % ("PASS" if ok else "FAIL"))
    return ok


def g3_gofra_diameter() -> bool:
    import thickness_extractor as tx
    items = _build_items()
    sm = tx.attribute_items(items)
    print("G3 gofra_total=%d with_diameter=%d"
          % (sm["gofra_total"], sm["gofra_with_diameter"]))
    ok = True
    if sm["gofra_total"] < 3:
        print("G3 FAIL: gofra_total %d < 3" % sm["gofra_total"])
        ok = False
    if sm["gofra_with_diameter"] != sm["gofra_total"]:
        print("G3 FAIL: %d/%d gofra items have diameter"
              % (sm["gofra_with_diameter"], sm["gofra_total"]))
        ok = False
    expected = {
        "Montazh gofry D16": 16,
        "Gofra PVH d.20":    20,
        "\u0413\u043e\u0444\u0440\u0430 PVH 25 \u043c\u043c": 25,
    }
    for it in items:
        nm = it.get("name")
        if nm in expected:
            got = it.get("diameter_mm")
            if got != expected[nm]:
                print("G3 FAIL: %r expected=%d got=%s"
                      % (nm, expected[nm], got))
                ok = False
    print("G3 %s" % ("PASS" if ok else "FAIL"))
    return ok


def g4_imports() -> bool:
    try:
        import thickness_extractor   # noqa: F401
        import route_classifier       # noqa: F401
        import height_bucketer        # noqa: F401
        import cable_length           # noqa: F401
        import importlib
        spec = importlib.util.find_spec("web_app")
        assert spec is not None
        print("G4 PASS: thickness_extractor + neighbours + web_app importable")
        return True
    except Exception as exc:
        print("G4 FAIL: %r" % exc)
        return False


def main() -> int:
    print("=" * 60)
    print("T072 acceptance verification")
    print("=" * 60)
    a = g1_cable_cs()
    print("-" * 60)
    b = g2_lotok_dims()
    print("-" * 60)
    c = g3_gofra_diameter()
    print("-" * 60)
    d = g4_imports()
    print("=" * 60)
    print("Overall: G1=%s G2=%s G3=%s G4=%s"
          % ("PASS" if a else "FAIL",
             "PASS" if b else "FAIL",
             "PASS" if c else "FAIL",
             "PASS" if d else "FAIL"))
    return 0 if (a and b and c and d) else 1


if __name__ == "__main__":
    raise SystemExit(main())
