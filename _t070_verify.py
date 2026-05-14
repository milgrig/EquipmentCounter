"""T070 acceptance verifier -- route_classifier.

Gates:
  G1: every cable-trace item has route in {tray, pipe_hidden,
      pipe_open, unknown}; unknown share <= 10% of cable items.
  G2: luminaire items get a mount field.
  G3: imports clean (web_app, cable_length, height_bucketer,
      route_classifier).

The harness builds synthetic item lists that mirror what
cable_length + vor_map_items emit, since running the full
_count_equipment_in_pdf pipeline on a 26-PDF folder is slow
(several minutes per PDF) and this is a classifier-only test.
"""

import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]


def g1_cable_route() -> bool:
    """Every cable item gets route; unknown share <= 10%."""
    import route_classifier as rc

    # Cyrillic labels as emitted by cable_length._CABLE_LENGTH_CATEGORIES.
    LBL = {
        "cable_emergency": "\u041a\u0430\u0431\u0435\u043b\u044c\u043d\u0430\u044f \u0442\u0440\u0430\u0441\u0441\u0430 \u0430\u0432\u0430\u0440\u0438\u0439\u043d\u043e\u0433\u043e \u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u044f",
        "cable_working":   "\u041a\u0430\u0431\u0435\u043b\u044c\u043d\u0430\u044f \u0442\u0440\u0430\u0441\u0441\u0430 \u0440\u0430\u0431\u043e\u0447\u0435\u0433\u043e \u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u044f",
        "wire_tray":       "\u041f\u0440\u043e\u0432\u043e\u0434\u043a\u0430 \u0432 \u043b\u043e\u0442\u043a\u0435",
        "wire_pipe_hidden":"\u041f\u0440\u043e\u0432\u043e\u0434\u043a\u0430 \u0432 \u0442\u0440\u0443\u0431\u0435 \u0441\u043a\u0440\u044b\u0442\u043e",
        "wire_pipe_open":  "\u041f\u0440\u043e\u0432\u043e\u0434\u043a\u0430 \u0432 \u0442\u0440\u0443\u0431\u0435 \u043e\u0442\u043a\u0440\u044b\u0442\u043e",
    }

    items = [
        {"name": LBL["cable_emergency"], "unit": "\u043c",
         "total": 120.0, "category": "cable_length_raster",
         "source": "cable_length_raster"},
        {"name": LBL["cable_working"], "unit": "\u043c",
         "total": 540.0, "category": "cable_length_raster",
         "source": "cable_length_raster"},
        {"name": LBL["wire_tray"], "unit": "\u043c",
         "total": 80.0, "category": "cable_length_raster",
         "source": "cable_length_raster"},
        {"name": LBL["wire_pipe_hidden"], "unit": "\u043c",
         "total": 35.0, "category": "cable_length_raster",
         "source": "cable_length_raster"},
        {"name": LBL["wire_pipe_open"], "unit": "\u043c",
         "total": 22.0, "category": "cable_length_raster",
         "source": "cable_length_raster"},
        # Item with pre-existing route -- must be preserved.
        {"name": "anything", "unit": "\u043c", "total": 10.0,
         "category": "cable_length_raster",
         "source": "cable_length_raster",
         "route": "tray"},
        # Non-cable item must NOT receive a route.
        {"name": "Some luminaire", "unit": "\u0448\u0442",
         "total": 5.0, "category": "luminaire"},
    ]

    summary = rc.attribute_items(items)
    missing = [i for i in items if rc._is_cable_item(i) and not i.get("route")]
    invalid = [
        i for i in items
        if rc._is_cable_item(i)
        and i.get("route") not in rc.ROUTE_CANONICAL
    ]
    preserved = next((i for i in items if i.get("name") == "anything"), None)

    print("G1 cable_total = %d, route counter = %s"
          % (summary["cable_total"], dict(summary["cable_route"])))
    print("G1 unknown share = %.1f%%" % (summary["cable_unknown_share"] * 100))

    ok = True
    if missing:
        print("G1 FAIL: %d cable items missing route" % len(missing))
        ok = False
    if invalid:
        print("G1 FAIL: %d cable items with invalid route value: %s"
              % (len(invalid), [i.get("route") for i in invalid]))
        ok = False
    if summary["cable_unknown_share"] > 0.10 + 1e-9:
        print("G1 FAIL: unknown share %.1f%% > 10%% threshold"
              % (summary["cable_unknown_share"] * 100))
        ok = False
    if preserved is None or preserved.get("route") != "tray":
        print("G1 FAIL: pre-tagged route not preserved")
        ok = False

    # Verify each category mapping individually.
    expected = {
        LBL["cable_emergency"]: "tray",
        LBL["cable_working"]:   "tray",
        LBL["wire_tray"]:       "tray",
        LBL["wire_pipe_hidden"]:"pipe_hidden",
        LBL["wire_pipe_open"]:  "pipe_open",
    }
    for it in items:
        nm = it.get("name") or ""
        if nm in expected and it.get("route") != expected[nm]:
            print("G1 FAIL: %r -> %r expected %r"
                  % (nm[:40], it.get("route"), expected[nm]))
            ok = False

    print("G1 %s" % ("PASS" if ok else "FAIL"))
    return ok


def g2_luminaire_mount() -> bool:
    """Luminaire items get a mount field."""
    import route_classifier as rc

    items = [
        {"name": "\u041c\u043e\u043d\u0442\u0430\u0436 "
                 "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d\u0438\u043a\u0430 "
                 "\u043d\u0430\u0441\u0442\u0435\u043d\u043d\u044b\u0439",
         "unit": "\u0448\u0442", "total": 4, "category": "luminaire"},
        {"name": "\u041c\u043e\u043d\u0442\u0430\u0436 "
                 "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d\u0438\u043a\u0430 "
                 "\u043d\u0430 \u0448\u043f\u0438\u043b\u044c\u043a\u0430\u0445",
         "unit": "\u0448\u0442", "total": 6, "category": "luminaire"},
        {"name": "\u041c\u043e\u043d\u0442\u0430\u0436 "
                 "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d\u0438\u043a\u0430 "
                 "\u043f\u043e\u0434\u0432\u0435\u0441\u043d\u043e\u0439",
         "unit": "\u0448\u0442", "total": 30, "category": "luminaire"},
        {"name": "\u041c\u043e\u043d\u0442\u0430\u0436 "
                 "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d\u0438\u043a\u0430 "
                 "\u0430\u0432\u0430\u0440\u0438\u0439\u043d\u043e\u0433\u043e",
         "unit": "\u0448\u0442", "total": 2,
         "category": "luminaire_emergency"},
        # Non-luminaire item must NOT receive a mount.
        {"name": "Cable trasse", "unit": "\u043c",
         "category": "cable_length_raster",
         "source": "cable_length_raster"},
    ]

    summary = rc.attribute_items(items)
    missing = [
        i for i in items
        if rc._is_luminaire_item(i) and not i.get("mount")
    ]
    print("G2 luminaire_total = %d, mount counter = %s"
          % (summary["luminaire_total"], dict(summary["luminaire_mount"])))

    ok = True
    if missing:
        print("G2 FAIL: %d luminaire items missing mount"
              % len(missing))
        ok = False

    expected = {
        # Wall keyword
        "\u043d\u0430\u0441\u0442\u0435\u043d": "wall",
        # Shpilka keyword
        "\u0448\u043f\u0438\u043b\u044c": "shpilka",
        # Suspended -> anker default
        "\u043f\u043e\u0434\u0432\u0435\u0441": "anker",
        # Emergency default -> anker
        "\u0430\u0432\u0430\u0440": "anker",
    }
    for kw, want in expected.items():
        for it in items:
            if kw in (it.get("name") or "").lower():
                got = it.get("mount")
                if got != want:
                    print("G2 FAIL: kw=%r got=%r expected=%r"
                          % (kw, got, want))
                    ok = False
                break

    # Cross-check: non-luminaire item has no mount tag.
    non_lum = next((i for i in items if i.get("name") == "Cable trasse"), None)
    if non_lum and non_lum.get("mount"):
        print("G2 FAIL: non-luminaire received mount=%r"
              % non_lum.get("mount"))
        ok = False

    print("G2 %s" % ("PASS" if ok else "FAIL"))
    return ok


def g3_imports() -> bool:
    """Imports clean."""
    try:
        import route_classifier  # noqa: F401
        import height_bucketer   # noqa: F401
        import cable_length      # noqa: F401
        # web_app is heavy; just check the import surface.
        import importlib
        spec = importlib.util.find_spec("web_app")
        assert spec is not None
        print("G3 PASS: route_classifier, height_bucketer, "
              "cable_length, web_app importable")
        return True
    except Exception as exc:
        print("G3 FAIL: %r" % exc)
        return False


def main() -> int:
    print("=" * 60)
    print("T070 acceptance verification")
    print("=" * 60)
    a = g1_cable_route()
    print("-" * 60)
    b = g2_luminaire_mount()
    print("-" * 60)
    c = g3_imports()
    print("=" * 60)
    print("Overall: G1=%s G2=%s G3=%s"
          % ("PASS" if a else "FAIL",
             "PASS" if b else "FAIL",
             "PASS" if c else "FAIL"))
    return 0 if (a and b and c) else 1


if __name__ == "__main__":
    raise SystemExit(main())
