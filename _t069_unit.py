"""T069 unit-test the height_bucketer module on canonical filenames."""

import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import height_bucketer as hb


CASES = [
    # No "otm" marker -> by-design unknown (avoids false-positive on sheet
    # numbers like "004.1" that look numerically like elevations).
    ("005-something-0.000.pdf", "unknown"),
    ("005-Plany osvescheniya-otm. 0.000.pdf", "<5m"),
    ("006-otm. +4.200.pdf", "<5m"),
    ("007-otm. +7.800 +9.000.pdf", "5-13m"),
    ("008-otm. +13.800.pdf", "13-20m"),
    ("010-otm. +23.400.pdf", "20-35m"),
    ("011-otm. +28.200.pdf", "20-35m"),
    ("019-Lotki otm. 0.000.pdf", "<5m"),
    ("022-Lotki na otm. +13.800.pdf", "13-20m"),
    ("001-General data.pdf", "unknown"),
]


def main() -> int:
    fails = 0
    for path, expected in CASES:
        got = hb.bucket_for_path(path)
        elevs = hb.extract_elevations(path)
        flag = "OK" if got == expected else "FAIL"
        if got != expected:
            fails += 1
        print(f"{flag}  {path:<42s} elevs={elevs}  bucket={got}  expected={expected}")
    print(f"Failures: {fails}")
    print()

    items = [
        {"name": "Lamp X", "unit": "sht", "count": 5, "total": 5},
        {"name": "Lamp X", "unit": "sht", "count": 3, "total": 3},
        {"name": "Cable Y", "unit": "m", "count": 0, "total": 12.5},
    ]
    hb.attribute_items(items, "007-otm. +7.800 +9.000.pdf")
    print("Tagged height_bucket:", [i.get("height_bucket") for i in items])
    groups = hb.group_items_by_bucket(items)
    print("Groups:", len(groups))
    for k, v in groups.items():
        print(
            "  ", k, "->",
            {kk: vv for kk, vv in v.items()
             if kk in ("count", "total", "height_bucket")}
        )
    return fails


if __name__ == "__main__":
    raise SystemExit(main())
