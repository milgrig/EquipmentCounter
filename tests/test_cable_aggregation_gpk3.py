"""T083: gpk3 cable totals per brand×cross-section (T081 baseline)."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from equipment_counter import CableItem
from vor_generator import _aggregate_cable_qty_by_brand_cross

# From .tayfa/common/discussions/T081_baseline.md — gpk3 cable totals
GPK3_CABLE_BASELINE: tuple[tuple[str, str, int], ...] = (
    ("ВБШвнг(А)-FRLS", "3х1,5", 4460),
    ("ВБШвнг(А)-FRLS", "3х2,5", 180),
    ("ВБШвнг(А)-LS", "3х1,5", 2317),
    ("ВБШвнг(А)-LS", "3х2,5", 1127),
    ("ППГнг-(А)-FRHF", "3х1,5", 1300),
    ("ППГнг-(А)-HF", "3х1,5", 842),
)


def _cables_from_baseline() -> list[CableItem]:
    return [
        CableItem(
            cable_type=f"{brand} {cross.replace('х', '×')}",
            count=1,
            total_length_m=qty_m,
        )
        for brand, cross, qty_m in GPK3_CABLE_BASELINE
    ]


def test_t083_gpk3_baseline_row_per_brand_cross() -> None:
    totals = _aggregate_cable_qty_by_brand_cross(_cables_from_baseline())
    assert len(totals) == len(GPK3_CABLE_BASELINE)
    for brand, cross, qty_m in GPK3_CABLE_BASELINE:
        norm_brand = brand.replace("ППГнг(A)", "ППГнг-(А)")
        assert totals[(norm_brand, cross)] == qty_m


def test_t083_merges_duplicate_cable_types_by_key() -> None:
    cables = [
        CableItem("ВБШвнг(А)-LS 3×1,5", count=2, total_length_m=1000),
        CableItem("ВБШвнг(А)-LS 3х1,5", count=1, total_length_m=317),
        CableItem("ВБШвнг(А)-LS 3×1,5", count=1, total_length_m=1000),
    ]
    totals = _aggregate_cable_qty_by_brand_cross(cables)
    assert totals[("ВБШвнг(А)-LS", "3х1,5")] == 2317
