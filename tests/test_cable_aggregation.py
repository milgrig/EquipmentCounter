"""T058/T083: cable aggregation by brand×cross-section."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from equipment_counter import CableItem
from vor_generator import (
    _aggregate_cable_qty_by_brand_cross,
    _cable_brand_cross_key,
    _normalize_brand_for_vor,
)


def test_t058_aggregates_brand_and_cross_section() -> None:
    cables = [
        CableItem(
            cable_type="ППГнг(А)-HF 3×1.5",
            count=10,
            total_length_m=120,
            length_by_laying={"в гофре": 80, "в кабель-канале": 40},
        ),
        CableItem(
            cable_type="ППГнг-(А)-HF 3x1,5",
            count=3,
            total_length_m=30,
            length_by_laying={"в лотке": 30},
        ),
    ]
    totals = _aggregate_cable_qty_by_brand_cross(cables)
    assert len(totals) == 1
    key = ("ППГнг-(А)-HF", "3х1,5")
    assert key in totals
    assert totals[key] == 150


def test_t058_canonical_key_normalizes_ppg_brand() -> None:
    a = _cable_brand_cross_key("ППГнг(А)-HF 3×2,5")
    b = _cable_brand_cross_key("ППГнг-(А)-HF 3х2,5")
    assert a == b
    assert _normalize_brand_for_vor(a[0]) == "ППГнг-(А)-HF"
