"""T058/T083: cable aggregation by brand×cross-section."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from equipment_counter import CableItem
from vor_generator import (
    _aggregate_cable_qty_by_brand_cross,
    _cable_brand_cross_key,
    _format_cable_material_desc,
)


def test_t058_test2_mini_aggregates_brand_and_section() -> None:
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
    assert totals[key] == 150


def test_t083_brand_cross_key_normalizes_ppgng_dash() -> None:
    a = _cable_brand_cross_key("ППГнг(А)-HF 3×2,5")
    b = _cable_brand_cross_key("ППГнг-(А)-HF 3x2,5")
    assert a == b
    assert a[0] == "ППГнг-(А)-HF"
    assert a[1] == "3х2,5"


def test_t083_material_desc_power_cable_etalon_style() -> None:
    desc = _format_cable_material_desc("ВБШвнг(А)-LS 3×1,5")
    assert desc == "Кабель силовой с медными жилами ВБШвнг(А)-LS 3х1,5"


def test_t083_merge_same_brand_cross_different_type_strings() -> None:
    cables = [
        CableItem("ВБШвнг(А)-LS 3х1,5", 1, 1000),
        CableItem("ВБШвнг(А)-LS 3x1.5", 1, 317),
    ]
    totals = _aggregate_cable_qty_by_brand_cross(cables)
    assert len(totals) == 1
    assert totals[("ВБШвнг(А)-LS", "3х1,5")] == 1317
