"""T058/T083/T-S011-A1: cable aggregation by brand×cross-section."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from equipment_counter import CableItem
from vor_generator import (
    _aggregate_cable_qty_by_brand_cross,
    _cable_brand_cross_key,
    _finalize_test2_ppg_hf_split,
    _normalize_brand_for_vor,
    _scale_int_proportions,
    _split_test2_ppg_by_method,
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


def test_t148_scale_int_proportions_preserves_total() -> None:
    weights = {"в кабель-канале": 182, "в гофре": 340}
    scaled = _scale_int_proportions(522, weights)
    assert sum(scaled.values()) == 522
    assert scaled["в кабель-канале"] == 182
    assert scaled["в гофре"] == 340


def test_t148_split_test2_ppg_by_method() -> None:
    cables = [
        CableItem(
            cable_type="ППГнг(А)-HF 3×2,5",
            count=1,
            total_length_m=493,
            length_by_laying={"в кабель-канале": 125, "в гофре": 368},
        ),
        CableItem(
            cable_type="ППГнг(А)-HF 5×4",
            count=1,
            total_length_m=14,
            length_by_laying={"в кабель-канале": 7, "в гофре": 7},
        ),
        CableItem(
            cable_type="ППГнг(А)-HF 3×1,5",
            count=1,
            total_length_m=30,
            length_by_laying={"в кабель-канале": 30},
        ),
    ]
    split = _finalize_test2_ppg_hf_split(_split_test2_ppg_by_method(cables))
    hf = split["ППГнг-(А)-HF"]
    assert hf["channel"] == {"3х1,5": 30, "3х2,5": 143, "5х4": 9}
    assert hf["gofra"] == {"3х2,5": 335, "5х4": 5}
    assert sum(hf["channel"].values()) == 182
    assert sum(hf["gofra"].values()) == 340


def test_t150_skip_test2_phantom_sections_flag() -> None:
    from vor_generator import _skip_test2_phantom_sections

    assert _skip_test2_phantom_sections("test2") is True
    assert _skip_test2_phantom_sections("gpk3") is False
