"""T083: gpk3 etalon cable totals by brand×cross-section (T081 baseline)."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from equipment_counter import CableItem
from vor_generator import (
    _aggregate_cable_qty_by_brand_cross,
    _format_cable_material_desc,
)

# From .tayfa/common/discussions/T081_baseline.md — gpk3 cable totals
GPK3_ETALON_CABLE_TOTALS: dict[tuple[str, str], int] = {
    ("ВБШвнг(А)-FRLS", "3х1,5"): 4460,
    ("ВБШвнг(А)-FRLS", "3х2,5"): 180,
    ("ВБШвнг(А)-LS", "3х1,5"): 2317,
    ("ВБШвнг(А)-LS", "3х2,5"): 1127,
    ("ППГнг-(А)-FRHF", "3х1,5"): 1300,
    ("ППГнг-(А)-HF", "3х1,5"): 842,
}


def _cables_from_etalon_totals() -> list[CableItem]:
    return [
        CableItem(
            cable_type=f"{brand} {cross}",
            count=1,
            total_length_m=qty_m,
        )
        for (brand, cross), qty_m in GPK3_ETALON_CABLE_TOTALS.items()
    ]


def test_t083_gpk3_aggregate_matches_etalon_per_brand_cross() -> None:
    totals = _aggregate_cable_qty_by_brand_cross(_cables_from_etalon_totals())
    assert totals == GPK3_ETALON_CABLE_TOTALS


def test_t083_gpk3_vbshvng_ls_3x1_5() -> None:
    cables = [
        CableItem("ВБШвнг(А)-LS 3х1,5", 1, 2000),
        CableItem("ВБШвнг(А)-LS 3x1.5", 1, 317),
    ]
    assert _aggregate_cable_qty_by_brand_cross(cables)[("ВБШвнг(А)-LS", "3х1,5")] == 2317


def test_t083_gpk3_vbshvng_frls_3x2_5() -> None:
    cables = [CableItem("ВБШвнг(А)-FRLS 3х2,5", 2, 180)]
    assert _aggregate_cable_qty_by_brand_cross(cables)[("ВБШвнг(А)-FRLS", "3х2,5")] == 180


def test_t083_gpk3_ppng_frhf_3x1_5() -> None:
    cables = [CableItem("ППГнг(А)-FRHF 3х1,5", 1, 1300)]
    assert _aggregate_cable_qty_by_brand_cross(cables)[("ППГнг-(А)-FRHF", "3х1,5")] == 1300


def test_t083_gpk3_ppng_hf_3x1_5() -> None:
    cables = [CableItem("ППГнг(A)-HF 3x1,5", 1, 842)]
    assert _aggregate_cable_qty_by_brand_cross(cables)[("ППГнг-(А)-HF", "3х1,5")] == 842


def test_t083_gpk3_material_desc_matches_compare_synonyms() -> None:
    # Merge S011 (07f7993d): формат материала — «Кабель <марка> сечением
    # <сечение>» (валидирован сверками с СО/эталонами июля). Июньская
    # T083-формулировка «Кабель силовой с медными жилами …» осталась в
    # истории main; возможный возврат к ней — отдельное решение с
    # ревалидацией фьюзи-матчинга эталонов.
    desc = _format_cable_material_desc("ВБШвнг(А)-FRLS 3х1,5")
    assert "Кабель" in desc
    assert "ВБШвнг(А)-FRLS" in desc
    assert "3х1,5" in desc


def test_t083_gpk3_ppng_material_desc_uses_latin_a() -> None:
    # См. комментарий выше: актуальный формат нормализует марку как
    # «ППГнг-(А)-HF» (кириллическая А с дефисом), а не латинскую «(A)».
    desc = _format_cable_material_desc("ППГнг(А)-HF 3х1,5")
    assert "ППГнг-(А)-HF" in desc
    assert "сечением 3х1,5" in desc
