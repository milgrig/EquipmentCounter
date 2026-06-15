"""Regression test for T-S011-B2 (T151): dataset-aware section header
taxonomy and PVC phantom suppression.

Background
----------
gpk3 etalon ВОР uses different section header wording than test2 etalon:

    test2:                      gpk3:
    ------------------------    --------------------------------
    Монтаж светильников и ламп  Светотехническое оборудование
    Монтаж ПВХ изделий и труб   ПВХ изделия и трубы

Before T151 the generator emitted test2-style headers unconditionally.
On gpk3 this created phantom sections — rows under our header had no
etalon counterpart and were bucketed as ONLY_OURS even when the
underlying row content paired by name.  Baseline gpk3 had 12 ONLY_OURS
rows across these two phantom sections (10 PVC + 2 light).

This test pins the mapping helper so a future refactor cannot silently
reintroduce the phantom by reverting the dataset switch or by adding a
test2 alias that would regress KB-008.
"""

import importlib
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

vor_generator = importlib.import_module("vor_generator")


GPK3_ALIASES = {
    "Монтаж светильников и ламп": "Светотехническое оборудование",
    "Монтаж ПВХ изделий и труб":  "ПВХ изделия и трубы",
}


class TestSectionHeaderTaxonomy:
    def test_gpk3_aliases_match_etalon_titles(self):
        for src, expected in GPK3_ALIASES.items():
            got = vor_generator._section_header_for_dataset(src, "gpk3")
            assert got == expected, (
                f"gpk3 header {src!r} must map to etalon {expected!r}, got {got!r}"
            )

    def test_test2_keeps_legacy_headers(self):
        for src in GPK3_ALIASES:
            got = vor_generator._section_header_for_dataset(src, "test2")
            assert got == src, (
                f"test2 header {src!r} must NOT change (KB-008), got {got!r}"
            )

    def test_none_dataset_keeps_legacy_headers(self):
        for src in GPK3_ALIASES:
            got = vor_generator._section_header_for_dataset(src, None)
            assert got == src

    def test_unmapped_header_passes_through_on_gpk3(self):
        got = vor_generator._section_header_for_dataset(
            "Щитовое оборудование", "gpk3",
        )
        assert got == "Щитовое оборудование"


class TestPvcMaterialSubrowSuppression:
    """The per-height 'Труба ПВХ гибкая гофр. д.XXмм' material sub-rows
    have no counterpart in the gpk3 etalon ПВХ section (which carries
    only the diameter-tagged work rows).  Emitting them there produced
    pure ONLY_OURS phantom noise."""

    def test_suppressed_on_gpk3(self):
        assert vor_generator._emit_pvc_material_subrows("gpk3") is False

    def test_kept_on_test2(self):
        # Test2 etalon carries 'Труба ПА 6 гибкая гофр. д.XXмм' material
        # rows that pair against our sub-rows; KB-008 protects them.
        assert vor_generator._emit_pvc_material_subrows("test2") is True

    def test_kept_when_dataset_omitted(self):
        assert vor_generator._emit_pvc_material_subrows(None) is True
        assert vor_generator._emit_pvc_material_subrows("") is True


class TestAggregateByHeightCallable:
    """Regression guard for the T-S011-B2 follow-up NameError fix.

    The pictogram spec→plan enrichment path in ``aggregate_by_height``
    calls a *local* helper ``_apply_spec_qty_to_indicator``.  A typo at
    two call sites (``apply_spec_qty_to_indicator`` without the leading
    underscore) raised ``NameError`` whenever a pictogram-class spec
    item was processed (every gpk3 run), short-circuiting both the
    section-alias rename and the PVC sub-row gate.  This test pins the
    module compiling and ``aggregate_by_height`` running on an empty
    input without raising — i.e. the local helper resolves.
    """

    def test_aggregate_by_height_does_not_raise_nameerror(self):
        agg = vor_generator.aggregate_by_height([], log=lambda *_: None)
        assert isinstance(agg, dict)
        assert "luminaires" in agg


class TestPhantomGateBookkeeping:
    """Regression guard for the T-S011-B2 gate.

    Phantom-named sections must not be emitted on gpk3 so that the
    `gpk3 phantom ONLY_OURS <= 3` gate stays satisfied without relying
    on row-content fuzz score volatility.
    """

    @pytest.mark.parametrize(
        "src,gpk3,test2",
        list(GPK3_ALIASES.items()) and [
            ("Монтаж светильников и ламп",
             "Светотехническое оборудование",
             "Монтаж светильников и ламп"),
            ("Монтаж ПВХ изделий и труб",
             "ПВХ изделия и трубы",
             "Монтаж ПВХ изделий и труб"),
        ],
    )
    def test_phantom_header_renamed_on_gpk3_only(self, src, gpk3, test2):
        assert vor_generator._section_header_for_dataset(src, "gpk3") == gpk3
        assert vor_generator._section_header_for_dataset(src, "test2") == test2
        assert vor_generator._section_header_for_dataset(src, "test2") != gpk3
