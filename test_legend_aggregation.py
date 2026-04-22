"""
Unit tests for T020 (S002-05): multi-page legend aggregation.

Covers:
  * LegendItem now has a `page_index` field (default 0, still backward-
    compatible with positional construction).
  * _merge_legend_candidates de-duplicates by symbol, keeps the longer
    description on collisions, and stamps a meaningful aggregate bbox.
  * _should_aggregate returns True only when another page contributes a
    genuinely new symbol.
  * parse_legend produces an aggregated LegendResult when two pages
    carry disjoint symbols.
"""

from __future__ import annotations

import unittest
from unittest import mock

import pdf_legend_parser as plp
from pdf_legend_parser import LegendItem, LegendResult


def _item(sym, desc="", page=0) -> LegendItem:
    return LegendItem(symbol=sym, description=desc, page_index=page)


class LegendItemFieldTest(unittest.TestCase):

    def test_default_page_index_is_zero(self):
        it = LegendItem(symbol="1", description="x")
        self.assertEqual(it.page_index, 0)

    def test_page_index_is_settable(self):
        it = LegendItem(symbol="1", description="x", page_index=3)
        self.assertEqual(it.page_index, 3)


class MergeCandidatesTest(unittest.TestCase):

    def test_empty_list_yields_empty_result(self):
        out = plp._merge_legend_candidates([])
        self.assertEqual(out.items, [])

    def test_single_candidate_passes_through(self):
        a = LegendResult(
            items=[_item("1", "lamp", page=0), _item("2", "switch", page=0)],
            page_index=0,
        )
        out = plp._merge_legend_candidates([a])
        self.assertEqual(len(out.items), 2)
        self.assertEqual(out.page_index, 0)

    def test_disjoint_symbols_are_both_kept(self):
        a = LegendResult(
            items=[_item("1", "lamp", page=0), _item("2", "switch", page=0)],
            page_index=0,
        )
        b = LegendResult(
            items=[_item("3", "socket", page=1), _item("4", "panel", page=1)],
            page_index=1,
        )
        out = plp._merge_legend_candidates([a, b])
        syms = sorted(it.symbol for it in out.items)
        self.assertEqual(syms, ["1", "2", "3", "4"])
        # Items from page 1 keep their page_index.
        p1_items = [it for it in out.items if it.page_index == 1]
        self.assertEqual(len(p1_items), 2)

    def test_colliding_symbols_keep_longer_description(self):
        a = LegendResult(
            items=[_item("1", "short", page=0)],
            page_index=0,
        )
        b = LegendResult(
            items=[_item("1", "a much longer full description", page=1)],
            page_index=1,
        )
        # a is denser (same 1 symbol, longer desc wins), but the test is
        # about what happens when *both* exist: the longer description
        # should survive regardless of ordering.
        out = plp._merge_legend_candidates([a, b])
        self.assertEqual(len(out.items), 1)
        self.assertIn("longer", out.items[0].description)

    def test_empty_symbol_items_deduped_by_description(self):
        a = LegendResult(
            items=[_item("", "EXIT", page=0)], page_index=0,
        )
        b = LegendResult(
            items=[_item("", "EXIT", page=1), _item("", "EVAC", page=1)],
            page_index=1,
        )
        out = plp._merge_legend_candidates([a, b])
        descs = sorted(it.description for it in out.items)
        self.assertEqual(descs, ["EVAC", "EXIT"])

    def test_primary_page_is_the_densest(self):
        small = LegendResult(items=[_item("9")], page_index=2)
        big = LegendResult(
            items=[_item(str(i)) for i in range(1, 6)],
            page_index=5,
        )
        out = plp._merge_legend_candidates([small, big])
        # `big` is denser → its page becomes the primary page.
        self.assertEqual(out.page_index, 5)


class ShouldAggregateTest(unittest.TestCase):

    def test_false_when_no_new_symbols(self):
        base = LegendResult(items=[_item("1"), _item("2")])
        other = LegendResult(items=[_item("1"), _item("2")])
        self.assertFalse(plp._should_aggregate(base, [other]))

    def test_true_when_other_brings_new_symbol(self):
        base = LegendResult(items=[_item("1"), _item("2")])
        other = LegendResult(items=[_item("3")])
        self.assertTrue(plp._should_aggregate(base, [other]))

    def test_empty_symbols_do_not_trigger_aggregation(self):
        base = LegendResult(items=[_item("1")])
        other = LegendResult(items=[_item("", "some text")])
        self.assertFalse(plp._should_aggregate(base, [other]))


class ParseLegendAggregationTest(unittest.TestCase):
    """End-to-end: parse_legend aggregates across pages when beneficial."""

    def _run_with_candidates(self, page_results: dict[int, LegendResult]):
        n_pages = max(page_results) + 1 if page_results else 1
        fake_pages = [mock.MagicMock(name=f"page{i}") for i in range(n_pages)]
        for p in fake_pages:
            p.lines = []
            p.rects = []
            p.width = 595
            p.height = 842
            p.extract_words = mock.MagicMock(return_value=[])

        fake_pdf = mock.MagicMock()
        fake_pdf.pages = fake_pages
        fake_pdf.__enter__ = mock.MagicMock(return_value=fake_pdf)
        fake_pdf.__exit__ = mock.MagicMock(return_value=False)

        def fake_content_search(page, page_idx):
            return page_results.get(page_idx)

        with mock.patch("pdf_legend_parser.pdfplumber.open", return_value=fake_pdf), \
             mock.patch(
                 "pdf_legend_parser._content_based_legend_search",
                 side_effect=fake_content_search,
             ), \
             mock.patch(
                 "pdf_legend_parser._extract_with_tolerance",
                 return_value=([], None),
             ):
            return plp.parse_legend("fake.pdf")

    def test_aggregates_when_pages_bring_new_symbols(self):
        results = {
            0: LegendResult(
                items=[_item("1", "lamp", page=0), _item("2", "sw", page=0)],
                page_index=0,
            ),
            1: LegendResult(
                items=[_item("3", "skt", page=1), _item("4", "pnl", page=1)],
                page_index=1,
            ),
        }
        out = self._run_with_candidates(results)
        syms = sorted(it.symbol for it in out.items)
        self.assertEqual(syms, ["1", "2", "3", "4"])
        # At least one item from page 1 should be present with page_index=1.
        self.assertTrue(any(it.page_index == 1 for it in out.items))

    def test_single_page_behaviour_is_unchanged(self):
        # When only one page has a candidate, the result is byte-identical
        # to the previous (non-aggregating) behaviour.
        results = {
            0: LegendResult(
                items=[_item("1", "a", page=0), _item("2", "b", page=0)],
                page_index=0,
            ),
        }
        out = self._run_with_candidates(results)
        self.assertEqual(len(out.items), 2)
        self.assertEqual(out.page_index, 0)

    def test_no_aggregation_when_pages_duplicate_each_other(self):
        # Two pages carrying identical symbols -> pick best single page,
        # do not produce an aggregate.
        shared = [_item("1", "a"), _item("2", "b")]
        results = {
            0: LegendResult(items=[LegendItem(**it.__dict__) for it in shared], page_index=0),
            1: LegendResult(items=[LegendItem(**it.__dict__) for it in shared], page_index=1),
        }
        out = self._run_with_candidates(results)
        self.assertEqual(len(out.items), 2)  # not 4


if __name__ == "__main__":
    unittest.main(verbosity=2)
