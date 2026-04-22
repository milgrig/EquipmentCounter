"""
Unit tests for T019 (S002-04): multi-page legend scanning in
pdf_legend_parser.parse_legend.

We cannot (and should not) depend on real PDFs for unit tests, so these
tests focus on the two surgical additions that make multi-page selection
work:

  * `_legend_density_score` — the scoring function that decides which
    candidate wins when multiple pages produce legends.
  * `parse_legend` — verified by monkey-patching pdfplumber.open with a
    fake multi-page PDF: we inject synthetic per-page LegendResult
    candidates via patched helpers and confirm that the densest one is
    returned regardless of page order.
"""

from __future__ import annotations

import unittest
from unittest import mock

import pdf_legend_parser as plp
from pdf_legend_parser import LegendItem, LegendResult


def _item(symbol: str = "1", desc: str = "desc") -> LegendItem:
    return LegendItem(symbol=symbol, description=desc)


class DensityScoreTest(unittest.TestCase):

    def test_empty_result_is_zero(self):
        self.assertEqual(plp._legend_density_score(LegendResult()), 0.0)

    def test_more_symbols_beats_fewer(self):
        small = LegendResult(items=[_item("1"), _item("2")])
        big = LegendResult(items=[_item(str(i)) for i in range(1, 11)])
        self.assertGreater(
            plp._legend_density_score(big),
            plp._legend_density_score(small),
        )

    def test_symbols_dominate_over_total(self):
        # A result with 3 symbolled items should beat one with 10 items
        # but no symbols.
        symbolled = LegendResult(items=[_item("1"), _item("2"), _item("3")])
        no_symbols = LegendResult(items=[_item("", "d") for _ in range(10)])
        self.assertGreater(
            plp._legend_density_score(symbolled),
            plp._legend_density_score(no_symbols),
        )

    def test_tiebreak_by_total_then_description(self):
        a = LegendResult(items=[_item("1", "one"), _item("2", "two")])
        b = LegendResult(items=[_item("1", ""), _item("2", "")])
        # Same symbol count and total, but `a` has more descriptions.
        self.assertGreater(
            plp._legend_density_score(a),
            plp._legend_density_score(b),
        )


class ParseLegendMultiPageTest(unittest.TestCase):
    """Verify parse_legend picks the densest candidate across all pages."""

    def test_picks_page2_when_it_has_more_symbols(self):
        # Simulate a 3-page PDF.  Page 0 has a tiny legend (1 item),
        # page 1 has a rich legend (8 items), page 2 has nothing.
        page0_result = LegendResult(
            items=[_item("1", "only one")],
            page_index=0,
            legend_bbox=(0, 0, 10, 10),
        )
        page1_result = LegendResult(
            items=[_item(str(i), f"item {i}") for i in range(1, 9)],
            page_index=1,
            legend_bbox=(0, 0, 20, 20),
        )
        # page 2 -> no candidate produced

        # Build fake pages and a fake pdfplumber context manager.
        fake_pages = [mock.MagicMock(name=f"page{i}") for i in range(3)]
        for i, p in enumerate(fake_pages):
            p.lines = []
            p.rects = []
            p.width = 595
            p.height = 842
            # Return empty words so the real header pass finds nothing.
            p.extract_words = mock.MagicMock(return_value=[])

        fake_pdf = mock.MagicMock()
        fake_pdf.pages = fake_pages
        fake_pdf.__enter__ = mock.MagicMock(return_value=fake_pdf)
        fake_pdf.__exit__ = mock.MagicMock(return_value=False)

        # We patch _content_based_legend_search so that:
        #   page 0 -> returns page0_result (small legend)
        #   page 1 -> returns page1_result (big legend)
        #   page 2 -> returns None
        def fake_content_search(page, page_idx):
            if page_idx == 0:
                return page0_result
            if page_idx == 1:
                return page1_result
            return None

        # Also suppress _extract_with_tolerance so header pass does nothing.
        def fake_extract_with_tolerance(page, page_idx, x_tol):
            return ([], None)

        with mock.patch("pdf_legend_parser.pdfplumber.open", return_value=fake_pdf), \
             mock.patch(
                 "pdf_legend_parser._content_based_legend_search",
                 side_effect=fake_content_search,
             ), \
             mock.patch(
                 "pdf_legend_parser._extract_with_tolerance",
                 side_effect=fake_extract_with_tolerance,
             ):
            result = plp.parse_legend("fake.pdf")

        self.assertIsNotNone(result)
        self.assertEqual(
            result.page_index, 1,
            f"Expected page 1 (dense legend), got page {result.page_index}",
        )
        self.assertEqual(len(result.items), 8)

    def test_returns_empty_when_no_pages_have_legends(self):
        fake_pages = [mock.MagicMock(name=f"page{i}") for i in range(2)]
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

        with mock.patch("pdf_legend_parser.pdfplumber.open", return_value=fake_pdf), \
             mock.patch(
                 "pdf_legend_parser._content_based_legend_search",
                 return_value=None,
             ), \
             mock.patch(
                 "pdf_legend_parser._extract_with_tolerance",
                 return_value=([], None),
             ):
            result = plp.parse_legend("fake.pdf")

        self.assertEqual(result.items, [])
        self.assertEqual(result.page_index, 0)  # default

    def test_tiebreak_prefers_earlier_page(self):
        # Two pages both produce identical-density legends -> earlier wins
        # so that single-legend PDFs keep their legacy behaviour.
        item_list = [_item(str(i)) for i in range(1, 6)]
        page0 = LegendResult(items=list(item_list), page_index=0)
        page1 = LegendResult(items=list(item_list), page_index=1)

        fake_pages = [mock.MagicMock(name=f"page{i}") for i in range(2)]
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
            return page0 if page_idx == 0 else page1

        with mock.patch("pdf_legend_parser.pdfplumber.open", return_value=fake_pdf), \
             mock.patch(
                 "pdf_legend_parser._content_based_legend_search",
                 side_effect=fake_content_search,
             ), \
             mock.patch(
                 "pdf_legend_parser._extract_with_tolerance",
                 return_value=([], None),
             ):
            result = plp.parse_legend("fake.pdf")

        self.assertEqual(result.page_index, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
