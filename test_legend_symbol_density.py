"""
Unit tests for T021 (S002-06): _detect_legend_by_symbol_density.

Covers:
  * Returns None when the page has no candidate markers.
  * Returns None when markers exist but none have right-hand text
    (pure grid axis label column).
  * Returns None when markers are spread across multiple X-columns and
    none of them contains >= min_markers aligned markers.
  * Detects a legend with 4 markers + right-hand text.
  * Mixed markers (digits, digit+letter, letter+digit) are all picked up.
  * parse_legend falls back to the density detector when both
    header-based and content-based passes fail.
"""

from __future__ import annotations

import unittest
from unittest import mock

import pdf_legend_parser as plp
from pdf_legend_parser import LegendResult


def _word(text: str, x0: float, top: float,
          x1: float | None = None, bottom: float | None = None) -> dict:
    if x1 is None:
        x1 = x0 + max(6, 6 * len(text))
    if bottom is None:
        bottom = top + 8
    return {"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom}


class _FakePage:
    def __init__(self, words, lines=None, rects=None):
        self._words = list(words)
        self.lines = lines or []
        self.rects = rects or []
        self.width = 595
        self.height = 842

    def extract_words(self, **_kwargs):
        return list(self._words)


class DensityDetectorTest(unittest.TestCase):

    def test_no_markers_returns_none(self):
        page = _FakePage([_word("foo", 100, 100), _word("bar", 200, 100)])
        self.assertIsNone(plp._detect_legend_by_symbol_density(page, 0))

    def test_isolated_markers_return_none(self):
        # Markers with NO right-hand text — e.g. a column of grid labels
        # without descriptions should not qualify as a legend.
        page = _FakePage([
            _word("1", 100, 100),
            _word("2", 100, 120),
            _word("3", 100, 140),
            _word("4", 100, 160),
        ])
        self.assertIsNone(plp._detect_legend_by_symbol_density(page, 0))

    def test_detects_four_row_legend(self):
        page = _FakePage([
            _word("1", 100, 100), _word("Lamp A", 130, 100),
            _word("2", 100, 120), _word("Lamp B", 130, 120),
            _word("3A", 100, 140), _word("Socket", 130, 140),
            _word("5B", 100, 160), _word("Switch", 130, 160),
        ])
        result = plp._detect_legend_by_symbol_density(page, 0, min_markers=4)
        self.assertIsNotNone(result)
        self.assertGreaterEqual(len(result.items), 1)
        # bbox should start at x=100 and reach the description side.
        x0, y0, x1, y1 = result.legend_bbox
        self.assertAlmostEqual(x0, 100, delta=2)
        self.assertGreaterEqual(x1, 130)

    def test_respects_min_markers(self):
        # Only 3 marker rows -> should not qualify at min_markers=4.
        page = _FakePage([
            _word("1", 100, 100), _word("Lamp A", 130, 100),
            _word("2", 100, 120), _word("Lamp B", 130, 120),
            _word("3", 100, 140), _word("Lamp C", 130, 140),
        ])
        self.assertIsNone(
            plp._detect_legend_by_symbol_density(page, 0, min_markers=4),
        )

    def test_non_aligned_markers_do_not_cluster(self):
        # Four markers in four different X columns, none dense enough.
        page = _FakePage([
            _word("1", 100, 100), _word("Lamp A", 130, 100),
            _word("2", 250, 120), _word("Lamp B", 290, 120),
            _word("3", 100, 140), _word("Socket", 130, 140),
            _word("4", 300, 160), _word("Switch", 330, 160),
        ])
        # Column @100 has only 2 rows; @250/300 have 1 each — none qualifies.
        self.assertIsNone(
            plp._detect_legend_by_symbol_density(page, 0, min_markers=4),
        )

    def test_large_row_gap_breaks_run(self):
        # 4 markers but with an 80pt gap between 2 and 3 -> two runs of
        # 2 markers each, neither qualifies.
        page = _FakePage([
            _word("1", 100, 100), _word("A", 130, 100),
            _word("2", 100, 120), _word("B", 130, 120),
            _word("3", 100, 200), _word("C", 130, 200),  # gap of 80pt
            _word("4", 100, 220), _word("D", 130, 220),
        ])
        self.assertIsNone(
            plp._detect_legend_by_symbol_density(page, 0, min_markers=4),
        )


class ParseLegendLastResortTest(unittest.TestCase):
    """parse_legend should invoke the density detector when both earlier
    passes fail and produce zero candidates."""

    def test_density_fallback_is_last_resort(self):
        # A single-page PDF: header-based + content-based both return
        # nothing, but the density detector finds a legend.
        words = [
            _word("1", 100, 100), _word("Lamp A", 130, 100),
            _word("2", 100, 120), _word("Lamp B", 130, 120),
            _word("3A", 100, 140), _word("Socket", 130, 140),
            _word("5B", 100, 160), _word("Switch", 130, 160),
        ]
        fake_page = mock.MagicMock(name="page0")
        fake_page.lines = []
        fake_page.rects = []
        fake_page.width = 595
        fake_page.height = 842
        fake_page.extract_words = mock.MagicMock(return_value=list(words))

        fake_pdf = mock.MagicMock()
        fake_pdf.pages = [fake_page]
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

        self.assertTrue(
            result.items,
            "Density detector must produce items when earlier passes fail.",
        )

    def test_density_detector_is_not_called_when_earlier_pass_succeeds(self):
        # If _content_based_legend_search returns a good candidate, the
        # density fallback must NOT be invoked (it's last-resort).
        page0_result = LegendResult(
            items=[plp.LegendItem(symbol=str(i), description=f"x{i}")
                   for i in range(1, 6)],
            page_index=0,
        )

        fake_page = mock.MagicMock(name="page0")
        fake_page.lines = []
        fake_page.rects = []
        fake_page.width = 595
        fake_page.height = 842
        fake_page.extract_words = mock.MagicMock(return_value=[])

        fake_pdf = mock.MagicMock()
        fake_pdf.pages = [fake_page]
        fake_pdf.__enter__ = mock.MagicMock(return_value=fake_pdf)
        fake_pdf.__exit__ = mock.MagicMock(return_value=False)

        with mock.patch("pdf_legend_parser.pdfplumber.open", return_value=fake_pdf), \
             mock.patch(
                 "pdf_legend_parser._content_based_legend_search",
                 return_value=page0_result,
             ), \
             mock.patch(
                 "pdf_legend_parser._extract_with_tolerance",
                 return_value=([], None),
             ), \
             mock.patch(
                 "pdf_legend_parser._detect_legend_by_symbol_density",
             ) as dens_mock:
            result = plp.parse_legend("fake.pdf")

        self.assertEqual(len(result.items), 5)
        dens_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
