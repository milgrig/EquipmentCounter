"""Unit tests for T022 (S002-07): _spec_table_as_legend."""

from __future__ import annotations

import unittest

import pdf_legend_parser as plp


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


class SpecTableDetectorTest(unittest.TestCase):

    def test_no_words_returns_none(self):
        self.assertIsNone(plp._spec_table_as_legend(_FakePage([]), 0))

    def test_returns_none_without_poz_header(self):
        # Header row contains Наименование but no Поз./Поз
        page = _FakePage([
            _word("Наименование", 100, 100),
            _word("Кол.", 300, 100),
            _word("Anchor", 100, 120),
            _word("10", 300, 120),
        ])
        self.assertIsNone(plp._spec_table_as_legend(page, 0))

    def test_detects_basic_poz_nam_kol_table(self):
        page = _FakePage([
            # header row
            _word("Поз.", 100, 100, x1=130),
            _word("Наименование", 150, 100, x1=260),
            _word("Кол.", 300, 100, x1=330),
            # 3 data rows
            _word("1", 100, 130), _word("Anchor M10", 150, 130, x1=260), _word("5", 300, 130),
            _word("2", 100, 150), _word("Bracket 200mm", 150, 150, x1=270), _word("3", 300, 150),
            _word("3", 100, 170), _word("Cable tray", 150, 170, x1=260), _word("10", 300, 170),
        ])
        result = plp._spec_table_as_legend(page, 0)
        self.assertIsNotNone(result)
        self.assertEqual(len(result.items), 3)
        self.assertEqual(result.items[0].symbol, "1")
        self.assertIn("Anchor", result.items[0].description)
        # page_index is stamped
        for it in result.items:
            self.assertEqual(it.page_index, 0)

    def test_requires_min_three_rows(self):
        page = _FakePage([
            _word("Поз.", 100, 100, x1=130),
            _word("Наименование", 150, 100, x1=260),
            _word("1", 100, 130), _word("Anchor", 150, 130),
            _word("2", 100, 150), _word("Bracket", 150, 150),
        ])
        self.assertIsNone(plp._spec_table_as_legend(page, 0))

    def test_large_y_gap_breaks_table(self):
        # 3 rows, then a 120pt gap, then something else — detector must
        # stop at the gap and still find the first table.
        page = _FakePage([
            _word("Поз.", 100, 100, x1=130),
            _word("Наименование", 150, 100, x1=260),
            _word("1", 100, 130), _word("Anchor", 150, 130),
            _word("2", 100, 150), _word("Bracket", 150, 150),
            _word("3", 100, 170), _word("Cable", 150, 170),
            # Big gap:
            _word("X", 100, 350), _word("unrelated", 150, 350),
        ])
        result = plp._spec_table_as_legend(page, 0)
        self.assertIsNotNone(result)
        self.assertEqual(len(result.items), 3)
        syms = [it.symbol for it in result.items]
        self.assertNotIn("X", syms)

    def test_picks_multi_table_page_with_most_items(self):
        # Two header rows on the same page; the second one has more data.
        words = [
            # Table A: 3 rows
            _word("Поз.", 50, 100, x1=80),
            _word("Наименование", 100, 100, x1=210),
            _word("1", 50, 130), _word("A1", 100, 130),
            _word("2", 50, 150), _word("A2", 100, 150),
            _word("3", 50, 170), _word("A3", 100, 170),
            # Big gap so first table ends
            # Table B: 5 rows, starting at y=400
            _word("Поз.", 50, 400, x1=80),
            _word("Наименование", 100, 400, x1=210),
            _word("1", 50, 430), _word("B1", 100, 430),
            _word("2", 50, 450), _word("B2", 100, 450),
            _word("3", 50, 470), _word("B3", 100, 470),
            _word("4", 50, 490), _word("B4", 100, 490),
            _word("5", 50, 510), _word("B5", 100, 510),
        ]
        page = _FakePage(words)
        result = plp._spec_table_as_legend(page, 0)
        self.assertIsNotNone(result)
        # The 5-row table should win.
        self.assertEqual(len(result.items), 5)
        descs = " ".join(it.description for it in result.items)
        self.assertIn("B", descs)


if __name__ == "__main__":
    unittest.main(verbosity=2)
