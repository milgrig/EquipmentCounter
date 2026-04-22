"""
Unit tests for pdf_legend_parser._content_based_legend_search and the
new _find_identifier_cluster helper introduced in T018 (S002-03).

Goals:
  * ROW_IDENT_RE matches the expected legend-identifier shapes and
    rejects clearly non-identifier tokens.
  * _find_identifier_cluster returns a plausible bbox when given a stack
    of >= 4 rows whose leftmost token is an identifier.
  * _content_based_legend_search does NOT early-return when a page has
    fewer than 3 horizontal lines — it falls through to the cluster
    pass and produces a LegendResult from word content alone.
"""

from __future__ import annotations

import unittest

import pdf_legend_parser as plp


def _word(text: str, x0: float, top: float, x1: float | None = None,
          bottom: float | None = None) -> dict:
    if x1 is None:
        x1 = x0 + 6 * len(text)
    if bottom is None:
        bottom = top + 8
    return {"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom}


class _FakePage:
    """Minimal stand-in for a pdfplumber.Page for _content_based_legend_search."""

    def __init__(self, words, lines=None, rects=None):
        self._words = words
        self.lines = lines or []
        self.rects = rects or []

    def extract_words(self, **_kwargs):
        return list(self._words)


class RowIdentRegexTest(unittest.TestCase):

    def test_digit_only_identifiers_match(self):
        for tok in ["1", "5", "9", "12", "99"]:
            self.assertTrue(
                plp.ROW_IDENT_RE.match(tok), f"{tok!r} should match ROW_IDENT_RE"
            )

    def test_letter_plus_digit_match(self):
        # Mixed: digit-letter and letter-digit, Latin and Cyrillic
        for tok in ["5A", "2B", "10A", "5\u0410", "\u04102", "\u04121"]:
            self.assertTrue(
                plp.ROW_IDENT_RE.match(tok), f"{tok!r} should match ROW_IDENT_RE"
            )

    def test_rejects_non_identifiers(self):
        for tok in ["abc", "100", "1.5", "-1", ""]:
            self.assertFalse(
                plp.ROW_IDENT_RE.match(tok), f"{tok!r} should NOT match ROW_IDENT_RE"
            )


class FindIdentifierClusterTest(unittest.TestCase):

    def test_returns_none_for_empty(self):
        self.assertIsNone(plp._find_identifier_cluster([]))

    def test_detects_cluster_of_four_rows(self):
        # Four rows of the form "<ID> <description>"
        words = [
            _word("1", 100, 200), _word("Svet-A", 130, 200),
            _word("2", 100, 220), _word("Svet-B", 130, 220),
            _word("3", 100, 240), _word("Socket",  130, 240),
            _word("4", 100, 260), _word("Switch",  130, 260),
        ]
        result = plp._find_identifier_cluster(words, min_rows=4)
        self.assertIsNotNone(result)
        x0, y0, x1, y1, row_tops = result
        self.assertAlmostEqual(x0, 100, delta=1)
        self.assertGreaterEqual(x1, 130)
        self.assertAlmostEqual(y0, 200, delta=1)
        self.assertGreaterEqual(y1, 260)
        self.assertEqual(len(row_tops), 4)

    def test_requires_min_rows(self):
        words = [
            _word("1", 100, 200),
            _word("2", 100, 220),
            _word("3", 100, 240),
        ]
        self.assertIsNone(plp._find_identifier_cluster(words, min_rows=4))

    def test_letter_prefixed_ids_are_clustered(self):
        words = [
            _word("5A", 50, 100), _word("Cable1", 80, 100),
            _word("5B", 50, 115), _word("Cable2", 80, 115),
            _word("5C", 50, 130), _word("Cable3", 80, 130),
            _word("5D", 50, 145), _word("Cable4", 80, 145),
        ]
        result = plp._find_identifier_cluster(words, min_rows=4)
        self.assertIsNotNone(result)

    def test_non_aligned_ids_do_not_cluster(self):
        # IDs wildly misaligned in X should NOT form a cluster.
        words = [
            _word("1", 100, 200), _word("X", 150, 200),
            _word("2", 300, 220), _word("Y", 340, 220),
            _word("3", 100, 240), _word("Z", 150, 240),
            _word("4", 300, 260), _word("W", 340, 260),
        ]
        self.assertIsNone(plp._find_identifier_cluster(words, min_rows=4))


class ContentBasedSearchFallbackTest(unittest.TestCase):
    """Verify the early-return guard is gone and the cluster path runs."""

    def test_no_early_return_on_few_lines(self):
        # Build a page with ZERO lines but a valid identifier cluster that
        # includes the equipment keyword "svetilnik" (Russian: lamp).
        words = [
            _word("1", 100, 200), _word("svetilnik LPO", 130, 200, x1=230),
            _word("2", 100, 220), _word("svetilnik LED", 130, 220, x1=230),
            _word("3", 100, 240), _word("rozetka",       130, 240, x1=200),
            _word("4", 100, 260), _word("kabel",         130, 260, x1=200),
        ]
        # Patch the equipment keyword list to include our Latin stand-ins,
        # then restore after the test.
        original_kw = list(plp._EQUIPMENT_KEYWORDS_LOWER)
        plp._EQUIPMENT_KEYWORDS_LOWER[:] = original_kw + [
            "svetilnik", "rozetka", "kabel",
        ]
        try:
            page = _FakePage(words=words, lines=[], rects=[])
            result = plp._content_based_legend_search(page, page_idx=0)
        finally:
            plp._EQUIPMENT_KEYWORDS_LOWER[:] = original_kw

        self.assertIsNotNone(
            result,
            "With the new cluster fallback, a page with no lines but a "
            "valid identifier cluster should still return a LegendResult.",
        )
        # The returned bbox should roughly enclose the four rows.
        x0, y0, x1, y1 = result.legend_bbox
        self.assertAlmostEqual(x0, 100, delta=2)
        self.assertAlmostEqual(y0, 200, delta=2)

    def test_returns_none_without_equipment_keyword(self):
        # Valid cluster shape but no equipment vocabulary -> rejected.
        words = [
            _word("1", 100, 200), _word("foo", 130, 200),
            _word("2", 100, 220), _word("bar", 130, 220),
            _word("3", 100, 240), _word("baz", 130, 240),
            _word("4", 100, 260), _word("qux", 130, 260),
        ]
        page = _FakePage(words=words, lines=[], rects=[])
        result = plp._content_based_legend_search(page, page_idx=0)
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main(verbosity=2)
