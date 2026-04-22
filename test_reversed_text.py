"""Unit tests for S1.8 (T023) reversed-text detection.

Covers:
  * `_is_text_reversed` — detects pages whose Cyrillic tokens are
    character-reversed (common on CAD-exported PDFs).
  * `_reverse_cyrillic_words` — flips the ``text`` field of each word
    while preserving geometry.
  * Integration via `_extract_with_tolerance` — pages flagged as
    reversed get their words auto-flipped and the page is tagged so
    downstream detection_method picks up a "reversed:" prefix.
"""

import unittest
from unittest.mock import patch

from pdf_legend_parser import (
    _is_text_reversed,
    _reverse_cyrillic_words,
    _extract_with_tolerance,
    _tag_method,
    LegendResult,
    REVERSED_TEXT_STEMS,
)


def _word(text: str, x0: float = 0, top: float = 0,
          x1: float | None = None, bottom: float | None = None) -> dict:
    if x1 is None:
        x1 = x0 + max(6, 6 * len(text))
    if bottom is None:
        bottom = top + 8
    return {"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom}


class _FakePage:
    def __init__(self, words):
        self._words = list(words)
        self.lines = []
        self.rects = []
        self.width = 595
        self.height = 842

    def extract_words(self, **_kwargs):
        return [dict(w) for w in self._words]


class IsTextReversedTests(unittest.TestCase):

    def test_empty_input_returns_false(self):
        self.assertFalse(_is_text_reversed([]))

    def test_forward_text_is_not_flagged(self):
        words = [_word(t) for t in (
            "\u0417\u0430\u0449\u0438\u0442\u0430",        # Защита
            "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435",  # Наименование
            "\u041e\u0431\u043e\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435",  # Обозначение
            "\u0410\u0432\u0442\u043e\u043c\u0430\u0442",   # Автомат
            "\u041a\u0430\u0431\u0435\u043b\u044c",          # Кабель
        )]
        self.assertFalse(_is_text_reversed(words))

    def test_reversed_text_is_flagged(self):
        # Same words as above but each reversed character-by-character
        raw = [
            "\u0417\u0430\u0449\u0438\u0442\u0430",
            "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435",
            "\u041e\u0431\u043e\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435",
            "\u0410\u0432\u0442\u043e\u043c\u0430\u0442",
            "\u041a\u0430\u0431\u0435\u043b\u044c",
        ]
        words = [_word(t[::-1]) for t in raw]
        self.assertTrue(_is_text_reversed(words))

    def test_mixed_numbers_and_short_tokens_ignored(self):
        # Numeric / short tokens mustn't influence the decision.
        words = [_word(t) for t in ("1", "2A", "10", "\u0414\u0412", "+1.000", "104")]
        self.assertFalse(_is_text_reversed(words))

    def test_reference_stems_cover_domain_vocabulary(self):
        self.assertIn("\u0437\u0430\u0449\u0438\u0442\u0430", REVERSED_TEXT_STEMS)
        self.assertIn("\u0430\u0432\u0442\u043e\u043c\u0430\u0442", REVERSED_TEXT_STEMS)
        self.assertIn("\u043f\u043e\u0437", REVERSED_TEXT_STEMS)

    def test_threshold_requires_at_least_three_reversed_hits(self):
        # Only 2 reversed hits — below the absolute threshold of 3.
        raw = ["\u0417\u0430\u0449\u0438\u0442\u0430", "\u0410\u0432\u0442\u043e\u043c\u0430\u0442"]
        words = [_word(t[::-1]) for t in raw]
        self.assertFalse(_is_text_reversed(words))

    def test_forward_dominance_suppresses_false_positive(self):
        # 3 reversed hits but 5 forward hits → must stay False.
        fwd = [_word(t) for t in (
            "\u0417\u0430\u0449\u0438\u0442\u0430",
            "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435",
            "\u041e\u0431\u043e\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435",
            "\u0410\u0432\u0442\u043e\u043c\u0430\u0442",
            "\u041a\u0430\u0431\u0435\u043b\u044c",
        )]
        rev = [_word(t[::-1]) for t in (
            "\u0417\u0430\u0449\u0438\u0442\u0430",
            "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435",
            "\u041e\u0431\u043e\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435",
        )]
        self.assertFalse(_is_text_reversed(fwd + rev))


class ReverseCyrillicWordsTests(unittest.TestCase):

    def test_preserves_geometry(self):
        words = [_word("\u0417\u0430\u0449\u0438\u0442\u0430", x0=10, top=20, x1=60, bottom=28)]
        out = _reverse_cyrillic_words(words)
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0]["x0"], 10)
        self.assertEqual(out[0]["x1"], 60)
        self.assertEqual(out[0]["top"], 20)
        self.assertEqual(out[0]["bottom"], 28)
        self.assertEqual(out[0]["text"], "\u0430\u0442\u0438\u0449\u0430\u0417")  # атищаЗ

    def test_is_shallow_copy(self):
        words = [_word("\u0410\u0412", x0=0)]
        out = _reverse_cyrillic_words(words)
        # Modifying input must not affect output, and vice versa.
        words[0]["text"] = "XX"
        self.assertEqual(out[0]["text"], "\u0412\u0410")

    def test_non_dict_entries_pass_through(self):
        out = _reverse_cyrillic_words([None, "not a dict", 42])
        self.assertEqual(out, [None, "not a dict", 42])

    def test_missing_text_field_is_tolerated(self):
        out = _reverse_cyrillic_words([{"x0": 0, "top": 0}])
        self.assertEqual(out, [{"x0": 0, "top": 0}])


class ExtractWithToleranceFlipsReversedPagesTests(unittest.TestCase):

    def test_reversed_page_auto_flipped_and_tagged(self):
        raw_reversed = [
            _word(t[::-1], x0=10 + i * 60, top=30)
            for i, t in enumerate((
                "\u0417\u0430\u0449\u0438\u0442\u0430",
                "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435",
                "\u041e\u0431\u043e\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435",
                "\u0410\u0432\u0442\u043e\u043c\u0430\u0442",
                "\u041a\u0430\u0431\u0435\u043b\u044c",
            ))
        ]
        page = _FakePage(raw_reversed)
        words, _header = _extract_with_tolerance(page, page_idx=0, x_tol=3)
        # Words should be flipped back to forward form.
        forms = [w["text"] for w in words]
        self.assertIn("\u0417\u0430\u0449\u0438\u0442\u0430", forms)
        self.assertIn("\u0410\u0432\u0442\u043e\u043c\u0430\u0442", forms)
        # Page flagged for detection_method tagging.
        self.assertTrue(getattr(page, "_legend_reversed", False))

    def test_forward_page_untouched(self):
        words_in = [
            _word(t, x0=10 + i * 60, top=30)
            for i, t in enumerate((
                "\u0417\u0430\u0449\u0438\u0442\u0430",
                "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435",
                "\u041e\u0431\u043e\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435",
            ))
        ]
        page = _FakePage(words_in)
        words, _header = _extract_with_tolerance(page, page_idx=0, x_tol=3)
        forms = [w["text"] for w in words]
        self.assertIn("\u0417\u0430\u0449\u0438\u0442\u0430", forms)
        # No flag set on a forward page.
        self.assertFalse(getattr(page, "_legend_reversed", False))


class TagMethodTests(unittest.TestCase):

    def test_untagged_page_returns_base(self):
        class P:
            pass
        self.assertEqual(_tag_method("header", P()), "header")

    def test_flagged_page_gets_reversed_prefix(self):
        class P:
            _legend_reversed = True
        self.assertEqual(_tag_method("spec", P()), "reversed:spec")

    def test_legend_result_default_detection_method_empty(self):
        self.assertEqual(LegendResult().detection_method, "")


if __name__ == "__main__":
    unittest.main()
