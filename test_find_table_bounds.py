"""
Unit tests for pdf_legend_parser._find_table_bounds.

Task T017 (S002-02):
  * Verify MIN_TABLE_LINE_LEN threshold is low enough for small GPC tables.
  * Verify the function supports tables that have only horizontal separators
    (no vertical frame).

These tests feed synthetic pdfplumber-style `lines` dicts directly into
`_find_table_bounds` — no real PDF is opened.  Each line dict mirrors the
subset of keys that the parser actually consumes: x0, x1, top, bottom.
"""

from __future__ import annotations

import unittest

import pdf_legend_parser as plp


def _hline(x0: float, x1: float, y: float) -> dict:
    """Build a horizontal line dict (top == bottom)."""
    return {"x0": x0, "x1": x1, "top": y, "bottom": y}


def _vline(x: float, y_top: float, y_bottom: float) -> dict:
    """Build a vertical line dict (x0 == x1)."""
    return {"x0": x, "x1": x, "top": y_top, "bottom": y_bottom}


class MinTableLineLenTest(unittest.TestCase):
    """MIN_TABLE_LINE_LEN should be low enough for small GPC tables."""

    def test_threshold_is_at_most_25pt(self):
        self.assertLessEqual(
            plp.MIN_TABLE_LINE_LEN,
            25,
            "MIN_TABLE_LINE_LEN must be <= 25pt to admit small GPC "
            "legend tables (abk_em, abk_eg).",
        )

    def test_threshold_is_above_tick_marks(self):
        # Guard against accidental over-lowering: ticks / hatches are < 10pt.
        self.assertGreaterEqual(
            plp.MIN_TABLE_LINE_LEN,
            15,
            "MIN_TABLE_LINE_LEN should stay above typical tick stroke length.",
        )


class SmallGpcTableTest(unittest.TestCase):
    """Tables with short (< 50pt) row separators should now be detected."""

    def test_two_row_table_with_40pt_rows(self):
        # Header somewhere near y=100, table x range 200..240 (40pt wide).
        header_y = 100.0
        header_x = 200.0

        lines = [
            _hline(200, 240, 120),  # top border
            _hline(200, 240, 140),  # mid separator
            _hline(200, 240, 160),  # bottom border
            _vline(200, 120, 160),  # left border
            _vline(240, 120, 160),  # right border
        ]

        bounds = plp._find_table_bounds(lines, header_y, header_x, page_width=595.0)
        self.assertIsNotNone(bounds, "Should detect a 40pt-wide legend table")
        self.assertAlmostEqual(bounds.x0, 200, delta=1)
        self.assertAlmostEqual(bounds.x1, 240, delta=1)
        # 3 horizontal lines -> 3 row_ys
        self.assertEqual(len(bounds.row_ys), 3)


class HorizontalOnlyTableTest(unittest.TestCase):
    """Table with horizontal separators but no vertical frame (abk_em/eg)."""

    def test_horizontal_only_table_is_detected(self):
        header_y = 100.0
        header_x = 210.0

        # Three parallel horizontals -> 2 rows.  No vertical lines at all.
        lines = [
            _hline(200, 260, 130),
            _hline(200, 260, 150),
            _hline(200, 260, 170),
        ]

        bounds = plp._find_table_bounds(lines, header_y, header_x, page_width=595.0)
        self.assertIsNotNone(
            bounds,
            "Tables with only horizontal separators must still be detected.",
        )
        self.assertAlmostEqual(bounds.x0, 200, delta=1)
        self.assertAlmostEqual(bounds.x1, 260, delta=1)
        # Fallback: col_xs should be synthesised from x0/x1.
        self.assertGreaterEqual(len(bounds.col_xs), 2)
        self.assertAlmostEqual(bounds.col_xs[0], 200, delta=1)
        self.assertAlmostEqual(bounds.col_xs[-1], 260, delta=1)
        self.assertEqual(len(bounds.row_ys), 3)

    def test_horizontal_only_with_internal_short_divider(self):
        """Internal vertical divider shorter than full height should still
        be picked up by the relaxed pass."""
        header_y = 100.0
        header_x = 210.0

        lines = [
            _hline(200, 260, 130),
            _hline(200, 260, 150),
            _hline(200, 260, 170),
            # Short internal divider that overlaps the Y range but does not
            # span it top-to-bottom.
            _vline(230, 135, 165),
        ]

        bounds = plp._find_table_bounds(lines, header_y, header_x, page_width=595.0)
        self.assertIsNotNone(bounds)
        # Should include the internal divider as a column.
        self.assertTrue(
            any(abs(x - 230) < 1 for x in bounds.col_xs),
            f"Expected an internal column divider near x=230, got {bounds.col_xs}",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
