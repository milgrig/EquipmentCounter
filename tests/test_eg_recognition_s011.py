# -*- coding: utf-8 -*-
"""Тесты no-СО извлечения ЭГ (S011, Test2707)."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pdf_cables_by_method import _bridge_dashes, _classify_color  # noqa: E402
from pdf_count_grounding import (  # noqa: E402
    _bundle_len,
    _holder_step,
    _purchase,
    classify_eg_sheet,
    is_eg_sheet,
)


class ClassifyEgSheetTests(unittest.TestCase):

    def test_kinds(self):
        self.assertEqual(classify_eg_sheet("003-План заземления на отм. 0.000.pdf"),
                         "grounding")
        self.assertEqual(classify_eg_sheet("005-План молниезащиты.pdf"),
                         "lightning")
        self.assertEqual(classify_eg_sheet("002-СУП.pdf"), "bonding")
        self.assertEqual(classify_eg_sheet("004-План УП на отм. +4.500.pdf"),
                         "bonding")
        self.assertEqual(classify_eg_sheet("007-План освещения.pdf"), "")
        self.assertFalse(is_eg_sheet("018-Планы лотков на отм 0.000.pdf"))


class DerivationTests(unittest.TestCase):

    def test_holder_step_from_notes(self):
        self.assertEqual(_holder_step("Шаг установки держателей 1,0м."), 1.0)
        self.assertEqual(_holder_step("шаг установки арт.294011 L=0,5 м"), 0.5)
        self.assertEqual(_holder_step("нет шага"), 1.0)  # default

    def test_bundle_len(self):
        t = "проволока (бухта 110 м); полоса МПП (бухта 38 м)"
        self.assertEqual(_bundle_len(t, 50, 200, 110.0), 110.0)
        self.assertEqual(_bundle_len(t, 20, 50, 38.0), 38.0)
        self.assertEqual(_bundle_len("", 20, 50, 38.0), 38.0)

    def test_purchase_rounds_up_to_bundles(self):
        # Эталон АБК-1: контур 203.9 м → 6 бухт × 38 = 228 м (эталон 229).
        n, m = _purchase(203.9, 38.0)
        self.assertEqual((n, m), (6, 228.0))
        n, m = _purchase(242.2, 110.0)
        self.assertEqual((n, m), (3, 330.0))


class BridgeDashesTests(unittest.TestCase):

    @staticmethod
    def _seg(x0, y0, x1, y1):
        return {"x0": x0, "top": y0, "x1": x1, "bottom": y1,
                "stroking_color": (1, 0, 0)}

    def test_dashes_merge_including_gaps(self):
        # Пунктир: штрихи 10 pt с пробелами 5 pt вдоль одной прямой.
        segs = [self._seg(x, 100, x + 10, 100) for x in (0, 15, 30, 45)]
        out = _bridge_dashes(segs, gap_pt=8.0)
        self.assertEqual(len(out), 1)
        total = abs(out[0]["x1"] - out[0]["x0"])
        self.assertAlmostEqual(total, 55.0)  # 45+10 − 0: пробелы вошли

    def test_distant_runs_not_merged(self):
        segs = [self._seg(0, 100, 10, 100), self._seg(100, 100, 110, 100)]
        out = _bridge_dashes(segs, gap_pt=8.0)
        self.assertEqual(len(out), 2)

    def test_different_offsets_not_merged(self):
        segs = [self._seg(0, 100, 10, 100), self._seg(15, 130, 25, 130)]
        out = _bridge_dashes(segs, gap_pt=8.0)
        self.assertEqual(len(out), 2)


class ColorPaletteTests(unittest.TestCase):

    def test_green_only_when_requested(self):
        green = (0.0, 1.0, 0.0)
        self.assertEqual(_classify_color(green), "")           # default red/blue
        self.assertEqual(_classify_color(green, ("green",)), "green")
        self.assertEqual(_classify_color((1.0, 0.0, 0.0)), "red")


if __name__ == "__main__":
    unittest.main()
