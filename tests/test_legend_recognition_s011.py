# -*- coding: utf-8 -*-
"""Тесты улучшений распознавания элементов легенды (S011, К1-К3/С4/С6).

Покрывает:
  * pdf_count_visual._template_similarity — guard перекрёстного NMS;
  * pdf_count_visual: мультимасштаб маркированных символов (константы);
  * pdf_count_text: фильтр ампер-номиналов (AMP_*);
  * pdf_count_anchored: канонизация ориентации, размерный матч, IoU масок.
"""
import os
import sys
import unittest

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pdf_count_visual as pcv  # noqa: E402
from pdf_count_anchored import (  # noqa: E402
    _glyph_clusters,
    _mask_iou,
    _size_match,
)
from pdf_count_text import AMP_CONTEXT_RE, AMP_VALUE_RE  # noqa: E402


class TemplateSimilarityTests(unittest.TestCase):

    @staticmethod
    def _tpl(draw):
        img = np.full((40, 40), 255, np.uint8)
        draw(img)
        return img

    def test_identical_templates_similar(self):
        a = self._tpl(lambda im: im.__setitem__((slice(10, 30), slice(10, 30)), 0))
        b = self._tpl(lambda im: im.__setitem__((slice(10, 30), slice(10, 30)), 0))
        self.assertGreaterEqual(pcv._template_similarity(a, b), 0.9)

    def test_disjoint_templates_dissimilar(self):
        a = self._tpl(lambda im: im.__setitem__((slice(2, 12), slice(2, 12)), 0))
        b = self._tpl(lambda im: im.__setitem__((slice(28, 38), slice(28, 38)), 0))
        self.assertLess(pcv._template_similarity(a, b), 0.2)

    def test_constants_recalibrated(self):
        # К1/К3: пороги подняты (S011); MARKED_SCALES введён (К2).
        self.assertGreaterEqual(pcv.SHAPE_VERIFY_MIN_RECALL, 0.28)
        self.assertGreaterEqual(pcv.SHAPE_VERIFY_MIN_PRECISION, 0.10)
        self.assertIn(1.0, pcv.MARKED_SCALES)
        self.assertGreater(len(pcv.MARKED_SCALES), 1)


class AmpFilterTests(unittest.TestCase):

    def test_amp_values(self):
        for t in ("16А", "25А", "63А"):
            self.assertTrue(AMP_VALUE_RE.match(t), t)
        for t in ("5АЭ", "7А1", "100А", "3А"):  # 3А — не типовой номинал
            self.assertFalse(AMP_VALUE_RE.match(t), t)

    def test_amp_context(self):
        for t in ("QF1", "ВА47", "C16", "кА", "3П"):
            self.assertTrue(AMP_CONTEXT_RE.match(t), t)
        for t in ("Светильник", "Гр.5", "ЩО3"):
            self.assertFalse(AMP_CONTEXT_RE.match(t), t)


class AnchoredHelpersTests(unittest.TestCase):

    def test_size_match_with_rotation(self):
        self.assertTrue(_size_match(100, 20, 100, 20))
        self.assertTrue(_size_match(20, 100, 100, 20))   # поворот 90°
        self.assertTrue(_size_match(120, 22, 100, 20))   # в допуске 35%
        self.assertFalse(_size_match(200, 20, 100, 20))  # вдвое длиннее

    def test_mask_iou(self):
        a = np.zeros((32, 32), bool); a[8:24, 8:24] = True
        b = np.zeros((32, 32), bool); b[8:24, 8:24] = True
        c = np.zeros((32, 32), bool); c[0:4, 0:4] = True
        self.assertAlmostEqual(_mask_iou(a, b), 1.0)
        self.assertEqual(_mask_iou(a, c), 0.0)

    def test_glyph_clusters_canonical_orientation(self):
        # Два одинаковых глифа: горизонтальный и вертикальный. После
        # канонизации оба должны получить cw >= ch и совпасть по размеру.
        mask = np.zeros((400, 400), np.uint8)
        mask[50:70, 50:150] = 255    # горизонтальный 100x20
        mask[200:300, 200:220] = 255  # вертикальный 20x100
        clusters = _glyph_clusters(mask)
        self.assertEqual(len(clusters), 2)
        for c in clusters:
            self.assertGreaterEqual(c["cw"], c["ch"])
        self.assertTrue(_size_match(clusters[0]["cw"], clusters[0]["ch"],
                                    clusters[1]["cw"], clusters[1]["ch"]))
        iou = _mask_iou(clusters[0]["mask32"], clusters[1]["mask32"])
        self.assertGreaterEqual(iou, 0.6)


if __name__ == "__main__":
    unittest.main()
