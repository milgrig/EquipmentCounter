# -*- coding: utf-8 -*-
"""Тесты улучшений распознавания проводки (S011, П1-П5).

Покрывает:
  * pdf_cable_height._merge_multiline_blocks — адаптивное слияние строк
    (баг «65 групп вместо 27», метраж стояков ×2.4);
  * pdf_cables_by_method: новые способы прокладки (кабель-канал, в земле)
    в правилах листа, извлечение аннотаций трасс и их привязка к
    ближайшей полилинии;
  * pdf_vor_nospec._raster_method_key — маппинг категорий растрового
    движка на способы прокладки.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pdf_cable_height import _merge_multiline_blocks, extract_annotations_from_page  # noqa: E402
from pdf_cables_by_method import (  # noqa: E402
    Polyline,
    _bind_annotations_to_polylines,
    _detect_sheet_default_method,
    _extract_route_annotations,
)
from pdf_vor_nospec import _raster_method_key  # noqa: E402


def _line(text, x0, y0, x1, y1, color=None):
    return {"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1,
            "color": color}


class MergeMultilineBlocksTests(unittest.TestCase):
    """Геометрия из реального листа 006 ГПК-3: h=6.2 pt, межстрочный
    зазор ~1.4 pt внутри аннотации, ~6 pt между соседними аннотациями."""

    def test_two_lines_of_one_annotation_merge(self):
        blocks = _merge_multiline_blocks([
            _line("Гр.1-Гр.8, Гр.15-Гр.31", 818.4, 830.7, 893.6, 836.8, "blue"),
            _line("на отм. 0.000, +9.000", 818.4, 838.2, 940.0, 844.4, "blue"),
        ])
        self.assertEqual(len(blocks), 1)
        self.assertIn("Гр.1-Гр.8", blocks[0]["text"])
        self.assertIn("на отм.", blocks[0]["text"])

    def test_adjacent_annotations_do_not_merge(self):
        # Зазор 6 pt (844.4 → 850.4) > 0.6 × 6.2 — разные аннотации.
        blocks = _merge_multiline_blocks([
            _line("на отм. 0.000, +9.000", 818.4, 838.2, 940.0, 844.4, "blue"),
            _line("Гр.1А-Гр.6А", 818.4, 850.4, 881.2, 856.6, "red"),
        ])
        self.assertEqual(len(blocks), 2)

    def test_different_colors_do_not_merge(self):
        # Даже впритык: красная и синяя строки — разные трассы.
        blocks = _merge_multiline_blocks([
            _line("Гр.33", 790.0, 921.0, 804.4, 927.2, "red"),
            _line("Гр.34", 790.0, 928.0, 804.4, 934.2, "blue"),
        ])
        self.assertEqual(len(blocks), 2)

    def test_horizontally_distant_label_not_merged(self):
        # Метка щита в 15 pt правее блока (было: ±80 pt затягивало её).
        blocks = _merge_multiline_blocks([
            _line("на отм. 0.000, +9.000", 818.4, 873.1, 940.1, 879.3, "red"),
            _line("ЩАО3-Гр.4А", 955.8, 874.0, 990.1, 880.2, "red"),
        ])
        self.assertEqual(len(blocks), 2)


class SheetDefaultMethodTests(unittest.TestCase):

    @staticmethod
    def _words(*texts):
        out = []
        for i, t in enumerate(texts):
            x = 10.0
            for w in t.split():
                out.append({"text": w, "x0": x, "x1": x + 8 * len(w),
                            "top": 100.0 + i * 30, "bottom": 106.0 + i * 30})
                x += 8 * len(w) + 4
        return out

    def test_kabel_kanal_detected(self):
        method, info = _detect_sheet_default_method(self._words(
            "Проводку выполнить в кабель-канале 20х10",
            "Сети прокладываются в кабельном канале",
        ))
        self.assertEqual(method, "kabel_kanal")

    def test_v_zemle_detected(self):
        method, _ = _detect_sheet_default_method(self._words(
            "Кабель проложить в земле в трубе ПНД",
            "Прокладка в траншее на глубине 0,7 м",
        ))
        self.assertEqual(method, "v_zemle")


class AnnotationBindingTests(unittest.TestCase):

    @staticmethod
    def _words_line(text, x0=100.0, top=200.0):
        out = []
        x = x0
        for w in text.split():
            out.append({"text": w, "x0": x, "x1": x + 6 * len(w),
                        "top": top, "bottom": top + 6})
            x += 6 * len(w) + 3
        return out

    def test_extract_and_bind_method_overrides(self):
        words = self._words_line("ВВГнг(А)-LS 3х1,5 в кабель-канале", top=200)
        anns = _extract_route_annotations(words, zones=[])
        self.assertEqual(len(anns), 1)
        self.assertEqual(anns[0]["method"], "kabel_kanal")
        self.assertTrue(anns[0]["mark"].startswith("ВВГ"))
        self.assertEqual(anns[0]["section"], "3x1.5")

        near = Polyline(color="blue", points=[(150.0, 210.0), (300.0, 210.0)],
                        length_pt=150.0, laying_method="lotok")
        far = Polyline(color="blue", points=[(900.0, 900.0), (990.0, 900.0)],
                       length_pt=90.0, laying_method="lotok")
        bound = _bind_annotations_to_polylines(anns, [near, far])
        self.assertEqual(bound, 1)
        self.assertEqual(near.laying_method, "kabel_kanal")
        self.assertEqual(near.method_source, "annotation")
        self.assertEqual(near.cable_mark, anns[0]["mark"])
        self.assertEqual(near.cross_section, "3x1.5")
        self.assertEqual(far.laying_method, "lotok")  # не тронута

    def test_label_length_binds(self):
        words = self._words_line("L=25м", top=300)
        anns = _extract_route_annotations(words, zones=[])
        self.assertEqual(len(anns), 1)
        self.assertAlmostEqual(anns[0]["label_m"], 25.0)
        pl = Polyline(color="red", points=[(110.0, 305.0), (120.0, 305.0)],
                      length_pt=10.0)
        _bind_annotations_to_polylines(anns, [pl])
        self.assertAlmostEqual(pl.nearby_label_m, 25.0)

    def test_tray_size_not_taken_as_cross_section(self):
        # "50х50" — размер лотка (жилы 50 > 19), не сечение кабеля.
        words = self._words_line("Лоток 50х50 в лотке", top=400)
        anns = _extract_route_annotations(words, zones=[])
        self.assertEqual(len(anns), 1)
        self.assertEqual(anns[0]["section"], "")

    def test_out_of_range_annotation_not_bound(self):
        words = self._words_line("в гофре", top=500)
        anns = _extract_route_annotations(words, zones=[])
        pl = Polyline(color="red", points=[(800.0, 800.0)], length_pt=10.0,
                      laying_method="po_konstrukciyam")
        bound = _bind_annotations_to_polylines(anns, [pl])
        self.assertEqual(bound, 0)
        self.assertEqual(pl.laying_method, "po_konstrukciyam")


class RasterMethodKeyTests(unittest.TestCase):

    def test_mapping(self):
        # Реальные ключи категорий cable_length + защита от пустого.
        self.assertEqual(_raster_method_key("wire_tray"), "lotok")
        self.assertEqual(_raster_method_key("wire_pipe_hidden"), "gofra/truba")
        self.assertEqual(_raster_method_key("wire_pipe_open"), "gofra/truba")
        self.assertEqual(_raster_method_key("cable_emergency"),
                         "po_konstrukciyam")
        self.assertEqual(_raster_method_key("cable_working"),
                         "po_konstrukciyam")
        self.assertEqual(_raster_method_key(""), "po_konstrukciyam")


if __name__ == "__main__":
    unittest.main()
