# -*- coding: utf-8 -*-
"""Тесты schema-cable fallback для CAD-ветки (S011, КПП-30)."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from vor_generator import _CABLE_LEN_RE  # noqa: E402


class CableLenReTests(unittest.TestCase):

    def test_plain_format(self):
        m = _CABLE_LEN_RE.search("Резервный ввод, ППГнг-FRHF 5х2,5 L=5м")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "ППГнг-FRHF 5х2,5")
        self.assertEqual(m.group(2), "5")

    def test_kpp30_format_with_laying_and_du(self):
        # Старый паттерн терял такие строки: между сечением и L= стоит
        # способ прокладки и падение напряжения.
        m = _CABLE_LEN_RE.search("ППГнг(А)-HF 5х2,5 в гофре ΔU=0,11% L=10м")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "ППГнг(А)-HF 5х2,5")
        self.assertEqual(m.group(2), "10")

    def test_fractional_length(self):
        m = _CABLE_LEN_RE.search("ВБШвнг(А)-LS 3х2,5 в земле L=7,5м")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(2), "7,5")

    def test_ls_mark_suffix_not_confused_with_len(self):
        # «-LS» в марке не должен обрывать зазор до L=.
        m = _CABLE_LEN_RE.search("ВБШвнг(А)-LS 3х2,5 открыто L=48м")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "ВБШвнг(А)-LS 3х2,5")
        self.assertEqual(m.group(2), "48")

    def test_no_length_no_match(self):
        self.assertIsNone(_CABLE_LEN_RE.search("Кабельная трасса, ППГнг(А)-HF"))


if __name__ == "__main__":
    unittest.main()
