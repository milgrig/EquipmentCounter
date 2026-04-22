"""Unit tests for S003-03 (T029) cable-to-work derivation.

Covers:
  * `_parse_cross_section_conductors` — parses "3x1.5" / "3х1,5" etc.
  * `_derive_work_items` — produces three VOR-ready rows per parseable run
    ("Подключение жил" chosen by <=10 / >10 mm², plus "Прокладка кабеля").
  * `CableResult.derived_work_items` exists and is populated by the
    derivation pass.
  * Counts: n_conductors * 2 ends * run_count (aggregated).
"""

import unittest
from pdf_count_cables import (
    CableRun,
    CableResult,
    _parse_cross_section_conductors,
    _derive_work_items,
    _make_cable_work_name,
)


# Literal Russian strings kept as unicode escapes so the file stays
# printable under cp1254 shells.
_PODKL_DO10 = "\u041f\u043e\u0434\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u0436\u0438\u043b \u043a\u0430\u0431\u0435\u043b\u0435\u0439 \u0434\u043e 10 \u043c\u043c2"
_PODKL_SVYSHE = "\u041f\u043e\u0434\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u0436\u0438\u043b \u043a\u0430\u0431\u0435\u043b\u0435\u0439 \u0441\u0432\u044b\u0448\u0435 10 \u043c\u043c2"
_CYR_X = "\u0445"  # х


class ParseCrossSectionTests(unittest.TestCase):

    def test_standard_latin_x_dot(self):
        self.assertEqual(_parse_cross_section_conductors("3x1.5"), (3, 1.5))

    def test_standard_latin_x_comma(self):
        self.assertEqual(_parse_cross_section_conductors("5x2,5"), (5, 2.5))

    def test_cyrillic_x_comma(self):
        self.assertEqual(
            _parse_cross_section_conductors(f"3{_CYR_X}1,5"),
            (3, 1.5),
        )

    def test_integer_section(self):
        self.assertEqual(_parse_cross_section_conductors("5x16"), (5, 16.0))

    def test_large_section(self):
        self.assertEqual(_parse_cross_section_conductors("4x25"), (4, 25.0))

    def test_empty_input(self):
        self.assertEqual(_parse_cross_section_conductors(""), (0, 0.0))

    def test_garbage_input(self):
        self.assertEqual(_parse_cross_section_conductors("not a section"), (0, 0.0))

    def test_whitespace_padding(self):
        self.assertEqual(_parse_cross_section_conductors("  3x1,5  "), (3, 1.5))


class DeriveWorkItemsTests(unittest.TestCase):

    def _run(self, cs, n=1, length=None, brand=""):
        return [
            CableRun(cross_section=cs, length_m=length, cable_type=brand)
            for _ in range(n)
        ]

    def test_small_section_emits_do_10(self):
        runs = self._run("3x1.5", n=5, length=10)
        items = _derive_work_items(runs)
        conn = [i for i in items if i["name"] == _PODKL_DO10]
        self.assertEqual(len(conn), 1)
        # 5 runs * 3 conductors * 2 ends = 30
        self.assertEqual(conn[0]["count"], 30)
        self.assertEqual(conn[0]["unit"], "\u0448\u0442")  # шт
        self.assertEqual(conn[0]["category"], "cable_strand_connection_small")
        self.assertEqual(conn[0]["source"], "cable_derivation")

    def test_large_section_emits_svyshe_10(self):
        runs = self._run("4x25", n=2, length=50)
        items = _derive_work_items(runs)
        conn = [i for i in items if i["name"] == _PODKL_SVYSHE]
        self.assertEqual(len(conn), 1)
        # 2 runs * 4 conductors * 2 ends = 16
        self.assertEqual(conn[0]["count"], 16)
        self.assertEqual(conn[0]["category"], "cable_strand_connection_large")

    def test_boundary_exactly_10_is_small(self):
        # mm2 == 10.0 -> <= 10 branch (small).
        runs = self._run("3x10", n=1, length=1)
        items = _derive_work_items(runs)
        conn = [i for i in items if "cable_strand_connection" in i["category"]]
        self.assertEqual(len(conn), 1)
        self.assertEqual(conn[0]["category"], "cable_strand_connection_small")

    def test_prokladka_emits_total_length(self):
        runs = [
            CableRun(cross_section="3x1.5", length_m=12.5, cable_type="\u0412\u0412\u0413\u043d\u0433"),  # ВВГнг
            CableRun(cross_section="3x1.5", length_m=7.5, cable_type="\u0412\u0412\u0413\u043d\u0433"),
        ]
        items = _derive_work_items(runs)
        lay = [i for i in items if i["unit"] == "\u043c"]  # м
        self.assertEqual(len(lay), 1)
        self.assertAlmostEqual(lay[0]["count"], 20.0)
        self.assertIn("\u0412\u0412\u0413\u043d\u0433", lay[0]["name"])  # ВВГнг

    def test_unparseable_cross_section_skipped(self):
        runs = [CableRun(cross_section="", length_m=10)]
        items = _derive_work_items(runs)
        # No "Подключение жил" entry (no n_cond); also no laying because
        # work_name composition requires parseable section OR mark — the
        # derivation rule only emits laying when length>0 regardless of cs.
        conn = [i for i in items if "cable_strand_connection" in i["category"]]
        self.assertEqual(conn, [])
        # Laying row still emitted (length > 0)
        lay = [i for i in items if i["unit"] == "\u043c"]
        self.assertEqual(len(lay), 1)

    def test_length_zero_skips_laying(self):
        runs = [CableRun(cross_section="3x1.5", length_m=None)]
        items = _derive_work_items(runs)
        lay = [i for i in items if i["unit"] == "\u043c"]
        self.assertEqual(lay, [])

    def test_aggregation_across_mixed_sections(self):
        runs = [
            CableRun(cross_section="3x1.5", length_m=10),   # small
            CableRun(cross_section="5x2.5", length_m=5),    # small
            CableRun(cross_section="4x16",  length_m=15),   # large
        ]
        items = _derive_work_items(runs)
        cats = {i["category"]: i["count"] for i in items
                if "cable_strand_connection" in i["category"]}
        # small: 3*2 + 5*2 = 16
        self.assertEqual(cats["cable_strand_connection_small"], 16)
        # large: 4*2 = 8
        self.assertEqual(cats["cable_strand_connection_large"], 8)

    def test_make_cable_work_name_with_and_without_brand(self):
        n1 = _make_cable_work_name("\u0412\u0412\u0413\u043d\u0433", "3x1.5")  # ВВГнг
        self.assertIn("\u0412\u0412\u0413\u043d\u0433", n1)
        self.assertIn("3x1.5", n1)

        n2 = _make_cable_work_name("", "3x1.5")
        self.assertNotIn("  ", n2)
        self.assertIn("3x1.5", n2)

        n3 = _make_cable_work_name("", "")
        self.assertTrue(n3.startswith("\u041f\u0440\u043e\u043a\u043b\u0430\u0434\u043a\u0430"))  # Прокладка


class CableResultFieldTests(unittest.TestCase):

    def test_derived_work_items_defaults_empty(self):
        r = CableResult()
        self.assertEqual(r.derived_work_items, [])

    def test_derived_work_items_accepts_list(self):
        r = CableResult(derived_work_items=[{"name": "x", "count": 1}])
        self.assertEqual(len(r.derived_work_items), 1)


if __name__ == "__main__":
    unittest.main()
