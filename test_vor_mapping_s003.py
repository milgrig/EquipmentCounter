"""S003-02 (T028) tests for new VOR categories in vor_work_mapping.

Covers the 7 new categories added from the S003-01 analyst mapping:
  * cable_strand_connection_small  (Подключение жил кабелей до 10 мм2)
  * cable_strand_connection_large  (Подключение жил кабелей свыше 10 мм2)
  * cable_tray (cable-channel form — Монтаж кабель-канала {size} мм)
  * conduit_pvc                    (Монтаж трубы ПВХ {detail})
  * conduit_corrugated             (Монтаж трубы гофрированной {detail})
  * fireproof_sealant              (Двухкомпонентная огнестойкая пена ...)
  * socket_in_tray                 (Монтаж розеток в кабель-канале ...)

Backward-compatibility of pre-existing categories is also asserted
(panel / cable / wire / socket / luminaire / conduit fallback / tray
fallback) so this change-set does not silently regress legacy rules.
"""

import unittest

from vor_work_mapping import map_equipment_to_work, map_items


# Russian literals written as unicode escapes so the file stays pure
# ASCII and bypasses any cp1254-style printing glitches on Windows.
_ZHIL_DO10 = "\u041f\u043e\u0434\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u0436\u0438\u043b \u043a\u0430\u0431\u0435\u043b\u0435\u0439 \u0434\u043e 10 \u043c\u043c2"
_ZHIL_SVYSHE = "\u041f\u043e\u0434\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u0436\u0438\u043b \u043a\u0430\u0431\u0435\u043b\u0435\u0439 \u0441\u0432\u044b\u0448\u0435 10 \u043c\u043c2"
_KABEL_KANAL = "\u041a\u0430\u0431\u0435\u043b\u044c-\u043a\u0430\u043d\u0430\u043b 40\u044525"  # Кабель-канал 40х25
_TRUBA_PVH = "\u0422\u0440\u0443\u0431\u0430 \u041f\u0412\u0425 20 \u043c\u043c"  # Труба ПВХ 20 мм
_TRUBA_GOFR = "\u0422\u0440\u0443\u0431\u0430 \u0433\u043e\u0444\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u043d\u0430\u044f 16 \u043c\u043c"  # Труба гофрированная 16 мм
_GOFROTRUBA = "\u0413\u043e\u0444\u0440\u043e\u0442\u0440\u0443\u0431\u0430 \u041f\u0412\u0425 d=20"  # Гофротруба ПВХ d=20
_PENA = "\u0414\u0432\u0443\u0445\u043a\u043e\u043c\u043f\u043e\u043d\u0435\u043d\u0442\u043d\u0430\u044f \u043e\u0433\u043d\u0435\u0441\u0442\u043e\u0439\u043a\u0430\u044f \u043f\u0435\u043d\u0430 DN1201"
_ROZ_V_KANALE = "\u0420\u043e\u0437\u0435\u0442\u043a\u0430 \u0432 \u043a\u0430\u0431\u0435\u043b\u044c-\u043a\u0430\u043d\u0430\u043b\u0435 ABB"  # Розетка в кабель-канале ABB

# Legacy names
_ROZETKA = "\u0420\u043e\u0437\u0435\u0442\u043a\u0430 \u0420\u04101\u0036-003"  # Розетка РА16-003
_LOTOK = "\u041b\u043e\u0442\u043e\u043a 300\u0445100"  # Лоток 300х100
_TRUBA_LEGACY = "\u0422\u0440\u0443\u0431\u0430 \u041f\u0412\u0425 \u0433\u0438\u0431\u043a\u0430\u044f \u0433\u043e\u0444\u0440. \u0434.16\u043c\u043c"  # Труба ПВХ гибкая гофр. д.16мм
_SHCHR = "\u0429\u0420-12"  # ЩР-12
_CABLE = "\u041a\u0430\u0431\u0435\u043b\u044c \u0412\u0412\u0413\u043d\u0433 3\u04451.5"  # Кабель ВВГнг 3х1.5
_SVETILNIK = "\u0421\u0432\u0435\u0442\u0438\u043b\u044c\u043d\u0438\u043a \u0441\u0432\u0435\u0442\u043e\u0434\u0438\u043e\u0434\u043d\u044b\u0439 \u0414\u0412\u041e"  # Светильник светодиодный ДВО


class NewCategoriesTests(unittest.TestCase):

    def test_cable_strand_connection_small(self):
        r = map_equipment_to_work(_ZHIL_DO10, unit="\u0448\u0442")
        self.assertEqual(r.category, "cable_strand_connection_small")
        self.assertEqual(r.unit, "\u0448\u0442")
        self.assertIn("\u041f\u043e\u0434\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435", r.work_name)  # Подключение
        self.assertIn("\u0434\u043e 10", r.work_name)  # до 10

    def test_cable_strand_connection_large(self):
        r = map_equipment_to_work(_ZHIL_SVYSHE, unit="\u0448\u0442")
        self.assertEqual(r.category, "cable_strand_connection_large")
        self.assertEqual(r.unit, "\u0448\u0442")
        self.assertIn("\u0441\u0432\u044b\u0448\u0435 10", r.work_name)  # свыше 10

    def test_cable_tray_cable_channel(self):
        r = map_equipment_to_work(_KABEL_KANAL, unit="\u043c")
        self.assertEqual(r.category, "cable_tray")
        self.assertEqual(r.unit, "\u043c")
        self.assertIn("\u043a\u0430\u0431\u0435\u043b\u044c-\u043a\u0430\u043d\u0430\u043b", r.work_name.lower())  # кабель-канал
        # size survives into the template
        self.assertIn("40", r.work_name)

    def test_conduit_pvc(self):
        r = map_equipment_to_work(_TRUBA_PVH, unit="\u043c")
        self.assertEqual(r.category, "conduit_pvc")
        self.assertEqual(r.unit, "\u043c")
        self.assertIn("\u041f\u0412\u0425", r.work_name)  # ПВХ

    def test_conduit_corrugated(self):
        r = map_equipment_to_work(_TRUBA_GOFR, unit="\u043c")
        self.assertEqual(r.category, "conduit_corrugated")
        self.assertEqual(r.unit, "\u043c")
        self.assertIn(
            "\u0433\u043e\u0444\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u043d\u043e\u0439",
            r.work_name.lower(),  # гофрированной
        )

    def test_fireproof_sealant(self):
        r = map_equipment_to_work(_PENA, unit="\u043a\u043e\u043c\u043f\u043b")
        self.assertEqual(r.category, "fireproof_sealant")
        self.assertEqual(r.unit, "\u043a\u043e\u043c\u043f")  # normalised компл -> комп
        self.assertIn("DN1201", r.work_name)

    def test_socket_in_tray(self):
        r = map_equipment_to_work(_ROZ_V_KANALE, unit="\u0448\u0442")
        self.assertEqual(r.category, "socket_in_tray")
        self.assertEqual(r.unit, "\u0448\u0442")
        self.assertIn(
            "\u043a\u0430\u0431\u0435\u043b\u044c-\u043a\u0430\u043d\u0430\u043b\u0435",
            r.work_name.lower(),  # кабель-канале
        )
        self.assertIn("ABB", r.work_name)


class BackwardCompatTests(unittest.TestCase):
    """Ensure legacy categories still resolve exactly as before."""

    def test_socket_plain(self):
        r = map_equipment_to_work(_ROZETKA, unit="\u0448\u0442")
        self.assertEqual(r.category, "socket")

    def test_cable_tray_lotok(self):
        r = map_equipment_to_work(_LOTOK, unit="\u043c")
        self.assertEqual(r.category, "cable_tray")
        # Legacy lotok template retained
        self.assertIn(
            "\u041b\u043e\u0442\u043a\u0430 \u043a\u0430\u0431\u0435\u043b\u044c\u043d\u043e\u0433\u043e".lower(),  # лотка кабельного
            r.work_name.lower(),
        )

    def test_conduit_fallback_truba_pvh_gofr(self):
        # "Труба ПВХ гибкая гофр." — the "гофр" token wins and routes to
        # the corrugated category (this is the desired behaviour since
        # the row is physically a corrugated PVC pipe).
        r = map_equipment_to_work(_TRUBA_LEGACY, unit="\u043c")
        self.assertIn(r.category, {"conduit_corrugated", "conduit_pvc"})
        self.assertEqual(r.unit, "\u043c")

    def test_panel(self):
        r = map_equipment_to_work(_SHCHR, unit="\u0448\u0442")
        self.assertEqual(r.category, "panel")

    def test_cable(self):
        r = map_equipment_to_work(_CABLE, unit="\u043c")
        self.assertEqual(r.category, "cable")
        self.assertEqual(r.unit, "\u043c")

    def test_luminaire(self):
        r = map_equipment_to_work(_SVETILNIK, unit="\u0448\u0442")
        self.assertEqual(r.category, "luminaire")

    def test_unknown_falls_through_to_other(self):
        r = map_equipment_to_work("\u042d\u043b\u0435\u043a\u0442\u0440\u043e\u0441\u0447\u0435\u0442\u0447\u0438\u043a")  # Электросчетчик
        self.assertEqual(r.category, "other")

    def test_map_items_batch_preserves_new_categories(self):
        rows = [
            {"name": _ZHIL_DO10, "count": 100, "unit": "\u0448\u0442"},
            {"name": _KABEL_KANAL, "count": 25, "unit": "\u043c"},
            {"name": _PENA, "count": 5, "unit": "\u043a\u043e\u043c\u043f\u043b"},
        ]
        mapped = map_items(rows)
        cats = [r["category"] for r in mapped]
        self.assertIn("cable_strand_connection_small", cats)
        self.assertIn("cable_tray", cats)
        self.assertIn("fireproof_sealant", cats)


if __name__ == "__main__":
    unittest.main()
