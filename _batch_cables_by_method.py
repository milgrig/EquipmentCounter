"""
_batch_cables_by_method.py — Run pdf_cables_by_method on a curated set
of lighting plans across projects and print a side-by-side summary.

Usage:
    python _batch_cables_by_method.py
"""

from __future__ import annotations

import io
import sys
import traceback
from collections import defaultdict

from pdf_cables_by_method import measure_cables, MM_PER_PT


# Curated set: (label, pdf_path)
PLANS = [
    ("KPP-007",
     r"Data\ДБТ разделы для ИИ\30. КПП\03_PDF\007-План освещения.pdf"),
    ("GPK-005",
     r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf"),
    ("GPK-006",
     r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\006-Планы освещения-отм. +4.200.pdf"),
    ("GPK-007",
     r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\007-Планы освещения-отм. +7.800 +9.000.pdf"),
    ("GPK-008",
     r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\008-Планы освещения-отм. +13.800.pdf"),
    ("GPK-009",
     r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\009-Планы освещения-отм. +18.600.pdf"),
    ("GPK-010",
     r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\010-Планы освещения-отм. +23.400.pdf"),
    ("GPK-011",
     r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\011-Планы освещения-отм. +28.200.pdf"),
    ("Sulfat-003",
     r"Data\ДБТ разделы для ИИ\8.2 Участок хранения сульфата аммония\ЭО\02_PDF\003-План освещения.pdf"),
    ("Nasosnaya-003",
     r"Data\ДБТ разделы для ИИ\12.3 Насосная станция поверхностных стоков\02_PDF\003 - Планы освещения.pdf"),
    ("Avtovesy-004",
     r"Data\ДБТ разделы для ИИ\28. Автовесы\02_PDF\1-Д-24-28-ЭОМ-004.Планы освещения.pdf"),
    ("ZhdKPP-005",
     r"Data\ДБТ разделы для ИИ\16.2_ЖД КПП\Обновленные файлы 2\02_PDF\005-План освещения.pdf"),
    ("Sklad27-015",
     r"Data\ДБТ разделы для ИИ\27_Склад вспомогательных материалов с участком погрузки крытых вагонов\Обновленные файлы\03_PDF\015-План освещения.pdf"),
    ("ABK-006",
     r"Data\ДБТ разделы для ИИ\01_АБК\Обновленные файлы\PDF_\ЭО\006-План освещения на отм. 0.000.pdf"),
    ("ABK-007",
     r"Data\ДБТ разделы для ИИ\01_АБК\Обновленные файлы\PDF_\ЭО\007-План освещения на отм. +4.500, +9.000.pdf"),
]


def main() -> None:
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                                      errors="replace")

    print(f"{'plan':<16} {'scale':>8} {'src':<28} "
          f"{'segs r/b':>10} {'plines':>6} "
          f"{'po_konstr':>10} {'gofra':>8} {'lotok':>7} "
          f"{'TOTAL,m':>9}")
    print("-" * 110)

    for label, path in PLANS:
        try:
            rep = measure_cables(path)
        except FileNotFoundError:
            print(f"{label:<16}  FILE NOT FOUND: {path}")
            continue
        except Exception as e:
            print(f"{label:<16}  ERROR: {e!r}")
            traceback.print_exc()
            continue

        real_N = (rep.scale_mm_per_pt / MM_PER_PT) if rep.scale_mm_per_pt else 0
        scale_s = f"1:{real_N:.0f}" if real_N else "—"

        by_method: dict[str, float] = defaultdict(float)
        for (_color, m), L in rep.totals_m.items():
            by_method[m] += L

        po = by_method.get("po_konstrukciyam", 0.0)
        go = by_method.get("gofra/truba", 0.0)
        lo = by_method.get("lotok", 0.0)
        tot = sum(rep.totals_m.values())

        print(f"{label:<16} {scale_s:>8} {rep.scale_source:<28} "
              f"{rep.raw_red_lines:>4}/{rep.raw_blue_lines:<4} "
              f"{len(rep.polylines):>6} "
              f"{po:>10.1f} {go:>8.1f} {lo:>7.1f} "
              f"{tot:>9.1f}")


if __name__ == "__main__":
    main()
