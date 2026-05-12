"""
_batch_gpk3.py — Run pdf_cables_by_method on every drawing PDF in
ГПК 3-я захватка folder.
"""

from __future__ import annotations

import io
import os
import sys
import traceback
from collections import defaultdict

from pdf_cables_by_method import measure_cables, MM_PER_PT


ROOT = r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF"

# Skip non-drawing pages (title, common data, schematics, СО-СО, VOR PDF)
SKIP = {
    "001-Общие данные начало.pdf",
    "002-Общие данные окончание.pdf",
    "003 - Схемы ЩО, ЩАО-ЩО-3.pdf",
    "004 - Схемы ЩО, ЩАО-ЩАО-3.pdf",
    "004.1- Схемы ЩО, ЩАО-ЦСАО.pdf",
    "СО-СО_01.pdf", "СО-СО_02.pdf", "СО-СО_03.pdf", "СО-СО_04.pdf",
    "ВОР ЭО, Захватка 3_ГПК.pdf",
    "1) Титульный лист (1Д-24-3-3-ЭО) изм.2_рев.1.pdf",
}


def main() -> None:
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                                      errors="replace")

    files = sorted(f for f in os.listdir(ROOT)
                   if f.lower().endswith(".pdf") and f not in SKIP)

    print(f"Folder: {ROOT}")
    print(f"Files to scan: {len(files)}")
    print()
    header = (f"{'file':<50} {'scale':>7} {'src':<24} "
              f"{'r/b':>9} {'pl':>5} "
              f"{'po_konstr':>10} {'gofra':>8} {'lotok':>7} "
              f"{'TOTAL,m':>9}")
    print(header)
    print("-" * len(header))

    grand_po = grand_go = grand_lo = grand_total = 0.0
    by_kind = defaultdict(lambda: dict(po=0.0, go=0.0, lo=0.0, tot=0.0,
                                       count=0))

    for name in files:
        path = os.path.join(ROOT, name)
        try:
            rep = measure_cables(path)
        except Exception as e:
            print(f"{name:<50}  ERROR: {e!r}")
            traceback.print_exc()
            continue

        real_N = (rep.scale_mm_per_pt / MM_PER_PT) if rep.scale_mm_per_pt else 0
        scale_s = f"1:{real_N:.0f}" if real_N else "—"

        bym = defaultdict(float)
        for (_c, m), L in rep.totals_m.items():
            bym[m] += L
        po = bym.get("po_konstrukciyam", 0.0)
        go = bym.get("gofra/truba", 0.0)
        lo = bym.get("lotok", 0.0)
        tot = sum(rep.totals_m.values())

        # Classify by drawing kind
        low = name.lower()
        if "освещения" in low or "освещ" in low:
            kind = "Освещение"
        elif "привязк" in low or "приявязк" in low:
            kind = "Привязка"
        elif "лотк" in low:
            kind = "Лотки"
        else:
            kind = "Прочее"

        b = by_kind[kind]
        b["po"] += po; b["go"] += go; b["lo"] += lo
        b["tot"] += tot; b["count"] += 1

        grand_po += po; grand_go += go; grand_lo += lo; grand_total += tot

        # Truncate filename to 50
        fn = name if len(name) <= 49 else name[:46] + "..."
        print(f"{fn:<50} {scale_s:>7} {rep.scale_source:<24} "
              f"{rep.raw_red_lines:>4}/{rep.raw_blue_lines:<4} "
              f"{len(rep.polylines):>5} "
              f"{po:>10.1f} {go:>8.1f} {lo:>7.1f} "
              f"{tot:>9.1f}")

    print("-" * len(header))
    print()
    print("=== Summary by drawing kind ===")
    print(f"{'kind':<14} {'sheets':>7} {'po_konstr':>11} "
          f"{'gofra':>8} {'lotok':>8} {'TOTAL,m':>10}")
    for k, b in by_kind.items():
        print(f"{k:<14} {b['count']:>7} {b['po']:>11.1f} "
              f"{b['go']:>8.1f} {b['lo']:>8.1f} {b['tot']:>10.1f}")
    print(f"{'GRAND TOTAL':<14} {'':>7} {grand_po:>11.1f} "
          f"{grand_go:>8.1f} {grand_lo:>8.1f} {grand_total:>10.1f}")


if __name__ == "__main__":
    main()
