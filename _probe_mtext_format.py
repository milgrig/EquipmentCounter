#!/usr/bin/env python3
"""Показывает сырой .text MTEXT в BLOCKS, ищем нестандартное форматирование,
мешающее увидеть FRLS / FRHF."""
from __future__ import annotations
import io, sys, re
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import ezdxf  # type: ignore

DXF = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG/"
    "003, 004 - Схемы ЩО, ЩАО.dxf"
)


def main() -> None:
    doc = ezdxf.readfile(str(DXF))

    # Возьмём первые 20 MTEXT из BLOCKS и напечатаем сырой .text
    samples_shown = 0
    for block in doc.blocks:
        if block.name.startswith("*"):
            continue
        for ent in block:
            if ent.dxftype() == "MTEXT":
                raw = ent.text or ""
                # Интересны MTEXT длиной > 50 (похожи на ячейки журнала)
                if 20 < len(raw) < 500 and ("ВБ" in raw or "ПП" in raw or "FR" in raw or "=" in raw):
                    print(f"--- block={block.name}, len={len(raw)} ---")
                    print(repr(raw[:500]))
                    print()
                    samples_shown += 1
                    if samples_shown >= 15:
                        return


if __name__ == "__main__":
    main()
