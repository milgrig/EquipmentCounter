#!/usr/bin/env python3
"""
Read-only: исследует структуру таблиц в "003, 004 - Схемы ЩО, ЩАО.dxf".

Цель — понять, почему сырой скан находит 102+74 = 176 пар
"марка FRLS/FRHF + L=", а эталон учитывает только 116 кабелей всего.
Задача: найти ключ дедупликации (номер линии / номер группы / обозначение трассы).

Стратегия:
  1. Читаем DXF через ezdxf.
  2. Собираем все MTEXT и TEXT с координатами вставки.
  3. Находим все MTEXT, содержащие FRLS 3х1,5 или FRHF 3х1,5.
  4. Для каждого такого MTEXT находим соседние (в радиусе 5000 юнитов) MTEXT/TEXT.
  5. Распечатываем первые 10 примеров с полным контекстом.
"""
from __future__ import annotations

import io
import re
import sys
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import ezdxf  # type: ignore

DXF_PATH = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG/"
    "003, 004 - Схемы ЩО, ЩАО.dxf"
)

FRLS_RE = re.compile(
    r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*1[,.]5",
    re.IGNORECASE,
)
FRHF_RE = re.compile(
    r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF\s*3\s*[xх×]\s*1[,.]5",
    re.IGNORECASE,
)
L_RE = re.compile(r"L\s*=\s*(\d{1,5}(?:[,.]\d+)?)", re.IGNORECASE)


def strip_mtext_codes(text: str) -> str:
    """Удаляет MTEXT форматирование: \\P, {…;…}, \\C1;, \\fISOCPEUR|..;"""
    text = re.sub(r"\\[A-Za-z]\s*[^\\;]*;?", "", text)
    text = re.sub(r"\\P", " ", text)
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def main() -> None:
    print(f"Читаю {DXF_PATH.name} ...")
    doc = ezdxf.readfile(str(DXF_PATH))
    msp = doc.modelspace()

    all_texts: list[tuple[float, float, str]] = []

    for ent in msp:
        if ent.dxftype() == "MTEXT":
            try:
                x, y = ent.dxf.insert.x, ent.dxf.insert.y
                raw = ent.text or ""
                cleaned = strip_mtext_codes(raw)
                all_texts.append((x, y, cleaned))
            except Exception:
                pass
        elif ent.dxftype() == "TEXT":
            try:
                x, y = ent.dxf.insert.x, ent.dxf.insert.y
                cleaned = (ent.dxf.text or "").strip()
                all_texts.append((x, y, cleaned))
            except Exception:
                pass

    print(f"  Всего MTEXT/TEXT в modelspace: {len(all_texts)}")

    # Также пройдёмся по BLOCKS (там могут быть таблицы)
    for block in doc.blocks:
        if block.name.startswith("*"):
            continue
        for ent in block:
            if ent.dxftype() == "MTEXT":
                try:
                    x, y = ent.dxf.insert.x, ent.dxf.insert.y
                    raw = ent.text or ""
                    cleaned = strip_mtext_codes(raw)
                    all_texts.append((x, y, cleaned))
                except Exception:
                    pass
            elif ent.dxftype() == "TEXT":
                try:
                    x, y = ent.dxf.insert.x, ent.dxf.insert.y
                    cleaned = (ent.dxf.text or "").strip()
                    all_texts.append((x, y, cleaned))
                except Exception:
                    pass

    print(f"  После BLOCKS: {len(all_texts)}")

    # Находим все с FRLS 3х1,5
    frls_hits = [t for t in all_texts if FRLS_RE.search(t[2])]
    frhf_hits = [t for t in all_texts if FRHF_RE.search(t[2])]
    print(f"\n  FRLS 3х1,5 найдено в ezdxf-тексте: {len(frls_hits)}")
    print(f"  FRHF 3х1,5 найдено в ezdxf-тексте: {len(frhf_hits)}")

    # Показываем первые 5 FRLS вместе с соседями (5 ближайших MTEXT/TEXT)
    def show_neighbors(hits, label):
        print(f"\n===== Первые 5 вхождений {label} с соседями =====\n")
        for idx, (x, y, txt) in enumerate(hits[:5], 1):
            print(f"--- #{idx}  @ ({x:.1f}, {y:.1f}) ---")
            print(f"  TEXT: {txt[:200]}")
            # Ищем соседей в радиусе 3000 единиц
            neighbors = []
            for (nx, ny, ntxt) in all_texts:
                if (nx, ny, ntxt) == (x, y, txt):
                    continue
                dx = nx - x
                dy = ny - y
                dist = (dx * dx + dy * dy) ** 0.5
                if dist < 3000 and ntxt.strip():
                    neighbors.append((dist, dx, dy, ntxt))
            neighbors.sort()
            for dist, dx, dy, ntxt in neighbors[:12]:
                print(f"    [d={dist:7.0f} dx={dx:+7.0f} dy={dy:+7.0f}] {ntxt[:120]}")
            print()

    show_neighbors(frls_hits, "FRLS 3х1,5")
    show_neighbors(frhf_hits, "FRHF 3х1,5")

    # Посмотрим, сколько MTEXT содержат и марку и L= одновременно (сразу, без окна)
    both_in_one = sum(
        1 for (_, _, t) in all_texts
        if (FRLS_RE.search(t) or FRHF_RE.search(t)) and L_RE.search(t)
    )
    print(f"\n  MTEXT где И марка (FRLS/FRHF 3х1,5) И L= в одном блоке: {both_in_one}")

    # Показать все уникальные L= значения вблизи FRLS/FRHF
    lengths_near_frls = []
    for (x, y, txt) in frls_hits + frhf_hits:
        for (nx, ny, ntxt) in all_texts:
            dist = ((nx - x) ** 2 + (ny - y) ** 2) ** 0.5
            if dist < 1500:
                for m in L_RE.finditer(ntxt):
                    try:
                        lengths_near_frls.append(float(m.group(1).replace(",", ".")))
                    except ValueError:
                        pass

    if lengths_near_frls:
        print(f"\n  L=NNN значений в радиусе 1500 от FRLS/FRHF: {len(lengths_near_frls)}")
        print(f"    Сумма: {sum(lengths_near_frls):.0f} м")
        print(f"    Min/Med/Max: {min(lengths_near_frls):.0f}/{sorted(lengths_near_frls)[len(lengths_near_frls)//2]:.0f}/{max(lengths_near_frls):.0f}")


if __name__ == "__main__":
    main()
