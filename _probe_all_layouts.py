#!/usr/bin/env python3
"""
Read-only: проходит ВСЕ layouts (modelspace + paperspace) через ezdxf,
плюс все блоки и все entities — и ищет FRLS/FRHF 3х1,5 с L=NNN в одном MTEXT.
"""
from __future__ import annotations
import io, re, sys
from pathlib import Path
from collections import defaultdict

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import ezdxf  # type: ignore

DXF = Path(
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
L_RE = re.compile(r"L\s*=\s*(\d{1,5}(?:[,.]\d+)?)\s*м?", re.IGNORECASE)


def strip_mtext_codes(text: str) -> str:
    # Убираем {\fISOCPEUR|b0|i1|c0; ... } форматирование
    text = re.sub(r"\\[A-Za-z]\s*[^\\;]*;?", "", text)
    text = re.sub(r"\\P", " ", text)
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def find_in_entities(iterable, source_label: str, bucket: list[dict]) -> None:
    for ent in iterable:
        if ent.dxftype() != "MTEXT":
            continue
        try:
            raw = ent.text or ""
        except Exception:
            continue
        cleaned = strip_mtext_codes(raw)
        is_frls = bool(FRLS_RE.search(cleaned))
        is_frhf = bool(FRHF_RE.search(cleaned))
        if not (is_frls or is_frhf):
            continue
        lm = L_RE.search(cleaned)
        if not lm:
            continue
        try:
            length = float(lm.group(1).replace(",", "."))
        except ValueError:
            continue
        brand = "FRLS 3x1,5" if is_frls else "FRHF 3x1,5"
        try:
            x, y = ent.dxf.insert.x, ent.dxf.insert.y
        except Exception:
            x, y = None, None
        bucket.append(
            {
                "source": source_label,
                "brand": brand,
                "length_m": length,
                "x": x,
                "y": y,
                "text": cleaned[:200],
            }
        )


def main() -> None:
    print(f"Читаю {DXF.name} ...")
    doc = ezdxf.readfile(str(DXF))

    found: list[dict] = []

    # 1) modelspace
    find_in_entities(doc.modelspace(), "modelspace", found)
    print(f"  modelspace: {sum(1 for f in found if f['source']=='modelspace')} hits")

    # 2) все paperspace layouts
    for layout in doc.layouts:
        if layout.name == "Model":
            continue
        pre = len(found)
        find_in_entities(layout, f"layout:{layout.name}", found)
        post = len(found)
        if post > pre:
            print(f"  layout '{layout.name}': {post-pre} hits")

    # 3) все блоки (BLOCKS section)
    for block in doc.blocks:
        if block.name.startswith("*Model") or block.name.startswith("*Paper"):
            continue
        pre = len(found)
        find_in_entities(block, f"block:{block.name}", found)
        post = len(found)
        if post > pre:
            print(f"  block '{block.name}': {post-pre} hits")

    # Итоги
    print(f"\n=========================")
    print(f"ВСЕГО найдено MTEXT с маркой+длиной: {len(found)}")

    by_brand = defaultdict(list)
    for f in found:
        by_brand[f["brand"]].append(f["length_m"])

    print("\n--- Сводка по маркам (до дедупликации) ---")
    for brand, lens in sorted(by_brand.items()):
        print(f"  {brand}: {len(lens)} шт, сумма {sum(lens):.0f} м")

    # Грубая дедупликация по (brand, x, y, length_m) — один MTEXT может быть
    # дублирован в разных layouts / blocks
    unique = {}
    for f in found:
        key = (f["brand"], round(f["x"] or 0, 1), round(f["y"] or 0, 1), round(f["length_m"], 1))
        if key not in unique:
            unique[key] = f

    print(f"\n--- После дедупликации по (brand, x, y, length): {len(unique)} уникальных ---")
    by_brand_u = defaultdict(list)
    for f in unique.values():
        by_brand_u[f["brand"]].append(f["length_m"])

    for brand, lens in sorted(by_brand_u.items()):
        print(f"  {brand}: {len(lens)} шт, сумма {sum(lens):.0f} м")

    print(f"\nЭталон:")
    print(f"  FRLS 3x1,5: 4460 м")
    print(f"  FRHF 3x1,5: 1300 м")

    # Показать по какому layout/block/source сколько
    print("\n--- Распределение по источнику (уникальные) ---")
    by_source = defaultdict(int)
    for f in unique.values():
        by_source[f["source"]] += 1
    for src, cnt in sorted(by_source.items(), key=lambda x: -x[1])[:15]:
        print(f"  {src}: {cnt}")


if __name__ == "__main__":
    main()
