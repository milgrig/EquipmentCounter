#!/usr/bin/env python3
"""
Прогон полного парсера по ВСЕМ 25 DXF третьей захватки ГПК.
Цель: увидеть, нет ли кабелей в других файлах (СО.dxf, планы кабеленесущих, ...).
"""
from __future__ import annotations

import io
import re
import sys
from collections import defaultdict
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

DIR = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG"
)

BRAND_PATTERNS = [
    ("ВБШвнг(А)-LS 3х1,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*LS\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ВБШвнг(А)-LS 3х2,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*LS\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
    ("ВБШвнг(А)-FRLS 3х1,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ВБШвнг(А)-FRLS 3х2,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-HF 3х1,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*HF\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-HF 3х2,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*HF\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-FRHF 3х1,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-FRHF 3х2,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
]

L_RE = re.compile(r"L\s*=\s*(\d{1,5}(?:[,.]\d+)?)", re.IGNORECASE)


def iter_entities(path: Path):
    section = "?"
    with path.open("r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()
    i = 0
    n = len(lines)
    current_entity: dict | None = None
    current_type: str | None = None
    while i < n - 1:
        code_line = lines[i].strip()
        value_line = lines[i + 1].rstrip("\n").rstrip("\r")
        try:
            code = int(code_line)
        except ValueError:
            i += 1
            continue
        i += 2
        if code == 0:
            if current_entity is not None and current_type is not None:
                yield current_type, section, current_entity
            current_type = value_line.strip()
            current_entity = {}
            if current_type == "SECTION":
                if i < n - 1 and lines[i].strip() == "2":
                    section = lines[i + 1].strip()
                    i += 2
                current_entity = None
                current_type = None
            elif current_type == "ENDSEC":
                section = "?"
                current_entity = None
                current_type = None
        else:
            if current_entity is not None:
                current_entity.setdefault(code, []).append(value_line)
    if current_entity is not None and current_type is not None:
        yield current_type, section, current_entity


def strip_mtext_codes(text: str) -> str:
    text = re.sub(r"\\[A-Za-z]\s*[^\\;]*;?", "", text)
    text = re.sub(r"\\P", " ", text)
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def mtext_content(e: dict) -> str:
    parts: list[str] = []
    parts.extend(e.get(3, []))
    parts.extend(e.get(1, []))
    return "".join(parts)


def scan_file(path: Path) -> list[dict]:
    cables: list[dict] = []
    for etype, section, e in iter_entities(path):
        if etype not in ("MTEXT", "TEXT"):
            continue
        if etype == "MTEXT":
            raw = mtext_content(e)
        else:
            raw = (e.get(1, [""])[0])
        if not raw:
            continue
        cleaned = strip_mtext_codes(raw) if etype == "MTEXT" else raw.strip()
        brand = None
        for bname, pat in BRAND_PATTERNS:
            if pat.search(cleaned):
                brand = bname
                break
        if not brand:
            continue
        lm = L_RE.search(cleaned)
        if not lm:
            continue
        try:
            length = float(lm.group(1).replace(",", "."))
        except ValueError:
            continue
        try:
            x = float(e.get(10, ["0"])[0])
            y = float(e.get(20, ["0"])[0])
        except (ValueError, IndexError):
            x, y = 0.0, 0.0
        cables.append({
            "brand": brand, "length_m": length,
            "x": x, "y": y, "section": section, "etype": etype,
        })
    return cables


def main() -> None:
    dxf_files = sorted(DIR.glob("*.dxf"))
    print(f"DXF файлов: {len(dxf_files)}\n")

    global_by_brand: dict[tuple[str, float, float, float], dict] = {}
    header = f"{'Файл':<55} | {'FRLS-1,5':>10} {'FRHF-1,5':>10} {'LS-1,5':>10} {'HF-1,5':>10} {'всего':>10}"
    print(header)
    print("-" * len(header))

    for path in dxf_files:
        cables = scan_file(path)
        if not cables:
            print(f"{path.name[:53]:<55} | {'-':>10} {'-':>10} {'-':>10} {'-':>10} {'-':>10}")
            continue

        by_brand = defaultdict(lambda: [0, 0.0])
        for c in cables:
            by_brand[c["brand"]][0] += 1
            by_brand[c["brand"]][1] += c["length_m"]

        # Агрегируем в общий
        for c in cables:
            key = (c["brand"], round(c["x"], 1), round(c["y"], 1), round(c["length_m"], 1))
            if key not in global_by_brand:
                global_by_brand[key] = c

        total_m = sum(v[1] for v in by_brand.values())
        name = path.name
        if len(name) > 53:
            name = name[:52] + "…"

        def fmt(br):
            v = by_brand.get(br)
            return f"{v[1]:.0f}" if v else "-"

        print(f"{name:<55} | "
              f"{fmt('ВБШвнг(А)-FRLS 3х1,5'):>10} "
              f"{fmt('ППГнг(А)-FRHF 3х1,5'):>10} "
              f"{fmt('ВБШвнг(А)-LS 3х1,5'):>10} "
              f"{fmt('ППГнг(А)-HF 3х1,5'):>10} "
              f"{total_m:>10.0f}")

    print("-" * len(header))
    # Итого уникальных (с учётом глобальной дедупликации между файлами)
    by = defaultdict(lambda: [0, 0.0])
    for v in global_by_brand.values():
        by[v["brand"]][0] += 1
        by[v["brand"]][1] += v["length_m"]

    print(f"\n--- Глобальная дедупликация (по всем файлам) ---")
    etalon = {
        "ВБШвнг(А)-LS 3х1,5": 2317, "ВБШвнг(А)-LS 3х2,5": 1127,
        "ВБШвнг(А)-FRLS 3х1,5": 4460, "ВБШвнг(А)-FRLS 3х2,5": 180,
        "ППГнг(А)-HF 3х1,5": 842, "ППГнг(А)-HF 3х2,5": 0,
        "ППГнг(А)-FRHF 3х1,5": 1300, "ППГнг(А)-FRHF 3х2,5": 0,
    }
    our_sum = 0
    print(f"  {'Марка':30} {'Найдено':>10} {'Эталон':>10} {'Δ':>10}")
    for br in sorted(set(list(by.keys()) + list(etalon.keys()))):
        our = by.get(br, [0, 0])[1]
        et = etalon.get(br, 0)
        our_sum += our
        if et:
            diff = our - et
            pct = diff / et * 100
            print(f"  {br:30} {our:>10.0f} {et:>10} {diff:>+10.0f}  ({pct:+.0f}%)")
        else:
            print(f"  {br:30} {our:>10.0f} {et:>10} {our - et:>+10.0f}")
    et_sum = sum(etalon.values())
    print(f"  {'ИТОГО':30} {our_sum:>10.0f} {et_sum:>10} {our_sum - et_sum:>+10.0f}  ({(our_sum-et_sum)/et_sum*100:+.1f}%)")


if __name__ == "__main__":
    main()
