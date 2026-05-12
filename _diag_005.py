"""Diagnose PDF 005 parse: show legend + per-phase counts."""
from __future__ import annotations

import sys
import json
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

PDF_PATH = Path(r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")

if not PDF_PATH.exists():
    print(f"[ERR] PDF not found: {PDF_PATH}")
    sys.exit(1)

print(f"PDF: {PDF_PATH.name}")
print(f"Size: {PDF_PATH.stat().st_size} bytes")
print("=" * 80)

# Phase 1: Legend
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF_PATH))
print(f"\n[PHASE 1] parse_legend -> {len(legend.items)} items:")
for i, item in enumerate(legend.items):
    sym = item.symbol or "(no-sym)"
    desc = item.description or "(no-desc)"
    print(f"  [{i:2}] symbol={sym!r:15} desc={desc!r}")

# Phase 2: Text counts
from pdf_count_text import count_symbols

print(f"\n[PHASE 2] count_symbols ->")
try:
    text_result = count_symbols(str(PDF_PATH), legend)
    for sym, cnt in sorted(text_result.counts.items()):
        print(f"  symbol {sym!r:15} -> {cnt}")
except Exception as e:
    print(f"  [ERR] {type(e).__name__}: {e}")
    text_result = None

# Phase 3: Visual match (only for symbols missing from text)
from pdf_count_visual import match_symbols, detect_pictograms

missing = []
if text_result:
    for idx, item in enumerate(legend.items):
        sym = item.symbol or ""
        if not sym or text_result.counts.get(sym, 0) == 0:
            missing.append((idx, item))

print(f"\n[PHASE 3] match_symbols (visual) -> {len(missing)} items need it:")
for idx, it in missing:
    print(f"  [missing] [{idx:2}] symbol={(it.symbol or '(empty)')!r:15} desc={(it.description or '')!r}")

try:
    vis = match_symbols(str(PDF_PATH), legend)
    print(f"  visual counts: {vis.counts}")
except Exception as e:
    print(f"  [ERR] match_symbols: {type(e).__name__}: {e}")

# Phase 4: Pictograms
print(f"\n[PHASE 4] detect_pictograms ->")
try:
    picto = detect_pictograms(str(PDF_PATH), legend)
    for name, cnt in picto.counts.items():
        print(f"  {name!r}: {cnt}")
except Exception as e:
    print(f"  [ERR] {type(e).__name__}: {e}")

# Phase 5: Cables
from pdf_count_cables import extract_cables

print(f"\n[PHASE 5] extract_cables ->")
try:
    cables = extract_cables(str(PDF_PATH), legend)
    sched = getattr(cables, "cable_schedule", None) or []
    print(f"  cable_schedule: {len(sched)} entries")
    for e in sched[:10]:
        print(f"    {e}")
    derived = getattr(cables, "derived_work_items", None) or []
    print(f"  derived_work_items: {len(derived)}")
    for d in derived[:10]:
        print(f"    {d}")
except Exception as e:
    print(f"  [ERR] {type(e).__name__}: {e}")

# Final orchestration: reuse test_vor_cross_format helper
print(f"\n[FINAL] _count_equipment_in_pdf(...) result:")
sys.path.insert(0, str(Path.cwd()))
import importlib.util

# T047: register the module in sys.modules BEFORE exec_module so that
# @dataclass decorators inside the loaded file can resolve
# `sys.modules.get(cls.__module__)` -> module object (not None).
# Without this, Python 3.12 raises
# `AttributeError: 'NoneType' object has no attribute '__dict__'`
# from dataclasses._is_type during decoration.
spec = importlib.util.spec_from_file_location("tvcf", "test_vor_cross_format.py")
tvcf = importlib.util.module_from_spec(spec)
sys.modules["tvcf"] = tvcf
spec.loader.exec_module(tvcf)

items = tvcf._count_equipment_in_pdf(str(PDF_PATH))
print(f"  {len(items)} items returned:")
for it in items:
    wn = it.get("work_name", "")
    rn = it.get("name", "")
    tot = it.get("total", 0)
    unit = it.get("unit", "")
    print(f"    [{tot:>6} {unit:3}] work_name={wn!r:60} raw_name={rn!r}")
