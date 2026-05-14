"""T086 follow-up: enumerate every distinct cable-category name."""
import json
from collections import Counter
from pathlib import Path

cache = Path(__file__).resolve().parent / ".tayfa" / "common" / "discussions" / "T066_items_cache.json"
data = json.loads(cache.read_text(encoding="utf-8"))

name_counts: Counter = Counter()
name_totals: Counter = Counter()
name_to_pdfs: dict = {}
for pdf, items in data.items():
    for it in items:
        if str(it.get("category") or "").lower() != "cable":
            continue
        nm = (it.get("name") or "").strip()
        try:
            m = float(it.get("total") or 0)
        except (TypeError, ValueError):
            m = 0.0
        name_counts[nm] += 1
        name_totals[nm] += m
        name_to_pdfs.setdefault(nm, []).append((pdf, m))

print(f"=== Distinct cable-category names ({len(name_counts)} unique) ===")
for nm, cnt in name_counts.most_common(50):
    print(f"  rows={cnt:3d}  total_m={name_totals[nm]:10.2f}  | {nm}")

print("\n=== Names with metres > 100 — per-PDF breakdown ===")
for nm, total in sorted(name_totals.items(), key=lambda x: -x[1]):
    if total < 100:
        continue
    print(f"\n[{nm}] total={total:.2f} m across {name_counts[nm]} rows")
    for pdf, m in name_to_pdfs[nm]:
        print(f"    {pdf}: {m:.2f} m")
