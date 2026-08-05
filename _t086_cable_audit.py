"""T086 (T102-investigate-cable-overcount) — diagnostic audit.

Reads .tayfa/common/discussions/T066_items_cache.json (per-PDF item dumps
produced by T073) and produces a structured per-PDF cable-contribution
report.  Goal: explain ~6500 m overcount (measured 16692 m vs spec 10226 m).

Two hypotheses to test:
  (a) Same cable line counted on 005 vs 012 PDFs (or similar duplicates).
  (b) 'трасса' (cable raceway / route) being summed into кабель totals.

Outputs:
  .tayfa/common/discussions/T086_cable_audit.json   (machine-readable)
  .tayfa/common/discussions/T086_cable_audit.md     (human-readable)
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
CACHE = ROOT / ".tayfa" / "common" / "discussions" / "T066_items_cache.json"
OUT_JSON = ROOT / ".tayfa" / "common" / "discussions" / "T086_cable_audit.json"
OUT_MD = ROOT / ".tayfa" / "common" / "discussions" / "T086_cable_audit.md"

# Spec ground truth from boss T086 message.
SPEC_BREAKDOWN = {
    "ВБШвнг-LS": 2317 + 1127,    # 3444
    "ВБШвнг-FRLS": 4460 + 180,   # 4640
    "ППГнг-HF": 842,
    "ППГнг-FRHF": 1300,
}
SPEC_SUM = sum(SPEC_BREAKDOWN.values())  # 10226

# What we consider a "real cable" vs a "raceway/трасса".
TRASSA_TOKENS_RE = re.compile(r"трасс", re.IGNORECASE)
# A real cable brand name appears as ВБШвнг / ППГнг / ВВГнг etc.
CABLE_BRAND_RE = re.compile(r"(ВБШвнг|ППГнг|ВВГнг|ВВГ|КПСнг|КПС|КСБнг)", re.IGNORECASE)


def classify_kind(item: dict[str, Any]) -> str:
    """Return one of: 'cable_real', 'cable_trassa', 'cable_other'.

    Logic
    -----
    - category != 'cable' -> not considered here (skipped).
    - name/work_name/equipment_name contains 'трасс' AND no cable brand
      -> cable_trassa  (this is hypothesis (b)).
    - name has recognised cable brand -> cable_real.
    - everything else under category=cable -> cable_other (suspect).
    """
    fields = [
        str(item.get("name") or ""),
        str(item.get("work_name") or ""),
        str(item.get("equipment_name") or ""),
        str(item.get("symbol") or ""),
    ]
    blob = " | ".join(fields)
    has_trassa = bool(TRASSA_TOKENS_RE.search(blob))
    has_brand = bool(CABLE_BRAND_RE.search(blob))
    if has_brand:
        return "cable_real"
    if has_trassa:
        return "cable_trassa"
    return "cable_other"


def safe_float(x: Any) -> float:
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def main() -> None:
    raw = json.loads(CACHE.read_text(encoding="utf-8"))
    # raw: dict[pdf_name, list[item]]

    # Per-PDF accumulation.
    per_pdf: dict[str, dict[str, Any]] = {}
    # Per-name-kind global accumulation (for dedupe detection).
    name_key_to_pdfs: dict[str, list[tuple[str, float]]] = defaultdict(list)

    grand_total_real = 0.0
    grand_total_trassa = 0.0
    grand_total_other = 0.0

    for pdf_name, items in raw.items():
        bucket = {
            "pdf": pdf_name,
            "cable_real_m": 0.0,
            "cable_trassa_m": 0.0,
            "cable_other_m": 0.0,
            "rows": [],
        }
        for it in items:
            if str(it.get("category") or "").lower() != "cable":
                continue
            kind = classify_kind(it)
            metres = safe_float(it.get("total"))
            if metres <= 0:
                # Sometimes count is metres -- fallback.
                metres = safe_float(it.get("count"))
            if metres <= 0:
                continue

            name = str(it.get("name") or "")
            bucket["rows"].append({
                "name": name,
                "kind": kind,
                "metres": round(metres, 2),
                "source": it.get("source"),
                "route": it.get("route"),
            })
            if kind == "cable_real":
                bucket["cable_real_m"] += metres
                grand_total_real += metres
            elif kind == "cable_trassa":
                bucket["cable_trassa_m"] += metres
                grand_total_trassa += metres
            else:
                bucket["cable_other_m"] += metres
                grand_total_other += metres

            # Track for dedupe analysis.
            norm_name = re.sub(r"\s+", " ", name.strip().lower())
            name_key_to_pdfs[norm_name].append((pdf_name, round(metres, 2)))

        bucket["cable_real_m"] = round(bucket["cable_real_m"], 2)
        bucket["cable_trassa_m"] = round(bucket["cable_trassa_m"], 2)
        bucket["cable_other_m"] = round(bucket["cable_other_m"], 2)
        bucket["cable_total_m"] = round(
            bucket["cable_real_m"] + bucket["cable_trassa_m"] + bucket["cable_other_m"],
            2,
        )
        per_pdf[pdf_name] = bucket

    grand_total = grand_total_real + grand_total_trassa + grand_total_other

    # Dedupe detection: same normalised name appearing on multiple PDFs with
    # non-trivial metres.
    suspected_dups = []
    for name, hits in name_key_to_pdfs.items():
        if len(hits) < 2:
            continue
        total_m = round(sum(m for _, m in hits), 2)
        if total_m < 50:  # ignore noise
            continue
        suspected_dups.append({
            "name": name,
            "appearances": len(hits),
            "total_m": total_m,
            "per_pdf": [{"pdf": p, "m": m} for p, m in hits],
        })
    suspected_dups.sort(key=lambda r: -r["total_m"])

    summary = {
        "spec_breakdown_m": SPEC_BREAKDOWN,
        "spec_sum_m": SPEC_SUM,
        "measured_grand_total_m": round(grand_total, 2),
        "measured_cable_real_m": round(grand_total_real, 2),
        "measured_cable_trassa_m": round(grand_total_trassa, 2),
        "measured_cable_other_m": round(grand_total_other, 2),
        "delta_vs_spec_m": round(grand_total - SPEC_SUM, 2),
        "delta_real_only_vs_spec_m": round(grand_total_real - SPEC_SUM, 2),
        "n_pdfs_with_cable": sum(
            1 for b in per_pdf.values() if b["cable_total_m"] > 0
        ),
        "n_suspected_dup_names": len(suspected_dups),
    }

    # Top contributing PDFs.
    pdf_list = sorted(
        per_pdf.values(), key=lambda b: -b["cable_total_m"]
    )
    top_pdfs = [b for b in pdf_list if b["cable_total_m"] > 0]

    out = {
        "summary": summary,
        "per_pdf_top": top_pdfs,
        "suspected_duplicate_names_top20": suspected_dups[:20],
    }

    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Human-readable markdown.
    lines: list[str] = []
    lines.append("# T086 — Cable overcount audit\n")
    lines.append("## Summary\n")
    lines.append(f"- Spec breakdown (m): {SPEC_BREAKDOWN}")
    lines.append(f"- Spec sum (m): **{SPEC_SUM}**")
    lines.append(f"- Measured grand total (m): **{summary['measured_grand_total_m']}**")
    lines.append(
        f"- Measured by kind: real cable={summary['measured_cable_real_m']}, "
        f"трассы={summary['measured_cable_trassa_m']}, "
        f"other={summary['measured_cable_other_m']}"
    )
    lines.append(
        f"- Delta vs spec (total): **{summary['delta_vs_spec_m']:+.2f} m**"
    )
    lines.append(
        f"- Delta vs spec (real-only, hypothesis-b fix): "
        f"**{summary['delta_real_only_vs_spec_m']:+.2f} m**"
    )
    lines.append(f"- PDFs contributing cable rows: {summary['n_pdfs_with_cable']}")
    lines.append(f"- Suspected duplicate-name groups: {summary['n_suspected_dup_names']}\n")

    lines.append("## Per-PDF cable contributions (sorted by total)\n")
    lines.append("| PDF | real m | трассы m | other m | total m |")
    lines.append("|-----|-------:|---------:|--------:|--------:|")
    for b in top_pdfs:
        lines.append(
            f"| {b['pdf']} | {b['cable_real_m']:.2f} | "
            f"{b['cable_trassa_m']:.2f} | {b['cable_other_m']:.2f} | "
            f"{b['cable_total_m']:.2f} |"
        )

    lines.append("\n## Suspected duplicate-name groups (top 20)\n")
    if not suspected_dups:
        lines.append("_(none above 50 m noise floor)_")
    else:
        for d in suspected_dups[:20]:
            per_pdf_str = ", ".join(
                f"{x['pdf']}={x['m']}m" for x in d["per_pdf"]
            )
            lines.append(
                f"- **{d['name']}** — {d['appearances']} PDFs, "
                f"total {d['total_m']} m → {per_pdf_str}"
            )

    lines.append("\n## Verdict\n")
    if summary["measured_cable_trassa_m"] > 1000:
        lines.append(
            f"- Hypothesis (b) **CONFIRMED**: трассы contribute "
            f"{summary['measured_cable_trassa_m']:.2f} m, which alone "
            f"explains {(summary['measured_cable_trassa_m']/abs(summary['delta_vs_spec_m']) * 100 if summary['delta_vs_spec_m'] else 0):.1f}% "
            f"of the {abs(summary['delta_vs_spec_m']):.0f} m overcount."
        )
    else:
        lines.append("- Hypothesis (b): трассы contribute <1000 m — not the main cause.")

    if abs(summary["delta_real_only_vs_spec_m"]) < abs(summary["delta_vs_spec_m"]) * 0.3:
        lines.append(
            "- After excluding трассы, real-cable total is within "
            f"{abs(summary['delta_real_only_vs_spec_m']):.0f} m of spec — "
            "patching the composer to exclude трассы closes most of the gap."
        )
    else:
        lines.append(
            "- Even after excluding трассы, real-cable total deviates "
            f"by {summary['delta_real_only_vs_spec_m']:+.0f} m — hypothesis (a) "
            "double-count likely contributes; see duplicate-name groups above."
        )

    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")

    # Console summary.
    print("=== T086 cable audit ===")
    print(f"spec_sum_m              = {SPEC_SUM}")
    print(f"measured_grand_total_m  = {summary['measured_grand_total_m']}")
    print(f"  cable_real_m          = {summary['measured_cable_real_m']}")
    print(f"  cable_trassa_m        = {summary['measured_cable_trassa_m']}")
    print(f"  cable_other_m         = {summary['measured_cable_other_m']}")
    print(f"delta_total_vs_spec_m   = {summary['delta_vs_spec_m']:+.2f}")
    print(f"delta_real_only_vs_spec = {summary['delta_real_only_vs_spec_m']:+.2f}")
    print(f"suspected_dup_groups    = {summary['n_suspected_dup_names']}")
    print(f"Wrote: {OUT_JSON}")
    print(f"Wrote: {OUT_MD}")


if __name__ == "__main__":
    main()
