#!/usr/bin/env python3
"""pdf_vor_nospec.py -- Plan-driven VOR generation for projects WITHOUT a СО spec.

When a project has no separate specification (СО) PDF, equipment quantities
cannot be read from a spec table. This module derives them by counting symbols
on the drawing plan pages -- text markers plus OpenCV visual template matching,
including marker-less pictograms (switches/sockets/exit signs that have no
numeric symbol) -- then synthesizes spec-like items and builds VOR sections
through the existing section builders.

Design: fully isolated from generate_vor_from_pdfs() so the authoritative
СО pipeline is left untouched (no regression). Activated only for no-СО input.

Usage:
    from pdf_vor_nospec import generate_vor_from_plans
    sections = generate_vor_from_plans("path/to/combined.pdf")
    # or a folder of plan PDFs
"""
from __future__ import annotations

import logging
import os
from collections import defaultdict
from pathlib import Path

from equipment_counter import SpecItem
from pdf_vor_pipeline import (
    PlanCountResult,
    VorSection,
    _build_devices_section,
    _build_lighting_section,
    build_height_ratios,
    elevation_to_height,
    _extract_elevations,
)

_log = logging.getLogger(__name__)

# VOR equipment buckets that we point-count from plan symbols.
_POINT_BUCKETS = ("luminaire", "indicator", "pictogram", "switch", "socket", "box")
_SKIP = "skip"  # line/area/panel features -- not point-counted on plans


def _route_bucket(desc: str, cat: str = "") -> str:
    """Route an equipment description to a VOR bucket.

    Description keywords are authoritative (legend categories are coarse and
    items like junction boxes / cable outlets carry an empty category). Returns
    ``_SKIP`` for line/area/panel features that are not point-counted here.
    """
    d = (desc or "").lower()
    if "коробк" in d:
        return "box"
    if "выключател" in d:
        return "switch"
    if "розетк" in d:
        return "socket"
    if "эвакуац" in d or "указател" in d or "выход" in d:
        return "indicator"
    if "светильник" in d:
        return "luminaire"
    # Panels, cable runs, wiring, trays, outlets -> not point-counted here.
    return _SKIP


# Line/area features whose legend symbol is a stroke (not a countable point):
# template-matching these produces noise, so they are never counted.
_LINE_FEATURE_KW = ("проводка", "кабельная трасса", "трасса", "лоток")


def _is_line_feature(desc: str, cat: str = "") -> bool:
    d = (desc or "").lower()
    if cat in ("проводка", "кабельная трасса", "лоток"):
        return True
    return any(k in d for k in _LINE_FEATURE_KW)


def _count_one_plan(
    path: str, legend, log=print, include_all: bool = False,
) -> dict[str, int]:
    """Merge text + visual symbol counts for a single plan PDF.

    Mirrors the merge strategy of pdf_vor_pipeline.count_equipment_on_plans:
      - legend items WITHOUT a text marker (sym == "") -> visual count only
        (this is the marker-less pictogram path: switches/sockets/exits)
      - legend items WITH a marker -> text count (precise); fall back to
        visual only when text found nothing.

    ``include_all=True`` counts every point symbol (panels, outlets, ...),
    skipping only line/area features — used by the main pipeline's no-СО
    synth, which routes the raw counts via classify_spec_item.  Default
    (False) keeps only the lighting/devices buckets for this module's own
    section assembly.

    Returns model_description -> count.
    """
    from pdf_count_text import count_symbols

    enable_visual = os.environ.get("VOR_VISUAL", "1") == "1"
    txt = count_symbols(path, legend)

    visual_counts: dict[int, int] = {}
    if enable_visual:
        try:
            from pdf_count_visual import match_symbols
            vis = match_symbols(path, legend_result=legend, page=legend.page_index)
            visual_counts = vis.counts
        except Exception as exc:  # noqa: BLE001
            log(f"    visual matching failed: {exc}")

    model_counts: dict[str, int] = defaultdict(int)
    for idx, item in enumerate(legend.items):
        if not item.description:
            continue
        if include_all:
            if _is_line_feature(item.description, item.category):
                continue
        elif _route_bucket(item.description, item.category) == _SKIP:
            continue
        sym = item.symbol or ""
        vis_count = visual_counts.get(idx, 0)
        txt_count = txt.counts.get(sym, 0) if sym else 0

        if not sym:
            count = vis_count          # marker-less -> visual is the only source
        elif txt_count == 0 and vis_count > 0:
            count = vis_count
        else:
            count = txt_count

        if count > 0:
            model_counts[item.description] += count
    return dict(model_counts)


def _resolve_plan_pdfs(source: str | Path) -> list[Path]:
    """A single PDF -> [pdf]; a folder -> all *.pdf inside (recursive)."""
    p = Path(source)
    if p.is_file() and p.suffix.lower() == ".pdf":
        return [p]
    if p.is_dir():
        return sorted(p.rglob("*.pdf"))
    return []


def count_plans_nospec(
    source: str | Path, log=print, include_all: bool = False,
) -> tuple[list[PlanCountResult], dict[str, str]]:
    """Count equipment on every plan PDF/page under *source* (no-СО mode).

    Returns (plan_results, desc_to_category) where desc_to_category maps each
    seen equipment description to its legend category. ``include_all`` is
    forwarded to :func:`_count_one_plan` (count panels/outlets too).
    """
    import pdfplumber
    from pdf_legend_parser import parse_legend

    results: list[PlanCountResult] = []
    desc_to_cat: dict[str, str] = {}

    for pdf in _resolve_plan_pdfs(source):
        try:
            with pdfplumber.open(str(pdf)) as doc:
                n_pages = len(doc.pages)
        except Exception as exc:  # noqa: BLE001
            log(f"  cannot open {pdf.name}: {exc}")
            continue

        # Combined no-СО PDFs put a different legend on every sheet
        # (power / lighting / grounding). Sweep pages individually so each
        # sheet's own legend scopes what is counted ON that sheet.
        file_elevs = _extract_elevations(pdf.name)
        for pg in range(n_pages):
            legend = parse_legend(str(pdf), restrict_page=pg)
            if not legend or not legend.items:
                continue
            for it in legend.items:
                if it.description and it.description not in desc_to_cat:
                    desc_to_cat[it.description] = it.category or ""
            counts = _count_one_plan(str(pdf), legend, log, include_all=include_all)
            if not counts:
                continue
            total = sum(counts.values())
            tag = f"{pdf.name}#p{pg}"
            if file_elevs:
                share = {m: max(1, c // len(file_elevs)) for m, c in counts.items()}
                for elev in file_elevs:
                    results.append(
                        PlanCountResult(tag, elev, elevation_to_height(elev), share))
            else:
                results.append(PlanCountResult(tag, 0.0, "до 5 метров", counts))
            log(f"  {pdf.name} p{pg}: {len(counts)} models, {total} items")
    return results, desc_to_cat


def _bucket_items(
    results: list[PlanCountResult], desc_to_cat: dict[str, str],
) -> dict[str, dict[str, int]]:
    """Aggregate plan counts into VOR buckets: bucket -> {description: total}."""
    buckets: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    totals: dict[str, int] = defaultdict(int)
    for r in results:
        for desc, c in r.counts.items():
            totals[desc] += c
    for desc, total in totals.items():
        bucket = _route_bucket(desc, desc_to_cat.get(desc, ""))
        if bucket == _SKIP:
            continue
        buckets[bucket][desc] = total
    return {b: dict(d) for b, d in buckets.items()}


def _synth_specitems(by_desc: dict[str, int]) -> list[SpecItem]:
    """Build synthetic spec items (unit шт) from plan-derived totals."""
    out: list[SpecItem] = []
    for i, (desc, qty) in enumerate(sorted(by_desc.items(), key=lambda kv: -kv[1]), 1):
        out.append(SpecItem(
            position=str(i), description=desc, model="",
            catalog_code="", supplier="", unit="шт", quantity=int(qty),
        ))
    return out


def generate_vor_from_plans(source: str | Path, log=print) -> list[VorSection]:
    """Generate VOR sections purely from drawing plans (no СО spec).

    Returns list[VorSection]; currently emits the lighting and devices
    sections derived from on-plan symbol counts.
    """
    log("No-СО mode: deriving quantities from plan symbol counts")
    results, desc_to_cat = count_plans_nospec(source, log)
    if not results:
        log("  No plan counts found")
        return []

    height_ratios = build_height_ratios(results)
    buckets = _bucket_items(results, desc_to_cat)

    # Single-floor when no real elevations were detected on any sheet.
    is_single_floor = all(r.elevation == 0.0 for r in results)

    sections: list[VorSection] = []

    lighting = _build_lighting_section(
        _synth_specitems(buckets.get("luminaire", {})),
        _synth_specitems(buckets.get("indicator", {})),
        _synth_specitems(buckets.get("pictogram", {})),
        [],
        height_ratios,
        log,
        is_single_floor=is_single_floor,
        has_trays=False,
        plan_counts=getattr(height_ratios, "raw_counts", None),
    )
    if lighting.rows:
        sections.append(lighting)

    # Switches + sockets share the devices section (the builder routes
    # "розетк"/"выключател" internally); junction boxes go to its box slot.
    devices = _build_devices_section(
        _synth_specitems({**buckets.get("switch", {}), **buckets.get("socket", {})}),
        _synth_specitems(buckets.get("box", {})),
        [], log,
    )
    if devices.rows:
        sections.append(devices)

    total_rows = sum(len(s.rows) for s in sections)
    log(f"No-СО VOR: {len(sections)} sections, {total_rows} rows")
    return sections


def render_nospec_docx(sections: list[VorSection]) -> bytes:
    """Render no-СО VOR sections to a ДБТ-format .docx (bytes).

    The renderer re-sections rows by name keyword, so we just flatten our
    section rows into the aggregate dict format it expects.
    """
    from vor_docx_renderer import render_vor_docx

    aggregated: list[dict] = []
    for sec in sections:
        for row in sec.rows:
            aggregated.append({
                "name": row.get("name", ""),
                "unit": row.get("unit", "шт"),
                "total": row.get("qty", 0),
                "drawing_refs": row.get("drawing_ref", ""),
            })
    return render_vor_docx(aggregated)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    if len(sys.argv) < 2:
        print("Usage: python pdf_vor_nospec.py <pdf_or_folder> [out.docx]")
        sys.exit(1)
    secs = generate_vor_from_plans(sys.argv[1])
    for s in secs:
        print(f"\n## {s.title}")
        for row in s.rows:
            mark = "    " if row.get("is_material") else "  "
            print(f"{mark}{row['name'][:60]:62} {row['unit']:5} {row['qty']}")
    if len(sys.argv) >= 3 and secs:
        out = sys.argv[2]
        Path(out).write_bytes(render_nospec_docx(secs))
        print(f"\nWrote {out}")
    print("\nNOTE: counts for marker-less symbols are VISUAL estimates "
          "(approximate); cross-check against schema/spec where available.")
