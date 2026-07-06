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
import re
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
    _vor_worker_count,
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


# ---------------------------------------------------------------------------
# Tray (cable-carrying systems) metraж — text callout parser.
#
# On "План кабеленесущих систем" sheets each tray segment is annotated in
# plain text next to its run as "<size> <length> мм", e.g. "50х50 6000 мм"
# or "100х80 4500 мм". The drawing layers carry no usable per-line attribute
# (single stroke colour, no dash, no OCG), so geometry cannot isolate tray
# runs — but these text callouts give a precise, summable length per size.
# This recovers the linear metraж the point-counter throws away as _SKIP.
# ---------------------------------------------------------------------------

# "50х50 6000 мм"  /  "100 х 80  4500 мм"  (Cyrillic х and Latin x; мм/mm)
_TRAY_CALLOUT_RE = re.compile(
    r"(\d{2,3})\s*[хx]\s*(\d{2,3})\s+(\d{3,5})\s*мм",
    re.IGNORECASE,
)
# Plausible tray cross-sections (mm). Guards against matching unrelated
# "WxH NNNN мм" dimension strings (door/window/grid sizes).
_TRAY_WIDTHS = {50, 80, 100, 150, 200, 300, 400, 500, 600}
_TRAY_HEIGHTS = {35, 50, 60, 80, 100, 150, 200}
# A single tray segment longer than this (mm) is almost surely a building
# dimension mis-read, not one length of tray; drop it.
_TRAY_SEG_MAX_MM = 12000


def _parse_tray_lengths(text: str) -> dict[str, list[int]]:
    """Extract tray segment lengths (mm) per size from a page's text.

    Returns {"50х50": [6000, 5800, ...], "100х80": [...], ...}. Only
    size/length pairs with a plausible tray cross-section and a sane segment
    length are kept, so stray dimension strings don't inflate the metraж.
    """
    out: dict[str, list[int]] = defaultdict(list)
    for m in _TRAY_CALLOUT_RE.finditer(text or ""):
        w, h, length = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if w not in _TRAY_WIDTHS or h not in _TRAY_HEIGHTS:
            continue
        if not (0 < length <= _TRAY_SEG_MAX_MM):
            continue
        out[f"{w}х{h}"].append(length)
    return {k: v for k, v in out.items()}


def _count_one_plan(
    path: str, legend, log=print, include_all: bool = False,
    force_visual: bool = False,
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

    # No-СО mode relies on visual matching for marker-less equipment
    # (grounding/lightning pictograms, luminaires) which carry no text marker.
    # force_visual overrides the global VOR_VISUAL=0 speed flag here, because
    # without visual the no-СО synthesis yields nothing for such sections.
    enable_visual = force_visual or os.environ.get("VOR_VISUAL", "1") == "1"
    txt = count_symbols(path, legend)

    visual_counts: dict[int, int] = {}
    if enable_visual:
        try:
            from pdf_count_visual import match_symbols
            vis = match_symbols(path, legend_result=legend, page=legend.page_index)
            visual_counts = vis.counts
        except Exception as exc:  # noqa: BLE001
            log(f"    visual matching failed: {exc}")

    # Якорный подсчёт глифов (S011): маркер на плане ставится на ГРУППУ
    # светильников, поэтому текстовый счётчик недосчитывает. Учимся виду
    # глифа у найденных маркеров и считаем сами глифы; берём max(text,
    # anchored) с защитным потолком от взрыва маски.
    anchored_counts: dict[str, int] = {}
    if enable_visual:
        try:
            from pdf_count_anchored import count_by_marker_anchors
            anch = count_by_marker_anchors(path, legend, txt)
            anchored_counts = anch.counts
        except Exception as exc:  # noqa: BLE001
            log(f"    anchored counting failed: {exc}")

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

        # Якорный подсчёт может только ПОДНЯТЬ количество маркированного
        # символа (нашлись немаркированные экземпляры группы). Потолок
        # 6×text+10 страхует от деградации маски на нетипичном листе.
        if sym:
            a = anchored_counts.get(sym, 0)
            if count < a <= count * 6 + 10:
                count = a

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


# ---------------------------------------------------------------------------
# Cable-run metraж (no-СО): vector engine + raster fallback.
#
# The point-counter above throws away line features, so without this pass
# the no-СО "cable" spec category stays empty and the «Кабельная продукция»
# VOR section is never built. pdf_cables_by_method measures red/blue routes
# per page (dedup, laying method, marks bound from on-plan annotations);
# PDFs whose routes live inside an embedded raster image get a fallback
# pass through cable_length.measure_cable_lengths_raster.
# ---------------------------------------------------------------------------

# Below this vector total (m) a PDF is considered "vector-empty" and is
# retried with the raster engine (routes drawn as an image stream).
_VECTOR_EMPTY_M = 1.0


def _raster_method_key(category_key: str) -> str:
    """Map a raster-engine category key to a laying-method key.

    cable_length category keys: cable_emergency / cable_working /
    wire_tray / wire_pipe_hidden / wire_pipe_open.
    """
    ck = (category_key or "").lower()
    if "tray" in ck or "lot" in ck or "лотк" in ck:
        return "lotok"
    if "pipe" in ck or "trub" in ck or "труб" in ck:
        return "gofra/truba"
    return "po_konstrukciyam"


# Листы, на которых растровому движку нечего мерить: легенда и цветные
# примеры «Общих данных»/схем дают ложный скелет (шумовые метры).
_RASTER_SKIP_NAME_RE = re.compile(
    r"общие\s+данные|\bОД\b|схем", re.IGNORECASE)


def measure_plan_cables(
    source: str | Path, plan_names: list[str] | None = None, log=print,
) -> tuple[dict[tuple[str, str], float], dict[str, float]]:
    """Измерить метраж кабельных трасс на планах (no-СО режим).

    ``plan_names`` — имена файлов-планов (освещение + привязки); ``None`` =
    все PDF под ``source``. Возвращает ``(by_mark, by_method)``:
      by_mark[(марка, сечение)] = метры (пустая марка — длина, не
      атрибутированная ни аннотацией, ни лист-уровневой маркой);
      by_method[способ прокладки] = метры.
    """
    from pdf_cables_by_method import measure_cables
    from pdf_legend_parser import parse_legend
    import pdfplumber

    pdfs = _resolve_plan_pdfs(source)
    if plan_names is not None:
        wanted = set(plan_names)
        pdfs = [p for p in pdfs if p.name in wanted]

    by_mark: dict[tuple[str, str], float] = defaultdict(float)
    by_method: dict[str, float] = defaultdict(float)

    raster_candidates: list[Path] = []
    for pdf in pdfs:
        try:
            with pdfplumber.open(str(pdf)) as doc:
                n_pages = len(doc.pages)
        except Exception as exc:  # noqa: BLE001
            log(f"    трассы: не открыть {pdf.name}: {exc}")
            continue
        pdf_total = 0.0
        for pg in range(n_pages):
            try:
                legend = parse_legend(str(pdf), restrict_page=pg)
            except Exception:  # noqa: BLE001
                legend = None
            try:
                rep = measure_cables(str(pdf), page_index=pg,
                                     legend_result=legend)
            except Exception as exc:  # noqa: BLE001
                log(f"    трассы: сбой на {pdf.name} p{pg}: {exc}")
                continue
            page_total = sum(rep.totals_m.values())
            if page_total <= 0:
                continue
            pdf_total += page_total
            for (_color, method), m in rep.totals_m.items():
                by_method[method] += m
            marked = 0.0
            for mk, m in rep.totals_by_mark_m.items():
                by_mark[mk] += m
                marked += m
            # Длина без пер-трассовой марки уходит на лист-уровневую марку
            # (самую упоминаемую в примечаниях), иначе — в безымянный ключ.
            unmarked = page_total - marked
            if unmarked > 0.5:
                sheet_mark, sheet_sec = "", ""
                if rep.sheet_cable_marks:
                    top = rep.sheet_cable_marks[0]
                    sheet_mark = top.get("mark") or ""
                    sheet_sec = top.get("section") or ""
                by_mark[(sheet_mark, sheet_sec)] += unmarked
            log(f"    {pdf.name} p{pg}: {page_total:.1f} м трасс "
                f"(масштаб {rep.scale_source})")
        if pdf_total < _VECTOR_EMPTY_M:
            raster_candidates.append(pdf)

    # Растровый fallback — только для PDF, где векторный движок не нашёл
    # трасс (характерно для листов, экспортированных картинкой). Листы
    # «Общие данные»/схемы пропускаем: трасс там нет, а цветная легенда
    # даёт ложный скелет.
    raster_candidates = [p for p in raster_candidates
                         if not _RASTER_SKIP_NAME_RE.search(p.stem)]
    if raster_candidates:
        try:
            from cable_length import measure_cable_lengths_raster
        except Exception as exc:  # noqa: BLE001
            log(f"    растровый движок недоступен ({exc}); "
                f"{len(raster_candidates)} PDF без метража трасс")
            return dict(by_mark), dict(by_method)
        for pdf in raster_candidates:
            try:
                rep = measure_cable_lengths_raster(str(pdf))
            except Exception as exc:  # noqa: BLE001
                log(f"    растровый сбой на {pdf.name}: {exc}")
                continue
            pdf_total = 0.0
            for e in rep.entries:
                if e.length_m <= 0:
                    continue
                pdf_total += e.length_m
                by_mark[("", "")] += e.length_m
                by_method[_raster_method_key(e.category_key)] += e.length_m
            if pdf_total > 0:
                log(f"    {pdf.name}: {pdf_total:.1f} м трасс (растр)")
    return dict(by_mark), dict(by_method)


def _nospec_one_pdf_worker(args):
    """Process one PDF (all pages) for no-СО counting.

    Standalone, picklable for ProcessPoolExecutor. Returns
    (results_tuples, desc_to_cat, log_lines, tray_lengths) where results_tuples
    is a list of (tag, elevation, height_cat, counts) so the parent can rebuild
    PlanCountResult objects in deterministic order, and tray_lengths maps each
    tray size to the list of segment lengths (mm) read from text callouts.
    """
    pdf_path, pdf_name, include_all, force_visual = args
    import pdfplumber
    from pdf_legend_parser import parse_legend

    out_results: list[tuple] = []
    out_desc: dict[str, str] = {}
    log_lines: list[str] = []
    tray_lengths: dict[str, list[int]] = defaultdict(list)

    try:
        doc = pdfplumber.open(pdf_path)
    except Exception as exc:  # noqa: BLE001
        log_lines.append(f"  cannot open {pdf_name}: {exc}")
        return (out_results, out_desc, log_lines, dict(tray_lengths))
    try:
        n_pages = len(doc.pages)
        file_elevs = _extract_elevations(pdf_name)
        for pg in range(n_pages):
            # Tray metraж is read from page text and is independent of the
            # legend gate below (tray callouts live on the plan, not the
            # legend), so harvest it first for every page.
            try:
                ptext = doc.pages[pg].extract_text() or ""
            except Exception:  # noqa: BLE001
                ptext = ""
            for size, lens in _parse_tray_lengths(ptext).items():
                tray_lengths[size].extend(lens)

            legend = parse_legend(pdf_path, restrict_page=pg)
            if not legend or not legend.items:
                continue
            for it in legend.items:
                if it.description and it.description not in out_desc:
                    out_desc[it.description] = it.category or ""
            counts = _count_one_plan(pdf_path, legend, lambda m: None,
                                     include_all=include_all,
                                     force_visual=force_visual)
            if not counts:
                continue
            total = sum(counts.values())
            tag = f"{pdf_name}#p{pg}"
            if file_elevs:
                share = {m: max(1, c // len(file_elevs)) for m, c in counts.items()}
                for elev in file_elevs:
                    out_results.append((tag, elev, elevation_to_height(elev), share))
            else:
                out_results.append((tag, 0.0, "до 5 метров", counts))
            log_lines.append(f"  {pdf_name} p{pg}: {len(counts)} models, {total} items")
    finally:
        doc.close()
    return (out_results, out_desc, log_lines, dict(tray_lengths))


def count_plans_nospec(
    source: str | Path, log=print, include_all: bool = False,
    force_visual: bool = True,
) -> tuple[list[PlanCountResult], dict[str, str], dict[str, list[int]]]:
    """Count equipment on every plan PDF/page under *source* (no-СО mode).

    Returns (plan_results, desc_to_category, tray_lengths) where
    desc_to_category maps each seen equipment description to its legend
    category and tray_lengths maps each tray size ("50х50") to the list of
    segment lengths (mm) read from on-plan text callouts. ``include_all`` is
    forwarded to :func:`_count_one_plan` (count panels/outlets too).

    ``force_visual`` defaults to True: in no-СО mode visual matching is the
    only source for marker-less equipment (grounding/lightning/luminaires),
    so it must run even when the global VOR_VISUAL=0 speed flag is set.
    """
    results: list[PlanCountResult] = []
    desc_to_cat: dict[str, str] = {}
    tray_lengths: dict[str, list[int]] = defaultdict(list)

    pdfs = _resolve_plan_pdfs(source)
    jobs = [(str(pdf), pdf.name, include_all, force_visual) for pdf in pdfs]

    # Each PDF is processed independently (open + per-page legend + count).
    # Heaviest no-СО pass: each worker renders large PDFs at high DPI for
    # visual matching, so cap tightly by RAM to avoid OOM / hangs.
    _w = _vor_worker_count(len(jobs) if jobs else 0, heavy=True)

    import time as _time
    _total = len(jobs)
    # raw[i] holds the result for jobs[i]; None until filled. This lets a single
    # crashed worker fall back to in-process counting for THAT pdf only, while
    # every already-completed result is kept (no catastrophic full restart).
    raw: list = [None] * _total
    _done = 0
    _t0 = _time.time()

    def _emit(name: str) -> None:
        _pct = int(_done * 100 / _total) if _total else 100
        _el = _time.time() - _t0
        _eta = (_el / _done) * (_total - _done) if _done else 0
        log(
            f"    ⏳ обработано {_done}/{_total} ({_pct}%) "
            f"— {name} | прошло {int(_el)}с, осталось ~{int(_eta)}с"
        )

    if _w > 1:
        import multiprocessing as _mp
        from concurrent.futures import ProcessPoolExecutor, as_completed
        log(f"  Sweeping {_total} PDFs (no-СО) across {_w} workers…")
        try:
            _ctx = _mp.get_context("spawn")
            with ProcessPoolExecutor(max_workers=_w, mp_context=_ctx) as ex:
                # Map each future back to its submission index so results stay
                # deterministic, but report live progress as each PDF finishes.
                _fut_to_idx = {
                    ex.submit(_nospec_one_pdf_worker, job): i
                    for i, job in enumerate(jobs)
                }
                for _fut in as_completed(_fut_to_idx):
                    _idx = _fut_to_idx[_fut]
                    try:
                        raw[_idx] = _fut.result()
                    except Exception as exc:  # noqa: BLE001
                        # A single worker died (e.g. abrupt termination). Retry
                        # just this PDF in-process below; keep the rest. Don't
                        # advance the progress counter — it's counted when the
                        # fallback loop actually finishes it.
                        log(f"  ⚠ worker failed on {jobs[_idx][1]} ({exc}); "
                            f"will retry in-process")
                        raw[_idx] = None
                        continue
                    _done += 1
                    _emit(jobs[_idx][1])
        except Exception as exc:  # noqa: BLE001
            # Pool-level failure (e.g. broken pool). Keep whatever completed;
            # the per-job fallback loop below fills the gaps in-process.
            log(f"  ⚠ Parallel no-СО pool failed ({exc}); "
                f"finishing remaining PDFs in-process")

    # Fill any gaps (workers that crashed, or _w<=1) by counting in-process.
    n_fallback = sum(1 for r in raw if r is None)
    if n_fallback and _w > 1:
        log(f"  Finishing {n_fallback} PDF(s) in-process…")
    for _i, job in enumerate(jobs):
        if raw[_i] is None:
            raw[_i] = _nospec_one_pdf_worker(job)
            _done += 1
            _emit(job[1])

    # Reassemble in deterministic submission order.
    for out_results, out_desc, log_lines, out_trays in raw:
        for d, c in out_desc.items():
            desc_to_cat.setdefault(d, c)
        for tag, elev, hcat, counts in out_results:
            results.append(PlanCountResult(tag, elev, hcat, counts))
        for size, lens in out_trays.items():
            tray_lengths[size].extend(lens)
        for ln in log_lines:
            log(ln)
    return results, desc_to_cat, dict(tray_lengths)


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


def _build_tray_metraj_section(
    tray_lengths: dict[str, list[int]], log=print,
) -> VorSection | None:
    """Build the cable-tray metraж section from parsed text callouts.

    ``tray_lengths`` maps a tray size ("50х50") to a list of segment lengths
    in mm. Lengths are summed per size and rounded up to whole metres; an
    overall total row is appended. Returns ``None`` when nothing was found.
    """
    section = VorSection(title="Монтаж кабеленесущих систем (лотки)")
    grand_total_m = 0.0
    for size in sorted(tray_lengths, key=lambda s: -sum(tray_lengths[s])):
        lens = tray_lengths[size]
        if not lens:
            continue
        total_m = sum(lens) / 1000.0
        grand_total_m += total_m
        section.rows.append({
            "name": f"Монтаж лотка {size} мм",
            "unit": "м",
            "qty": round(total_m, 1),
            "is_material": False,
            "drawing_ref": "",
        })
        log(f"    лоток {size}: {total_m:.1f} м ({len(lens)} участков)")
    if not section.rows:
        return None
    log(
        f"  Section (Кабеленесущие системы): {len(section.rows)} "
        f"типоразмеров, итого {grand_total_m:.1f} м"
    )
    return section


def generate_vor_from_plans(source: str | Path, log=print) -> list[VorSection]:
    """Generate VOR sections purely from drawing plans (no СО spec).

    Returns list[VorSection]; emits lighting + devices sections derived from
    on-plan symbol counts, plus a cable-tray metraж section recovered from
    on-plan text length callouts ("<size> <length> мм").
    """
    log("No-СО mode: deriving quantities from plan symbol counts")
    results, desc_to_cat, tray_lengths = count_plans_nospec(source, log)
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

    # Cable-tray metraж recovered from on-plan text callouts (lengths the
    # point-counter cannot see). Independent of symbol counts above.
    trays = _build_tray_metraj_section(tray_lengths, log)
    if trays is not None:
        sections.append(trays)

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
