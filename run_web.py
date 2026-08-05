"""
run_web.py — точка входа для веб-приложения с расширением VOR-экспорта.

Использование:
    python run_web.py
    # или
    uvicorn run_web:app --host 0.0.0.0 --port 8050 --reload

Стартер импортирует web_app.app, регистрирует:
  • быстрый /api/folder/{id}/export_xlsx_v2 (из кэша, формула по отметкам)
  • новый    /api/folder/{id}/export_docx (формат ДБТ-эталона, формула по отметкам)
  • POST     /api/folder/{id}/cache_vor (приём агрегата от клиента)
  • HTML-инжектор <script src="/static/vor_patch.js"> на странице "/"
"""

from __future__ import annotations

import logging

from web_app import app  # noqa: F401 — реэкспорт для uvicorn run_web:app

# 1. Подменяем простой DOCX-рендер на эталонный (ДБТ-формат)
import vor_export_patch as _vep
from vor_docx_renderer import render_vor_docx as _render_etalon
from vor_elevation_grouper import regroup_aggregate_by_elevation

# T077 / S019-compose-adapter: composer + per-PDF sidecar
from vor_compose import compose_vor_table


# ---------------------------------------------------------------------------
# T077 — sidecar holding the upstream per-PDF dict so the patched DOCX
# builder can call vor_compose.compose_vor_table.  Populated by a wrapped
# store_vor_result; consumed by _build_docx_bytes_etalon via rel_folder.
# ---------------------------------------------------------------------------
_PER_PDF_BY_FOLDER: dict[str, dict[str, list[dict]]] = {}

_orig_store_vor_result = _vep.store_vor_result


def _store_with_per_pdf(app_obj, folder_id: str, *, aggregated, all_results, rel_folder):
    """Wrap upstream store_vor_result so we cache all_results by rel_folder
    in addition to the original folder_id-keyed cache.  T076 confirmed the
    cache is passthrough -- all_results is the exact dict[str,list[dict]]
    shape compose_vor_table accepts (keys = PDF file names, values = items
    out of _count_equipment_in_pdf after Steps 8-10 from T080 + T081)."""
    _orig_store_vor_result(
        app_obj, folder_id,
        aggregated=aggregated, all_results=all_results, rel_folder=rel_folder,
    )
    if isinstance(all_results, dict):
        _PER_PDF_BY_FOLDER[rel_folder] = all_results


_vep.store_vor_result = _store_with_per_pdf


def _composed_row_to_renderer_shape(row: dict, idx: int) -> dict:
    """Map vor_compose row {name, unit, qty, rd, formula, sheet_ref,
    notes, _category, _height_bucket, _thickness, _route, _kind,
    _formula_audit} -> renderer shape used by vor_docx_renderer
    (compatible with _render_etalon expectations on aggregated list).

    The renderer pipeline expects keys {row, name, unit, total, formula,
    drawing_refs, extra_info}.  Per T077 (3): map qty->total,
    sheet_ref->drawing_refs, notes->extra_info."""
    return {
        "row": idx,
        "name": row.get("name", ""),
        "unit": row.get("unit", ""),
        "total": row.get("qty", row.get("rd", 0)) or 0,
        "formula": row.get("formula", ""),
        "drawing_refs": row.get("sheet_ref", ""),
        "extra_info": row.get("notes", ""),
        # T077: preserve composer attributes as underscore-prefixed
        # passthrough so downstream renderers / xlsx exporter can read
        # category-grouping hints if they choose.
        "_category": row.get("_category"),
        "_height_bucket": row.get("_height_bucket"),
        "_route": row.get("_route"),
        "_thickness": row.get("_thickness"),
        "_kind": row.get("_kind"),
        "_formula_audit": row.get("_formula_audit"),
    }


def _reconstruct_per_pdf_from_aggregated(aggregated) -> dict[str, list[dict]]:
    """Fallback when sidecar miss: rebuild a per-PDF dict from the already-
    aggregated list using drawing_refs as the source-PDF list.  This loses
    the per-PDF granularity (every PDF gets the full aggregate share) but
    is enough for compose_vor_table to fire without KeyError when the
    cache sidecar is empty (e.g. boss invokes DOCX export from a stored
    JSON snapshot or after process restart).  Best-effort only."""
    if not aggregated:
        return {}
    # Single synthetic bucket: '_aggregate' acts as one virtual PDF.  This
    # keeps compose_vor_table operating on the correct shape; the
    # downstream renderer accepts whatever rows come back.
    return {"_aggregate": list(aggregated)}


def _build_docx_bytes_etalon(aggregated, rel_folder: str) -> bytes:
    """Patched DOCX builder.

    T077 / S019-compose-adapter behaviour:
      1. Pre-existing pre-aggregation by elevation runs first (kept for
         backwards-compat when sidecar misses).
      2. We try to obtain the upstream per-PDF dict via the
         _PER_PDF_BY_FOLDER sidecar populated by the wrapped
         store_vor_result.  If found, we call
         vor_compose.compose_vor_table(per_pdf) and map composed rows
         back to renderer shape.
      3. If the sidecar misses (cold path), we fall back to passing the
         elevation-pre-aggregated list straight through, preserving
         pre-T077 behaviour.
      4. Emit CHECKED-S019-COMPOSE marker log for G1 verification.

    Шапка содержит только строки, в которых мы уверены (по образцу ВОР_ЭО.docx):
      • "Ведомость объемов работ"
      • "Основание_<раздел>"  (раздел определяется по имени папки: ЭО/ЭМ/ЭГ)
    Шифр проекта, номер захватки, имена объекта/стройки/исполнителей не
    извлекаются автоматически и не подставляются.
    """
    log = logging.getLogger("run_web._build_docx_bytes_etalon")
    import sys, time as _t
    def _mark(msg):
        # KB-006 safe: encode->ascii via repr fallback for non-cp1254 chars
        try:
            line = f"[DOCX-TRACE {_t.strftime('%H:%M:%S')}] {msg}"
            sys.stdout.write(line.encode("ascii", "backslashreplace").decode("ascii") + "\n")
            sys.stdout.flush()
        except Exception:
            pass

    _mark(f"enter rel_folder_len={len(rel_folder or '')} aggregated_len={len(aggregated) if aggregated else 0}")

    # 1. Pre-grouping by floor elevation (kept for fallback path)
    aggregated = regroup_aggregate_by_elevation(aggregated, labeled=False)
    _mark(f"after regroup_aggregate_by_elevation -> {len(aggregated)} rows")

    # 2. T077: try to find the per-PDF dict from sidecar.
    per_pdf = _PER_PDF_BY_FOLDER.get(rel_folder) if rel_folder else None
    use_compose = bool(per_pdf)
    final_rows = aggregated
    composed_count = 0
    fallback_reason = ""
    _mark(f"sidecar lookup use_compose={use_compose} per_pdf_keys={len(per_pdf) if per_pdf else 0}")

    if use_compose:
        try:
            _mark("calling compose_vor_table...")
            composed = compose_vor_table(per_pdf)
            composed_count = len(composed)
            _mark(f"compose_vor_table returned {composed_count} rows")
            final_rows = [
                _composed_row_to_renderer_shape(r, i + 1)
                for i, r in enumerate(composed)
            ]
            log.info(
                "CHECKED-S019-COMPOSE rows=%d (per_pdf_keys=%d, folder=%s)",
                composed_count, len(per_pdf), rel_folder,
            )
        except Exception as exc:  # noqa: BLE001
            fallback_reason = f"compose_vor_table raised {type(exc).__name__}: {exc}"
            _mark(f"compose FAILED: {fallback_reason}")
            log.warning(
                "CHECKED-S019-COMPOSE FALLBACK (%s) rows=%d folder=%s",
                fallback_reason, len(aggregated), rel_folder,
            )
            final_rows = aggregated
    else:
        fallback_reason = "sidecar miss"
        log.info(
            "CHECKED-S019-COMPOSE FALLBACK (%s) rows=%d folder=%s",
            fallback_reason, len(aggregated), rel_folder,
        )

    # 3. Section title from folder (unchanged from pre-T077 logic)
    folder_lower = (rel_folder or "").lower()
    if "эо" in folder_lower or "освещ" in folder_lower:
        razdel_full = "Электроосвещение"
    elif "эм" in folder_lower:
        razdel_full = "Электрооборудование"
    elif "эг" in folder_lower:
        razdel_full = "Заземление"
    else:
        razdel_full = ""

    section_basis = f"Основание_{razdel_full}" if razdel_full else ""
    _mark(f"calling _render_etalon with {len(final_rows)} rows, section_basis_len={len(section_basis)}")

    out = _render_etalon(
        final_rows,
        rel_folder=rel_folder,
        section_basis=section_basis,
    )
    _mark(f"_render_etalon done, bytes={len(out)}")
    return out


_vep._build_docx_bytes = _build_docx_bytes_etalon  # подмена


# 2. Также пере-группируем XLSX (через monkey-patch агрегата на этапе экспорта).
#    Делаем это путём wrapper'а над get_cached_vor.
_orig_get_cached = _vep.get_cached_vor


def _get_cached_with_elevation(app_obj, folder_id: str):
    cached = _orig_get_cached(app_obj, folder_id)
    if not cached:
        return None
    # Пере-группируем агрегат, если ещё не пере-группирован
    if cached.get("_elevation_done"):
        return cached
    new_agg = regroup_aggregate_by_elevation(cached["aggregated"], labeled=False)
    new_cached = dict(cached)
    new_cached["aggregated"] = new_agg
    new_cached["_elevation_done"] = True
    return new_cached


_vep.get_cached_vor = _get_cached_with_elevation


# 3. middleware-инжектор скрипта
from vor_inject_middleware import install_html_inject
install_html_inject(app)

# 4. регистрация эндпоинтов экспорта
from vor_export_patch import register_vor_export
register_vor_export(app)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("run_web:app", host="0.0.0.0", port=8050, reload=False)
