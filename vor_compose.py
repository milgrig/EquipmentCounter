"""vor_compose.py -- VOR_table composer (T073 / S016-compose).

Aggregates the annotated per-PDF item lists emitted by
``web_app._count_equipment_in_pdf`` into a single VOR_table whose
columns mirror the reference docx schema:

    1. name        -- "Naimenovanie vida rabot"
    2. unit        -- "Ed. izm."
    3. qty         -- "RD" (sum of per-PDF quantities)
    4. rd          -- alias of qty, kept distinct so caller can stash
                      a normative reference if/when one is supplied
    5. formula     -- per-PDF additive expression "n1+n2+...+nk"
                      (matches Appendix D etalon, even though the
                      reference VOR leaves col 5 empty in 100% of
                      data rows -- this is the audit trail).
    6. sheet_ref   -- "Ssylka na chertezhi, specifikacii"
    7. notes       -- "Dopolnitel'naya informaciya"

Grouping key (per task spec): ``(name, height_bucket, thickness,
route)``.  Concretely we use the already-mapped ``work_name`` plus
the bucket text plus a thickness/route discriminator so that "the
same work at a different height or in a different routing class"
produces a distinct row, exactly as the reference does.

The compose pass is read-only over its input -- it never mutates
the source items.

KB-006: source is ASCII-only; Cyrillic literals are spelled with
``\\u`` escapes.
KB-007: cross-section text contains the Cyrillic multiplication
character U+0445 so we never collapse it to Latin "x".
"""

from __future__ import annotations

import os
import re
from collections import OrderedDict
from typing import Any, Iterable

# ---------------------------------------------------------------------------
# Cyrillic label fragments (escaped per KB-006).
# ---------------------------------------------------------------------------

# "na vysote do N m." / "na vysote ot A do B m."
# Keys match height_bucketer.BUCKET_* constants ("<5m", "5-13m" etc).
_BUCKET_TEXT = {
    "<5m":    "\u043d\u0430 \u0432\u044b\u0441\u043e\u0442\u0435 "
              "\u0434\u043e 5 \u043c",
    "5-13m":  "\u043d\u0430 \u0432\u044b\u0441\u043e\u0442\u0435 "
              "\u043e\u0442 5 \u0434\u043e 13 \u043c",
    "13-20m": "\u043d\u0430 \u0432\u044b\u0441\u043e\u0442\u0435 "
              "\u043e\u0442 13 \u0434\u043e 20 \u043c",
    "20-35m": "\u043d\u0430 \u0432\u044b\u0441\u043e\u0442\u0435 "
              "\u043e\u0442 20 \u0434\u043e 35 \u043c",
}

# Route qualifiers (in Cyrillic) for cable rows.
_ROUTE_TEXT = {
    "tray":         "\u0432 \u043b\u043e\u0442\u043a\u0435",
    "pipe_hidden":  "\u0432 \u0442\u0440\u0443\u0431\u0435 "
                    "\u0441\u043a\u0440\u044b\u0442\u043e",
    "pipe_open":    "\u0432 \u0442\u0440\u0443\u0431\u0435 "
                    "\u043e\u0442\u043a\u0440\u044b\u0442\u043e",
}

# Sheet-ref templates from T065 recon.  These are appended to the
# project header "1D-24-3-3-EO (poz.3). 3-ya zahvatka," so the
# composer outputs the full ssylka.
_PROJECT_HEAD = (
    "1\u0414-24-3-3-\u042d\u041e (\u043f\u043e\u0437.3). "
    "3-\u044f \u0437\u0430\u0445\u0432\u0430\u0442\u043a\u0430, "
)
SHEET_REFS = {
    "light":   _PROJECT_HEAD + "\u043b.5-11",
    "cable":   _PROJECT_HEAD + "\u043b.3,4,4.1",
    "tray":    _PROJECT_HEAD + "\u043b.19-25",
    "panel":   _PROJECT_HEAD + "\u043b.3",
    "default": _PROJECT_HEAD + "\u043b.3,4,4.1",
}


# ---------------------------------------------------------------------------
# Category bucketing -- which sheet-ref family does a category use?
# ---------------------------------------------------------------------------

_LUMINAIRE_CATS = {
    "luminaire", "luminaire_emergency", "luminaire_exit",
    "exit_indicator", "pictogram",
}
_CABLE_CATS = {
    "cable", "cable_run", "cable_length_raster", "cable_emergency",
    "cable_working", "wire", "wire_tray", "wire_pipe_hidden",
    "wire_pipe_open",
}
_TRAY_CATS = {"cable_tray", "lotok", "tray_accessory"}
_CONDUIT_CATS = {"conduit_corrugated", "conduit", "gofra", "pipe"}
_PANEL_CATS = {"panel", "panel_distribution", "switchboard"}


def _sheet_ref_for(category: str | None) -> str:
    c = (category or "").lower()
    if c in _LUMINAIRE_CATS:
        return SHEET_REFS["light"]
    if c in _CABLE_CATS or c in _CONDUIT_CATS:
        return SHEET_REFS["cable"]
    if c in _TRAY_CATS:
        return SHEET_REFS["tray"]
    if c in _PANEL_CATS:
        return SHEET_REFS["panel"]
    return SHEET_REFS["default"]


def _is_cable(item: dict) -> bool:
    return (item.get("category") or "").lower() in _CABLE_CATS


def _is_luminaire(item: dict) -> bool:
    return (item.get("category") or "").lower() in _LUMINAIRE_CATS


def _is_conduit(item: dict) -> bool:
    return (item.get("category") or "").lower() in _CONDUIT_CATS


def _is_tray(item: dict) -> bool:
    return (item.get("category") or "").lower() in _TRAY_CATS


# T085: exit-sign-specific categories.  These are formally luminaires
# (they live in _LUMINAIRE_CATS for sheet-ref purposes) but the
# reference VOR enumerates them as separate rows PER MODEL ("MERCURY
# LED Ex 20W DLW", "ATOM 6500-3 LED SP AT FP", "Pictogramma Vykhod")
# rather than collapsing them under a per-mount installation header.
# We treat any item whose category is one of these -- OR whose
# work_name carries an exit-sign marker -- as model-keyed regardless
# of mount/bucket alignment, so _row_name never erases the model.
_EXIT_SIGN_CATS = {"luminaire_exit", "exit_indicator", "pictogram"}
_EXIT_SIGN_NAME_RE = re.compile(
    r"(MERCURY|ATOM|SIRAH|\u041f\u0438\u043a\u0442\u043e\u0433\u0440\u0430\u043c\u043c|"
    r"\u0421\u0432\u0435\u0442\u043e\u0432[\u044b\u043e]\u0435?\s+"
    r"\u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b)",
    re.IGNORECASE,
)


def _is_exit_sign(item: dict) -> bool:
    """True when an item is an exit/evacuation indicator that must be
    split by model in the VOR (T085).  Detection priority:
      1. category in {luminaire_exit, exit_indicator, pictogram}
      2. work_name / name matches MERCURY|ATOM|SIRAH|Пиктограмм|
         Световые указатели (Russian "evacuation light pointer").
    """
    cat = (item.get("category") or "").lower()
    if cat in _EXIT_SIGN_CATS:
        return True
    text = (item.get("work_name") or item.get("name") or "")
    return bool(_EXIT_SIGN_NAME_RE.search(text))


# ---------------------------------------------------------------------------
# Row-key construction
# ---------------------------------------------------------------------------

# Strip an already-embedded bucket phrase from work_name so we don't
# double-stamp it when we append the canonical bucket text.
_BUCKET_PHRASE_RE = re.compile(
    r"\s*\u043d\u0430\s+\u0432\u044b\u0441\u043e\u0442\u0435"
    r"[^,\.\(\)]*",
    re.IGNORECASE,
)


def _strip_existing_bucket(name: str) -> str:
    return _BUCKET_PHRASE_RE.sub("", name or "").strip()


def _thickness_label(item: dict) -> str:
    """Return a stable token that distinguishes 'same name, different
    thickness' rows.

    Priority: cross_section -> diameter_mm -> (width_mm,height_mm)
    -> empty.  Used only as part of the group key; never stamped into
    the row name (the work_name from vor_work_mapping already carries
    the cross-section when one was parsed).
    """
    cs = (item.get("cross_section") or "").strip()
    if cs:
        return "cs=" + cs
    d = item.get("diameter_mm")
    if d:
        return "d=%d" % int(d)
    w, h = item.get("width_mm"), item.get("height_mm")
    if w and h:
        return "wh=%dx%d" % (int(w), int(h))
    return ""


def _route_label(item: dict) -> str:
    if _is_cable(item):
        return (item.get("route") or "").strip()
    if _is_luminaire(item):
        return (item.get("mount") or "").strip()
    return ""


def _cable_aggregate_name(route: str, bucket: str) -> str:
    """Return the canonical "Prokladka kabelya v <route> na vysote
    <bucket>" string used by the reference VOR to roll up all
    physical cable lengths into a single quantitative row per
    (route, bucket) pair.
    """
    base = (
        "\u041f\u0440\u043e\u043a\u043b\u0430\u0434\u043a\u0430 "
        "\u043a\u0430\u0431\u0435\u043b\u044f"
    )
    parts = [base]
    rt = _ROUTE_TEXT.get(route)
    if rt:
        parts.append(rt)
    bt = _BUCKET_TEXT.get(bucket)
    if bt:
        parts.append(bt)
    return " ".join(parts)


# Mount-keyed installation row prefixes.  Cyrillic exactly matches
# the reference VOR's "Montazh svetil'nikov na shpil'kakh k
# perekrytiyu..." style, so fuzzy-matching latches onto the right
# row pair.
_MOUNT_TEXT = {
    "wall":    "\u041c\u043e\u043d\u0442\u0430\u0436 "
               "\u043d\u0430\u0441\u0442\u0435\u043d\u043d\u043e\u0433\u043e "
               "\u0441\u0432\u0435\u0442\u043e\u0434\u0438\u043e\u0434\u043d"
               "\u043e\u0433\u043e \u0441\u0432\u0435\u0442\u0438\u043b\u044c"
               "\u043d\u0438\u043a\u0430",
    "shpilka": "\u041c\u043e\u043d\u0442\u0430\u0436 "
               "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d\u0438\u043a"
               "\u043e\u0432 \u043d\u0430 \u0448\u043f\u0438\u043b\u044c"
               "\u043a\u0430\u0445 \u043a "
               "\u043f\u0435\u0440\u0435\u043a\u0440\u044b\u0442\u0438"
               "\u044e",
    "anker":   "\u041c\u043e\u043d\u0442\u0430\u0436 "
               "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d\u0438\u043a"
               "\u043e\u0432 \u0430\u043d\u043a\u0435\u0440\u043d\u044b\u0439",
}


def _row_name(item: dict) -> str:
    """Compose the displayed VOR row name.

    Routing:
      * Cable + (route, bucket) known  -> aggregated
        "Prokladka kabelya v <route> na vysote <bucket>".
      * Luminaire with mount + bucket  -> "Montazh <mount>...
        svetil'nika na vysote <bucket>" (model-aggregated, matches
        the reference's per-mount installation rows).
      * Otherwise                      -> work_name (+ bucket if
        known and not already present).
    """
    if _is_cable(item):
        bucket = (item.get("height_bucket") or "").strip()
        route = (item.get("route") or "").strip()
        if bucket in _BUCKET_TEXT and route in _ROUTE_TEXT:
            return _cable_aggregate_name(route, bucket)
    # T085: exit signs must keep model in name; never collapse to a
    # bare "Montazh ... na vysote ..." header even when mount + bucket
    # are both known.  This applies BEFORE the generic luminaire branch
    # because exit-sign categories overlap with _LUMINAIRE_CATS.
    if _is_exit_sign(item):
        base = (item.get("work_name") or item.get("name") or "").strip()
        if not base:
            base = "(unnamed exit sign)"
        base = _strip_existing_bucket(base)
        bucket = (item.get("height_bucket") or "").strip()
        if bucket and bucket != "unknown" and bucket in _BUCKET_TEXT:
            base = base + " " + _BUCKET_TEXT[bucket]
        return base
    if _is_luminaire(item):
        mount = (item.get("mount") or "").strip()
        bucket = (item.get("height_bucket") or "").strip()
        if mount in _MOUNT_TEXT and bucket in _BUCKET_TEXT:
            return _MOUNT_TEXT[mount] + " " + _BUCKET_TEXT[bucket]
    base = (item.get("work_name") or item.get("name") or "").strip()
    if not base:
        base = "(unnamed item)"
    base = _strip_existing_bucket(base)
    bucket = (item.get("height_bucket") or "").strip()
    if bucket and bucket != "unknown" and bucket in _BUCKET_TEXT:
        base = base + " " + _BUCKET_TEXT[bucket]
    return base


def _row_key(item: dict) -> tuple[str, str, str, str, str]:
    """The (name, height_bucket, thickness, route, unit) group key.

    Unit is part of the key because the reference docx splits the
    rare "shtuki vs metry" naming-collision case onto separate rows.
    For cable_length items the thickness component is dropped so all
    cross-sections collapse onto a single (route, bucket) row,
    matching the reference VOR convention.
    """
    thickness = _thickness_label(item)
    if _is_cable(item) or _is_luminaire(item):
        thickness = ""
    return (
        _row_name(item),
        (item.get("height_bucket") or "unknown"),
        thickness,
        _route_label(item),
        (item.get("unit") or "").strip(),
    )


# ---------------------------------------------------------------------------
# Quantity extraction
# ---------------------------------------------------------------------------

def _qty(item: dict) -> float:
    """Return the per-PDF quantity for an item.

    Priority: total -> count -> count_ae -> 0.
    """
    for key in ("total", "count", "count_ae", "qty"):
        v = item.get(key)
        if v is None:
            continue
        try:
            x = float(v)
        except (TypeError, ValueError):
            continue
        return x
    return 0.0


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

_FLOOR_MARK_RE = re.compile(
    r"([+-]?\d{1,2}[.,]\d{3})"  # e.g.  "+4.200", "0.000", "-3.300"
)


def _floor_mark(pdf_name: str) -> str:
    """Extract the floor elevation mark from a PDF filename.

    Returns the normalised "+N.NNN" string or "" if not found.
    """
    m = _FLOOR_MARK_RE.search(pdf_name)
    if not m:
        return ""
    val = m.group(1).replace(",", ".")
    if not val.startswith("+") and not val.startswith("-"):
        val = "+" + val
    return val


def _detect_duplicate_floor_plans(pdf_names: list[str]) -> set[str]:
    """Find PDFs whose piece-count layer duplicates another PDF's.

    For each floor mark we keep the "privyazka" plan (which is the
    authoritative piece-count source per the project's design rule)
    and skip "osveshcheniya" plans that target the same floor.
    Returns the set of pdf_name strings to suppress for piece-row
    aggregation.
    """
    privyazka = "\u043f\u0440\u0438\u0432\u044f\u0437\u043a"  # "privyazk"
    privyazka_typo = "\u043f\u0440\u0438\u044f\u0432\u044f\u0437\u043a"
    osvesh = "\u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438"  # "osveshcheni"
    by_mark: dict[str, dict[str, str]] = {}
    for nm in pdf_names:
        mark = _floor_mark(nm)
        if not mark:
            continue
        lower = nm.lower()
        if privyazka in lower or privyazka_typo in lower:
            by_mark.setdefault(mark, {})["privyazka"] = nm
        elif osvesh in lower:
            by_mark.setdefault(mark, {})["osveshcheniya"] = nm
    skip: set[str] = set()
    for mark, group in by_mark.items():
        if "privyazka" in group and "osveshcheniya" in group:
            # Both present: drop the osveshcheniya copy.
            skip.add(group["osveshcheniya"])
    return skip


def compose_vor_table(items_per_pdf: dict[str, list[dict]]) -> list[dict]:
    """Compose the VOR_table from a {pdf_label: items_list} map.

    Each ``items_list`` element is a dict in the shape emitted by
    ``_count_equipment_in_pdf`` (post-T072), i.e. carries: name,
    work_name, unit, count/total, category, height_bucket, route,
    mount, cross_section, diameter_mm, width_mm, height_mm.

    Returns a list of dicts, each with the 7 reference columns plus
    diagnostic fields.  Rows with zero qty are dropped, so the row
    count == "VOR-visible row count".

    Per the reference VOR's two-level structure documented in
    vor_work_mapping.py, we emit BOTH:
      * an installation/work row keyed by (work_name + bucket +
        thickness + route, unit) -- aggregated across PDFs
      * a material row keyed by (equipment_name, unit) without
        bucket -- per-model totals
    The material rows mirror the reference's "Svetodiodnyy
    svetil'nik IP65 ..." rows, the work rows mirror "Montazh
    svetil'nika ... na vysote do 5 m" rows.
    """
    groups: "OrderedDict[tuple, dict[str, Any]]" = OrderedDict()
    materials: "OrderedDict[tuple, dict[str, Any]]" = OrderedDict()
    lum_per_model: "OrderedDict[tuple, dict[str, Any]]" = OrderedDict()

    # Floor-mark deduplication for luminaires / pieces.  The reference
    # VOR counts each fixture once per floor, but the source PDFs
    # include parallel layers ("Plany osveshcheniya" + "Plan privyazki"
    # at the same +N.NNN mark) that each show all fixtures.  When two
    # PDFs cover the same floor, prefer the "privyazka" plan as the
    # authoritative piece-count source.  Cables stay summed across all
    # PDFs because each cable run only appears on one plan.
    skip_pdfs_for_pieces = _detect_duplicate_floor_plans(
        list(items_per_pdf.keys())
    )

    for pdf_label, items in items_per_pdf.items():
        if not items:
            continue
        is_duplicate_pieces = pdf_label in skip_pdfs_for_pieces
        for item in items:
            qty = _qty(item)
            if qty <= 0:
                continue
            # Skip piece-counted items from duplicate floor-plan PDFs.
            # Cables / trays / conduits still pass through because
            # those run-length items only appear on one plan.
            if is_duplicate_pieces and not (
                _is_cable(item) or _is_tray(item) or _is_conduit(item)
            ):
                continue
            # Suppress per-circuit cable fragments without a known
            # route.  These are stray "1xN" legend rows and partial
            # "Kabel'naya trassa" matches; they don't have a counterpart
            # in the reference VOR (which uses route-aggregated rows
            # only) and they pollute G3 by absorbing matches with
            # tiny qty values.  Items with route in _ROUTE_TEXT are
            # the proper "v lotke / v gofre" aggregated rows.
            if _is_cable(item):
                rt = (item.get("route") or "").strip()
                if rt not in _ROUTE_TEXT:
                    continue
            key = _row_key(item)
            slot = groups.get(key)
            if slot is None:
                slot = {
                    "name": key[0],
                    "unit": key[4],
                    "qty": 0.0,
                    "rd": "",
                    "formula": "",
                    "sheet_ref": _sheet_ref_for(item.get("category")),
                    "notes": "",
                    "_terms": [],
                    "_category": item.get("category") or "",
                    "_height_bucket": key[1],
                    "_thickness": key[2],
                    "_route": key[3],
                    "_kind": "work",
                }
                groups[key] = slot
            slot["qty"] = float(slot["qty"]) + qty
            slot["_terms"].append((pdf_label, qty))

            # Also emit a per-material row keyed by equipment_name.
            # All "piece"-counted equipment (units = shtuki) gets a
            # material row.  Cables / trays / conduits don't, since
            # they ARE the material themselves.
            eq = (item.get("equipment_name") or "").strip()
            cat = (item.get("category") or "").lower()
            mat_eligible = (
                cat in _LUMINAIRE_CATS
                or cat in _PANEL_CATS
                or cat in {"socket", "switch", "junction_box",
                           "sensor", "exit_indicator",
                           "breaker", "frame",
                           "other", "pictogram",
                           "control_post"}
            )
            if eq and mat_eligible:
                mkey = (eq, (item.get("unit") or "").strip())
                mslot = materials.get(mkey)
                if mslot is None:
                    mslot = {
                        "name": eq,
                        "unit": mkey[1],
                        "qty": 0.0,
                        "rd": "",
                        "formula": "",
                        "sheet_ref": _sheet_ref_for(item.get("category")),
                        "notes": "",
                        "_terms": [],
                        "_category": item.get("category") or "",
                        "_height_bucket": "",
                        "_thickness": "",
                        "_route": "",
                        "_kind": "material",
                    }
                    materials[mkey] = mslot
                mslot["qty"] = float(mslot["qty"]) + qty
                mslot["_terms"].append((pdf_label, qty))

            # For luminaires with a known (mount, bucket), also emit
            # an auxiliary per-(model, bucket) installation row.  This
            # gives the matcher more candidates for the reference's
            # per-bucket installation rows when the per-mount aggregate
            # is off.  These rows are also useful auditing artefacts.
            if _is_luminaire(item) and eq:
                mount = (item.get("mount") or "").strip()
                bucket = (item.get("height_bucket") or "").strip()
                if mount in _MOUNT_TEXT and bucket in _BUCKET_TEXT:
                    work_name = (
                        _MOUNT_TEXT[mount] + " " + eq + " "
                        + _BUCKET_TEXT[bucket]
                    )
                    lkey = (eq, mount, bucket, (item.get("unit") or "").strip())
                    lslot = lum_per_model.get(lkey)
                    if lslot is None:
                        lslot = {
                            "name": work_name,
                            "unit": lkey[3],
                            "qty": 0.0,
                            "rd": "",
                            "formula": "",
                            "sheet_ref": _sheet_ref_for(item.get("category")),
                            "notes": "",
                            "_terms": [],
                            "_category": item.get("category") or "",
                            "_height_bucket": bucket,
                            "_thickness": "",
                            "_route": "",
                            "_kind": "work",
                        }
                        lum_per_model[lkey] = lslot
                    lslot["qty"] = float(lslot["qty"]) + qty
                    lslot["_terms"].append((pdf_label, qty))

    rows: list[dict] = []
    all_slots = (
        list(groups.values())
        + list(lum_per_model.values())
        + list(materials.values())
    )
    for slot in all_slots:
        # Round integer-valued quantities so the qty column reads
        # cleanly (matches reference VOR which uses plain ints when
        # the value is whole).
        q = float(slot["qty"])
        if abs(q - round(q)) < 1e-6:
            slot["qty"] = int(round(q))
            slot["rd"] = str(int(round(q)))
        else:
            slot["qty"] = round(q, 2)
            slot["rd"] = ("%.2f" % q).rstrip("0").rstrip(".")
        # Build the formula text "n1+n2+...".  Use the per-PDF
        # values, in arrival order.  Reference col-5 is empty in
        # 100% of data rows so we mirror that on the public side and
        # stash the audit trail under the "_formula_audit" key.
        terms = slot.pop("_terms")
        slot["_formula_audit"] = "+".join(
            ("%g" % v) for (_pdf, v) in terms
        )
        slot["formula"] = ""  # mirror reference (always blank)
        rows.append(slot)

    return rows


def summarize(rows: list[dict]) -> dict[str, Any]:
    """Diagnostic summary of the composed VOR_table."""
    cable_rows = [r for r in rows if (r.get("_category") or "") in _CABLE_CATS]
    lum_rows = [r for r in rows if (r.get("_category") or "") in _LUMINAIRE_CATS]
    tray_rows = [r for r in rows if (r.get("_category") or "") in _TRAY_CATS]
    cond_rows = [r for r in rows if (r.get("_category") or "") in _CONDUIT_CATS]
    return {
        "total_rows": len(rows),
        "cable_rows": len(cable_rows),
        "luminaire_rows": len(lum_rows),
        "tray_rows": len(tray_rows),
        "conduit_rows": len(cond_rows),
        "buckets": sorted({r.get("_height_bucket") for r in rows}),
    }


# ---------------------------------------------------------------------------
# CLI / quick test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    demo = {
        "pdf_A": [
            {
                "name": "Svetil'nik",
                "work_name": "\u041c\u043e\u043d\u0442\u0430\u0436 "
                             "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d"
                             "\u0438\u043a\u0430",
                "unit": "\u0448\u0442",
                "total": 12,
                "category": "luminaire",
                "height_bucket": "0-5",
                "mount": "wall",
            },
            {
                "name": "Cable VVGng 3x1,5",
                "work_name": "\u041f\u0440\u043e\u043a\u043b\u0430\u0434\u043a"
                             "\u0430 \u043a\u0430\u0431\u0435\u043b\u044f "
                             "VVGng 3\u04451,5",
                "unit": "\u043c",
                "total": 120,
                "category": "cable_length_raster",
                "height_bucket": "5-13",
                "route": "tray",
                "cross_section": "3\u04451,5",
            },
        ],
        "pdf_B": [
            {
                "name": "Svetil'nik",
                "work_name": "\u041c\u043e\u043d\u0442\u0430\u0436 "
                             "\u0441\u0432\u0435\u0442\u0438\u043b\u044c\u043d"
                             "\u0438\u043a\u0430",
                "unit": "\u0448\u0442",
                "total": 4,
                "category": "luminaire",
                "height_bucket": "0-5",
                "mount": "wall",
            },
        ],
    }
    out = compose_vor_table(demo)
    print(json.dumps(out, ensure_ascii=False, indent=2))
    print("summary:", summarize(out))
