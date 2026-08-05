r"""height_bucketer.py -- T069 / S016 height-bucket attribution.

Per T060 / T065 recon: the reference VOR groups installation work into
four height buckets

    "<5m"      ("do 5 metrov")           floor 0..5
    "5-13m"    ("ot 5 do 13 metrov")     floor 5..13
    "13-20m"   ("ot 13 do 20 metrov")    floor 13..20
    "20-35m"   ("ot 20 do 35 metrov")    floor 20..35

The bucket is derived from the *otmetka* (floor elevation) printed in
the title block of each PDF.  Today's GPK3 layouts encode the otmetka
in the FILENAME of every plan sheet, e.g.

    "005-Plany osvescheniya-otm. 0.000.pdf"
    "007-Plany osvescheniya-otm. +7.800 +9.000.pdf"
    "022-Lotki na otm. +13.800.pdf"

When a sheet carries two otmetki (e.g. +7.800 +9.000) we attribute to
the HIGHER one because that is the ceiling height of the lower zone --
matching the engineer's hand-bucket rule in T065_recon.md.

KB-006: source is ASCII-only.  Cyrillic strings only appear via
``\uNNNN`` escapes.  The labels emitted into items use the Cyrillic VOR
wording so downstream VOR mapping can match etalon row names verbatim.

KB-004: this module is project source (within scope).

Public API
----------

``bucket_for_path(pdf_path) -> str``         -- canonical bucket key
``bucket_for_elevations(elevs) -> str``       -- explicit list of floats
``bucket_label_ru(bucket_key) -> str``        -- Cyrillic VOR label
``attribute_items(items, pdf_path) -> None``  -- in-place tag each item
                                                 with ``height_bucket``
``group_items_by_bucket(items) -> dict``      -- group by (name, unit,
                                                 height_bucket)
"""

from __future__ import annotations

import os
import re
from typing import Iterable, Optional

try:
    from height_mapping import (
        extract_elevations_from_filename,
        height_to_category,
        DEFAULT_FLOOR_HEIGHT_M,
    )
except Exception:  # pragma: no cover -- defensive
    extract_elevations_from_filename = None  # type: ignore
    height_to_category = None  # type: ignore
    DEFAULT_FLOOR_HEIGHT_M = 4.8

# ---------------------------------------------------------------------------
# Canonical bucket keys (ASCII; KB-006)
# ---------------------------------------------------------------------------

BUCKET_LT5 = "<5m"
BUCKET_5_13 = "5-13m"
BUCKET_13_20 = "13-20m"
BUCKET_20_35 = "20-35m"
BUCKET_UNKNOWN = "unknown"

ALL_BUCKETS = (BUCKET_LT5, BUCKET_5_13, BUCKET_13_20, BUCKET_20_35)

# Cyrillic labels matching the reference VOR substring rules
# (see T065_recon.md Q1).  Used by vor_composer to write row names.
_BUCKET_LABEL_RU = {
    BUCKET_LT5: "\u0434\u043e 5 \u043c\u0435\u0442\u0440\u043e\u0432",
    BUCKET_5_13: (
        "\u043e\u0442 5 \u0434\u043e 13 \u043c\u0435\u0442\u0440\u043e\u0432"
    ),
    BUCKET_13_20: (
        "\u043e\u0442 13 \u0434\u043e 20 \u043c\u0435\u0442\u0440\u043e\u0432"
    ),
    BUCKET_20_35: (
        "\u043e\u0442 20 \u0434\u043e 35 \u043c\u0435\u0442\u0440\u043e\u0432"
    ),
    BUCKET_UNKNOWN: "",
}

# Fallback regex: matches "+12.345" / "12.345" / "0.000" / "+7,800" /
# "+7-800" anywhere in the filename.  Used when
# ``extract_elevations_from_filename`` returns empty because the file
# does not carry the "otm" / "\u043e\u0442\u043c" marker but does carry
# a numeric elevation in a different shape.
_RAW_ELEV_RE = re.compile(
    r"(?<![\d,-])([+-]?\d{1,3}[.,-]\d{1,3})(?!\d)"
)

# Bucket boundaries (closed-low, open-high) matching T065 recon.
# Bucket selection is by the SHEET's reported ceiling height (or the
# higher of the two otmetki when both ceiling and floor are encoded
# into the same filename).
_BOUNDS = (
    (5.0, BUCKET_LT5),
    (13.0, BUCKET_5_13),
    (20.0, BUCKET_13_20),
    (float("inf"), BUCKET_20_35),
)


def _height_m_to_bucket(height_m: float) -> str:
    for boundary, key in _BOUNDS:
        if height_m < boundary:
            return key
    return BUCKET_20_35


def _parse_raw_token(tok: str) -> Optional[float]:
    """Fallback parser for the bare numeric form ``+7.800``.

    Mirrors ``height_mapping.parse_elevation_token`` but is self-contained
    so this module degrades gracefully when ``height_mapping`` is
    unavailable.
    """
    s = tok.strip()
    if s.startswith("+"):
        s = s[1:]
    s = s.replace(",", ".").replace("-", ".")
    if s.startswith(".."):
        s = "-" + s[2:]
    try:
        v = float(s)
    except ValueError:
        return None
    if abs(v) >= 100.0:
        v = v / 1000.0
    return v


def extract_elevations(path_or_name: str) -> list[float]:
    """Return every numeric otmetka encoded in the supplied path.

    Tries the project-wide ``extract_elevations_from_filename`` first
    (which requires an "\u043e\u0442\u043c" marker) and falls back to a
    raw-token sweep for filenames that omit the marker.
    """
    name = os.path.basename(path_or_name)
    elevs: list[float] = []
    if extract_elevations_from_filename is not None:
        try:
            elevs = extract_elevations_from_filename(name)
        except Exception:
            elevs = []
    if elevs:
        return elevs

    # Fallback: require an explicit otmetka marker.  We accept the
    # Cyrillic "\u043e\u0442\u043c" (already tried above by
    # extract_elevations_from_filename), the Latin transliteration
    # "otm", or the full word "elev" -- anything else is treated as
    # "no otmetka encoded in this filename" and yields an empty list.
    # This avoids matching sheet numbers like "004.1" or "1-2" that
    # look numerically like elevations but are not.
    marker = re.search(r"otm|elev", name, re.IGNORECASE)
    if marker is None:
        return []
    scan = name[marker.end():]
    seen: set[float] = set()
    out: list[float] = []
    for m in _RAW_ELEV_RE.finditer(scan):
        v = _parse_raw_token(m.group(1))
        if v is None:
            continue
        if not (-50.0 <= v <= 200.0):
            continue
        # Drop "0.0" emitted by sheet numbers like "1.5" -- but DO keep
        # the literal "0.000" floor marker which is a valid otmetka.
        if abs(v) < 1e-6 and "0.000" not in name and "0-000" not in name:
            continue
        if v in seen:
            continue
        seen.add(v)
        out.append(v)
    return out


def bucket_for_elevations(elevs: Iterable[float]) -> str:
    """Pick the canonical bucket for an iterable of floor otmetki.

    Rule: take the MAX otmetka in the list (the ceiling height of the
    enclosed zone, per T065 recon), then map it through the bucket
    boundaries.  Empty list returns ``BUCKET_UNKNOWN``.
    """
    elevs = [float(e) for e in elevs]
    if not elevs:
        return BUCKET_UNKNOWN
    # Use the MAX otmetka.  This matches the T065 recon rule and is
    # robust to "+7.800 +9.000" double-otmetka sheets where the upper
    # value is the right pick.  A literal 0.000 sheet correctly stays
    # in the <5m bucket because 0 < 5.
    ref = max(elevs)
    # The bucket is on the CEILING height; for the top elevation we
    # add the default floor height to estimate the ceiling.  For lower
    # elevations the next otmetka in the filename (if present) is
    # already the ceiling, so we simply use ``ref`` directly when it
    # already exceeds 0.
    if ref >= 5.0:
        return _height_m_to_bucket(ref)
    # ref < 5.0 (typical "0.000" / "+4.200" sheets) -- the ceiling is
    # approximately ref + DEFAULT_FLOOR_HEIGHT_M, which lands in <5m
    # for ref=0 and 5-13m for ref=+4.200.  Match the engineer's hand
    # rule: ref=0 -> <5m; ref=+4.200 -> still <5m if the next floor is
    # +9.000 (ceiling 9.0 -> 5-13m bucket).  Without that knowledge we
    # default to <5m for any ref<5.
    return BUCKET_LT5


def bucket_for_path(pdf_path: str) -> str:
    """Return the canonical bucket key for a PDF file path."""
    return bucket_for_elevations(extract_elevations(pdf_path))


def bucket_label_ru(bucket_key: str) -> str:
    """Return the Cyrillic VOR label for a canonical bucket key."""
    return _BUCKET_LABEL_RU.get(bucket_key, "")


# ---------------------------------------------------------------------------
# Item attribution & grouping
# ---------------------------------------------------------------------------


def attribute_items(items: list[dict], pdf_path: str) -> str:
    """Tag every item in-place with ``height_bucket`` derived from the
    PDF path.  Returns the bucket key applied (so the caller can log).

    G1 acceptance criterion: every item produced by
    ``_count_equipment_in_pdf`` must have a non-empty ``height_bucket``
    field after this call.  We choose the value:

    - the resolved canonical bucket key, OR
    - ``BUCKET_UNKNOWN`` when the filename carries no parseable otmetka

    Items already carrying a ``height_bucket`` are left untouched so
    upstream stages (e.g. per-symbol height-from-template) can override.
    """
    bucket = bucket_for_path(pdf_path)
    for it in items:
        if not it.get("height_bucket"):
            it["height_bucket"] = bucket
            it["height_bucket_label"] = bucket_label_ru(bucket)
    return bucket


def group_items_by_bucket(items: list[dict]) -> dict[tuple, dict]:
    """Group items by ``(name, unit, height_bucket)`` summing ``count``
    and ``total`` per group.  The return value mirrors the input dict
    schema but collapses duplicates produced when the same kind of work
    appears on multiple pages of the same bucket.

    Output key: ``(name, unit, height_bucket)``.
    Output value: aggregated item dict.
    """
    groups: dict[tuple, dict] = {}
    for it in items:
        name = it.get("name", "")
        unit = it.get("unit", "")
        bucket = it.get("height_bucket") or BUCKET_UNKNOWN
        key = (name, unit, bucket)
        if key not in groups:
            groups[key] = {
                "name": name,
                "unit": unit,
                "height_bucket": bucket,
                "height_bucket_label": bucket_label_ru(bucket),
                "count": 0,
                "count_ae": 0,
                "total": 0.0,
                "symbol": it.get("symbol", ""),
                "category": it.get("category", ""),
                "source": it.get("source", ""),
                "pdf_sources": [],
            }
        g = groups[key]
        try:
            g["count"] = int(g["count"]) + int(it.get("count", 0) or 0)
        except (TypeError, ValueError):
            pass
        try:
            g["count_ae"] = int(g["count_ae"]) + int(it.get("count_ae", 0) or 0)
        except (TypeError, ValueError):
            pass
        try:
            g["total"] = float(g["total"]) + float(it.get("total", 0) or 0)
        except (TypeError, ValueError):
            pass
        src = it.get("pdf_source") or it.get("pdf_path")
        if src and src not in g["pdf_sources"]:
            g["pdf_sources"].append(src)
    # Round totals for clean reporting; counts stay integer.
    for g in groups.values():
        if isinstance(g["total"], float):
            g["total"] = round(g["total"], 1)
    return groups
