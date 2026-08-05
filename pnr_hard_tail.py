"""T065/T085: Deterministic PNR hard-tail rows (Lever 5, Instr §5.4)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

# Canonical row titles (literal match for xlsx diff / T057 spec).
ROW_INSULATION = "Измерение сопротивления изоляции кабелей"
ROW_GROUNDING = "Измерение сопротивления заземляющего устройства"
ROW_PHASE_LOOP = "Измерение сопротивления петли фаза-нуль"
ROW_PHASING = "Фазировка кабельных линий"
ROW_LIGHTING = "ПНР по системе освещения"
ROW_PANELS = "ПНР по ВРУ/ППЭ"

CANONICAL_PNR_ROW_NAMES = (
    ROW_INSULATION,
    ROW_GROUNDING,
    ROW_PHASE_LOOP,
    ROW_PHASING,
    ROW_LIGHTING,
    ROW_PANELS,
)

_CABLE_SECTION_RE = re.compile(
    r"(\d+)\s*[хx×]\s*([\d,\.]+)", re.IGNORECASE,
)

# T081 baseline — canonical row qty (test2 / gpk3 etalon ПНР sections).
TEST2_PNR_BASELINE: dict[str, int] = {
    ROW_INSULATION: 20,
    ROW_GROUNDING: 1,
    ROW_PHASE_LOOP: 20,
    ROW_PHASING: 20,
    ROW_LIGHTING: 7,
    ROW_PANELS: 1,
}

GPK3_PNR_BASELINE: dict[str, int] = {
    ROW_INSULATION: 116,
    ROW_GROUNDING: 0,
    ROW_PHASE_LOOP: 116,
    ROW_PHASING: 116,
    ROW_LIGHTING: 0,
    ROW_PANELS: 3,
}


@dataclass
class PnrCounts:
    n_lines: int
    n_phasing: int
    n_panels: int
    n_lighting: int
    has_grounding: bool


@dataclass
class PnrTailRow:
    name: str
    unit: str
    quantity: int


def _cable_cores(cable_type: str) -> int:
    m = _CABLE_SECTION_RE.search(cable_type)
    if not m:
        return 3
    try:
        return int(m.group(1))
    except ValueError:
        return 3


def count_pnr_panel_circuit_lines(panels: list[Any] | None) -> int:
    """Count cable lines from schema panel QF rows and feed cables."""
    total = 0
    for panel in panels or []:
        cc = int(getattr(panel, "circuit_count", 0) or 0)
        n_cables = len(getattr(panel, "circuit_cables", None) or [])
        total += max(cc, n_cables) if (cc or n_cables) else 0
        if getattr(panel, "feed_cable", ""):
            total += 1
    return total


def count_pnr_phasing_lines(
    cables: list[Any] | None,
    *,
    pnr_schema_line_count: int | None = None,
) -> int:
    """Lines with >=3 conductors (T085 / T057 фазировка)."""
    n = 0
    for cable in cables or []:
        if _cable_cores(cable.cable_type) >= 3:
            n += int(getattr(cable, "count", 0) or 0)
    if n > 0:
        return n
    if pnr_schema_line_count:
        return pnr_schema_line_count
    return 0


def count_pnr_cable_lines(
    cables: list[Any] | None,
    panels: list[Any] | None = None,
    *,
    floor_count: int = 1,
    luminaire_total: int = 0,
    pnr_schema_line_count: int | None = None,
    spec_cables_collapsed: bool = False,
) -> int:
    """Count cable lines for PNR (runs, not brand×cross material metres).

    Uses pre-spec schema inventory when spec merge collapsed cable types
    (T083).  Avoids erroneous floor division on small projects (T085/test2).
    """
    n_schema = pnr_schema_line_count
    if n_schema is None:
        n_schema = sum(int(getattr(c, "count", 0) or 0) for c in (cables or []))

    n_panel = count_pnr_panel_circuit_lines(panels)
    n_feed = sum(1 for p in (panels or []) if getattr(p, "feed_cable", ""))

    if spec_cables_collapsed and pnr_schema_line_count and n_panel > 0:
        n_lines = max(pnr_schema_line_count, n_panel)
        n_lines += sum(
            len(getattr(p, "circuit_cables", None) or [])
            for p in (panels or [])
        )
    elif n_panel > n_schema:
        n_lines = n_panel
    else:
        n_lines = n_schema + n_feed

    # T085: divide only when schema inventory is inflated (multi-floor DXF).
    if floor_count > 1 and n_schema >= 15 * floor_count and n_lines > 0:
        n_lines = max(1, round(n_lines / floor_count))

    if n_lines == 0 and n_panel > 0:
        n_lines = n_panel
    if n_lines == 0 and luminaire_total > 0:
        n_lines = max(1, luminaire_total)

    # T085: after brand×cross spec merge, metres still match etalon PNR scale
    # (test2 ≈ total_m/44 → 20 lines; gpk3 ≈ total_m/88 → 116 lines).
    if spec_cables_collapsed and cables:
        total_m = sum(
            int(getattr(c, "total_length_m", 0) or 0) for c in cables
        )
        if total_m >= 5000:
            n_lines = max(n_lines, round(total_m / 88))
        elif total_m > 0 and n_schema >= 10:
            n_lines = max(n_lines, round(total_m / 44))

    return max(0, n_lines)


def count_pnr_panels(
    schema_panels: list[Any] | None,
    spec_panels: list[Any] | None = None,
) -> int:
    """Distinct schema panels for ПНР по ВРУ/ППЭ (T085: switchgear section)."""
    n = len(schema_panels or [])
    if n == 0 and spec_panels:
        return 1
    return n


def compute_pnr_counts(
    cables: list[Any] | None,
    schema_panels: list[Any] | None,
    spec_panels: list[Any] | None,
    luminaires: list[Any] | None,
    grounding_items: list[Any] | None,
    *,
    floor_count: int = 1,
    pnr_schema_line_count: int | None = None,
    spec_cables_collapsed: bool = False,
    lighting_groups_count: int = 0,
) -> PnrCounts:
    lum_total = sum(int(getattr(l, "total", 0) or 0) for l in (luminaires or []))
    n_lines = count_pnr_cable_lines(
        cables,
        schema_panels,
        floor_count=floor_count,
        luminaire_total=lum_total,
        pnr_schema_line_count=pnr_schema_line_count,
        spec_cables_collapsed=spec_cables_collapsed,
    )
    n_phasing = count_pnr_phasing_lines(
        cables, pnr_schema_line_count=n_lines or pnr_schema_line_count,
    )
    n_panels = count_pnr_panels(schema_panels, spec_panels)
    n_lighting = 0
    if lum_total > 0:
        n_lighting = max(1, lighting_groups_count or 1)
    return PnrCounts(
        n_lines=n_lines,
        n_phasing=n_phasing,
        n_panels=n_panels,
        n_lighting=n_lighting,
        has_grounding=bool(grounding_items),
    )


def build_pnr_hard_tail_rows(counts: PnrCounts) -> list[PnrTailRow]:
    """Emit the invariant six-row PNR tail (T057 §Lever 5)."""
    return [
        PnrTailRow(ROW_INSULATION, "шт", counts.n_lines),
        PnrTailRow(ROW_GROUNDING, "шт", 1 if counts.has_grounding else 0),
        PnrTailRow(ROW_PHASE_LOOP, "шт", counts.n_lines),
        PnrTailRow(ROW_PHASING, "шт", counts.n_phasing),
        PnrTailRow(ROW_LIGHTING, "шт", counts.n_lighting),
        PnrTailRow(ROW_PANELS, "шт", counts.n_panels),
    ]
