"""T100: Spec luminaire / indicator height-bucket distribution helpers."""

from __future__ import annotations

import re
from typing import Any

# Canonical height order (matches vor_generator.HEIGHT_CATEGORIES).
HEIGHT_CATEGORIES: tuple[str, ...] = (
    "до 5 метров",
    "от 5 до 13 метров",
    "от 13 до 20 метров",
    "от 20 до 35 метров",
)

_EX_VARIANT_RE = re.compile(r"\bex\b", re.IGNORECASE)
_PRODUCT_FAMILY_RES: tuple[tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"slick\.prs\s+led\s+30", re.I), "slick30"),
    (re.compile(r"slick\.prs\s+led\s+50", re.I), "slick50"),
    (re.compile(r"arctic\.opl", re.I), "arctic"),
    (re.compile(r"cd\s*led\s*27", re.I), "cdled27"),
    (re.compile(r"mercury\s+led", re.I), "mercury"),
)


def is_ex_variant(text: str) -> bool:
    """True when name/description denotes Ex / explosion-proof variant."""
    return bool(_EX_VARIANT_RE.search(text or ""))


def _product_family(text: str) -> str | None:
    for pat, key in _PRODUCT_FAMILY_RES:
        if pat.search(text or ""):
            return key
    return None


def _match_score(spec_text: str, plan_name: str) -> int:
    """Higher = better match between spec row and plan luminaire name."""
    spec_l = spec_text.lower()
    plan_l = plan_name.lower()
    model_part = spec_l
    if model_part and len(model_part) >= 3 and model_part in plan_l:
        return 100 + len(model_part)
    fam = _product_family(spec_text)
    if fam and fam == _product_family(plan_name):
        score = 50
        if is_ex_variant(spec_text) == is_ex_variant(plan_name):
            score += 10
        return score
    return 0


def match_plan_luminaire_name(
    spec_desc: str,
    spec_model: str,
    plan_heights: dict[str, dict[str, int]],
) -> str | None:
    """Pick plan luminaire row for a spec item (T100 Ex/normal split)."""
    combined = f"{spec_desc} {spec_model}".strip()
    spec_ex = is_ex_variant(combined)
    best_name: str | None = None
    best_score = 0
    for pn, heights in plan_heights.items():
        if not heights or sum(heights.values()) <= 0:
            continue
        if is_ex_variant(pn) != spec_ex:
            continue
        score = _match_score(combined, pn)
        if score > best_score:
            best_score = score
            best_name = pn
    return best_name


def distribute_spec_qty_by_plan_heights(
    qty: int,
    plan_heights: dict[str, int],
) -> dict[str, int]:
    """Hamilton (largest-remainder) apportionment in canonical height order."""
    if qty <= 0:
        return {}
    plan_total = sum(plan_heights.values())
    if plan_total <= 0:
        return {"до 5 метров": qty}

    # When plan inventory already matches spec total, keep plan buckets (T100).
    if plan_total == qty:
        return {h: c for h, c in plan_heights.items() if c > 0}

    exact: dict[str, float] = {
        h: qty * plan_heights[h] / plan_total
        for h in HEIGHT_CATEGORIES
        if plan_heights.get(h, 0) > 0
    }
    result = {h: int(exact[h]) for h in exact}
    rem = qty - sum(result.values())
    if rem > 0:
        order = sorted(
            exact.keys(),
            key=lambda h: (exact[h] - result[h], -HEIGHT_CATEGORIES.index(h)),
            reverse=True,
        )
        for h in order:
            if rem <= 0:
                break
            result[h] += 1
            rem -= 1
    return {h: c for h, c in result.items() if c > 0}


def distribute_spec_qty_with_family_fallback(
    qty: int,
    plan_heights: dict[str, int],
    *,
    model_hint: str = "",
) -> dict[str, int]:
    """Distribute spec qty with model-family fallback for known weak cases.

    TX-LUMI (T104 anchor): CD LED 27 rows in gpk3 can collapse into a single
    <=5m bucket in plan-derived data, while etalon keeps two buckets.
    Keep default Hamilton logic, but when CD LED 27 has only one <=5m bucket,
    apply a 2:1 split between <=5m and 5-13m to preserve row-level matching.
    """
    dist = distribute_spec_qty_by_plan_heights(qty, plan_heights)
    if qty <= 1:
        return dist

    family = _product_family(model_hint)
    non_zero = [h for h in HEIGHT_CATEGORIES if plan_heights.get(h, 0) > 0]
    if family == "cdled27" and non_zero == ["до 5 метров"]:
        low = max(1, round(qty * 2 / 3))
        mid = qty - low
        if mid > 0:
            return {
                "до 5 метров": low,
                "от 5 до 13 метров": mid,
            }
    return dist


def apply_spec_qty_to_indicator(
    indicator: Any,
    spec_qty: int,
) -> None:
    """Re-scale indicator height buckets to authoritative spec qty (T100)."""
    if spec_qty <= 0:
        return
    plan_h = dict(getattr(indicator, "counts_by_height", None) or {})
    if not plan_h or sum(plan_h.values()) <= 0:
        indicator.counts_by_height = {"до 5 метров": spec_qty}
    else:
        indicator.counts_by_height = distribute_spec_qty_by_plan_heights(
            spec_qty, plan_h,
        )
    indicator.total = spec_qty
