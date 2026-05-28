"""T070: Spec-authoritative resolution for hardware and cable trays."""

from __future__ import annotations

from typing import Any, Callable

SPEC_AUTHORITATIVE_FLAG = True
SPEC_AUTH_TOLERANCE_PCT = 0.05

_METER_UNITS = frozenset({"м", "м.", "м.п.", "м. п."})

_FASTENER_KW = ("винт", "гайка", "шайба", "саморез")
_MOUNT_KW = ("стойка", "кронштейн", "опора")
_TRAY_KW = ("лоток", "лотка", "лотков", "кабельный канал", "кабель-канал")
_ACCESSORY_KW = (
    "тройник", "угол", "заглушк", "разделител",
    "соединитель лотк", "соединительная пластина",
    "т-ответвитель", "скоба для настенного", "консоль сварн",
    "подвес для лотк", "прижим лотк",
)


def hardware_kind(description: str, category: str = "") -> str | None:
    """Map a spec/plan row to a spec-authoritative hardware family."""
    nl = (description or "").lower().replace("ё", "е")
    if category == "tray":
        if any(k in nl for k in _TRAY_KW):
            return "tray"
        if any(k in nl for k in _ACCESSORY_KW):
            return "accessory"
        return "accessory"
    if any(k in nl for k in _FASTENER_KW):
        return "fastener"
    if any(k in nl for k in ("болт",)) and "анкер" not in nl:
        return "fastener"
    if any(k in nl for k in _MOUNT_KW):
        return "mount"
    if any(k in nl for k in _TRAY_KW):
        return "tray"
    if any(k in nl for k in _ACCESSORY_KW):
        return "accessory"
    return None


def is_spec_authoritative(description: str, category: str = "") -> bool:
    return hardware_kind(description, category) is not None


def qty_delta_pct(reference: float, other: float) -> float:
    return abs(float(reference) - float(other)) / max(abs(float(reference)), 1.0)


def _spec_tray_total_m(tray_items: list[Any]) -> float:
    total = 0.0
    for item in tray_items:
        unit = (getattr(item, "unit", None) or "").strip().lower()
        qty = getattr(item, "quantity", 0)
        if unit in _METER_UNITS and isinstance(qty, (int, float)) and qty > 0:
            if hardware_kind(getattr(item, "description", ""), "tray") == "tray":
                total += float(qty)
    return total


def apply_spec_auth_tray_lengths(
    tray_lengths: dict[str, int],
    tray_items: list[Any],
    events: list[dict],
    *,
    name_key_fn: Callable[[str], str] | None = None,
) -> dict[str, int]:
    """Drop geometric tray totals when spec tray metres disagree by >5%."""
    if not SPEC_AUTHORITATIVE_FLAG:
        return tray_lengths
    spec_m = _spec_tray_total_m(tray_items)
    geom_m = float(sum(tray_lengths.values()))
    if spec_m <= 0 or geom_m <= 0:
        return tray_lengths
    if qty_delta_pct(spec_m, geom_m) <= SPEC_AUTH_TOLERANCE_PCT:
        return tray_lengths
    events.append({
        "name": "cable_tray_total",
        "category": "tray",
        "category_event": "hardware",
        "hardware_kind": "tray",
        "reason": "spec_auth",
        "chosen_source": "spec_table",
        "chosen_quantity": round(spec_m),
        "rejected_source": "plan_derived",
        "rejected_quantity": round(geom_m),
        "delta_pct": round(qty_delta_pct(spec_m, geom_m) * 100.0, 2),
        "action": "suppress_geometric_tray_lengths",
    })
    return {}


def resolve_canonical_with_spec_auth(
    facts: list[Any],
    base_resolver: Callable[[list[Any]], Any],
    events: list[dict],
    *,
    key: str = "",
) -> Any:
    """Prefer spec_table qty for hardware/tray when derived sources differ >5%."""
    row = base_resolver(facts)
    if not SPEC_AUTHORITATIVE_FLAG or row is None:
        return row

    spec_facts = [
        f for f in facts
        if getattr(f, "source", None) == "spec_table"
        and isinstance(getattr(f, "quantity", None), (int, float))
        and float(f.quantity) > 0
    ]
    if not spec_facts:
        return row

    spec_pick = max(spec_facts, key=lambda f: float(f.quantity))
    hw_kind = hardware_kind(
        getattr(spec_pick, "name", ""),
        getattr(spec_pick, "category", ""),
    )
    if not hw_kind:
        return row

    spec_qty = float(spec_pick.quantity)
    rejected: list[dict] = []
    for fact in facts:
        if getattr(fact, "source", None) == "spec_table":
            continue
        other_qty = float(getattr(fact, "quantity", 0) or 0)
        if other_qty <= 0:
            continue
        delta_pct = qty_delta_pct(spec_qty, other_qty)
        if delta_pct > SPEC_AUTH_TOLERANCE_PCT:
            rejected.append({
                "source": getattr(fact, "source", ""),
                "quantity": other_qty,
                "delta_pct": round(delta_pct * 100.0, 2),
            })

    if not rejected:
        return row

    events.append({
        "key": key,
        "name": getattr(spec_pick, "name", row.name),
        "category": getattr(spec_pick, "category", row.category),
        "category_event": "hardware",
        "hardware_kind": hw_kind,
        "reason": "spec_auth",
        "chosen_source": "spec_table",
        "chosen_quantity": int(round(spec_qty)),
        "unit": getattr(spec_pick, "unit", row.unit),
        "rejected_over_5pct": rejected,
    })

    heights = dict(getattr(spec_pick, "counts_by_height", None) or {})
    if not heights:
        heights = dict(getattr(row, "counts_by_height", None) or {})
    if not heights:
        heights = {"до 5 метров": int(round(spec_qty))}

    from vor_generator import CanonicalRow

    return CanonicalRow(
        source="spec_table",
        name=getattr(spec_pick, "name", row.name),
        category=getattr(spec_pick, "category", row.category),
        quantity=int(round(spec_qty)),
        unit=(getattr(spec_pick, "unit", None) or getattr(row, "unit", None) or "шт"),
        counts_by_height=heights,
    )
