"""
legend_validator.py — Validate that each legend item can be uniquely identified.

For each pair of legend items, checks whether the parser can distinguish
them using available signals: text markers, visual templates, and color.

Validation logic:
  - Same text marker across items -> ERROR (legend malformed)
  - Different text markers -> OK (text distinguishes them)
  - No text markers + different visual templates -> OK
  - No text markers + similar visual templates + different colors -> OK (if pipeline uses color)
  - No text markers + similar visual templates + same/no color -> PROBLEM (can't distinguish)

Usage:
    python legend_validator.py <path.pdf>
"""

from __future__ import annotations

import io
import sys
from dataclasses import dataclass, field
from typing import Optional

import cv2
import numpy as np

from pdf_legend_parser import parse_legend, LegendResult, LegendItem
from pdf_count_text import count_symbols
from pdf_count_visual import (
    _extract_symbol_images,
    _preprocess_template,
    _preprocess_color_template,
)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class ConflictPair:
    """A pair of legend items that may be confused."""
    index_a: int
    index_b: int
    symbol_a: str
    symbol_b: str
    description_a: str
    description_b: str
    conflict_type: str       # "same_text", "visual_similar_no_text",
                             # "visual_similar_same_color", "visual_similar_no_color"
    visual_similarity: float = 0.0  # 0..1 (template match score)
    color_a: str = ""
    color_b: str = ""
    distinguishable: bool = False  # True if the pipeline CAN tell them apart
    resolution: str = ""     # how they can be distinguished ("text", "color", "")


@dataclass
class SymbolValidation:
    """Validation status for a single legend item."""
    index: int
    symbol: str              # text marker
    description: str         # equipment name
    color: str               # detected color
    has_text_marker: bool    # whether it has a non-empty text symbol
    has_visual_template: bool  # whether a visual template was extracted
    conflicts: list[int] = field(default_factory=list)  # indices of conflicting items
    status: str = "ok"       # "ok", "conflict", "unresolvable"
    notes: str = ""


@dataclass
class ValidationResult:
    """Complete validation result."""
    items: list[SymbolValidation] = field(default_factory=list)
    conflicts: list[ConflictPair] = field(default_factory=list)
    total: int = 0
    ok_count: int = 0
    conflict_count: int = 0
    unresolvable_count: int = 0


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Threshold for considering two visual templates "similar"
VISUAL_SIMILARITY_THRESHOLD = 0.70

# Minimum template size for meaningful comparison
MIN_TEMPLATE_PIXELS = 50


# ---------------------------------------------------------------------------
# Visual template similarity
# ---------------------------------------------------------------------------

def _compute_template_similarity(
    img_a: np.ndarray,
    img_b: np.ndarray,
) -> float:
    """
    Compute similarity between two symbol template images.

    Uses OpenCV template matching (normalized cross-correlation).
    Resizes the smaller template to match the larger one's dimensions.

    Returns similarity score 0..1.
    """
    gray_a = _preprocess_template(img_a)
    gray_b = _preprocess_template(img_b)

    if gray_a.size < MIN_TEMPLATE_PIXELS or gray_b.size < MIN_TEMPLATE_PIXELS:
        return 0.0

    # Make both the same size for comparison
    h_a, w_a = gray_a.shape[:2]
    h_b, w_b = gray_b.shape[:2]

    # Resize smaller to match larger
    target_h = max(h_a, h_b)
    target_w = max(w_a, w_b)

    if target_h < 5 or target_w < 5:
        return 0.0

    resized_a = cv2.resize(gray_a, (target_w, target_h), interpolation=cv2.INTER_LINEAR)
    resized_b = cv2.resize(gray_b, (target_w, target_h), interpolation=cv2.INTER_LINEAR)

    # Normalized cross-correlation
    result = cv2.matchTemplate(resized_a, resized_b, cv2.TM_CCOEFF_NORMED)
    similarity = float(result[0, 0]) if result.size > 0 else 0.0

    return max(0.0, min(1.0, similarity))


# ---------------------------------------------------------------------------
# Core validation logic
# ---------------------------------------------------------------------------

def validate_legend_symbols(pdf_path: str) -> ValidationResult:
    """
    For each pair of legend items, verify the parser can distinguish them.

    Uses text markers, visual template similarity, and color to determine
    whether items can be uniquely identified.

    Args:
        pdf_path: Path to the PDF file.

    Returns:
        ValidationResult with per-item status and conflict pairs.
    """
    # 1. Parse legend
    legend_result = parse_legend(pdf_path)

    if not legend_result.items:
        return ValidationResult()

    n = len(legend_result.items)

    # 2. Extract visual templates for all items
    symbol_images: dict[int, Optional[np.ndarray]] = {}
    try:
        symbol_data = _extract_symbol_images(pdf_path, legend_result)
        for idx, item, img in symbol_data:
            symbol_images[idx] = img
    except Exception:
        pass

    # 3. Build per-item validation records
    validations: list[SymbolValidation] = []
    for idx, item in enumerate(legend_result.items):
        sv = SymbolValidation(
            index=idx,
            symbol=item.symbol,
            description=item.description,
            color=item.color,
            has_text_marker=bool(item.symbol),
            has_visual_template=symbol_images.get(idx) is not None,
        )
        validations.append(sv)

    # 4. Check all pairs for potential conflicts
    conflicts: list[ConflictPair] = []

    for i in range(n):
        for j in range(i + 1, n):
            item_a = legend_result.items[i]
            item_b = legend_result.items[j]
            sv_a = validations[i]
            sv_b = validations[j]

            # Case 1: Both have text markers
            if item_a.symbol and item_b.symbol:
                if item_a.symbol == item_b.symbol:
                    # Same text marker => ERROR
                    cp = ConflictPair(
                        index_a=i, index_b=j,
                        symbol_a=item_a.symbol, symbol_b=item_b.symbol,
                        description_a=item_a.description, description_b=item_b.description,
                        conflict_type="same_text",
                        color_a=item_a.color, color_b=item_b.color,
                        distinguishable=False,
                        resolution="",
                    )
                    conflicts.append(cp)
                    sv_a.conflicts.append(j)
                    sv_b.conflicts.append(i)
                # Different text markers => OK, no conflict
                continue

            # Case 2: One or both have no text marker => need visual distinction
            # Only relevant if both have visual templates
            img_a = symbol_images.get(i)
            img_b = symbol_images.get(j)

            if img_a is None or img_b is None:
                # Can't compare visually. If one has text and other doesn't,
                # they're already distinguishable by that.
                if item_a.symbol != item_b.symbol:
                    continue  # one has text, other doesn't => different methods
                # Both have no text and no visual => problem but rare
                if img_a is None and img_b is None and not item_a.symbol and not item_b.symbol:
                    cp = ConflictPair(
                        index_a=i, index_b=j,
                        symbol_a=item_a.symbol, symbol_b=item_b.symbol,
                        description_a=item_a.description, description_b=item_b.description,
                        conflict_type="no_distinguishing_features",
                        color_a=item_a.color, color_b=item_b.color,
                        distinguishable=False,
                        resolution="",
                    )
                    conflicts.append(cp)
                    sv_a.conflicts.append(j)
                    sv_b.conflicts.append(i)
                continue

            # Compare visual templates
            similarity = _compute_template_similarity(img_a, img_b)

            if similarity < VISUAL_SIMILARITY_THRESHOLD:
                # Visually different => OK
                continue

            # Visually similar — check if color can distinguish them
            color_a = item_a.color
            color_b = item_b.color

            if color_a and color_b and color_a != color_b:
                # Different colors => distinguishable by color
                cp = ConflictPair(
                    index_a=i, index_b=j,
                    symbol_a=item_a.symbol, symbol_b=item_b.symbol,
                    description_a=item_a.description, description_b=item_b.description,
                    conflict_type="visual_similar_different_color",
                    visual_similarity=round(similarity, 3),
                    color_a=color_a, color_b=color_b,
                    distinguishable=True,
                    resolution="color",
                )
                conflicts.append(cp)
                # Not a real problem — they CAN be distinguished
                continue

            # Same color or no color => can't distinguish
            if color_a == color_b and color_a:
                conflict_type = "visual_similar_same_color"
            else:
                conflict_type = "visual_similar_no_color"

            # Check if text markers help (one has text, other doesn't)
            if item_a.symbol and not item_b.symbol:
                # Text method handles A, visual handles B => distinguishable
                cp = ConflictPair(
                    index_a=i, index_b=j,
                    symbol_a=item_a.symbol, symbol_b=item_b.symbol,
                    description_a=item_a.description, description_b=item_b.description,
                    conflict_type=conflict_type,
                    visual_similarity=round(similarity, 3),
                    color_a=color_a, color_b=color_b,
                    distinguishable=True,
                    resolution="text",
                )
                conflicts.append(cp)
                continue
            if item_b.symbol and not item_a.symbol:
                cp = ConflictPair(
                    index_a=i, index_b=j,
                    symbol_a=item_a.symbol, symbol_b=item_b.symbol,
                    description_a=item_a.description, description_b=item_b.description,
                    conflict_type=conflict_type,
                    visual_similarity=round(similarity, 3),
                    color_a=color_a, color_b=color_b,
                    distinguishable=True,
                    resolution="text",
                )
                conflicts.append(cp)
                continue

            # Both have no text, similar visuals, same/no color => PROBLEM
            cp = ConflictPair(
                index_a=i, index_b=j,
                symbol_a=item_a.symbol, symbol_b=item_b.symbol,
                description_a=item_a.description, description_b=item_b.description,
                conflict_type=conflict_type,
                visual_similarity=round(similarity, 3),
                color_a=color_a, color_b=color_b,
                distinguishable=False,
                resolution="",
            )
            conflicts.append(cp)
            sv_a.conflicts.append(j)
            sv_b.conflicts.append(i)

    # 5. Set final status for each item
    ok_count = 0
    conflict_count = 0
    unresolvable_count = 0

    for sv in validations:
        if not sv.conflicts:
            sv.status = "ok"
            ok_count += 1
        else:
            # Check if all conflicts are resolvable
            all_resolvable = True
            for cp in conflicts:
                if (cp.index_a == sv.index or cp.index_b == sv.index):
                    if not cp.distinguishable:
                        all_resolvable = False
                        break
            if all_resolvable:
                sv.status = "ok"
                sv.notes = "conflicts resolvable"
                ok_count += 1
            else:
                # Check if it's just a warning (color could help) or unresolvable
                has_unresolvable = False
                for cp in conflicts:
                    if (cp.index_a == sv.index or cp.index_b == sv.index):
                        if not cp.distinguishable:
                            has_unresolvable = True
                            break
                if has_unresolvable:
                    sv.status = "unresolvable"
                    sv.notes = "cannot distinguish from similar item(s)"
                    unresolvable_count += 1
                else:
                    sv.status = "conflict"
                    sv.notes = "conflicts exist but resolvable"
                    conflict_count += 1

    return ValidationResult(
        items=validations,
        conflicts=conflicts,
        total=n,
        ok_count=ok_count,
        conflict_count=conflict_count,
        unresolvable_count=unresolvable_count,
    )


# ---------------------------------------------------------------------------
# CLI interface
# ---------------------------------------------------------------------------

def main() -> None:
    """CLI entry point for testing legend validation."""
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", errors="replace"
        )

    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print("Usage: python legend_validator.py <path.pdf>")
        sys.exit(1)

    pdf_path = args[0]
    print(f"Validating legend symbols: {pdf_path}")
    print()

    result = validate_legend_symbols(pdf_path)

    if result.total == 0:
        print("No legend found.")
        sys.exit(0)

    print(f"Legend items: {result.total}")
    print(f"  OK: {result.ok_count}")
    print(f"  Conflicts (resolvable): {result.conflict_count}")
    print(f"  Unresolvable: {result.unresolvable_count}")
    print()

    # Print per-item status
    print("Per-item status:")
    print(f"  {'#':<4s} {'Symbol':<10s} {'Color':<8s} {'Status':<14s} Description")
    print("-" * 80)

    for sv in result.items:
        sym = sv.symbol if sv.symbol else "(none)"
        color = sv.color if sv.color else "-"
        status_icon = {
            "ok": "[OK]",
            "conflict": "[WARN]",
            "unresolvable": "[FAIL]",
        }.get(sv.status, "[?]")
        desc = sv.description[:50]
        notes = f"  -- {sv.notes}" if sv.notes else ""
        print(f"  {sv.index:<3d} {sym:<10s} {color:<8s} {status_icon:<14s} {desc}{notes}")

    # Print conflicts
    if result.conflicts:
        print()
        print("Conflict pairs:")
        print("-" * 80)
        for cp in result.conflicts:
            sym_a = cp.symbol_a or "(none)"
            sym_b = cp.symbol_b or "(none)"
            dist = "YES" if cp.distinguishable else "NO"
            res = f" via {cp.resolution}" if cp.resolution else ""
            print(f"  #{cp.index_a} ({sym_a}) vs #{cp.index_b} ({sym_b})")
            print(f"    Type: {cp.conflict_type}")
            if cp.visual_similarity > 0:
                print(f"    Visual similarity: {cp.visual_similarity:.1%}")
            if cp.color_a or cp.color_b:
                print(f"    Colors: {cp.color_a or '-'} vs {cp.color_b or '-'}")
            print(f"    Distinguishable: {dist}{res}")
            print(f"    A: {cp.description_a[:60]}")
            print(f"    B: {cp.description_b[:60]}")
            print()

    print("-" * 80)


if __name__ == "__main__":
    main()
