"""
pdf_color_layers.py -- Split PDF drawing into independent color layers.

Separates drawing content by color into independent layers for targeted
processing. Each layer can be processed independently -- red symbols are
searched only in the red layer, etc.

Color layers:
  - black:  walls, axes, dimensions (H:any, S<30, V<80)
  - grey:   furniture, hatching (H:any, S<30, 80<=V<200)
  - red:    emergency equipment (H:0-10 or 170-180, S>50, V>50)
  - blue:   working equipment (H:100-130, S>50, V>50)

Two approaches combined:
  1. Vector: pdfplumber stroking_color / non_stroking_color for lines, rects, chars
  2. Raster: HSV masks on PyMuPDF-rendered page image

Usage:
    python pdf_color_layers.py <path.pdf> [--page N] [--dpi 150] [--save]

Dependencies:
    pdfplumber, fitz (PyMuPDF), cv2 (OpenCV), numpy
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from typing import Optional

import cv2
import fitz
import numpy as np
import pdfplumber

from pdf_legend_parser import _normalize_color, _classify_rgb


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

LAYER_NAMES = ("red", "blue", "black", "grey")

# HSV ranges for raster layer extraction (OpenCV: H 0-180, S 0-255, V 0-255)
# Red wraps around hue 0, so we need two ranges
_HSV_RANGES = {
    "red": [
        ((0, 60, 60), (10, 255, 255)),
        ((170, 60, 60), (180, 255, 255)),
    ],
    "blue": [
        ((100, 60, 60), (130, 255, 255)),
    ],
    "black": [
        ((0, 0, 0), (180, 50, 80)),
    ],
    "grey": [
        ((0, 0, 80), (180, 50, 200)),
    ],
}

DEFAULT_RENDER_DPI = 150


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class VectorLayer:
    """Vector elements (lines, rects, chars) belonging to one color layer."""
    color: str  # "red", "blue", "black", "grey"
    lines: list[dict] = field(default_factory=list)
    rects: list[dict] = field(default_factory=list)
    chars: list[dict] = field(default_factory=list)

    @property
    def element_count(self) -> int:
        return len(self.lines) + len(self.rects) + len(self.chars)


@dataclass
class RasterLayer:
    """Raster (pixel) representation of one color layer."""
    color: str  # "red", "blue", "black", "grey"
    mask: np.ndarray = field(default=None, repr=False)  # binary mask H x W
    image: np.ndarray = field(default=None, repr=False)  # BGR image with only this color

    @property
    def pixel_count(self) -> int:
        if self.mask is None:
            return 0
        return int(np.count_nonzero(self.mask))


@dataclass
class ColorLayers:
    """Complete color layer decomposition of a PDF page."""
    page_index: int = 0
    page_width: float = 0.0
    page_height: float = 0.0
    dpi: int = DEFAULT_RENDER_DPI
    vector: dict[str, VectorLayer] = field(default_factory=dict)
    raster: dict[str, RasterLayer] = field(default_factory=dict)
    full_image: np.ndarray = field(default=None, repr=False)  # full rendered BGR

    def get_vector(self, color: str) -> VectorLayer:
        """Get vector layer by color name."""
        return self.vector.get(color, VectorLayer(color=color))

    def get_raster(self, color: str) -> RasterLayer:
        """Get raster layer by color name."""
        return self.raster.get(color, RasterLayer(color=color))

    def get_layer_image(self, color: str) -> Optional[np.ndarray]:
        """Get masked BGR image for a specific color layer."""
        layer = self.raster.get(color)
        if layer is None or layer.image is None:
            return None
        return layer.image

    def get_combined_image(self, colors: list[str]) -> np.ndarray:
        """Get BGR image combining multiple color layers."""
        h, w = self.full_image.shape[:2]
        combined_mask = np.zeros((h, w), dtype=np.uint8)
        for c in colors:
            layer = self.raster.get(c)
            if layer is not None and layer.mask is not None:
                combined_mask = cv2.bitwise_or(combined_mask, layer.mask)
        result = np.full((h, w, 3), 255, dtype=np.uint8)
        result[combined_mask > 0] = self.full_image[combined_mask > 0]
        return result

    def summary(self) -> dict:
        """Return a summary dict of layer statistics."""
        return {
            "page_index": self.page_index,
            "page_size": (self.page_width, self.page_height),
            "dpi": self.dpi,
            "layers": {
                c: {
                    "vector_elements": self.get_vector(c).element_count,
                    "raster_pixels": self.get_raster(c).pixel_count,
                }
                for c in LAYER_NAMES
            },
        }


# ---------------------------------------------------------------------------
# Vector layer extraction (pdfplumber)
# ---------------------------------------------------------------------------

def _classify_element_color(element: dict, color_key: str = "stroking_color") -> str:
    """Classify a pdfplumber element into a color layer.

    Returns one of: "red", "blue", "black", "grey", or "" (unclassified).
    """
    raw = element.get(color_key)
    rgb = _normalize_color(raw)
    if rgb is None:
        return "black"  # default: uncolored elements are black
    return _classify_rgb(rgb) or "black"


def extract_vector_layers(
    pdf_path: str,
    page_index: int = 0,
) -> dict[str, VectorLayer]:
    """Extract vector elements from a PDF page, grouped by color layer.

    Uses pdfplumber to read lines, rects, and chars with their
    stroking_color / non_stroking_color attributes.
    """
    layers: dict[str, VectorLayer] = {c: VectorLayer(color=c) for c in LAYER_NAMES}

    with pdfplumber.open(pdf_path) as pdf:
        if page_index >= len(pdf.pages):
            return layers
        page = pdf.pages[page_index]

        # Lines
        for line in (page.lines or []):
            color = _classify_element_color(line, "stroking_color")
            if color in layers:
                layers[color].lines.append(line)

        # Rects
        for rect in (page.rects or []):
            stroke_color = _classify_element_color(rect, "stroking_color")
            fill_color = _classify_element_color(rect, "non_stroking_color")
            # Prefer the more "interesting" color (non-black)
            color = stroke_color
            if stroke_color == "black" and fill_color in ("red", "blue", "grey"):
                color = fill_color
            if color in layers:
                layers[color].rects.append(rect)

        # Chars (text characters)
        for ch in (page.chars or []):
            color = _classify_element_color(ch, "non_stroking_color")
            if color in layers:
                layers[color].chars.append(ch)

    return layers


# ---------------------------------------------------------------------------
# Raster layer extraction (PyMuPDF + OpenCV HSV)
# ---------------------------------------------------------------------------

def _render_page(pdf_path: str, page_index: int, dpi: int) -> np.ndarray:
    """Render a PDF page to BGR numpy array using PyMuPDF."""
    doc = fitz.open(pdf_path)
    if page_index >= len(doc):
        doc.close()
        raise ValueError(f"Page {page_index} out of range (0-{len(doc) - 1})")
    page = doc[page_index]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, 3
    )
    # PyMuPDF returns RGB, OpenCV uses BGR
    bgr = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    doc.close()
    return bgr


def _build_hsv_mask(hsv: np.ndarray, color: str) -> np.ndarray:
    """Build a binary mask for a specific color using HSV ranges."""
    ranges = _HSV_RANGES.get(color, [])
    if not ranges:
        return np.zeros(hsv.shape[:2], dtype=np.uint8)

    combined = np.zeros(hsv.shape[:2], dtype=np.uint8)
    for lo, hi in ranges:
        lo_arr = np.array(lo, dtype=np.uint8)
        hi_arr = np.array(hi, dtype=np.uint8)
        mask = cv2.inRange(hsv, lo_arr, hi_arr)
        combined = cv2.bitwise_or(combined, mask)

    return combined


def extract_raster_layers(
    pdf_path: str,
    page_index: int = 0,
    dpi: int = DEFAULT_RENDER_DPI,
    page_bgr: Optional[np.ndarray] = None,
) -> tuple[dict[str, RasterLayer], np.ndarray]:
    """Extract raster color layers from a rendered PDF page.

    Uses HSV color space to separate the rendered image into color layers.
    Each layer gets a binary mask and a masked BGR image.

    Args:
        pdf_path: Path to the PDF file.
        page_index: Page index (0-based).
        dpi: Render resolution.
        page_bgr: Pre-rendered page image (optional, avoids re-rendering).

    Returns:
        (layers_dict, full_bgr_image)
    """
    if page_bgr is None:
        page_bgr = _render_page(pdf_path, page_index, dpi)

    hsv = cv2.cvtColor(page_bgr, cv2.COLOR_BGR2HSV)

    layers: dict[str, RasterLayer] = {}

    # White background mask (to exclude from all layers)
    white_mask = cv2.inRange(hsv, np.array([0, 0, 220]), np.array([180, 30, 255]))

    for color in LAYER_NAMES:
        mask = _build_hsv_mask(hsv, color)
        # Exclude white background pixels
        mask = cv2.bitwise_and(mask, cv2.bitwise_not(white_mask))

        # Build masked image: white background + colored pixels
        layer_img = np.full_like(page_bgr, 255)
        layer_img[mask > 0] = page_bgr[mask > 0]

        layers[color] = RasterLayer(color=color, mask=mask, image=layer_img)

    return layers, page_bgr


# ---------------------------------------------------------------------------
# Main API: split page into color layers
# ---------------------------------------------------------------------------

def split_page_colors(
    pdf_path: str,
    page_index: int = 0,
    dpi: int = DEFAULT_RENDER_DPI,
    vector: bool = True,
    raster: bool = True,
    page_bgr: Optional[np.ndarray] = None,
) -> ColorLayers:
    """Split a PDF page into independent color layers.

    This is the main entry point. Returns a ColorLayers object with both
    vector (pdfplumber) and raster (OpenCV HSV) decompositions.

    Args:
        pdf_path: Path to the PDF file.
        page_index: Page index (0-based).
        dpi: Render DPI for raster layers.
        vector: Whether to extract vector layers.
        raster: Whether to extract raster layers.
        page_bgr: Pre-rendered page (avoids re-rendering).

    Returns:
        ColorLayers with vector and raster layers for each color.
    """
    result = ColorLayers(page_index=page_index, dpi=dpi)

    # Get page dimensions
    with pdfplumber.open(pdf_path) as pdf:
        if page_index < len(pdf.pages):
            p = pdf.pages[page_index]
            result.page_width = float(p.width)
            result.page_height = float(p.height)

    if vector:
        result.vector = extract_vector_layers(pdf_path, page_index)

    if raster:
        raster_layers, full_img = extract_raster_layers(
            pdf_path, page_index, dpi, page_bgr
        )
        result.raster = raster_layers
        result.full_image = full_img

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    """CLI entry point for testing color layer extraction."""
    import argparse
    import json
    from pathlib import Path

    parser = argparse.ArgumentParser(
        description="Split PDF drawing into color layers"
    )
    parser.add_argument("pdf_path", help="Path to PDF file")
    parser.add_argument("--page", type=int, default=0, help="Page index (0-based)")
    parser.add_argument("--dpi", type=int, default=150, help="Render DPI")
    parser.add_argument("--save", action="store_true",
                        help="Save layer images to disk")
    args = parser.parse_args()

    pdf_path = args.pdf_path
    if not Path(pdf_path).exists():
        print(f"File not found: {pdf_path}")
        sys.exit(1)

    print(f"Splitting: {Path(pdf_path).name}, page {args.page}, dpi={args.dpi}")
    print("-" * 60)

    layers = split_page_colors(pdf_path, args.page, args.dpi)

    # Print summary
    summary = layers.summary()
    for color in LAYER_NAMES:
        info = summary["layers"][color]
        vec = info["vector_elements"]
        pix = info["raster_pixels"]
        print(f"  {color:8s}: {vec:6d} vector elements, {pix:8d} raster pixels")

    print(f"\n  Page size: {summary['page_size'][0]:.0f} x {summary['page_size'][1]:.0f} pt")

    if args.save:
        stem = Path(pdf_path).stem
        out_dir = Path(pdf_path).parent
        for color in LAYER_NAMES:
            img = layers.get_layer_image(color)
            if img is not None:
                out_path = out_dir / f"{stem}_layer_{color}_p{args.page}.png"
                cv2.imwrite(str(out_path), img)
                print(f"  Saved: {out_path.name}")

        # Also save combined red+blue (equipment only)
        equip_img = layers.get_combined_image(["red", "blue"])
        out_path = out_dir / f"{stem}_layer_equipment_p{args.page}.png"
        cv2.imwrite(str(out_path), equip_img)
        print(f"  Saved: {out_path.name}")

    print("\nDone.")


if __name__ == "__main__":
    main()
