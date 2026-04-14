"""
web_app.py — FastAPI web application for PDF legend analysis.

Provides endpoints for browsing PDFs, rendering pages, parsing legends,
and debugging word extraction.

Usage:
    uvicorn web_app:app --host 0.0.0.0 --port 8050 --reload
    # or
    python web_app.py
"""

from __future__ import annotations

import asyncio
import hashlib
import io
import json as json_mod
import math
import os
import time
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
import pdfplumber
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

import re as re_mod

from equipment_counter import process_pdf as ec_process_pdf
from pdf_legend_parser import parse_legend, LegendResult
from pdf_count_text import count_symbols
from pdf_count_cables import extract_cables
from pdf_count_geometry import measure_cables
from pdf_count_visual import match_symbols, _extract_symbol_images, build_equipment_cluster_bboxes
from vor_work_mapping import map_items as vor_map_items
from legend_validator import validate_legend_symbols

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "Data"
WEB_DIR = BASE_DIR / "web"

# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------

app = FastAPI(
    title="PDF Legend Viewer",
    description="Web viewer for PDF legend analysis in electrical drawings",
    version="1.0.0",
)

# Mount static files and templates
app.mount("/static", StaticFiles(directory=str(WEB_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(WEB_DIR / "templates"))


# ---------------------------------------------------------------------------
# Helpers: folder operations
# ---------------------------------------------------------------------------

def _folder_id(rel_folder: str) -> str:
    """Generate stable URL-safe folder ID from relative folder path."""
    return hashlib.sha256(("FOLDER:" + rel_folder).encode("utf-8")).hexdigest()[:16]


def _id_to_folder(folder_id: str) -> Optional[str]:
    """Resolve folder ID back to relative folder path."""
    seen: set[str] = set()
    for pdf_path in DATA_DIR.rglob("*.pdf"):
        rel = str(pdf_path.relative_to(DATA_DIR)).replace("\\", "/")
        parts = rel.rsplit("/", 1)
        folder = parts[0] if len(parts) > 1 else ""
        if folder not in seen:
            seen.add(folder)
            if _folder_id(folder) == folder_id:
                return folder
    return None


def _folder_files(rel_folder: str) -> list[Path]:
    """Return all PDF files in a specific folder under DATA_DIR."""
    folder_path = DATA_DIR / rel_folder
    if not folder_path.is_dir():
        return []
    return sorted(folder_path.glob("*.pdf"))


templates.env.globals["folder_id"] = _folder_id


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _file_id(rel_path: str) -> str:
    """Generate a stable, URL-safe file ID from a relative path."""
    return hashlib.sha256(rel_path.encode("utf-8")).hexdigest()[:16]


def _id_to_path(file_id: str) -> Optional[Path]:
    """Resolve file ID back to an absolute path by scanning all PDFs."""
    for pdf_path in DATA_DIR.rglob("*.pdf"):
        rel = str(pdf_path.relative_to(DATA_DIR))
        if _file_id(rel) == file_id:
            return pdf_path
    return None


def _guess_type(filename: str) -> str:
    """Guess drawing type from filename/path segments."""
    upper = filename.upper()
    # Check path segments and filename
    if "/ЭО/" in upper or "\\ЭО\\" in upper or "ЭО" in upper:
        return "ЭО"
    if "/ЭМ/" in upper or "\\ЭМ\\" in upper or "ЭМ" in upper:
        return "ЭМ"
    if "/ЭГ/" in upper or "\\ЭГ\\" in upper or "ЭГ" in upper:
        return "ЭГ"
    if "/ЭС/" in upper or "\\ЭС\\" in upper or "ЭС" in upper:
        return "ЭС"
    return ""


def _detect_section_type(folder_path: str) -> str:
    """Detect electrical section type from folder path.

    Returns "ЭО", "ЭМ", or "ЭГ" based on folder name segments.
    Default: "ЭО" (most common — electrical lighting).
    """
    # Normalize separators
    normalized = folder_path.replace("\\", "/")
    # Check for section markers in path segments
    # Use segment boundaries to avoid false matches (e.g. "ЭОС" should not match "ЭО")
    for segment in normalized.split("/"):
        seg = segment.strip()
        if seg == "ЭО" or seg.startswith("ЭО "):
            return "ЭО"
        if seg == "ЭМ" or seg.startswith("ЭМ "):
            return "ЭМ"
        if seg == "ЭГ" or seg.startswith("ЭГ "):
            return "ЭГ"
    # Fallback: check if any segment contains ЭМ or ЭГ as a substring
    upper = normalized.upper()
    if "/ЭМ/" in upper or "/ЭМ" == upper[-3:]:
        return "ЭМ"
    if "/ЭГ/" in upper or "/ЭГ" == upper[-3:]:
        return "ЭГ"
    return "ЭО"


def _scan_pdfs() -> list[dict]:
    """Recursively scan Data/ for PDF files and return metadata."""
    if not DATA_DIR.exists():
        return []

    results = []
    for pdf_path in sorted(DATA_DIR.rglob("*.pdf")):
        try:
            stat = pdf_path.stat()
        except OSError:
            continue

        rel = str(pdf_path.relative_to(DATA_DIR))
        fid = _file_id(rel)

        # Compute folder group
        rel_posix = rel.replace("\\", "/")
        parts = rel_posix.rsplit("/", 1)
        folder = parts[0] if len(parts) > 1 else ""

        results.append({
            "id": fid,
            "filename": pdf_path.name,
            "path": rel,
            "folder": folder,
            "size_kb": round(stat.st_size / 1024, 1),
            "modified": time.strftime(
                "%Y-%m-%d %H:%M", time.localtime(stat.st_mtime)
            ),
            "type_guess": _guess_type(rel),
        })

    return results


# ---------------------------------------------------------------------------
# 1. GET / — main page
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Render the main page with file list."""
    files = _scan_pdfs()
    return templates.TemplateResponse(
        "index.html",
        {"request": request, "files": files, "total": len(files)},
    )


# ---------------------------------------------------------------------------
# GET /viewer/{file_id} — viewer page for a specific file
# ---------------------------------------------------------------------------

@app.get("/viewer/{file_id}", response_class=HTMLResponse)
async def viewer(request: Request, file_id: str):
    """Render the PDF viewer page for a specific file."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    # Get page count
    try:
        doc = fitz.open(str(pdf_path))
        page_count = len(doc)
        doc.close()
    except Exception:
        page_count = 1

    rel = str(pdf_path.relative_to(DATA_DIR))
    return templates.TemplateResponse(
        "viewer.html",
        {
            "request": request,
            "file_id": file_id,
            "filename": pdf_path.name,
            "filepath": rel,
            "page_count": page_count,
        },
    )


# ---------------------------------------------------------------------------
# 2. GET /api/files — JSON file list
# ---------------------------------------------------------------------------

@app.get("/api/files")
async def api_files():
    """Return JSON list of all PDFs in Data/ directory."""
    files = _scan_pdfs()
    return JSONResponse(content=files)


# ---------------------------------------------------------------------------
# 3. GET /api/file/{id}/render — render PDF page as PNG
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/render")
async def api_render(
    file_id: str,
    page: int = Query(0, ge=0, description="Page index (0-based)"),
    dpi: int = Query(150, ge=72, le=600, description="Render DPI"),
):
    """Render a PDF page as a PNG image using PyMuPDF."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Cannot open PDF: {e}")

    if page >= len(doc):
        doc.close()
        raise HTTPException(
            status_code=400,
            detail=f"Page {page} out of range (0-{len(doc) - 1})",
        )

    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = doc[page].get_pixmap(matrix=mat, alpha=False)

    img_bytes = pix.tobytes("png")
    doc.close()

    return Response(content=img_bytes, media_type="image/png")


# ---------------------------------------------------------------------------
# 4. GET /api/file/{id}/legend — parse legend
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/legend")
async def api_legend(file_id: str):
    """Parse legend from PDF and return structured JSON result."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        result = parse_legend(str(pdf_path))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Legend parse error: {e}")

    # Determine legend_type
    has_numbered = any(item.symbol and item.symbol[0].isdigit() for item in result.items)
    has_graphical = any(not item.symbol for item in result.items)
    if has_numbered and has_graphical:
        legend_type = "mixed"
    elif has_numbered:
        legend_type = "numbered"
    elif has_graphical:
        legend_type = "graphical"
    else:
        legend_type = "numbered"

    # Count raw words for debug info
    raw_words_count = 0
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                words = page.extract_words(x_tolerance=3, y_tolerance=3) or []
                raw_words_count += len(words)
    except Exception:
        pass

    return JSONResponse(content={
        "legend_found": len(result.items) > 0,
        "legend_bbox": {
            "x0": round(result.legend_bbox[0], 1),
            "y0": round(result.legend_bbox[1], 1),
            "x1": round(result.legend_bbox[2], 1),
            "y1": round(result.legend_bbox[3], 1),
        },
        "page": result.page_index,
        "items": [
            {
                "symbol": item.symbol,
                "description": item.description,
                "category": item.category,
                "color": item.color,
                "bbox": {
                    "x0": round(item.bbox[0], 1),
                    "y0": round(item.bbox[1], 1),
                    "x1": round(item.bbox[2], 1),
                    "y1": round(item.bbox[3], 1),
                },
                "image_url": f"/api/file/{file_id}/symbol_image/{i}",
            }
            for i, item in enumerate(result.items)
        ],
        "legend_type": legend_type,
        "raw_words_count": raw_words_count,
        "columns_detected": result.columns_detected,
    })


# ---------------------------------------------------------------------------
# 4b. GET /api/file/{id}/validate_legend — validate legend symbol uniqueness
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/validate_legend")
async def api_validate_legend(file_id: str):
    """Validate that each legend item can be uniquely identified.

    Checks all pairs of legend items for potential conflicts:
    - Same text marker -> ERROR
    - Similar visual templates with no distinguishing features -> PROBLEM
    - Similar visuals but different colors -> OK (color distinguishes)
    - Different text markers -> OK

    Returns per-item status and conflict details.
    """
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        import time as _time
        t0 = _time.time()
        result = await asyncio.to_thread(
            validate_legend_symbols, str(pdf_path)
        )
        elapsed = round(_time.time() - t0, 2)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Validation error: {e}")

    # Serialize items
    items_json = []
    for sv in result.items:
        items_json.append({
            "index": sv.index,
            "symbol": sv.symbol,
            "description": sv.description,
            "color": sv.color,
            "has_text_marker": sv.has_text_marker,
            "has_visual_template": sv.has_visual_template,
            "conflicts": sv.conflicts,
            "status": sv.status,
            "notes": sv.notes,
        })

    # Serialize conflicts
    conflicts_json = []
    for cp in result.conflicts:
        conflicts_json.append({
            "index_a": cp.index_a,
            "index_b": cp.index_b,
            "symbol_a": cp.symbol_a,
            "symbol_b": cp.symbol_b,
            "description_a": cp.description_a,
            "description_b": cp.description_b,
            "conflict_type": cp.conflict_type,
            "visual_similarity": cp.visual_similarity,
            "color_a": cp.color_a,
            "color_b": cp.color_b,
            "distinguishable": cp.distinguishable,
            "resolution": cp.resolution,
        })

    return JSONResponse(content={
        "total": result.total,
        "ok_count": result.ok_count,
        "conflict_count": result.conflict_count,
        "unresolvable_count": result.unresolvable_count,
        "items": items_json,
        "conflicts": conflicts_json,
        "elapsed_s": elapsed,
    })


# ---------------------------------------------------------------------------
# 5. GET /api/file/{id}/render_with_overlay — render with legend highlight
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/render_with_overlay")
async def api_render_with_overlay(
    file_id: str,
    page: int = Query(0, ge=0, description="Page index (0-based)"),
    dpi: int = Query(150, ge=72, le=600, description="Render DPI"),
):
    """Render a PDF page with the legend bbox highlighted as a semi-transparent overlay."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    # First, parse legend to get bbox and page
    try:
        legend_result = parse_legend(str(pdf_path))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Legend parse error: {e}")

    # Open PDF with PyMuPDF
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Cannot open PDF: {e}")

    # Use the legend page if found, otherwise use requested page
    target_page = legend_result.page_index if legend_result.items else page
    if target_page >= len(doc):
        doc.close()
        raise HTTPException(
            status_code=400,
            detail=f"Page {target_page} out of range (0-{len(doc) - 1})",
        )

    zoom = dpi / 72.0

    # Get the page and add legend highlight annotation
    fitz_page = doc[target_page]

    if legend_result.items:
        bbox = legend_result.legend_bbox
        # pdfplumber coordinates: origin at top-left, Y increases downward
        # fitz coordinates: same convention for page rendering
        rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])

        # Draw a semi-transparent rectangle
        shape = fitz_page.new_shape()
        shape.draw_rect(rect)
        shape.finish(
            color=(1, 0, 0),        # red border
            fill=(1, 0.9, 0.8),     # light orange fill
            fill_opacity=0.25,
            width=3,
        )
        shape.commit()

        # Also highlight individual item bboxes
        for item in legend_result.items:
            ib = item.bbox
            item_rect = fitz.Rect(ib[0], ib[1], ib[2], ib[3])
            item_shape = fitz_page.new_shape()
            item_shape.draw_rect(item_rect)
            item_shape.finish(
                color=(0, 0.4, 0.8),     # blue border
                fill=(0.8, 0.9, 1.0),    # light blue fill
                fill_opacity=0.15,
                width=1,
            )
            item_shape.commit()

    mat = fitz.Matrix(zoom, zoom)
    pix = fitz_page.get_pixmap(matrix=mat, alpha=False)
    img_bytes = pix.tobytes("png")
    doc.close()

    return Response(content=img_bytes, media_type="image/png")


# ---------------------------------------------------------------------------
# 6. GET /api/file/{id}/debug_words — debug word extraction
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/debug_words")
async def api_debug_words(
    file_id: str,
    page: int = Query(0, ge=0, description="Page index (0-based)"),
):
    """Return all pdfplumber words with coordinates for debugging."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if page >= len(pdf.pages):
                raise HTTPException(
                    status_code=400,
                    detail=f"Page {page} out of range (0-{len(pdf.pages) - 1})",
                )

            p = pdf.pages[page]
            words = p.extract_words(x_tolerance=3, y_tolerance=3) or []

            return JSONResponse(content={
                "page": page,
                "page_width": round(p.width, 1),
                "page_height": round(p.height, 1),
                "words_count": len(words),
                "words": [
                    {
                        "text": w["text"],
                        "x0": round(w["x0"], 1),
                        "top": round(w["top"], 1),
                        "x1": round(w["x1"], 1),
                        "bottom": round(w["bottom"], 1),
                    }
                    for w in words
                ],
            })
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Word extraction error: {e}")


# ---------------------------------------------------------------------------
# 6b. GET /api/folders — JSON list of folders
# ---------------------------------------------------------------------------

@app.get("/api/folders")
async def api_folders():
    """Return JSON list of folders containing PDFs with IDs and counts."""
    files = _scan_pdfs()
    folder_map: dict[str, dict] = {}
    for f in files:
        folder = f["folder"]
        if folder not in folder_map:
            folder_map[folder] = {"id": _folder_id(folder), "path": folder,
                                   "name": folder.rsplit("/", 1)[-1] if "/" in folder else folder,
                                   "count": 0, "types": set()}
        folder_map[folder]["count"] += 1
        if f["type_guess"]:
            folder_map[folder]["types"].add(f["type_guess"])
    result = []
    for info in sorted(folder_map.values(), key=lambda x: x["path"]):
        info["types"] = sorted(info["types"])
        result.append(info)
    return JSONResponse(content=result)


# ---------------------------------------------------------------------------
# Helpers: equipment aggregation
# ---------------------------------------------------------------------------

def _aggregate_equipment(results: dict[str, list[dict]]) -> list[dict]:
    """Aggregate equipment items from multiple files into VOR table.

    Uses 'work_name' (VOR work description) for aggregation if available,
    falling back to 'name'. Preserves 'equipment_name' (original equipment
    name from the PDF legend) for the 'Доп. информация' column.
    """
    import re
    agg: dict[str, dict] = {}  # normalized_key -> {...}

    for filename, items in results.items():
        drawing_ref = filename.replace(".pdf", "")
        for item in items:
            # Prefer work_name for VOR display; fall back to name
            work_name = item.get("work_name", "").strip()
            raw_name = item.get("name", "").strip()
            display_name = work_name or raw_name
            if not display_name:
                continue
            key = re.sub(r"\s+", " ", display_name).strip().lower()
            total = item.get("total", item.get("count", 0) + item.get("count_ae", 0))
            unit = item.get("unit", "шт")
            if total <= 0:
                continue
            if key not in agg:
                # equipment_name is the original name from PDF legend
                equip_name = item.get("equipment_name", raw_name)
                agg[key] = {
                    "name": display_name, "unit": unit, "total": 0,
                    "per_file": {}, "files": [],
                    "equipment_names": set(),
                }
                if equip_name:
                    agg[key]["equipment_names"].add(equip_name)
            else:
                equip_name = item.get("equipment_name", raw_name)
                if equip_name:
                    agg[key]["equipment_names"].add(equip_name)
            agg[key]["total"] += total
            agg[key]["per_file"][drawing_ref] = agg[key]["per_file"].get(drawing_ref, 0) + total
            if drawing_ref not in agg[key]["files"]:
                agg[key]["files"].append(drawing_ref)

    result = []
    for i, (key, info) in enumerate(sorted(agg.items(), key=lambda x: x[1]["name"]), 1):
        formula_parts = [str(v) for v in info["per_file"].values()]
        formula = "+".join(formula_parts) if len(formula_parts) > 1 else (formula_parts[0] if formula_parts else "")
        # Build extra info from original equipment names
        equip_names = sorted(info.get("equipment_names", set()))
        extra_info = "; ".join(equip_names) if equip_names else ""
        result.append({
            "row": i, "name": info["name"], "unit": info["unit"],
            "total": info["total"], "formula": formula,
            "drawing_refs": ", ".join(info["files"]),
            "extra_info": extra_info,
        })
    return result


def _detect_section_type(pdf_path: str) -> str:
    """Detect the drawing section type from file/folder name.

    Returns one of: "ЭО", "ЭМ", "ЭГ", "ЭОМ", or "unknown".

    Detection order:
      1. Folder name containing section code (e.g. .../ЭМ/..., .../ЭГ/...)
      2. Filename containing section code (e.g. 1Д-24-1-ЭМ.pdf)
    """
    import re
    path_str = str(pdf_path).replace("\\", "/")
    # Check from most specific to least; ЭОМ before ЭО to avoid false match
    for code in ("ЭОМ", "ЭО", "ЭМ", "ЭГ"):
        # Folder match: /ЭМ/ or /ЭМ\  or path ends with /ЭМ
        if re.search(rf"[/\\]{code}(?:[/\\]|$)", path_str):
            return code
        # Filename match: -ЭМ.pdf, -ЭМ_, _ЭМ.pdf, _ЭМ_, " ЭМ "
        basename = path_str.rsplit("/", 1)[-1]
        if re.search(rf"[-_\s]{code}(?:[-_.\s]|$)", basename, re.IGNORECASE):
            return code
        # Also match if basename STARTS with the section code
        if basename.upper().startswith(code):
            return code
    return "unknown"


def _count_equipment_in_pdf(pdf_path: str) -> list[dict]:
    """Run legend extraction + counting methods on a single PDF.

    Supports ЭО, ЭМ, and ЭГ section types:
      - ЭО: full legend + text/visual counting + cables
      - ЭМ: legend (panels/equipment) + cables (heavy cable schedule)
      - ЭГ: cables (grounding conductors) + geometric measurement
      - unknown: same as ЭО (generic)

    Steps:
      1. Detect section type from filename/folder
      2. Parse legend via pdf_legend_parser
      3. Run count_symbols (text) + match_symbols (visual fallback)
      4. Build items from legend + counts
      5. Extract cables (all section types)
      6. For ЭГ: also run measure_cables for geometric lengths
      7. Apply VOR work-name mapping
    """
    import logging
    log = logging.getLogger("web_app._count_equipment_in_pdf")

    section = _detect_section_type(pdf_path)
    log.info("Processing %s (section=%s)", pdf_path, section)

    # Step 1: parse legend
    legend_result = parse_legend(pdf_path)
    has_legend = bool(legend_result.items)

    items: list[dict] = []

    # Step 2-4: legend-based equipment counting (skip if no legend found)
    if has_legend:
        # Step 2: run VISUAL counting first (primary method)
        visual_counts: dict[int, int] = {}  # symbol_index -> count
        try:
            vis_result = match_symbols(pdf_path, legend_result)
            visual_counts = vis_result.counts  # symbol_index -> count
        except Exception as exc:
            log.warning("Visual counting failed for %s: %s", pdf_path, exc)

        # Step 2b: build equipment cluster zones for text filtering
        equip_zones: dict[str, list] | None = None
        try:
            equip_zones = build_equipment_cluster_bboxes(
                pdf_path, legend_result.page_index
            )
        except Exception as exc:
            log.warning("Equipment zone detection failed: %s", exc)

        # Step 3: run text counting as secondary/fallback
        # Pass equipment zones so standalone digit markers are only
        # accepted near coloured equipment clusters.
        text_counts: dict[str, int] = {}
        try:
            text_result = count_symbols(
                pdf_path, legend_result, equipment_zones=equip_zones,
            )
            text_counts = text_result.counts  # symbol -> count
        except Exception as exc:
            log.warning("Text counting failed for %s: %s", pdf_path, exc)

        # Step 4: build enriched items from legend + counts
        for idx, item in enumerate(legend_result.items):
            sym = item.symbol or ""
            name = item.description or ""
            if not name:
                continue

            # Determine count: smart priority between visual and text.
            # Visual is preferred, but if it's suspiciously high compared
            # to text (>3×), the visual result likely contains false positives
            # so we trust text instead.
            vis_count = visual_counts.get(idx, 0)
            txt_count = text_counts.get(sym, 0) if sym else 0
            count = 0
            if vis_count > 0 and txt_count > 0 and vis_count > txt_count * 3:
                count = txt_count  # visual likely has false positives
            elif vis_count > 0:
                count = vis_count
            elif txt_count > 0:
                count = txt_count

            if count <= 0:
                continue

            items.append({
                "symbol": sym,
                "name": name,
                "count": count,
                "count_ae": 0,
                "total": count,
            })

    # Step 5: extract cables and add to items (ALL section types)
    try:
        cable_result = extract_cables(pdf_path, legend_result)
        for entry in cable_result.cable_schedule:
            group = entry.get("group", "")
            panel = entry.get("panel", "")
            cable_types = entry.get("cable_types", [])
            cross_sections = entry.get("cross_sections", [])
            run_count = entry.get("run_count", 0) or 0
            total_length_m = entry.get("total_length_m", 0) or 0

            # Use cable_type if available, otherwise cross_section
            type_label = (cable_types[0] if cable_types
                          else cross_sections[0] if cross_sections
                          else "")
            if not type_label and not group:
                continue

            if run_count > 0:
                cable_name = f"Кабель {type_label}" if type_label else "Кабель"
                if group:
                    cable_name += f" ({panel}-{group})" if panel else f" ({group})"
                items.append({
                    "symbol": "",
                    "name": cable_name,
                    "count": run_count,
                    "count_ae": 0,
                    "total": run_count,
                    "unit": "шт",
                })
            if total_length_m > 0:
                cable_name_m = (f"Кабель {type_label} (прокладка)"
                                if type_label else "Кабель (прокладка)")
                if group:
                    cable_name_m += f" ({panel}-{group})" if panel else f" ({group})"
                items.append({
                    "symbol": "",
                    "name": cable_name_m,
                    "count": 0,
                    "count_ae": 0,
                    "total": round(total_length_m, 1),
                    "unit": "м",
                })
    except Exception as exc:
        log.warning("Cable extraction failed for %s: %s", pdf_path, exc)

    # Step 6: for ЭГ section — use geometric measurement for cable lengths
    # ЭГ drawings often have grounding conductors measured by line geometry
    # rather than cable schedule tables
    if section == "ЭГ":
        try:
            geo_result = measure_cables(pdf_path, legend_result)
            # Add red-line measurements (typically grounding conductor runs)
            if geo_result.total_red_length_m > 0:
                items.append({
                    "symbol": "",
                    "name": "Проводник заземления (горизонтальный)",
                    "count": 0,
                    "count_ae": 0,
                    "total": round(geo_result.total_red_length_m, 1),
                    "unit": "м",
                })
            # Add blue-line measurements (typically equipotential bonding)
            if geo_result.total_blue_length_m > 0:
                items.append({
                    "symbol": "",
                    "name": "Проводник уравнивания потенциалов",
                    "count": 0,
                    "count_ae": 0,
                    "total": round(geo_result.total_blue_length_m, 1),
                    "unit": "м",
                })
            # Add individual route details by linewidth if available
            for lw_key, lw_info in geo_result.by_linewidth.items():
                length_m = lw_info.get("length_m", 0)
                segments = lw_info.get("segments", 0)
                if length_m > 0 and segments > 0:
                    log.info("ЭГ geometry: linewidth=%s  length=%.1f m  segments=%d",
                             lw_key, length_m, segments)
        except Exception as exc:
            log.warning("Geometric measurement failed for %s: %s", pdf_path, exc)

    # Step 7: apply VOR work-name mapping
    items = vor_map_items(items)

    return items


# ---------------------------------------------------------------------------
# 6c. GET /api/folder/{folder_id}/process — SSE batch processing
# ---------------------------------------------------------------------------

@app.get("/api/folder/{folder_id}/process")
async def api_folder_process(folder_id: str):
    """Process all PDFs in a folder via SSE stream with progress."""
    rel_folder = _id_to_folder(folder_id)
    if rel_folder is None:
        raise HTTPException(status_code=404, detail="Folder not found")
    pdf_files = _folder_files(rel_folder)
    if not pdf_files:
        raise HTTPException(status_code=404, detail="No PDFs in folder")

    async def event_stream():
        all_results: dict[str, list[dict]] = {}
        errors: list[dict] = []
        total = len(pdf_files)
        yield f"event: start\ndata: {json_mod.dumps({'total': total, 'folder': rel_folder}, ensure_ascii=False)}\n\n"

        for i, pdf_path in enumerate(pdf_files):
            filename = pdf_path.name
            yield f"event: progress\ndata: {json_mod.dumps({'current': i + 1, 'total': total, 'filename': filename}, ensure_ascii=False)}\n\n"
            try:
                file_result = await asyncio.to_thread(
                    _count_equipment_in_pdf, str(pdf_path)
                )
                all_results[filename] = file_result
                yield f"event: file_done\ndata: {json_mod.dumps({'filename': filename, 'items': len(file_result), 'status': 'ok'}, ensure_ascii=False)}\n\n"
            except Exception as e:
                errors.append({"filename": filename, "error": str(e)})
                yield f"event: file_done\ndata: {json_mod.dumps({'filename': filename, 'items': 0, 'status': 'error', 'error': str(e)}, ensure_ascii=False)}\n\n"

        aggregated = _aggregate_equipment(all_results)
        yield f"event: done\ndata: {json_mod.dumps({'aggregated': aggregated, 'files_processed': len(all_results), 'errors': errors, 'total_files': total}, ensure_ascii=False)}\n\n"

    return StreamingResponse(event_stream(), media_type="text/event-stream",
                            headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ---------------------------------------------------------------------------
# 6d. GET /api/folder/{folder_id}/export_xlsx — Excel VOR export
# ---------------------------------------------------------------------------

@app.get("/api/folder/{folder_id}/export_xlsx")
async def api_folder_export_xlsx(folder_id: str):
    """Generate and download VOR Excel file for a folder."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    rel_folder = _id_to_folder(folder_id)
    if rel_folder is None:
        raise HTTPException(status_code=404, detail="Folder not found")
    pdf_files = _folder_files(rel_folder)
    if not pdf_files:
        raise HTTPException(status_code=404, detail="No PDFs in folder")

    # Process all files using counting methods for actual quantities
    all_results: dict[str, list[dict]] = {}
    for pdf_path in pdf_files:
        try:
            file_result = await asyncio.to_thread(
                _count_equipment_in_pdf, str(pdf_path)
            )
            all_results[pdf_path.name] = file_result
        except Exception:
            continue

    aggregated = _aggregate_equipment(all_results)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ВОР"

    header_font = Font(bold=True, size=10)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    thin_border = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

    headers = ["№ п/п", "Наименование вида работ", "Ед. изм.", "Объем работ",
               "Формула расчета", "Ссылка на чертежи", "Доп. информация"]
    col_widths = [7, 72, 9, 12, 20, 26, 27]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

    for row_data in aggregated:
        ri = row_data["row"] + 1
        vals = [row_data["row"], row_data["name"], row_data["unit"],
                row_data["total"], row_data["formula"], row_data["drawing_refs"],
                row_data.get("extra_info", "")]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border = thin_border
            if ci in (1, 3, 4):
                cell.alignment = Alignment(horizontal="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    folder_name = rel_folder.rsplit("/", 1)[-1] if "/" in rel_folder else rel_folder
    safe_name = "".join(c if c.isalnum() or c in "._- " else "_" for c in folder_name)

    return Response(content=buf.getvalue(),
                   media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   headers={"Content-Disposition": f'attachment; filename="VOR_{safe_name}.xlsx"'})


# ---------------------------------------------------------------------------
# Helpers: color analysis
# ---------------------------------------------------------------------------

# Known color labels for electrical drawings
_KNOWN_COLORS: dict[tuple, str] = {
    (0, 0, 0): "Чёрный / Рамка",
    (1, 0, 0): "Красный / Аварийное",
    (0, 0, 1): "Синий / Рабочее",
    (0, 1, 0): "Зелёный",
    (1, 1, 0): "Жёлтый",
    (1, 0, 1): "Пурпурный",
    (0, 1, 1): "Голубой",
}


def _normalize_color(c) -> tuple | None:
    """Normalize a pdfplumber color value to a tuple of floats, or None."""
    if c is None:
        return None
    if isinstance(c, (int, float)):
        # Grayscale
        v = float(c)
        return (v, v, v)
    if isinstance(c, (tuple, list)):
        if len(c) == 3:
            return tuple(round(float(x), 4) for x in c)
        if len(c) == 4:
            # CMYK → RGB approximation
            cc, m, y, k = [float(x) for x in c]
            r = (1 - cc) * (1 - k)
            g = (1 - m) * (1 - k)
            b = (1 - y) * (1 - k)
            return (round(r, 4), round(g, 4), round(b, 4))
        if len(c) == 1:
            v = float(c[0])
            return (v, v, v)
    return None


def _color_to_hex(rgb: tuple) -> str:
    """Convert (r,g,b) floats 0-1 to hex string like 'FF0000'."""
    r = max(0, min(255, int(round(rgb[0] * 255))))
    g = max(0, min(255, int(round(rgb[1] * 255))))
    b = max(0, min(255, int(round(rgb[2] * 255))))
    return f"{r:02X}{g:02X}{b:02X}"


def _label_color(rgb: tuple) -> str:
    """Auto-label a color if it matches known colors (with tolerance)."""
    # Exact match first
    rounded = tuple(round(x, 2) for x in rgb)
    for known, label in _KNOWN_COLORS.items():
        if all(abs(a - b) < 0.05 for a, b in zip(rounded, known)):
            return label
    # Gray detection
    if abs(rgb[0] - rgb[1]) < 0.05 and abs(rgb[1] - rgb[2]) < 0.05:
        level = rgb[0]
        if level < 0.15:
            return "Чёрный / Рамка"
        if level > 0.85:
            return "Белый / Фон"
        return f"Серый ({int(level * 100)}%)"
    return ""


# ---------------------------------------------------------------------------
# 7. GET /api/file/{id}/colors — color palette analysis
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/colors")
async def api_colors(
    file_id: str,
    page: int = Query(0, ge=0, description="Page index (0-based)"),
):
    """Analyze PDF page color palette using pdfplumber."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if page >= len(pdf.pages):
                raise HTTPException(
                    status_code=400,
                    detail=f"Page {page} out of range (0-{len(pdf.pages) - 1})",
                )

            p = pdf.pages[page]

            # Collect color stats: rgb_tuple -> {lines, rects, chars}
            color_stats: dict[tuple, dict] = {}

            def _inc(rgb, kind):
                if rgb is None:
                    return
                if rgb not in color_stats:
                    color_stats[rgb] = {"lines": 0, "rects": 0, "chars": 0}
                color_stats[rgb][kind] += 1

            # Lines
            for line in (p.lines or []):
                rgb = _normalize_color(line.get("stroking_color"))
                _inc(rgb, "lines")

            # Rects
            for rect in (p.rects or []):
                rgb_s = _normalize_color(rect.get("stroking_color"))
                rgb_f = _normalize_color(rect.get("non_stroking_color"))
                _inc(rgb_s, "rects")
                if rgb_f and rgb_f != rgb_s:
                    _inc(rgb_f, "rects")

            # Chars
            for ch in (p.chars or []):
                rgb = _normalize_color(ch.get("non_stroking_color"))
                _inc(rgb, "chars")

            # Build response
            colors = []
            for rgb, stats in sorted(
                color_stats.items(),
                key=lambda kv: kv[1]["lines"] + kv[1]["rects"] + kv[1]["chars"],
                reverse=True,
            ):
                total = stats["lines"] + stats["rects"] + stats["chars"]
                hex_val = _color_to_hex(rgb)
                label = _label_color(rgb)
                colors.append({
                    "rgb": list(rgb),
                    "hex": hex_val,
                    "label": label,
                    "lines": stats["lines"],
                    "rects": stats["rects"],
                    "chars": stats["chars"],
                    "total": total,
                })

            return JSONResponse(content={
                "page": page,
                "colors": colors,
            })
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Color analysis error: {e}")


# ---------------------------------------------------------------------------
# 8. GET /api/file/{id}/render_filtered — render with color filter
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/render_filtered")
async def api_render_filtered(
    file_id: str,
    page: int = Query(0, ge=0, description="Page index (0-based)"),
    dpi: int = Query(150, ge=72, le=600, description="Render DPI"),
    show: str = Query("", description="Comma-separated hex colors to show (e.g. FF0000,0000FF)"),
    hide: str = Query("", description="Comma-separated hex colors to hide"),
):
    """Render PDF page showing only elements matching specified colors.

    Uses PyMuPDF to render the full page, then composites with pdfplumber
    color data to show/hide elements by color.
    """
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    # Parse color filters
    show_set: set[str] = set()
    hide_set: set[str] = set()
    if show.strip():
        show_set = {c.strip().upper() for c in show.split(",") if c.strip()}
    if hide.strip():
        hide_set = {c.strip().upper() for c in hide.split(",") if c.strip()}

    if not show_set and not hide_set:
        # No filter — just render normally
        return await api_render(file_id, page, dpi)

    try:
        # Step 1: Get element data from pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            if page >= len(pdf.pages):
                raise HTTPException(
                    status_code=400,
                    detail=f"Page {page} out of range (0-{len(pdf.pages) - 1})",
                )
            p = pdf.pages[page]
            pg_w, pg_h = float(p.width), float(p.height)

            # Collect elements grouped by visibility
            visible_lines = []
            visible_rects = []
            visible_chars = []

            def _is_visible(rgb_tuple) -> bool:
                if rgb_tuple is None:
                    return not bool(show_set)
                hex_val = _color_to_hex(rgb_tuple)
                if show_set:
                    return hex_val in show_set
                if hide_set:
                    return hex_val not in hide_set
                return True

            for line in (p.lines or []):
                rgb = _normalize_color(line.get("stroking_color"))
                if _is_visible(rgb):
                    visible_lines.append(line)

            for rect in (p.rects or []):
                rgb_s = _normalize_color(rect.get("stroking_color"))
                rgb_f = _normalize_color(rect.get("non_stroking_color"))
                if _is_visible(rgb_s) or _is_visible(rgb_f):
                    visible_rects.append(rect)

            for ch in (p.chars or []):
                rgb = _normalize_color(ch.get("non_stroking_color"))
                if _is_visible(rgb):
                    visible_chars.append(ch)

        # Step 2: Render using PyMuPDF shapes on a blank page
        doc = fitz.open(str(pdf_path))
        if page >= len(doc):
            doc.close()
            raise HTTPException(status_code=400, detail=f"Page out of range")

        src_page = doc[page]
        # Create a new blank document with same page size
        new_doc = fitz.open()
        new_page = new_doc.new_page(width=pg_w, height=pg_h)

        # Draw visible lines
        if visible_lines:
            shape = new_page.new_shape()
            for line in visible_lines:
                x0, y0 = float(line["x0"]), float(line["top"])
                x1, y1 = float(line["x1"]), float(line["bottom"])
                lw = float(line.get("linewidth", 1) or 1)
                rgb = _normalize_color(line.get("stroking_color"))
                color = rgb if rgb else (0, 0, 0)
                shape.draw_line(fitz.Point(x0, y0), fitz.Point(x1, y1))
                shape.finish(color=color, width=max(0.3, lw))
            shape.commit()

        # Draw visible rects
        if visible_rects:
            shape = new_page.new_shape()
            for rect in visible_rects:
                x0, y0 = float(rect["x0"]), float(rect["top"])
                x1, y1 = float(rect["x1"]), float(rect["bottom"])
                lw = float(rect.get("linewidth", 0.5) or 0.5)
                rgb_s = _normalize_color(rect.get("stroking_color"))
                rgb_f = _normalize_color(rect.get("non_stroking_color"))
                s_color = rgb_s if rgb_s else (0, 0, 0)
                f_color = rgb_f if rgb_f else None
                shape.draw_rect(fitz.Rect(x0, y0, x1, y1))
                shape.finish(
                    color=s_color,
                    fill=f_color,
                    width=max(0.1, lw),
                )
            shape.commit()

        # Draw visible chars (group by position for efficiency)
        for ch in visible_chars:
            rgb = _normalize_color(ch.get("non_stroking_color"))
            color = rgb if rgb else (0, 0, 0)
            x0 = float(ch["x0"])
            y0 = float(ch["top"])
            y1 = float(ch["bottom"])
            font_size = y1 - y0
            text = ch.get("text", "")
            if not text or not text.strip():
                continue
            try:
                new_page.insert_text(
                    fitz.Point(x0, y1 - font_size * 0.15),
                    text,
                    fontsize=max(1, font_size * 0.85),
                    color=color,
                )
            except Exception:
                pass

        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = new_page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")

        new_doc.close()
        doc.close()

        return Response(content=img_bytes, media_type="image/png")

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Filtered render error: {e}")


# ---------------------------------------------------------------------------
# 9. Find equipment positions (for interactive highlight)
# ---------------------------------------------------------------------------

@app.get("/api/file/{file_id}/find/{row_index}")
async def api_find_positions(file_id: str, row_index: int):
    """Find all instances of a legend item on the drawing.

    Returns positions from Method A (text markers) and optionally
    Method D (visual template matching) if available.
    """
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Legend parse error: {e}")

    if row_index < 0 or row_index >= len(legend.items):
        raise HTTPException(
            status_code=400,
            detail=f"Row index {row_index} out of range (0-{len(legend.items) - 1})",
        )

    item = legend.items[row_index]
    symbol = item.symbol or ""
    description = item.description or ""

    positions: list[dict] = []
    methods_used: list[str] = []
    excluded_zones: list[dict] = []

    # Method D: visual template matching — PRIMARY
    try:
        t0 = time.time()
        vis_result = await asyncio.to_thread(
            match_symbols, str(pdf_path), legend
        )
        elapsed_vis = round(time.time() - t0, 2)
        methods_used.append("visual")

        for m in vis_result.matches:
            if m.symbol_index == row_index:
                positions.append({
                    "x": round(m.x, 1),
                    "y": round(m.y, 1),
                    "width": 20,
                    "height": 20,
                    "confidence": round(m.confidence, 3),
                    "method": "visual",
                })
    except Exception:
        pass

    # Method A: text markers — SECONDARY (add non-duplicate positions)
    if symbol:
        try:
            # Build equipment zones for spatial filtering
            _eq_zones = None
            try:
                _eq_zones = await asyncio.to_thread(
                    build_equipment_cluster_bboxes,
                    str(pdf_path), legend.page_index,
                )
            except Exception:
                pass

            t0 = time.time()
            text_result = await asyncio.to_thread(
                count_symbols, str(pdf_path), legend, _eq_zones,
            )
            elapsed_text = round(time.time() - t0, 2)
            methods_used.append("text")

            for p in text_result.positions:
                if p.symbol == symbol:
                    # Skip if already found by visual (within 15pt)
                    duplicate = any(
                        abs(existing["x"] - p.x) < 15 and abs(existing["y"] - p.y) < 15
                        for existing in positions
                    )
                    if not duplicate:
                        positions.append({
                            "x": round(p.x, 1),
                            "y": round(p.y, 1),
                            "width": 12,
                            "height": 12,
                            "confidence": 1.0,
                            "method": "text",
                        })

            # Build exclusion zones
            excluded_zones = [
                {
                    "x0": round(z[1][0], 1), "y0": round(z[1][1], 1),
                    "x1": round(z[1][2], 1), "y1": round(z[1][3], 1),
                    "reason": z[0],
                }
                for z in text_result.exclusion_zones
            ]
        except Exception:
            pass

    needs_visual = len(positions) == 0

    return JSONResponse(content={
        "symbol": symbol,
        "description": description,
        "category": item.category or "",
        "row_index": row_index,
        "positions": positions,
        "count": len(positions),
        "excluded_zones": excluded_zones if symbol else [],
        "methods_used": methods_used,
        "needs_visual": needs_visual,
    })


# ---------------------------------------------------------------------------
# 9b. Cable highlight analysis (T107)
# ---------------------------------------------------------------------------

def _extract_cable_segments(pdf_path: str, legend_result, pages):
    """Extract cable line segments grouped by color/linewidth with connected routes.

    Returns a dict with segments, routes, annotations, and scale info.
    """
    from pdf_count_geometry import (
        measure_cables as geo_measure,
        _classify_line_color, _segment_length,
        _build_routes, _build_exclusion_zones, _detect_scale,
        MIN_SEGMENT_LENGTH_PT, ENDPOINT_TOLERANCE,
    )
    from pdf_count_cables import extract_cables as cable_extract

    legend_bbox = None
    legend_page = -1
    if legend_result and legend_result.items:
        legend_bbox = legend_result.legend_bbox
        legend_page = legend_result.page_index

    segments_out = []  # all cable line segments for overlay
    routes_out = []    # connected route polylines
    annotations_out = []  # cable run annotations (text-based)
    scale_info = None

    with pdfplumber.open(pdf_path) as pdf:
        scan_pages = pages if pages is not None else list(range(len(pdf.pages)))

        for page_idx in scan_pages:
            if page_idx >= len(pdf.pages):
                continue
            page = pdf.pages[page_idx]
            pdf_lines = page.lines or []
            words = page.extract_words(x_tolerance=3, y_tolerance=3) or []

            if not pdf_lines:
                continue

            # Detect scale
            page_scale = _detect_scale(page, words, pdf_lines)
            if scale_info is None or (
                scale_info.get("confidence") != "high"
                and page_scale.confidence == "high"
            ):
                scale_info = {
                    "mm_per_pt": round(page_scale.mm_per_pt, 4),
                    "source": page_scale.source,
                    "confidence": page_scale.confidence,
                }

            # Build exclusion zones
            lb = legend_bbox if page_idx == legend_page else None
            zones = _build_exclusion_zones(page, pdf_lines, lb)

            # Classify colored lines
            red_segs = []
            blue_segs = []

            for ln in pdf_lines:
                # Check exclusion
                mx = (ln["x0"] + ln["x1"]) / 2
                my = (ln["top"] + ln["bottom"]) / 2
                excluded = False
                for _, zb in zones:
                    if zb[0] <= mx <= zb[2] and zb[1] <= my <= zb[3]:
                        excluded = True
                        break
                if excluded:
                    continue

                seg_len = _segment_length(ln)
                if seg_len < MIN_SEGMENT_LENGTH_PT:
                    continue

                color = _classify_line_color(ln.get("stroking_color"))
                if color == "other":
                    continue

                lw = round(ln.get("linewidth", 0), 3)
                seg = {
                    "x0": round(ln["x0"], 1), "y0": round(ln["top"], 1),
                    "x1": round(ln["x1"], 1), "y1": round(ln["bottom"], 1),
                    "color": color, "lw": lw, "page": page_idx,
                }
                segments_out.append(seg)

                if color == "red":
                    red_segs.append(ln)
                else:
                    blue_segs.append(ln)

            # Build connected routes
            for color, segs in [("red", red_segs), ("blue", blue_segs)]:
                if not segs:
                    continue
                route_groups = _build_routes(segs, ENDPOINT_TOLERANCE)
                mm_per_pt = scale_info["mm_per_pt"] if scale_info else 35.0

                for ri, route_segs in enumerate(route_groups):
                    total_pt = sum(_segment_length(s) for s in route_segs)
                    total_m = total_pt * mm_per_pt / 1000.0

                    # Compute bounding box of route
                    all_x = []
                    all_y = []
                    for s in route_segs:
                        all_x.extend([s["x0"], s["x1"]])
                        all_y.extend([s["top"], s["bottom"]])

                    routes_out.append({
                        "id": len(routes_out),
                        "color": color,
                        "segment_count": len(route_segs),
                        "length_pt": round(total_pt, 1),
                        "length_m": round(total_m, 1),
                        "bbox": {
                            "x0": round(min(all_x), 1),
                            "y0": round(min(all_y), 1),
                            "x1": round(max(all_x), 1),
                            "y1": round(max(all_y), 1),
                        },
                        "page": page_idx,
                    })

    # Get cable annotations (text-based)
    try:
        cable_result = cable_extract(pdf_path, legend_result, pages)
        for r in cable_result.runs:
            annotations_out.append({
                "panel": r.panel, "group": r.group,
                "group_full": r.group_full,
                "cross_section": r.cross_section,
                "length_m": r.length_m,
                "cable_type": r.cable_type,
                "x": r.position[0], "y": r.position[1],
                "color": r.color,
                "page": r.page_index,
            })
    except Exception:
        pass

    return {
        "segments": segments_out,
        "routes": routes_out,
        "annotations": annotations_out,
        "scale": scale_info,
        "total_segments": len(segments_out),
        "total_routes": len(routes_out),
        "red_segments": sum(1 for s in segments_out if s["color"] == "red"),
        "blue_segments": sum(1 for s in segments_out if s["color"] == "blue"),
    }


@app.get("/api/file/{file_id}/cables")
async def api_cables(
    file_id: str,
    all_pages: bool = Query(False, description="Scan all pages"),
):
    """Cable highlight data: line segments, connected routes, annotations.

    Three filter modes supported by the frontend:
      1. By type: red=emergency, blue=working
      2. By group: ЩО1-Гр.7
      3. By route: connected polyline ID
    """
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        t0 = time.time()
        legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
        pages = None if all_pages else (
            [legend.page_index] if legend.items else None
        )
        data = await asyncio.to_thread(
            _extract_cable_segments, str(pdf_path), legend, pages
        )
        elapsed = round(time.time() - t0, 2)
        data["elapsed_s"] = elapsed
        return JSONResponse(content=data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Cable analysis error: {e}")


# ---------------------------------------------------------------------------
# 10. Counting API endpoints
# ---------------------------------------------------------------------------

# Cache for legend results (keyed by file_id)
_legend_cache: dict[str, LegendResult] = {}


def _get_legend(pdf_path: Path, file_id: str) -> LegendResult:
    """Get cached or freshly parsed legend result."""
    if file_id not in _legend_cache:
        _legend_cache[file_id] = parse_legend(str(pdf_path))
    return _legend_cache[file_id]


# Cache for symbol images extracted from legend (keyed by file_id)
import numpy as np
import cv2

_symbol_image_cache: dict[str, list] = {}


def _get_symbol_images(pdf_path: Path, file_id: str, legend_result: LegendResult) -> list:
    """Get cached or freshly extracted symbol images from legend."""
    if file_id not in _symbol_image_cache:
        _symbol_image_cache[file_id] = _extract_symbol_images(
            str(pdf_path), legend_result
        )
    return _symbol_image_cache[file_id]


def _make_symbol_png_transparent(img: np.ndarray) -> bytes:
    """Convert BGR symbol image to PNG with transparent background."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # White/near-white pixels become transparent
    _, mask = cv2.threshold(gray, 235, 255, cv2.THRESH_BINARY)
    # Convert BGR to BGRA (add alpha channel)
    bgra = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)
    # Set alpha: white pixels -> 0 (transparent), rest -> 255 (opaque)
    bgra[:, :, 3] = 255 - mask
    _, png_buf = cv2.imencode(".png", bgra)
    return png_buf.tobytes()


@app.get("/api/file/{file_id}/symbol_image/{row_index}")
async def api_symbol_image(file_id: str, row_index: int):
    """Render a single legend symbol cell as PNG with transparent background."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
    images = await asyncio.to_thread(_get_symbol_images, pdf_path, file_id, legend)

    # Find image for the requested row index
    for idx, _item, img in images:
        if idx == row_index and img is not None:
            png_bytes = _make_symbol_png_transparent(img)
            return Response(
                content=png_bytes,
                media_type="image/png",
                headers={"Cache-Control": "max-age=3600"},
            )

    # No image for this index — return 204
    return Response(status_code=204)


@app.get("/api/file/{file_id}/count/text")
async def api_count_text(file_id: str):
    """Run Method A: count text markers on drawing."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        t0 = time.time()
        legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
        result = await asyncio.to_thread(count_symbols, str(pdf_path), legend)
        elapsed = round(time.time() - t0, 2)

        positions_by_sym: dict[str, list[dict]] = {}
        for p in result.positions:
            if p.symbol not in positions_by_sym:
                positions_by_sym[p.symbol] = []
            positions_by_sym[p.symbol].append({
                "x": round(p.x, 1), "y": round(p.y, 1),
                "merged": p.merged,
            })

        return JSONResponse(content={
            "method": "text",
            "page_index": result.page_index,
            "counts": result.counts,
            "positions": positions_by_sym,
            "total_found": sum(result.counts.values()),
            "symbols_searched": result.symbols_searched,
            "elapsed_s": elapsed,
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Text count error: {e}")


@app.get("/api/file/{file_id}/count/cables")
async def api_count_cables(
    file_id: str,
    all_pages: bool = Query(False, description="Scan all pages"),
):
    """Run Method B: extract cable annotations."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        t0 = time.time()
        legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
        pages = None if all_pages else (
            [legend.page_index] if legend.items else None
        )
        result = await asyncio.to_thread(
            extract_cables, str(pdf_path), legend, pages
        )
        elapsed = round(time.time() - t0, 2)

        runs_json = []
        for r in result.runs:
            runs_json.append({
                "panel": r.panel, "group": r.group,
                "group_full": r.group_full,
                "cross_section": r.cross_section,
                "length_m": r.length_m, "cable_type": r.cable_type,
                "position": {"x": r.position[0], "y": r.position[1]},
                "color": r.color, "page_index": r.page_index,
                "is_reversed": r.is_reversed,
            })

        schedule_json = []
        for entry in result.cable_schedule:
            schedule_json.append({
                "group": entry["group"],
                "panel": entry["panel"],
                "cross_sections": entry["cross_sections"],
                "cable_types": entry["cable_types"],
                "run_count": entry["run_count"],
                "total_length_m": entry["total_length_m"],
                "colors": entry.get("colors", []),
            })

        return JSONResponse(content={
            "method": "cables",
            "total_runs": result.total_runs,
            "runs": runs_json,
            "panels": {k: len(v) for k, v in result.panels.items()},
            "cable_schedule": schedule_json,
            "pages_scanned": result.pages_scanned,
            "elapsed_s": elapsed,
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Cable extraction error: {e}")


@app.get("/api/file/{file_id}/count/geometry")
async def api_count_geometry(
    file_id: str,
    all_pages: bool = Query(False, description="Scan all pages"),
):
    """Run Method C: measure cable routes by geometry."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        t0 = time.time()
        legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
        pages = None if all_pages else (
            [legend.page_index] if legend.items else None
        )
        result = await asyncio.to_thread(
            measure_cables, str(pdf_path), legend, pages
        )
        elapsed = round(time.time() - t0, 2)

        routes_json = []
        for r in result.routes:
            routes_json.append({
                "color": r.color, "linewidth": r.linewidth,
                "total_length_pt": round(r.total_length_pt, 1),
                "total_length_m": round(r.total_length_m, 1),
                "segment_count": r.segment_count,
                "route_count": r.route_count,
                "page_index": r.page_index,
            })

        scale_info = None
        if result.scale:
            scale_info = {
                "mm_per_pt": round(result.scale.mm_per_pt, 4),
                "source": result.scale.source,
                "confidence": result.scale.confidence,
            }

        return JSONResponse(content={
            "method": "geometry",
            "routes": routes_json,
            "scale": scale_info,
            "pages_scanned": result.pages_scanned,
            "elapsed_s": elapsed,
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Geometry measurement error: {e}")


@app.get("/api/file/{file_id}/count/visual")
async def api_count_visual(
    file_id: str,
    page: Optional[int] = Query(None, description="Page to scan (0-based)"),
    threshold: float = Query(0.75, ge=0.3, le=1.0),
):
    """Run Method D: visual symbol template matching."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    try:
        t0 = time.time()
        legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
        result = await asyncio.to_thread(
            match_symbols, str(pdf_path), legend, page, threshold
        )
        elapsed = round(time.time() - t0, 2)

        matches_json = []
        for m in result.matches:
            matches_json.append({
                "symbol_index": m.symbol_index,
                "description": m.description,
                "x": m.x, "y": m.y,
                "confidence": m.confidence,
                "scale": m.scale, "rotation": m.rotation,
                "color": m.color, "page_index": m.page_index,
            })

        return JSONResponse(content={
            "method": "visual",
            "counts": result.counts,
            "descriptions": result.descriptions,
            "matches": matches_json,
            "symbols_extracted": result.symbols_extracted,
            "page_index": result.page_index,
            "threshold": result.threshold,
            "elapsed_s": elapsed,
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Visual matching error: {e}")


@app.get("/api/file/{file_id}/count/all")
async def api_count_all(file_id: str):
    """Run ALL counting methods and return combined results."""
    pdf_path = _id_to_path(file_id)
    if pdf_path is None:
        raise HTTPException(status_code=404, detail="File not found")

    t0 = time.time()
    results = {}
    errors = {}

    # Parse legend once
    try:
        legend = await asyncio.to_thread(_get_legend, pdf_path, file_id)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Legend parse error: {e}")

    legend_page = legend.page_index if legend.items else 0

    # Method A: Text markers
    try:
        t1 = time.time()
        text_result = await asyncio.to_thread(
            count_symbols, str(pdf_path), legend
        )
        positions_by_sym: dict[str, list[dict]] = {}
        for p in text_result.positions:
            if p.symbol not in positions_by_sym:
                positions_by_sym[p.symbol] = []
            positions_by_sym[p.symbol].append({
                "x": round(p.x, 1), "y": round(p.y, 1),
            })
        results["text"] = {
            "counts": text_result.counts,
            "positions": positions_by_sym,
            "total": sum(text_result.counts.values()),
            "elapsed_s": round(time.time() - t1, 2),
        }
    except Exception as e:
        errors["text"] = str(e)

    # Method B: Cable annotations
    try:
        t1 = time.time()
        cable_result = await asyncio.to_thread(
            extract_cables, str(pdf_path), legend, [legend_page]
        )
        cable_runs = []
        for r in cable_result.runs:
            cable_runs.append({
                "panel": r.panel, "group_full": r.group_full,
                "cross_section": r.cross_section,
                "length_m": r.length_m, "cable_type": r.cable_type,
                "color": r.color,
                "position": {"x": r.position[0], "y": r.position[1]},
            })
        results["cables"] = {
            "total_runs": cable_result.total_runs,
            "runs": cable_runs,
            "panels": {k: len(v) for k, v in cable_result.panels.items()},
            "schedule": cable_result.cable_schedule,
            "elapsed_s": round(time.time() - t1, 2),
        }
    except Exception as e:
        errors["cables"] = str(e)

    # Method C: Geometry
    try:
        t1 = time.time()
        geo_result = await asyncio.to_thread(
            measure_cables, str(pdf_path), legend, [legend_page]
        )
        routes_json = []
        for r in geo_result.routes:
            routes_json.append({
                "color": r.color,
                "total_length_m": round(r.total_length_m, 1),
                "segment_count": r.segment_count,
            })
        scale_info = None
        if geo_result.scale:
            scale_info = {
                "mm_per_pt": round(geo_result.scale.mm_per_pt, 4),
                "source": geo_result.scale.source,
            }
        results["geometry"] = {
            "routes": routes_json,
            "scale": scale_info,
            "elapsed_s": round(time.time() - t1, 2),
        }
    except Exception as e:
        errors["geometry"] = str(e)

    # Method D: Visual matching
    try:
        t1 = time.time()
        vis_result = await asyncio.to_thread(
            match_symbols, str(pdf_path), legend
        )
        vis_matches = []
        for m in vis_result.matches:
            vis_matches.append({
                "symbol_index": m.symbol_index,
                "description": m.description,
                "x": m.x, "y": m.y,
                "confidence": m.confidence, "color": m.color,
            })
        results["visual"] = {
            "counts": vis_result.counts,
            "descriptions": vis_result.descriptions,
            "matches": vis_matches,
            "symbols_extracted": vis_result.symbols_extracted,
            "elapsed_s": round(time.time() - t1, 2),
        }
    except Exception as e:
        errors["visual"] = str(e)

    total_elapsed = round(time.time() - t0, 2)

    return JSONResponse(content={
        "results": results,
        "errors": errors,
        "legend_page": legend_page,
        "legend_items": len(legend.items),
        "elapsed_s": total_elapsed,
    })


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_app:app", host="0.0.0.0", port=8050, reload=True)
