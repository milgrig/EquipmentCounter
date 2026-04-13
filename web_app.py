"""
web_app.py — FastAPI web application for PDF legend analysis.

Provides endpoints for browsing PDFs, rendering pages, parsing legends,
and debugging word extraction.

Usage:
    uvicorn web_app:app --host 0.0.0.0 --port 8000 --reload
    # or
    python web_app.py
"""

from __future__ import annotations

import hashlib
import io
import os
import time
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
import pdfplumber
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from pdf_legend_parser import parse_legend

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
                "bbox": {
                    "x0": round(item.bbox[0], 1),
                    "y0": round(item.bbox[1], 1),
                    "x1": round(item.bbox[2], 1),
                    "y1": round(item.bbox[3], 1),
                },
            }
            for item in result.items
        ],
        "legend_type": legend_type,
        "raw_words_count": raw_words_count,
        "columns_detected": result.columns_detected,
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
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_app:app", host="0.0.0.0", port=8000, reload=True)
