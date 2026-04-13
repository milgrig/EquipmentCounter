/**
 * app.js — Client-side logic for PDF Legend Viewer
 * Works on both index.html and viewer.html pages.
 */

(function () {
    "use strict";

    // ── Theme toggle (shared across pages) ──────────────────────
    const themeBtn = document.getElementById("btn-theme");
    if (themeBtn) {
        function applyTheme(dark) {
            document.body.classList.toggle("dark", dark);
            localStorage.setItem("theme", dark ? "dark" : "light");
            themeBtn.textContent = dark ? "\u2600" : "\u263D";
        }
        themeBtn.addEventListener("click", () =>
            applyTheme(!document.body.classList.contains("dark"))
        );
        applyTheme(localStorage.getItem("theme") === "dark");
    }

    // Only run viewer logic on the viewer page
    if (document.body.dataset.page !== "viewer") return;

    // ================================================================
    // VIEWER PAGE LOGIC
    // ================================================================

    const FILE_ID = window.__FILE_ID__;
    const FILENAME = window.__FILENAME__;
    const PAGE_COUNT = window.__PAGE_COUNT__;

    // ── State ────────────────────────────────────────────────────
    let currentPage = 0;
    let zoomLevel = 1.0;
    let legendData = null;
    let pageWidth = 0;
    let pageHeight = 0;
    let renderDPI = 150;

    // ── DOM refs ─────────────────────────────────────────────────
    const pdfImage = document.getElementById("pdf-image");
    const pdfViewport = document.getElementById("pdf-viewport");
    const pdfScroll = document.getElementById("pdf-scroll");
    const highlightCanvas = document.getElementById("highlight-canvas");
    const pdfLoading = document.getElementById("pdf-loading");
    const legendLoading = document.getElementById("legend-loading");
    const legendStatus = document.getElementById("legend-status");
    const legendTableWrap = document.getElementById("legend-table-wrap");
    const legendTbody = document.getElementById("legend-tbody");
    const debugPanel = document.getElementById("debug-panel");
    const debugTbody = document.getElementById("debug-tbody");
    const chkOverlay = document.getElementById("chk-overlay");

    // ── Initialize ───────────────────────────────────────────────
    loadLegend();
    loadPage(0);

    // ── Page navigation ──────────────────────────────────────────
    document.getElementById("btn-prev-page").addEventListener("click", () => {
        if (currentPage > 0) loadPage(currentPage - 1);
    });
    document.getElementById("btn-next-page").addEventListener("click", () => {
        if (currentPage < PAGE_COUNT - 1) loadPage(currentPage + 1);
    });

    // ── Zoom controls ────────────────────────────────────────────
    const ZOOM_STEPS = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 2.0, 3.0, 4.0];
    const zoomInfo = document.getElementById("zoom-info");

    function setZoom(level) {
        zoomLevel = Math.max(0.1, Math.min(5.0, level));
        pdfScroll.style.transform = `scale(${zoomLevel})`;
        zoomInfo.textContent = Math.round(zoomLevel * 100) + "%";
    }

    document.getElementById("btn-zoom-in").addEventListener("click", () => {
        const next = ZOOM_STEPS.find((z) => z > zoomLevel + 0.01) || zoomLevel * 1.25;
        setZoom(next);
    });
    document.getElementById("btn-zoom-out").addEventListener("click", () => {
        const prev = [...ZOOM_STEPS].reverse().find((z) => z < zoomLevel - 0.01) || zoomLevel * 0.8;
        setZoom(prev);
    });
    document.getElementById("btn-zoom-fit").addEventListener("click", () => {
        if (!pdfImage.naturalWidth) return;
        const vpW = pdfViewport.clientWidth;
        setZoom(vpW / pdfImage.naturalWidth);
    });
    document.getElementById("btn-zoom-reset").addEventListener("click", () => setZoom(1.0));

    // Mouse wheel zoom
    pdfViewport.addEventListener("wheel", (e) => {
        if (e.ctrlKey || e.metaKey) {
            e.preventDefault();
            const delta = e.deltaY > 0 ? 0.9 : 1.1;
            setZoom(zoomLevel * delta);
        }
    }, { passive: false });

    // ── Pan via mouse drag ───────────────────────────────────────
    let isPanning = false;
    let panStart = { x: 0, y: 0 };
    let scrollStart = { x: 0, y: 0 };

    pdfViewport.addEventListener("mousedown", (e) => {
        if (e.button !== 0) return;
        isPanning = true;
        panStart = { x: e.clientX, y: e.clientY };
        scrollStart = { x: pdfViewport.scrollLeft, y: pdfViewport.scrollTop };
        pdfViewport.style.cursor = "grabbing";
        e.preventDefault();
    });

    window.addEventListener("mousemove", (e) => {
        if (!isPanning) return;
        pdfViewport.scrollLeft = scrollStart.x - (e.clientX - panStart.x);
        pdfViewport.scrollTop = scrollStart.y - (e.clientY - panStart.y);
    });

    window.addEventListener("mouseup", () => {
        if (isPanning) {
            isPanning = false;
            pdfViewport.style.cursor = "grab";
        }
    });

    // ── Overlay toggle ───────────────────────────────────────────
    chkOverlay.addEventListener("change", () => drawOverlay());

    // ── Load PDF page ────────────────────────────────────────────
    async function loadPage(pageIdx) {
        currentPage = pageIdx;
        document.getElementById("current-page").textContent = pageIdx + 1;

        pdfLoading.classList.remove("hidden");
        const url = `/api/file/${FILE_ID}/render?page=${pageIdx}&dpi=${renderDPI}`;

        try {
            const resp = await fetch(url);
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const blob = await resp.blob();

            // Revoke old object URL
            if (pdfImage.src && pdfImage.src.startsWith("blob:")) {
                URL.revokeObjectURL(pdfImage.src);
            }

            pdfImage.src = URL.createObjectURL(blob);
            pdfImage.onload = () => {
                drawOverlay();
                // Auto-fit on first load
                if (zoomLevel === 1.0 && pdfImage.naturalWidth > pdfViewport.clientWidth) {
                    setZoom(pdfViewport.clientWidth / pdfImage.naturalWidth);
                }
            };
        } catch (err) {
            console.error("Page load error:", err);
        } finally {
            pdfLoading.classList.add("hidden");
        }
    }

    // ── Load legend data ─────────────────────────────────────────
    async function loadLegend() {
        const t0 = performance.now();
        legendLoading.classList.remove("hidden");

        try {
            const resp = await fetch(`/api/file/${FILE_ID}/legend`);
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            legendData = await resp.json();

            const elapsed = ((performance.now() - t0) / 1000).toFixed(2);
            renderLegendStatus(legendData);
            renderLegendItems(legendData);
            renderStatusBar(legendData, elapsed);

            // If legend is on a different page, go there
            if (legendData.legend_found && legendData.page !== currentPage) {
                loadPage(legendData.page);
            } else {
                drawOverlay();
            }
        } catch (err) {
            console.error("Legend load error:", err);
            legendLoading.classList.add("hidden");
            legendStatus.classList.remove("hidden");
            document.getElementById("badge-found").textContent = "Ошибка загрузки";
            document.getElementById("badge-found").className = "status-badge badge-red";
        } finally {
            legendLoading.classList.add("hidden");
        }
    }

    // ── Render legend status badges ──────────────────────────────
    function renderLegendStatus(data) {
        const foundEl = document.getElementById("badge-found");
        const typeEl = document.getElementById("badge-type");
        const countEl = document.getElementById("badge-count");

        if (data.legend_found) {
            foundEl.textContent = "\u2713 Легенда найдена";
            foundEl.className = "status-badge badge-green";
        } else {
            foundEl.textContent = "\u2717 Легенда не найдена";
            foundEl.className = "status-badge badge-red";
        }

        const typeNames = {
            numbered: "Нумерованная",
            graphical: "Графическая",
            mixed: "Смешанная",
        };
        typeEl.textContent = typeNames[data.legend_type] || data.legend_type;
        typeEl.className = "status-badge badge-blue";

        countEl.textContent = `${data.items ? data.items.length : 0} элементов`;
        countEl.className = "status-badge badge-yellow";

        legendStatus.classList.remove("hidden");
    }

    // ── Render legend items table ────────────────────────────────
    function renderLegendItems(data) {
        legendTbody.innerHTML = "";

        if (!data.items || data.items.length === 0) {
            legendTableWrap.classList.add("hidden");
            return;
        }

        data.items.forEach((item, i) => {
            const tr = document.createElement("tr");
            tr.dataset.index = i;

            const catClass = item.category
                ? `cat-${item.category.replace(/\s/g, "-")}`
                : "";

            const symHtml = item.symbol
                ? `<span>${escapeHtml(item.symbol)}</span>`
                : `<span class="sym-none">&#9679;</span>`;

            const catHtml = item.category
                ? `<span class="category-tag ${catClass}">${escapeHtml(item.category)}</span>`
                : "";

            tr.innerHTML = `
                <td style="text-align:center;color:var(--text-muted)">${i + 1}</td>
                <td class="sym-cell">${symHtml}</td>
                <td class="desc-cell">${escapeHtml(item.description)}</td>
                <td class="cat-cell">${catHtml}</td>
            `;

            // Click → scroll PDF to item position & highlight
            tr.addEventListener("click", () => scrollToItem(i, item));

            legendTbody.appendChild(tr);
        });

        legendTableWrap.classList.remove("hidden");
    }

    // ── Scroll PDF to an item's bbox ─────────────────────────────
    function scrollToItem(index, item) {
        // Highlight the table row
        legendTbody.querySelectorAll("tr").forEach((r) => r.classList.remove("active-row"));
        const targetRow = legendTbody.querySelector(`tr[data-index="${index}"]`);
        if (targetRow) {
            targetRow.classList.add("active-row");
            targetRow.classList.remove("flash-row");
            void targetRow.offsetWidth; // force reflow
            targetRow.classList.add("flash-row");
        }

        // Scroll PDF viewport to the item's bbox
        if (!pdfImage.naturalWidth || !legendData) return;

        const bb = item.bbox;
        const scaleX = pdfImage.naturalWidth / (legendData.legend_bbox.x1 > 3000 ? legendData.legend_bbox.x1 * 1.1 : 3370);
        const scaleY = pdfImage.naturalHeight / (legendData.legend_bbox.y1 > 2000 ? legendData.legend_bbox.y1 * 1.2 : 2384);

        // Use DPI-based scale: pdfplumber coords are in PDF points (72 dpi),
        // rendered image is at renderDPI
        const pxPerPt = renderDPI / 72.0;
        const cx = ((bb.x0 + bb.x1) / 2) * pxPerPt * zoomLevel;
        const cy = ((bb.y0 + bb.y1) / 2) * pxPerPt * zoomLevel;

        pdfViewport.scrollTo({
            left: cx - pdfViewport.clientWidth / 2,
            top: cy - pdfViewport.clientHeight / 2,
            behavior: "smooth",
        });

        // Flash-highlight the item on canvas
        flashHighlightItem(item, pxPerPt);
    }

    // ── Draw overlay rectangles on canvas ────────────────────────
    function drawOverlay() {
        if (!pdfImage.naturalWidth) return;

        const w = pdfImage.naturalWidth;
        const h = pdfImage.naturalHeight;
        highlightCanvas.width = w;
        highlightCanvas.height = h;
        highlightCanvas.style.width = w + "px";
        highlightCanvas.style.height = h + "px";

        const ctx = highlightCanvas.getContext("2d");
        ctx.clearRect(0, 0, w, h);

        if (!chkOverlay.checked || !legendData || !legendData.legend_found) return;

        const pxPerPt = renderDPI / 72.0;

        // Draw legend bbox (red outline, light fill)
        const lb = legendData.legend_bbox;
        ctx.strokeStyle = "rgba(220, 50, 50, 0.7)";
        ctx.lineWidth = 3;
        ctx.fillStyle = "rgba(255, 200, 150, 0.15)";
        const lx = lb.x0 * pxPerPt;
        const ly = lb.y0 * pxPerPt;
        const lw = (lb.x1 - lb.x0) * pxPerPt;
        const lh = (lb.y1 - lb.y0) * pxPerPt;
        ctx.fillRect(lx, ly, lw, lh);
        ctx.strokeRect(lx, ly, lw, lh);

        // Draw individual item bboxes (blue)
        ctx.strokeStyle = "rgba(74, 144, 217, 0.5)";
        ctx.lineWidth = 1;
        ctx.fillStyle = "rgba(74, 144, 217, 0.08)";

        legendData.items.forEach((item) => {
            const bb = item.bbox;
            const ix = bb.x0 * pxPerPt;
            const iy = bb.y0 * pxPerPt;
            const iw = (bb.x1 - bb.x0) * pxPerPt;
            const ih = (bb.y1 - bb.y0) * pxPerPt;
            ctx.fillRect(ix, iy, iw, ih);
            ctx.strokeRect(ix, iy, iw, ih);
        });
    }

    // ── Flash-highlight a specific item on canvas ────────────────
    function flashHighlightItem(item, pxPerPt) {
        const ctx = highlightCanvas.getContext("2d");
        const bb = item.bbox;
        const ix = bb.x0 * pxPerPt;
        const iy = bb.y0 * pxPerPt;
        const iw = (bb.x1 - bb.x0) * pxPerPt;
        const ih = (bb.y1 - bb.y0) * pxPerPt;
        const pad = 4;

        let opacity = 0.5;
        const fadeInterval = setInterval(() => {
            drawOverlay(); // redraw base
            const ctx2 = highlightCanvas.getContext("2d");
            ctx2.fillStyle = `rgba(255, 100, 50, ${opacity})`;
            ctx2.strokeStyle = `rgba(255, 50, 0, ${Math.min(1, opacity * 2)})`;
            ctx2.lineWidth = 3;
            ctx2.fillRect(ix - pad, iy - pad, iw + pad * 2, ih + pad * 2);
            ctx2.strokeRect(ix - pad, iy - pad, iw + pad * 2, ih + pad * 2);
            opacity -= 0.04;
            if (opacity <= 0) {
                clearInterval(fadeInterval);
                drawOverlay();
            }
        }, 50);
    }

    // ── Status bar ───────────────────────────────────────────────
    function renderStatusBar(data, elapsed) {
        document.getElementById("stat-items").textContent =
            `Элементов: ${data.items ? data.items.length : 0}`;

        if (data.legend_found) {
            const bb = data.legend_bbox;
            document.getElementById("stat-bbox").textContent =
                `BBox: (${bb.x0}, ${bb.y0}) \u2014 (${bb.x1}, ${bb.y1})`;
        } else {
            document.getElementById("stat-bbox").textContent = "BBox: \u2014";
        }

        document.getElementById("stat-time").textContent = `Время: ${elapsed}с`;
        document.getElementById("stat-words").textContent =
            `Слов: ${data.raw_words_count || 0}`;
    }

    // ── Export JSON ───────────────────────────────────────────────
    document.getElementById("btn-export").addEventListener("click", () => {
        if (!legendData) return;
        const json = JSON.stringify(legendData, null, 2);
        const blob = new Blob([json], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = FILENAME.replace(/\.pdf$/i, "") + "_legend.json";
        a.click();
        URL.revokeObjectURL(url);
    });

    // ── Debug words toggle ───────────────────────────────────────
    let debugLoaded = false;

    document.getElementById("btn-debug-toggle").addEventListener("click", async () => {
        if (!debugPanel.classList.contains("hidden")) {
            debugPanel.classList.add("hidden");
            return;
        }

        if (!debugLoaded) {
            try {
                const resp = await fetch(`/api/file/${FILE_ID}/debug_words?page=${currentPage}`);
                if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
                const data = await resp.json();

                document.getElementById("debug-info").textContent =
                    `${data.page_width} \u00d7 ${data.page_height} | ${data.words_count} слов`;

                debugTbody.innerHTML = "";
                const limit = Math.min(data.words.length, 1000);
                for (let i = 0; i < limit; i++) {
                    const w = data.words[i];
                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td>${i + 1}</td>
                        <td>${escapeHtml(w.text)}</td>
                        <td>${w.x0}</td>
                        <td>${w.top}</td>
                        <td>${w.x1}</td>
                        <td>${w.bottom}</td>
                    `;
                    debugTbody.appendChild(tr);
                }

                if (data.words.length > limit) {
                    const tr = document.createElement("tr");
                    tr.innerHTML = `<td colspan="6" style="text-align:center;color:var(--text-muted)">
                        \u2026 ${data.words.length - limit} ещё (показано ${limit})
                    </td>`;
                    debugTbody.appendChild(tr);
                }

                debugLoaded = true;
            } catch (err) {
                console.error("Debug load error:", err);
                return;
            }
        }

        debugPanel.classList.remove("hidden");
    });

    // ── Resizer (drag to resize panels) ──────────────────────────
    const resizer = document.getElementById("resizer");
    const pdfPanel = document.getElementById("pdf-panel");
    const legendPanel = document.getElementById("legend-panel");

    let isResizing = false;

    resizer.addEventListener("mousedown", (e) => {
        isResizing = true;
        resizer.classList.add("active");
        e.preventDefault();
    });

    window.addEventListener("mousemove", (e) => {
        if (!isResizing) return;
        const container = document.querySelector(".viewer-container");
        const containerRect = container.getBoundingClientRect();
        const ratio = (e.clientX - containerRect.left) / containerRect.width;
        const clamped = Math.max(0.2, Math.min(0.85, ratio));

        pdfPanel.style.flex = "none";
        legendPanel.style.flex = "none";
        pdfPanel.style.width = (clamped * 100) + "%";
        legendPanel.style.width = ((1 - clamped) * 100 - 0.3) + "%";
    });

    window.addEventListener("mouseup", () => {
        if (isResizing) {
            isResizing = false;
            resizer.classList.remove("active");
        }
    });

    // ── Helper ───────────────────────────────────────────────────
    function escapeHtml(text) {
        const div = document.createElement("div");
        div.textContent = text;
        return div.innerHTML;
    }

})();
