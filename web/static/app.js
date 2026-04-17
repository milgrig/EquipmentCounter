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

    // Color filter state (T123): tracks whether we're showing a color-filtered page
    let legendColorFilterActive = false;  // true when color filter is applied from legend click
    let legendColorFilterHex = null;      // hex string of filtered color, e.g. "FF0000"
    let legendColorFilterName = "";       // display name, e.g. "red"

    // Map legend item color names to hex values for render_filtered API
    const COLOR_NAME_TO_HEX = {
        "red":   ["FF0000"],
        "blue":  ["0000FF"],
        "green": ["00FF00"],
        "black": ["000000"],
        "grey":  [],  // grey includes many shades; show all greys
    };

    // Map color names to CSS display colors for UI
    const COLOR_NAME_TO_CSS = {
        "red":   "#dc3545",
        "blue":  "#4a90d9",
        "green": "#28a745",
        "black": "#333",
        "grey":  "#999",
    };

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
                if (findState && findState.active) {
                    invalidateFindPixelMask();
                }
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

            const textFallback = item.symbol
                ? escapeHtml(item.symbol)
                : `<span class="sym-none">&#9679;</span>`;
            const imgUrl = item.image_url || `/api/file/${FILE_ID}/symbol_image/${i}`;
            const symHtml = `<img class="sym-img" src="${imgUrl}" alt="" `
                + `onerror="this.style.display='none';this.nextElementSibling.style.display='inline'" />`
                + `<span class="sym-text-fallback" style="display:none">${textFallback}</span>`;

            const catHtml = item.category
                ? `<span class="category-tag ${catClass}">${escapeHtml(item.category)}</span>`
                : "";

            // Color dot indicator (T123): shows the equipment color from the PDF
            const colorDotHtml = item.color && COLOR_NAME_TO_CSS[item.color]
                ? `<span class="legend-color-dot" style="background:${COLOR_NAME_TO_CSS[item.color]}" title="${item.color}"></span>`
                : "";

            tr.innerHTML = `
                <td style="text-align:center;color:var(--text-muted)">${i + 1}</td>
                <td class="sym-cell">${symHtml}</td>
                <td class="desc-cell">${colorDotHtml}${escapeHtml(item.description)}</td>
                <td class="cat-cell">${catHtml}</td>
            `;

            // Click → find all instances on PDF & highlight
            tr.addEventListener("click", (e) => {
                if (e.ctrlKey || e.metaKey) {
                    // Ctrl+click: add/remove from multi-selection
                    toggleFindSelection(i, item);
                } else {
                    findEquipment(i, item);
                }
            });

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

    // ── Color filters ─────────────────────────────────────────────
    const colorPanel = document.getElementById("color-filter-panel");
    const cfList = document.getElementById("cf-list");
    const cfLoading = document.getElementById("cf-loading");
    let colorData = null;
    let colorsLoaded = false;
    let activeColorFilter = null; // null = no filter (all visible)

    document.getElementById("btn-colors-toggle").addEventListener("click", async () => {
        if (!colorPanel.classList.contains("hidden")) {
            colorPanel.classList.add("hidden");
            return;
        }

        if (!colorsLoaded) {
            colorPanel.classList.remove("hidden");
            cfLoading.classList.remove("hidden");
            try {
                const resp = await fetch(`/api/file/${FILE_ID}/colors?page=${currentPage}`);
                if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
                colorData = await resp.json();
                renderColorList(colorData.colors);
                colorsLoaded = true;
            } catch (err) {
                console.error("Color load error:", err);
                cfList.innerHTML = '<div style="padding:1rem;color:var(--text-muted)">Ошибка загрузки цветов</div>';
            } finally {
                cfLoading.classList.add("hidden");
            }
        }

        colorPanel.classList.remove("hidden");
    });

    function renderColorList(colors) {
        cfList.innerHTML = "";
        colors.forEach((c) => {
            const item = document.createElement("label");
            item.className = "cf-item";
            item.dataset.hex = c.hex;

            const checked = activeColorFilter === null || (activeColorFilter && activeColorFilter.has(c.hex));
            item.innerHTML = `
                <input type="checkbox" class="cf-check" data-hex="${c.hex}" ${checked ? "checked" : ""} />
                <span class="cf-swatch" style="background:#${c.hex}"></span>
                <span class="cf-label">${c.label || "#" + c.hex}</span>
                <span class="cf-count">${c.total}</span>
            `;
            cfList.appendChild(item);
        });

        // Bind change events
        cfList.querySelectorAll(".cf-check").forEach((cb) => {
            cb.addEventListener("change", () => applyColorFilter());
        });
    }

    function applyColorFilter() {
        const checkboxes = cfList.querySelectorAll(".cf-check");
        const total = checkboxes.length;
        const checked = Array.from(checkboxes).filter((cb) => cb.checked);

        if (checked.length === total || checked.length === 0) {
            // All or none checked → show normal render
            activeColorFilter = null;
            loadPage(currentPage);
            return;
        }

        // Build show list
        const showHexes = checked.map((cb) => cb.dataset.hex);
        activeColorFilter = new Set(showHexes);
        loadFilteredPage(showHexes);
    }

    async function loadFilteredPage(showHexes) {
        pdfLoading.classList.remove("hidden");
        const showParam = showHexes.join(",");
        const url = `/api/file/${FILE_ID}/render_filtered?page=${currentPage}&dpi=${renderDPI}&show=${showParam}`;

        try {
            const resp = await fetch(url);
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const blob = await resp.blob();

            if (pdfImage.src && pdfImage.src.startsWith("blob:")) {
                URL.revokeObjectURL(pdfImage.src);
            }

            pdfImage.src = URL.createObjectURL(blob);
            pdfImage.onload = () => {
                if (findState && findState.active) {
                    invalidateFindPixelMask();
                }
                drawOverlay();
            };
        } catch (err) {
            console.error("Filtered render error:", err);
        } finally {
            pdfLoading.classList.add("hidden");
        }
    }

    // Quick filter buttons
    document.getElementById("cf-all").addEventListener("click", () => {
        cfList.querySelectorAll(".cf-check").forEach((cb) => { cb.checked = true; });
        applyColorFilter();
    });

    document.getElementById("cf-electric").addEventListener("click", () => {
        if (!colorData) return;
        cfList.querySelectorAll(".cf-check").forEach((cb) => {
            const hex = cb.dataset.hex;
            // Red (FF0000) or Blue (0000FF) — with tolerance
            cb.checked = _isRedHex(hex) || _isBlueHex(hex);
        });
        applyColorFilter();
    });

    document.getElementById("cf-alarm").addEventListener("click", () => {
        if (!colorData) return;
        cfList.querySelectorAll(".cf-check").forEach((cb) => {
            cb.checked = _isRedHex(cb.dataset.hex);
        });
        applyColorFilter();
    });

    document.getElementById("cf-work").addEventListener("click", () => {
        if (!colorData) return;
        cfList.querySelectorAll(".cf-check").forEach((cb) => {
            cb.checked = _isBlueHex(cb.dataset.hex);
        });
        applyColorFilter();
    });

    function _isRedHex(hex) {
        // Match reds: R channel high, G and B low
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        return r > 180 && g < 80 && b < 80;
    }

    function _isBlueHex(hex) {
        // Match blues: B channel high, R and G low
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        return b > 180 && r < 80 && g < 80;
    }

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

    // ================================================================
    // TAB SWITCHING
    // ================================================================

    document.querySelectorAll(".tab-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
            document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
            document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
            btn.classList.add("active");
            const target = document.getElementById(btn.dataset.tab);
            if (target) target.classList.add("active");
        });
    });

    // ================================================================
    // COUNTING DASHBOARD
    // ================================================================

    let countData = null; // cached results from /count/all
    let countOverlayState = { text: false, visual: false, cables: false };
    let countEventSource = null; // SSE connection

    const countLoading = document.getElementById("count-loading");
    const countLoadingText = document.getElementById("count-loading-text");
    const countProgress = document.getElementById("count-progress");
    const countProgressBar = document.getElementById("count-progress-bar");
    const countProgressStep = document.getElementById("count-progress-step");
    const countProgressLabel = document.getElementById("count-progress-label");
    const countProgressLog = document.getElementById("count-progress-log");
    const countStats = document.getElementById("count-stats");
    const countStatsBadges = document.getElementById("count-stats-badges");
    const overlayControls = document.getElementById("overlay-controls");
    const equipSection = document.getElementById("equip-section");
    const cableSection = document.getElementById("cable-section");
    const equipTbody = document.getElementById("equip-tbody");
    const cableTbody = document.getElementById("cable-tbody");

    // ── Run all counting methods (SSE streaming) ────────────────
    document.getElementById("btn-run-count").addEventListener("click", runCounting);

    function appendLogLine(text, icon) {
        const div = document.createElement("div");
        div.className = "log-line running";
        div.innerHTML = `<span class="log-icon">${icon || "\u23f3"}</span><span class="log-text">${escapeHtml(text)}</span>`;
        countProgressLog.appendChild(div);
        countProgressLog.scrollTop = countProgressLog.scrollHeight;
        return div;
    }

    function finishLastLogLine(countVal, isError) {
        const lines = countProgressLog.querySelectorAll(".log-line.running");
        const last = lines[lines.length - 1];
        if (!last) return;
        last.classList.remove("running");
        last.classList.add(isError ? "error" : "done");
        const iconEl = last.querySelector(".log-icon");
        if (iconEl) iconEl.textContent = isError ? "\u274c" : "\u2705";
        if (countVal !== undefined && countVal !== null && !isError) {
            const span = document.createElement("span");
            span.className = "log-count";
            span.textContent = `${countVal} \u0448\u0442.`;
            last.appendChild(span);
        }
    }

    async function runCounting() {
        if (!legendData) {
            alert("\u041b\u0435\u0433\u0435\u043d\u0434\u0430 \u043d\u0435 \u0437\u0430\u0433\u0440\u0443\u0436\u0435\u043d\u0430.");
            return;
        }

        // Close previous SSE if running
        if (countEventSource) {
            countEventSource.close();
            countEventSource = null;
        }

        // Reset UI
        countLoading.classList.add("hidden");
        countProgress.classList.remove("hidden");
        countProgressBar.style.width = "0%";
        countProgressStep.textContent = "";
        countProgressLabel.textContent = "\u041f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043a\u0430...";
        countProgressLog.innerHTML = "";
        countStats.classList.add("hidden");
        equipSection.classList.add("hidden");
        cableSection.classList.add("hidden");
        overlayControls.classList.add("hidden");
        equipTbody.innerHTML = "";

        // Accumulate results
        countData = { results: {}, errors: {} };
        let estimatedTotal = 5; // will be updated from start event
        let currentStep = 0;

        const es = new EventSource(`/api/file/${FILE_ID}/count/stream`);
        countEventSource = es;

        es.addEventListener("start", (e) => {
            const d = JSON.parse(e.data);
            // legend_items + 4 (legend parse, text, cables, geometry)
            estimatedTotal = (d.legend_items || 0) + 4;
        });

        es.addEventListener("progress", (e) => {
            const d = JSON.parse(e.data);
            currentStep = d.step;
            const pct = Math.min(95, Math.round((currentStep / estimatedTotal) * 100));
            countProgressBar.style.width = `${pct}%`;
            countProgressStep.textContent = `${currentStep}`;
            countProgressLabel.textContent = d.label || "";
            // Only add log line for non-symbol progress (text/cables/geometry)
            // Symbol log lines are created from step_done where we have the real label
            if (d.label && !d.label.startsWith("\u041f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043a\u0430") && !d.label.startsWith("\u0412\u0438\u0437\u0443\u0430\u043b\u044c\u043d\u044b\u0439")) {
                appendLogLine(d.label, "\u23f3");
            }
        });

        es.addEventListener("step_done", (e) => {
            const d = JSON.parse(e.data);
            currentStep = d.step;
            const pct = Math.min(95, Math.round((currentStep / estimatedTotal) * 100));
            countProgressBar.style.width = `${pct}%`;
            countProgressStep.textContent = `${currentStep}`;

            if (d.error) {
                // If there's a pending running line, finish it as error
                finishLastLogLine(null, true);
            } else if (d.type === "symbol") {
                // Create log line with symbol image for visual items
                const div = document.createElement("div");
                div.className = "log-line done";
                const imgUrl = `/api/file/${FILE_ID}/symbol_image/${d.symbol_index}`;
                const descText = d.description || "";
                const symText = d.symbol && d.symbol !== "?" ? d.symbol + " — " : "";
                div.innerHTML = `<span class="log-icon">\u2705</span>`
                    + `<img class="log-sym-img" src="${imgUrl}" alt="" onerror="this.style.display='none'" />`
                    + `<span class="log-text">${escapeHtml(symText + descText)}</span>`;
                if (d.count !== undefined) {
                    const span = document.createElement("span");
                    span.className = "log-count";
                    span.textContent = `${d.count} \u0448\u0442.`;
                    div.appendChild(span);
                }
                countProgressLog.appendChild(div);
                countProgressLog.scrollTop = countProgressLog.scrollHeight;
                countProgressLabel.textContent = d.label || "";
                // Accumulate visual counts
                if (!countData.results.visual) {
                    countData.results.visual = { counts: {}, descriptions: {}, matches: [] };
                }
                countData.results.visual.counts[String(d.symbol_index)] = d.count;
            } else {
                finishLastLogLine(d.count, false);
            }
        });

        es.addEventListener("done", (e) => {
            const d = JSON.parse(e.data);
            es.close();
            countEventSource = null;

            countProgressBar.style.width = "100%";
            countProgressLabel.textContent = `\u0413\u043e\u0442\u043e\u0432\u043e \u0437\u0430 ${d.total_elapsed_s}\u0441`;

            // Merge full results from done event
            if (d.results) {
                countData.results = d.results;
            }
            if (d.errors) {
                countData.errors = d.errors;
            }
            countData.elapsed_s = d.total_elapsed_s;
            countData.legend_page = d.legend_page;
            countData.legend_items = d.legend_items;

            renderCountStats(countData);
            renderEquipTable(countData);
            renderCableTable(countData);
            overlayControls.classList.remove("hidden");
        });

        es.addEventListener("error", (e) => {
            // SSE error (connection lost, etc.)
            if (es.readyState === EventSource.CLOSED) return;
            es.close();
            countEventSource = null;
            finishLastLogLine(null, true);
            countProgressLabel.textContent = "\u041e\u0448\u0438\u0431\u043a\u0430 \u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u044f";
            countStatsBadges.innerHTML = `<span class="status-badge badge-red">\u041e\u0448\u0438\u0431\u043a\u0430 SSE</span>`;
            countStats.classList.remove("hidden");
        });
    }

    // ── Render count stats badges ───────────────────────────────
    function renderCountStats(data) {
        const badges = [];
        const elapsed = data.elapsed_s || 0;
        badges.push(`<span class="status-badge badge-blue">Total: ${elapsed}s</span>`);

        const methods = ["text", "cables", "geometry", "visual"];
        for (const m of methods) {
            if (data.results && data.results[m]) {
                const t = data.results[m].elapsed_s || 0;
                badges.push(`<span class="status-badge badge-green">${m}: ${t}s</span>`);
            }
            if (data.errors && data.errors[m]) {
                badges.push(`<span class="status-badge badge-yellow">${m}: err</span>`);
            }
        }

        countStatsBadges.innerHTML = badges.join("");
        countStats.classList.remove("hidden");
    }

    // ── Render equipment comparison table ────────────────────────
    function renderEquipTable(data) {
        equipTbody.innerHTML = "";
        if (!legendData || !legendData.items) return;

        const textCounts = (data.results && data.results.text)
            ? data.results.text.counts || {}
            : {};
        const visCounts = (data.results && data.results.visual)
            ? data.results.visual.counts || {}
            : {};
        const visDescs = (data.results && data.results.visual)
            ? data.results.visual.descriptions || {}
            : {};

        let hasData = false;

        legendData.items.forEach((item, i) => {
            const sym = item.symbol || "";
            const textVal = sym ? (textCounts[sym] || 0) : 0;

            // Visual counts are keyed by symbol_index
            const visVal = visCounts[String(i)] || 0;

            // Only show rows that have any count or a symbol
            if (textVal === 0 && visVal === 0 && !sym) return;
            hasData = true;

            // Visual is primary: use visual count when available, text as fallback
            const total = visVal > 0 ? visVal : textVal;
            const mismatch = textVal > 0 && visVal > 0 && textVal !== visVal;

            const tr = document.createElement("tr");
            tr.dataset.index = i;
            tr.dataset.symbol = sym;

            const imgUrl = `/api/file/${FILE_ID}/symbol_image/${i}`;
            const textFallback = sym
                ? `<span style="font-weight:700;font-family:var(--mono)">${escapeHtml(sym)}</span>`
                : `<span style="color:var(--text-faint)">&#9679;</span>`;
            const symHtml = `<img class="sym-img" src="${imgUrl}" alt="" `
                + `onerror="this.style.display='none';this.nextElementSibling.style.display='inline'" />`
                + `<span style="display:none">${textFallback}</span>`;

            const textClass = textVal > 0 ? "cnt-cell" : "cnt-cell cnt-zero";
            const visClass = visVal > 0 ? "cnt-cell" : "cnt-cell cnt-zero";
            const mismatchClass = mismatch ? " cnt-mismatch" : "";

            tr.innerHTML = `
                <td style="text-align:center;color:var(--text-muted)">${i + 1}</td>
                <td style="text-align:center">${symHtml}</td>
                <td class="desc-cell">${escapeHtml(item.description).substring(0, 60)}</td>
                <td class="${textClass}${mismatchClass}">${textVal}</td>
                <td class="${visClass}${mismatchClass}">${visVal}</td>
                <td class="cnt-cell" style="font-weight:700">${total}</td>
            `;

            // Click to highlight positions on PDF
            tr.addEventListener("click", () => {
                highlightCountPositions(sym, i, data);
            });

            equipTbody.appendChild(tr);
        });

        if (hasData) equipSection.classList.remove("hidden");
    }

    // ── Render cable comparison table ────────────────────────────
    function renderCableTable(data) {
        cableTbody.innerHTML = "";

        const cableData = (data.results && data.results.cables)
            ? data.results.cables
            : null;
        const geoData = (data.results && data.results.geometry)
            ? data.results.geometry
            : null;

        if (!cableData || cableData.total_runs === 0) return;

        // Build cable schedule rows
        const schedule = cableData.schedule || [];

        let totalAnnotLen = 0;
        let totalGeoRed = 0;
        let totalGeoBlue = 0;

        // Get geometric lengths by color
        if (geoData && geoData.routes) {
            for (const r of geoData.routes) {
                if (r.color === "red") totalGeoRed = r.total_length_m;
                if (r.color === "blue") totalGeoBlue = r.total_length_m;
            }
        }

        schedule.forEach((entry) => {
            const tr = document.createElement("tr");

            const colors = entry.colors || [];
            let colorDots = "";
            for (const c of colors) {
                colorDots += `<span class="cable-color-dot cable-color-${c}"></span>`;
            }

            const lengthAnnot = entry.total_length_m
                ? `${entry.total_length_m}m`
                : "\u2014";
            if (entry.total_length_m) totalAnnotLen += entry.total_length_m;

            const cs = (entry.cross_sections || []).join(", ");
            const types = (entry.cable_types || []).join(", ") || "\u2014";

            tr.innerHTML = `
                <td>${escapeHtml(entry.panel || "\u2014")}</td>
                <td>${escapeHtml(entry.group || "\u2014")}</td>
                <td style="font-family:var(--mono);font-size:0.78rem">${escapeHtml(cs)}</td>
                <td style="text-align:center">${lengthAnnot}</td>
                <td style="text-align:center;color:var(--text-muted)">\u2014</td>
                <td style="font-size:0.75rem">${escapeHtml(types)}</td>
                <td>${colorDots}</td>
            `;
            cableTbody.appendChild(tr);
        });

        // Summary row
        const summaryTr = document.createElement("tr");
        summaryTr.className = "cable-summary-row";

        let geoSummary = "\u2014";
        if (totalGeoRed > 0 || totalGeoBlue > 0) {
            const parts = [];
            if (totalGeoRed > 0) parts.push(`<span class="cable-color-dot cable-color-red"></span>${totalGeoRed.toFixed(1)}m`);
            if (totalGeoBlue > 0) parts.push(`<span class="cable-color-dot cable-color-blue"></span>${totalGeoBlue.toFixed(1)}m`);
            geoSummary = parts.join(" ");
        }

        summaryTr.innerHTML = `
            <td colspan="3" style="text-align:right;font-weight:600">ИТОГО</td>
            <td style="text-align:center;font-weight:600">${totalAnnotLen > 0 ? totalAnnotLen.toFixed(1) + "m" : "\u2014"}</td>
            <td style="text-align:center">${geoSummary}</td>
            <td colspan="2"></td>
        `;
        cableTbody.appendChild(summaryTr);

        cableSection.classList.remove("hidden");
    }

    // ── Highlight positions on PDF ──────────────────────────────
    function highlightCountPositions(symbol, index, data) {
        if (!pdfImage.naturalWidth) return;
        const pxPerPt = renderDPI / 72.0;
        const ctx = highlightCanvas.getContext("2d");

        // Redraw base overlay first
        drawOverlay();

        const positions = [];

        // Text positions
        if (data.results && data.results.text && data.results.text.positions) {
            const symPositions = data.results.text.positions[symbol] || [];
            for (const p of symPositions) {
                positions.push({ x: p.x * pxPerPt, y: p.y * pxPerPt, type: "text" });
            }
        }

        // Visual positions
        if (data.results && data.results.visual && data.results.visual.matches) {
            for (const m of data.results.visual.matches) {
                if (m.symbol_index === index) {
                    positions.push({ x: m.x * pxPerPt, y: m.y * pxPerPt, type: "visual" });
                }
            }
        }

        // Draw markers
        for (const p of positions) {
            const r = 8;
            if (p.type === "text") {
                ctx.beginPath();
                ctx.arc(p.x, p.y, r, 0, 2 * Math.PI);
                ctx.fillStyle = "rgba(74, 144, 217, 0.5)";
                ctx.fill();
                ctx.strokeStyle = "rgba(74, 144, 217, 0.9)";
                ctx.lineWidth = 2;
                ctx.stroke();
            } else {
                ctx.strokeStyle = "rgba(40, 167, 69, 0.9)";
                ctx.lineWidth = 2;
                ctx.fillStyle = "rgba(40, 167, 69, 0.3)";
                ctx.fillRect(p.x - r, p.y - r, r * 2, r * 2);
                ctx.strokeRect(p.x - r, p.y - r, r * 2, r * 2);
            }
        }

        // Scroll to first position
        if (positions.length > 0) {
            const first = positions[0];
            pdfViewport.scrollTo({
                left: first.x * zoomLevel - pdfViewport.clientWidth / 2,
                top: first.y * zoomLevel - pdfViewport.clientHeight / 2,
                behavior: "smooth",
            });
        }
    }

    // ── Overlay toggle handlers ─────────────────────────────────
    const chkText = document.getElementById("chk-overlay-text");
    const chkVisual = document.getElementById("chk-overlay-visual");
    const chkCables = document.getElementById("chk-overlay-cables");

    if (chkText) chkText.addEventListener("change", () => {
        countOverlayState.text = chkText.checked;
        drawCountOverlay();
    });
    if (chkVisual) chkVisual.addEventListener("change", () => {
        countOverlayState.visual = chkVisual.checked;
        drawCountOverlay();
    });
    if (chkCables) chkCables.addEventListener("change", () => {
        countOverlayState.cables = chkCables.checked;
        drawCountOverlay();
    });

    function drawCountOverlay() {
        if (!countData || !pdfImage.naturalWidth) return;

        // Redraw base
        drawOverlay();

        const pxPerPt = renderDPI / 72.0;
        const ctx = highlightCanvas.getContext("2d");

        // Method A: Text markers (blue circles)
        if (countOverlayState.text && countData.results && countData.results.text) {
            const positions = countData.results.text.positions || {};
            ctx.fillStyle = "rgba(74, 144, 217, 0.4)";
            ctx.strokeStyle = "rgba(74, 144, 217, 0.8)";
            ctx.lineWidth = 1.5;
            for (const sym in positions) {
                for (const p of positions[sym]) {
                    ctx.beginPath();
                    ctx.arc(p.x * pxPerPt, p.y * pxPerPt, 6, 0, 2 * Math.PI);
                    ctx.fill();
                    ctx.stroke();
                }
            }
        }

        // Method D: Visual matches (green rectangles)
        if (countOverlayState.visual && countData.results && countData.results.visual) {
            const matches = countData.results.visual.matches || [];
            ctx.strokeStyle = "rgba(40, 167, 69, 0.8)";
            ctx.fillStyle = "rgba(40, 167, 69, 0.2)";
            ctx.lineWidth = 1.5;
            for (const m of matches) {
                const r = 7;
                ctx.fillRect(m.x * pxPerPt - r, m.y * pxPerPt - r, r * 2, r * 2);
                ctx.strokeRect(m.x * pxPerPt - r, m.y * pxPerPt - r, r * 2, r * 2);
            }
        }

        // Method B: Cable positions (red/blue dots)
        if (countOverlayState.cables && countData.results && countData.results.cables) {
            const runs = countData.results.cables.runs || [];
            ctx.lineWidth = 1;
            for (const r of runs) {
                const px = r.position.x * pxPerPt;
                const py = r.position.y * pxPerPt;
                if (r.color === "red") {
                    ctx.fillStyle = "rgba(220, 53, 69, 0.5)";
                    ctx.strokeStyle = "rgba(220, 53, 69, 0.9)";
                } else if (r.color === "blue") {
                    ctx.fillStyle = "rgba(74, 144, 217, 0.5)";
                    ctx.strokeStyle = "rgba(74, 144, 217, 0.9)";
                } else {
                    ctx.fillStyle = "rgba(128, 128, 128, 0.4)";
                    ctx.strokeStyle = "rgba(128, 128, 128, 0.8)";
                }
                ctx.beginPath();
                ctx.arc(px, py, 4, 0, 2 * Math.PI);
                ctx.fill();
                ctx.stroke();
            }
        }
    }

    // ================================================================
    // INTERACTIVE EQUIPMENT FIND & HIGHLIGHT (T106)
    // ================================================================

    // State for find mode
    let findState = {
        active: false,
        selections: [],       // [{rowIndex, item, data}] — multi-select support
        currentIndex: 0,      // index into allPositions for navigation
        allPositions: [],     // flat list of all found positions across selections
        pulseTimer: null,     // animation timer
        pulsePhase: 0,
        pixelMaskCanvas: null, // offscreen canvas with highlighted equipment pixels
        pixelMaskCount: 0,     // number of pixels in mask (debug/status)
    };

    function invalidateFindPixelMask() {
        findState.pixelMaskCanvas = null;
        findState.pixelMaskCount = 0;
    }

    function buildFindPixelMask() {
        if (!pdfImage.naturalWidth || !findState.active || findState.allPositions.length === 0) {
            return null;
        }

        const w = pdfImage.naturalWidth;
        const h = pdfImage.naturalHeight;
        const pxPerPt = renderDPI / 72.0;

        // Read rendered PDF pixels once (base layer).
        const srcCanvas = document.createElement("canvas");
        srcCanvas.width = w;
        srcCanvas.height = h;
        const srcCtx = srcCanvas.getContext("2d", { willReadFrequently: true });
        srcCtx.drawImage(pdfImage, 0, 0, w, h);
        const srcData = srcCtx.getImageData(0, 0, w, h).data;

        // Build output mask canvas: only equipment-related pixels are painted.
        const maskCanvas = document.createElement("canvas");
        maskCanvas.width = w;
        maskCanvas.height = h;
        const maskCtx = maskCanvas.getContext("2d");
        const maskImg = maskCtx.createImageData(w, h);
        const out = maskImg.data;
        const seen = new Uint8Array(w * h);

        let selectedPixels = 0;
        const radius = 28;   // local window around each found position
        const r2 = radius * radius;
        const lumaThreshold = 228;
        const seedSearchRadius = 8;
        const maxBlobPixels = 3200;

        function inBounds(x, y) {
            return x >= 0 && y >= 0 && x < w && y < h;
        }

        function isDarkPixelAt(x, y) {
            const idx = (y * w + x) * 4;
            const a = srcData[idx + 3];
            if (a < 8) return false;
            const luma = (srcData[idx] + srcData[idx + 1] + srcData[idx + 2]) / 3;
            return luma <= lumaThreshold;
        }

        for (const pos of findState.allPositions) {
            const cx = Math.round(pos.x * pxPerPt);
            const cy = Math.round(pos.y * pxPerPt);
            const x0 = Math.max(0, cx - radius);
            const x1 = Math.min(w - 1, cx + radius);
            const y0 = Math.max(0, cy - radius);
            const y1 = Math.min(h - 1, cy + radius);

            // 1) Find a seed dark pixel nearest to reported position.
            let seedX = -1;
            let seedY = -1;
            let bestD2 = Infinity;
            for (let y = Math.max(y0, cy - seedSearchRadius); y <= Math.min(y1, cy + seedSearchRadius); y++) {
                for (let x = Math.max(x0, cx - seedSearchRadius); x <= Math.min(x1, cx + seedSearchRadius); x++) {
                    if (!isDarkPixelAt(x, y)) continue;
                    const dx = x - cx;
                    const dy = y - cy;
                    const d2 = dx * dx + dy * dy;
                    if (d2 < bestD2) {
                        bestD2 = d2;
                        seedX = x;
                        seedY = y;
                    }
                }
            }

            // If no near seed found, fallback to circular selection around center.
            if (seedX < 0 || seedY < 0) {
                for (let y = y0; y <= y1; y++) {
                    for (let x = x0; x <= x1; x++) {
                        const dx = x - cx;
                        const dy = y - cy;
                        if (dx * dx + dy * dy > r2) continue;
                        if (!isDarkPixelAt(x, y)) continue;
                        const pi = y * w + x;
                        if (seen[pi]) continue;
                        seen[pi] = 1;
                        selectedPixels++;
                        const idx = pi * 4;
                        out[idx] = 255;
                        out[idx + 1] = 190;
                        out[idx + 2] = 30;
                        out[idx + 3] = 255;
                    }
                }
                continue;
            }

            // 2) Flood-fill connected dark pixels from seed within local window.
            const queue = [[seedX, seedY]];
            const visitedLocal = new Uint8Array((x1 - x0 + 1) * (y1 - y0 + 1));
            const localWidth = x1 - x0 + 1;
            let qIndex = 0;
            let blobCount = 0;

            while (qIndex < queue.length && blobCount < maxBlobPixels) {
                const [x, y] = queue[qIndex++];
                if (!inBounds(x, y)) continue;
                if (x < x0 || x > x1 || y < y0 || y > y1) continue;
                const dx = x - cx;
                const dy = y - cy;
                if (dx * dx + dy * dy > r2) continue;

                const li = (y - y0) * localWidth + (x - x0);
                if (visitedLocal[li]) continue;
                visitedLocal[li] = 1;

                if (!isDarkPixelAt(x, y)) continue;

                const pi = y * w + x;
                if (!seen[pi]) {
                    seen[pi] = 1;
                    selectedPixels++;
                    blobCount++;
                    const idx = pi * 4;
                    out[idx] = 255;
                    out[idx + 1] = 190;
                    out[idx + 2] = 30;
                    out[idx + 3] = 255;
                }

                queue.push([x + 1, y]);
                queue.push([x - 1, y]);
                queue.push([x, y + 1]);
                queue.push([x, y - 1]);
                queue.push([x + 1, y + 1]);
                queue.push([x + 1, y - 1]);
                queue.push([x - 1, y + 1]);
                queue.push([x - 1, y - 1]);
            }
        }

        if (selectedPixels === 0) {
            return null;
        }

        maskCtx.putImageData(maskImg, 0, 0);
        findState.pixelMaskCount = selectedPixels;
        return maskCanvas;
    }

    function ensureFindPixelMask() {
        if (!findState.pixelMaskCanvas) {
            findState.pixelMaskCanvas = buildFindPixelMask();
        }
        return findState.pixelMaskCanvas;
    }

    function drawCurrentFindCursor(ctx, pxPerPt) {
        const pos = findState.allPositions[findState.currentIndex];
        if (!pos) return;
        const px = pos.x * pxPerPt;
        const py = pos.y * pxPerPt;
        const r = 8;
        ctx.save();
        ctx.strokeStyle = "rgba(255,255,255,0.95)";
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.moveTo(px - r, py);
        ctx.lineTo(px + r, py);
        ctx.moveTo(px, py - r);
        ctx.lineTo(px, py + r);
        ctx.stroke();
        ctx.restore();
    }

    // DOM refs for find nav bar
    const findNav = document.getElementById("find-nav");
    const findNavSymbol = document.getElementById("find-nav-symbol");
    const findNavDesc = document.getElementById("find-nav-desc");
    const findNavCounter = document.getElementById("find-nav-counter");
    const findNavStatus = document.getElementById("find-nav-status");

    // Navigation buttons
    document.getElementById("find-nav-prev").addEventListener("click", () => findNavigate(-1));
    document.getElementById("find-nav-next").addEventListener("click", () => findNavigate(1));
    document.getElementById("find-nav-close").addEventListener("click", clearFind);

    // T123: Reset color filter button
    document.getElementById("find-nav-filter-reset").addEventListener("click", () => {
        resetLegendColorFilter();
    });

    // Esc to close find
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && findState.active) {
            clearFind();
            e.preventDefault();
        }
        if (findState.active) {
            if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
                findNavigate(-1);
                e.preventDefault();
            }
            if (e.key === "ArrowRight" || e.key === "ArrowDown") {
                findNavigate(1);
                e.preventDefault();
            }
        }
    });

    /**
     * Main entry: find all instances of a legend item and highlight them.
     * Single-click replaces selection; Ctrl+click adds (handled by toggleFindSelection).
     */
    async function findEquipment(rowIndex, item) {
        // If clicking the same row again while find is active — clear
        if (findState.active && findState.selections.length === 1
            && findState.selections[0].rowIndex === rowIndex) {
            clearFind();
            return;
        }

        // Reset to single selection
        findState.selections = [];
        findState.currentIndex = 0;
        findState.active = true;

        // Highlight table row
        legendTbody.querySelectorAll("tr").forEach((r) => r.classList.remove("active-row", "find-selected"));
        const targetRow = legendTbody.querySelector(`tr[data-index="${rowIndex}"]`);
        if (targetRow) {
            targetRow.classList.add("active-row", "find-selected");
        }

        // T123: Apply color filter if item has a color
        applyLegendColorFilter(item);

        await addFindSelection(rowIndex, item);
    }

    /**
     * Ctrl+click: toggle a row in multi-selection mode.
     */
    async function toggleFindSelection(rowIndex, item) {
        // Check if already selected
        const existingIdx = findState.selections.findIndex((s) => s.rowIndex === rowIndex);
        if (existingIdx >= 0) {
            // Remove from selection
            findState.selections.splice(existingIdx, 1);
            const row = legendTbody.querySelector(`tr[data-index="${rowIndex}"]`);
            if (row) row.classList.remove("active-row", "find-selected");

            if (findState.selections.length === 0) {
                clearFind();
                return;
            }

            // Rebuild allPositions
            rebuildAllPositions();
            findState.currentIndex = 0;
            updateFindNavBar();
            drawFindOverlay();
            return;
        }

        // Add to selection
        if (!findState.active) {
            findState.active = true;
            findState.selections = [];
        }

        const row = legendTbody.querySelector(`tr[data-index="${rowIndex}"]`);
        if (row) row.classList.add("active-row", "find-selected");

        await addFindSelection(rowIndex, item);
    }

    /**
     * Add a single row to the current find selection, fetch positions from API.
     */
    async function addFindSelection(rowIndex, item) {
        // Show loading state
        findNav.classList.remove("hidden");
        const sym = item.symbol || "\u25CF";
        findNavSymbol.textContent = sym;
        findNavDesc.textContent = item.description || "";
        findNavStatus.innerHTML = '<span class="spinner" style="width:16px;height:16px;border-width:2px"></span>';
        findNavCounter.textContent = "...";

        try {
            const resp = await fetch(`/api/file/${FILE_ID}/find/${rowIndex}`);
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const data = await resp.json();

            findState.selections.push({ rowIndex, item, data });
            rebuildAllPositions();
            invalidateFindPixelMask();
            findState.currentIndex = 0;
            updateFindNavBar();
            drawFindOverlay();

            // Navigate to first position
            if (findState.allPositions.length > 0) {
                scrollToFindPosition(0);
            }
        } catch (err) {
            console.error("Find error:", err);
            findNavStatus.textContent = "Error";
            findNavStatus.className = "find-nav-status find-nav-error";
        }
    }

    /**
     * Rebuild flat allPositions array from all selections.
     */
    function rebuildAllPositions() {
        findState.allPositions = [];
        for (const sel of findState.selections) {
            if (sel.data && sel.data.positions) {
                for (const p of sel.data.positions) {
                    findState.allPositions.push({
                        ...p,
                        rowIndex: sel.rowIndex,
                        symbol: sel.data.symbol || "",
                        description: sel.data.description || "",
                        category: sel.data.category || "",
                    });
                }
            }
        }
        invalidateFindPixelMask();
    }

    /**
     * Update the navigation bar text & status.
     */
    function updateFindNavBar() {
        const total = findState.allPositions.length;

        // Update symbol/description for multi-select
        if (findState.selections.length === 1) {
            const sel = findState.selections[0];
            findNavSymbol.textContent = sel.data.symbol || "\u25CF";
            findNavDesc.textContent = sel.data.description || "";
        } else {
            const syms = findState.selections.map((s) => s.data.symbol || "\u25CF").join(", ");
            findNavSymbol.textContent = syms;
            findNavDesc.textContent = findState.selections.length + " elements";
        }

        if (total === 0) {
            findNavCounter.textContent = "0 / 0";
            // Determine message
            const hasSymbol = findState.selections.some((s) => s.data.symbol);
            const needsVisual = findState.selections.some((s) => s.data.needs_visual);
            if (needsVisual) {
                findNavStatus.textContent = "Requires visual search (Method D)";
                findNavStatus.className = "find-nav-status find-nav-warning";
            } else if (!hasSymbol) {
                findNavStatus.textContent = "No text marker (graphical symbol)";
                findNavStatus.className = "find-nav-status find-nav-warning";
            } else {
                findNavStatus.textContent = "Not found on drawing";
                findNavStatus.className = "find-nav-status find-nav-warning";
            }
        } else {
            findNavCounter.textContent = `${findState.currentIndex + 1} / ${total}`;
            const count = total;
            findNavStatus.textContent = `found: ${count}`;
            findNavStatus.className = "find-nav-status find-nav-ok";
        }
    }

    /**
     * Navigate to prev/next found position.
     */
    function findNavigate(delta) {
        if (findState.allPositions.length === 0) return;
        findState.currentIndex =
            (findState.currentIndex + delta + findState.allPositions.length)
            % findState.allPositions.length;

        updateFindNavBar();
        scrollToFindPosition(findState.currentIndex);
        drawFindOverlay(); // redraw to highlight current
    }

    /**
     * Scroll PDF viewport to a specific found position.
     */
    function scrollToFindPosition(idx) {
        if (!pdfImage.naturalWidth) return;
        const pos = findState.allPositions[idx];
        if (!pos) return;

        const pxPerPt = renderDPI / 72.0;
        const cx = pos.x * pxPerPt * zoomLevel;
        const cy = pos.y * pxPerPt * zoomLevel;

        pdfViewport.scrollTo({
            left: cx - pdfViewport.clientWidth / 2,
            top: cy - pdfViewport.clientHeight / 2,
            behavior: "smooth",
        });
    }

    /**
     * Draw the find overlay: dim layer + glowing markers.
     */
    function drawFindOverlay() {
        if (!pdfImage.naturalWidth) return;

        const w = pdfImage.naturalWidth;
        const h = pdfImage.naturalHeight;
        highlightCanvas.width = w;
        highlightCanvas.height = h;
        highlightCanvas.style.width = w + "px";
        highlightCanvas.style.height = h + "px";

        const ctx = highlightCanvas.getContext("2d");
        ctx.clearRect(0, 0, w, h);

        if (!findState.active || findState.allPositions.length === 0) return;

        const pxPerPt = renderDPI / 72.0;
        const manyPositions = findState.allPositions.length > 50;

        // 1. Dim layer (30% dark overlay)
        ctx.fillStyle = "rgba(0, 0, 0, 0.3)";
        ctx.fillRect(0, 0, w, h);

        // 1b. Blink-able pixel mask (equipment pixels around found markers).
        const maskCanvas = ensureFindPixelMask();
        if (maskCanvas) {
            const t = findState.pulsePhase / 60;
            const blinkAlpha = 0.35 + 0.35 * (0.5 + 0.5 * Math.sin(t * 2 * Math.PI));
            ctx.save();
            ctx.globalCompositeOperation = "lighter";
            ctx.globalAlpha = blinkAlpha;
            ctx.drawImage(maskCanvas, 0, 0);
            ctx.restore();
        }
        const renderClassicMarkers = !maskCanvas;

        // 2. Draw markers for each found position
        const catColors = {
            "default":    { fill: "rgba(74, 144, 217, 0.25)", stroke: "rgba(74, 144, 217, 0.80)" },
            "text":       { fill: "rgba(74, 144, 217, 0.25)", stroke: "rgba(74, 144, 217, 0.80)" },
            "visual":     { fill: "rgba(40, 167, 69, 0.25)",  stroke: "rgba(40, 167, 69, 0.80)" },
        };

        if (renderClassicMarkers) {
            findState.allPositions.forEach((pos, i) => {
            const px = pos.x * pxPerPt;
            const py = pos.y * pxPerPt;
            const radius = 18;

            const colors = catColors[pos.method] || catColors["default"];
            const isCurrent = (i === findState.currentIndex);

            // Clear a circle in the dim layer to reveal the PDF underneath
            ctx.save();
            ctx.globalCompositeOperation = "destination-out";
            ctx.beginPath();
            ctx.arc(px, py, radius + 12, 0, 2 * Math.PI);
            ctx.fill();
            ctx.restore();

            // Draw marker circle
            ctx.beginPath();
            ctx.arc(px, py, radius, 0, 2 * Math.PI);
            ctx.fillStyle = colors.fill;
            ctx.fill();
            ctx.strokeStyle = isCurrent ? "#fff" : colors.stroke;
            ctx.lineWidth = isCurrent ? 3 : 2;
            ctx.stroke();

            // Draw glow for current marker
            if (isCurrent && !manyPositions) {
                ctx.save();
                ctx.shadowColor = colors.stroke;
                ctx.shadowBlur = 20;
                ctx.beginPath();
                ctx.arc(px, py, radius + 2, 0, 2 * Math.PI);
                ctx.strokeStyle = colors.stroke;
                ctx.lineWidth = 3;
                ctx.stroke();
                ctx.restore();
            }

            // Draw label with symbol
            const label = pos.symbol || (i + 1).toString();
            ctx.font = "bold 11px " + getComputedStyle(document.body).getPropertyValue("--mono");
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillStyle = isCurrent ? "#fff" : colors.stroke.replace("0.80", "1.0");
            ctx.fillText(label, px, py);
            });
        } else {
            drawCurrentFindCursor(ctx, pxPerPt);
        }

        // Start pulse animation if not too many positions and not already running
        if (!manyPositions && !findState.pulseTimer) {
            startPulseAnimation();
        }
        if (manyPositions && findState.pulseTimer) {
            stopPulseAnimation();
        }
    }

    /**
     * Pulsating glow animation for current marker.
     */
    function startPulseAnimation() {
        findState.pulsePhase = 0;
        findState.pulseTimer = setInterval(() => {
            if (!findState.active || !pdfImage.naturalWidth) {
                stopPulseAnimation();
                return;
            }

            findState.pulsePhase = (findState.pulsePhase + 1) % 60; // ~2s at 30fps
            const t = findState.pulsePhase / 60;
            const scale = 1.0 + 0.15 * Math.sin(t * 2 * Math.PI);

            // Redraw just the current marker with pulse
            drawFindOverlayFrame(scale);
        }, 33); // ~30fps
    }

    function stopPulseAnimation() {
        if (findState.pulseTimer) {
            clearInterval(findState.pulseTimer);
            findState.pulseTimer = null;
        }
    }

    /**
     * Draw a single animation frame with pulse scale for current marker.
     */
    function drawFindOverlayFrame(pulseScale) {
        if (!pdfImage.naturalWidth || findState.allPositions.length === 0) return;

        const w = pdfImage.naturalWidth;
        const h = pdfImage.naturalHeight;
        const ctx = highlightCanvas.getContext("2d");
        ctx.clearRect(0, 0, w, h);

        const pxPerPt = renderDPI / 72.0;

        // Dim layer
        ctx.fillStyle = "rgba(0, 0, 0, 0.3)";
        ctx.fillRect(0, 0, w, h);

        // Draw blinking equipment-pixel mask.
        const maskCanvas = ensureFindPixelMask();
        if (maskCanvas) {
            const blinkAlpha = 0.25 + 0.55 * (0.5 + 0.5 * Math.sin((findState.pulsePhase / 60) * 2 * Math.PI));
            ctx.save();
            ctx.globalCompositeOperation = "lighter";
            ctx.globalAlpha = blinkAlpha;
            ctx.drawImage(maskCanvas, 0, 0);
            ctx.restore();
        }
        const renderClassicMarkers = !maskCanvas;

        const catColors = {
            "default":    { fill: "rgba(74, 144, 217, 0.25)", stroke: "rgba(74, 144, 217, 0.80)" },
            "text":       { fill: "rgba(74, 144, 217, 0.25)", stroke: "rgba(74, 144, 217, 0.80)" },
            "visual":     { fill: "rgba(40, 167, 69, 0.25)",  stroke: "rgba(40, 167, 69, 0.80)" },
        };

        if (renderClassicMarkers) {
            findState.allPositions.forEach((pos, i) => {
            const px = pos.x * pxPerPt;
            const py = pos.y * pxPerPt;
            const isCurrent = (i === findState.currentIndex);
            const radius = isCurrent ? 18 * pulseScale : 18;
            const clearRadius = isCurrent ? (18 + 12) * pulseScale : 18 + 12;

            const colors = catColors[pos.method] || catColors["default"];

            // Clear dim layer around marker
            ctx.save();
            ctx.globalCompositeOperation = "destination-out";
            ctx.beginPath();
            ctx.arc(px, py, clearRadius, 0, 2 * Math.PI);
            ctx.fill();
            ctx.restore();

            // Draw marker
            ctx.beginPath();
            ctx.arc(px, py, radius, 0, 2 * Math.PI);
            ctx.fillStyle = colors.fill;
            ctx.fill();
            ctx.strokeStyle = isCurrent ? "#fff" : colors.stroke;
            ctx.lineWidth = isCurrent ? 3 : 2;
            ctx.stroke();

            // Glow for current
            if (isCurrent) {
                ctx.save();
                ctx.shadowColor = colors.stroke;
                ctx.shadowBlur = 15 + 10 * Math.sin((findState.pulsePhase / 60) * 2 * Math.PI);
                ctx.beginPath();
                ctx.arc(px, py, radius + 2, 0, 2 * Math.PI);
                ctx.strokeStyle = colors.stroke;
                ctx.lineWidth = 2;
                ctx.stroke();
                ctx.restore();
            }

            // Label
            const label = pos.symbol || (i + 1).toString();
            ctx.font = "bold 11px " + getComputedStyle(document.body).getPropertyValue("--mono");
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillStyle = isCurrent ? "#fff" : colors.stroke.replace("0.80", "1.0");
            ctx.fillText(label, px, py);
            });
        } else {
            drawCurrentFindCursor(ctx, pxPerPt);
        }
    }

    /**
     * T123: Apply color filter to PDF background based on legend item's color.
     * Fetches colors data if needed, finds matching hex values, and loads filtered render.
     */
    async function applyLegendColorFilter(item) {
        const colorName = item.color;
        const filterResetBtn = document.getElementById("find-nav-filter-reset");

        if (!colorName || colorName === "black" || !COLOR_NAME_TO_HEX[colorName]) {
            // No meaningful color to filter by (black is the default, grey is ambiguous)
            legendColorFilterActive = false;
            legendColorFilterHex = null;
            legendColorFilterName = "";
            if (filterResetBtn) filterResetBtn.classList.add("hidden");
            return;
        }

        // We need to find the actual hex values from the PDF's color palette
        // that match this color category, so render_filtered shows the right elements
        try {
            // Load color palette if not already loaded
            if (!colorData) {
                const resp = await fetch(`/api/file/${FILE_ID}/colors?page=${currentPage}`);
                if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
                colorData = await resp.json();
                colorsLoaded = true;
            }

            // Find hex values matching the color name using the same heuristics
            // as the existing _isRedHex / _isBlueHex helpers
            const matchFn = {
                "red": _isRedHex,
                "blue": _isBlueHex,
                "green": (hex) => {
                    const r = parseInt(hex.substring(0, 2), 16);
                    const g = parseInt(hex.substring(2, 4), 16);
                    const b = parseInt(hex.substring(4, 6), 16);
                    return g > 150 && r < 100 && b < 100;
                },
            };

            const matcher = matchFn[colorName];
            let hexes = [];

            if (matcher && colorData.colors) {
                hexes = colorData.colors
                    .filter((c) => matcher(c.hex))
                    .map((c) => c.hex);
            }

            // If no palette matches found, the PDF may not have colored elements
            // (e.g. all black/grey). In that case, skip filtering entirely —
            // using hardcoded hex like FF0000 on a black-only PDF would show nothing.
            if (hexes.length === 0) {
                legendColorFilterActive = false;
                legendColorFilterHex = null;
                legendColorFilterName = "";
                if (filterResetBtn) filterResetBtn.classList.add("hidden");
                return;
            }

            // Apply filter
            legendColorFilterActive = true;
            legendColorFilterHex = hexes;
            legendColorFilterName = colorName;

            // Show reset button with color badge
            if (filterResetBtn) {
                filterResetBtn.classList.remove("hidden");
            }

            // Add a color filter badge to the find nav status area
            const statusEl = document.getElementById("find-nav-status");
            if (statusEl) {
                const cssColor = COLOR_NAME_TO_CSS[colorName] || "#999";
                statusEl.innerHTML += ` <span class="find-filter-badge"><span class="legend-color-dot" style="background:${cssColor}"></span>${colorName}</span>`;
            }

            // Load the filtered page render
            await loadFilteredPage(hexes);
        } catch (err) {
            console.error("Color filter error:", err);
            legendColorFilterActive = false;
        }
    }

    /**
     * T123: Remove color filter and restore normal PDF render.
     */
    function resetLegendColorFilter() {
        legendColorFilterActive = false;
        legendColorFilterHex = null;
        legendColorFilterName = "";

        const filterResetBtn = document.getElementById("find-nav-filter-reset");
        if (filterResetBtn) filterResetBtn.classList.add("hidden");

        // Reload unfiltered page
        loadPage(currentPage);
    }

    /**
     * Clear all find state and restore normal view.
     */
    function clearFind() {
        findState.active = false;
        findState.selections = [];
        findState.currentIndex = 0;
        findState.allPositions = [];
        invalidateFindPixelMask();
        stopPulseAnimation();

        // Hide nav bar
        findNav.classList.add("hidden");

        // Remove active state from legend rows
        legendTbody.querySelectorAll("tr").forEach((r) => r.classList.remove("active-row", "find-selected"));

        // T123: Restore unfiltered page if color filter was active
        if (legendColorFilterActive) {
            resetLegendColorFilter();
        }

        // Redraw canvas to clear overlay
        drawOverlay();
    }

    // Keep the old scrollToItem as a simpler alternative when find is not needed
    // (the new findEquipment replaces its role for legend row clicks)

    // ================================================================
    // CABLE HIGHLIGHTING (T107)
    // ================================================================

    let cableData = null;  // cached cable analysis results
    let cableFilterMode = "type";  // 'type', 'group', 'route'
    let cableActiveFilters = { red: true, blue: true };
    let cableActiveGroups = new Set();   // group_full strings to show
    let cableActiveRoutes = new Set();   // route IDs to show
    let cableOverlayActive = false;
    let cableHoveredRoute = null;

    // DOM refs
    const cableLoading = document.getElementById("cable-loading");
    const cableStats = document.getElementById("cable-stats");
    const cableStatsBadges = document.getElementById("cable-stats-badges");
    const cableFilterBar = document.getElementById("cable-filter-bar");
    const cableFilterType = document.getElementById("cable-filter-type");
    const cableFilterGroup = document.getElementById("cable-filter-group");
    const cableFilterRoute = document.getElementById("cable-filter-route");
    const cableGroupList = document.getElementById("cable-group-list");
    const cableRouteList = document.getElementById("cable-route-list");
    const cableAnnotSection = document.getElementById("cable-annot-section");
    const cableAnnotTbody = document.getElementById("cable-annot-tbody");
    const cableTooltip = document.getElementById("cable-tooltip");

    // Run cable analysis
    document.getElementById("btn-run-cables").addEventListener("click", runCableAnalysis);

    async function runCableAnalysis() {
        cableLoading.classList.remove("hidden");
        cableStats.classList.add("hidden");
        cableFilterBar.classList.add("hidden");
        cableAnnotSection.classList.add("hidden");

        try {
            const resp = await fetch(`/api/file/${FILE_ID}/cables`);
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            cableData = await resp.json();

            renderCableStats();
            renderCableFilters();
            renderCableAnnotations();
            cableOverlayActive = true;
            drawCableOverlay();
        } catch (err) {
            console.error("Cable analysis error:", err);
            cableStatsBadges.innerHTML = `<span class="status-badge badge-red">Error: ${escapeHtml(err.message)}</span>`;
            cableStats.classList.remove("hidden");
        } finally {
            cableLoading.classList.add("hidden");
        }
    }

    function renderCableStats() {
        if (!cableData) return;
        const badges = [];
        badges.push(`<span class="status-badge badge-blue">${cableData.elapsed_s}s</span>`);
        badges.push(`<span class="status-badge badge-red">${cableData.red_segments} red</span>`);
        badges.push(`<span class="status-badge badge-blue">${cableData.blue_segments} blue</span>`);
        badges.push(`<span class="status-badge badge-green">${cableData.total_routes} routes</span>`);
        if (cableData.scale) {
            badges.push(`<span class="status-badge badge-yellow">${cableData.scale.source}</span>`);
        }
        cableStatsBadges.innerHTML = badges.join("");
        cableStats.classList.remove("hidden");
    }

    function renderCableFilters() {
        if (!cableData) return;

        // Update type counts
        document.getElementById("cable-red-count").textContent = cableData.red_segments;
        document.getElementById("cable-blue-count").textContent = cableData.blue_segments;

        // Build group list
        const groups = new Map();
        for (const a of cableData.annotations) {
            if (!a.group_full) continue;
            if (!groups.has(a.group_full)) {
                groups.set(a.group_full, { color: a.color, count: 0 });
            }
            groups.get(a.group_full).count++;
        }
        cableGroupList.innerHTML = "";
        cableActiveGroups.clear();
        for (const [gf, info] of groups) {
            cableActiveGroups.add(gf);
            const label = document.createElement("label");
            label.className = "cable-filter-item";
            const dotColor = info.color === "red" ? "#dc3545" : (info.color === "blue" ? "#4a90d9" : "#888");
            label.innerHTML = `
                <input type="checkbox" class="cable-group-chk" data-group="${escapeHtml(gf)}" checked />
                <span class="cable-type-dot" style="background:${dotColor}"></span>
                <span>${escapeHtml(gf)}</span>
                <span class="cable-type-count">${info.count}</span>
            `;
            cableGroupList.appendChild(label);
        }

        // Build route list (top 20 by length)
        const sortedRoutes = [...cableData.routes]
            .sort((a, b) => b.length_m - a.length_m)
            .slice(0, 30);
        cableRouteList.innerHTML = "";
        cableActiveRoutes.clear();
        for (const r of sortedRoutes) {
            cableActiveRoutes.add(r.id);
            const label = document.createElement("label");
            label.className = "cable-filter-item";
            const dotColor = r.color === "red" ? "#dc3545" : "#4a90d9";
            label.innerHTML = `
                <input type="checkbox" class="cable-route-chk" data-route="${r.id}" checked />
                <span class="cable-type-dot" style="background:${dotColor}"></span>
                <span>Route #${r.id} (${r.segment_count} seg)</span>
                <span class="cable-type-count">${r.length_m}m</span>
            `;
            cableRouteList.appendChild(label);
        }

        // Bind filter events
        cableFilterType.querySelectorAll(".cable-type-chk").forEach((cb) => {
            cb.addEventListener("change", () => {
                cableActiveFilters[cb.dataset.color] = cb.checked;
                drawCableOverlay();
            });
        });
        cableGroupList.querySelectorAll(".cable-group-chk").forEach((cb) => {
            cb.addEventListener("change", () => {
                if (cb.checked) cableActiveGroups.add(cb.dataset.group);
                else cableActiveGroups.delete(cb.dataset.group);
                drawCableOverlay();
            });
        });
        cableRouteList.querySelectorAll(".cable-route-chk").forEach((cb) => {
            cb.addEventListener("change", () => {
                const rid = parseInt(cb.dataset.route);
                if (cb.checked) cableActiveRoutes.add(rid);
                else cableActiveRoutes.delete(rid);
                drawCableOverlay();
            });
        });

        cableFilterBar.classList.remove("hidden");
    }

    // Filter mode switching
    document.querySelectorAll(".cable-mode-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
            document.querySelectorAll(".cable-mode-btn").forEach((b) => b.classList.remove("active"));
            btn.classList.add("active");
            cableFilterMode = btn.dataset.mode;

            // Show/hide panels
            cableFilterType.classList.toggle("hidden", cableFilterMode !== "type");
            cableFilterGroup.classList.toggle("hidden", cableFilterMode !== "group");
            cableFilterRoute.classList.toggle("hidden", cableFilterMode !== "route");

            drawCableOverlay();
        });
    });

    function renderCableAnnotations() {
        if (!cableData || !cableData.annotations.length) return;
        cableAnnotTbody.innerHTML = "";

        for (const a of cableData.annotations) {
            const tr = document.createElement("tr");
            const dotColor = a.color === "red" ? "#dc3545" : (a.color === "blue" ? "#4a90d9" : "#888");
            tr.innerHTML = `
                <td>${escapeHtml(a.group_full || "\u2014")}</td>
                <td style="font-family:var(--mono);font-size:0.78rem">${escapeHtml(a.cross_section)}</td>
                <td style="font-size:0.75rem">${escapeHtml(a.cable_type || "\u2014")}</td>
                <td><span class="cable-type-dot" style="background:${dotColor}"></span></td>
            `;
            tr.addEventListener("click", () => {
                scrollToCableAnnotation(a);
            });
            cableAnnotTbody.appendChild(tr);
        }
        cableAnnotSection.classList.remove("hidden");
    }

    function scrollToCableAnnotation(annot) {
        if (!pdfImage.naturalWidth) return;
        const pxPerPt = renderDPI / 72.0;
        pdfViewport.scrollTo({
            left: annot.x * pxPerPt * zoomLevel - pdfViewport.clientWidth / 2,
            top: annot.y * pxPerPt * zoomLevel - pdfViewport.clientHeight / 2,
            behavior: "smooth",
        });
    }

    /**
     * Check if a segment should be visible based on current filter mode.
     */
    function isSegmentVisible(seg) {
        if (cableFilterMode === "type") {
            return cableActiveFilters[seg.color] === true;
        }
        if (cableFilterMode === "group") {
            // Find annotations near this segment
            if (!cableData.annotations.length) return true;
            // Match by color at minimum
            for (const a of cableData.annotations) {
                if (cableActiveGroups.has(a.group_full) && a.color === seg.color) {
                    return true;
                }
            }
            return false;
        }
        if (cableFilterMode === "route") {
            // Check if segment belongs to any active route
            // We match by checking if segment endpoints are within route bbox
            for (const r of cableData.routes) {
                if (!cableActiveRoutes.has(r.id)) continue;
                if (r.color !== seg.color) continue;
                const b = r.bbox;
                const mx = (seg.x0 + seg.x1) / 2;
                const my = (seg.y0 + seg.y1) / 2;
                if (mx >= b.x0 - 5 && mx <= b.x1 + 5 && my >= b.y0 - 5 && my <= b.y1 + 5) {
                    return true;
                }
            }
            return false;
        }
        return true;
    }

    /**
     * Draw cable overlay on canvas.
     */
    function drawCableOverlay() {
        if (!cableData || !pdfImage.naturalWidth || !cableOverlayActive) return;

        // Clear any find overlay first
        if (findState.active) {
            clearFind();
        }

        const w = pdfImage.naturalWidth;
        const h = pdfImage.naturalHeight;
        highlightCanvas.width = w;
        highlightCanvas.height = h;
        highlightCanvas.style.width = w + "px";
        highlightCanvas.style.height = h + "px";

        const ctx = highlightCanvas.getContext("2d");
        ctx.clearRect(0, 0, w, h);

        const pxPerPt = renderDPI / 72.0;

        // Dim layer
        ctx.fillStyle = "rgba(0, 0, 0, 0.25)";
        ctx.fillRect(0, 0, w, h);

        // Draw cable segments
        const colors = {
            red: { line: "rgba(220, 53, 69, 0.85)", glow: "rgba(220, 53, 69, 0.35)" },
            blue: { line: "rgba(74, 144, 217, 0.85)", glow: "rgba(74, 144, 217, 0.35)" },
        };

        // Draw glow first (wider, translucent)
        ctx.lineCap = "round";
        for (const seg of cableData.segments) {
            if (!isSegmentVisible(seg)) continue;
            const c = colors[seg.color] || colors.blue;
            ctx.beginPath();
            ctx.moveTo(seg.x0 * pxPerPt, seg.y0 * pxPerPt);
            ctx.lineTo(seg.x1 * pxPerPt, seg.y1 * pxPerPt);
            ctx.strokeStyle = c.glow;
            ctx.lineWidth = 6;
            ctx.stroke();
        }

        // Draw actual cable lines
        for (const seg of cableData.segments) {
            if (!isSegmentVisible(seg)) continue;
            const c = colors[seg.color] || colors.blue;
            ctx.beginPath();
            ctx.moveTo(seg.x0 * pxPerPt, seg.y0 * pxPerPt);
            ctx.lineTo(seg.x1 * pxPerPt, seg.y1 * pxPerPt);
            ctx.strokeStyle = c.line;
            ctx.lineWidth = 2;
            ctx.stroke();
        }

        // Draw annotation markers (text positions)
        for (const a of cableData.annotations) {
            // Filter by mode
            if (cableFilterMode === "type" && !cableActiveFilters[a.color]) continue;
            if (cableFilterMode === "group" && !cableActiveGroups.has(a.group_full)) continue;

            const px = a.x * pxPerPt;
            const py = a.y * pxPerPt;

            // Clear dim around annotation
            ctx.save();
            ctx.globalCompositeOperation = "destination-out";
            ctx.beginPath();
            ctx.arc(px, py, 22, 0, 2 * Math.PI);
            ctx.fill();
            ctx.restore();

            // Draw annotation dot
            const dotColor = a.color === "red" ? "rgba(220, 53, 69, 0.9)" : "rgba(74, 144, 217, 0.9)";
            ctx.beginPath();
            ctx.arc(px, py, 5, 0, 2 * Math.PI);
            ctx.fillStyle = dotColor;
            ctx.fill();
            ctx.strokeStyle = "#fff";
            ctx.lineWidth = 1.5;
            ctx.stroke();

            // Label
            if (a.group_full) {
                ctx.font = "bold 9px " + getComputedStyle(document.body).getPropertyValue("--sans");
                ctx.textAlign = "left";
                ctx.textBaseline = "bottom";
                ctx.fillStyle = "#fff";

                // Background for label
                const text = a.group_full;
                const tm = ctx.measureText(text);
                ctx.fillStyle = a.color === "red" ? "rgba(180,30,40,0.8)" : "rgba(50,110,180,0.8)";
                ctx.fillRect(px + 7, py - 14, tm.width + 6, 14);
                ctx.fillStyle = "#fff";
                ctx.fillText(text, px + 10, py - 2);
            }
        }
    }

    // Tooltip on mouse move over PDF viewport (when cable overlay is active)
    pdfViewport.addEventListener("mousemove", (e) => {
        if (!cableOverlayActive || !cableData || !pdfImage.naturalWidth) {
            cableTooltip.classList.add("hidden");
            return;
        }

        const pxPerPt = renderDPI / 72.0;
        const rect = pdfImage.getBoundingClientRect();
        const imgX = (e.clientX - rect.left) / zoomLevel;
        const imgY = (e.clientY - rect.top) / zoomLevel;
        const pdfX = imgX / pxPerPt;
        const pdfY = imgY / pxPerPt;

        // Find nearest annotation
        let nearest = null;
        let nearestDist = 30; // max distance in PDF pts
        for (const a of cableData.annotations) {
            const d = Math.sqrt((a.x - pdfX) ** 2 + (a.y - pdfY) ** 2);
            if (d < nearestDist) {
                nearestDist = d;
                nearest = a;
            }
        }

        if (nearest) {
            const colorLabel = nearest.color === "red" ? "Аварийное" : (nearest.color === "blue" ? "Рабочее" : "");
            const lengthStr = nearest.length_m ? `${nearest.length_m}m` : "";
            cableTooltip.innerHTML = `
                <div class="cable-tt-group">${escapeHtml(nearest.group_full || "—")}</div>
                <div class="cable-tt-detail">${escapeHtml(nearest.cross_section)} ${escapeHtml(nearest.cable_type || "")}</div>
                <div class="cable-tt-meta">${colorLabel} ${lengthStr}</div>
            `;
            cableTooltip.style.left = (e.clientX + 12) + "px";
            cableTooltip.style.top = (e.clientY - 10) + "px";
            cableTooltip.classList.remove("hidden");
        } else {
            cableTooltip.classList.add("hidden");
        }
    });

    // Hide tooltip when leaving viewport
    pdfViewport.addEventListener("mouseleave", () => {
        cableTooltip.classList.add("hidden");
    });

    // Clear cable overlay when switching tabs away from cables
    const origTabHandler = document.querySelectorAll(".tab-btn");
    origTabHandler.forEach((btn) => {
        btn.addEventListener("click", () => {
            if (btn.dataset.tab !== "tab-cables" && cableOverlayActive) {
                cableOverlayActive = false;
                drawOverlay(); // redraw base overlay
                cableTooltip.classList.add("hidden");
            }
            if (btn.dataset.tab === "tab-cables" && cableData) {
                cableOverlayActive = true;
                drawCableOverlay();
            }
        });
    });

})();
