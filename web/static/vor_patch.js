// vor_patch.js — клиентский патч модалки ВОР.
// Чинит кнопку «Скачать Excel» и добавляет «Скачать DOCX».
// Подключается стартером run_web.py через инжекцию <script> в /.
(function () {
    'use strict';

    function escHtml(t) {
        var d = document.createElement('div');
        d.textContent = t == null ? '' : String(t);
        return d.innerHTML;
    }

    // Найти контейнер кнопок и добавить DOCX, если его ещё нет.
    function ensureDocxButton() {
        var actions = document.querySelector('.vor-actions');
        if (!actions) return null;
        var btnD = document.getElementById('btn-vor-docx');
        if (!btnD) {
            btnD = document.createElement('button');
            btnD.id = 'btn-vor-docx';
            btnD.className = 'btn btn-primary';
            btnD.style.marginLeft = '8px';
            btnD.innerHTML = '\u{1F4C4} Скачать DOCX';
            actions.appendChild(btnD);
        }
        return btnD;
    }

    // Послать агрегат на сервер для кэширования.
    function cacheAggregated(folderId, aggregated) {
        try {
            fetch('/api/folder/' + folderId + '/cache_vor', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ aggregated: aggregated || [] })
            }).catch(function () {});
        } catch (e) { /* ignore */ }
    }

    // Полная замена startVor: чиним XLSX-кнопку и добавляем DOCX.
    function patchedStartVor(folderId, folderName) {
        var modal = document.getElementById('vor-modal');
        var fill = document.getElementById('vor-fill');
        var ptext = document.getElementById('vor-ptext');
        var log = document.getElementById('vor-log');
        var results = document.getElementById('vor-results');
        var tbody = document.getElementById('vor-tbody');

        document.getElementById('vor-folder-name').textContent = folderName;
        modal.classList.remove('hidden');
        results.classList.add('hidden');
        log.innerHTML = '';
        fill.style.width = '0%';
        ptext.textContent = 'Подготовка...';

        var btnX = document.getElementById('btn-vor-xlsx');
        var btnD = ensureDocxButton();

        // Сразу навешиваем onclick — чтобы кнопки работали даже если done не дойдёт.
        if (btnX) {
            btnX.disabled = true;
            btnX.onclick = function () {
                window.open('/api/folder/' + folderId + '/export_xlsx_v2', '_blank');
            };
        }
        if (btnD) {
            btnD.disabled = true;
            btnD.onclick = function () {
                window.open('/api/folder/' + folderId + '/export_docx', '_blank');
            };
        }

        var src = new EventSource('/api/folder/' + folderId + '/process');

        src.addEventListener('start', function (e) {
            var d = JSON.parse(e.data);
            ptext.textContent = '0 / ' + d.total + ' файлов...';
        });

        src.addEventListener('progress', function (e) {
            var d = JSON.parse(e.data);
            var pct = Math.round((d.current / d.total) * 100);
            fill.style.width = pct + '%';
            ptext.textContent = d.current + ' / ' + d.total + ': ' + d.filename;
        });

        src.addEventListener('file_done', function (e) {
            var d = JSON.parse(e.data);
            var icon = d.status === 'ok' ? '\u2713' : '\u2717';
            var cls = d.status === 'ok' ? 'log-ok' : 'log-err';
            log.innerHTML += '<div class="log-entry ' + cls + '">' + icon + ' ' + escHtml(d.filename) + ' (' + d.items + ' эл.)</div>';
            log.scrollTop = log.scrollHeight;
        });

        src.addEventListener('done', function (e) {
            src.close();
            var d = JSON.parse(e.data);
            fill.style.width = '100%';
            ptext.textContent = 'Готово: ' + d.files_processed + ' / ' + d.total_files + ' файлов';

            tbody.innerHTML = '';
            (d.aggregated || []).forEach(function (item) {
                var tr = document.createElement('tr');
                tr.innerHTML = '<td>' + item.row + '</td><td>' + escHtml(item.name) + '</td><td>' + item.unit + '</td><td>' + item.total + '</td><td class="mono">' + escHtml(item.formula) + '</td><td>' + escHtml(item.drawing_refs) + '</td>';
                tbody.appendChild(tr);
            });
            results.classList.remove('hidden');

            // Кладём агрегат в серверный кэш — чтобы XLSX/DOCX отдавались мгновенно.
            cacheAggregated(folderId, d.aggregated || []);

            if (btnX) btnX.disabled = false;
            if (btnD) btnD.disabled = false;
        });

        src.onerror = function () {
            src.close();
            ptext.textContent = 'Ошибка соединения';
            // Даже при ошибке оставляем кнопки активными (бэк сам пересчитает по запросу).
            if (btnX) btnX.disabled = false;
            if (btnD) btnD.disabled = false;
        };
    }

    // Подменяем глобальный startVor сразу — onclick кнопок «✏ ВОР» вызывает window.startVor.
    function applyPatch() {
        ensureDocxButton();
        window.startVor = patchedStartVor;
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', applyPatch);
    } else {
        applyPatch();
    }
})();
