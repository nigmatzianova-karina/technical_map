(function() {
    'use strict';

    const escapeHtml = window.escapeHtml || function(text) {
        if (typeof text !== 'string') return text;
        const d = document.createElement('div');
        d.textContent = text;
        return d.innerHTML;
    };
    const showStatus = window.showStatus || function(msg, type = 'info', duration = 5000) {
        let el = document.getElementById('status-message');
        if (!el) {
            el = document.createElement('div');
            el.id = 'status-message';
            document.body.appendChild(el);
        }
        el.textContent = msg;
        el.className = type;
        el.style.display = 'block';
        if (duration > 0 && type === 'success') {
            setTimeout(() => { el.style.display = 'none'; }, duration);
        }
    };
    const formatFileSize = window.formatFileSize || function(bytes) {
        if (bytes === 0) return '0 Б';
        const k = 1024;
        const sizes = ['Б', 'КБ', 'МБ', 'ГБ'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    };

    const HISTORY_KEY = 'pdf_parser_history';
    let selectedFile = null;
    let lastResult = null;

    window.addEventListener('DOMContentLoaded', () => {
        initFileUpload();
        document.getElementById('parseBtn').addEventListener('click', handleParse);
        document.getElementById('downloadTextBtn').addEventListener('click', downloadText);
        document.getElementById('downloadExcelBtn').addEventListener('click', downloadTables);
        document.getElementById('historySelect').addEventListener('change', onHistorySelect);
        document.getElementById('clearHistoryBtn').addEventListener('click', clearHistory);
        loadHistoryDropdown();
    });

    function initFileUpload() {
        const input = document.getElementById('fileInput');
        const dropZone = document.getElementById('fileDropZone');
        if (!input || !dropZone) return;
        input.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                selectedFile = e.target.files[0];
                document.getElementById('fileName').textContent =
                    `${selectedFile.name} (${formatFileSize(selectedFile.size)})`;
                dropZone.classList.add('has-file');
            }
        });
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            if (e.dataTransfer.files[0]) {
                selectedFile = e.dataTransfer.files[0];
                document.getElementById('fileName').textContent =
                    `${selectedFile.name} (${formatFileSize(selectedFile.size)})`;
                dropZone.classList.add('has-file');
            }
        });
    }

    async function handleParse() {
        if (!selectedFile) {
            showStatus('Выберите файл', 'error');
            return;
        }
        const indicator = document.getElementById('loadingIndicator');
        indicator.style.display = 'inline';
        try {
            const result = await parseDocument(selectedFile);
            displayResult(result);
            addToHistory(result);
            showStatus('Файл успешно распарсен', 'success');
        } catch (e) {
            showStatus(`Ошибка: ${e.message}`, 'error');
        } finally {
            indicator.style.display = 'none';
        }
    }

    async function parseDocument(file) {
        const formData = new FormData();
        formData.append('file', file);
        const response = await fetch('/pdf_parser/api/parse', {
            method: 'POST',
            body: formData
        });
        if (!response.ok) {
            const err = await response.json().catch(() => ({ detail: 'Ошибка сервера' }));
            throw new Error(err.detail || 'Ошибка сервера');
        }
        const json = await response.json();
        if (!json.success) throw new Error(json.error || 'Неизвестная ошибка');
        lastResult = {
            ...json.data,
            xlsx_file: json.xlsx_file,
            xlsx_filename: json.xlsx_filename
        };
        return lastResult;
    }

    function displayResult(data) {
        const section = document.getElementById('resultsSection');
        // section.classList.add('visible');
        const container = document.getElementById('parsedContent');
        container.innerHTML = '';

        if (data.pages_text && data.pages_text.length) {
            const textBlock = document.createElement('div');
            textBlock.className = 'parsed-text-block';
            let combinedText = data.pages_text
                .map((txt, idx) => `=== Страница/Параграф ${idx+1} ===\n${txt}`)
                .join('\n\n');
            textBlock.innerHTML = `<pre style="white-space:pre-wrap; font-family: inherit; margin:0;">${escapeHtml(combinedText)}</pre>`;
            container.appendChild(textBlock);
        }

        if (data.tables && data.tables.length) {
            data.tables.forEach((table, tIdx) => {
                if (!table.length) return;
                const wrapper = document.createElement('div');
                wrapper.className = 'ai-text-result';
                wrapper.innerHTML = `<h3>Таблица ${tIdx+1}</h3>`;
                const htmlTable = '<div class="table-wrapper" style="margin-top:8px;"><table class="result-table"><thead><tr>' +
                    table[0].map(h => `<th>${escapeHtml(h)}</th>`).join('') +
                    '</tr></thead><tbody>' +
                    table.slice(1).map(row => '<tr>' + row.map(c => `<td>${escapeHtml(c)}</td>`).join('') + '</tr>').join('') +
                    '</tbody></table></div>';
                wrapper.innerHTML += htmlTable;
                container.appendChild(wrapper);
            });
        }

        document.getElementById('downloadTextBtn').style.display = data.pages_text?.length ? 'inline-flex' : 'none';
        document.getElementById('downloadExcelBtn').style.display = data.tables?.length ? 'inline-flex' : 'none';
    }

    function downloadText() {
        if (!lastResult?.pages_text) return;
        const text = lastResult.pages_text.join('\n\n');
        const blob = new Blob([text], { type: 'text/plain;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'parsed_text.txt';
        a.click();
        URL.revokeObjectURL(url);
    }

    function downloadTables() {
        if (!lastResult || !lastResult.tables || lastResult.tables.length === 0) {
            showStatus('Нет данных для скачивания', 'error');
            return;
        }

        if (lastResult.xlsx_file) {
            const link = document.createElement('a');
            link.href = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${lastResult.xlsx_file}`;
            link.download = lastResult.xlsx_filename || 'parsed_tables.xlsx';
            link.click();
            return;
        }

        const wb = XLSX.utils.book_new();
        lastResult.tables.forEach((table, idx) => {
            if (!table.length) return;
            const ws = XLSX.utils.aoa_to_sheet(table);
            XLSX.utils.book_append_sheet(wb, ws, `Таблица ${idx + 1}`);
        });

        const baseName = selectedFile
            ? selectedFile.name.replace(/\.(pdf|docx)$/i, '')
            : 'parsed_tables';
        const filename = baseName + '.xlsx';

        XLSX.writeFile(wb, filename);
    }

    function loadHistoryDropdown() {
        const select = document.getElementById('historySelect');
        const history = getHistory();
        select.innerHTML = '<option value="">— Выбрать предыдущий результат —</option>';
        history.forEach((entry, idx) => {
            const opt = document.createElement('option');
            opt.value = idx;
            opt.textContent = `${entry.filename} (${new Date(entry.timestamp).toLocaleString()})`;
            select.appendChild(opt);
        });
    }

    function getHistory() {
        try {
            return JSON.parse(localStorage.getItem(HISTORY_KEY)) || [];
        } catch { return []; }
    }

    function saveHistory(history) {
        localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
    }

    function addToHistory(result) {
        const history = getHistory();
        history.unshift({
            filename: selectedFile.name,
            timestamp: Date.now(),
            data: result
        });
        if (history.length > 20) history.pop();
        saveHistory(history);
        loadHistoryDropdown();
    }

    function onHistorySelect(e) {
        const idx = e.target.value;
        if (idx === '') return;
        const history = getHistory();
        const entry = history[parseInt(idx)];
        if (entry) {
            lastResult = entry.data;
            selectedFile = { name: entry.filename };
            displayResult(entry.data);
        }
    }

    function clearHistory() {
        if (confirm('Удалить всю историю парсинга?')) {
            localStorage.removeItem(HISTORY_KEY);
            loadHistoryDropdown();
        }
    }
})();