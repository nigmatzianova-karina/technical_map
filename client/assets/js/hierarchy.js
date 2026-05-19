let globalData = [];
let globalHeaders = [];
let currentEditRow = -1;
let currentEditCol = -1;
const HISTORY_KEY = 'hierarchy_history';


document.getElementById('fileInput').addEventListener('change', (e) => {
    document.getElementById('fileName').textContent = e.target.files[0]?.name || '';
});

function getHistory() {
    try { return JSON.parse(localStorage.getItem(HISTORY_KEY)) || []; } catch { return []; }
}
function saveHistory(history) {
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
}
function addToHistory(filename, data, headers) {
    const history = getHistory();
    history.unshift({
        filename: filename,
        timestamp: Date.now(),
        data: data,
        headers: headers
    });
    if (history.length > 20) history.pop();
    saveHistory(history);
    loadHistoryDropdown();
}
function loadHistoryDropdown() {
    const select = document.getElementById('historySelect');
    const history = getHistory();
    select.innerHTML = '<option value="">— Выбрать предыдущий результат —</option>';
    history.forEach((entry, idx) => {
        const opt = document.createElement('option');
        opt.value = idx;
        const date = new Date(entry.timestamp).toLocaleString();
        opt.textContent = `${date} – ${entry.filename}`;
        select.appendChild(opt);
    });
}
document.getElementById('historySelect').addEventListener('change', function(e) {
    const idx = e.target.value;
    if (idx === '') return;
    const history = getHistory();
    const entry = history[parseInt(idx)];
    if (entry) {
        globalData = entry.data;
        globalHeaders = entry.headers;
        renderTable();
        showStatus(`Загружен результат "${entry.filename}"`, 'success');
    }
});
document.getElementById('clearHistoryBtn').addEventListener('click', function() {
    if (confirm('Удалить всю историю?')) {
        localStorage.removeItem(HISTORY_KEY);
        loadHistoryDropdown();
        showStatus('История очищена', 'success');
    }
});

async function uploadFile() {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files.length) return showStatus('Выберите файл.', 'error');
    const formData = new FormData(); 
    formData.append('file', fileInput.files[0]);
    showStatus('Загрузка файла...', 'info');

    try {
        const res = await fetch('/hierarchy/uploadfile/', { method: 'POST', body: formData });
        const result = await res.json();
        
        if (result.status === 'success') {
            globalData = result.data;
            if (globalData.length) {
                globalHeaders = Object.keys(globalData[0]);
                showStatus(`Успешно загружено: ${globalData.length} записей`, 'success');
                renderTable();
                document.getElementById('edit-block').style.display = 'block';
                document.getElementById('normalize-block').style.display = 'block';
                addToHistory(fileInput.files[0].name, globalData, globalHeaders);
            } else {
                showStatus('Файл пуст.', 'error');
            }
        } else {
            showStatus(`Ошибка: ${result.message || 'Неизвестная ошибка'}`, 'error');
        }
    } catch (e) { 
        showStatus('Ошибка сервера: ' + e.message, 'error'); 
    }
}

async function runNormalization() {
    if (!globalData.length) return showStatus('Нет данных', 'error');
    showStatus('Нормализация...', 'info');
    try {
        const res = await fetch('/hierarchy/normalize', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: globalData })
        });
        const result = await res.json();
        if (result.status === 'success') {
            globalData = result.data; 
            globalHeaders = Object.keys(globalData[0]);
            renderTable(); 
            showStatus('Нормализация завершена', 'success');
            const lastFile = document.getElementById('fileName').textContent || 'нормализованные_данные';
            addToHistory(lastFile + ' (нормализ.)', globalData, globalHeaders);
        } else { showStatus(`Ошибка: ${result.message}`, 'error'); }
    } catch (e) { showStatus('Ошибка сервера', 'error'); }
}

function renderTable() {
    const container = document.getElementById('view-table');
    if (!globalData.length) return;
    
    let html = '<table><thead><tr>';
    globalHeaders.forEach(h => html += `<th>${escapeHtml(h)}</th>`);
    html += '</tr></thead><tbody>';

    globalData.forEach((row, rI) => {
        let rowClass = row['needs_review'] ? 'needs-review-row' : '';
        html += `<tr class="${rowClass}">`;
        globalHeaders.forEach((h, cI) => {
            const val = row[h] !== undefined && row[h] !== null ? row[h] : '';
            html += `<td oncontextmenu="showContextMenu(event, ${rI}, ${cI})">${escapeHtml(val)}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
}

function showContextMenu(e, rowIdx, colIdx) {
    e.preventDefault();
    currentEditRow = rowIdx;
    currentEditCol = colIdx;
    const menu = document.getElementById('contextMenu');
    menu.innerHTML = `
        <div onclick="editCell()">✏️ Редактировать</div>
        <div onclick="createTechCard()">📋 Создать техкарту</div>
    `;
    menu.style.display = 'block';
    menu.style.left = e.clientX + 'px';
    menu.style.top = e.clientY + 'px';
    document.addEventListener('click', function hide() {
        menu.style.display = 'none';
        document.removeEventListener('click', hide);
    });
}

function editCell() {
    document.getElementById('contextMenu').style.display = 'none';
    if (currentEditRow < 0 || currentEditCol < 0) return;
    const colName = globalHeaders[currentEditCol];
    const oldVal = globalData[currentEditRow][colName] || '';
    const newVal = prompt('Изменить значение', oldVal);
    if (newVal !== null) {
        globalData[currentEditRow][colName] = newVal;
        renderTable();
        showStatus('Ячейка обновлена', 'success');
    }
}

function createTechCard() {
    document.getElementById('contextMenu').style.display = 'none';
    if (currentEditRow < 0 || currentEditCol < 0) return;
    const row = globalData[currentEditRow];
    let model = '';
    for (let key of globalHeaders) {
        if (key.toLowerCase().includes('модель')) {
            model = row[key] || '';
            break;
        }
    }
    if (!model) {
        model = row[globalHeaders[currentEditCol]] || '';
    }
    window.location.href = `/technical_map/?model=${encodeURIComponent(model)}`;
}

function saveAndDownload() {
    if (!globalData.length) return showStatus('Нет данных', 'error');
    const headerRow = globalHeaders.join(',') + '\n';
    const rows = globalData.map(row => globalHeaders.map(h => {
        let v = (row[h]||'').toString();
        if (v.includes(',') || v.includes('"')) v = `"${v.replace(/"/g, '""')}"`;
        return v;
    }).join(',')).join('\n');
    const blob = new Blob([headerRow+rows], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob); link.download = "classified_data.csv";
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    showStatus('CSV скачан', 'success');
}

function saveAndDownloadJSON() {
    if (!globalData.length) return showStatus('Нет данных', 'error');
    const blob = new Blob([JSON.stringify(globalData, null, 2)], { type: 'application/json;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob); link.download = "classified_data.json";
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    showStatus('JSON скачан', 'success');
}

loadHistoryDropdown();