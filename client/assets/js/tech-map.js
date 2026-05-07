(function() {
    'use strict';

    const state = {
        selectedFile: null,
        settings: {},
        history: [],
        tableHistory: [],
        generating: false,
        pendingMsgId: null,
        currentXlsxData: null,
        currentXlsxName: '',
        currentRows: [] 
    };

    window.TechMap = {
        init,
        openSettings,
        closeSettings,
        saveSettings,
        updateModelDropdown,
        validateKey
    };

    async function init() {
        try {
            await loadSettings();
        } catch (e) {
            console.warn('Настройки не загружены:', e);
        }
        loadHistoryFromStorage();
        loadTableHistory();
        renderHistory();
        initFileUpload();
        initEventListeners();
        renderEmptyTable();
        refreshHistorySelect();

        if (state.tableHistory.length > 0) {
            const last = state.tableHistory[0];
            if (last.rows) {
                fillTable(last.rows);
                state.currentXlsxData = last.xlsxFile || null;
                state.currentXlsxName = last.xlsxFileName || '';
                updateDownloadButton();
            }
        }

        console.log('✅ TechMap initialized');
    }

    async function loadSettings() {
        const s = await getSettings().catch(() => ({}));
        state.settings = s;
        document.getElementById('settingsApiKey').value = s.api_key || '';
        document.getElementById('settingsModel').value = s.model || 'openai/gpt-4o-mini';
        document.getElementById('settingsTemperature').value = s.temperature || 0.3;
        document.getElementById('settingsMaxTokens').value = s.max_tokens || 3000;
        document.getElementById('settingsMasterPrompt').value = s.master_prompt || '';
        updateModelDropdown();
    }

    function loadHistoryFromStorage() {
        try {
            state.history = JSON.parse(localStorage.getItem('techmap_history')) || [];
        } catch {
            state.history = [];
        }
    }

    function saveHistoryToStorage() {
        localStorage.setItem('techmap_history', JSON.stringify(state.history));
    }

    function renderHistory() {
        const container = document.getElementById('chatMessages');
        container.innerHTML = '';
        state.history.forEach(msg => addMessageToContainer(msg));
        container.scrollTop = container.scrollHeight;
    }

    function addMessageToContainer(msg, append = true) {
        const container = document.getElementById('chatMessages');
        const div = document.createElement('div');
        div.className = `message ${msg.role}`;
        div.innerHTML = `<div class="label">${msg.role === 'user' ? 'Вы' : 'ИИ'}</div><div>${msg.role === 'ai' ? renderMarkdown(msg.content) : escapeHtml(msg.content)}</div>`;
        if (msg.id) div.id = msg.id;
        if (append) container.appendChild(div);
        else container.prepend(div);
        if (append) container.scrollTop = container.scrollHeight;
        return div;
    }

    function renderMarkdown(text) {
        if (!text) return '';
        const rawHtml = marked.parse(text);
        return DOMPurify.sanitize(rawHtml);
    }

    function initFileUpload() {
        const input = document.getElementById('fileInput');
        const dropZone = document.getElementById('fileDropZone');
        if (!input || !dropZone) return;
        input.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                state.selectedFile = e.target.files[0];
                document.getElementById('fileName').textContent =
                    `${state.selectedFile.name} (${formatFileSize(state.selectedFile.size)})`;
                dropZone.classList.add('has-file');
            }
        });
        dropZone.addEventListener('dragover', (e) => e.preventDefault());
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            if (e.dataTransfer.files[0]) {
                state.selectedFile = e.dataTransfer.files[0];
                document.getElementById('fileName').textContent =
                    `${state.selectedFile.name} (${formatFileSize(state.selectedFile.size)})`;
                dropZone.classList.add('has-file');
            }
        });
    }

    function initEventListeners() {
        document.getElementById('sendBtn')?.addEventListener('click', handleSend);
        document.getElementById('chatInput')?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') handleSend();
        });
        document.getElementById('downloadBtn')?.addEventListener('click', downloadXlsx);
        document.getElementById('historySelect')?.addEventListener('change', onHistorySelect);
        document.getElementById('clearHistoryBtn')?.addEventListener('click', clearTableHistory);
    }

    function loadTableHistory() {
        try {
            state.tableHistory = JSON.parse(localStorage.getItem('techmap_table_history')) || [];
        } catch {
            state.tableHistory = [];
        }
    }

    function saveTableHistory() {
        localStorage.setItem('techmap_table_history', JSON.stringify(state.tableHistory));
    }

    function refreshHistorySelect() {
        const select = document.getElementById('historySelect');
        if (!select) return;
        select.innerHTML = '<option value="">— Выбрать предыдущую генерацию —</option>';
        state.tableHistory.forEach((entry, idx) => {
            const date = new Date(entry.timestamp).toLocaleString();
            const opt = document.createElement('option');
            opt.value = idx;
            opt.textContent = `${date} – ${entry.model || 'Без названия'}`;
            select.appendChild(opt);
        });
    }

    function clearTableHistory() {
        if (confirm('Удалить всю историю генераций таблиц?')) {
            state.tableHistory = [];
            saveTableHistory();
            refreshHistorySelect();
        }
    }

    function onHistorySelect(e) {
        const idx = e.target.value;
        if (idx === '' || idx === null) return;
        const entry = state.tableHistory[parseInt(idx)];
        if (!entry) return;

        if (entry.rows && entry.rows.length) {
            fillTable(entry.rows);
        }

        state.currentXlsxData = entry.xlsxFile || null;
        state.currentXlsxName = entry.xlsxFileName || '';

        updateDownloadButton();

        addAIMessage(`📂 Загружена предыдущая генерация: **${entry.model || 'Без названия'}** (${new Date(entry.timestamp).toLocaleString()})`);
        showStatus(`Загружена генерация ${entry.model || ''}`, 'success');
    }

    function addToTableHistory(modelName, textDesc, rows, xlsxFile = null, xlsxFileName = '') {
        const entry = {
            model: modelName,
            text: textDesc,
            rows: rows,
            timestamp: Date.now(),
            xlsxFile: xlsxFile,
            xlsxFileName: xlsxFileName
        };
        state.tableHistory.unshift(entry);
        if (state.tableHistory.length > 20) state.tableHistory.pop();
        saveTableHistory();
        refreshHistorySelect();
    }

    function renderEmptyTable() {
        const headers = [
            "Элемент","Подэлемент","Наименование операции","Краткое содержание работ",
            "Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей",
            "Профессия/Квалификация","Трудоёмкость, человеко/часов",
            "Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ",
            "Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"
        ];
        const thead = document.getElementById('resultTableHead');
        thead.innerHTML = `<tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>`;
        document.getElementById('resultTableBody').innerHTML = '';
    }

    function fillTable(rows) {
        if (!rows || !rows.length) return;
        const headers = [
            "Элемент","Подэлемент","Наименование операции","Краткое содержание работ",
            "Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей",
            "Профессия/Квалификация","Трудоёмкость, человеко/часов",
            "Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ",
            "Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"
        ];
        const tbody = document.getElementById('resultTableBody');
        tbody.innerHTML = rows.map(row => {
            const cells = headers.map(h => escapeHtml(row[h] ?? ''));
            return `<tr>${cells.map(c => `<td>${c}</td>`).join('')}</tr>`;
        }).join('');
        state.currentRows = rows;
        updateDownloadButton();
    }

    function updateDownloadButton() {
        const btn = document.getElementById('downloadBtn');
        if (!btn) return;
        btn.style.display = state.currentRows.length > 0 ? 'inline-flex' : 'none';
    }

    async function handleSend() {
        const input = document.getElementById('chatInput');
        const message = input.value.trim();
        const modelName = document.getElementById('modelName')?.value.trim();

        if (!message && !modelName) {
            showStatus('Введите сообщение или укажите модель для генерации техкарты', 'error');
            return;
        }

        if (!message && modelName) {
            await generateTechCardForModel(modelName);
            input.value = '';
            return;
        }

        if (message) {
            input.value = '';
            await sendChatMessage(message);
        }
    }

    async function generateTechCardForModel(modelName) {
        if (state.generating) return;
        const userMsg = `🔧 Генерация техкарты для: ${modelName}`;
        addUserMessage(userMsg);
        showPendingMessage();
        setLoading(true);
        try {
            const result = await generateTechCard(
                state.selectedFile,
                modelName,
                {
                    equipmentClass: document.getElementById('equipmentClass')?.value.trim() || '',
                    subclass: document.getElementById('subclass')?.value.trim() || '',
                    model: state.settings.model,
                    apiKey: state.settings.api_key,
                    temperature: state.settings.temperature,
                    maxTokens: state.settings.max_tokens,
                    masterPrompt: state.settings.master_prompt
                }
            );
            const textDesc = result.data?.text || 'Техкарта сгенерирована.';
            const rows = result.data?.rows || [];
            replacePendingMessage(textDesc);
            if (rows.length) {
                fillTable(rows);
                state.currentXlsxData = result.data?.xlsx_file;
                state.currentXlsxName = result.data?.xlsx_filename || 'Техкарта.xlsx';
                document.getElementById('downloadBtn').style.display = 'inline-flex';
                addToTableHistory(modelName, textDesc, rows, state.currentXlsxData, state.currentXlsxName);
            }
            showStatus('Техкарта готова', 'success');
        } catch (e) {
            replacePendingMessage(`❌ Ошибка: ${e.message}`);
            showStatus(e.message, 'error');
        } finally {
            setLoading(false);
        }
    }

    async function sendChatMessage(message) {
        if (state.generating) return;
        addUserMessage(message);
        showPendingMessage();
        setLoading(true);
        try {
            const data = await window.sendChatMessage(message, state.history);
            replacePendingMessage(data.reply || 'Нет ответа');
        } catch (e) {
            replacePendingMessage(`❌ Ошибка: ${e.message}`);
        } finally {
            setLoading(false);
        }
    }

    function addUserMessage(content) {
        const msg = { role: 'user', content };
        state.history.push(msg);
        saveHistoryToStorage();
        addMessageToContainer(msg);
    }

    function addAIMessage(content) {
        const msg = { role: 'ai', content };
        state.history.push(msg);
        saveHistoryToStorage();
        addMessageToContainer(msg);
    }

    function showPendingMessage() {
        if (state.pendingMsgId) {
            document.getElementById(state.pendingMsgId)?.remove();
            state.history = state.history.filter(m => m.id !== state.pendingMsgId);
        }
        const id = 'pending-' + Date.now();
        const msg = { role: 'ai', content: '⏳ Обработка запроса…', id };
        state.history.push(msg);
        saveHistoryToStorage();
        const div = addMessageToContainer(msg);
        state.pendingMsgId = id;
    }

    function replacePendingMessage(newContent) {
        const pendingId = state.pendingMsgId;
        if (pendingId) {
            const el = document.getElementById(pendingId);
            if (el) el.remove();
            state.history = state.history.filter(m => m.id !== pendingId);
            state.pendingMsgId = null;
        }
        const msg = { role: 'ai', content: newContent };
        state.history.push(msg);
        saveHistoryToStorage();
        addMessageToContainer(msg);
    }

    function setLoading(isLoading) {
        state.generating = isLoading;
        document.getElementById('sendBtn').disabled = isLoading;
        document.getElementById('chatInput').disabled = isLoading;
    }

    function downloadXlsx() {
        if (state.currentXlsxData) {
            const link = document.createElement('a');
            link.href = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${state.currentXlsxData}`;
            link.download = state.currentXlsxName || 'Техкарта.xlsx';
            link.click();
            return;
        }

        if (!state.currentRows || state.currentRows.length === 0) {
            showStatus('Нет данных для скачивания', 'error');
            return;
        }

        let modelName = '';
        const histSelect = document.getElementById('historySelect');
        if (histSelect && histSelect.value !== '') {
            const entry = state.tableHistory[parseInt(histSelect.value)];
            if (entry) modelName = entry.model || '';
        }
        if (!modelName) {
            modelName = document.getElementById('modelName')?.value.trim() || '';
        }
        const filename = `ТК_${modelName || 'модель'}.xlsx`;

        const headers = [
            "Элемент","Подэлемент","Наименование операции","Краткое содержание работ",
            "Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей",
            "Профессия/Квалификация","Трудоёмкость, человеко/часов",
            "Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ",
            "Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"
        ];
        const data = [headers, ...state.currentRows.map(row => headers.map(h => row[h] || ''))];
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Техкарта");
        XLSX.writeFile(wb, filename);
    }

    function updateModelDropdown() {
        const provider = document.getElementById('settingsProvider')?.value;
        getModels(provider).then(models => {
            const sel = document.getElementById('settingsModel');
            if (sel) sel.innerHTML = models.map(m => `<option value="${m.value}">${m.label}</option>`).join('');
        }).catch(() => {});
    }

    function openSettings() {
        document.getElementById('settingsModal')?.classList.add('active');
    }

    function closeSettings() {
        document.getElementById('settingsModal')?.classList.remove('active');
    }

    async function saveSettings() {
        const settings = {
            provider: document.getElementById('settingsProvider')?.value,
            api_key: document.getElementById('settingsApiKey')?.value.trim(),
            model: document.getElementById('settingsModel')?.value,
            max_tokens: parseInt(document.getElementById('settingsMaxTokens')?.value) || 3000,
            temperature: parseFloat(document.getElementById('settingsTemperature')?.value) || 0.3,
            master_prompt: document.getElementById('settingsMasterPrompt')?.value
        };
        try {
            await saveSettings(settings);
            state.settings = settings;
            closeSettings();
            showStatus('Настройки сохранены', 'success');
        } catch (e) {
            showStatus('Ошибка сохранения: ' + e.message, 'error');
        }
    }

    async function validateKey() {
        const key = document.getElementById('settingsApiKey')?.value.trim();
        if (!key) return showStatus('Введите ключ', 'error');
        const btn = document.querySelector('#settingsModal .btn-outline');
        if (btn) btn.disabled = true;
        try {
            const res = await window.validateKey(key);
            showStatus(res.message, res.valid ? 'success' : 'error');
        } catch (e) {
            showStatus('Ошибка проверки: ' + (e.message || 'неизвестная ошибка'), 'error');
        } finally {
            if (btn) btn.disabled = false;
        }
    }
})();