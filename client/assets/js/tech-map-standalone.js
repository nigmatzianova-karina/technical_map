(function() {
    'use strict';
    
    const state = {
        selectedFile: null,
        isGenerating: false,
        settings: {
            apiKey: '',
            model: 'openai/gpt-4o-mini'
        }
    };

    window.TechMap = {
        init,
        openSettings,
        closeSettings,
        saveSettings,
        handleGenerateTechCard,
        handleChatSend
    };

    function init() {
        loadSettings();
        initFileUpload();
        initEventListeners();
        console.log('✅ TechMap initialized (standalone mode)');
    }

    function loadSettings() {
        try {
            const saved = localStorage.getItem('techmap_settings');
            if (saved) {
                state.settings = { ...state.settings, ...JSON.parse(saved) };
                const apiKeyEl = document.getElementById('settingsApiKey');
                const modelEl = document.getElementById('settingsModel');
                if (apiKeyEl && state.settings.apiKey) apiKeyEl.value = state.settings.apiKey;
                if (modelEl && state.settings.model) modelEl.value = state.settings.model;
            }
        } catch (e) {
            console.warn('Не удалось загрузить настройки из localStorage');
        }
    }

    function initFileUpload() {
        const input = document.getElementById('fileInput');
        const dropZone = document.getElementById('fileDropZone');
        
        if (!input || !dropZone) return;
        
        input.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                state.selectedFile = e.target.files[0];
                const nameEl = document.getElementById('fileName');
                if (nameEl) {
                    nameEl.textContent = `${state.selectedFile.name} (${formatFileSize(state.selectedFile.size)})`;
                }
                dropZone.classList.add('has-file');
            }
        });

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = 'var(--primary, #6e8efb)';
        });
        dropZone.addEventListener('dragleave', () => {
            dropZone.style.borderColor = '';
        });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '';
            if (e.dataTransfer.files[0]) {
                state.selectedFile = e.dataTransfer.files[0];
                const nameEl = document.getElementById('fileName');
                if (nameEl) {
                    nameEl.textContent = `${state.selectedFile.name} (${formatFileSize(state.selectedFile.size)})`;
                }
                dropZone.classList.add('has-file');
            }
        });
    }

    function initEventListeners() {
        document.getElementById('sendBtn')?.addEventListener('click', handleGenerateTechCard);
        document.getElementById('chatSendBtn')?.addEventListener('click', handleChatSend);
        document.getElementById('chatInput')?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') handleChatSend();
        });
        document.getElementById('downloadBtn')?.addEventListener('click', downloadXlsx);
    }

    async function handleGenerateTechCard() {
        if (state.isGenerating) return;
        
        const modelName = document.getElementById('modelName')?.value.trim();
        if (!modelName) {
            showStatus('Введите название модели оборудования', 'error');
            return;
        }
        
        setLoading(true);
        addChatMessage('user', `🔧 Сгенерировать техкарту для: ${modelName}`);
        
        try {
            const apiUrl = getApiBaseUrl();
            const endpoint = `${apiUrl}/api/parsing/extract/tech-card`;
            
            const formData = new FormData();
            
            if (state.selectedFile) {
                formData.append('file', state.selectedFile);
                formData.append('has_document', 'true');
            } else {
                formData.append('has_document', 'false');
            }
            
            formData.append('model_name', modelName);
            formData.append('model', state.settings.model);
            formData.append('api_key', state.settings.apiKey);
            
            const equipClass = document.getElementById('equipmentClass')?.value.trim();
            const subclass = document.getElementById('subclass')?.value.trim();
            if (equipClass) formData.append('equipment_class', equipClass);
            if (subclass) formData.append('subclass', subclass);
            

            
            const response = await fetch(endpoint, { method: 'POST', body: formData });
            
            if (!response.ok) {
                const err = await response.json().catch(() => ({}));
                throw new Error(err.detail || err.message || `HTTP ${response.status}`);
            }
            
            const result = await response.json();
            
            if (result.success) {
                addChatMessage('ai', '✅ Технологическая карта сгенерирована!');
                displayResults(result.data);
                showStatus('Готово!', 'success');
            } else {
                throw new Error(result.error || 'Неизвестная ошибка');
            }
            
        } catch (error) {
            console.error('Ошибка генерации:', error);
            addChatMessage('ai', `❌ Ошибка: ${error.message}`);
            showStatus(error.message, 'error');
        } finally {
            setLoading(false);
        }
    }

    async function fallbackToOldChat(message) {
        try {
            const apiUrl = getApiBaseUrl();
            const formData = new FormData();
            formData.append('message', message);
            formData.append('model_name', document.getElementById('modelName')?.value || '');
            formData.append('equipment_class', document.getElementById('equipmentClass')?.value || '');
            formData.append('subclass', document.getElementById('subclass')?.value || '');
            formData.append('provider', 'openrouter');
            if (state.selectedFile) formData.append('file', state.selectedFile);
            
            const response = await fetch(`${apiUrl}/technical_map/api/chat`, {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            const result = await response.json();
            
            addChatMessage('ai', result.text || 'Данные получены');
            if (result.table_rows?.length > 0) {
                displayResults({ rows: result.table_rows, text: result.text });
            }
        } catch (e) {
            addChatMessage('ai', `❌ Ошибка чата: ${e.message}`);
        }
    }

    function displayResults(data) {
        console.log('📊 Displaying results:', data); 
        
        const section = document.getElementById('resultsSection');
        if (!section) {
            console.error('❌ Не найден элемент #resultsSection');
            showStatus('Ошибка: не удалось найти блок результатов', 'error');
            return;
        }
        
        section.classList.remove('hidden');
        section.classList.add('visible');
        console.log('✅ Section visible');
        
        const textEl = document.getElementById('aiTextResult');
        if (textEl) {
            if (data.text) {
                textEl.textContent = data.text;
                textEl.style.display = 'block';
            } else {
                textEl.style.display = 'none';
            }
        }
        
        if (data.rows && data.rows.length > 0) {
            renderTable(data.rows);
        } else {
            console.warn('⚠️ Нет данных для таблицы (rows)');
        }
        
        const downloadBtn = document.getElementById('downloadBtn');
        if (downloadBtn) {
            if (data.xlsx_file) {
                window._currentXlsxData = data.xlsx_file;
                window._currentXlsxName = data.xlsx_filename || 'Техкарта.xlsx';
                downloadBtn.style.display = 'inline-flex';
            } else {
                downloadBtn.style.display = 'none';
            }
        }
        
        setTimeout(() => {
            section.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
        
        console.log('✅ Results displayed');
    }

    function renderTable(rows) {
        const thead = document.getElementById('resultTableHead');
        const tbody = document.getElementById('resultTableBody');
        if (!thead || !tbody || !rows.length) return;
        
        const headers = Object.keys(rows[0]);
        thead.innerHTML = `<tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>`;
        tbody.innerHTML = rows.map(row => 
            `<tr>${headers.map(h => `<td>${escapeHtml(row[h] ?? '')}</td>`).join('')}</tr>`
        ).join('');
    }

    function downloadXlsx() {
        if (!window._currentXlsxData) return;
        const link = document.createElement('a');
        link.href = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${window._currentXlsxData}`;
        link.download = window._currentXlsxName;
        link.click();
    }

    async function handleChatSend() {
        if (state.isGenerating) return;
        const input = document.getElementById('chatInput');
        const message = input?.value.trim();
        if (!message) return;
        
        setLoading(true);
        addChatMessage('user', message);
        if (input) input.value = '';
        
        await fallbackToOldChat(message);
        setLoading(false);
    }

    function addChatMessage(role, text) {
        const container = document.getElementById('chatMessages');
        if (!container) return;
        
        const div = document.createElement('div');
        div.className = `message ${role}`;
        div.style.cssText = `
            align-self: ${role === 'user' ? 'flex-end' : 'flex-start'};
            background: ${role === 'user' ? 'linear-gradient(135deg, #6e8efb, #a777e3)' : '#1a1a2e'};
            border: ${role === 'ai' ? '1px solid #2a2a4a' : 'none'};
            border-radius: 12px; padding: 10px 14px; max-width: 90%; color: ${role === 'user' ? 'white' : '#e0e0e0'};
        `;
        div.innerHTML = `
            <div style="font-size: 11px; color: #888; margin-bottom: 4px;">${role === 'user' ? 'Вы' : 'ИИ-ассистент'}</div>
            <div>${escapeHtml(text)}</div>
        `;
        container.appendChild(div);
        container.scrollTop = container.scrollHeight;
    }

    function openSettings() {
        const modal = document.getElementById('settingsModal');
        if (modal) modal.classList.remove('hidden');
    }

    function closeSettings() {
        const modal = document.getElementById('settingsModal');
        if (modal) modal.classList.add('hidden');
    }

    function saveSettings() {
        const apiKey = document.getElementById('settingsApiKey')?.value || '';
        const model = document.getElementById('settingsModel')?.value || 'openai/gpt-4o-mini';
        
        state.settings.apiKey = apiKey;
        state.settings.model = model;
        
        try {
            localStorage.setItem('techmap_settings', JSON.stringify(state.settings));
            closeSettings();
            showStatus('Настройки сохранены', 'success');
        } catch (e) {
            showStatus('Ошибка сохранения', 'error');
        }
    }

    function setLoading(isLoading) {
        state.isGenerating = isLoading;
        
        const sendBtn = document.getElementById('sendBtn');
        const chatBtn = document.getElementById('chatSendBtn');
        const loader = document.getElementById('resultsLoader');
        
        if (sendBtn) {
            sendBtn.disabled = isLoading;
            sendBtn.innerHTML = isLoading ? '<span class="spinner-inline"></span> Генерация...' : 'Отправить';
        }
        if (chatBtn) chatBtn.disabled = isLoading;
        if (loader) loader.classList.toggle('active', isLoading);
        
        ['chatInput', 'mainMessage'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.disabled = isLoading;
        });
    }

    function getApiBaseUrl() {
        if (window.location.protocol === 'file:') {
            return 'http://localhost:8000';
        }
        return '';
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Б';
        const k = 1024;
        const sizes = ['Б', 'КБ', 'МБ', 'ГБ'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    if (typeof escapeHtml !== 'function') {
        window.escapeHtml = function(text) {
            if (typeof text !== 'string') return text;
            const d = document.createElement('div');
            d.textContent = text;
            return d.innerHTML;
        };
    }
    if (typeof showStatus !== 'function') {
        window.showStatus = function(msg, type = 'info') {
            let el = document.getElementById('status-message');
            if (!el) {
                el = document.createElement('div');
                el.id = 'status-message';
                document.body.appendChild(el);
            }
            el.textContent = msg;
            el.style.cssText = `
                border-color: ${type === 'success' ? '#10b981' : type === 'error' ? '#ef4444' : '#6e8efb'};
                color: ${type === 'success' ? '#10b981' : type === 'error' ? '#ef4444' : '#6e8efb'};
                display: block;
            `;
            if (type === 'success') setTimeout(() => { el.style.display = 'none'; }, 5000);
        };
    }

})();