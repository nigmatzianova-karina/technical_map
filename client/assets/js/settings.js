const SettingsManager = {
    defaultPrompt: `Ты инженер, специалист по формированию технологических карт и работ по ТОиР оборудования.

{file_instruction}

Необходимо заполнить:
1. Столбец "Элемент" — основной крупный элемент, входящий в состав узла. Например: Система смазки.
2. Столбец "Подэлемент" — более мелкий элемент, входящий в состав элемента. Например: Картер.

Правила:
• Каждый новый узел, элемент и подэлемент — в отдельной строке по порядку.
• НЕ вноси как "Элемент" или "Подэлемент": гайки, шайбы, винты, шпильки, хомуты, болты, штифты, шпонки.
• Если в столбцах несколько слов — первое слово всегда существительное, остальные после него.
• Элемент и подэлемент — в единственном числе, именительном падеже.
• Слова нельзя сокращать и заменять синонимами.
• Другие столбцы таблицы не удаляй и не изменяй.

ОТВЕТ ДОЛЖЕН БЫТЬ В СТРОГОМ ФОРМАТЕ:

[ТЕКСТ_ОТВЕТ]
Краткое текстовое описание результата для пользователя.
[/ТЕКСТ_ОТВЕТ]

[ТАБЛИЦА]
Элемент|Подэлемент|Наименование операции|Краткое содержание работ|Вид ТОиР|Периодичность|Норма времени, часов|Количество исполнителей|Профессия/Квалификация|Трудоёмкость, человеко/часов|Наименование ТМЦ|Количество ТМЦ|Единицы измерения ТМЦ|Наименование инструменты|Средства индивидуальной защиты|Требования по безопасности
Система смазки|Картер|Осмотр|Визуальный осмотр картера на наличие трещин и подтёков|ТО-1|4320|2.0|1|Слесарь по ремонту автомобилей, 3 разряд|2.0|||||Каска защитная, 1 шт; Очки защитные, 1 шт; Перчатки защитные, 1 пара|Затормозить технику; Выполнять работы при неработающем двигателе
[/ТАБЛИЦА]

ВАЖНО: Каждая строка таблицы — значения через "|". Всего 16 столбцов. Если данных нет — оставьте пусто (||).`,

    async load() {
        try {
            const res = await fetch('/api/settings');
            const settings = await res.json();
            
            document.getElementById('settingsProvider').value = settings.provider || 'openrouter';
            document.getElementById('settingsApiKey').value = settings.api_key || '';
            document.getElementById('settingsModel').value = settings.model || 'openai/gpt-4o-mini';
            document.getElementById('settingsMaxTokens').value = settings.max_tokens || 3000;
            document.getElementById('settingsTemperature').value = settings.temperature || 0.3;
            document.getElementById('settingsMasterPrompt').value = settings.master_prompt || this.defaultPrompt;
            
            await this.updateModelDropdown();
            
        } catch (e) {
            console.error('Ошибка загрузки настроек:', e);
            showStatus('Не удалось загрузить настройки', 'error');
        }
    },

    async save() {
        const settings = {
            provider: document.getElementById('settingsProvider').value,
            api_key: document.getElementById('settingsApiKey').value.trim(),
            model: document.getElementById('settingsModel').value,
            max_tokens: parseInt(document.getElementById('settingsMaxTokens').value) || 3000,
            temperature: parseFloat(document.getElementById('settingsTemperature').value) || 0.3,
            master_prompt: document.getElementById('settingsMasterPrompt').value.trim() || this.defaultPrompt
        };

        if (!settings.api_key) {
            alert('Введите API ключ');
            return false;
        }

        try {
            const res = await fetch('/api/settings', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(settings)
            });

            if (!res.ok) throw new Error('Ошибка сохранения');

            showStatus('Настройки сохранены', 'success');
            return true;
        } catch (e) {
            showStatus('Ошибка сохранения: ' + e.message, 'error');
            return false;
        }
    },

    async updateModelDropdown() {
        const provider = document.getElementById('settingsProvider').value;
        const modelSelect = document.getElementById('settingsModel');
        const currentVal = modelSelect.value;

        try {
            const res = await fetch(`/api/models/${provider}`);
            const models = await res.json();

            modelSelect.innerHTML = '';
            models.forEach(m => {
                const opt = document.createElement('option');
                opt.value = m.value;
                opt.textContent = m.label;
                modelSelect.appendChild(opt);
            });

            const exists = models.some(m => m.value === currentVal);
            if (exists) {
                modelSelect.value = currentVal;
            }
        } catch (e) {
            console.error('Ошибка загрузки моделей:', e);
        }
    },

    async validateKey() {
        const apiKey = document.getElementById('settingsApiKey').value.trim();
        if (!apiKey) {
            showStatus('Введите API ключ', 'error');
            return;
        }

        const btn = document.getElementById('validateKeyBtn');
        const originalHTML = btn.innerHTML;
        
        btn.disabled = true;
        btn.innerHTML = `<span class="spinner-inline" style="width:12px;height:12px;border-width:2px;margin-right:6px;"></span> Проверка...`;

        try {
            const res = await fetch('/api/key/validate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ api_key: apiKey })
            });

            const data = await res.json();

            if (data.valid) {
                showStatus('✅ ' + data.message, 'success');
                btn.innerHTML = '✅ Валиден';
                setTimeout(() => {
                    btn.innerHTML = originalHTML;
                    btn.disabled = false;
                }, 2000);
            } else {
                showStatus('❌ ' + data.message, 'error');
                btn.innerHTML = originalHTML;
                btn.disabled = false;
            }
        } catch (e) {
            showStatus('Ошибка сети: ' + e.message, 'error');
            btn.innerHTML = originalHTML;
            btn.disabled = false;
        }
    }
};

window.openSettings = function() {
    document.getElementById('settingsModal').classList.add('active');
};

window.closeSettings = function() {
    document.getElementById('settingsModal').classList.remove('active');
};

window.saveSettings = async function() {
    const success = await SettingsManager.save();
    if (success) {
        closeSettings();
    }
};

window.updateModelDropdown = function() {
    SettingsManager.updateModelDropdown();
};

window.validateKey = function() {
    SettingsManager.validateKey();
};

document.addEventListener('DOMContentLoaded', () => {
    SettingsManager.load();
    
    document.getElementById('settingsProvider')?.addEventListener('change', () => {
        SettingsManager.updateModelDropdown();
    });
});