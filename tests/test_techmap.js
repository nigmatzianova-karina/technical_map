// ============================================================
// Создаём временные DOM-элементы для полной совместимости с TechMap
// ============================================================
function setupTestDOM() {
    const container = document.createElement('div');
    container.style.display = 'none';
    container.innerHTML = `
        <input type="text" id="modelName">
        <input type="text" id="equipmentClass">
        <input type="text" id="subclass">
        <input type="file" id="fileInput">
        <div id="fileDropZone"><div id="fileName"></div></div>
        <button id="sendBtn"></button>
        <input type="text" id="chatInput">
        <div id="chatMessages"><div class="message ai"><div class="label"></div><div></div></div></div>
        <div id="resultsSection">
            <button id="downloadBtn"></button>
            <div class="table-wrapper">
                <table><thead id="resultTableHead"></thead><tbody id="resultTableBody"></tbody></table>
            </div>
            <select id="historySelect"></select>
            <button id="clearHistoryBtn"></button>
        </div>
        <div class="modal-overlay" id="settingsModal">
            <input type="password" id="settingsApiKey">
            <input type="number" id="settingsTemperature">
            <input type="number" id="settingsMaxTokens">
            <textarea id="settingsMasterPrompt"></textarea>
            <select id="settingsProvider"></select>
            <select id="settingsModel"></select>
            <button class="btn-outline"></button>
        </div>
        <div id="status-message"></div>
        <div id="resultsLoader"></div>
    `;
    document.body.appendChild(container);
}

function cleanupTestDOM() {
    const div = document.querySelector('div[style*="display: none"]');
    if (div) div.remove();
}

// ============================================================
// Подмена fetch
// ============================================================
const originalFetch = window.fetch;
let mockResponses = {};

function mockFetch(url, options) {
    const urlStr = typeof url === 'string' ? url : url.url;
    for (let [pattern, response] of Object.entries(mockResponses)) {
        if (urlStr.includes(pattern)) {
            return Promise.resolve({
                ok: true,
                json: () => Promise.resolve(response)
            });
        }
    }
    return originalFetch.apply(this, arguments);
}

const delay = ms => new Promise(res => setTimeout(res, ms));
let testResults = [];

function log(msg, pass = true) {
    const cls = pass ? 'pass' : 'fail';
    document.getElementById('log').innerHTML +=
        `<div class="result ${cls}"><b>${pass ? '✅' : '❌'} ${msg}</b></div>`;
    testResults.push({ msg, pass });
}

// ============================================================
// Запуск тестов
// ============================================================
async function runAllTests() {
    testResults = [];
    document.getElementById('log').innerHTML = '';

    setupTestDOM();
    window.fetch = mockFetch;

    // 1. Инициализация
    try {
        await TechMap.init();
        log('Инициализация TechMap без ошибок');
    } catch (e) {
        log(`Ошибка инициализации: ${e.message}`, false);
    }

    // 2. Проверка наличия ключевых элементов
    const ids = ['modelName', 'sendBtn', 'chatInput', 'resultsSection', 'historySelect'];
    let allFound = true;
    ids.forEach(id => {
        if (!document.getElementById(id)) {
            allFound = false;
            log(`Элемент #${id} не найден`, false);
        }
    });
    if (allFound) log('Все ключевые элементы присутствуют');

    // 3. Отправка с пустыми полями (ожидается статус ошибки)
    document.getElementById('modelName').value = '';
    document.getElementById('chatInput').value = '';
    try {
        document.getElementById('sendBtn').click();   // кнопка вызывает handleSend
        await delay(50);
        const statusEl = document.getElementById('status-message');
        if (statusEl && statusEl.textContent.includes('Введите сообщение или укажите модель')) {
            log('Пустая отправка показывает правильную ошибку');
        } else {
            log('Пустая отправка не показала ошибку', false);
        }
    } catch (e) {
        log(`Ошибка при пустой отправке: ${e.message}`, false);
    }

    // 4. Чат-сообщение
    mockResponses['/tech_map/api/chat'] = { reply: 'Тестовый ответ чата' };
    document.getElementById('chatInput').value = 'Что такое ТОиР?';
    try {
        document.getElementById('sendBtn').click();
        await delay(100);
        const lastMsg = document.querySelector('#chatMessages .message:last-child');
        if (lastMsg && lastMsg.textContent.includes('Тестовый ответ чата')) {
            log('Чат-сообщение обработано и ответ получен');
        } else {
            log('Чат-сообщение не отобразилось', false);
        }
    } catch (e) {
        log(`Ошибка чата: ${e.message}`, false);
    }

    // 5. Генерация техкарты
    mockResponses['/tech_map/api/generate'] = {
        success: true,
        data: {
            text: 'Тестовая техкарта',
            rows: [{
                "Элемент":"Двигатель",
                "Подэлемент":"Поршень",
                "Наименование операции":"Замена",
                "Краткое содержание работ":"...",
                "Вид ТОиР":"ТО-2",
                "Периодичность":"1000ч",
                "Норма времени, часов":"2",
                "Количество исполнителей":"1",
                "Профессия/Квалификация":"Механик",
                "Трудоёмкость, человеко/часов":"2",
                "Наименование ТМЦ":"",
                "Количество ТМЦ":"",
                "Единицы измерения ТМЦ":"",
                "Наименование инструмента":"",
                "Средства индивидуальной защиты":"",
                "Требования по безопасности":""
            }],
            xlsx_file: 'dGVzdA==',
            xlsx_filename: 'Тест.xlsx'
        }
    };
    document.getElementById('modelName').value = 'TestModel';
    document.getElementById('chatInput').value = '';
    try {
        document.getElementById('sendBtn').click();
        await delay(100);
        const tbody = document.getElementById('resultTableBody');
        if (tbody && tbody.innerHTML.includes('Двигатель')) {
            log('Таблица заполнена после генерации');
        } else {
            log('Таблица не заполнилась', false);
        }
        const histSelect = document.getElementById('historySelect');
        if (histSelect && histSelect.options.length > 1 &&
            histSelect.options[1].text.includes('TestModel')) {
            log('История генераций обновлена');
        } else {
            log('История генераций не обновилась', false);
        }
    } catch (e) {
        log(`Ошибка генерации: ${e.message}`, false);
    }

    // 6. Проверка ключа
    mockResponses['/api/key/validate'] = { valid: true, message: 'Ключ валиден' };
    document.getElementById('settingsApiKey').value = 'test-key';
    try {
        await TechMap.validateKey();
        await delay(50);
        const statusEl = document.getElementById('status-message');
        if (statusEl && statusEl.textContent.includes('валиден')) {
            log('Проверка ключа работает');
        } else {
            log('Проверка ключа не показала правильный статус', false);
        }
    } catch (e) {
        log(`Ошибка проверки ключа: ${e.message}`, false);
    }

    // 7. Очистка истории
    try {
        document.getElementById('clearHistoryBtn').click();
        await delay(50);
        const histSelect = document.getElementById('historySelect');
        if (histSelect && histSelect.options.length === 1) {
            log('Очистка истории таблиц работает');
        } else {
            log('История после очистки не пуста', false);
        }
    } catch (e) {
        log(`Ошибка очистки истории: ${e.message}`, false);
    }

    const total = testResults.length;
    const passed = testResults.filter(r => r.pass).length;
    document.getElementById('log').innerHTML +=
        `<hr><b>Итого: ${passed}/${total} тестов пройдено</b>`;

    window.fetch = originalFetch;
    cleanupTestDOM();
}