// Подмена fetch
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

window.fetch = mockFetch;

const delay = ms => new Promise(res => setTimeout(res, ms));
let testResults = [];

function log(msg, pass = true) {
    const cls = pass ? 'pass' : 'fail';
    document.getElementById('log').innerHTML +=
        `<div class="result ${cls}"><b>${pass ? '✅' : '❌'} ${msg}</b></div>`;
    testResults.push({ msg, pass });
}

function setupTestDOM() {
    const container = document.createElement('div');
    container.innerHTML = `
        <div id="fileDropZone" style="margin-bottom:10px;">
            <input type="file" id="fileInput" accept=".pdf,.docx">
            <div id="fileName" style="color:var(--text-light);"></div>
        </div>
        <button id="parseBtn">Распарсить</button>
        <span id="loadingIndicator" style="display:none;"></span>
        <div id="resultsSection" class="results-section" style="display:none; margin-top:20px;">
            <div id="parsedContent"></div>
            <button id="downloadTextBtn" style="display:none;">Текст</button>
            <button id="downloadExcelBtn" style="display:none;">Excel</button>
            <select id="historySelect" style="margin-top:10px;">
                <option value="">— Выбрать —</option>
            </select>
            <button id="clearHistoryBtn">Очистить</button>
        </div>
        <div id="status-message"></div>
    `;
    document.body.appendChild(container);
    return container;
}

function cleanupTestDOM(container) {
    if (container) container.remove();
}

async function runAllTests() {
    testResults = [];
    document.getElementById('log').innerHTML = '';

    const container = setupTestDOM();

    // Динамически загружаем pdf_parser.js ПОСЛЕ создания DOM
    await new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = '/assets/js/pdf_parser.js';
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
    // Даём время на выполнение скрипта (DOMContentLoaded уже был, но в pdf_parser используется window.addEventListener('DOMContentLoaded', ...))
    // Принудительно вызываем DOMContentLoaded, чтобы инициализация сработала.
    document.dispatchEvent(new Event('DOMContentLoaded'));
    await delay(100);

    // 1. Проверка элементов
    const ids = ['fileInput', 'parseBtn', 'resultsSection', 'downloadTextBtn', 'downloadExcelBtn', 'historySelect'];
    let allFound = true;
    ids.forEach(id => {
        if (!document.getElementById(id)) {
            allFound = false;
            log(`Элемент #${id} не найден`, false);
        }
    });
    if (allFound) log('Все ключевые элементы присутствуют');

    // 2. Выбор файла
    const input = document.getElementById('fileInput');
    const file = new File(['dummy content'], 'test.pdf', { type: 'application/pdf' });
    Object.defineProperty(input, 'files', { value: [file] });
    input.dispatchEvent(new Event('change'));
    await delay(10);
    const fileNameEl = document.getElementById('fileName');
    if (fileNameEl && fileNameEl.textContent.includes('test.pdf')) {
        log('Файл успешно выбран и отображён');
    } else {
        log('Имя файла не отобразилось', false);
    }

    // Мок ответа API
    mockResponses['/pdf_parser/api/parse'] = {
        success: true,
        filename: 'test.pdf',
        file_type: 'PDF',
        data: {
            pages_text: ['Страница 1 текст', 'Страница 2 текст'],
            tables: [
                [ ['Заголовок1', 'Заголовок2'], ['Ячейка1', 'Ячейка2'] ]
            ],
            page_count: 2,
            has_tables: true
        },
        xlsx_file: 'dGVzdA==',
        xlsx_filename: 'test.xlsx'
    };

    // 3. Нажатие кнопки "Распарсить"
    document.getElementById('parseBtn').click();
    await delay(200); // ожидаем ответа и отрисовки

    const section = document.getElementById('resultsSection');
    if (section && section.style.display !== 'none') {
        log('Блок результатов отображается после парсинга');
    } else {
        log('Блок результатов не отобразился', false);
    }

    const content = document.getElementById('parsedContent');
    if (content && content.innerHTML.includes('Заголовок1')) {
        log('Таблица корректно отображена в результатах');
    } else {
        log('Данные таблицы не найдены в DOM', false);
    }
    if (content && content.innerHTML.includes('Страница 1 текст')) {
        log('Текст корректно отображён в результатах');
    } else {
        log('Текст не отобразился', false);
    }

    // 4. Кнопки скачивания
    const downloadText = document.getElementById('downloadTextBtn');
    const downloadExcel = document.getElementById('downloadExcelBtn');
    if (downloadText.style.display === 'inline-flex' || downloadText.style.display === 'flex') {
        log('Кнопка скачивания текста активна');
    } else {
        log('Кнопка скачивания текста не активна', false);
    }
    if (downloadExcel.style.display === 'inline-flex' || downloadExcel.style.display === 'flex') {
        log('Кнопка скачивания Excel активна');
    } else {
        log('Кнопка скачивания Excel не активна', false);
    }

    // 5. История
    const historySelect = document.getElementById('historySelect');
    if (historySelect && historySelect.options.length > 1) {
        log('История содержит новую запись');
    } else {
        log('История не обновлена', false);
    }

    // 6. Очистка истории
    document.getElementById('clearHistoryBtn').click();
    await delay(50);
    if (historySelect && historySelect.options.length === 1) {
        log('Очистка истории работает');
    } else {
        log('История после очистки не пуста', false);
    }

    const total = testResults.length;
    const passed = testResults.filter(r => r.pass).length;
    document.getElementById('log').innerHTML +=
        `<hr><b>Итого: ${passed}/${total} тестов пройдено</b>`;

    window.fetch = originalFetch;
    cleanupTestDOM(container);
}