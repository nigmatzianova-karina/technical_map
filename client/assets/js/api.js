const BASE_URL = '';

async function apiRequest(endpoint, options = {}) {
    const url = `${BASE_URL}${endpoint}`;
    const config = { headers: { 'Content-Type': 'application/json', ...options.headers }, ...options };
    const response = await fetch(url, config);
    if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.detail || errData.message || `HTTP ${response.status}`);
    }
    return await response.json();
}

async function uploadFile(endpoint, file, additionalFields = {}) {
    const formData = new FormData();
    if (file) formData.append('file', file);
    Object.entries(additionalFields).forEach(([key, value]) => formData.append(key, value));
    const response = await fetch(`${BASE_URL}${endpoint}`, { method: 'POST', body: formData });
    if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.detail || errData.message || `HTTP ${response.status}`);
    }
    return await response.json();
}

async function generateTechCard(file, modelName, options = {}) {
    return uploadFile('/tech_map/api/generate', file, {
        model_name: modelName,
        equipment_class: options.equipmentClass || '',
        subclass: options.subclass || '',
        model: options.model || 'openai/gpt-4o-mini',
        api_key: options.apiKey || '',
        temperature: options.temperature || 0.3,
        max_tokens: options.maxTokens || 3000,
        master_prompt: options.masterPrompt || ''
    });
}

async function sendChatMessage(message, history = []) {
    const formData = new FormData();
    formData.append('message', message);
    formData.append('history', JSON.stringify(history));
    const response = await fetch('/tech_map/api/chat', { method: 'POST', body: formData });
    if (!response.ok) throw new Error('Ошибка чата');
    return response.json();
}

async function parsePdf(file) {
    return uploadFile('/pdf_parser/api/parse', file);
}

async function getSettings() {
    return apiRequest('/api/settings');
}

async function saveSettings(settings) {
    return apiRequest('/api/settings', { method: 'POST', body: JSON.stringify(settings) });
}

async function getModels(provider) {
    return apiRequest(`/api/models/${provider}`);
}

async function validateKey(apiKey) {
    return apiRequest('/api/key/validate', { method: 'POST', body: JSON.stringify({ api_key: apiKey }) });
}