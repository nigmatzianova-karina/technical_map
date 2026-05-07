function escapeHtml(text) {
    if (typeof text !== 'string') return text;
    const d = document.createElement('div');
    d.textContent = text;
    return d.innerHTML;
}

function showStatus(msg, type = 'info', duration = 5000) {
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
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Б';
    const k = 1024;
    const sizes = ['Б', 'КБ', 'МБ', 'ГБ'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}