// Конфигурация
const API_URL = 'http://localhost:8000';

// Элементы DOM
const statusEl = document.getElementById('status');
const timestampEl = document.getElementById('timestamp');
const fileInput = document.getElementById('fileInput');
const fileLabel = document.getElementById('fileLabel');
const uploadForm = document.getElementById('uploadForm');
const uploadBtn = document.getElementById('uploadBtn');
const uploadResult = document.getElementById('uploadResult');
const reconcileBtn = document.getElementById('reconcileBtn');
const reconcileResult = document.getElementById('reconcileResult');
const logsBtn = document.getElementById('logsBtn');
const logsOutput = document.getElementById('logsOutput');

// Обновление статуса
function updateStatus(message, isError = false) {
    statusEl.textContent = message;
    statusEl.style.color = isError ? '#dc2626' : '#16a34a';
}

// Обновление времени
function updateTimestamp() {
    const now = new Date();
    timestampEl.textContent = now.toLocaleString('ru-RU');
}
setInterval(updateTimestamp, 10000);
updateTimestamp();

// Показ результата
function showResult(element, message, type = 'success') {
    element.textContent = message;
    element.className = `result ${type}`;
    element.classList.remove('hidden');
}

function hideResult(element) {
    element.classList.add('hidden');
}

// Обработка выбора файла
fileInput.addEventListener('change', function() {
    if (this.files.length > 0) {
        fileLabel.textContent = ` ${this.files[0].name}`;
    } else {
        fileLabel.textContent = 'Выберите файл...';
    }
});

// Загрузка файла
uploadForm.addEventListener('submit', async function(e) {
    e.preventDefault();
    hideResult(uploadResult);
    
    const file = fileInput.files[0];
    if (!file) {
        showResult(uploadResult, '❌ Пожалуйста, выберите файл', 'error');
        return;
    }
    
    uploadBtn.disabled = true;
    uploadBtn.textContent = '⏳ Загрузка...';
    updateStatus('⏳ Загрузка файла...');
    
    try {
        const formData = new FormData();
        formData.append('file', file);
        
        const response = await fetch(`${API_URL}/upload`, {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.status === 'success') {
            showResult(uploadResult, 
                `✅ Файл загружен: ${data.filename}\n Листы: ${data.sheets.join(', ')}`,
                'success'
            );
            updateStatus('✅ Файл загружен');
        } else {
            showResult(uploadResult, 
                `❌ Ошибка: ${data.message || data.error || 'Неизвестная ошибка'}`,
                'error'
            );
            updateStatus('❌ Ошибка загрузки', true);
        }
    } catch (error) {
        showResult(uploadResult, `❌ Ошибка соединения: ${error.message}`, 'error');
        updateStatus('❌ Ошибка соединения', true);
    }
    
    uploadBtn.disabled = false;
    uploadBtn.textContent = ' Загрузить';
});

// Запуск сверки
reconcileBtn.addEventListener('click', async function() {
    hideResult(reconcileResult);
    reconcileBtn.disabled = true;
    reconcileBtn.textContent = '⏳ Выполняется...';
    updateStatus('⏳ Выполняется сверка...');
    
    try {
        const response = await fetch(`${API_URL}/reconcile`, {
            method: 'POST'
        });
        
        const data = await response.json();
        
        if (response.ok) {
            showResult(reconcileResult, 
                `✅ Сверка запущена\n Статус: ${data.status}\n ${data.timestamp}`,
                'success'
            );
            updateStatus('✅ Сверка выполнена');
        } else {
            showResult(reconcileResult, `❌ Ошибка: ${data.detail || 'Неизвестная ошибка'}`, 'error');
            updateStatus('❌ Ошибка сверки', true);
        }
    } catch (error) {
        showResult(reconcileResult, `❌ Ошибка соединения: ${error.message}`, 'error');
        updateStatus('❌ Ошибка соединения', true);
    }
    
    reconcileBtn.disabled = false;
    reconcileBtn.textContent = ' Выполнить сверку';
});

// Показ логов
logsBtn.addEventListener('click', async function() {
    if (logsOutput.classList.contains('visible')) {
        logsOutput.classList.remove('visible');
        logsOutput.innerHTML = '';
        logsBtn.textContent = ' Показать логи';
        return;
    }
    
    logsBtn.disabled = true;
    logsBtn.textContent = '⏳ Загрузка...';
    
    try {
        const response = await fetch(`${API_URL}/logs?limit=100`);
        const data = await response.json();
        
        if (response.ok) {
            if (data.logs && data.logs.length > 0) {
                logsOutput.innerHTML = data.logs.map(line => 
                    `<div class="log-line">${line.trim()}</div>`
                ).join('');
                logsOutput.classList.add('visible');
                logsBtn.textContent = ' Скрыть логи';
            } else {
                showResult(uploadResult, ' Логов пока нет', 'success');
                setTimeout(() => hideResult(uploadResult), 3000);
            }
        } else {
            showResult(uploadResult, `❌ Ошибка: ${data.detail || 'Неизвестная ошибка'}`, 'error');
            setTimeout(() => hideResult(uploadResult), 3000);
        }
    } catch (error) {
        showResult(uploadResult, `❌ Ошибка соединения: ${error.message}`, 'error');
        setTimeout(() => hideResult(uploadResult), 3000);
    }
    
    logsBtn.disabled = false;
    if (!logsOutput.classList.contains('visible')) {
        logsBtn.textContent = ' Показать логи';
    }
});

// Инициализация
updateStatus(' Система готова');
console.log('RN Reconciliation App initialized');
