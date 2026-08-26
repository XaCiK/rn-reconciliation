// RN Reconciliation Frontend

const API_URL = 'http://localhost:8000';

let bankUploaded = false;
let rnUploaded = false;

const bankFileInput = document.getElementById('bankFile');
const rnFileInput = document.getElementById('rnFile');
const bankStatus = document.getElementById('bankStatus');
const rnStatus = document.getElementById('rnStatus');
const reconcileBtn = document.getElementById('reconcileBtn');
const loading = document.getElementById('loading');
const resultsDiv = document.getElementById('results');
const resultContent = document.getElementById('resultContent');
const errorDiv = document.getElementById('error');
const statusSpan = document.getElementById('status');

// File upload handlers
bankFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) {
        bankStatus.textContent = 'Файл не выбран';
        bankStatus.style.color = '#666';
        bankUploaded = false;
        updateButton();
        return;
    }
    bankStatus.textContent = ` ${file.name} (загрузка...)`;
    bankStatus.style.color = '#ffa500';
    
    try {
        const result = await uploadFile(file, 'bank');
        bankUploaded = true;
        bankStatus.textContent = `✅ ${file.name} (${result.rows || 0} строк)`;
        bankStatus.style.color = '#4caf50';
        updateButton();
    } catch (err) {
        bankUploaded = false;
        bankStatus.textContent = `❌ Ошибка: ${err.message}`;
        bankStatus.style.color = '#f44336';
    }
});

rnFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) {
        rnStatus.textContent = 'Файл не выбран';
        rnStatus.style.color = '#666';
        rnUploaded = false;
        updateButton();
        return;
    }
    rnStatus.textContent = ` ${file.name} (загрузка...)`;
    rnStatus.style.color = '#ffa500';
    
    try {
        const result = await uploadFile(file, 'rn');
        rnUploaded = true;
        rnStatus.textContent = `✅ ${file.name} (${result.rows || 0} строк)`;
        rnStatus.style.color = '#4caf50';
        updateButton();
    } catch (err) {
        rnUploaded = false;
        rnStatus.textContent = `❌ Ошибка: ${err.message}`;
        rnStatus.style.color = '#f44336';
    }
});

async function uploadFile(file, type) {
    const formData = new FormData();
    formData.append('file', file);
    
    const endpoint = type === 'bank' ? '/upload-bank' : '/upload-rn';
    const response = await fetch(`${API_URL}${endpoint}`, {
        method: 'POST',
        body: formData
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.detail || error.message || 'Ошибка загрузки');
    }
    
    return await response.json();
}

function updateButton() {
    if (bankUploaded && rnUploaded) {
        reconcileBtn.disabled = false;
        reconcileBtn.textContent = '▶ Запустить сверку';
        statusSpan.textContent = '✅ Файлы загружены, можно запускать сверку';
    } else {
        reconcileBtn.disabled = true;
        reconcileBtn.textContent = '▶ Запустить сверку (ожидание файлов)';
        statusSpan.textContent = '✅ Система готова';
    }
}

// Reconcile handler
reconcileBtn.addEventListener('click', async () => {
    if (!bankUploaded || !rnUploaded) {
        showError('Загрузите оба файла');
        return;
    }
    
    reconcileBtn.disabled = true;
    loading.classList.remove('hidden');
    resultsDiv.classList.add('hidden');
    errorDiv.classList.add('hidden');
    statusSpan.textContent = '⏳ Выполняется сверка...';
    
    try {
        const response = await fetch(`${API_URL}/reconcile`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || error.message || 'Ошибка сверки');
        }
        
        const data = await response.json();
        if (data.status === 'success' && data.results) {
            displayResults(data.results);
            statusSpan.textContent = '✅ Сверка завершена';
        } else {
            throw new Error(data.message || 'Неизвестная ошибка');
        }
    } catch (err) {
        showError(err.message);
        statusSpan.textContent = '❌ Ошибка сверки';
    } finally {
        reconcileBtn.disabled = false;
        loading.classList.add('hidden');
    }
});

function displayResults(results) {
    resultsDiv.classList.remove('hidden');
    
    const summary = results.summary || {};
    const matches = results.matches || [];
    const unmatchedBank = results.unmatched_bank || [];
    const unmatchedRn = results.unmatched_rn || [];
    const linkAnalysis = results.link_analysis || {};
    const baseComparison = results.base_comparison || {};
    
    let html = '<div class="results-grid">';
    
    // Summary
    html += `<div class="result-card">
        <h3> Статистика</h3>
        <div class="stat-row"><span>Банковских платежей:</span> <strong>${results.bank_count || 0}</strong></div>
        <div class="stat-row"><span>Записей РН-Карт:</span> <strong>${results.rn_count || 0}</strong></div>
        <div class="stat-row"><span>Совпадений:</span> <strong>${summary.total_matches || 0}</strong></div>
        <div class="stat-row"><span>Не совпало (банк):</span> <strong>${summary.unmatched_bank || 0}</strong></div>
        <div class="stat-row"><span>Не совпало (РН-Карт):</span> <strong>${summary.unmatched_rn || 0}</strong></div>
        <div class="stat-row"><span>Точность сверки (банк):</span> <strong>${summary.match_rate_bank ? summary.match_rate_bank.toFixed(1) : 0}%</strong></div>
        <div class="stat-row"><span>Точность сверки (РН):</span> <strong>${summary.match_rate_rn ? summary.match_rate_rn.toFixed(1) : 0}%</strong></div>
    </div>`;
    
    // Link analysis
    if (linkAnalysis.by_contragent && Object.keys(linkAnalysis.by_contragent).length > 0) {
        html += `<div class="result-card">
            <h3> Связи по контрагентам</h3>`;
        for (const [contragent, items] of Object.entries(linkAnalysis.by_contragent).slice(0, 10)) {
            const total = items.reduce((sum, i) => sum + i.bank_amount, 0);
            html += `<div class="stat-row"><span>${contragent}:</span> <strong>${items.length} связей, ${total.toFixed(2)} руб.</strong></div>`;
        }
        if (Object.keys(linkAnalysis.by_contragent).length > 10) {
            html += `<div class="stat-row"><span>...</span> <strong>и ещё ${Object.keys(linkAnalysis.by_contragent).length - 10} контрагентов</strong></div>`;
        }
        html += `</div>`;
    }
    
    // Base comparison
    if (baseComparison && baseComparison.total_checked > 0) {
        html += `<div class="result-card">
            <h3> Сравнение с базой</h3>
            <div class="stat-row"><span>Проверено контрагентов:</span> <strong>${baseComparison.total_checked}</strong></div>
            <div class="stat-row"><span>В базе:</span> <strong>${(baseComparison.in_base || []).length}</strong></div>
            <div class="stat-row"><span>Не найдены:</span> <strong>${(baseComparison.not_in_base || []).length}</strong></div>`;
        if (baseComparison.not_in_base && baseComparison.not_in_base.length > 0) {
            html += `<div class="not-found">⚠️ Не в базе: ${baseComparison.not_in_base.slice(0, 5).join(', ')}${baseComparison.not_in_base.length > 5 ? ` и ещё ${baseComparison.not_in_base.length - 5}` : ''}</div>`;
        }
        html += `</div>`;
    }
    
    // Matches list (first 10)
    if (matches.length > 0) {
        html += `<div class="result-card full-width">
            <h3>✅ Найденные совпадения (первые ${Math.min(matches.length, 10)})</h3>
            <table class="matches-table">
                <thead><tr><th>#</th><th>Контрагент</th><th>Сумма (банк)</th><th>Сумма (РН)</th><th>Разница</th></tr></thead>
                <tbody>`;
        matches.slice(0, 10).forEach((match, i) => {
            const bank = match.bank || {};
            const rn = match.rn || {};
            html += `<tr>
                <td>${i + 1}</td>
                <td>${match.contragent || '—'}</td>
                <td>${bank.Amount ? Number(bank.Amount).toFixed(2) : '—'}</td>
                <td>${rn.Amount ? Number(rn.Amount).toFixed(2) : '—'}</td>
                <td>${match.difference !== undefined ? match.difference.toFixed(2) : '—'}</td>
            </tr>`;
        });
        html += `</tbody></table>
            ${matches.length > 10 ? `<div class="more-rows">... и ещё ${matches.length - 10} совпадений</div>` : ''}
        </div>`;
    }
    
    // Unmatched
    if (unmatchedBank.length > 0) {
        html += `<div class="result-card half-width">
            <h3>❌ Не совпало в банке (${unmatchedBank.length})</h3>
            <ul class="unmatched-list">`;
        unmatchedBank.slice(0, 10).forEach((item, i) => {
            html += `<li>${i + 1}. ${item.Contragent || '—'}: ${item.Amount ? Number(item.Amount).toFixed(2) : '—'} руб.</li>`;
        });
        if (unmatchedBank.length > 10) {
            html += `<li class="more-rows">... и ещё ${unmatchedBank.length - 10}</li>`;
        }
        html += `</ul></div>`;
    }
    
    if (unmatchedRn.length > 0) {
        html += `<div class="result-card half-width">
            <h3>❌ Не совпало в РН-Карт (${unmatchedRn.length})</h3>
            <ul class="unmatched-list">`;
        unmatchedRn.slice(0, 10).forEach((item, i) => {
            html += `<li>${i + 1}. ${item.Contragent || '—'}: ${item.Amount ? Number(item.Amount).toFixed(2) : '—'} руб.</li>`;
        });
        if (unmatchedRn.length > 10) {
            html += `<li class="more-rows">... и ещё ${unmatchedRn.length - 10}</li>`;
        }
        html += `</ul></div>`;
    }
    
    html += '</div>';
    resultContent.innerHTML = html;
}

function showError(message) {
    errorDiv.textContent = '❌ ' + message;
    errorDiv.classList.remove('hidden');
    setTimeout(() => errorDiv.classList.add('hidden'), 8000);
}

// Init
updateButton();
console.log('Frontend ready');
