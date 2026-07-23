// =====================================================================
// NUFLO AVAILABILITY & BILLING DASHBOARD — PT MOBILKOM
// =====================================================================

// ---------------------- GLOBAL STATE ----------------------
const appState = {
    allData: [],            // seluruh hasil grouping (per sumur per tanggal), tidak berubah setelah upload
    defaultPrice: 18000,
    customPrices: {},       // { nodeID: harga }
    filters: { search: '', well: 'ALL', status: 'ALL' },
    sort: { key: 'date', dir: 'asc' },
    chart: null,
    uploadMode: 'new',      // 'new' atau 'append'
    uploadHistory: []       // riwayat tiap file yang diproses
};

const REQUIRED_COLS = ['NodeID', 'Date', 'ScanTime', 'Flowrate', 'InternalTemperature', 'SupplyVoltage', 'BatteryVoltage'];
const AVAILABILITY_THRESHOLD = 90;
const MINUTES_PER_DAY = 1440;

// ---------------------- DOM SHORTCUTS ----------------------
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const fileNameDisplay = document.getElementById('file-name-display');
const processBtn = document.getElementById('process-btn');
const uploadError = document.getElementById('upload-error');
const loadingOverlay = document.getElementById('loading-overlay');
const loadingMessage = document.getElementById('loading-message');
const progressBar = document.getElementById('progress-bar');
const dashboardSection = document.getElementById('dashboard-section');
const tableBody = document.getElementById('table-body');

let selectedFile = null;

// ---------------------- UPLOAD UI ----------------------
dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFileSelected(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFileSelected(e.target.files[0]);
});

function handleFileSelected(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xls', 'xlsx', 'csv'].includes(ext)) {
        showUploadError(`Format ".${ext}" tidak didukung. Gunakan file .xls, .xlsx, atau .csv.`);
        return;
    }
    selectedFile = file;
    uploadError.classList.add('hidden');
    fileNameDisplay.textContent = `File terpilih: ${file.name}`;
    fileInfo.classList.remove('hidden');
}

document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.mode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        appState.uploadMode = btn.dataset.mode;
    });
});

function showUploadError(msg) {
    uploadError.textContent = `⚠ ${msg}`;
    uploadError.classList.remove('hidden');
}

processBtn.addEventListener('click', () => {
    if (selectedFile) processFile(selectedFile);
});

// ---------------------- HELPER: NORMALISASI NAMA KOLOM ----------------------
function normalizeColName(s) {
    return String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Cari nama kolom asli di dalam baris sampel, cocokkan dengan nama kolom baku.
// Pakai startsWith untuk InternalTemperature karena suka ada embel-embel satuan, mis. "InternalTemperature(°C)"
function buildColumnMap(sampleRow) {
    const keys = Object.keys(sampleRow || {});
    const map = {};
    const matchers = {
        NodeID: k => normalizeColName(k) === 'nodeid',
        Date: k => normalizeColName(k) === 'date',
        ScanTime: k => normalizeColName(k) === 'scantime',
        Flowrate: k => normalizeColName(k) === 'flowrate',
        InternalTemperature: k => normalizeColName(k).startsWith('internaltemperature'),
        SupplyVoltage: k => normalizeColName(k).startsWith('supplyvoltage'),
        BatteryVoltage: k => normalizeColName(k).startsWith('batteryvoltage')
    };
    for (const [field, test] of Object.entries(matchers)) {
        const found = keys.find(test);
        if (found) map[field] = found;
    }
    return map;
}

// ---------------------- HELPER: PARSING TANGGAL & JAM (aman terhadap timezone) ----------------------
function excelSerialToDate(serial) {
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + Math.floor(serial) * 86400000);
}

function parseDateFlexible(value) {
    if (value instanceof Date) {
        return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    }
    if (typeof value === 'number') {
        return excelSerialToDate(value);
    }
    if (typeof value === 'string') {
        const trimmed = value.trim();
        let m = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/); // M/D/YYYY
        if (m) return new Date(Date.UTC(parseInt(m[3], 10), parseInt(m[1], 10) - 1, parseInt(m[2], 10)));
        m = trimmed.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/); // YYYY-MM-DD
        if (m) return new Date(Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)));
    }
    return null;
}

function formatDateStr(d) {
    if (!d) return 'Unknown Date';
    const y = d.getUTCFullYear();
    const mo = String(d.getUTCMonth() + 1).padStart(2, '0');
    const da = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${mo}-${da}`;
}

function excelTimeToHHMM(value) {
    let fraction = null;
    if (typeof value === 'number') fraction = value;
    else if (typeof value === 'string' && /^\d{1,2}:\d{2}/.test(value.trim())) return value.trim();
    if (fraction === null) return '-';
    const totalMinutes = Math.round((fraction % 1) * MINUTES_PER_DAY);
    const hh = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
    const mm = String(totalMinutes % 60).padStart(2, '0');
    return `${hh}:${mm}`;
}

function formatRupiah(n) {
    return 'Rp ' + Math.round(n).toLocaleString('id-ID');
}

// ---------------------- PROSES FILE ----------------------
async function processFile(file) {
    fileInfo.classList.add('hidden');
    loadingOverlay.classList.remove('hidden');

    try {
        await setLoadingStep('Membaca file...', 15);
        const ext = file.name.split('.').pop().toLowerCase();
        let workbook;
        if (ext === 'csv') {
            const text = await file.text();
            workbook = XLSX.read(text, { type: 'string' });
        } else {
            const buffer = await file.arrayBuffer();
            workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
        }

        await setLoadingStep('Mendeteksi sheet & kolom...', 35);

        const usedSheetNames = [];
        const sheetDiagnostics = [];
        let rawData = [];

        for (const sheetName of workbook.SheetNames) {
            const ws = workbook.Sheets[sheetName];
            const sample = XLSX.utils.sheet_to_json(ws, { defval: null });
            const cols = sample.length ? Object.keys(sample[0]) : [];
            const colMap = buildColumnMap(sample[0] || {});
            sheetDiagnostics.push({ name: sheetName, rowCount: sample.length, columns: cols, colMap });

            if (colMap.NodeID && colMap.Date) {
                usedSheetNames.push(sheetName);
                for (const row of sample) {
                    rawData.push({
                        NodeID: row[colMap.NodeID],
                        Date: row[colMap.Date],
                        ScanTime: colMap.ScanTime ? row[colMap.ScanTime] : null,
                        Flowrate: colMap.Flowrate ? row[colMap.Flowrate] : null,
                        InternalTemperature: colMap.InternalTemperature ? row[colMap.InternalTemperature] : null,
                        SupplyVoltage: colMap.SupplyVoltage ? row[colMap.SupplyVoltage] : null,
                        BatteryVoltage: colMap.BatteryVoltage ? row[colMap.BatteryVoltage] : null
                    });
                }
            }
        }

        if (usedSheetNames.length === 0) {
            const sheetList = sheetDiagnostics.map(s => `"${s.name}" (kolom: ${s.columns.join(', ') || '-'})`).join(' | ');
            throw new Error(`Tidak ada sheet yang punya kolom 'NodeID' dan 'Date'. Sheet yang terbaca: ${sheetList}`);
        }

        rawData = rawData.filter(r => r.NodeID !== null && r.NodeID !== '' && r.Date !== null && r.Date !== '');
        if (rawData.length === 0) throw new Error('Sheet ditemukan tapi tidak ada baris data valid (NodeID/Date kosong semua).');

        await setLoadingStep('Mengelompokkan per sumur & tanggal...', 60);
        const grouped = groupData(rawData);

        await setLoadingStep('Menghitung availability & rata-rata...', 80);
        const processedData = computeStats(grouped);

        await setLoadingStep('Menyusun laporan...', 100);

        const isFirstUpload = appState.allData.length === 0;
        const mode = isFirstUpload ? 'new' : appState.uploadMode;

        let mergeReport = null;
        if (mode === 'new') {
            appState.allData = processedData;
            mergeReport = { mode: 'new', added: processedData.length, updated: 0, skipped: 0, details: [] };
        } else {
            mergeReport = mergeIntoExistingData(processedData);
        }

        appState.sheetDiagnostics = sheetDiagnostics;
        appState.usedSheetNames = usedSheetNames;
        appState.validRowCount = rawData.length;
        appState.uploadHistory.push({
            fileName: file.name,
            timestamp: new Date(),
            mode,
            added: mergeReport.added,
            updated: mergeReport.updated,
            skipped: mergeReport.skipped
        });

        populateWellDropdowns();
        renderDiagnosticPanel();
        renderUploadHistory();
        applyFiltersAndRender();

        loadingOverlay.classList.add('hidden');
        dashboardSection.classList.remove('hidden');
        document.getElementById('upload-mode-toggle').classList.remove('hidden');
        fileInput.value = '';
        selectedFile = null;

        if (mode === 'append') {
            showMergeReportModal(mergeReport);
        } else {
            showToast(`✅ Data berhasil diproses: ${processedData.length} baris (sumur × tanggal).`, 'success');
        }

    } catch (err) {
        loadingOverlay.classList.add('hidden');
        fileInfo.classList.remove('hidden');
        showUploadError(err.message);
    }
}

// Membandingkan 2 hasil hitungan (grup sumur+tanggal) apakah benar-benar identik
function isGroupIdentical(a, b) {
    return a.rowCount === b.rowCount &&
        Math.abs(a.sumFlowrate - b.sumFlowrate) < 0.0001 &&
        Math.abs(a.sumTemp - b.sumTemp) < 0.0001 &&
        Math.abs(a.sumSupplyV - b.sumSupplyV) < 0.0001 &&
        Math.abs(a.sumBatteryV - b.sumBatteryV) < 0.0001;
}

// Gabungkan hasil hitungan baru ke data yang sudah ada:
// - Tanggal baru (belum ada sebelumnya)   -> DITAMBAHKAN
// - Tanggal sudah ada & datanya identik   -> DILEWATI (skip)
// - Tanggal sudah ada & datanya beda      -> DIPERBARUI (update/replace)
function mergeIntoExistingData(newGroups) {
    const existingMap = new Map(appState.allData.map(g => [`${g.nodeID}_${g.date}`, g]));
    const details = [];
    let added = 0, updated = 0, skipped = 0;

    for (const newGroup of newGroups) {
        const key = `${newGroup.nodeID}_${newGroup.date}`;
        const existing = existingMap.get(key);

        if (!existing) {
            existingMap.set(key, newGroup);
            added++;
            details.push({ type: 'added', nodeID: newGroup.nodeID, date: newGroup.date, info: `${newGroup.rowCount} baris` });
        } else if (isGroupIdentical(existing, newGroup)) {
            skipped++;
            details.push({ type: 'skipped', nodeID: newGroup.nodeID, date: newGroup.date, info: `${existing.rowCount} baris (sama persis, tidak dihitung ulang)` });
        } else {
            existingMap.set(key, newGroup);
            updated++;
            details.push({ type: 'updated', nodeID: newGroup.nodeID, date: newGroup.date, info: `${existing.rowCount} baris → ${newGroup.rowCount} baris` });
        }
    }

    appState.allData = Array.from(existingMap.values()).sort((a, b) => a.nodeID.localeCompare(b.nodeID) || a.date.localeCompare(b.date));
    return { mode: 'append', added, updated, skipped, details };
}

function setLoadingStep(message, percent) {
    loadingMessage.textContent = message;
    progressBar.style.width = percent + '%';
    return new Promise(resolve => setTimeout(resolve, 220));
}

function groupData(rawData) {
    const grouped = {};
    for (const row of rawData) {
        const nodeID = String(row.NodeID).trim();
        const dateObj = parseDateFlexible(row.Date);
        const dateStr = formatDateStr(dateObj);
        const key = `${nodeID}_${dateStr}`;

        if (!grouped[key]) {
            grouped[key] = {
                nodeID, date: dateStr, rowCount: 0,
                sumFlowrate: 0, sumTemp: 0, sumSupplyV: 0, sumBatteryV: 0,
                rawRows: []
            };
        }
        const g = grouped[key];
        g.rowCount++;
        g.sumFlowrate += parseFloat(row.Flowrate) || 0;
        g.sumTemp += parseFloat(row.InternalTemperature) || 0;
        g.sumSupplyV += parseFloat(row.SupplyVoltage) || 0;
        g.sumBatteryV += parseFloat(row.BatteryVoltage) || 0;
        g.rawRows.push({
            jam: excelTimeToHHMM(row.ScanTime),
            flowrate: row.Flowrate,
            temp: row.InternalTemperature,
            supplyV: row.SupplyVoltage,
            batteryV: row.BatteryVoltage
        });
    }
    return grouped;
}

function computeStats(grouped) {
    return Object.values(grouped).map(g => {
        const availability = (g.rowCount / MINUTES_PER_DAY) * 100;
        const status = availability >= AVAILABILITY_THRESHOLD ? 1 : 0;
        return {
            nodeID: g.nodeID,
            date: g.date,
            rowCount: g.rowCount,
            availability,
            status,
            avgFlowrate: g.sumFlowrate / g.rowCount,
            avgTemp: g.sumTemp / g.rowCount,
            avgSupplyV: g.sumSupplyV / g.rowCount,
            avgBatteryV: g.sumBatteryV / g.rowCount,
            sumFlowrate: g.sumFlowrate,
            sumTemp: g.sumTemp,
            sumSupplyV: g.sumSupplyV,
            sumBatteryV: g.sumBatteryV,
            rawRows: g.rawRows
        };
    }).sort((a, b) => a.nodeID.localeCompare(b.nodeID) || a.date.localeCompare(b.date));
}

// ---------------------- HARGA ----------------------
function getPriceForWell(nodeID) {
    return appState.customPrices[nodeID] ?? appState.defaultPrice;
}

function getBilling(row) {
    return row.status === 1 ? getPriceForWell(row.nodeID) : 0;
}

// ---------------------- DIAGNOSTIC PANEL ----------------------
function renderDiagnosticPanel() {
    const body = document.getElementById('diagnostic-body');
    const { sheetDiagnostics, usedSheetNames, validRowCount } = appState;
    let html = `<p>File kamu punya <b>${sheetDiagnostics.length}</b> sheet. Sheet yang digabung untuk perhitungan: <b>${usedSheetNames.map(n => `"${n}"`).join(', ')}</b>.</p>`;

    sheetDiagnostics.forEach(s => {
        const isUsed = usedSheetNames.includes(s.name);
        html += `<h4 style="margin-top:0.9rem;">${isUsed ? '✅ Dipakai' : '⏭️ Dilewati'} — Sheet "${s.name}" (${s.rowCount} baris terbaca)</h4>`;
        html += `<table><thead><tr><th>Kolom dibutuhkan</th><th>Status</th></tr></thead><tbody>`;
        REQUIRED_COLS.forEach(col => {
            const found = !!s.colMap[col];
            html += `<tr><td><code>${col}</code></td><td class="${found ? 'tag-ok' : 'tag-miss'}">${found ? 'Ditemukan (' + s.colMap[col] + ')' : 'Tidak ada'}</td></tr>`;
        });
        html += `</tbody></table>`;
    });

    html += `<p><b>Total baris valid dipakai:</b> ${validRowCount} baris.</p>`;
    body.innerHTML = html;
}

document.getElementById('diagnostic-toggle').addEventListener('click', () => {
    const body = document.getElementById('diagnostic-body');
    const arrow = document.getElementById('diagnostic-arrow');
    body.classList.toggle('hidden');
    arrow.textContent = body.classList.contains('hidden') ? '▼' : '▲';
});

// ---------------------- TOAST NOTIFICATION ----------------------
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('show'));
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 5000);
}

// ---------------------- MODAL LAPORAN PENGGABUNGAN ----------------------
function showMergeReportModal(report) {
    const body = document.getElementById('merge-modal-body');
    const iconFor = t => t === 'added' ? '🟢' : t === 'updated' ? '🟡' : '⚪';
    const labelFor = t => t === 'added' ? 'Ditambahkan (tanggal baru)' : t === 'updated' ? 'Diperbarui (data berubah)' : 'Dilewati (sudah sama persis)';

    let html = `
        <div class="merge-summary-grid">
            <div class="merge-summary-item merge-added"><h2>${report.added}</h2><p>Ditambahkan</p></div>
            <div class="merge-summary-item merge-updated"><h2>${report.updated}</h2><p>Diperbarui</p></div>
            <div class="merge-summary-item merge-skipped"><h2>${report.skipped}</h2><p>Dilewati</p></div>
        </div>
        <p class="merge-explain">Data lama tetap aman. Sistem hanya menambah tanggal baru, memperbarui tanggal yang datanya berubah, dan melewati tanggal yang sudah persis sama — tidak ada yang dihitung dua kali.</p>
    `;

    if (report.details.length > 0) {
        html += `<div class="verify-table-wrap"><table>
            <thead><tr><th>Status</th><th>Nama Sumur</th><th>Tanggal</th><th>Keterangan</th></tr></thead>
            <tbody>${report.details.map(d => `
                <tr>
                    <td>${iconFor(d.type)} ${labelFor(d.type)}</td>
                    <td>${d.nodeID}</td>
                    <td>${d.date}</td>
                    <td>${d.info}</td>
                </tr>
            `).join('')}</tbody>
        </table></div>`;
    }

    body.innerHTML = html;
    document.getElementById('merge-modal').classList.remove('hidden');
}
document.getElementById('merge-modal-close').addEventListener('click', () => document.getElementById('merge-modal').classList.add('hidden'));
document.getElementById('merge-modal').addEventListener('click', (e) => {
    if (e.target.id === 'merge-modal') e.target.classList.add('hidden');
});

// ---------------------- RIWAYAT UPLOAD ----------------------
function renderUploadHistory() {
    const panel = document.getElementById('upload-history-panel');
    const body = document.getElementById('history-body');
    if (appState.uploadHistory.length === 0) { panel.classList.add('hidden'); return; }
    panel.classList.remove('hidden');

    body.innerHTML = `<table><thead><tr><th>Waktu</th><th>Nama File</th><th>Mode</th><th>Ditambah</th><th>Diperbarui</th><th>Dilewati</th></tr></thead><tbody>
        ${appState.uploadHistory.map(h => `
            <tr>
                <td>${h.timestamp.toLocaleTimeString('id-ID')}</td>
                <td>${h.fileName}</td>
                <td>${h.mode === 'new' ? '🆕 Upload Baru' : '➕ Tambah Data'}</td>
                <td>${h.added}</td><td>${h.updated}</td><td>${h.skipped}</td>
            </tr>
        `).join('')}
    </tbody></table>`;
}
document.getElementById('history-toggle').addEventListener('click', () => {
    const body = document.getElementById('history-body');
    const arrow = document.getElementById('history-arrow');
    body.classList.toggle('hidden');
    arrow.textContent = body.classList.contains('hidden') ? '▼' : '▲';
});

// ---------------------- DROPDOWN SUMUR ----------------------
function populateWellDropdowns() {
    const wells = [...new Set(appState.allData.map(r => r.nodeID))].sort();

    const filterSelect = document.getElementById('well-filter-select');
    filterSelect.innerHTML = '<option value="ALL">Semua Sumur</option>' + wells.map(w => `<option value="${w}">${w}</option>`).join('');

    const priceSelect = document.getElementById('price-well-select');
    priceSelect.innerHTML = wells.map(w => `<option value="${w}">${w}</option>`).join('');
}

// ---------------------- FILTER BAR EVENTS ----------------------
document.getElementById('search-input').addEventListener('input', (e) => {
    appState.filters.search = e.target.value.trim().toLowerCase();
    applyFiltersAndRender();
});
document.getElementById('well-filter-select').addEventListener('change', (e) => {
    appState.filters.well = e.target.value;
    applyFiltersAndRender();
});
document.querySelectorAll('.status-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.status-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        appState.filters.status = btn.dataset.status;
        applyFiltersAndRender();
    });
});
document.getElementById('reset-filter-btn').addEventListener('click', () => {
    appState.filters = { search: '', well: 'ALL', status: 'ALL' };
    document.getElementById('search-input').value = '';
    document.getElementById('well-filter-select').value = 'ALL';
    document.querySelectorAll('.status-btn').forEach(b => b.classList.remove('active'));
    document.querySelector('.status-btn[data-status="ALL"]').classList.add('active');
    applyFiltersAndRender();
});

function getFilteredData() {
    const { search, well, status } = appState.filters;
    return appState.allData.filter(row => {
        if (well !== 'ALL' && row.nodeID !== well) return false;
        if (status !== 'ALL' && String(row.status) !== status) return false;
        if (search) {
            const hay = (row.nodeID + ' ' + row.date).toLowerCase();
            if (!hay.includes(search)) return false;
        }
        return true;
    });
}

function applyFiltersAndRender() {
    const filtered = getFilteredData();
    renderSummaryCards(filtered);
    renderTable(sortData(filtered));
    renderChart(filtered);
    renderInvoiceSummary();
}

// ---------------------- SORTING ----------------------
document.querySelectorAll('#main-table thead th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
        const key = th.dataset.sort;
        if (appState.sort.key === key) {
            appState.sort.dir = appState.sort.dir === 'asc' ? 'desc' : 'asc';
        } else {
            appState.sort.key = key;
            appState.sort.dir = 'asc';
        }
        applyFiltersAndRender();
    });
});

function sortData(data) {
    const { key, dir } = appState.sort;
    const factor = dir === 'asc' ? 1 : -1;
    return [...data].sort((a, b) => {
        let av = key === 'billing' ? getBilling(a) : a[key];
        let bv = key === 'billing' ? getBilling(b) : b[key];
        if (typeof av === 'string') return av.localeCompare(bv) * factor;
        return (av - bv) * factor;
    });
}

// ---------------------- SUMMARY CARDS ----------------------
function renderSummaryCards(data) {
    const wells = new Set(data.map(r => r.nodeID));
    const avgAvail = data.length ? data.reduce((s, r) => s + r.availability, 0) / data.length : 0;
    const green = data.filter(r => r.status === 1).length;
    const red = data.filter(r => r.status === 0).length;
    const totalBilling = data.reduce((s, r) => s + getBilling(r), 0);

    document.getElementById('sum-total-wells').textContent = wells.size;
    document.getElementById('sum-avg-availability').textContent = avgAvail.toFixed(2) + '%';
    document.getElementById('sum-status-green').textContent = green;
    document.getElementById('sum-status-red').textContent = red;
    document.getElementById('sum-total-billing').textContent = formatRupiah(totalBilling);
}

// ---------------------- TABLE ----------------------
function renderTable(data) {
    tableBody.innerHTML = '';
    const emptyMsg = document.getElementById('table-empty-msg');

    if (data.length === 0) {
        emptyMsg.classList.remove('hidden');
        return;
    }
    emptyMsg.classList.add('hidden');

    data.forEach((row) => {
        const tr = document.createElement('tr');
        tr.className = row.status === 1 ? 'status-green' : 'status-red';
        const billing = getBilling(row);
        const badge = row.status === 1
            ? '<span class="status-badge status-badge-green">🟢 HIJAU</span>'
            : '<span class="status-badge status-badge-red">🔴 MERAH</span>';

        tr.innerHTML = `
            <td>${row.nodeID}</td>
            <td>${row.date}</td>
            <td>${row.rowCount}</td>
            <td>${row.availability.toFixed(2)}%</td>
            <td>${badge}</td>
            <td>${row.avgFlowrate.toFixed(2)}</td>
            <td>${row.avgTemp.toFixed(2)}</td>
            <td>${row.avgSupplyV.toFixed(2)}</td>
            <td>${row.avgBatteryV.toFixed(2)}</td>
            <td>${formatRupiah(billing)}</td>
            <td><button class="btn-verify" data-node="${row.nodeID}" data-date="${row.date}">Cek Manual</button></td>
        `;
        tableBody.appendChild(tr);
    });

    tableBody.querySelectorAll('.btn-verify').forEach(btn => {
        btn.addEventListener('click', () => {
            const row = appState.allData.find(r => r.nodeID === btn.dataset.node && r.date === btn.dataset.date);
            openVerifyModal(row);
        });
    });
}

// ---------------------- CHART ----------------------
function renderChart(data) {
    const ctx = document.getElementById('availability-chart').getContext('2d');
    const title = document.getElementById('chart-title');
    if (appState.chart) appState.chart.destroy();

    const wellFilter = appState.filters.well;

    if (wellFilter !== 'ALL') {
        // Line chart trend per tanggal untuk 1 sumur
        title.textContent = `📈 Trend Availability Harian — ${wellFilter}`;
        const sorted = [...data].sort((a, b) => a.date.localeCompare(b.date));
        const gradient = ctx.createLinearGradient(0, 0, 0, 320);
        gradient.addColorStop(0, 'rgba(10,44,92,0.35)');
        gradient.addColorStop(1, 'rgba(10,44,92,0.02)');

        appState.chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: sorted.map(r => r.date),
                datasets: [{
                    label: 'Availability (%)',
                    data: sorted.map(r => r.availability),
                    borderColor: '#0a2c5c',
                    backgroundColor: gradient,
                    fill: true,
                    tension: 0.35,
                    pointRadius: 6,
                    pointHoverRadius: 8,
                    pointBackgroundColor: sorted.map(r => r.status === 1 ? '#2e7d32' : '#c62828'),
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    borderWidth: 3
                }, {
                    label: 'Batas Minimum (90%)',
                    data: sorted.map(() => 90),
                    borderColor: '#d4a017',
                    borderDash: [8, 4],
                    borderWidth: 2,
                    pointRadius: 0,
                    fill: false
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => ctx.dataset.label === 'Availability (%)'
                                ? `Availability: ${ctx.parsed.y.toFixed(2)}%`
                                : 'Batas minimum: 90%'
                        }
                    }
                },
                scales: { y: { min: 0, max: 105, ticks: { callback: v => v + '%' } } }
            }
        });
    } else {
        // Bar chart perbandingan rata-rata availability antar sumur
        title.textContent = '📊 Perbandingan Rata-rata Availability antar Sumur';
        const wellMap = {};
        data.forEach(r => {
            if (!wellMap[r.nodeID]) wellMap[r.nodeID] = [];
            wellMap[r.nodeID].push(r.availability);
        });
        const wells = Object.keys(wellMap).sort();
        const avgs = wells.map(w => wellMap[w].reduce((a, b) => a + b, 0) / wellMap[w].length);

        appState.chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: wells,
                datasets: [{
                    label: 'Rata-rata Availability (%)',
                    data: avgs,
                    backgroundColor: avgs.map(v => v >= 90 ? 'rgba(46,125,50,0.75)' : 'rgba(198,40,40,0.75)'),
                    borderRadius: 6,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { callbacks: { label: (ctx) => `${ctx.parsed.y.toFixed(2)}%` } }
                },
                scales: { y: { min: 0, max: 105, ticks: { callback: v => v + '%' } } }
            }
        });
    }
}

// ---------------------- VERIFY MODAL ----------------------
function openVerifyModal(row) {
    document.getElementById('verify-title').textContent = `Verifikasi Manual — ${row.nodeID} / ${row.date}`;
    const billing = getBilling(row);
    const price = getPriceForWell(row.nodeID);

    const previewRows = row.rawRows.slice(0, 15);
    const tableRowsHtml = previewRows.map((r, idx) => `
        <tr>
            <td>${idx + 1}</td><td>${r.jam}</td><td>${r.flowrate ?? '-'}</td>
            <td>${r.temp ?? '-'}</td><td>${r.supplyV ?? '-'}</td><td>${r.batteryV ?? '-'}</td>
        </tr>
    `).join('');

    document.getElementById('verify-body').innerHTML = `
        <div class="verify-formula"><b>Availability</b>
            <code>(${row.rowCount} / 1440) × 100 = ${row.availability.toFixed(4)}%</code>
        </div>
        <div class="verify-formula"><b>Status</b>
            <code>${row.availability.toFixed(2)}% ${row.status === 1 ? '≥' : '<'} 90% → Status ${row.status} (${row.status === 1 ? 'Bisa Ditagih' : 'Belum Bisa Ditagih'})</code>
        </div>
        <div class="verify-formula"><b>Rata-rata</b>
            <code>Flowrate = ${row.sumFlowrate.toFixed(2)} / ${row.rowCount} = ${row.avgFlowrate.toFixed(4)}
Temp     = ${row.sumTemp.toFixed(2)} / ${row.rowCount} = ${row.avgTemp.toFixed(4)}
SupplyV  = ${row.sumSupplyV.toFixed(2)} / ${row.rowCount} = ${row.avgSupplyV.toFixed(4)}
BatteryV = ${row.sumBatteryV.toFixed(2)} / ${row.rowCount} = ${row.avgBatteryV.toFixed(4)}</code>
        </div>
        <div class="verify-formula"><b>Tagihan</b>
            <code>Harga/hari = ${formatRupiah(price)} → Tagihan = ${formatRupiah(billing)}</code>
        </div>
        <p><b>Contoh ${previewRows.length} dari ${row.rowCount} baris mentah:</b></p>
        <div class="verify-table-wrap">
            <table><thead><tr><th>#</th><th>Jam</th><th>Flowrate</th><th>Temp</th><th>SupplyV</th><th>BatteryV</th></tr></thead>
            <tbody>${tableRowsHtml}</tbody></table>
        </div>
    `;
    document.getElementById('verify-modal').classList.remove('hidden');
}
document.getElementById('verify-close').addEventListener('click', () => document.getElementById('verify-modal').classList.add('hidden'));
document.getElementById('verify-modal').addEventListener('click', (e) => {
    if (e.target.id === 'verify-modal') e.target.classList.add('hidden');
});

// ---------------------- PANEL HARGA ----------------------
document.getElementById('default-price-input').addEventListener('change', (e) => {
    const val = parseFloat(e.target.value);
    if (!isNaN(val) && val >= 0) {
        appState.defaultPrice = val;
        applyFiltersAndRender();
    }
});

document.getElementById('apply-custom-price-btn').addEventListener('click', () => {
    const well = document.getElementById('price-well-select').value;
    const priceInput = document.getElementById('custom-price-input');
    const val = parseFloat(priceInput.value);
    if (!well || isNaN(val) || val < 0) return;
    appState.customPrices[well] = val;
    priceInput.value = '';
    renderCustomPriceChips();
    applyFiltersAndRender();
});

function renderCustomPriceChips() {
    const container = document.getElementById('custom-price-list');
    const entries = Object.entries(appState.customPrices);
    container.innerHTML = entries.map(([well, price]) => `
        <span class="price-chip">${well}: ${formatRupiah(price)} <button data-well="${well}">&times;</button></span>
    `).join('');
    container.querySelectorAll('button').forEach(btn => {
        btn.addEventListener('click', () => {
            delete appState.customPrices[btn.dataset.well];
            renderCustomPriceChips();
            applyFiltersAndRender();
        });
    });
}

// ---------------------- RINGKASAN TAGIHAN PER SUMUR ----------------------
function renderInvoiceSummary() {
    const wells = [...new Set(appState.allData.map(r => r.nodeID))].sort();
    const tbody = document.getElementById('invoice-table-body');
    const tfoot = document.getElementById('invoice-table-foot');

    let grandTotalDays = 0, grandGreen = 0, grandRed = 0, grandBilling = 0;

    tbody.innerHTML = wells.map((well, idx) => {
        const rows = appState.allData.filter(r => r.nodeID === well);
        const totalDays = rows.length;
        const greenDays = rows.filter(r => r.status === 1).length;
        const redDays = totalDays - greenDays;
        const price = getPriceForWell(well);
        const totalBilling = greenDays * price;

        grandTotalDays += totalDays; grandGreen += greenDays; grandRed += redDays; grandBilling += totalBilling;

        return `<tr>
            <td>${idx + 1}</td><td>${well}</td><td>${totalDays}</td>
            <td>${greenDays}</td><td>${redDays}</td>
            <td>${formatRupiah(price)}</td><td>${formatRupiah(totalBilling)}</td>
        </tr>`;
    }).join('');

    tfoot.innerHTML = `<tr>
        <td colspan="2">GRAND TOTAL</td><td>${grandTotalDays}</td>
        <td>${grandGreen}</td><td>${grandRed}</td><td>-</td><td>${formatRupiah(grandBilling)}</td>
    </tr>`;
}

// ---------------------- EXPORT EXCEL (ExcelJS — rapi, berwarna, ada rumus) ----------------------
document.getElementById('export-excel-btn').addEventListener('click', exportToExcel);

async function exportToExcel() {
    const data = sortData(getFilteredData());
    if (data.length === 0) { alert('Tidak ada data untuk di-export.'); return; }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'PT Mobilkom Telekomindo';
    workbook.created = new Date();

    // ============ SHEET 1: DETAIL HARIAN ============
    const ws = workbook.addWorksheet('Detail Harian');
    ws.mergeCells('A1:K1');
    ws.getCell('A1').value = 'NUFLO AVAILABILITY & BILLING REPORT';
    ws.getCell('A1').font = { bold: true, size: 15, color: { argb: 'FF0A2C5C' } };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    ws.mergeCells('A2:K2');
    ws.getCell('A2').value = 'PT MOBILKOM TELEKOMINDO — Laporan Detail Availability Sumur';
    ws.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF667085' } };
    ws.getCell('A2').alignment = { horizontal: 'center' };

    ws.mergeCells('A3:K3');
    ws.getCell('A3').value = `Dicetak: ${new Date().toLocaleString('id-ID')}`;
    ws.getCell('A3').font = { size: 9, color: { argb: 'FF667085' } };
    ws.getCell('A3').alignment = { horizontal: 'center' };

    const headerRowIdx = 5;
    const headers = ['Nama Sumur', 'Tanggal', 'Jml Baris', 'Availability (%)', 'Status', 'Avg Flowrate', 'Avg Temp (°C)', 'Avg Supply V', 'Avg Battery V', 'Harga/Hari (IDR)', 'Tagihan (IDR)'];
    const headerRow = ws.getRow(headerRowIdx);
    headerRow.values = headers;
    headerRow.eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A2C5C' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });

    const firstDataRow = headerRowIdx + 1;
    data.forEach((row, i) => {
        const price = getPriceForWell(row.nodeID);
        const billing = getBilling(row);
        const r = ws.getRow(firstDataRow + i);
        r.values = [
            row.nodeID, row.date, row.rowCount, row.availability,
            row.status, row.avgFlowrate, row.avgTemp, row.avgSupplyV, row.avgBatteryV,
            price, billing
        ];
        // Format availability sebagai angka + simbol % literal (BUKAN format Percentage bawaan Excel,
        // supaya tidak dikali 100 lagi saat dibuka / diubah formatnya)
        r.getCell(4).numFmt = '0.00"%"';
        r.getCell(6).numFmt = '#,##0.00';
        r.getCell(7).numFmt = '#,##0.00';
        r.getCell(8).numFmt = '#,##0.00';
        r.getCell(9).numFmt = '#,##0.00';
        r.getCell(10).numFmt = '#,##0';
        r.getCell(11).numFmt = '#,##0';

        const fillColor = row.status === 1 ? 'FFC8E6C9' : 'FFFFCDD2';
        r.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            cell.border = { top: { style: 'thin', color: { argb: 'FFD0D7E2' } }, bottom: { style: 'thin', color: { argb: 'FFD0D7E2' } }, left: { style: 'thin', color: { argb: 'FFD0D7E2' } }, right: { style: 'thin', color: { argb: 'FFD0D7E2' } } };
        });
        r.getCell(5).value = row.status === 1 ? '1 (Bisa Ditagih)' : '0 (Belum Bisa Ditagih)';
        r.getCell(5).font = { bold: true, color: { argb: row.status === 1 ? 'FF1B5E20' : 'FFB71C1C' } };
    });

    const lastDataRow = firstDataRow + data.length - 1;
    const totalRowIdx = lastDataRow + 1;
    ws.getCell(`A${totalRowIdx}`).value = 'RINGKASAN';
    ws.getCell(`A${totalRowIdx}`).font = { bold: true };
    ws.getCell(`D${totalRowIdx}`).value = { formula: `AVERAGE(D${firstDataRow}:D${lastDataRow})` };
    ws.getCell(`D${totalRowIdx}`).numFmt = '0.00"%"';
    ws.getCell(`F${totalRowIdx}`).value = { formula: `AVERAGE(F${firstDataRow}:F${lastDataRow})` };
    ws.getCell(`F${totalRowIdx}`).numFmt = '#,##0.00';
    ws.getCell(`G${totalRowIdx}`).value = { formula: `AVERAGE(G${firstDataRow}:G${lastDataRow})` };
    ws.getCell(`G${totalRowIdx}`).numFmt = '#,##0.00';
    ws.getCell(`K${totalRowIdx}`).value = { formula: `SUM(K${firstDataRow}:K${lastDataRow})` };
    ws.getCell(`K${totalRowIdx}`).numFmt = '#,##0';
    ws.getRow(totalRowIdx).eachCell(cell => {
        cell.font = { ...(cell.font || {}), bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    });

    ws.columns.forEach(col => { col.width = 16; });
    ws.getColumn(1).width = 14;
    ws.getColumn(5).width = 20;

    // ============ SHEET 2: RINGKASAN & TAGIHAN ============
    const ws2 = workbook.addWorksheet('Ringkasan & Tagihan');
    ws2.mergeCells('A1:G1');
    ws2.getCell('A1').value = 'RINGKASAN TAGIHAN PER SUMUR — PT MOBILKOM TELEKOMINDO';
    ws2.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF0A2C5C' } };
    ws2.getCell('A1').alignment = { horizontal: 'center' };

    const invHeaderIdx = 3;
    const invHeaders = ['No', 'Nama Sumur', 'Total Hari Tercatat', 'Hari Bisa Ditagih (≥90%)', 'Hari Belum Bisa Ditagih (<90%)', 'Harga per Hari (IDR)', 'Total Tagihan (IDR)'];
    ws2.getRow(invHeaderIdx).values = invHeaders;
    ws2.getRow(invHeaderIdx).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A2C5C' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });

    const wells = [...new Set(appState.allData.map(r => r.nodeID))].sort();
    const invFirstRow = invHeaderIdx + 1;
    wells.forEach((well, i) => {
        const rows = appState.allData.filter(r => r.nodeID === well);
        const totalDays = rows.length;
        const greenDays = rows.filter(r => r.status === 1).length;
        const redDays = totalDays - greenDays;
        const price = getPriceForWell(well);
        const rowIdx = invFirstRow + i;
        const r = ws2.getRow(rowIdx);
        r.values = [i + 1, well, totalDays, greenDays, redDays, price];
        r.getCell(7).value = { formula: `D${rowIdx}*F${rowIdx}` };
        r.getCell(6).numFmt = '#,##0';
        r.getCell(7).numFmt = '#,##0';
        r.eachCell(cell => cell.border = { top: { style: 'thin', color: { argb: 'FFD0D7E2' } }, bottom: { style: 'thin', color: { argb: 'FFD0D7E2' } }, left: { style: 'thin', color: { argb: 'FFD0D7E2' } }, right: { style: 'thin', color: { argb: 'FFD0D7E2' } } });
    });

    const invLastRow = invFirstRow + wells.length - 1;
    const invTotalRow = invLastRow + 1;
    ws2.mergeCells(`A${invTotalRow}:C${invTotalRow}`);
    ws2.getCell(`A${invTotalRow}`).value = 'GRAND TOTAL';
    ws2.getCell(`D${invTotalRow}`).value = { formula: `SUM(D${invFirstRow}:D${invLastRow})` };
    ws2.getCell(`E${invTotalRow}`).value = { formula: `SUM(E${invFirstRow}:E${invLastRow})` };
    ws2.getCell(`G${invTotalRow}`).value = { formula: `SUM(G${invFirstRow}:G${invLastRow})` };
    ws2.getCell(`G${invTotalRow}`).numFmt = '#,##0';
    ws2.getRow(invTotalRow).eachCell(cell => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF176' } };
    });

    ws2.columns.forEach(col => { col.width = 20; });
    ws2.getColumn(1).width = 6;
    ws2.getColumn(2).width = 16;

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(buffer, 'Laporan_NUFLO_Mobilkom.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

function downloadBlob(buffer, filename, mime) {
    const blob = new Blob([buffer], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
}

// ---------------------- EXPORT PDF ----------------------
document.getElementById('export-pdf-btn').addEventListener('click', exportToPDF);

function exportToPDF() {
    const data = sortData(getFilteredData());
    if (data.length === 0) { alert('Tidak ada data untuk di-export.'); return; }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape' });

    doc.setFontSize(14);
    doc.setTextColor(10, 44, 92);
    doc.text('NUFLO AVAILABILITY & BILLING REPORT', 148, 12, { align: 'center' });
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text('PT MOBILKOM TELEKOMINDO', 148, 18, { align: 'center' });
    doc.setFontSize(8);
    doc.text(`Dicetak: ${new Date().toLocaleString('id-ID')}`, 148, 23, { align: 'center' });

    const body = data.map(row => [
        row.nodeID, row.date, row.rowCount, row.availability.toFixed(2) + '%',
        row.status === 1 ? 'Bisa Ditagih' : 'Belum Bisa Ditagih',
        row.avgFlowrate.toFixed(2), row.avgTemp.toFixed(2), row.avgSupplyV.toFixed(2), row.avgBatteryV.toFixed(2),
        formatRupiah(getBilling(row))
    ]);

    doc.autoTable({
        startY: 28,
        head: [['Nama Sumur', 'Tanggal', 'Jml Baris', 'Availability', 'Status', 'Avg Flowrate', 'Avg Temp', 'Avg SupplyV', 'Avg BatteryV', 'Tagihan']],
        body,
        headStyles: { fillColor: [10, 44, 92], textColor: 255, fontStyle: 'bold' },
        didParseCell: (hookData) => {
            if (hookData.section === 'body') {
                const row = data[hookData.row.index];
                hookData.cell.styles.fillColor = row.status === 1 ? [200, 230, 201] : [255, 205, 210];
                hookData.cell.styles.textColor = row.status === 1 ? [27, 94, 32] : [183, 28, 28];
            }
        },
        styles: { fontSize: 8, cellPadding: 2 }
    });

    // Ringkasan tagihan per sumur di halaman baru
    doc.addPage();
    doc.setFontSize(13);
    doc.setTextColor(10, 44, 92);
    doc.text('RINGKASAN TAGIHAN PER SUMUR', 148, 15, { align: 'center' });

    const wells = [...new Set(appState.allData.map(r => r.nodeID))].sort();
    let grandBilling = 0;
    const invBody = wells.map((well, idx) => {
        const rows = appState.allData.filter(r => r.nodeID === well);
        const totalDays = rows.length;
        const greenDays = rows.filter(r => r.status === 1).length;
        const redDays = totalDays - greenDays;
        const price = getPriceForWell(well);
        const total = greenDays * price;
        grandBilling += total;
        return [idx + 1, well, totalDays, greenDays, redDays, formatRupiah(price), formatRupiah(total)];
    });
    invBody.push(['', 'GRAND TOTAL', '', '', '', '', formatRupiah(grandBilling)]);

    doc.autoTable({
        startY: 22,
        head: [['No', 'Nama Sumur', 'Total Hari', 'Bisa Ditagih', 'Belum Ditagih', 'Harga/Hari', 'Total Tagihan']],
        body: invBody,
        headStyles: { fillColor: [10, 44, 92], textColor: 255, fontStyle: 'bold' },
        didParseCell: (hookData) => {
            if (hookData.row.index === invBody.length - 1) {
                hookData.cell.styles.fillColor = [255, 241, 118];
                hookData.cell.styles.fontStyle = 'bold';
            }
        }
    });

    doc.save('Laporan_NUFLO_Mobilkom.pdf');
}
