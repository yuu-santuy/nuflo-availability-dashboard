// --- STATE MANAGEMENT ---
let rawData = [];
let processedData = [];
let chartInstance = null;
let sortDirection = 1; 

// --- DOM ELEMENTS ---
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const btnBrowse = document.getElementById('btn-browse');
const fileInfo = document.getElementById('file-info');
const fileNameDisplay = document.getElementById('file-name');
const btnProcess = document.getElementById('btn-process');
const errorMessage = document.getElementById('error-message');

const uploadSection = document.getElementById('upload-section');
const loadingOverlay = document.getElementById('loading-overlay');
const loadingText = document.getElementById('loading-text');
const dashboardSection = document.getElementById('dashboard-section');

// --- UPLOAD HANDLERS ---
btnBrowse.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
        handleFileSelect(e.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFileSelect(e.target.files[0]);
    }
});

let selectedFile = null;
function handleFileSelect(file) {
    const validExtensions = ['xls', 'xlsx'];
    const fileExt = file.name.split('.').pop().toLowerCase();
    
    if (!validExtensions.includes(fileExt)) {
        showError("Format tidak valid! Harap upload file Excel (.xls atau .xlsx).");
        selectedFile = null;
        fileInfo.classList.add('hidden');
        return;
    }
    
    showError("");
    selectedFile = file;
    fileNameDisplay.textContent = `File terpilih: ${file.name}`;
    fileInfo.classList.remove('hidden');
}

function showError(msg) {
    errorMessage.textContent = msg;
    if(msg) errorMessage.classList.remove('hidden');
    else errorMessage.classList.add('hidden');
}

function updateLoading(msg) {
    loadingText.textContent = msg;
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Konversi angka serial tanggal Excel (contoh: 46205) ke objek Date JavaScript.
// Excel menghitung hari sejak 30 Desember 1899.
function excelSerialToDate(serial) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const days = Math.floor(serial);
    return new Date(excelEpoch.getTime() + days * 86400000);
}

// Konversi ScanTime (pecahan hari, contoh 0.588 = jam 14:07) jadi teks jam "HH:MM"
function excelTimeToHHMM(fraction) {
    if (typeof fraction !== 'number') return '-';
    const totalMinutes = Math.round((fraction % 1) * 1440);
    const hh = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
    const mm = String(totalMinutes % 60).padStart(2, '0');
    return `${hh}:${mm}`;
}

// --- LOGIKA PEMROSESAN UTAMA ---
btnProcess.addEventListener('click', async () => {
    if (!selectedFile) return;
    
    uploadSection.classList.add('hidden');
    loadingOverlay.classList.remove('hidden');

    try {
        updateLoading("Membaca file Excel...");
        await sleep(100); // Memberi waktu UI untuk render teks
        
        const dataBuffer = await selectedFile.arrayBuffer();
        const workbook = XLSX.read(dataBuffer, { type: 'array', cellDates: false });

        // --- PERBAIKAN PENTING ---
        // File NUFLO bisa punya LEBIH DARI 1 sheet, dan bisa jadi LEBIH DARI SATU
        // sheet yang berisi data valid (misal data dipecah per bulan/batch:
        // Sheet2, Sheet3, Sheet4, dst — semua punya kolom NodeID & Date).
        // Contoh: "Sheet1" biasanya cuma ringkasan status terakhir (TIDAK
        // punya kolom NodeID/Date), sedangkan sheet lain berisi data mentah.
        // Jadi: SEMUA sheet yang punya kolom NodeID+Date akan digabung (merge)
        // jadi satu kumpulan data, bukan cuma ambil sheet pertama yang ketemu.
        const usedSheetNames = [];
        const sheetDiagnostics = [];
        rawData = [];

        for (const sheetName of workbook.SheetNames) {
            const ws = workbook.Sheets[sheetName];
            const sample = XLSX.utils.sheet_to_json(ws, { defval: null, range: 0 });
            const cols = sample.length ? Object.keys(sample[0]) : [];
            sheetDiagnostics.push({ name: sheetName, rowCount: sample.length, columns: cols });

            const hasNodeID = cols.some(c => c.trim().toLowerCase() === 'nodeid');
            const hasDate = cols.some(c => c.trim().toLowerCase() === 'date');

            if (hasNodeID && hasDate) {
                usedSheetNames.push(sheetName);
                rawData = rawData.concat(sample);
            }
        }

        window.__lastSheetDiagnostics = sheetDiagnostics; // dipakai panel diagnostik
        window.__usedSheetNames = usedSheetNames;

        if (usedSheetNames.length === 0) {
            const sheetList = sheetDiagnostics.map(s => `"${s.name}" (kolom: ${s.columns.join(', ') || '-'})`).join(' | ');
            throw new Error(`Tidak ada sheet yang punya kolom 'NodeID' dan 'Date'. Sheet yang terbaca: ${sheetList}`);
        }

        // Buang baris "sampah" (baris kosong / tidak punya NodeID atau Date)
        rawData = rawData.filter(r => r.NodeID !== null && r.NodeID !== '' && r.Date !== null && r.Date !== '');

        if (rawData.length === 0) throw new Error("Sheet " + usedSheetNames.join(', ') + " ditemukan tapi tidak ada baris data yang valid (NodeID/Date kosong semua).");



        updateLoading("Mengelompokkan data per sumur dan tanggal...");
        await sleep(100);
        
        // Algoritma Grouping & Agregasi
        const grouped = {};
        
        for (let i = 0; i < rawData.length; i++) {
            const row = rawData[i];
            const nodeID = row.NodeID;

            // Format Tanggal dengan aman.
            // Kolom "Date" di file NUFLO adalah ANGKA SERIAL EXCEL (mis. 46205),
            // bukan teks tanggal biasa. Jadi harus dikonversi pakai rumus serial Excel.
            let dateStr = "Unknown Date";
            if (row.Date instanceof Date) {
                dateStr = row.Date.toISOString().split('T')[0];
            } else if (typeof row.Date === 'number') {
                const jsDate = excelSerialToDate(row.Date);
                dateStr = jsDate.toISOString().split('T')[0];
            } else if (typeof row.Date === 'string') {
                dateStr = row.Date.split(' ')[0];
            }

            const groupKey = `${nodeID}_${dateStr}`;
            
            if (!grouped[groupKey]) {
                grouped[groupKey] = {
                    nodeID: nodeID,
                    date: dateStr,
                    rowCount: 0,
                    sumFlowrate: 0,
                    sumInternalTemp: 0,
                    sumSupplyVolt: 0,
                    sumBatteryVolt: 0,
                    rawRows: [] // simpan baris asli untuk fitur Verifikasi Manual
                };
            }
            
            grouped[groupKey].rowCount++;
            // Parsing Float agar tidak NaN jika ada data string kosong
            grouped[groupKey].sumFlowrate += parseFloat(row.Flowrate) || 0;
            grouped[groupKey].sumInternalTemp += parseFloat(row.InternalTemperature) || 0;
            grouped[groupKey].sumSupplyVolt += parseFloat(row.SupplyVoltage) || 0;
            grouped[groupKey].sumBatteryVolt += parseFloat(row.BatteryVoltage) || 0;

            grouped[groupKey].rawRows.push({
                jam: excelTimeToHHMM(row.ScanTime),
                flowrate: row.Flowrate,
                temp: row.InternalTemperature,
                supplyV: row.SupplyVoltage,
                batteryV: row.BatteryVoltage
            });
        }

        updateLoading("Menghitung availability dan status...");
        await sleep(100);

        processedData = [];
        let totalSumAvailability = 0;
        let countStatus1 = 0;
        let countStatus0 = 0;
        const uniqueNodes = new Set();

        // Finalisasi Kalkulasi (Availability & Average)
        for (const key in grouped) {
            const g = grouped[key];
            
            // 1. Availability = (jumlah baris / 1440) * 100
            let availability = (g.rowCount / 1440) * 100;
            if (availability > 100) availability = 100; // Cap at 100% jika ada data duplikat > 1 menit
            
            // 2. Status Billing
            const status = availability >= 90 ? 1 : 0;
            
            // 3. Simple Averages
            const avgFlowrate = g.sumFlowrate / g.rowCount;
            const avgTemp = g.sumInternalTemp / g.rowCount;
            const avgSupplyV = g.sumSupplyVolt / g.rowCount;
            const avgBatteryV = g.sumBatteryVolt / g.rowCount;
            
            // Record Statisitcs
            totalSumAvailability += availability;
            if (status === 1) countStatus1++;
            else countStatus0++;
            uniqueNodes.add(g.nodeID);

            processedData.push({
                nodeID: g.nodeID,
                date: g.date,
                rowCount: g.rowCount,
                availability: availability,
                status: status,
                avgFlowrate: avgFlowrate,
                avgTemp: avgTemp,
                avgSupplyV: avgSupplyV,
                avgBatteryV: avgBatteryV,
                sumFlowrate: g.sumFlowrate,
                sumInternalTemp: g.sumInternalTemp,
                sumSupplyVolt: g.sumSupplyVolt,
                sumBatteryVolt: g.sumBatteryVolt,
                rawRows: g.rawRows
            });
        }

        updateLoading("Menyusun laporan akhir...");
        await sleep(100);

        // Update Summary Cards UI
        document.getElementById('sum-total-wells').textContent = uniqueNodes.size;
        document.getElementById('sum-avg-avail').textContent = 
            (totalSumAvailability / processedData.length || 0).toFixed(2) + '%';
        document.getElementById('sum-status-1').textContent = countStatus1;
        document.getElementById('sum-status-0').textContent = countStatus0;

        // Populate Selectors for Chart
        const selector = document.getElementById('well-selector');
        selector.innerHTML = '<option value="">Pilih Sumur untuk Grafik...</option>';
        uniqueNodes.forEach(node => {
            const opt = document.createElement('option');
            opt.value = node;
            opt.textContent = node;
            selector.appendChild(opt);
        });

        // Render Table & Selesai
        renderTable(processedData);
        renderDiagnosticPanel(sheetDiagnostics, usedSheetNames, rawData.length);

        loadingOverlay.classList.add('hidden');
        dashboardSection.classList.remove('hidden');

    } catch (error) {
        loadingOverlay.classList.add('hidden');
        uploadSection.classList.remove('hidden');
        showError("Terjadi kesalahan: " + error.message);
    }
});

// --- PANEL DIAGNOSTIK FILE ---
const REQUIRED_COLS = ['NodeID', 'Date', 'ScanTime', 'Flowrate', 'InternalTemperature', 'SupplyVoltage', 'BatteryVoltage'];

function renderDiagnosticPanel(sheetDiagnostics, usedSheetNames, validRowCount) {
    const body = document.getElementById('diagnostic-body');
    let html = '';

    html += `<p>File kamu punya <b>${sheetDiagnostics.length}</b> sheet. Sheet yang DIGABUNG untuk perhitungan: <b>${usedSheetNames.map(n => `"${n}"`).join(', ')}</b> (semua sheet yang punya kolom NodeID & Date otomatis digabung jadi satu).</p>`;

    sheetDiagnostics.forEach(s => {
        const isUsed = usedSheetNames.includes(s.name);
        html += `<h4 style="margin-top:1rem;">${isUsed ? '✅ Dipakai' : '⏭️ Dilewati'} — Sheet "${s.name}" (${s.rowCount} baris terbaca)</h4>`;
        html += `<table><thead><tr><th>Kolom yang dibutuhkan</th><th>Status di sheet ini</th></tr></thead><tbody>`;
        REQUIRED_COLS.forEach(col => {
            const found = s.columns.some(c => c.trim().toLowerCase() === col.toLowerCase());
            html += `<tr><td><code>${col}</code></td><td class="${found ? 'tag-ok' : 'tag-miss'}">${found ? 'Ditemukan' : 'Tidak ada'}</td></tr>`;
        });
        html += `</tbody></table>`;
        html += `<p class="verify-note">Semua kolom terdeteksi di sheet ini: ${s.columns.join(', ') || '(tidak ada)'}</p>`;
    });

    html += `<p style="margin-top:1rem;"><b>Jumlah baris valid dipakai untuk perhitungan:</b> ${validRowCount} baris (baris dengan NodeID/Date kosong sudah dibuang otomatis).</p>`;

    body.innerHTML = html;
}

document.getElementById('diagnostic-toggle').addEventListener('click', () => {
    const body = document.getElementById('diagnostic-body');
    const arrow = document.getElementById('diagnostic-arrow');
    body.classList.toggle('hidden');
    arrow.textContent = body.classList.contains('hidden') ? '▼' : '▲';
});

// --- MODAL VERIFIKASI MANUAL ---
function openVerifyModal(row) {
    document.getElementById('verify-title').textContent = `Verifikasi Manual — ${row.nodeID} / ${row.date}`;

    const previewRows = row.rawRows.slice(0, 15);
    const tableRowsHtml = previewRows.map((r, idx) => `
        <tr>
            <td>${idx + 1}</td>
            <td>${r.jam}</td>
            <td>${r.flowrate ?? '-'}</td>
            <td>${r.temp ?? '-'}</td>
            <td>${r.supplyV ?? '-'}</td>
            <td>${r.batteryV ?? '-'}</td>
        </tr>
    `).join('');

    const html = `
        <div class="verify-formula">
            <b>1. Availability</b>
            <code>Availability = (Jumlah baris / 1440) × 100
= (${row.rowCount} / 1440) × 100
= ${row.availability.toFixed(4)}%</code>
        </div>
        <div class="verify-formula">
            <b>2. Status Billing</b>
            <code>Status = 1 jika Availability >= 90%, selain itu 0
Availability = ${row.availability.toFixed(2)}% → Status = ${row.status} (${row.status === 1 ? 'Bisa Ditagih' : 'Belum Bisa Ditagih'})</code>
        </div>
        <div class="verify-formula">
            <b>3. Rata-rata (Average)</b>
            <code>Avg Flowrate   = ${row.sumFlowrate.toFixed(2)} / ${row.rowCount} = ${row.avgFlowrate.toFixed(4)}
Avg Temp       = ${row.sumInternalTemp.toFixed(2)} / ${row.rowCount} = ${row.avgTemp.toFixed(4)}
Avg SupplyV    = ${row.sumSupplyVolt.toFixed(2)} / ${row.rowCount} = ${row.avgSupplyV.toFixed(4)}
Avg BatteryV   = ${row.sumBatteryVolt.toFixed(2)} / ${row.rowCount} = ${row.avgBatteryV.toFixed(4)}</code>
        </div>
        <p><b>Contoh baris mentah yang dipakai</b> (menampilkan ${previewRows.length} dari total ${row.rowCount} baris):</p>
        <div class="verify-table-wrap">
            <table>
                <thead><tr><th>#</th><th>Jam</th><th>Flowrate</th><th>Temp</th><th>SupplyV</th><th>BatteryV</th></tr></thead>
                <tbody>${tableRowsHtml}</tbody>
            </table>
        </div>
        <p class="verify-note">Silakan cocokkan dengan hitungan manual kakak senior menggunakan Excel: filter NodeID = "${row.nodeID}" dan Date = "${row.date}" di file asli, lalu COUNT baris dan AVERAGE tiap kolom. Hasilnya harus sama persis dengan angka di atas.</p>
    `;

    document.getElementById('verify-body').innerHTML = html;
    document.getElementById('verify-modal').classList.remove('hidden');
}

document.getElementById('verify-close').addEventListener('click', () => {
    document.getElementById('verify-modal').classList.add('hidden');
});
document.getElementById('verify-modal').addEventListener('click', (e) => {
    if (e.target.id === 'verify-modal') {
        document.getElementById('verify-modal').classList.add('hidden');
    }
});

// --- RENDER TABLE & FILTER ---
const tableBody = document.getElementById('table-body');
const searchInput = document.getElementById('search-input');
const filterStatus = document.getElementById('filter-status');

function renderTable(data) {
    tableBody.innerHTML = '';
    data.forEach((row, idx) => {
        const tr = document.createElement('tr');
        // Warna highlight baris berdasarkan Status
        tr.className = row.status === 1 ? 'status-green' : 'status-red';
        
        tr.innerHTML = `
            <td>${row.nodeID}</td>
            <td>${row.date}</td>
            <td>${row.rowCount}</td>
            <td>${row.availability.toFixed(2)}%</td>
            <td>${row.status}</td>
            <td>${row.avgFlowrate.toFixed(2)}</td>
            <td>${row.avgTemp.toFixed(2)}</td>
            <td>${row.avgSupplyV.toFixed(2)}</td>
            <td>${row.avgBatteryV.toFixed(2)}</td>
            <td><button class="btn-verify" data-idx="${idx}">Cek Manual</button></td>
        `;
        tableBody.appendChild(tr);
    });

    tableBody.querySelectorAll('.btn-verify').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.target.getAttribute('data-idx'));
            openVerifyModal(data[idx]);
        });
    });
}

function filterAndSearch() {
    const searchTerm = searchInput.value.toLowerCase();
    const statusVal = filterStatus.value;
    
    const filtered = processedData.filter(row => {
        const matchSearch = row.nodeID.toLowerCase().includes(searchTerm) || row.date.toLowerCase().includes(searchTerm);
        const matchStatus = statusVal === 'all' ? true : (row.status == statusVal);
        return matchSearch && matchStatus;
    });
    
    renderTable(filtered);
}

searchInput.addEventListener('input', filterAndSearch);
filterStatus.addEventListener('change', filterAndSearch);

// --- SORTING TABLE ---
document.querySelectorAll('th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
        const key = th.getAttribute('data-sort');
        const sortMap = {
            'node': 'nodeID', 'date': 'date', 'count': 'rowCount', 
            'avail': 'availability', 'status': 'status'
        };
        const actualKey = sortMap[key];

        processedData.sort((a, b) => {
            let valA = a[actualKey];
            let valB = b[actualKey];
            if(typeof valA === 'string') valA = valA.toLowerCase();
            if(typeof valB === 'string') valB = valB.toLowerCase();
            
            if (valA < valB) return -1 * sortDirection;
            if (valA > valB) return 1 * sortDirection;
            return 0;
        });
        
        sortDirection *= -1; // Toggle arah sorting
        filterAndSearch(); // Render ulang dengan mempertahankan filter
    });
});

// --- CHART RENDERING ---
document.getElementById('well-selector').addEventListener('change', (e) => {
    const selectedNode = e.target.value;
    if (!selectedNode) return;

    // Ambil data khusus untuk node ini dan sort berdasarkan tanggal
    const nodeData = processedData
        .filter(d => d.nodeID === selectedNode)
        .sort((a, b) => new Date(a.date) - new Date(b.date));

    const labels = nodeData.map(d => d.date);
    const dataPoints = nodeData.map(d => d.availability);

    const ctx = document.getElementById('trendChart').getContext('2d');
    
    if (chartInstance) chartInstance.destroy();
    
    chartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: `Availability (%) - ${selectedNode}`,
                data: dataPoints,
                borderColor: '#0d47a1',
                backgroundColor: 'rgba(13, 71, 161, 0.1)',
                borderWidth: 2,
                fill: true,
                tension: 0.1
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: { beginAtZero: true, max: 100 }
            }
        }
    });
});

// --- EXPORT FEATURES ---
document.getElementById('btn-export-excel').addEventListener('click', () => {
    const ws = XLSX.utils.json_to_sheet(processedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Laporan_Availability");
    XLSX.writeFile(wb, "Laporan_Sumur_NUFLO.xlsx");
});

document.getElementById('btn-export-pdf').addEventListener('click', () => {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    doc.text("Laporan Availability Sumur NUFLO MCIII", 14, 15);
    
    const tableColumn = ["Nama Sumur", "Tanggal", "Jml Baris", "Availability", "Status", "Avg Flowrate"];
    const tableRows = [];

    processedData.forEach(row => {
        tableRows.push([
            row.nodeID, row.date, row.rowCount, 
            row.availability.toFixed(2) + '%', row.status, 
            row.avgFlowrate.toFixed(2)
        ]);
    });

    doc.autoTable({
        head: [tableColumn],
        body: tableRows,
        startY: 20,
        styles: { fontSize: 8 },
        headStyles: { fillColor: [13, 71, 161] }
    });
    
    doc.save("Laporan_Sumur_NUFLO.pdf");
});

// --- RESET UTILITY ---
document.getElementById('btn-reset').addEventListener('click', () => {
    dashboardSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    fileInput.value = '';
    selectedFile = null;
    fileInfo.classList.add('hidden');
});