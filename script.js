// Data global
let uploadedFiles = [];          // { file, bank, password, transactions, error }
let pendingPasswordResolve = null;
let currentBankForPassword = null;

// Elemen DOM
const fileInput = document.getElementById('files');
const fileListDiv = document.getElementById('fileList');
const calculateBtn = document.getElementById('calculateBtn');
const passwordModal = document.getElementById('passwordModal');
const bankNameSpan = document.getElementById('bankName');
const fileNamesSpan = document.getElementById('fileNames');
const bankPasswordInput = document.getElementById('bankPassword');
const submitPasswordBtn = document.getElementById('submitPassword');
const cancelPasswordBtn = document.getElementById('cancelPassword');
const resultDiv = document.getElementById('result');
const errorDiv = document.getElementById('error');
const minBalanceSpan = document.getElementById('min_balance');
const nisabSpan = document.getElementById('nisab');
const statusSpan = document.getElementById('status');
const zakatSpan = document.getElementById('zakat');

// ====================== UTILS ======================
function parseDate(dateStr) {
    // Format: "04 Feb 2026"
    const parts = dateStr.trim().split(' ');
    if (parts.length === 3) {
        const day = parts[0];
        const month = parts[1];
        const year = parts[2];
        const monthMap = {
            'Jan':0, 'Feb':1, 'Mar':2, 'Apr':3, 'Mei':4, 'Jun':5,
            'Jul':6, 'Agu':7, 'Sep':8, 'Okt':9, 'Nov':10, 'Des':11
        };
        if (monthMap.hasOwnProperty(month)) {
            return new Date(year, monthMap[month], day);
        }
    }
    return null;
}

function formatRupiah(angka) {
    return new Intl.NumberFormat('id-ID').format(angka);
}

// ====================== DETEKSI BANK ======================
function detectBankFromExcel(data) {
    // data adalah array of rows (array of nilai)
    for (let row of data) {
        if (Array.isArray(row)) {
            const rowText = row.join(' ').toLowerCase();
            if (rowText.includes('bank mandiri')) return 'mandiri';
            if (rowText.includes('bca')) return 'bca';
            if (rowText.includes('bni')) return 'bni';
            if (rowText.includes('bri')) return 'bri';
        }
    }
    return 'unknown';
}

// ====================== PARSER PER BANK ======================
function parseMandiriExcel(data) {
    // Cari baris header yang mengandung "Tanggal" dan "Saldo"
    let headerRowIndex = -1;
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (Array.isArray(row) && 
            row.some(cell => cell && cell.toString().includes('Tanggal')) && 
            row.some(cell => cell && cell.toString().includes('Saldo'))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error('Header tidak ditemukan');

    const header = data[headerRowIndex];
    const dateCol = header.findIndex(cell => cell && cell.toString().includes('Tanggal'));
    const balanceCol = header.findIndex(cell => cell && cell.toString().includes('Saldo'));
    if (dateCol === -1 || balanceCol === -1) throw new Error('Kolom tanggal atau saldo tidak ditemukan');

    const transactions = [];
    for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        const dateStr = row[dateCol];
        const balanceStr = row[balanceCol];
        if (dateStr && balanceStr) {
            const date = parseDate(dateStr);
            if (date) {
                // Bersihkan format angka: "16.593.021,00" -> 16593021.00
                const clean = balanceStr.replace(/\./g, '').replace(',', '.');
                const balance = parseFloat(clean);
                if (!isNaN(balance)) {
                    transactions.push({ date, balance });
                }
            }
        }
    }
    return transactions;
}

// ====================== BACA FILE EXCEL ======================
async function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
                resolve(rows);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// ====================== PROSES SEMUA FILE ======================
async function processAllFiles() {
    // Reset status
    for (let f of uploadedFiles) {
        f.transactions = null;
        f.error = null;
    }

    // Kelompokkan file berdasarkan bank untuk keperluan password (PDF nanti)
    // Sementara kita proses Excel dulu
    for (let f of uploadedFiles) {
        if (f.file.type === 'application/pdf') {
            // TODO: handle PDF dengan password
            f.error = 'PDF belum didukung di versi ini. Gunakan Excel/CSV.';
            continue;
        }

        try {
            const rows = await readExcel(f.file);
            const bank = detectBankFromExcel(rows);
            f.bank = bank;

            if (bank === 'mandiri') {
                f.transactions = parseMandiriExcel(rows);
            } else {
                f.error = `Bank ${bank} belum memiliki parser.`;
            }
        } catch (err) {
            f.error = err.message;
        }
    }
    renderFileList();
}

// ====================== RENDER DAFTAR FILE ======================
function renderFileList() {
    fileListDiv.innerHTML = '';
    uploadedFiles.forEach((f, index) => {
        const div = document.createElement('div');
        div.className = 'file-item';
        div.innerHTML = `
            <span class="bank">${f.bank || 'Deteksi...'} - ${f.file.name}</span>
            <span class="status ${f.error ? 'error' : ''}">
                ${f.error ? `❌ ${f.error}` : (f.transactions ? `✅ ${f.transactions.length} transaksi` : '⏳')}
            </span>
        `;
        fileListDiv.appendChild(div);
    });
}

// ====================== HITUNG ZAKAT ======================
function calculateZakat() {
    const startDate = new Date(document.getElementById('start_date').value);
    const endDate = new Date(document.getElementById('end_date').value);
    const goldPrice = parseFloat(document.getElementById('gold_price').value);

    if (!startDate || !endDate || !goldPrice) {
        showError('Tanggal dan harga emas harus diisi.');
        return;
    }
    if (startDate > endDate) {
        showError('Tanggal mulai harus sebelum tanggal akhir.');
        return;
    }

    // Kumpulkan semua transaksi yang berhasil
    const allTransactions = uploadedFiles.filter(f => f.transactions).map(f => f.transactions);
    if (allTransactions.length === 0) {
        showError('Tidak ada data transaksi valid.');
        return;
    }

    // Buat daily balance untuk setiap akun
    const dailyMaps = [];
    for (let tx of allTransactions) {
        const map = generateDailyBalance(tx, startDate, endDate);
        dailyMaps.push(map);
    }

    // Jumlahkan semua akun per hari
    const totalMap = sumDailyBalances(dailyMaps);

    // Cari saldo terendah
    let minBalance = Infinity;
    for (let bal of totalMap.values()) {
        if (bal < minBalance) minBalance = bal;
    }

    const nisab = 85 * goldPrice;
    const wajib = minBalance >= nisab;
    const zakat = wajib ? minBalance * 0.025 : 0;

    // Tampilkan hasil
    minBalanceSpan.innerText = formatRupiah(minBalance);
    nisabSpan.innerText = formatRupiah(nisab);
    statusSpan.innerText = wajib ? 'Wajib Zakat' : 'Tidak Wajib Zakat';
    zakatSpan.innerText = formatRupiah(zakat);
    resultDiv.style.display = 'block';
    errorDiv.style.display = 'none';
}

function generateDailyBalance(transactions, start, end) {
    // Buat Map tanggal -> saldo
    const daily = new Map();
    let current = new Date(start);
    while (current <= end) {
        daily.set(current.toISOString().split('T')[0], null);
        current.setDate(current.getDate() + 1);
    }

    // Urutkan transaksi
    transactions.sort((a,b) => a.date - b.date);

    // Isi saldo dari transaksi
    let lastBalance = null;
    for (let t of transactions) {
        const dateStr = t.date.toISOString().split('T')[0];
        if (daily.has(dateStr)) {
            daily.set(dateStr, t.balance);
            lastBalance = t.balance;
        }
    }

    // Forward fill
    let prev = null;
    for (let [date, bal] of daily) {
        if (bal !== null) {
            prev = bal;
        } else {
            daily.set(date, prev);
        }
    }
    return daily;
}

function sumDailyBalances(maps) {
    const firstMap = maps[0];
    const total = new Map();
    for (let [date, _] of firstMap) {
        let sum = 0;
        for (let map of maps) {
            sum += map.get(date) || 0;
        }
        total.set(date, sum);
    }
    return total;
}

function showError(msg) {
    errorDiv.innerText = msg;
    errorDiv.style.display = 'block';
    resultDiv.style.display = 'none';
}

// ====================== EVENT LISTENERS ======================
fileInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    uploadedFiles = files.map(file => ({ file, bank: null, password: null, transactions: null, error: null }));
    renderFileList();
    // Proses otomatis untuk mendeteksi bank (tanpa password dulu)
    processAllFiles();
});

calculateBtn.addEventListener('click', () => {
    if (uploadedFiles.length === 0) {
        showError('Unggah file terlebih dahulu.');
        return;
    }
    // Pastikan semua file sudah diproses
    const unfinished = uploadedFiles.some(f => f.transactions === null && !f.error);
    if (unfinished) {
        showError('Tunggu hingga semua file selesai diproses.');
        return;
    }
    calculateZakat();
});

// Password modal (sementara hanya untuk placeholder, PDF belum diimplementasi)
submitPasswordBtn.addEventListener('click', () => {
    if (pendingPasswordResolve) {
        pendingPasswordResolve(bankPasswordInput.value);
    }
    passwordModal.style.display = 'none';
    bankPasswordInput.value = '';
});

cancelPasswordBtn.addEventListener('click', () => {
    if (pendingPasswordResolve) {
        pendingPasswordResolve(null); // batal
    }
    passwordModal.style.display = 'none';
    bankPasswordInput.value = '';
});

// Tutup modal jika klik di luar
window.addEventListener('click', (e) => {
    if (e.target === passwordModal) {
        cancelPasswordBtn.click();
    }
});