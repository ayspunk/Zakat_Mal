// ====================== KONFIGURASI DINAMIS ======================
const CURRENT_GOLD_PRICE = 2780420; // 09 Maret 2026
const RECOMMENDED_START_DATE = '2025-03-01'; // 1 Ramadhan 1446 H
const RECOMMENDED_END_DATE = '2026-02-18';   // 1 Ramadhan 1447 H
const RECOMMENDED_START_TEXT = '1 Maret 2025';
const RECOMMENDED_END_TEXT = '18 Februari 2026';

// ====================== Data Global ======================
let uploadedFiles = [];
let pendingPasswordResolve = null;
let currentFileForPassword = null;
let bankPasswords = {};
let balanceChartInstance = null;
let lastTotalMap = null;
let lastBankMaps = []; // [{label, map}] per rekening untuk CSV multi-kolom

// Elemen DOM
let fileInput, fileListDiv, calculateBtn, passwordModal, bankNameSpan, fileNamesSpan,
    bankPasswordInput, submitPasswordBtn, cancelPasswordBtn, resultDiv, errorDiv,
    minBalanceSpan, nisabSpan, statusSpan, zakatSpan, periodeStartSpan, periodeEndSpan;

// ====================== UTILS ======================
function toLocalDateStr(date) {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function parseDate(dateStr) {
    if (dateStr == null || typeof dateStr !== 'string') return null;
    const parts = dateStr.trim().split(' ');
    if (parts.length === 3) {
        const day = parts[0];
        const month = parts[1].charAt(0).toUpperCase() + parts[1].slice(1).toLowerCase();
        const year = parts[2];
        const monthMap = {
            'Jan':0, 'Feb':1, 'Mar':2, 'Apr':3, 'Mei':4, 'Jun':5,
            'Jul':6, 'Agu':7, 'Sep':8, 'Okt':9, 'Nov':10, 'Des':11,
            'May':4, 'Aug':7, 'Oct':9, 'Dec':11
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

// ====================== PARSER ANGKA CERDAS ======================
function parseNumber(str) {
    if (str == null) return NaN;
    str = str.toString().trim();
    if (str === '') return NaN;

    const lastComma = str.lastIndexOf(',');
    const lastDot = str.lastIndexOf('.');

    if (lastComma > -1 && lastDot > -1) {
        if (lastDot > lastComma) {
            str = str.replace(/,/g, '');
            return parseFloat(str);
        } else {
            str = str.replace(/\./g, '');
            str = str.replace(',', '.');
            return parseFloat(str);
        }
    } else if (lastComma > -1) {
        if (str.match(/,\d{2}$/)) {
            str = str.replace(',', '.');
            return parseFloat(str);
        } else {
            str = str.replace(/,/g, '');
            return parseFloat(str);
        }
    } else if (lastDot > -1) {
        if (str.match(/\.\d{2}$/)) {
            return parseFloat(str);
        } else {
            str = str.replace(/\./g, '');
            return parseFloat(str);
        }
    } else {
        return parseFloat(str);
    }
}

// ====================== DETEKSI BANK ======================
function guessBankFromFilename(filename) {
    const lower = filename.toLowerCase();
    if (lower.includes('mandiri')) return 'mandiri';
    if (lower.includes('bca')) return 'bca';
    if (lower.includes('bni')) return 'bni';
    if (lower.includes('bri')) return 'bri';
    if (lower.includes('danamon') || lower.includes('bdi')) return 'danamon';
    // Pola nomor nasabah Danamon (10 digit, tanpa nama bank di filename)
    if (lower.match(/\d{10}/)) return 'danamon_candidate';
    return 'unknown';
}

function detectBankFromExcel(rows) {
    for (let row of rows) {
        if (Array.isArray(row)) {
            const rowText = row.join(' ').toLowerCase();
            if (rowText.includes('bank mandiri') || rowText.includes('mandiri')) return 'mandiri';
            if (rowText.includes('bca')) return 'bca';
            if (rowText.includes('bni')) return 'bni';
            if (rowText.includes('bri')) return 'bri';
            if (rowText.includes('danamon')) return 'danamon';
        }
    }
    return 'unknown';
}

function detectBankFromPDFLines(lines) {
    // Cari di semua baris (logo/nama bank bisa muncul di posisi mana saja)
    const fullText = lines.join(' ').toLowerCase();
    if (fullText.includes('danamon')) return 'danamon';
    if (fullText.includes('mandiri')) return 'mandiri';
    if (fullText.includes('bca')) return 'bca';
    if (fullText.includes('bni')) return 'bni';
    if (fullText.includes('bri')) return 'bri';
    return 'unknown';
}

// ====================== BACA PDF DENGAN PDF.JS ======================
async function readPDF(file) {
    if (typeof pdfjsLib === 'undefined') throw new Error('PDF.js belum dimuat. Coba refresh halaman.');
    const buffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
    const allLines = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const textContent = await page.getTextContent();

        // Kelompokkan item teks berdasarkan koordinat Y (toleransi ±3pt)
        const yGroups = [];
        for (const item of textContent.items) {
            if (!item.str || !item.str.trim()) continue;
            const y = item.transform[5];
            const x = item.transform[4];
            let group = yGroups.find(g => Math.abs(g.y - y) <= 3);
            if (!group) { group = { y, items: [] }; yGroups.push(group); }
            group.items.push({ x, text: item.str });
        }

        // Urutkan baris atas→bawah, item kiri→kanan, lalu gabungkan
        yGroups.sort((a, b) => b.y - a.y);
        for (const group of yGroups) {
            const lineText = group.items
                .sort((a, b) => a.x - b.x)
                .map(i => i.text.trim())
                .filter(t => t)
                .join(' ')
                .trim();
            if (lineText) allLines.push(lineText);
        }
    }
    return allLines;
}

// ====================== PARSER DANAMON PDF (IDR ONLY) ======================
function parseDanamonPDF(lines) {
    const BULAN_ID = {
        'Januari':1,'Februari':2,'Maret':3,'April':4,'Mei':5,'Juni':6,
        'Juli':7,'Agustus':8,'September':9,'Oktober':10,'November':11,'Desember':12
    };

    // Cari periode dari header statement
    let periodYear = null, periodMonth = null;
    for (const line of lines) {
        const m = line.match(/Periode[:\s]+\d+\s*[-–]\s*\d+\s+(\w+)\s+(\d{4})/i);
        if (m) {
            periodMonth = BULAN_ID[m[1]] || null;
            periodYear = parseInt(m[2]);
            if (periodMonth && periodYear) break;
        }
    }
    if (!periodYear || !periodMonth) throw new Error('Periode tidak ditemukan dalam PDF Danamon.');

    // Regex untuk angka format Indonesia: 1.234.567,89
    const RE_IDR_NUM = /\d+(?:\.\d{3})*,\d{2}/g;

    // ── STRATEGI PARSING ──────────────────────────────────────────────────────
    // PDF.js menggabungkan teks dari dua kolom ke satu baris berdasarkan Y-coord.
    // Akibatnya baris "SALDO BULAN LALU 5.273.019,72" bisa tergabung dengan
    // angka dari tabel ringkasan di kolom sebelah (misal "192.466.019,37").
    // Solusi: gunakan dua penanda yang BERBEDA antara tabel ringkasan vs RINCIAN:
    //   • Tabel ringkasan  → "DANAMON LEBIH PRO (IDR) 903... IDR  5.263.377,65"  (tanpa dash)
    //   • Header RINCIAN   → "DANAMON LEBIH PRO (IDR) - IDR - 903..."            (ada dash)
    // Hanya format DENGAN DASH yang menandai masuk ke section transaksi.
    // ─────────────────────────────────────────────────────────────────────────

    let inRincian = false, inIDRSection = false;
    let saldoAwal = null;
    const transactions = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // Masuk ke bagian rincian transaksi per rekening
        if (/RINCIAN TRANSAKSI PER REKENING/i.test(line)) {
            inRincian = true;
            continue;
        }
        if (!inRincian) continue;

        // Header RINCIAN IDR → format dengan dash: "DANAMON LEBIH PRO (IDR) - IDR - 9036..."
        // BUKAN baris ringkasan tanpa dash
        if (/DANAMON LEBIH PRO \(IDR\)\s*-\s*IDR\s*-/i.test(line)) {
            inIDRSection = true;
            continue;
        }

        // Hentikan saat masuk rekening berikutnya (format RINCIAN juga pakai dash)
        if (inIDRSection && (
            /PRIMADOLLAR[^(]*-\s*USD\s*-/i.test(line) ||
            /DANAMON LEBIH PRO \((?!IDR)/i.test(line) ||
            /^TOTAL\s+[\d.,]+/.test(line.trim()) ||   // baris TOTAL penutup tabel
            /D-POINT BANKING|AKHIR LAPORAN/i.test(line)
        )) {
            break;
        }
        if (!inIDRSection) continue;

        // Baris "SALDO BULAN LALU" → ambil angka PERTAMA yang ditemukan (bukan terakhir)
        // agar tidak terpengaruh angka dari kolom ringkasan yang tergabung oleh PDF.js
        if (/SALDO BULAN LALU/i.test(line)) {
            const nums = [...line.matchAll(RE_IDR_NUM)].map(m => m[0]);
            if (nums.length) {
                // Ambil angka terkecil yang masuk akal (> 0) sebagai saldo rekening IDR,
                // bukan angka total gabungan yang bisa saja tergabung di baris yang sama
                const parsed = nums.map(n => parseNumber(n)).filter(n => !isNaN(n) && n > 0);
                saldoAwal = parsed.length ? Math.min(...parsed) : null;
                console.log(`Danamon saldo awal: ${saldoAwal} (dari angka: ${nums.join(', ')})`);
            }
            continue;
        }

        // Baris transaksi: dimulai dengan DD/MM DD/MM
        const txMatch = line.match(/^(\d{2})\/(\d{2})\s+(\d{2})\/(\d{2})\s/);
        if (txMatch) {
            // Gunakan TGL VALUTA (kolom ke-2) sebagai tanggal efektif
            const day   = parseInt(txMatch[3]);
            const month = parseInt(txMatch[4]);
            let year = periodYear;
            // Tangani batas tahun: valuta Desember di statement Januari
            if (month === 12 && periodMonth === 1) year--;
            const date = new Date(year, month - 1, day);

            const nums = [...line.matchAll(RE_IDR_NUM)].map(m => m[0]);
            if (nums.length) {
                // Saldo selalu di kolom terakhir; tapi jika ikut tergabung dengan
                // kolom ringkasan, ambil yang terkecil (saldo IDR, bukan total gabungan)
                const parsed = nums.map(n => parseNumber(n)).filter(n => !isNaN(n) && n >= 0);
                if (parsed.length) {
                    // Heuristik: jika ada lebih dari 3 angka (kemungkinan kolom tergabung),
                    // ambil angka yang masuk akal sbg saldo IDR (< 50 jt untuk rekening ini)
                    let balance;
                    if (parsed.length <= 3) {
                        balance = parsed[parsed.length - 1]; // kolom SALDO = terakhir
                    } else {
                        // Ambil angka terakhir yang < 50 jt (threshold aman untuk IDR account)
                        const IDR_THRESHOLD = 50_000_000;
                        const candidates = parsed.filter(n => n < IDR_THRESHOLD);
                        balance = candidates.length
                            ? candidates[candidates.length - 1]
                            : parsed[parsed.length - 1];
                    }
                    transactions.push({ date, balance });
                    console.log(`Danamon tx: ${toLocalDateStr(date)} saldo=${balance}`);
                }
            }
        }
    }

    if (saldoAwal === null && transactions.length === 0)
        throw new Error('Data transaksi IDR tidak ditemukan dalam PDF Danamon. Pastikan format PDF benar.');

    // Tambahkan saldo awal sebagai titik awal periode bulan ini
    if (saldoAwal !== null) {
        const openDate = new Date(periodYear, periodMonth - 1, 1);
        transactions.unshift({ date: openDate, balance: saldoAwal });
    }

    transactions.sort((a, b) => a.date - b.date);
    console.log(`Danamon parser: ${transactions.length} titik saldo untuk periode ${periodMonth}/${periodYear}`);

    // Ambil nomor rekening IDR dari header section RINCIAN
    let accountId = null;
    for (const line of lines) {
        const m = line.match(/DANAMON LEBIH PRO \(IDR\)\s*-\s*IDR\s*-\s*(\d{6,})/i);
        if (m) { accountId = m[1]; break; }
    }
    console.log('Danamon accountId:', accountId);
    return { transactions, accountId };
}



function parseMandiriExcel(rows) {
    let periodeAwal = null;
    let periodeAkhir = null;
    let saldoAwal = null;

    const fullText = rows.map(row => Array.isArray(row) ? row.join(' ') : '').join(' ');

    // Cari Periode
    const BULAN_PATTERN = '(?:Jan|Feb|Mar|Apr|Mei|May|Jun|Jul|Agu|Aug|Sep|Okt|Oct|Nov|Des|Dec)';
    const periodeRegex = new RegExp(
        '(\\d{1,2}\\s+' + BULAN_PATTERN + '\\s+\\d{4})\\s*[-–]\\s*(\\d{1,2}\\s+' + BULAN_PATTERN + '\\s+\\d{4})',
        'i'
    );
    const periodeMatch = fullText.match(periodeRegex);
    if (periodeMatch) {
        periodeAwal = parseDate(periodeMatch[1]);
        periodeAkhir = parseDate(periodeMatch[2]);
    } else {
        for (let row of rows) {
            if (!Array.isArray(row)) continue;
            if (!row.some(cell => cell && cell.toString().toLowerCase().includes('periode'))) continue;
            const rowText = row.map(c => c != null ? c.toString() : '').join(' ');
            const match = rowText.match(periodeRegex);
            if (match) {
                periodeAwal = parseDate(match[1]);
                periodeAkhir = parseDate(match[2]);
                break;
            }
        }
    }

    function getNumberFromCell(cell) {
        if (cell == null) return NaN;
        return parseNumber(cell.toString());
    }

    // Cari Saldo Awal (multimetode)
    console.log('Mencari saldo awal...');

    // Metode 1: Baris dengan "Saldo Awal" atau "Initial Balance"
    for (let row of rows) {
        if (!Array.isArray(row)) continue;
        const rowText = row.join(' ').toLowerCase();
        if (rowText.includes('saldo awal') || rowText.includes('initial balance')) {
            for (let cell of row) {
                const num = getNumberFromCell(cell);
                if (!isNaN(num) && num > 1000) {
                    saldoAwal = num;
                    console.log('Metode 1 berhasil:', saldoAwal);
                    break;
                }
            }
            if (saldoAwal) break;
        }
    }

    // Metode 2: Regex di seluruh teks
    if (!saldoAwal) {
        const match = fullText.match(/(?:saldo\s+awal|initial\s+balance)[^\d]*([\d\.\,]+)/i);
        if (match) {
            saldoAwal = parseNumber(match[1]);
            console.log('Metode 2 berhasil:', saldoAwal);
        }
    }

    // Metode 3: Baris yang mengandung kata produk (Tabungan, Giro, dll) dan ambil angka di kolom terakhir
    if (!saldoAwal) {
        for (let row of rows) {
            if (!Array.isArray(row)) continue;
            const rowText = row.join(' ').toLowerCase();
            if (rowText.includes('tabungan') || rowText.includes('giro') || rowText.includes('rekening') || rowText.includes('deposito')) {
                for (let i = row.length - 1; i >= 0; i--) {
                    const num = getNumberFromCell(row[i]);
                    if (!isNaN(num) && num > 1000) {
                        saldoAwal = num;
                        console.log('Metode 3 berhasil:', saldoAwal);
                        break;
                    }
                }
                if (saldoAwal) break;
            }
        }
    }

    // Metode 4: Cari angka besar di baris sebelum header tabel
    if (!saldoAwal) {
        let headerIdx = -1;
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (Array.isArray(row) && row.some(cell => cell && (cell.toString().includes('Tanggal') || cell.toString().includes('Date')))) {
                headerIdx = i;
                break;
            }
        }
        if (headerIdx > 0) {
            for (let i = 0; i < headerIdx; i++) {
                const row = rows[i];
                if (!Array.isArray(row)) continue;
                for (let cell of row) {
                    const num = getNumberFromCell(cell);
                    if (!isNaN(num) && num > 10000) {
                        saldoAwal = num;
                        console.log('Metode 4 berhasil:', saldoAwal);
                        break;
                    }
                }
                if (saldoAwal) break;
            }
        }
    }

    // Metode 5: Cari baris yang mengandung kata "Saldo" dan "Awal" secara terpisah, atau cek baris berikutnya
    if (!saldoAwal) {
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!Array.isArray(row)) continue;
            const rowText = row.join(' ').toLowerCase();
            if (rowText.includes('saldo') && (rowText.includes('awal') || rowText.includes('initial'))) {
                // Coba baris ini
                for (let cell of row) {
                    const num = getNumberFromCell(cell);
                    if (!isNaN(num) && num > 1000) {
                        saldoAwal = num;
                        console.log('Metode 5 berhasil:', saldoAwal);
                        break;
                    }
                }
                if (saldoAwal) break;
                // Coba baris berikutnya
                if (i+1 < rows.length) {
                    const nextRow = rows[i+1];
                    if (Array.isArray(nextRow)) {
                        for (let cell of nextRow) {
                            const num = getNumberFromCell(cell);
                            if (!isNaN(num) && num > 1000) {
                                saldoAwal = num;
                                console.log('Metode 5 (baris berikut) berhasil:', saldoAwal);
                                break;
                            }
                        }
                    }
                }
                if (saldoAwal) break;
            }
        }
    }

    if (!periodeAwal || !periodeAkhir || !saldoAwal) {
        console.error('Debug info (Mandiri parser):', { periodeAwal, periodeAkhir, saldoAwal });
        const detectedBank = detectBankFromExcel(rows);
        if (detectedBank !== 'unknown') {
            throw new Error(`File terdeteksi sebagai bank ${detectedBank.toUpperCase()}, namun saldo awal tidak ditemukan.`);
        } else {
            throw new Error('Tidak dapat menemukan periode atau saldo awal dalam file. Pastikan file e-statement Mandiri valid.');
        }
    }

    const transactions = [{ date: periodeAwal, balance: saldoAwal }];

    // Cari header tabel transaksi
    let headerRowIndex = -1, dateCol = -1, balanceCol = -1;
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!Array.isArray(row)) continue;
        const hasDate = row.some(cell => cell && (cell.toString().includes('Tanggal') || cell.toString().includes('Date')));
        const hasBalance = row.some(cell => cell && (cell.toString().includes('Saldo') || cell.toString().includes('Balance')));
        if (hasDate && hasBalance) {
            headerRowIndex = i;
            dateCol = row.findIndex(cell => cell && (cell.toString().includes('Tanggal') || cell.toString().includes('Date')));
            balanceCol = row.findIndex(cell => cell && (cell.toString().includes('Saldo') || cell.toString().includes('Balance')));
            break;
        }
    }
    if (headerRowIndex === -1) {
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!Array.isArray(row)) continue;
            if (row.some(cell => cell && (cell.toString().includes('Tanggal') || cell.toString().includes('Date')))) {
                headerRowIndex = i;
                dateCol = row.findIndex(cell => cell && (cell.toString().includes('Tanggal') || cell.toString().includes('Date')));
                const nextRow = rows[i+1];
                if (nextRow && Array.isArray(nextRow)) {
                    for (let j = nextRow.length-1; j>=0; j--) {
                        if (nextRow[j] && nextRow[j].toString().match(/[\d\.\,]+/)) {
                            balanceCol = j;
                            break;
                        }
                    }
                }
                if (balanceCol === -1) balanceCol = row.length-1;
                break;
            }
        }
    }
    if (headerRowIndex === -1 || dateCol === -1 || balanceCol === -1) {
        throw new Error('Header tabel transaksi tidak ditemukan.');
    }

    console.log(`Header ditemukan di baris ${headerRowIndex}, kolom tanggal ${dateCol}, kolom saldo ${balanceCol}`);

    let transactionCount = 0;
    for (let i = headerRowIndex+1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;
        const dateCell = row[dateCol];
        if (dateCell == null) continue;
        const dateStr = dateCell.toString().trim();
        const date = parseDate(dateStr);
        if (!date) continue;
        const balanceCell = row[balanceCol];
        if (balanceCell == null) continue;
        const balance = parseNumber(balanceCell.toString());
        if (isNaN(balance)) continue;
        transactions.push({ date, balance });
        transactionCount++;
    }
    console.log(`Berhasil membaca ${transactionCount} transaksi`);

    transactions.sort((a, b) => a.date - b.date);

    // Ambil nomor rekening dari baris 13 kolom J (index 9, 0-based)
    // Format SheetJS: rows[12][9] = '1570004240025'
    let accountId = null;
    for (let row of rows) {
        if (!Array.isArray(row)) continue;
        const rowText = row.join(' ').toLowerCase();
        if (rowText.includes('nomor rekening') || rowText.includes('account number')) {
            // Cari angka panjang (≥6 digit) di baris yang sama
            for (let cell of row) {
                if (cell == null) continue;
                const s = cell.toString().trim();
                if (/^\d{6,}$/.test(s)) { accountId = s; break; }
            }
            if (accountId) break;
        }
    }
    // Fallback: cari di fullText
    if (!accountId) {
        const m = fullText.match(/(?:nomor rekening|account number)[^:\d]*:?\s*(\d{6,})/i);
        if (m) accountId = m[1];
    }
    console.log('Mandiri accountId:', accountId);
    return { transactions, accountId };
}
function getCellValue(cell) {
    if (cell == null || cell === undefined) return null;
    if (typeof cell === 'string') return cell;
    if (typeof cell === 'number') return cell;
    if (typeof cell === 'boolean') return cell;
    if (cell instanceof Date) return cell;
    // ExcelJS RichText: { richText: [{text: '...'}, ...] }
    if (cell && Array.isArray(cell.richText)) {
        return cell.richText.map(rt => rt.text || '').join('');
    }
    // ExcelJS formula result
    if (cell && cell.result !== undefined) return cell.result;
    if (cell && cell.text !== undefined) return cell.text;
    return String(cell);
}

// ====================== BACA FILE EXCEL DENGAN SHEETJS (UTAMA) ======================
async function readExcelWithSheetJS(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const wb = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
                const ws = wb.Sheets[wb.SheetNames[0]];
                // header:1 → array-of-arrays; defval:null → missing cells = null
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
                resolve(rows);
            } catch (err) {
                const msg = (err.message || '').toLowerCase();
                if (msg.includes('password') || msg.includes('encrypted') || msg.includes('cfb')) {
                    reject(new Error('password_required'));
                } else {
                    reject(err);
                }
            }
        };
        reader.onerror = () => reject(new Error('Gagal membaca file'));
        reader.readAsArrayBuffer(file);
    });
}

// ====================== BACA FILE EXCEL DENGAN EXCELJS (FALLBACK, MENDUKUNG PASSWORD) ======================
async function readExcelWithPassword(file, password) {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.load(buffer, { password: password });
    } catch (loadError) {
        const msg = loadError.message.toLowerCase();
        if (msg.includes('central directory') || msg.includes('password') || msg.includes('encrypted')) {
            if (password) throw new Error('wrong_password');
            else throw new Error('password_required');
        }
        throw new Error(`Gagal membaca file: ${loadError.message}`);
    }
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) throw new Error('Tidak ada sheet dalam file.');
    const rows = [];
    worksheet.eachRow({ includeEmpty: true }, (row) => {
        // getCellValue menangani RichText dan tipe aneh lainnya dari ExcelJS
        rows.push(row.values.slice(1).map(getCellValue));
    });
    return rows;
}

// ====================== COBA BACA DENGAN CACHE PASSWORD ======================
async function tryReadWithCachedPasswords(file) {
    // Coba SheetJS terlebih dahulu (lebih andal untuk file tidak berpassword)
    try {
        const rows = await readExcelWithSheetJS(file);
        return { success: true, rows, passwordUsed: null };
    } catch (err) {
        if (err.message !== 'password_required') throw err;
        // File terenkripsi → lanjutkan ke ExcelJS dengan password
    }
    // Coba password yang sudah tersimpan menggunakan ExcelJS
    for (let [bank, pwd] of Object.entries(bankPasswords)) {
        try {
            const rows = await readExcelWithPassword(file, pwd);
            return { success: true, rows, passwordUsed: pwd, bankUsed: bank };
        } catch (err) {
            if (err.message !== 'wrong_password') throw err;
        }
    }
    return { success: false, needPassword: true };
}

// ====================== PROSES SEMUA FILE ======================
async function processAllFiles() {
    for (let f of uploadedFiles) {
        f.transactions = null;
        f.error = null;
        f.bank = guessBankFromFilename(f.file.name);
    }
    renderFileList();
    for (let f of uploadedFiles) {
        const isPDF = f.file.type === 'application/pdf' || f.file.name.toLowerCase().endsWith('.pdf');

        if (isPDF) {
            try {
                const lines = await readPDF(f.file);
                const detectedBank = detectBankFromPDFLines(lines);
                f.bank = detectedBank !== 'unknown' ? detectedBank : (f.bank || 'unknown');

                if (f.bank === 'danamon') {
                    const result = parseDanamonPDF(lines);
                    f.transactions = result.transactions;
                    f.accountId = result.accountId;
                } else {
                    f.error = `PDF bank "${f.bank || 'tidak dikenal'}" belum didukung. Tersedia: Danamon.`;
                }
            } catch (err) {
                f.error = err.message;
            }
            renderFileList();
            continue;
        }
        try {
            const result = await tryReadWithCachedPasswords(f.file);
            if (!result.success) {
                const password = await askPasswordForFile(f.file.name);
                if (!password) {
                    f.error = 'Password tidak diberikan.';
                    renderFileList();
                    continue;
                }
                try {
                    const rows = await readExcelWithPassword(f.file, password);
                    const detectedBank = detectBankFromExcel(rows);
                    f.bank = detectedBank !== 'unknown' ? detectedBank : f.bank;
                    if (f.bank !== 'unknown') bankPasswords[f.bank] = password;
                    if (f.bank === 'mandiri') {
                        const parsed = parseMandiriExcel(rows);
                        f.transactions = parsed.transactions;
                        f.accountId = parsed.accountId;
                    } else f.error = `Bank ${f.bank} belum memiliki parser.`;
                } catch (pwdErr) {
                    f.error = pwdErr.message === 'wrong_password' ? 'Password salah.' : pwdErr.message;
                }
            } else {
                const rows = result.rows;
                const detectedBank = detectBankFromExcel(rows);
                f.bank = detectedBank !== 'unknown' ? detectedBank : f.bank;
                if (result.passwordUsed && f.bank !== 'unknown') bankPasswords[f.bank] = result.passwordUsed;
                if (f.bank === 'mandiri') {
                    const parsed = parseMandiriExcel(rows);
                    f.transactions = parsed.transactions;
                    f.accountId = parsed.accountId;
                } else f.error = `Bank ${f.bank} belum memiliki parser.`;
            }
        } catch (err) {
            f.error = err.message;
        }
        renderFileList();
    }
}

function renderFileList() {
    fileListDiv.innerHTML = '';
    uploadedFiles.forEach((f) => {
        const div = document.createElement('div');
        div.className = 'file-item';
        const rekLabel = f.accountId ? ` [${f.accountId}]` : '';
        div.innerHTML = `
            <span class="bank">${(f.bank || 'Deteksi...').toUpperCase()}${rekLabel} — ${f.file.name}</span>
            <span class="status ${f.error ? 'error' : ''}">
                ${f.error ? `❌ ${f.error}` : (f.transactions ? `✅ ${f.transactions.length} tx` : '⏳')}
            </span>
        `;
        fileListDiv.appendChild(div);
    });
}

function askPasswordForFile(fileName) {
    return new Promise((resolve) => {
        currentFileForPassword = fileName;
        pendingPasswordResolve = resolve;
        if (bankNameSpan) bankNameSpan.innerText = 'File ini';
        if (fileNamesSpan) fileNamesSpan.innerText = fileName;
        if (passwordModal) passwordModal.style.display = 'block';
    });
}

// ====================== GENERATE DAILY BALANCE ======================
function generateDailyBalance(transactions, start, end) {
    const dailyMap = new Map();
    transactions.forEach(t => {
        const dateStr = toLocalDateStr(t.date);
        dailyMap.set(dateStr, t.balance);
    });

    const result = new Map();
    let current = new Date(start);
    while (current <= end) {
        const dateStr = toLocalDateStr(current);
        result.set(dateStr, dailyMap.has(dateStr) ? dailyMap.get(dateStr) : null);
        current.setDate(current.getDate() + 1);
    }

    let prev = null;
    for (let [date, bal] of result) {
        if (bal !== null) {
            prev = bal;
        } else {
            result.set(date, prev);
        }
    }
    return result;
}

function sumDailyBalances(maps) {
    const firstMap = maps[0];
    const total = new Map();
    for (let [date, _] of firstMap) {
        let sum = 0;
        for (let map of maps) {
            const val = map.get(date);
            sum += val !== null ? val : 0;
        }
        total.set(date, sum);
    }
    return total;
}

// ====================== SIMPAN & MUAT DATA ======================
const STORAGE_KEY = 'zakatMal_savedData';
function saveToStorage(totalMap, minDate, minBalance, nisab, goldPrice, startDateStr, endDateStr) {
    try {
        const payload = {
            savedAt: new Date().toISOString(),
            startDate: startDateStr,
            endDate: endDateStr,
            goldPrice,
            nisab,
            minDate,
            minBalance,
            dailyBalances: Array.from(totalMap.entries()),
        };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (e) { console.warn('Gagal menyimpan data:', e); }
}
function loadFromStorage() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return null;
        const payload = JSON.parse(raw);
        payload.totalMap = new Map(payload.dailyBalances);
        return payload;
    } catch (e) { return null; }
}
function clearStorage() { localStorage.removeItem(STORAGE_KEY); }
function restoreSavedResult(payload) {
    document.getElementById('start_date').value = payload.startDate;
    document.getElementById('end_date').value = payload.endDate;
    const goldInput = document.getElementById('gold_price');
    goldInput.value = payload.goldPrice.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    const nisab = payload.nisab, minBalance = payload.minBalance, minDate = payload.minDate;
    const wajib = minBalance >= nisab, zakat = wajib ? minBalance * 0.025 : 0;
    minBalanceSpan.innerText = formatRupiah(minBalance);
    nisabSpan.innerText = formatRupiah(nisab);
    statusSpan.innerText = wajib ? 'Wajib Zakat' : 'Tidak Wajib Zakat';
    zakatSpan.innerText = formatRupiah(zakat);
    const minDateSpan = document.getElementById('min_balance_date');
    if (minDateSpan && minDate) {
        const [y, m, d] = minDate.split('-');
        const dateObj = new Date(y, m-1, d);
        minDateSpan.innerText = '(pada ' + dateObj.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }) + ')';
    }
    resultDiv.style.display = 'block';
    errorDiv.style.display = 'none';
    lastTotalMap = payload.totalMap;
    renderChart(payload.totalMap, minDate, minBalance, nisab);
    const savedInfo = document.getElementById('savedDataInfo');
    if (savedInfo) {
        const savedAt = new Date(payload.savedAt).toLocaleString('id-ID');
        savedInfo.innerHTML = `📂 Menampilkan data tersimpan dari <strong>${savedAt}</strong>. 
            <button type="button" id="clearSavedData" class="btn-small btn-danger">Hapus Data Tersimpan</button>`;
        savedInfo.style.display = 'block';
        document.getElementById('clearSavedData')?.addEventListener('click', () => {
            clearStorage();
            savedInfo.style.display = 'none';
            resultDiv.style.display = 'none';
        });
    }
}

// ====================== HITUNG ZAKAT ======================
function calculateZakat() {
    const startDate = new Date(document.getElementById('start_date').value);
    const endDate = new Date(document.getElementById('end_date').value);
    const goldPriceStr = document.getElementById('gold_price').value.replace(/\./g, '');
    const goldPrice = parseFloat(goldPriceStr);
    if (!startDate || !endDate || isNaN(goldPrice)) {
        showError('Tanggal dan harga emas harus diisi dengan benar.');
        return;
    }
    if (startDate > endDate) {
        showError('Tanggal mulai harus sebelum tanggal akhir.');
        return;
    }
    const options = { day: '2-digit', month: 'long', year: 'numeric' };
    periodeStartSpan.innerText = startDate.toLocaleDateString('id-ID', options);
    periodeEndSpan.innerText = endDate.toLocaleDateString('id-ID', options);

    const allFiles = uploadedFiles.filter(f => f.transactions && f.transactions.length > 0);
    if (allFiles.length === 0) {
        showError('Tidak ada data transaksi valid.');
        return;
    }

    // Kelompokkan per rekening (accountId) agar beberapa PDF bulanan rekening yang sama
    // digabung sebagai satu kolom, bukan dijumlah
    const bankGroups = new Map(); // key → { label, txList }
    for (let f of allFiles) {
        // Gunakan accountId sebagai key utama; fallback ke bank+index jika tidak tersedia
        const key = f.accountId || `${f.bank || 'unknown'}_${allFiles.indexOf(f)}`;
        const bankName = (f.bank || 'rek').toUpperCase();
        const label = f.accountId
            ? `${bankName}-${f.accountId.slice(-4)}`   // contoh: MANDIRI-0025
            : bankName;
        if (!bankGroups.has(key)) bankGroups.set(key, { label, txList: [] });
        bankGroups.get(key).txList.push(...f.transactions);
    }

    const dailyMaps = [];
    lastBankMaps = [];
    for (let [, { label, txList }] of bankGroups) {
        txList.sort((a, b) => a.date - b.date);
        const map = generateDailyBalance(txList, startDate, endDate);
        dailyMaps.push(map);
        lastBankMaps.push({ label, map });
    }

    const totalMap = sumDailyBalances(dailyMaps);
    lastTotalMap = totalMap;

    let minBalance = Infinity, minDate = null;
    for (let [date, bal] of totalMap) {
        if (bal === null) {
            showError('Tidak ada data saldo untuk periode tersebut.');
            return;
        }
        if (bal < minBalance) {
            minBalance = bal;
            minDate = date;
        }
    }

    const nisab = 85 * goldPrice;
    const wajib = minBalance >= nisab;
    const zakat = wajib ? minBalance * 0.025 : 0;

    minBalanceSpan.innerText = formatRupiah(minBalance);
    nisabSpan.innerText = formatRupiah(nisab);
    statusSpan.innerText = wajib ? 'Wajib Zakat' : 'Tidak Wajib Zakat';
    zakatSpan.innerText = formatRupiah(zakat);
    const minDateSpan = document.getElementById('min_balance_date');
    if (minDateSpan && minDate) {
        const [y, m, d] = minDate.split('-');
        const dateObj = new Date(y, m-1, d);
        minDateSpan.innerText = '(pada ' + dateObj.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }) + ')';
    }
    resultDiv.style.display = 'block';
    errorDiv.style.display = 'none';

    const startDateStr = document.getElementById('start_date').value;
    const endDateStr   = document.getElementById('end_date').value;
    saveToStorage(totalMap, minDate, minBalance, nisab, goldPrice, startDateStr, endDateStr);

    const savedInfo = document.getElementById('savedDataInfo');
    if (savedInfo) savedInfo.style.display = 'none';

    renderChart(totalMap, minDate, minBalance, nisab);
}

// ====================== DOWNLOAD CSV ======================
function downloadCSV() {
    if (!lastTotalMap || lastTotalMap.size === 0) {
        alert('Tidak ada data saldo untuk diunduh.');
        return;
    }

    const fmt = n => `"${(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}"`;
    const fmtDate = date => { const [y,m,d] = date.split('-'); return `${m}/${d}/${y}`; };

    // Header: Tanggal | rek-1 | rek-2 | ... | Total
    const bankLabels = lastBankMaps.map(b => b.label);
    const header = ['Tanggal', ...bankLabels, 'Total'].join(',');

    const sortedDates = Array.from(lastTotalMap.keys()).sort();
    const rows = sortedDates.map(date => {
        const cols = [fmtDate(date)];
        for (const { map } of lastBankMaps) {
            cols.push(fmt(map.get(date) ?? 0));
        }
        cols.push(fmt(lastTotalMap.get(date) ?? 0));
        return cols.join(',');
    });

    const csvContent = [header, ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.setAttribute('download', 'saldo_harian.csv');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// ====================== RENDER GRAFIK ======================
function renderChart(totalMap, minDate, minBalance, nisab) {
    const canvas = document.getElementById('balanceChart');
    if (!canvas) return;
    const labels = [], data = [];
    for (let [date, bal] of totalMap) {
        labels.push(date);
        data.push(bal);
    }
    if (balanceChartInstance) {
        balanceChartInstance.destroy();
        balanceChartInstance = null;
    }
    function parseDateLocal(str) {
        const [y, m, d] = str.split('-').map(Number);
        return new Date(y, m-1, d);
    }
    function fmtDate(str, opts) {
        return parseDateLocal(str).toLocaleDateString('id-ID', opts);
    }
    const allVals = data.filter(v => v !== null);
    const dataMin = Math.min(...allVals);
    const dataMax = Math.max(...allVals, nisab);
    const range = dataMax - dataMin || 1;
    const yMin = Math.max(0, dataMin - range * 0.12);
    const yMax = dataMax + range * 0.18;
    const pointRadius = labels.map(d => d === minDate ? 7 : 0);
    const pointColors = labels.map(d => d === minDate ? '#e74c3c' : '#2c3e50');
    const minIdx = labels.indexOf(minDate);
    const minDateShort = fmtDate(minDate, { day: 'numeric', month: 'short', year: 'numeric' });
    const labelXAdjust = minIdx < labels.length * 0.4 ? 70 : -70;
    const labelYAdjust = minBalance < (yMin + (yMax - yMin) * 0.3) ? -60 : 60;

    balanceChartInstance = new Chart(canvas, {
        type: 'line',
        data: { labels, datasets: [{
            label: 'Saldo Harian',
            data,
            borderColor: '#2c3e50',
            backgroundColor: 'rgba(44,62,80,0.07)',
            borderWidth: 1.8,
            fill: true,
            tension: 0.15,
            pointBackgroundColor: pointColors,
            pointBorderColor: '#fff',
            pointBorderWidth: 2,
            pointRadius,
            pointHoverRadius: 5,
        }] },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(30,40,50,0.9)',
                    callbacks: {
                        title: ctx => fmtDate(ctx[0].label, { day: 'numeric', month: 'long', year: 'numeric' }),
                        label: ctx => 'Saldo: Rp ' + formatRupiah(Math.round(ctx.parsed.y)),
                        afterLabel: ctx => ctx.label === minDate ? '** Saldo Terendah **' : ''
                    }
                },
                annotation: {
                    annotations: {
                        nisabLine: {
                            type: 'line',
                            yMin: nisab, yMax: nisab,
                            borderColor: '#e67e22',
                            borderWidth: 1.5,
                            borderDash: [6,4],
                            label: {
                                display: true,
                                content: 'Nisab: Rp ' + formatRupiah(Math.round(nisab)),
                                position: 'end',
                                backgroundColor: '#e67e22',
                                color: '#fff',
                                font: { size: 11 },
                                padding: { x:7, y:4 },
                                borderRadius:4,
                            }
                        },
                        minLabel: {
                            type: 'label',
                            xValue: minIdx,
                            yValue: minBalance,
                            xAdjust: labelXAdjust,
                            yAdjust: labelYAdjust,
                            content: ['Saldo Terendah', 'Rp ' + formatRupiah(Math.round(minBalance)), minDateShort],
                            backgroundColor: 'rgba(231,76,60,0.92)',
                            color: '#fff',
                            font: { size:11, weight:'bold' },
                            padding: { x:10, y:6 },
                            borderRadius:5,
                            callout: { display:true, borderColor:'#e74c3c', borderWidth:1.5 }
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        maxTicksLimit: 9,
                        callback: function(val) {
                            const lbl = this.getLabelForValue(val);
                            return lbl ? fmtDate(lbl, { day:'numeric', month:'short' }) : '';
                        }
                    },
                    grid: { display:false }
                },
                y: {
                    min: yMin, max: yMax,
                    ticks: {
                        callback: val => val >= 1e6 ? 'Rp ' + (val/1e6).toFixed(1) + ' jt' : 'Rp ' + formatRupiah(val),
                        maxTicksLimit:7
                    }
                }
            }
        }
    });
}

function showError(msg) {
    errorDiv.innerText = msg;
    errorDiv.style.display = 'block';
    resultDiv.style.display = 'none';
}

// ====================== INISIALISASI ======================
document.addEventListener('DOMContentLoaded', function() {
    // Inisialisasi PDF.js worker
    if (typeof pdfjsLib !== 'undefined') {
        pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    // Ambil elemen DOM
    fileInput = document.getElementById('files');
    fileListDiv = document.getElementById('fileList');
    calculateBtn = document.getElementById('calculateBtn');
    passwordModal = document.getElementById('passwordModal');
    bankNameSpan = document.getElementById('bankName');
    fileNamesSpan = document.getElementById('fileNames');
    bankPasswordInput = document.getElementById('bankPassword');
    submitPasswordBtn = document.getElementById('submitPassword');
    cancelPasswordBtn = document.getElementById('cancelPassword');
    resultDiv = document.getElementById('result');
    errorDiv = document.getElementById('error');
    minBalanceSpan = document.getElementById('min_balance');
    nisabSpan = document.getElementById('nisab');
    statusSpan = document.getElementById('status');
    zakatSpan = document.getElementById('zakat');
    periodeStartSpan = document.getElementById('periode_start');
    periodeEndSpan = document.getElementById('periode_end');

    // Tombol download
    const downloadBtn = document.getElementById('downloadCsv');
    if (downloadBtn) downloadBtn.addEventListener('click', downloadCSV);

    // Tampilkan harga emas terkini
    const currentGoldDisplay = document.getElementById('currentGoldDisplay');
    if (currentGoldDisplay) currentGoldDisplay.innerText = formatRupiah(CURRENT_GOLD_PRICE);

    // Tombol gunakan harga emas
    const useCurrentPriceBtn = document.getElementById('useCurrentPrice');
    if (useCurrentPriceBtn) {
        useCurrentPriceBtn.addEventListener('click', function() {
            const priceInput = document.getElementById('gold_price');
            if (priceInput) {
                priceInput.value = CURRENT_GOLD_PRICE.toString();
                priceInput.dispatchEvent(new Event('input', { bubbles: true }));
            }
        });
    }

    // Tampilkan teks rekomendasi tanggal
    const recStart = document.getElementById('recommendedStartText');
    if (recStart) recStart.innerText = RECOMMENDED_START_TEXT;
    const recEnd = document.getElementById('recommendedEndText');
    if (recEnd) recEnd.innerText = RECOMMENDED_END_TEXT;

    // Tombol gunakan tanggal mulai
    const useStartBtn = document.getElementById('useStartDate');
    if (useStartBtn) {
        useStartBtn.addEventListener('click', function() {
            const startInput = document.getElementById('start_date');
            if (startInput) startInput.value = RECOMMENDED_START_DATE;
        });
    }

    // Tombol gunakan tanggal akhir
    const useEndBtn = document.getElementById('useEndDate');
    if (useEndBtn) {
        useEndBtn.addEventListener('click', function() {
            const endInput = document.getElementById('end_date');
            if (endInput) endInput.value = RECOMMENDED_END_DATE;
        });
    }

    // Tombol lihat harga emas
    const viewGoldBtn = document.getElementById('viewGoldPrice');
    if (viewGoldBtn) {
        viewGoldBtn.addEventListener('click', (e) => {
            e.preventDefault();
            window.open('https://harga-emas.org', '_blank');
        });
    }

    // Format input harga emas
    const goldInput = document.getElementById('gold_price');
    if (goldInput) {
        goldInput.addEventListener('input', function(e) {
            let raw = e.target.value.replace(/[^\d]/g, '');
            if (raw) {
                let formatted = raw.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
                if (e.target.value !== formatted) e.target.value = formatted;
            }
        });
    }

    // Event listener file input
    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            uploadedFiles = files.map(file => ({ file, bank: null, password: null, transactions: null, error: null }));
            renderFileList();
            processAllFiles().catch(err => console.error(err));
        });
    }

    // Hitung zakat
    if (calculateBtn) {
        calculateBtn.addEventListener('click', () => {
            if (uploadedFiles.length === 0) {
                showError('Unggah file terlebih dahulu.');
                return;
            }
            const unfinished = uploadedFiles.some(f => f.transactions === null && !f.error);
            if (unfinished) {
                showError('Tunggu hingga semua file selesai diproses.');
                return;
            }
            calculateZakat();
        });
    }

    // Modal password
    if (submitPasswordBtn) {
        submitPasswordBtn.addEventListener('click', () => {
            if (pendingPasswordResolve) {
                pendingPasswordResolve(bankPasswordInput.value);
                pendingPasswordResolve = null;
            }
            passwordModal.style.display = 'none';
            bankPasswordInput.value = '';
        });
    }
    if (cancelPasswordBtn) {
        cancelPasswordBtn.addEventListener('click', () => {
            if (pendingPasswordResolve) {
                pendingPasswordResolve(null);
                pendingPasswordResolve = null;
            }
            passwordModal.style.display = 'none';
            bankPasswordInput.value = '';
        });
    }
    window.addEventListener('click', (e) => {
        if (e.target === passwordModal) cancelPasswordBtn?.click();
    });

    // Muat data tersimpan
    const saved = loadFromStorage();
    if (saved) restoreSavedResult(saved);
});