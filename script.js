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

// ====================== PARSER DANAMON PDF (IDR + USD → 2 rekening terpisah) ======================
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

    const RE_IDR_NUM = /\d+(?:\.\d{3})*,\d{2}/g;

    // ── Ekstrak kurs implisit dari tabel ringkasan ─────────────────────────────
    // Prioritaskan DANAMON LEBIH PRO (USD) row karena saldo lebih besar & akurat.
    // Fallback ke PRIMADOLLAR jika USD-0747 = 0.
    let kursImplisit = null;
    let accountIdUSD = null;  // accountId PRIMADOLLAR
    for (const line of lines) {
        // Selalu catat accountId PRIMADOLLAR
        if (/PRIMADOLLAR/i.test(line) && !accountIdUSD) {
            const m = line.match(/(\d{10,12})/);
            if (m) accountIdUSD = m[1];
        }
        // Ambil kurs dari baris DANAMON LEBIH PRO (USD) atau PRIMADOLLAR
        if (kursImplisit) continue;
        if (!/DANAMON LEBIH PRO \(USD\)|PRIMADOLLAR/i.test(line)) continue;
        const idrs = [...line.matchAll(RE_IDR_NUM)].map(m => parseNumber(m[0]));
        if (idrs.length >= 2) {
            const usdSaldo = idrs[idrs.length - 2];
            const idrSaldo = idrs[idrs.length - 1];
            if (usdSaldo > 0 && idrSaldo > 0) {
                kursImplisit = idrSaldo / usdSaldo;
                console.log(`Danamon kurs implisit: ${idrSaldo} / ${usdSaldo} = ${kursImplisit.toFixed(2)}`);
            }
        }
    }
    if (!kursImplisit) {
        console.warn('Kurs implisit tidak ditemukan, fallback ke 16.000');
        kursImplisit = 16000;
    }

    // ── Helper ────────────────────────────────────────────────────────────────
    function stripRefNumbers(line) { return line.replace(/\d{13,}/g, ''); }

    function parseSection(headerRegex, stopRegex, isUSD) {
        let inRincian = false, inSection = false;
        let saldoAwal = null;
        const txList = [];
        for (const line of lines) {
            if (/RINCIAN TRANSAKSI PER REKENING/i.test(line)) { inRincian = true; continue; }
            if (!inRincian) continue;
            if (headerRegex.test(line)) { inSection = true; continue; }
            if (inSection && stopRegex.test(line)) break;
            if (!inSection) continue;
            const cleanLine = stripRefNumbers(line);
            if (/SALDO BULAN LALU/i.test(cleanLine)) {
                const vals = [...cleanLine.matchAll(RE_IDR_NUM)].map(m => parseNumber(m[0])).filter(n => n >= 0);
                saldoAwal = vals.length ? Math.min(...vals) : null;
                console.log(`Danamon saldo awal ${isUSD?'USD':'IDR'}:`, saldoAwal);
                continue;
            }
            const txMatch = line.match(/^(\d{2})\/(\d{2})\s+(\d{2})\/(\d{2})\s/);
            if (!txMatch) continue;
            const day = parseInt(txMatch[3]), month = parseInt(txMatch[4]);
            let year = periodYear;
            if (month === 12 && periodMonth === 1) year--;
            const date = new Date(year, month - 1, day);
            const vals = [...cleanLine.matchAll(RE_IDR_NUM)].map(m => parseNumber(m[0])).filter(n => n >= 0);
            if (!vals.length) continue;
            const threshold = isUSD ? 1_000_000 : 500_000_000;
            const balance = vals.length <= 3
                ? vals[vals.length - 1]
                : (vals.filter(n => n < threshold).slice(-1)[0] ?? vals[vals.length - 1]);
            if (!isNaN(balance)) txList.push({ date, balance });
            console.log(`Danamon ${isUSD?'USD':'IDR'} tx: ${toLocalDateStr(date)} bal=${balance}`);
        }
        if (saldoAwal !== null)
            txList.unshift({ date: new Date(periodYear, periodMonth - 1, 1), balance: saldoAwal });
        txList.sort((a, b) => a.date - b.date);
        return txList;
    }

    // ── Parse section IDR (903...0747) ────────────────────────────────────────
    // Stop saat masuk DANAMON LEBIH PRO (USD) ATAU PRIMADOLLAR
    const txIDR = parseSection(
        /DANAMON LEBIH PRO \(IDR\)\s*-\s*IDR\s*-/i,
        /DANAMON LEBIH PRO \(USD\)\s*-\s*USD\s*-|PRIMADOLLAR[^(]*-\s*USD\s*-|D-POINT BANKING|AKHIR LAPORAN/i,
        false
    );

    // ── Parse section USD sub-account (903...0747) ────────────────────────────
    // Hadir mulai Sept 2025; tidak ada = saldo nol, filter .length > 0 di akhir
    // Stop saat masuk PRIMADOLLAR
    const txDanaUSD_raw = parseSection(
        /DANAMON LEBIH PRO \(USD\)\s*-\s*USD\s*-/i,
        /PRIMADOLLAR[^(]*-\s*USD\s*-|D-POINT BANKING|AKHIR LAPORAN/i,
        true
    );
    const txDanaUSD_idr = txDanaUSD_raw.map(t => ({ date: t.date, balance: t.balance * kursImplisit }));

    // ── Parse section PRIMADOLLAR (003...5703) ────────────────────────────────
    // Stop hanya di D-POINT / AKHIR LAPORAN (rekening terakhir dalam PDF)
    const txPrima_raw = parseSection(
        /PRIMADOLLAR[^(]*-\s*USD\s*-/i,
        /D-POINT BANKING|AKHIR LAPORAN/i,
        true
    );
    const txPrima_idr = txPrima_raw.map(t => ({ date: t.date, balance: t.balance * kursImplisit }));

    if (txIDR.length === 0 && txDanaUSD_idr.length === 0 && txPrima_idr.length === 0)
        throw new Error('Data transaksi tidak ditemukan dalam PDF Danamon.');

    // Nomor rekening IDR
    let accountIdIDR = null;
    for (const line of lines) {
        const m = line.match(/DANAMON LEBIH PRO \(IDR\)\s*-\s*IDR\s*-\s*(\d{6,})/i);
        if (m) { accountIdIDR = m[1]; break; }
    }

    console.log(`Danamon: IDR(${accountIdIDR})=${txIDR.length}tx, USD-0747=${txDanaUSD_idr.length}tx, PRIMA(${accountIdUSD})=${txPrima_idr.length}tx, kurs=${kursImplisit.toFixed(2)}`);

    return [
        {
            accountId: accountIdIDR,
            key: accountIdIDR,
            label: `DANAMON-${(accountIdIDR||'IDR').slice(-4)}`,
            transactions: txIDR
        },
        {
            accountId: accountIdIDR,
            key: `${accountIdIDR}_USD`,
            label: `DANAMON-${(accountIdIDR||'USD').slice(-4)}U`,  // e.g. DANAMON-0747U
            transactions: txDanaUSD_idr
        },
        {
            accountId: accountIdUSD,
            key: accountIdUSD,
            label: `DANAMON-${(accountIdUSD||'PRIMA').slice(-4)}`, // e.g. DANAMON-5703
            transactions: txPrima_idr
        },
    ].filter(a => a.transactions.length > 0);
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
                    const accounts = parseDanamonPDF(lines); // array [{accountId, label, transactions}]
                    f.accounts    = accounts;
                    f.transactions = accounts.flatMap(a => a.transactions); // agar validasi "ada data" lolos
                    f.accountId   = accounts[0]?.accountId ?? null;
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
        let rekLabel = '';
        if (f.accounts && f.accounts.length > 0) {
            rekLabel = ' [' + f.accounts.map(a => a.label).join(', ') + ']';
        } else if (f.accountId) {
            const bankName = (f.bank || 'rek').toUpperCase();
            rekLabel = ` [${bankName}-${f.accountId.slice(-4)}]`;
        }
        const txCount = f.accounts
            ? f.accounts.reduce((s, a) => s + a.transactions.length, 0)
            : (f.transactions?.length ?? 0);
        div.innerHTML = `
            <span class="bank">${(f.bank || 'Deteksi...').toUpperCase()}${rekLabel} — ${f.file.name}</span>
            <span class="status ${f.error ? 'error' : ''}">
                ${f.error ? `❌ ${f.error}` : (f.transactions ? `✅ ${txCount} tx` : '⏳')}
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

    // Kumpulkan semua "rekening" dari semua file
    // File Danamon: punya f.accounts = [{accountId, label, transactions}, ...]
    // File lain (Mandiri, dll): satu rekening per file
    const allAccounts = []; // [{key, label, transactions}]
    for (const f of allFiles) {
        if (f.accounts && f.accounts.length > 0) {
            for (const acc of f.accounts) {
                allAccounts.push({
                    key: acc.key || acc.accountId || `${f.bank}_${allAccounts.length}`,
                    label: acc.label,
                    transactions: acc.transactions,
                });
            }
        } else {
            // Mandiri dll: satu rekening per file
            const bankName = (f.bank || 'rek').toUpperCase();
            const key = f.accountId || `${f.bank || 'unknown'}_${allAccounts.length}`;
            const label = f.accountId ? `${bankName}-${f.accountId.slice(-4)}` : bankName;
            allAccounts.push({ key, label, transactions: f.transactions });
        }
    }

    // Gabungkan transaksi yang punya accountId sama (misal beberapa PDF bulanan rekening yang sama)
    const bankGroups = new Map(); // key → { label, txList }
    for (const acc of allAccounts) {
        if (!bankGroups.has(acc.key)) bankGroups.set(acc.key, { label: acc.label, txList: [] });
        bankGroups.get(acc.key).txList.push(...acc.transactions);
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
	let zakat = wajib ? minBalance * 0.025 : 0;

	// pembulatan ke atas dalam ribuan
	if (zakat > 0) {
		zakat = Math.ceil(zakat / 1000) * 1000;
	}

    minBalanceSpan.innerText = formatRupiah(minBalance);
    nisabSpan.innerText = formatRupiah(nisab);
    statusSpan.innerText = wajib ? 'Wajib Zakat' : 'Tidak Wajib Zakat';
    zakatSpan.innerText = formatRupiah(zakat);
	const zakatNote = document.getElementById('zakat_note');
	if (zakatNote && zakat > 0) {
		zakatNote.innerText = '(dibulatkan ke atas ke ribuan)';
	}
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
	
	const copyBtn = document.getElementById('copyZakat');
	if(copyBtn){
		copyBtn.addEventListener('click', copyZakatValue);
	}

	const pdfBtn = document.getElementById('exportPdf');
	if(pdfBtn){
		pdfBtn.addEventListener('click', exportZakatPDF);
	}

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

async function copyZakatValue() {

    const zakatText = document.getElementById('zakat').innerText;

    try {

        await navigator.clipboard.writeText(zakatText);

        alert('Nilai zakat berhasil disalin:\nRp ' + zakatText);

    } catch (err) {

        alert('Gagal menyalin nilai zakat');

    }
}

function exportZakatPDF() {

    const { jsPDF } = window.jspdf;

    const doc = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4"
    });

    const periodeStart = document.getElementById('periode_start').innerText;
    const periodeEnd = document.getElementById('periode_end').innerText;

    const saldoMin = document.getElementById('min_balance').innerText;
    const saldoDate = document.getElementById('min_balance_date').innerText;

    const nisab = document.getElementById('nisab').innerText;
    const status = document.getElementById('status').innerText;
    const zakat = document.getElementById('zakat').innerText;

    let y = 20;

    doc.setFontSize(16);
    doc.text("Laporan Perhitungan Zakat Mal", 20, y);

    y += 10;

    doc.setFontSize(11);

    doc.text(`Periode: ${periodeStart} s.d. ${periodeEnd}`, 20, y);
    y += 8;

    doc.text(`Saldo terendah: Rp ${saldoMin} ${saldoDate}`, 20, y);
    y += 8;

    doc.text(`Nisab: Rp ${nisab}`, 20, y);
    y += 8;

    doc.text(`Status: ${status}`, 20, y);
    y += 8;

    doc.text(`Zakat: Rp ${zakat} (dibulatkan ke atas)`, 20, y);

    y += 12;

    // ambil grafik dari canvas
    const canvas = document.getElementById("balanceChart");

    if (canvas) {

        const imgData = canvas.toDataURL("image/png");

        const pageWidth = doc.internal.pageSize.getWidth();
        const chartWidth = pageWidth - 40;
        const chartHeight = chartWidth * 0.5;

        doc.addImage(
            imgData,
            "PNG",
            20,
            y,
            chartWidth,
            chartHeight
        );

        y += chartHeight + 10;
    }

    doc.setFontSize(9);

    doc.text(
        "Grafik menunjukkan perubahan saldo harian selama periode perhitungan zakat.",
        20,
        y
    );

    y += 6;

    doc.text(
        "Laporan dibuat otomatis oleh Kalkulator Zakat.",
        20,
        y
    );

    doc.save("laporan_zakat.pdf");
}