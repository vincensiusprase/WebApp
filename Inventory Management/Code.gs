// ╔══════════════════════════════════════════════════════════════════╗
// ║   Code.gs  —  ISG Inventory Management System              ║
// ║   Backend : Google Apps Script                             ║
// ║   Versi   : 2.0  (Final, sesuai SRS + klarifikasi)         ║
// ╚══════════════════════════════════════════════════════════════════╝
//
// ARSITEKTUR SINGKAT:
//  • Sheet "Expired Product & BS" → HANYA DIBACA oleh Code.gs.
//    Formula ArrayFormula diisi langsung di Google Sheets oleh user.
//  • Sheet "Inventory" → DITULIS oleh Code.gs setiap kali ada transaksi.
//  • Sheet "In-Out"    → DITULIS oleh Code.gs setiap kali ada transaksi.
//  • Sheet "Masterdata" & "Users" → HANYA DIBACA oleh Code.gs.
//
// PRINSIP EFISIENSI:
//  • Setiap fungsi hanya melakukan getValues() SATU KALI per sheet.
//  • Tidak ada query ke Spreadsheet di dalam loop.
//  • Data di-index ke JavaScript Object (Map) untuk lookup O(1).
// ─────────────────────────────────────────────────────────────────────

// ══════════════════════════════════════════════════════════════════════
// §0  KONSTANTA
// ══════════════════════════════════════════════════════════════════════

/** Nama sheet — ubah di sini jika nama sheet berubah */
const SH = {
  USERS    : 'Users',
  MASTER   : 'Masterdata',
  IN_OUT   : 'In-Out',
  INVENTORY: 'Inventory',
  EXPIRED  : 'Expired Product & BS',
};

/**
 * Header resmi setiap sheet.
 * Urutan ini HARUS SAMA dengan urutan kolom di Spreadsheet.
 */
const HDR = {
  USERS: [
    'Username','Password','Role'
  ],
  MASTER: [
    'SKU','Item Name','Kategori','SS','A-B-C',
    'Jenis','Tebal','Ukuran','Kualitas','Merk Kaca',
    'Kualitas + Supplier','Red Expired','Yellow Expired','Green Expired',
    'Supplier','Customer','Lot Number'
  ],
  IN_OUT: [
    'Inbound Date','Date','Status','Jenis','Tebal','Kualitas',
    'SKU#','SKU','Item Name','Qty In','Qty Out',
    'Customer','Supplier','Merk Kaca','No Dokumen','Lot Number'
  ],
  INVENTORY: [
    'SKU','Kategori','Item Name','Total In','Total Out',
    'Current Inventory','%Available','Expired Product','% Expired Product',
    'SS','A-B-C'
  ],
  EXPIRED: [
    'Inbound Date','Jenis','Kualitas','SKU','Supplier','Merk Kaca',
    'Lot Number','Umur Kaca (Bulan)','Exp Qty','Qty','Code',
    'Expired Status','Days in Inventory','SKU#','Check In','Check Out','Current BS'
  ],
};

// ══════════════════════════════════════════════════════════════════════
// §1  ENTRY POINT
// ══════════════════════════════════════════════════════════════════════

/** Entry point Google Apps Script Web App */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ISG — Inventory Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ══════════════════════════════════════════════════════════════════════
// §2  SETUP DATABASE
// ══════════════════════════════════════════════════════════════════════

/**
 * setupDatabase()
 * ─────────────────────────────────────────────────────────────────────
 * Jalankan SATU KALI dari Apps Script Editor (▶ Run) saat pertama kali
 * setup. Fungsi ini akan:
 *  1. Membuat sheet yang belum ada.
 *  2. Menulis header di baris 1 (hanya jika sheet kosong).
 *  3. Memberi styling header (bold, warna navy).
 *  4. Membekukan baris 1 (freeze header row).
 *  5. Menulis formula ArrayFormula di sheet Expired Product & BS.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  // ── Buat / styling semua sheet ─────────────────────────────────────
  Object.entries(SH).forEach(([key, name]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      Logger.log(`✅ Sheet "${name}" dibuat.`);
    }
    // Tulis header hanya jika baris 1 kosong
    if (sh.getLastColumn() === 0 || sh.getRange(1,1).getValue() === '') {
      const headers = HDR[key];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      _styleHeader(sh, headers.length);
      sh.setFrozenRows(1);
      Logger.log(`  → Header "${name}" ditulis.`);
    }
  });

  // ── Tulis ArrayFormula di sheet Expired Product & BS ──────────────
  // PENTING: Formula ini menggantikan entri manual.
  // Code.gs hanya akan MEMBACA sheet ini — tidak pernah menulis data baris.
  _writeExpiredFormulas(ss);

  // ── Notifikasi ─────────────────────────────────────────────────────
  SpreadsheetApp.getUi().alert(
    '✅ Setup selesai!\n\n' +
    'Semua sheet & header berhasil dibuat.\n' +
    'Formula otomatis sudah dipasang di sheet "Expired Product & BS".\n\n' +
    'Langkah berikutnya:\n' +
    '1. Isi sheet Users (tambahkan akun login).\n' +
    '2. Isi sheet Masterdata (data produk).\n' +
    '3. Deploy sebagai Web App.'
  );
}

/** Helper: styling header row */
function _styleHeader(sh, colCount) {
  const r = sh.getRange(1, 1, 1, colCount);
  r.setFontWeight('bold')
   .setBackground('#0d1f35')
   .setFontColor('#ffffff')
   .setHorizontalAlignment('center')
   .setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
}

/**
 * Tulis formula ArrayFormula di sheet Expired Product & BS.
 * Hanya menulis ke baris 2 (formula akan auto-expand ke bawah).
 * Jika formula sudah ada, dilewati agar tidak overwrite data.
 */
function _writeExpiredFormulas(ss) {
  const sh = ss.getSheetByName(SH.EXPIRED);
  if (!sh) return;

  // Cek apakah formula sudah ada di A2
  const a2 = sh.getRange('A2').getFormula();
  if (a2 && a2.length > 5) {
    Logger.log('  ℹ️  Formula Expired sudah ada, dilewati.');
    return;
  }

  // A2:G2 — QUERY dari In-Out (Inbound Date, Jenis, Kualitas, SKU, Supplier, Merk Kaca, Lot Number)
  sh.getRange('A2').setFormula(
    "=IFERROR(UNIQUE(QUERY('In-Out'!A2:P,\"select A,D,F,H,M,N,P WHERE C='Inbound'\",-1)),)"
  );

  // H2 — Umur Kaca (Bulan): selisih bulan dari Inbound Date ke hari ini
  sh.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IFERROR(DATEDIF(A2:A,NOW(),"m"),)))'
  );

  // I2 — Exp Qty: hanya tampil jika Expired Status = "Merah"
  sh.getRange('I2').setFormula(
    '=MAP(G2:G,L2:L,LAMBDA(vG,vL,IF(vL="Merah",SUMIF(\'In-Out\'!P:P,vG,\'In-Out\'!J:J)-SUMIF(\'In-Out\'!P:P,vG,\'In-Out\'!K:K),"")))'
  );

  // J2 — Qty: saldo per Lot Number (In - Out)
  sh.getRange('J2').setFormula(
    '=MAP(A2:A,G2:G,LAMBDA(vA,vG,IF(vA="","",SUMIF(\'In-Out\'!P:P,vG,\'In-Out\'!J:J)-SUMIF(\'In-Out\'!P:P,vG,\'In-Out\'!K:K))))'
  );

  // K2 — Code: gabungan Jenis + Merk Kaca + Kualitas (untuk VLOOKUP ke Masterdata)
  sh.getRange('K2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",B2:B&" "&F2:F&" "&C2:C))'
  );

  // L2 — Expired Status: Merah / Kuning / Hijau berdasarkan VLOOKUP ke Masterdata
  sh.getRange('L2').setFormula(
    '=ARRAYFORMULA(IFNA(IF(H2:H>=VLOOKUP(K2:K,Masterdata!$K$1:$N,2,0),"Merah",' +
    'IF(H2:H>=VLOOKUP(K2:K,Masterdata!$K$1:$N,3,0),"Kuning",' +
    'IF(H2:H<=VLOOKUP(K2:K,Masterdata!$K$1:$N,4,0),"Hijau",""))),""))'
  );

  // M2 — Days in Inventory: bulan dari Inbound Date ke tanggal keluar pertama per Lot
  sh.getRange('M2').setFormula(
    '=MAP(A2:A,G2:G,LAMBDA(vA,vG,IF(vA="","",IFERROR(' +
    'DATEDIF(vA,XLOOKUP(vG,\'In-Out\'!P:P,\'In-Out\'!B:B,,0,-1),"m"),))))'
  );

  Logger.log('  ✅ Formula Expired Product & BS berhasil ditulis.');
}

// ══════════════════════════════════════════════════════════════════════
// §3  AUTENTIKASI
// ══════════════════════════════════════════════════════════════════════

/**
 * login(username, password)
 * ─────────────────────────────────────────────────────────────────────
 * Validasi kredensial dari sheet Users.
 * Return: { success:true, username, role } atau { success:false, message }
 *
 * EFISIENSI: Baca sheet Users SEKALI ke array, lalu loop di memori.
 */
function login(username, password) {
  try {
    const sh   = _getSheet(SH.USERS);
    const rows = _readAllRows(sh, HDR.USERS.length);   // array 2D, baris 2–last
    const C    = _idx(HDR.USERS);

    for (const row of rows) {
      const u = String(row[C['Username']] || '').trim().toLowerCase();
      const p = String(row[C['Password']] || '').trim();
      const r = String(row[C['Role']]     || '').trim();

      if (u === username.trim().toLowerCase() && p === password.trim()) {
        return { success: true, username: row[C['Username']], role: r };
      }
    }
    return { success: false, message: 'Username atau password salah.' };

  } catch (e) {
    return { success: false, message: 'Error sistem: ' + e.message };
  }
}

// ══════════════════════════════════════════════════════════════════════
// §4  MASTERDATA — Dropdown & Referensi
// ══════════════════════════════════════════════════════════════════════

/**
 * getMasterdata()
 * ─────────────────────────────────────────────────────────────────────
 * Ambil semua data referensi yang dibutuhkan oleh form Inbound/Outbound.
 * SATU KALI baca sheet Masterdata → hasilkan objek terstruktur.
 *
 * Return:
 * {
 *   skus       : ['SKU1','SKU2',...],          // sorted ascending
 *   skuMap     : { SKU: { itemName, merkKaca, jenis, tebal } },
 *   merkList   : ['Asahi','Guardian',...],     // unique, sorted
 *   merkBySKU  : { SKU: 'Asahi' },
 *   suppliers  : ['Supplier A',...],
 *   customers  : ['Customer X',...],
 *   lotNumbers : ['LOT001',...],
 * }
 */
function getMasterdata() {
  try {
    const sh   = _getSheet(SH.MASTER);
    const rows = _readAllRows(sh, HDR.MASTER.length);
    const C    = _idx(HDR.MASTER);

    const skuSet  = new Set();
    const skuMap  = {};
    const merkSet = new Set();
    const supSet  = new Set();
    const custSet = new Set();
    const lotSet  = new Set();

    rows.forEach(row => {
      const sku      = _str(row[C['SKU']]);
      const itemName = _str(row[C['Item Name']]);
      const merk     = _str(row[C['Merk Kaca']]);
      const jenis    = _str(row[C['Jenis']]);
      const tebal    = _str(row[C['Tebal']]);

      // Supplier & Customer bisa multi-value dipisahkan koma
      const supRaw   = _str(row[C['Supplier']]);
      const custRaw  = _str(row[C['Customer']]);
      const lot      = _str(row[C['Lot Number']]);

      if (!sku) return; // skip baris kosong

      skuSet.add(sku);
      skuMap[sku] = { itemName, merkKaca: merk, jenis, tebal };
      if (merk) merkSet.add(merk);

      supRaw.split(',').map(s => s.trim()).filter(Boolean).forEach(s => supSet.add(s));
      custRaw.split(',').map(s => s.trim()).filter(Boolean).forEach(s => custSet.add(s));
      if (lot) lot.split(',').map(s => s.trim()).filter(Boolean).forEach(s => lotSet.add(s));
    });

    return {
      skus      : [...skuSet].sort(),
      skuMap,
      merkList  : [...merkSet].sort(),
      suppliers : [...supSet].sort(),
      customers : [...custSet].sort(),
      lotNumbers: [...lotSet].sort(),
    };

  } catch (e) {
    return { error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════════════
// §5  TRANSAKSI INBOUND
// ══════════════════════════════════════════════════════════════════════

/**
 * submitInbound(payload)
 * ─────────────────────────────────────────────────────────────────────
 * Proses transaksi barang masuk (penerimaan dari supplier).
 *
 * Flow:
 *  1. Validasi input.
 *  2. Ambil referensi SKU dari Masterdata (1x baca, O(1) lookup).
 *  3. Append baris ke sheet In-Out.
 *  4. Update (atau insert) baris SKU di sheet Inventory.
 *
 * payload: {
 *   inboundDate, date, sku, merkKaca, qtyIn,
 *   supplier, lotNumber, noDO
 * }
 */
function submitInbound(payload) {
  try {
    // ── Validasi wajib ──────────────────────────────────────────────
    if (!payload.sku)           return _err('SKU wajib dipilih.');
    if (!payload.qtyIn)         return _err('Qty Masuk wajib diisi.');
    const qty = Number(payload.qtyIn);
    if (isNaN(qty) || qty <= 0) return _err('Qty Masuk harus angka positif.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Ambil referensi Masterdata (1x baca) ────────────────────────
    const masterRef = _getMasterRef(ss, payload.sku);

    // ── Parse komponen SKU ──────────────────────────────────────────
    const parsed = _parseSKU(payload.sku);

    // ── Tulis ke sheet In-Out ────────────────────────────────────────
    const inOutSh = _getSheet(SH.IN_OUT, ss);
    inOutSh.appendRow([
      _fmtDate(payload.inboundDate),   // col A: Inbound Date
      _fmtDate(payload.date),           // col B: Date
      'Inbound',                         // col C: Status
      parsed.jenis,                      // col D: Jenis
      parsed.tebal,                      // col E: Tebal
      parsed.kualitas,                   // col F: Kualitas
      parsed.skuHash,                    // col G: SKU#
      payload.sku,                       // col H: SKU
      masterRef.itemName,                // col I: Item Name
      qty,                               // col J: Qty In
      '',                                // col K: Qty Out
      '',                                // col L: Customer
      payload.supplier   || '',          // col M: Supplier
      payload.merkKaca   || masterRef.merkKaca || '', // col N: Merk Kaca
      payload.noDO       || '',          // col O: No Dokumen
      payload.lotNumber  || '',          // col P: Lot Number
    ]);

    // ── Update sheet Inventory ───────────────────────────────────────
    _updateInventory(ss, payload.sku, qty, 0, masterRef);

    return { success: true, message: `✅ Inbound ${qty} unit [${payload.sku}] berhasil dicatat.` };

  } catch (e) {
    return _err('Error sistem: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════════════
// §6  TRANSAKSI OUTBOUND
// ══════════════════════════════════════════════════════════════════════

/**
 * submitOutbound(payload)
 * ─────────────────────────────────────────────────────────────────────
 * Proses transaksi barang keluar.
 *
 * GUARD: Sistem menolak transaksi jika Current Inventory < Qty Out.
 *
 * Flow:
 *  1. Validasi input.
 *  2. Cek stok (Current Inventory dari sheet Inventory).
 *  3. Append baris ke sheet In-Out.
 *  4. Update baris SKU di sheet Inventory.
 *
 * payload: {
 *   date, sku, merkKaca, qtyOut,
 *   customer, lotNumber, noDO
 * }
 */
function submitOutbound(payload) {
  try {
    // ── Validasi wajib ──────────────────────────────────────────────
    if (!payload.sku)            return _err('SKU wajib dipilih.');
    if (!payload.qtyOut)         return _err('Qty Keluar wajib diisi.');
    const qty = Number(payload.qtyOut);
    if (isNaN(qty) || qty <= 0)  return _err('Qty Keluar harus angka positif.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── GUARD: Cek stok mencukupi ───────────────────────────────────
    // _getInventoryRow() membaca sheet Inventory SEKALI dan mencari baris SKU.
    const invRow = _getInventoryRow(ss, payload.sku);
    const C      = _idx(HDR.INVENTORY);
    const currentStock = invRow ? (Number(invRow.values[C['Current Inventory']]) || 0) : 0;

    if (currentStock < qty) {
      return _err(
        `⚠️ Stok tidak mencukupi!\n` +
        `Current Inventory: ${currentStock} unit\n` +
        `Qty Out diminta  : ${qty} unit`
      );
    }

    // ── Ambil referensi Masterdata ───────────────────────────────────
    const masterRef = _getMasterRef(ss, payload.sku);
    const parsed    = _parseSKU(payload.sku);

    // ── Tulis ke sheet In-Out ────────────────────────────────────────
    const inOutSh = _getSheet(SH.IN_OUT, ss);
    inOutSh.appendRow([
      '',                                // col A: Inbound Date (kosong untuk outbound)
      _fmtDate(payload.date),            // col B: Date
      'Outbound',                        // col C: Status
      parsed.jenis,                      // col D: Jenis
      parsed.tebal,                      // col E: Tebal
      parsed.kualitas,                   // col F: Kualitas
      parsed.skuHash,                    // col G: SKU#
      payload.sku,                       // col H: SKU
      masterRef.itemName,                // col I: Item Name
      '',                                // col J: Qty In
      qty,                               // col K: Qty Out
      payload.customer  || '',           // col L: Customer
      '',                                // col M: Supplier
      payload.merkKaca  || masterRef.merkKaca || '', // col N: Merk Kaca
      payload.noDO      || '',           // col O: No Dokumen
      payload.lotNumber || '',           // col P: Lot Number
    ]);

    // ── Update sheet Inventory ───────────────────────────────────────
    _updateInventory(ss, payload.sku, 0, qty, masterRef);

    return { success: true, message: `✅ Outbound ${qty} unit [${payload.sku}] berhasil dicatat.` };

  } catch (e) {
    return _err('Error sistem: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════════════
// §7  DATA UNTUK UI — Dashboard, Inventory, Lot Number Details, History
// ══════════════════════════════════════════════════════════════════════

/**
 * getDashboardData()
 * ─────────────────────────────────────────────────────────────────────
 * Ambil data Dashboard: scorecard + tabel inventory.
 *
 * EFISIENSI:
 *  • Baca sheet Inventory SEKALI.
 *  • Baca sheet Expired SEKALI → bangun map { SKU: totalExpiredQty }.
 *  • Semua kalkulasi di memori JavaScript.
 *
 * Return: { scorecards, tableData, categories, abcList, skuList }
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Baca Inventory (1x) ─────────────────────────────────────────────
    const invSh   = _getSheet(SH.INVENTORY, ss);
    const invRows = _readAllRows(invSh, HDR.INVENTORY.length);
    const C       = _idx(HDR.INVENTORY);

    // Baca Expired (1x) dan bangun map SKU → totalExpiredQty ──────────
    const expMap = _buildExpiredMap(ss);

    // Akumulasi scorecard
    let totalInv      = 0;
    let totalExpired  = 0;
    let totalBS       = 0;
    let totalStockout = 0;

    const catSet  = new Set();
    const abcSet  = new Set();
    const tableData = [];

    invRows.forEach(row => {
      const sku      = _str(row[C['SKU']]);
      const kategori = _str(row[C['Kategori']]);
      const itemName = _str(row[C['Item Name']]);
      const totalIn  = _num(row[C['Total In']]);
      const totalOut = _num(row[C['Total Out']]);
      const current  = _num(row[C['Current Inventory']]);
      const ss_val   = _num(row[C['SS']]);
      const abc      = _str(row[C['A-B-C']]);

      if (!sku) return; // skip baris kosong

      // Expired qty dari sheet Expired (sudah di-map)
      const expQty   = expMap[sku] || 0;
      const pctAvail = ss_val > 0 ? current / ss_val : null;
      const pctExp   = current > 0 ? expQty / current : null;
      const isBS     = itemName.toLowerCase().includes('(bs)');
      const isStockout = current === 0;

      // Scorecard akumulasi
      totalInv     += current;
      totalExpired += expQty;
      if (isBS)       totalBS       += current;
      if (isStockout) totalStockout += 1;

      catSet.add(kategori);
      if (abc) abcSet.add(abc);

      tableData.push({
        sku, kategori, itemName,
        totalIn, totalOut, current,
        pctAvail, expQty, pctExp,
        ss: ss_val, abc, isStockout,
      });
    });

    // Sort SKU ascending ───────────────────────────────────────────────
    tableData.sort((a, b) => a.sku.localeCompare(b.sku));

    return {
      scorecards: { totalInv, totalExpired, totalBS, totalStockout },
      tableData,
      categories : [...catSet].filter(Boolean).sort(),
      abcList    : [...abcSet].filter(Boolean).sort(),
      skuList    : tableData.map(r => r.sku),
    };

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * getInventoryData()
 * ─────────────────────────────────────────────────────────────────────
 * Data untuk halaman Inventory.
 * Filter opsional: filterSKU (string contains) & filterKategori (exact).
 */
function getInventoryData(filterSKU, filterKategori) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const invSh   = _getSheet(SH.INVENTORY, ss);
    const rows    = _readAllRows(invSh, HDR.INVENTORY.length);
    const C       = _idx(HDR.INVENTORY);

    let data = rows
      .filter(row => _str(row[C['SKU']]))
      .map(row => ({
        sku      : _str(row[C['SKU']]),
        kategori : _str(row[C['Kategori']]),
        itemName : _str(row[C['Item Name']]),
        totalIn  : _num(row[C['Total In']]),
        totalOut : _num(row[C['Total Out']]),
        current  : _num(row[C['Current Inventory']]),
      }));

    if (filterSKU)      data = data.filter(r => r.sku.toUpperCase().includes(filterSKU.toUpperCase()));
    if (filterKategori) data = data.filter(r => r.kategori === filterKategori);

    data.sort((a, b) => a.sku.localeCompare(b.sku));
    return { tableData: data };

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * getLotDetails(filters)
 * ─────────────────────────────────────────────────────────────────────
 * Data untuk halaman Lot Number Details.
 * Membaca sheet "Expired Product & BS" yang sudah diisi oleh ArrayFormula.
 *
 * filters: {
 *   dateFrom  : 'YYYY-MM-DD'  — Inbound Date >=
 *   dateTo    : 'YYYY-MM-DD'  — Inbound Date <=
 *   sku       : string        — exact match
 *   kategori  : string        — (requires cross-ref ke Inventory)
 *   expStatus : string        — 'Merah'|'Kuning'|'Hijau'|''
 * }
 *
 * Return: { tableData: [...] }
 */
function getLotDetails(filters) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const expSh = _getSheet(SH.EXPIRED, ss);
    const rows  = _readAllRows(expSh, HDR.EXPIRED.length);
    const C     = _idx(HDR.EXPIRED);

    // Baca Inventory untuk lookup Kategori per SKU (1x) ──────────────
    const catBySKU = _buildCategoryMap(ss);

    const from = filters.dateFrom ? new Date(filters.dateFrom) : null;
    const to   = filters.dateTo   ? new Date(filters.dateTo)   : null;
    if (to) to.setHours(23, 59, 59);

    let data = rows
      .filter(row => _str(row[C['SKU']]))   // skip empty
      .map(row => {
        const ibDate    = row[C['Inbound Date']];
        const ibDateObj = ibDate ? new Date(ibDate) : null;
        const sku       = _str(row[C['SKU']]);
        return {
          inboundDate : ibDateObj ? _fmtDateDisplay(ibDateObj) : '',
          inboundDateRaw: ibDateObj,
          sku,
          supplier    : _str(row[C['Supplier']]),
          merkKaca    : _str(row[C['Merk Kaca']]),
          lotNumber   : _str(row[C['Lot Number']]),
          umurBulan   : _num(row[C['Umur Kaca (Bulan)']]),
          expQty      : _num(row[C['Exp Qty']]),
          qty         : _num(row[C['Qty']]),
          expStatus   : _str(row[C['Expired Status']]),
          kategori    : catBySKU[sku] || '',
        };
      });

    // ── Apply filters ─────────────────────────────────────────────────
    if (from) data = data.filter(r => r.inboundDateRaw && r.inboundDateRaw >= from);
    if (to)   data = data.filter(r => r.inboundDateRaw && r.inboundDateRaw <= to);
    if (filters.sku)       data = data.filter(r => r.sku === filters.sku);
    if (filters.kategori)  data = data.filter(r => r.kategori === filters.kategori);
    if (filters.expStatus) data = data.filter(r => r.expStatus === filters.expStatus);

    // Bersihkan field internal sebelum kirim ke client
    data.forEach(r => delete r.inboundDateRaw);
    data.sort((a, b) => a.sku.localeCompare(b.sku));

    return { tableData: data };

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * getInOutHistory(limit)
 * ─────────────────────────────────────────────────────────────────────
 * Ambil N baris terakhir dari sheet In-Out (default 200).
 * Diurutkan terbaru di atas.
 */
function getInOutHistory(limit) {
  try {
    const sh       = _getSheet(SH.IN_OUT);
    const maxRows  = Number(limit) || 200;
    const lastRow  = sh.getLastRow();
    if (lastRow < 2) return { rows: [] };

    const startRow = Math.max(2, lastRow - maxRows + 1);
    const numRows  = lastRow - startRow + 1;
    const C        = _idx(HDR.IN_OUT);

    const raw = sh.getRange(startRow, 1, numRows, HDR.IN_OUT.length).getValues();
    const rows = raw.map(row => ({
      inboundDate: _fmtDateDisplay(row[C['Inbound Date']]),
      date       : _fmtDateDisplay(row[C['Date']]),
      status     : _str(row[C['Status']]),
      sku        : _str(row[C['SKU']]),
      itemName   : _str(row[C['Item Name']]),
      qtyIn      : row[C['Qty In']]  !== '' ? _num(row[C['Qty In']])  : '',
      qtyOut     : row[C['Qty Out']] !== '' ? _num(row[C['Qty Out']]) : '',
      customer   : _str(row[C['Customer']]),
      supplier   : _str(row[C['Supplier']]),
      merkKaca   : _str(row[C['Merk Kaca']]),
      lotNumber  : _str(row[C['Lot Number']]),
      noDokumen  : _str(row[C['No Dokumen']]),
    })).reverse();  // terbaru di atas

    return { rows };

  } catch (e) {
    return { error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════════════
// §8  HELPER INTERNAL — Inventory Update
// ══════════════════════════════════════════════════════════════════════

/**
 * _updateInventory(ss, sku, qtyIn, qtyOut, masterRef)
 * ─────────────────────────────────────────────────────────────────────
 * FUNGSI DINONAKTIFKAN SESUAI PERMINTAAN:
 * Code.gs tidak lagi menulis/update baris ke sheet Inventory.
 * Perhitungan Inventory sepenuhnya dikelola menggunakan formula di Spreadsheet.
 */
function _updateInventory(ss, sku, qtyIn, qtyOut, masterRef) {
  // Return langsung; tidak melakukan penulisan apa-apa ke sheet Inventory.
  return;
}

// ══════════════════════════════════════════════════════════════════════
// §9  HELPER INTERNAL — Data Maps
// ══════════════════════════════════════════════════════════════════════

/**
 * _getMasterRef(ss, sku)
 * ─────────────────────────────────────────────────────────────────────
 * Baca Masterdata SEKALI dan kembalikan objek referensi untuk SKU.
 * Return: { itemName, merkKaca, jenis, tebal, kualitas, ss, abc, kategori }
 */
function _getMasterRef(ss, sku) {
  const sh   = _getSheet(SH.MASTER, ss);
  const rows = _readAllRows(sh, HDR.MASTER.length);
  const C    = _idx(HDR.MASTER);

  for (const row of rows) {
    if (_str(row[C['SKU']]) === sku) {
      return {
        itemName : _str(row[C['Item Name']]),
        merkKaca : _str(row[C['Merk Kaca']]),
        jenis    : _str(row[C['Jenis']]),
        tebal    : _str(row[C['Tebal']]),
        kualitas : _str(row[C['Kualitas']]),
        ss       : _num(row[C['SS']]),
        abc      : _str(row[C['A-B-C']]),
        kategori : _str(row[C['Kategori']]),
      };
    }
  }
  // Fallback jika SKU tidak di Masterdata
  return { itemName: sku, merkKaca: '', jenis: '', tebal: '', kualitas: '', ss: 0, abc: '', kategori: '' };
}

/**
 * _buildExpiredMap(ss)
 * ─────────────────────────────────────────────────────────────────────
 * Baca sheet Expired SEKALI → kembalikan map { SKU: totalQty }.
 * Digunakan untuk kalkulasi "Expired Product" di sheet Inventory.
 */
function _buildExpiredMap(ss) {
  const sh = ss.getSheetByName(SH.EXPIRED);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;

  const C    = _idx(HDR.EXPIRED);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HDR.EXPIRED.length).getValues();

  rows.forEach(row => {
    const sku = _str(row[C['SKU']]);
    if (!sku) return;
    map[sku] = (map[sku] || 0) + _num(row[C['Qty']]);
  });
  return map;
}

/**
 * _buildCategoryMap(ss)
 * Baca sheet Inventory → map { SKU: Kategori }
 */
function _buildCategoryMap(ss) {
  const sh  = ss.getSheetByName(SH.INVENTORY);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;

  const C    = _idx(HDR.INVENTORY);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HDR.INVENTORY.length).getValues();
  rows.forEach(row => {
    const sku = _str(row[C['SKU']]);
    if (sku) map[sku] = _str(row[C['Kategori']]);
  });
  return map;
}

// ══════════════════════════════════════════════════════════════════════
// §10  HELPER UTILITY
// ══════════════════════════════════════════════════════════════════════

/** Ambil sheet; lempar error deskriptif jika tidak ada */
function _getSheet(name, ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sh = spreadsheet.getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" tidak ditemukan. Jalankan setupDatabase() terlebih dahulu.`);
  return sh;
}

/** Baca semua baris data (mulai baris 2) sebagai array 2D */
function _readAllRows(sh, colCount) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, colCount).getValues();
}

/** Build map header → index (0-based) */
function _idx(headers) {
  const m = {};
  headers.forEach((h, i) => { m[h] = i; });
  return m;
}

/**
 * _parseSKU(sku)
 * ─────────────────────────────────────────────────────────────────────
 * Parse SKU format: "P 5 72x48MIRR"
 *   Token[0] = Jenis   → "P"
 *   Token[1] = Tebal   → "5"
 *   Token[2] = chunk   → "72x48MIRR"
 *     chunk = ukuran (72x48) + kualitas (MIRR)
 *   SKU#     → "P 5 72x48"  (tanpa suffix kualitas)
 *   Kualitas → "MIRR"
 */
function _parseSKU(sku) {
  const parts    = sku.trim().split(/\s+/);
  const jenis    = parts[0] || '';
  const tebal    = parts[1] || '';
  const chunk    = parts[2] || '';

  // Pisahkan ukuran (digit x digit) dari kualitas (huruf di belakang)
  const m        = chunk.match(/^(\d[\d.]*x[\d.]+)(.*)$/i);
  const ukuran   = m ? m[1] : chunk;
  const kualitas = m ? m[2] : '';
  const skuHash  = [jenis, tebal, ukuran].filter(Boolean).join(' ');

  return { jenis, tebal, kualitas, skuHash };
}

/** Format Date ke string yyyy-MM-dd untuk disimpan di Sheets */
function _fmtDate(d) {
  if (!d) return '';
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt)) return String(d);
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Format Date ke dd/MM/yyyy untuk ditampilkan di UI */
function _fmtDateDisplay(d) {
  if (!d || d === '') return '';
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt)) return String(d);
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/** Konversi ke string, trim, handle null/undefined */
function _str(v) { return String(v == null ? '' : v).trim(); }

/** Konversi ke number, default 0 */
function _num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }

/** Standard error response */
function _err(msg) { return { success: false, message: msg }; }
