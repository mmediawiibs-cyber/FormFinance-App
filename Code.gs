/**
 * FORM FINANCE STIT — GOOGLE APPS SCRIPT (PP, PUM, PJUM, VENDOR & SEO)
 *
 * Fitur:
 * 1. Mendukung penuh operasi CRUD (Insert, Update/Upsert di Nomor yang Sama, Delete)
 *    untuk Sheet: "Data PP", "Data PUM", dan "Data PJUM" (Table range bernama Tabel PP/Tabel PUM).
 * 2. Mendukung upload file Lampiran ke Google Drive untuk PP & PUM.
 * 3. Mempertahankan endpoint Dashboard Keuangan SEO (/exec dan ?action=getDataPPForSEO).
 * 4. Mendukung Cetak PDF dinamis untuk "Form PP", "Form PUM", dan "Form PJUM".
 */

var SPREADSHEET_ID = "1qdxGHKNCSkuhEccI_W63ShWv-bFuKVsTK09nrYdqdnw";
var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// ============================================================
// HELPER UMUM
// ============================================================

function getBulanRomawi(tanggalStr) {
  var d = tanggalStr ? new Date(tanggalStr) : new Date();
  if (isNaN(d.getTime())) d = new Date();
  var romawi = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"];
  return romawi[d.getMonth()] || "I";
}

function getSheetOrCreate(sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function normalizeNominal(rawValue) {
  if (typeof rawValue === "number") return rawValue;
  var raw = String(rawValue == null ? "0" : rawValue);
  var isNegative = raw.indexOf("-") !== -1;
  var cleaned = raw.replace(/[^0-9]/g, "");
  var num = cleaned === "" ? 0 : Number(cleaned);
  return isNegative ? -num : num;
}

function createJsonResponse_(payload, prefix) {
  var json = JSON.stringify(payload);
  var callback = String(prefix || "").trim();

  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, message: "Parameter prefix tidak valid." })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function findRowIndexByNomor(sheet, nomor) {
  if (!sheet) return -1;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(nomor)) return i + 1;
  }
  return -1;
}

function deleteRowsByNomor(sheetName, nomor) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  var rows = sheet.getDataRange().getValues();
  var deleted = false;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (Number(rows[i][0]) === Number(nomor)) {
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }
  return deleted;
}

function uploadLampiranIfAny_(data) {
  var lampiranLink = data.lampiranText || data.lampiran || "Terlampir";
  if (data.fileLampiran && data.fileLampiran.data) {
    try {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(data.fileLampiran.data),
        data.fileLampiran.mimeType,
        data.fileLampiran.name
      );
      var file = DriveApp.createFile(blob);
      lampiranLink = file.getUrl();
    } catch (errFile) {
      lampiranLink = "Gagal upload: " + errFile.toString();
    }
  }
  return lampiranLink;
}

// ============================================================
// DATA VENDOR
// ============================================================

function ensureVendorSeedData() {
  var sheet = getSheetOrCreate("Vendor");
  var expectedHeader = ["NO", "PENERIMA", "SEBAGAI", "A/N PENERIMA", "BANK", "NOMOR REKENING"];
  var values = sheet.getDataRange().getValues();

  if (!values.length || String(values[0][0] || "").trim() !== "NO") {
    sheet.clear();
    sheet.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader]);
  }
  return sheet;
}

function getVendorNextNo(sheet) {
  if (!sheet) return 1;
  var rows = sheet.getDataRange().getValues();
  var maxNo = 0;
  for (var i = 1; i < rows.length; i++) {
    var val = Number(rows[i][0]);
    if (!isNaN(val) && val > maxNo) maxNo = val;
  }
  return maxNo + 1;
}

function vendorRowData(data, nomorOverride) {
  var noValue = Number(nomorOverride);
  if (isNaN(noValue)) noValue = 1;
  return [
    noValue,
    data.nama || "",
    data.sebagai || "",
    data.anPenerima || "",
    data.bank || "",
    data.noRekening || ""
  ];
}

function findVendorRowIndexByName(sheet, nama) {
  if (!sheet || !nama) return -1;
  var rows = sheet.getDataRange().getValues();
  var target = String(nama).trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || "").trim().toLowerCase() === target) {
      return i + 1;
    }
  }
  return -1;
}

// ============================================================
// BUILDER BARIS: DATA PP, DATA PUM, DATA PJUM
// ============================================================

// Nama lengkap divisi yang disimpan di kolom DIVISI.
// Singkatan hanya dipakai untuk membentuk nomor surat PP/PUM/PJUM.
var DIVISI_FULL_NAMES_ = {
  "MD": "Media Development",
  "P-MPI": "Prodi Manajemen Pendidikan Islam",
  "IT": "Information And Technology",
  "LPM": "Lembaga Penjamin Mutu",
  "LPPM": "Lembaga Pendidikan & Pengabdian Masy.",
  "BAAU": "Biro Administrasi Akademik & Umum",
  "PMB": "Penerimaan Mahasiswa Baru",
  "JRL": "Jurnal",
  "GA": "General Affairs"
};

function getNamaDivisiLengkap_(kode, nilaiTersimpan) {
  var nilai = String(nilaiTersimpan || "").trim();
  var kodeUpper = String(kode || "").trim().toUpperCase();
  var nilaiUpper = nilai.toUpperCase();

  if (nilai && DIVISI_FULL_NAMES_[nilaiUpper]) {
    return DIVISI_FULL_NAMES_[nilaiUpper];
  }
  if (nilai) return nilai; // Sudah berupa nama lengkap.
  return DIVISI_FULL_NAMES_[kodeUpper] || kode || "";
}

function getKodeDivisiDariNomor_(nomorDokumen, namaDivisi) {
  var nomor = String(nomorDokumen || "");
  var match = nomor.match(/(?:PP|PUM|PJUM)\.STIT\.([^\s/]+)/i);
  if (match && match[1]) return String(match[1]).toUpperCase();

  var nama = String(namaDivisi || "").trim().toLowerCase();
  var keys = Object.keys(DIVISI_FULL_NAMES_);
  for (var i = 0; i < keys.length; i++) {
    if (DIVISI_FULL_NAMES_[keys[i]].toLowerCase() === nama) return keys[i];
  }
  if (DIVISI_FULL_NAMES_[String(namaDivisi || "").trim().toUpperCase()]) {
    return String(namaDivisi).trim().toUpperCase();
  }
  return String(namaDivisi || "").trim();
}

// Jalankan satu kali secara manual bila data lama di kolom DIVISI masih berisi singkatan.
// Fungsi mencari kolom berdasarkan header, sehingga tetap aman setelah kolom B dan I dihapus.
function migrasiNamaDivisiPP_PUM() {
  var namaSheet = ["Data PP", "Data PUM"];
  var totalDiubah = 0;

  namaSheet.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    var indexDivisi = headers.findIndex(function(h) {
      return String(h || "").trim().toUpperCase() === "DIVISI";
    });
    if (indexDivisi < 0) return;

    var range = sheet.getRange(2, indexDivisi + 1, sheet.getLastRow() - 1, 1);
    var values = range.getValues();
    var changed = false;
    for (var i = 0; i < values.length; i++) {
      var original = String(values[i][0] || "").trim();
      var fullName = getNamaDivisiLengkap_(original, original);
      if (original && fullName !== original) {
        values[i][0] = fullName;
        totalDiubah++;
        changed = true;
      }
    }
    if (changed) range.setValues(values);
  });

  Logger.log("Migrasi selesai. Jumlah nilai DIVISI yang diubah: " + totalDiubah);
  return totalDiubah;
}

// Susunan baru 23 kolom Sheet "Data PP":
// NO, NOMOR PP, NAMA PEMBUAT, DIVISI (nama lengkap), WAKA TERKAIT,
// TANGGAL PP, TANGGAL PEMENUHAN, URAIAN I, URAIAN II, URAIAN III,
// KEPADA, SEBAGAI, A/N PENERIMA, BANK, NO. REKENING, NOMINAL,
// LAMPIRAN, TANGGAL ENTRY, STATUS, TANGGAL TF, JUMLAH TF, TANGGAL, KETERANGAN.
function buildRowDataPP_(sheet, data) {
  var nominalValue = normalizeNominal(data.nominal);
  var lampiranLink = uploadLampiranIfAny_(data);
  var divisiNama = getNamaDivisiLengkap_(data.divisi, data.divisiNama);

  return [
    Number(data.nomor || 0),
    data.nomorPp || "",
    data.namaPembuat || "",
    divisiNama,
    data.wakaTerkait || "",
    data.tanggalPp || "",
    data.tanggalPemenuhan || "",
    data.uraian1 || "",
    data.uraian2 || "",
    data.uraian3 || "",
    data.kepada || "",
    data.sebagai || "",
    data.anPenerima || "",
    data.bank || "",
    data.noRekening || "",
    nominalValue,
    lampiranLink,
    data.tanggalEntry || "",
    data.status || "NOT PAID",
    data.tanggalTf || "",
    data.jumlahTf ? normalizeNominal(data.jumlahTf) : "",
    data.tanggalRekap || "",
    data.keterangan || ""
  ];
}

// Susunan baru 23 kolom Sheet "Data PUM" dengan urutan yang sama seperti PP.
function buildRowDataPUM_(data) {
  var nominalValue = normalizeNominal(data.nominal);
  var lampiranLink = uploadLampiranIfAny_(data);
  var divisiNama = getNamaDivisiLengkap_(data.divisi, data.divisiNama);

  return [
    Number(data.nomor || 0),
    data.nomorPum || "",
    data.namaPembuat || "",
    divisiNama,
    data.wakaTerkait || "",
    data.tanggalPp || "",
    data.tanggalPemenuhan || "",
    data.uraian1 || "",
    data.uraian2 || "",
    data.uraian3 || "",
    data.kepada || "",
    data.sebagai || "",
    data.anPenerima || "",
    data.bank || "",
    data.noRekening || "",
    nominalValue,
    lampiranLink,
    data.tanggalEntry || "",
    data.status || "NOT PAID",
    data.tanggalTf || "",
    data.jumlahTf ? normalizeNominal(data.jumlahTf) : "",
    data.tanggalRekap || "",
    data.keterangan || ""
  ];
}

// Susunan 13 kolom Data PJUM. Kolom B menggantikan kolom PP (PUM)
// dengan nomor dokumen PUM agar referensi administrasi terlihat jelas.
// NO (PUM) pada kolom A tetap menjadi primary key internal relasi PUM-PJUM.
function buildRowDataPJUM_(data) {
  var uangMukaVal = normalizeNominal(data.uangMuka);
  var terpakaiVal = normalizeNominal(data.terpakai);
  var selisihVal = uangMukaVal - terpakaiVal;

  return [
    Number(data.nomor || 0),               // A. NO (PUM) — primary key internal
    data.nomorPum || "",                  // B. NOMOR PUM (PUM) — referensi dokumen
    data.nomorPjum || "",                 // C. NOMOR PJUM
    data.namaPembuat || "",               // D. NAMA PEMBUAT (PUM)
    getNamaDivisiLengkap_(data.divisi, data.divisiNama), // E. DIVISI (PUM) — nama lengkap
    data.wakaTerkait || "",               // F. WAKA TERKAIT (PUM)
    data.tanggalPp || "",                 // G. TANGGAL PP (PUM)
    data.tanggalPemenuhan || "",          // H. TANGGAL PEMENUHAN (PUM)
    data.uraianPjum || "",                // I. URAIAN (Input PJUM)
    uangMukaVal,                          // J. UANG MUKA (PUM)
    terpakaiVal,                          // K. TERPAKAI (Input PJUM)
    selisihVal,                           // L. KURANG/LEBIH
    data.keterangan || (selisihVal >= 0 ? "LEBIH" : "KURANG") // M. KETERANGAN
  ];
}

// ============================================================
// PARSER BARIS FLEKSIBEL (ANTI GESER KOLOM)
// ============================================================

function parseSheetPP_(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var list = [];

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[0] || "").trim() === "" && String(r[1] || "").trim() === "") continue;

    var nomorPp = r[1] || "";
    var divisiCode = getKodeDivisiDariNomor_(nomorPp, r[3]);
    list.push({
      nomor: Number(r[0]) || i,
      nomorPp: nomorPp,
      namaPembuat: r[2] || "",
      divisi: divisiCode,
      divisiNama: getNamaDivisiLengkap_(divisiCode, r[3]),
      wakaTerkait: r[4] || "",
      tanggalPp: r[5] || "",
      tanggalPemenuhan: r[6] || "",
      uraian1: r[7] || "",
      uraian2: r[8] || "",
      uraian3: r[9] || "",
      kepada: r[10] || "",
      sebagai: r[11] || "",
      anPenerima: r[12] || "",
      bank: r[13] || "",
      noRekening: r[14] || "",
      nominal: r[15] || 0,
      lampiran: r[16] || "-",
      tanggalEntry: r[17] || "",
      status: r[18] || "NOT PAID",
      tanggalTf: r[19] || "",
      jumlahTf: r[20] || "",
      tanggalRekap: r[21] || "",
      keterangan: r[22] || ""
    });
  }
  return list;
}

function parseSheetPUM_(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var list = [];

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[0] || "").trim() === "" && String(r[1] || "").trim() === "") continue;

    var nomorPum = r[1] || "";
    var divisiCode = getKodeDivisiDariNomor_(nomorPum, r[3]);
    list.push({
      nomor: Number(r[0]) || i,
      nomorPum: nomorPum,
      namaPembuat: r[2] || "",
      divisi: divisiCode,
      divisiNama: getNamaDivisiLengkap_(divisiCode, r[3]),
      wakaTerkait: r[4] || "",
      tanggalPp: r[5] || "",
      tanggalPemenuhan: r[6] || "",
      uraian1: r[7] || "",
      uraian2: r[8] || "",
      uraian3: r[9] || "",
      kepada: r[10] || "",
      sebagai: r[11] || "",
      anPenerima: r[12] || "",
      bank: r[13] || "",
      noRekening: r[14] || "",
      nominal: r[15] || 0,
      lampiran: r[16] || "-",
      tanggalEntry: r[17] || "",
      status: r[18] || "NOT PAID",
      tanggalTf: r[19] || "",
      jumlahTf: r[20] || "",
      tanggalRekap: r[21] || "",
      keterangan: r[22] || ""
    });
  }
  return list;
}

function parseSheetPJUM_(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[0] === "" && r[2] === "") continue;
    var nomorPum = r[1] || "";
    var nomorPjum = r[2] || "";
    var divisiCode = getKodeDivisiDariNomor_(nomorPjum, r[4]);
    list.push({
      nomor: Number(r[0]) || i,
      nomorPum: nomorPum,
      nomorPjum: nomorPjum,
      namaPembuat: r[3] || "",
      divisi: divisiCode,
      divisiNama: getNamaDivisiLengkap_(divisiCode, r[4]),
      wakaTerkait: r[5] || "",
      tanggalPp: r[6] || "",
      tanggalPemenuhan: r[7] || "",
      uraianPjum: r[8] || "",
      uangMuka: r[9] || 0,
      terpakai: r[10] || 0,
      kurangLebih: r[11] || 0,
      keterangan: r[12] || ""
    });
  }
  return list;
}

// ============================================================
// MIGRASI REFERENSI NOMOR PUM KE DATA PJUM
// Jalankan SATU KALI setelah kolom B lama (PP) dihapus dan
// kolom B baru dibuat dengan header "NOMOR PUM".
// Fungsi mengisi nomor dokumen dari Data PUM berdasarkan NO (PUM).
// ============================================================
function migrasiNomorPUMKePJUM() {
  var sheetPUM = ss.getSheetByName("Data PUM");
  var sheetPJUM = ss.getSheetByName("Data PJUM");
  if (!sheetPUM) throw new Error('Sheet "Data PUM" tidak ditemukan.');
  if (!sheetPJUM) throw new Error('Sheet "Data PJUM" tidak ditemukan.');

  var pjumLastRow = sheetPJUM.getLastRow();
  var pjumLastCol = sheetPJUM.getLastColumn();
  if (pjumLastRow < 1 || pjumLastCol < 2) {
    throw new Error('Struktur Data PJUM belum siap. Pastikan kolom B "NOMOR PUM" sudah dibuat.');
  }

  var headerB = String(sheetPJUM.getRange(1, 2).getDisplayValue() || "").trim().toUpperCase();
  if (headerB !== "NOMOR PUM") {
    throw new Error('Header kolom B Data PJUM harus "NOMOR PUM". Hapus kolom B lama, sisipkan kolom B baru, lalu beri header NOMOR PUM.');
  }

  var pumRows = sheetPUM.getDataRange().getDisplayValues();
  var nomorPumByNo = {};
  for (var i = 1; i < pumRows.length; i++) {
    var noPum = Number(pumRows[i][0]);
    var nomorDokumenPum = String(pumRows[i][1] || "").trim();
    if (noPum > 0 && nomorDokumenPum) nomorPumByNo[noPum] = nomorDokumenPum;
  }

  if (pjumLastRow < 2) return "Data PJUM belum memiliki baris data untuk dimigrasikan.";

  var noValues = sheetPJUM.getRange(2, 1, pjumLastRow - 1, 1).getValues();
  var nomorPumRange = sheetPJUM.getRange(2, 2, pjumLastRow - 1, 1);
  var existingValues = nomorPumRange.getValues();
  var existingFormulas = nomorPumRange.getFormulas();
  var updated = 0;

  for (var r = 0; r < noValues.length; r++) {
    // Formula VLOOKUP yang sudah dipasang di kolom B tidak ditimpa.
    if (existingFormulas[r][0]) continue;
    var no = Number(noValues[r][0]);
    var docNo = nomorPumByNo[no] || "";
    if (docNo && String(existingValues[r][0] || "").trim() !== docNo) {
      nomorPumRange.getCell(r + 1, 1).setValue(docNo);
      updated++;
    }
  }

  SpreadsheetApp.flush();
  var message = "Migrasi selesai. " + updated + " nomor PUM berhasil diisi pada Data PJUM.";
  Logger.log(message);
  return message;
}

function getDataPPForSEO_() {
  var sheet = ss.getSheetByName("Data PP");
  if (!sheet) {
    return { success: false, message: 'Sheet "Data PP" tidak ditemukan.', sheetName: "Data PP", headers: [], rows: [], rowCount: 0 };
  }
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    return { success: true, sheetName: "Data PP", headers: [], rows: [], rowCount: 0, updatedAt: new Date().toISOString() };
  }
  var matrix = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var headers = (matrix[0] || []).map(function (value) { return String(value == null ? "" : value).trim(); });
  var rows = matrix.slice(1).filter(function (row) {
    return row.some(function (cell) { return String(cell == null ? "" : cell).trim() !== ""; });
  });
  return { success: true, sheetName: "Data PP", headers: headers, rows: rows, rowCount: rows.length, updatedAt: new Date().toISOString() };
}

// ============================================================
// POST API — CRUD PP, PUM, PJUM & VENDOR
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Request POST kosong atau tidak valid.");
    }

    var data = JSON.parse(e.postData.contents);
    var action = String(data.action || "").trim();

    // --------------------------------------------------------
    // 1. OPERASI VENDOR
    // --------------------------------------------------------
    if (action === "saveVendor") {
      var vendorSheet = ensureVendorSeedData();
      var vendorName = String(data.nama || "").trim();
      if (!vendorName) return createJsonResponse_({ status: "error", message: "Nama vendor wajib diisi." });
      if (findVendorRowIndexByName(vendorSheet, vendorName) > -1) {
        return createJsonResponse_({ status: "error", message: "Nama vendor sudah ada." });
      }
      vendorSheet.appendRow(vendorRowData(data, getVendorNextNo(vendorSheet)));
      return createJsonResponse_({ status: "success", message: "Vendor baru berhasil ditambahkan." });
    }

    if (action === "updateVendor") {
      var vendorSheetUpdate = ensureVendorSeedData();
      var originalName = String(data.originalNama || data.nama || "").trim();
      var vendorRowIndex = findVendorRowIndexByName(vendorSheetUpdate, originalName);
      if (vendorRowIndex === -1) vendorRowIndex = findVendorRowIndexByName(vendorSheetUpdate, data.nama || "");
      if (vendorRowIndex === -1) return createJsonResponse_({ status: "error", message: "Vendor tidak ditemukan." });
      var currentNo = Number(vendorSheetUpdate.getRange(vendorRowIndex, 1).getValue()) || vendorRowIndex;
      vendorSheetUpdate.getRange(vendorRowIndex, 1, 1, 6).setValues([[
        currentNo, data.nama || "", data.sebagai || "", data.anPenerima || "", data.bank || "", data.noRekening || ""
      ]]);
      return createJsonResponse_({ status: "success", message: "Vendor berhasil diperbarui." });
    }

    if (action === "deleteVendor") {
      var vendorSheetDelete = ensureVendorSeedData();
      var deleteVendorRow = findVendorRowIndexByName(vendorSheetDelete, data.nama || "");
      if (deleteVendorRow === -1) return createJsonResponse_({ status: "error", message: "Vendor tidak ditemukan." });
      vendorSheetDelete.deleteRow(deleteVendorRow);
      return createJsonResponse_({ status: "success", message: "Vendor berhasil dihapus." });
    }

    // --------------------------------------------------------
    // 2. OPERASI PUM (Data PUM)
    // --------------------------------------------------------
    if (action === "savePUM") {
      var sheetPUM = getSheetOrCreate("Data PUM");
      var rowPUM = buildRowDataPUM_(data);
      var existingPumRow = findRowIndexByNomor(sheetPUM, data.nomor);
      if (existingPumRow > -1) {
        sheetPUM.getRange(existingPumRow, 1, 1, rowPUM.length).setValues([rowPUM]);
      } else {
        sheetPUM.appendRow(rowPUM);
      }
      return createJsonResponse_({ status: "success", message: "Data PUM berhasil disimpan!" });
    }

    if (action === "updatePUM") {
      var sheetPUMUpd = getSheetOrCreate("Data PUM");
      var rIdxPum = findRowIndexByNomor(sheetPUMUpd, data.nomor);
      if (rIdxPum > -1) {
        var updatedPumRow = buildRowDataPUM_(data);
        sheetPUMUpd.getRange(rIdxPum, 1, 1, updatedPumRow.length).setValues([updatedPumRow]);
        return createJsonResponse_({ status: "success", message: "Data PUM berhasil diperbarui!" });
      }
      return createJsonResponse_({ status: "error", message: "Data PUM tidak ditemukan." });
    }

    if (action === "deletePUM") {
      var delPum = deleteRowsByNomor("Data PUM", data.nomor);
      return createJsonResponse_({ status: delPum ? "success" : "error", message: delPum ? "Data PUM dihapus." : "Nomor PUM tidak ditemukan." });
    }

    // --------------------------------------------------------
    // 3. OPERASI PJUM (Data PJUM — UPSERT PADA NOMOR YANG SAMA)
    // --------------------------------------------------------
    if (action === "upsertPJUM" || action === "updatePJUM") {
      var sheetPJUM = getSheetOrCreate("Data PJUM");
      var rIdxPjum = findRowIndexByNomor(sheetPJUM, data.nomor);
      var rowPJUM = buildRowDataPJUM_(data);

      if (rIdxPjum > -1) {
        // Jika baris dengan NO (PUM) tersebut sudah ada, UPDATE di baris yang sama!
        // Cek apakah kolom di sheet memakai rumus VLOOKUP untuk kolom (PUM).
        // Struktur A:M: NOMOR PUM berada di B (referensi), sedangkan kolom input PJUM adalah C, I, K, L, M.
        var formulas = sheetPJUM.getRange(rIdxPjum, 1, 1, 13).getFormulas()[0];
        for (var col = 0; col < 13; col++) {
          // Jangan timpa sel jika di dalamnya sudah ada rumus VLOOKUP dari user, kecuali kolom input PJUM (9, 11, 12, 13)
          var isInputColumn = (col === 2 || col === 8 || col === 10 || col === 11 || col === 12);
          if (!formulas[col] || isInputColumn) {
            sheetPJUM.getRange(rIdxPjum, col + 1).setValue(rowPJUM[col]);
          }
        }
        return createJsonResponse_({ status: "success", message: "Data PJUM Nomor #" + data.nomor + " berhasil diperbarui pada baris yang sama!" });
      } else {
        // Jika belum ada di Sheet Data PJUM, tambahkan baris baru dengan nomor PUM tersebut
        sheetPJUM.appendRow(rowPJUM);
        return createJsonResponse_({ status: "success", message: "Data PJUM Nomor #" + data.nomor + " berhasil disimpan!" });
      }
    }

    if (action === "deletePJUM") {
      var delPjum = deleteRowsByNomor("Data PJUM", data.nomor);
      return createJsonResponse_({ status: delPjum ? "success" : "error", message: delPjum ? "Data PJUM dihapus." : "Data PJUM tidak ditemukan." });
    }

    // --------------------------------------------------------
    // 4. OPERASI DATA PP (Update, Delete, & Default Insert)
    // --------------------------------------------------------
    var sheetPP = getSheetOrCreate("Data PP");

    if (action === "updateData") {
      var rowPP = findRowIndexByNomor(sheetPP, data.nomor);
      if (rowPP > -1) {
        var updatedRowPP = buildRowDataPP_(sheetPP, data);
        sheetPP.getRange(rowPP, 1, 1, updatedRowPP.length).setValues([updatedRowPP]);
        return createJsonResponse_({ status: "success", message: "Data PP berhasil diperbarui!" });
      }
      return createJsonResponse_({ status: "error", message: "Data PP tidak ditemukan untuk diupdate." });
    }

    if (action === "deleteData") {
      var deletedA = deleteRowsByNomor("Data PP", data.nomor);
      return createJsonResponse_({ status: deletedA ? "success" : "error", message: deletedA ? "Data PP berhasil dihapus!" : "Data tidak ditemukan." });
    }

    // Default Action: Simpan PP Baru
    var nomorUrut = Number(data.nomor) || 1;
    if (!data.nomor) {
      var lastRow = sheetPP.getLastRow();
      if (lastRow > 1) {
        var lastNum = sheetPP.getRange(lastRow, 1).getValue();
        nomorUrut = isNaN(lastNum) || lastNum === "" ? lastRow : Number(lastNum) + 1;
      }
    }
    var strNoUrut = ("000" + nomorUrut).slice(-3);
    var tahun = data.tanggalPp ? new Date(data.tanggalPp).getFullYear() : new Date().getFullYear();
    var bulanRomawi = getBulanRomawi(data.tanggalPp);
    var nomorPp = data.nomorPp || (strNoUrut + " / PP.STIT." + data.divisi + " / " + bulanRomawi + " / " + tahun);

    data.nomor = nomorUrut;
    data.nomorPp = nomorPp;
    var rowDataNewPP = buildRowDataPP_(sheetPP, data);
    sheetPP.appendRow(rowDataNewPP);

    return createJsonResponse_({
      status: "success",
      message: "Data berhasil disimpan ke Data PP!",
      nomorUrut: nomorUrut,
      nomorPp: nomorPp
    });

  } catch (error) {
    return createJsonResponse_({ status: "error", message: error.toString() });
  }
}

// ============================================================
// GET API — SEO + DATA PP, PUM, PJUM, VENDOR & CETAK PDF
// ============================================================

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = String(params.action || "").trim();

  // 1. Cetak PDF Dinamis (Form PP, Form PUM, Form PJUM)
  if (action === "cetakPDF" || (params.nomor && !action)) {
    var formName = params.formName || "Form PP";
    var targetFormSheet = getSheetOrCreate(formName);
    targetFormSheet.getRange("A1").setValue(params.nomor);
    SpreadsheetApp.flush();
    var pdfUrl =
      "https://docs.google.com/spreadsheets/d/" +
      ss.getId() +
      "/export?exportFormat=pdf&format=pdf&size=7&fzr=true&portrait=true&fitw=true&gridlines=false&gid=" +
      targetFormSheet.getSheetId();
    return HtmlService.createHtmlOutput("<script>window.location.href='" + pdfUrl + "';</script>");
  }

  // 2. Endpoint SEO
  if (action === "getDataPPForSEO" || (!action && params.prefix)) {
    return createJsonResponse_(getDataPPForSEO_(), params.prefix);
  }

  // 3. Ambil Data PP, PUM, atau PJUM
  if (action === "getAllData") {
    var sheetName = params.sheetName || "Data PP";
    if (sheetName === "Data PUM") {
      return createJsonResponse_(parseSheetPUM_(getSheetOrCreate("Data PUM")), params.prefix);
    }
    if (sheetName === "Data PJUM") {
      return createJsonResponse_(parseSheetPJUM_(getSheetOrCreate("Data PJUM")), params.prefix);
    }
    return createJsonResponse_(parseSheetPP_(getSheetOrCreate("Data PP")), params.prefix);
  }

  // 4. Ambil Daftar Vendor
  if (action === "getVendors") {
    var vendorSheet = ensureVendorSeedData();
    var vendorRows = vendorSheet.getDataRange().getValues();
    var vendors = [];
    for (var j = 1; j < vendorRows.length; j++) {
      if (vendorRows[j][1] && String(vendorRows[j][1]).trim() !== "") {
        vendors.push({
          no: Number(vendorRows[j][0]) || j,
          nama: vendorRows[j][1] || "",
          sebagai: vendorRows[j][2] || "",
          anPenerima: vendorRows[j][3] || "",
          bank: vendorRows[j][4] || "",
          noRekening: vendorRows[j][5] || ""
        });
      }
    }
    return createJsonResponse_(vendors, params.prefix);
  }

  // Default /exec tanpa parameter untuk Dashboard SEO
  if (!action) {
    return createJsonResponse_(getDataPPForSEO_(), params.prefix);
  }

  return createJsonResponse_({ success: false, message: "Action tidak dikenal: " + action }, params.prefix);
}