/**
 * FORM FINANCE STIT — GOOGLE APPS SCRIPT (FULL MULTI-SHEET: PP, PUM, PJUM, VENDOR & SEO)
 *
 * Fitur:
 * 1. Mendukung penuh operasi CRUD (Insert, Update/Upsert di Nomor yang Sama, Delete)
 *    untuk Sheet: "Data PP", "Data PUM", dan "Data PJUM".
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

// Mendeteksi otomatis apakah Sheet Data PP memakai 23 kolom atau 24 kolom
function buildRowDataPP_(sheet, data) {
  var nominalValue = normalizeNominal(data.nominal);
  var lampiranLink = uploadLampiranIfAny_(data);
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var col5Header = String(headers[4] || "").trim().toUpperCase();

  // Jika kolom ke-5 adalah DIVISI (24 kolom)
  if (col5Header === "DIVISI") {
    return [
      Number(data.nomor || 0),
      data.divisi || "",
      data.nomorPp || "",
      data.namaPembuat || "",
      data.divisiNama || data.divisi || "",
      data.tanggalPp || "",
      data.tanggalPemenuhan || "",
      data.tujuan || "",
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

  // Standar 23 kolom Data PP
  return [
    Number(data.nomor || 0),
    data.divisi || "",
    data.nomorPp || "",
    data.namaPembuat || "",
    data.tanggalPp || "",
    data.tanggalPemenuhan || "",
    data.tujuan || "",
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

// Susunan 25 Kolom Sheet "Data PUM" sesuai Gambar 1
function buildRowDataPUM_(data) {
  var nominalValue = normalizeNominal(data.nominal);
  var lampiranLink = uploadLampiranIfAny_(data);

  return [
    Number(data.nomor || 0),               // 1. NO
    data.divisi || "",                     // 2. PP
    data.nomorPum || "",                   // 3. NOMOR PUM
    data.namaPembuat || "",                // 4. NAMA PEMBUAT
    data.divisiNama || data.divisi || "",  // 5. DIVISI
    data.wakaTerkait || "",                // 6. WAKA TERKAIT
    data.tanggalPp || "",                  // 7. TANGGAL PP
    data.tanggalPemenuhan || "",           // 8. TANGGAL PEMENUHAN
    data.tujuan || "",                     // 9. TUJUAN
    data.uraian1 || "",                    // 10. URAIAN I
    data.uraian2 || "",                    // 11. URAIAN II
    data.uraian3 || "",                    // 12. URAIAN III
    data.kepada || "",                     // 13. KEPADA
    data.sebagai || "",                    // 14. SEBAGAI
    data.anPenerima || "",                 // 15. A/N PENERIMA
    data.bank || "",                       // 16. BANK
    data.noRekening || "",                 // 17. NO. REKENING
    nominalValue,                          // 18. NOMINAL
    lampiranLink,                          // 19. LAMPIRAN
    data.tanggalEntry || "",               // 20. TANGGAL ENTRY
    data.status || "NOT PAID",             // 21. STATUS
    data.tanggalTf || "",                  // 22. TANGGAL TF
    data.jumlahTf ? normalizeNominal(data.jumlahTf) : "", // 23. JUMLAH TF
    data.tanggalRekap || "",               // 24. TANGGAL
    data.keterangan || ""                  // 25. KETERANGAN
  ];
}

// Susunan 13 Kolom Sheet "Data PJUM" sesuai Gambar 2
function buildRowDataPJUM_(data) {
  var uangMukaVal = normalizeNominal(data.uangMuka);
  var terpakaiVal = normalizeNominal(data.terpakai);
  var selisihVal = uangMukaVal - terpakaiVal;

  return [
    Number(data.nomor || 0),               // 1. NO (PUM)
    data.pp || data.divisi || "",          // 2. PP (PUM)
    data.nomorPjum || "",                  // 3. NOMOR PJUM
    data.namaPembuat || "",                // 4. NAMA PEMBUAT (PUM)
    data.divisi || "",                     // 5. DIVISI (PUM)
    data.wakaTerkait || "",                // 6. WAKA TERKAIT (PUM)
    data.tanggalPp || "",                  // 7. TANGGAL PP (PUM)
    data.tanggalPemenuhan || "",           // 8. TANGGAL PEMENUHAN (PUM)
    data.uraianPjum || "",                 // 9. URAIAN (Input PJUM)
    uangMukaVal,                           // 10. UANG MUKA (PUM)
    terpakaiVal,                           // 11. TERPAKAI (Input PJUM)
    selisihVal,                            // 12. KURANG/LEBIH
    data.keterangan || (selisihVal >= 0 ? "LEBIH" : "KURANG") // 13. KETERANGAN
  ];
}

// ============================================================
// PARSER BARIS FLEKSIBEL (ANTI GESER KOLOM)
// ============================================================

function parseSheetPP_(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var headers = rows[0].map(function(h) { return String(h || "").trim().toUpperCase(); });
  var hasExtraDivisiCol = (headers[4] === "DIVISI");
  var list = [];

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[0] === "" && r[2] === "") continue;
    if (hasExtraDivisiCol) {
      list.push({
        nomor: Number(r[0]) || i,
        divisi: r[1] || "",
        nomorPp: r[2] || "",
        namaPembuat: r[3] || "",
        divisiNama: r[4] || r[1] || "",
        tanggalPp: r[5] || "",
        tanggalPemenuhan: r[6] || "",
        tujuan: r[7] || "",
        uraian1: r[8] || "",
        uraian2: r[9] || "",
        uraian3: r[10] || "",
        kepada: r[11] || "",
        sebagai: r[12] || "",
        anPenerima: r[13] || "",
        bank: r[14] || "",
        noRekening: r[15] || "",
        nominal: r[16] || 0,
        lampiran: r[17] || "-",
        tanggalEntry: r[18] || "",
        status: r[19] || "NOT PAID",
        tanggalTf: r[20] || "",
        jumlahTf: r[21] || "",
        tanggalRekap: r[22] || "",
        keterangan: r[23] || ""
      });
    } else {
      list.push({
        nomor: Number(r[0]) || i,
        divisi: r[1] || "",
        nomorPp: r[2] || "",
        namaPembuat: r[3] || "",
        tanggalPp: r[4] || "",
        tanggalPemenuhan: r[5] || "",
        tujuan: r[6] || "",
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
  }
  return list;
}

function parseSheetPUM_(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[0] === "" && r[2] === "") continue;
    list.push({
      nomor: Number(r[0]) || i,
      divisi: r[1] || "",
      nomorPum: r[2] || "",
      namaPembuat: r[3] || "",
      divisiNama: r[4] || r[1] || "",
      wakaTerkait: r[5] || "",
      tanggalPp: r[6] || "",
      tanggalPemenuhan: r[7] || "",
      tujuan: r[8] || "",
      uraian1: r[9] || "",
      uraian2: r[10] || "",
      uraian3: r[11] || "",
      kepada: r[12] || "",
      sebagai: r[13] || "",
      anPenerima: r[14] || "",
      bank: r[15] || "",
      noRekening: r[16] || "",
      nominal: r[17] || 0,
      lampiran: r[18] || "-",
      tanggalEntry: r[19] || "",
      status: r[20] || "NOT PAID",
      tanggalTf: r[21] || "",
      jumlahTf: r[22] || "",
      tanggalRekap: r[23] || "",
      keterangan: r[24] || ""
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
    list.push({
      nomor: Number(r[0]) || i,
      pp: r[1] || "",
      nomorPjum: r[2] || "",
      namaPembuat: r[3] || "",
      divisi: r[4] || "",
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
        // Untuk keamanan penuh: kita tulis kolom input PJUM (Kolom 3, 9, 11, 12, 13) & kolom referensi jika kosong
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