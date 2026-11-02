var ss = SpreadsheetApp.getActiveSpreadsheet();

function getBulanRomawi(tanggalStr) {
  var d = new Date(tanggalStr);
  var bulanAngka = d.getMonth() + 1;
  var romawi = [
    "I",
    "II",
    "III",
    "IV",
    "V",
    "VI",
    "VII",
    "VIII",
    "IX",
    "X",
    "XI",
    "XII",
  ];
  return romawi[bulanAngka - 1] || "I";
}

function getSheetOrCreate(sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function ensureVendorSeedData() {
  var sheet = getSheetOrCreate("Data Vendor");
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    sheet.appendRow([
      "Nama Penerima",
      "Sebagai",
      "A/N Penerima",
      "Bank",
      "No Rekening",
    ]);
    sheet.appendRow([
      "Ustadz Afghani",
      "Kabag Media",
      "AFGHANI",
      "BSI",
      "7147151588",
    ]);
    sheet.appendRow([
      "Salwa Nurshalihah",
      "Staff Media Development",
      "SALWA NURSHALIHAH",
      "BSI",
      "7218899001",
    ]);
    sheet.appendRow([
      "Dr. Ahmad Fauzi, M.Pd.",
      "Kaprodi MPI",
      "AHMAD FAUZI",
      "Bank Mandiri",
      "1370098765432",
    ]);
    sheet.appendRow([
      "Budi Santoso, S.Kom.",
      "Kepala Divisi IT",
      "BUDI SANTOSO",
      "BCA",
      "8420991122",
    ]);
    sheet.appendRow([
      "Siti Aminah, M.E.",
      "Kepala BAAU",
      "SITI AMINAH",
      "BRI",
      "002901002345501",
    ]);
    sheet.appendRow([
      "Toko Berkah Jaya",
      "Vendor Peralatan & Percetakan",
      "CV BERKAH JAYA UTAMA",
      "BCA",
      "4561122334",
    ]);
  }
  return sheet;
}

function normalizeNominal(rawValue) {
  var raw = String(rawValue == null ? "0" : rawValue);
  var cleaned = raw.replace(/[^0-9]/g, "");
  return cleaned === "" ? 0 : Number(cleaned);
}

function vendorRowData(data) {
  return [
    data.nama || "",
    data.sebagai || "",
    data.anPenerima || "",
    data.bank || "",
    data.noRekening || "",
  ];
}

function findVendorRowIndexByName(sheet, nama) {
  if (!sheet || !nama) return -1;
  var rows = sheet.getDataRange().getValues();
  var target = String(nama).trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (
      String(rows[i][0] || "")
        .trim()
        .toLowerCase() === target
    ) {
      return i + 1;
    }
  }
  return -1;
}

function findRowIndexByNomor(sheet, nomor) {
  if (!sheet) return -1;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(nomor)) {
      return i + 1;
    }
  }
  return -1;
}

function buildRowData(data) {
  var nominalValue = normalizeNominal(data.nominal);
  var lampiranLink = data.lampiranText || data.lampiran || "-";

  if (data.fileLampiran) {
    try {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(data.fileLampiran.data),
        data.fileLampiran.mimeType,
        data.fileLampiran.name,
      );
      var file = DriveApp.createFile(blob);
      lampiranLink = file.getUrl();
    } catch (errFile) {
      lampiranLink = "Gagal upload: " + errFile.toString();
    }
  }

  // Urutan kolom spreadsheet: A=1, ..., O=15, P=No Rekening, Q=Nominal, R=Lampiran, ...
  var rowData = [
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
    data.noRekening || "", // P
    nominalValue, // Q
    lampiranLink, // R
    data.tanggalEntry || "",
    data.status || "NOT PAID",
    data.tanggalTf || "",
    data.jumlahTf || "",
    data.tanggalRekap || "",
    data.keterangan || "",
  ];

  return rowData;
}

function deleteRowsByNomor(sheetName, nomor) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return false;
  }

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

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheetA = getSheetOrCreate("Data PP");
    var sheetB = getSheetOrCreate("Form PP");

    if (data.action === "updateData") {
      var rowA = findRowIndexByNomor(sheetA, data.nomor);
      var updatedRow = buildRowData(data);

      if (rowA > -1) {
        sheetA.getRange(rowA, 1, 1, updatedRow.length).setValues([updatedRow]);
      }

      if (rowA > -1) {
        return ContentService.createTextOutput(
          JSON.stringify({
            status: "success",
            message: "Data berhasil diperbarui!",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(
        JSON.stringify({
          status: "error",
          message: "Data tidak ditemukan untuk diupdate.",
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "deleteData") {
      var deletedA = deleteRowsByNomor("Data PP", data.nomor);
      var deletedB = deleteRowsByNomor("Form PP", data.nomor);

      if (deletedA || deletedB) {
        return ContentService.createTextOutput(
          JSON.stringify({
            status: "success",
            message: "Data berhasil dihapus dari Data PP dan Form PP!",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "Data tidak ditemukan." }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "saveVendor") {
      var vendorSheet = ensureVendorSeedData();
      var vendorName = String(data.nama || "").trim();
      if (!vendorName) {
        return ContentService.createTextOutput(
          JSON.stringify({
            status: "error",
            message: "Nama vendor wajib diisi.",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }
      if (findVendorRowIndexByName(vendorSheet, vendorName) > -1) {
        return ContentService.createTextOutput(
          JSON.stringify({
            status: "error",
            message: "Nama vendor sudah ada.",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }
      vendorSheet.appendRow(vendorRowData(data));
      return ContentService.createTextOutput(
        JSON.stringify({
          status: "success",
          message: "Vendor baru berhasil ditambahkan.",
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "updateVendor") {
      var vendorSheet = ensureVendorSeedData();
      var originalName = String(data.originalNama || data.nama || "").trim();
      var rowIndex = findVendorRowIndexByName(vendorSheet, originalName);
      if (rowIndex === -1) {
        rowIndex = findVendorRowIndexByName(vendorSheet, data.nama || "");
      }
      if (rowIndex === -1) {
        return ContentService.createTextOutput(
          JSON.stringify({
            status: "error",
            message: "Vendor yang diedit tidak ditemukan.",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }
      vendorSheet.getRange(rowIndex, 1, 1, 5).setValues([vendorRowData(data)]);
      return ContentService.createTextOutput(
        JSON.stringify({
          status: "success",
          message: "Vendor berhasil diperbarui.",
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "deleteVendor") {
      var vendorSheet = ensureVendorSeedData();
      var rowIndex = findVendorRowIndexByName(vendorSheet, data.nama || "");
      if (rowIndex === -1) {
        return ContentService.createTextOutput(
          JSON.stringify({
            status: "error",
            message: "Vendor tidak ditemukan.",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }
      vendorSheet.deleteRow(rowIndex);
      return ContentService.createTextOutput(
        JSON.stringify({
          status: "success",
          message: "Vendor berhasil dihapus.",
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var lastRow = sheetA.getLastRow();
    var nomorUrut = 1;
    if (lastRow > 1) {
      var lastNum = sheetA.getRange(lastRow, 1).getValue();
      nomorUrut =
        isNaN(lastNum) || lastNum === "" ? lastRow : Number(lastNum) + 1;
    }

    var strNoUrut = ("000" + nomorUrut).slice(-3);
    var tahun = new Date(data.tanggalPp).getFullYear();
    var bulanRomawi = getBulanRomawi(data.tanggalPp);
    var nomorPp =
      strNoUrut +
      " / PP.STIT." +
      data.divisi +
      " / " +
      bulanRomawi +
      " / " +
      tahun;

    data.nomor = nomorUrut;
    data.nomorPp = nomorPp;

    var rowData = buildRowData(data);
    sheetA.appendRow(rowData);

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Data berhasil disimpan ke Data PP!",
        nomorUrut: nomorUrut,
        nomorPp: nomorPp,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var action = e.parameter.action;
  var sheetA = getSheetOrCreate("Data PP");
  var sheetB = getSheetOrCreate("Form PP");

  if (action === "getAllData") {
    var sourceSheet = sheetA;
    if (sheetA.getLastRow() <= 1 && sheetB.getLastRow() > 1) {
      sourceSheet = sheetB;
    }
    var rows = sourceSheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] !== "") {
        result.push({
          nomor: Number(rows[i][0]) || 0,
          divisi: rows[i][1] || "",
          nomorPp: rows[i][2] || "",
          namaPembuat: rows[i][3] || "",
          tanggalPp: rows[i][5] || "",
          tanggalPemenuhan: rows[i][6] || "",
          tujuan: rows[i][7] || "",
          uraian1: rows[i][8] || "",
          uraian2: rows[i][9] || "",
          uraian3: rows[i][10] || "",
          kepada: rows[i][11] || "",
          sebagai: rows[i][12] || "",
          anPenerima: rows[i][13] || "",
          bank: rows[i][14] || "",
          noRekening: rows[i][15] || "",
          nominal: rows[i][16] || 0,
          lampiran: rows[i][17] || "-",
          tanggalEntry: rows[i][18] || "",
          status: rows[i][19] || "NOT PAID",
          tanggalTf: rows[i][20] || "",
          jumlahTf: rows[i][21] || "",
          tanggalRekap: rows[i][22] || "",
          keterangan: rows[i][23] || "",
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
      ContentService.MimeType.JSON,
    );
  }

  if (action === "getVendors") {
    var vendorSheet = ensureVendorSeedData();
    var vendorRows = vendorSheet.getDataRange().getValues();
    var vendors = [];
    for (var i = 1; i < vendorRows.length; i++) {
      if (vendorRows[i][0] && String(vendorRows[i][0]).trim() !== "") {
        vendors.push({
          nama: vendorRows[i][0] || "",
          sebagai: vendorRows[i][1] || "",
          anPenerima: vendorRows[i][2] || "",
          bank: vendorRows[i][3] || "",
          noRekening: vendorRows[i][4] || "",
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(vendors)).setMimeType(
      ContentService.MimeType.JSON,
    );
  }

  var nomorCari = e.parameter.nomor;
  if (sheetB && nomorCari) {
    sheetB.getRange("A1").setValue(nomorCari);
    SpreadsheetApp.flush();
    var ssId = ss.getId();
    var sheetId = sheetB.getSheetId();
    var url =
      "https://docs.google.com/spreadsheets/d/" +
      ssId +
      "/export?exportFormat=pdf&format=pdf&size=7&fzr=true&portrait=true&fitw=true&gridlines=true&gid=" +
      sheetId;
    return HtmlService.createHtmlOutput(
      "<script>window.location.href='" + url + "';</script>",
    );
  }

  return ContentService.createTextOutput("API Aktif").setMimeType(
    ContentService.MimeType.JSON,
  );
}
