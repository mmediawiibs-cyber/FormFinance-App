const SHEET_INPUT = "Sheet-A";
const SHEET_REPORT = "Sheet-B";
const SPREADSHEET_ID = "PASTE_SPREADSHEET_ID_HERE";

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "PASTE_SPREADSHEET_ID_HERE") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureHeaders(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = firstRow.every((cell) => cell === "");
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function toRowFromData(data) {
  return [
    data.nomor || "",
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
    data.nominal || "",
    data.lampiran || "",
    data.tanggalEntry || "",
    data.status || "NOT PAID",
    data.tanggalTf || "",
    data.jumlahTf || "",
    data.tanggalRekap || "",
    data.keterangan || "",
  ];
}

function normalizeRow(rawRow) {
  const safe = rawRow || [];
  return {
    nomor: safe[0] || "",
    divisi: safe[1] || "",
    nomorPp: safe[2] || "",
    namaPembuat: safe[3] || "",
    tanggalPp: safe[4] || "",
    tanggalPemenuhan: safe[5] || "",
    tujuan: safe[6] || "",
    uraian1: safe[7] || "",
    uraian2: safe[8] || "",
    uraian3: safe[9] || "",
    kepada: safe[10] || "",
    sebagai: safe[11] || "",
    anPenerima: safe[12] || "",
    bank: safe[13] || "",
    noRekening: safe[14] || "",
    nominal: safe[15] || "",
    lampiran: safe[16] || "",
    tanggalEntry: safe[17] || "",
    status: safe[18] || "NOT PAID",
    tanggalTf: safe[19] || "",
    jumlahTf: safe[20] || "",
    tanggalRekap: safe[21] || "",
    keterangan: safe[22] || "",
  };
}

function listRows(sheetName) {
  const sheet = ensureSheet(sheetName);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map((row) => normalizeRow(row));
}

function doPost(e) {
  try {
    const raw =
      e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const data = JSON.parse(raw);
    const sheet = ensureSheet(SHEET_INPUT);

    const headers = [
      "No",
      "PP",
      "Nomor PP",
      "Nama Pembuat",
      "Tanggal PP",
      "Tanggal Pemenuhan",
      "Tujuan",
      "Uraian I",
      "Uraian II",
      "Uraian III",
      "Kepada",
      "Sebagai",
      "A/N Penerima",
      "Bank",
      "No Rekening",
      "Nominal",
      "Lampiran",
      "Tanggal Entry",
      "Status",
      "Tanggal TF",
      "Jumlah TF",
      "Tanggal Rekap",
      "Keterangan",
    ];

    ensureHeaders(sheet, headers);

    if (data.action === "updateData") {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (Number(rows[i][0]) === Number(data.nomor)) {
          const target = i + 1;
          sheet
            .getRange(target, 8, 1, 16)
            .setValues([
              [
                data.uraian1 || "",
                data.uraian2 || "",
                data.uraian3 || "",
                data.kepada || "",
                data.sebagai || "",
                data.anPenerima || "",
                data.bank || "",
                data.noRekening || "",
                data.nominal || "",
                data.lampiran || "",
                data.tanggalEntry || "",
                data.status || "NOT PAID",
                data.tanggalTf || "",
                data.jumlahTf || "",
                data.tanggalRekap || "",
                data.keterangan || "",
              ],
            ]);
          return ContentService.createTextOutput(
            JSON.stringify({
              status: "success",
              message: "Data berhasil diperbarui",
            }),
          ).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    if (data.action === "deleteData") {
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (Number(rows[i][0]) === Number(data.nomor)) {
          sheet.deleteRow(i + 1);
          return ContentService.createTextOutput(
            JSON.stringify({
              status: "success",
              message: "Baris berhasil dihapus",
            }),
          ).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    const row = toRowFromData(data);
    sheet.appendRow(row);

    const reportSheet = ensureSheet(SHEET_REPORT);
    ensureHeaders(reportSheet, headers);
    reportSheet.appendRow(row);

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Data tersimpan ke Sheet-A dan Sheet-B",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "getAllData";
    const reportSheet = ensureSheet(SHEET_REPORT);
    const sourceSheet = ensureSheet(SHEET_INPUT);

    if (action === "syncFromSheetA") {
      const rows = sourceSheet.getDataRange().getValues();
      const all = rows.slice(1);
      reportSheet.clear();
      const headers = [
        "No",
        "PP",
        "Nomor PP",
        "Nama Pembuat",
        "Tanggal PP",
        "Tanggal Pemenuhan",
        "Tujuan",
        "Uraian I",
        "Uraian II",
        "Uraian III",
        "Kepada",
        "Sebagai",
        "A/N Penerima",
        "Bank",
        "No Rekening",
        "Nominal",
        "Lampiran",
        "Tanggal Entry",
        "Status",
        "Tanggal TF",
        "Jumlah TF",
        "Tanggal Rekap",
        "Keterangan",
      ];
      reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (all.length > 0) {
        reportSheet.getRange(2, 1, all.length, headers.length).setValues(all);
      }
      return ContentService.createTextOutput(
        JSON.stringify({ status: "success", total: all.length }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "getAllData") {
      const data = listRows(SHEET_REPORT);
      return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
        ContentService.MimeType.JSON,
      );
    }

    if (action === "cetakPDF") {
      const sheet = reportSheet || sourceSheet;
      const targetRow = Number(e.parameter.nomor || 1);
      const range = sheet.getDataRange();
      const values = range.getValues();

      let rowIndex = 1;
      for (let i = 1; i < values.length; i++) {
        if (Number(values[i][0]) === targetRow) {
          rowIndex = i + 1;
          break;
        }
      }

      const url =
        "https://docs.google.com/spreadsheets/d/" +
        getSpreadsheet().getId() +
        "/export?format=pdf&portrait=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=true&gid=" +
        sheet.getSheetId();
      return HtmlService.createHtmlOutput(
        "<script>window.location.href = " + JSON.stringify(url) + ";</script>",
      );
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: "ok" }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
