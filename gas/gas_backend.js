// ============================================================
// ハウジングロビー 評価シート — Google Apps Script バックエンド
// このコードをGoogle Apps Scriptに貼り付けてデプロイする
// ============================================================

const SHEET_NAME = "評価データ";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    if (data.action === "save") {
      saveEvaluation(sheet, data);
      return jsonResponse({ status: "ok" });
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (e.parameter.action === "getAll") {
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return jsonResponse({ records: [] });
    const headers = rows[0];
    const records = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return jsonResponse({ records });
  }
  return jsonResponse({ status: "ok" });
}

function saveEvaluation(sheet, data) {
  const id = data.name + "_" + data.quarter;
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  const newRow = headers.map(h => {
    if (h === "ID") return id;
    if (h === "タイムスタンプ") return new Date().toLocaleString("ja-JP");
    return data[h] !== undefined ? data[h] : "";
  });

  // 既存行を探して上書き
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([newRow]);
      return;
    }
  }
  // 新規追加
  sheet.appendRow(newRow);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "ID", "タイムスタンプ", "社員名", "部署", "役職",
      "四半期", "KPIデータ(JSON)", "合計点"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#1F4E79")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
    sheet.setColumnWidth(7, 400);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
