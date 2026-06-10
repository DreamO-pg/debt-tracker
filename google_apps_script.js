// ============================================================
// Google Apps Script — вставьте этот код в редактор Apps Script
// (Extensions > Apps Script) вашей Google Таблицы
// ============================================================

const SHEET_CLIENTS = 'Клиенты';
const SHEET_DEBTS = 'Долги';
const SHEET_PAYMENTS = 'Платежи';
const SHEET_META = 'Мета';
const SHEET_LOG = 'Журнал';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const action = e.parameter.action;

    if (action === 'load') {
      output.setContent(JSON.stringify({ success: true, data: loadAllData() }));
    } else if (action === 'save') {
      const data = JSON.parse(e.postData.contents);
      saveAllData(data);
      output.setContent(JSON.stringify({ success: true }));
    } else {
      output.setContent(JSON.stringify({ success: false, error: 'Unknown action' }));
    }
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.toString() }));
  }

  return output;
}

function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_CLIENTS)) {
    const s = ss.insertSheet(SHEET_CLIENTS);
    s.appendRow(['id', 'name', 'phone', 'createdAt']);
  }
  if (!ss.getSheetByName(SHEET_DEBTS)) {
    const s = ss.insertSheet(SHEET_DEBTS);
    s.appendRow(['id', 'clientId', 'amount', 'createdAt', 'dueDate', 'comment']);
  }
  if (!ss.getSheetByName(SHEET_PAYMENTS)) {
    const s = ss.insertSheet(SHEET_PAYMENTS);
    s.appendRow(['id', 'debtId', 'amount', 'date', 'comment']);
  }
  if (!ss.getSheetByName(SHEET_META)) {
    const s = ss.insertSheet(SHEET_META);
    s.appendRow(['key', 'value']);
    s.appendRow(['nextClientId', '1']);
    s.appendRow(['nextDebtId', '1']);
    s.appendRow(['nextPayId', '1']);
  }
  if (!ss.getSheetByName(SHEET_LOG)) {
    const s = ss.insertSheet(SHEET_LOG);
    s.appendRow(['user', 'userName', 'action', 'detail', 'timestamp']);
  }
}

function loadAllData() {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const clients = sheetToArray(ss.getSheetByName(SHEET_CLIENTS));
  const debts = sheetToArray(ss.getSheetByName(SHEET_DEBTS));
  const payments = sheetToArray(ss.getSheetByName(SHEET_PAYMENTS));
  const meta = loadMeta(ss.getSheetByName(SHEET_META));
  const log = sheetToArray(ss.getSheetByName(SHEET_LOG));

  return {
    clients: clients,
    debts: debts,
    payments: payments,
    log: log,
    nextClientId: parseInt(meta.nextClientId) || 1,
    nextDebtId: parseInt(meta.nextDebtId) || 1,
    nextPayId: parseInt(meta.nextPayId) || 1
  };
}

function saveAllData(data) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  arrayToSheet(ss.getSheetByName(SHEET_CLIENTS), data.clients,
    ['id', 'name', 'phone', 'createdAt']);
  arrayToSheet(ss.getSheetByName(SHEET_DEBTS), data.debts,
    ['id', 'clientId', 'amount', 'createdAt', 'dueDate', 'comment']);
  arrayToSheet(ss.getSheetByName(SHEET_PAYMENTS), data.payments,
    ['id', 'debtId', 'amount', 'date', 'comment']);
  if (data.log) {
    arrayToSheet(ss.getSheetByName(SHEET_LOG), data.log,
      ['user', 'userName', 'action', 'detail', 'timestamp']);
  }

  const metaSheet = ss.getSheetByName(SHEET_META);
  saveMeta(metaSheet, {
    nextClientId: data.nextClientId || 1,
    nextDebtId: data.nextDebtId || 1,
    nextPayId: data.nextPayId || 1
  });
}

function sheetToArray(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (headers[j] === 'id' || headers[j] === 'clientId' || headers[j] === 'debtId' ||
          headers[j] === 'amount' || headers[j] === 'nextClientId' || headers[j] === 'nextDebtId' || headers[j] === 'nextPayId') {
        val = Number(val);
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return result;
}

function arrayToSheet(sheet, arr, headers) {
  sheet.clear();
  sheet.appendRow(headers);
  arr.forEach(obj => {
    const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
    sheet.appendRow(row);
  });
}

function loadMeta(sheet) {
  const data = sheet.getDataRange().getValues();
  const meta = {};
  for (let i = 1; i < data.length; i++) {
    meta[data[i][0]] = data[i][1];
  }
  return meta;
}

function saveMeta(sheet, meta) {
  sheet.clear();
  sheet.appendRow(['key', 'value']);
  Object.keys(meta).forEach(k => {
    sheet.appendRow([k, meta[k]]);
  });
}
