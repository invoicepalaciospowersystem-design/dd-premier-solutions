// =====================================================
// FILE: 07_Notifications.gs
// =====================================================

function addNotification_(companyId, to, wo, type, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_NOTIFICATIONS);

  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_NOTIFICATIONS);
    sh.appendRow(["COMPANY_ID", "TIMESTAMP", "TO", "WO_NUMBER", "TYPE", "MESSAGE", "READ"]);
  }

  sh.appendRow([companyId || CFG.DEFAULT_COMPANY_ID, new Date(), to, wo, type, message, "NO"]);
}

function getNotifications(companyId, role) {
  companyId = String(companyId || "").trim().toUpperCase();
  role = String(role || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NOTIFICATIONS);
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(String);
  const idxCompany = headers.indexOf("COMPANY_ID");

  return data.slice(1).map(function(row) {
    const obj = {};

    headers.forEach(function(h, i) {
      let value = row[i];

      if (value instanceof Date) {
        value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a");
      }

      obj[h] = value;
    });

    return obj;
  }).filter(function(n) {
    if (!companyId) return false;

    return String(n.COMPANY_ID || "").trim().toUpperCase() === companyId;
  }).reverse();
}

function getNotificationsByUser(name, companyId) {
  const all = getNotifications(companyId, "");
  const target = String(name || "").trim().toLowerCase();

  return all.filter(function(n) {
    const to = String(n.TO || "").trim().toLowerCase();
    return to === target || to === "tech";
  });
}

function markNotificationsRead(companyId) {
  companyId = String(companyId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NOTIFICATIONS);
  if (!sh || sh.getLastRow() < 2) return true;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);

  const colCompany = headers.indexOf("COMPANY_ID");
  const colRead = headers.indexOf("READ");

  if (colCompany === -1 || colRead === -1) return true;

  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][colCompany] || "").trim().toUpperCase();

    if (rowCompany === companyId) {
      sh.getRange(i + 1, colRead + 1).setValue("YES");
    }
  }

  return true;
}

function markNotificationsReadByUser(name, companyId) {
  companyId = String(companyId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NOTIFICATIONS);
  if (!sh || sh.getLastRow() < 2) return true;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);

  const colCompany = headers.indexOf("COMPANY_ID");
  const colTo = headers.indexOf("TO");
  const colRead = headers.indexOf("READ");

  if (colCompany === -1 || colTo === -1 || colRead === -1) return true;

  const target = String(name || "").trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][colCompany] || "").trim().toUpperCase();
    const rowTo = String(data[i][colTo] || "").trim().toLowerCase();

    if (rowCompany === companyId && rowTo === target) {
      sh.getRange(i + 1, colRead + 1).setValue("YES");
    }
  }

  return true;
}