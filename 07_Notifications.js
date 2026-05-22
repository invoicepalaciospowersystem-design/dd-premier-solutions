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

function getNotifications(companyId, role, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  role = String(role || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES", "ECONOMIA"], companyId);
  return readNotificationsForCompany_(companyId);
}

function readNotificationsForCompany_(companyId) {
  companyId = String(companyId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NOTIFICATIONS);
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(String);

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

function getNotificationsByUser(name, companyId, sessionToken) {
  const session = requireNamedSession_(sessionToken, ["TECH", "SUPERVISOR", "OWNER", "ADMIN"], companyId, name, "usuario");
  const effectiveName = String(session.role || "").toUpperCase() === "OWNER" || String(session.role || "").toUpperCase() === "ADMIN"
    ? name
    : session.name;
  const all = readNotificationsForCompany_(companyId);
  const target = String(effectiveName || "").trim().toLowerCase();

  return all.filter(function(n) {
    const to = String(n.TO || n.ROLE || "").trim().toLowerCase();
    return to === target || (String(session.role || "").toUpperCase() === "TECH" && to === "tech");
  });
}

function markNotificationsRead(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES", "ECONOMIA"], companyId);

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

function markNotificationsReadByUser(name, companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireNamedSession_(sessionToken, ["TECH", "SUPERVISOR", "OWNER", "ADMIN"], companyId, name, "usuario");

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NOTIFICATIONS);
  if (!sh || sh.getLastRow() < 2) return true;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);

  const colCompany = headers.indexOf("COMPANY_ID");
  let colTo = headers.indexOf("TO");
  if (colTo === -1) colTo = headers.indexOf("ROLE");
  const colRead = headers.indexOf("READ");

  if (colCompany === -1 || colTo === -1 || colRead === -1) return true;

  const target = String(
    String(session.role || "").toUpperCase() === "OWNER" || String(session.role || "").toUpperCase() === "ADMIN"
      ? name
      : session.name
  || "").trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][colCompany] || "").trim().toUpperCase();
    const rowTo = String(data[i][colTo] || "").trim().toLowerCase();

    if (rowCompany === companyId && rowTo === target) {
      sh.getRange(i + 1, colRead + 1).setValue("YES");
    }
  }

  return true;
}
