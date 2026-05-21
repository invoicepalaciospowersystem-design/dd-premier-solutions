// =====================================================
// FILE: 99_Helpers.gs
// =====================================================

function addLog_(companyId, woNumber, action, oldStatus, newStatus, user, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WO_LOG);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WO_LOG);

  sh.appendRow([companyId || CFG.DEFAULT_COMPANY_ID, new Date(), woNumber, action, oldStatus, newStatus, user, notes]);
}

function notifySystemError_(context, err, details) {
  details = details || {};
  context = String(context || "SYSTEM_ERROR").trim() || "SYSTEM_ERROR";

  const info = normalizeSystemError_(err);
  const safeDetails = safeStringifySystemError_(details);

  try {
    appendSystemErrorLog_(context, info, safeDetails);
  } catch (logErr) {
    Logger.log("ERROR appendSystemErrorLog_: " + logErr);
  }

  if (CFG.SYSTEM_ALERTS_ENABLED === false || !CFG.SYSTEM_ALERT_EMAIL) return;

  try {
    if (!shouldSendSystemErrorEmail_(context, info, details)) return;

    const wo = details.woNumber || details.WO_NUMBER || "";
    const subjectParts = ["APP ALERT", context];
    if (wo) subjectParts.push("WO " + wo);

    MailApp.sendEmail({
      to: CFG.SYSTEM_ALERT_EMAIL,
      subject: subjectParts.join(" - "),
      htmlBody:
        "<h2>App error alert</h2>" +
        "<p><b>Context:</b> " + escapeHtmlForEmail_(context) + "</p>" +
        "<p><b>Message:</b> " + escapeHtmlForEmail_(info.message) + "</p>" +
        "<p><b>WO:</b> " + escapeHtmlForEmail_(wo) + "</p>" +
        "<p><b>Details:</b></p><pre>" + escapeHtmlForEmail_(safeDetails) + "</pre>" +
        "<p><b>Stack:</b></p><pre>" + escapeHtmlForEmail_(info.stack) + "</pre>"
    });
  } catch (emailErr) {
    Logger.log("ERROR sending system alert email: " + emailErr);
  }
}

function appendSystemErrorLog_(context, info, safeDetails) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CFG.SHEET_ERROR_LOGS || "ERROR_LOGS";
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "TIMESTAMP",
      "CONTEXT",
      "MESSAGE",
      "STACK",
      "DETAILS",
      "USER_EMAIL"
    ]);
  }

  sh.appendRow([
    new Date(),
    context,
    info.message,
    info.stack,
    safeDetails,
    getActiveUserEmailSafe_()
  ]);
}

function normalizeSystemError_(err) {
  return {
    message: err && err.message ? String(err.message) : String(err || "Unknown error"),
    stack: err && err.stack ? String(err.stack) : ""
  };
}

function shouldSendSystemErrorEmail_(context, info, details) {
  const minutes = Number(CFG.SYSTEM_ALERT_THROTTLE_MINUTES || 10);
  if (minutes <= 0) return true;

  const raw = [
    context,
    info.message,
    details.woNumber || details.WO_NUMBER || "",
    details.docType || "",
    details.module || ""
  ].join("|");

  const digest = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw)
  ).slice(0, 18);

  const key = "SYSTEM_ALERT_SENT_" + digest;
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const last = Number(props.getProperty(key) || 0);

  if (last && now - last < minutes * 60 * 1000) return false;

  props.setProperty(key, String(now));
  return true;
}

function getActiveUserEmailSafe_() {
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (err) {
    return "";
  }
}

function safeStringifySystemError_(value) {
  try {
    return JSON.stringify(value || {}, null, 2).slice(0, 20000);
  } catch (err) {
    return String(value || "").slice(0, 20000);
  }
}

function escapeHtmlForEmail_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function testSystemErrorAlert() {
  notifySystemError_("TEST_SYSTEM_ALERT", new Error("System alert test."), {
    module: "TEST",
    expectedTo: CFG.SYSTEM_ALERT_EMAIL || ""
  });

  return "Alert test sent to " + (CFG.SYSTEM_ALERT_EMAIL || "");
}

function createOrderFolders_(companyId, nsn, woNumber) {
  const root = DriveApp.getFolderById(CFG.ROOT_FOLDER_ID);
  const companyFolder = getOrCreateFolder_(root, companyId || CFG.DEFAULT_COMPANY_ID);
  const clientFolder = getOrCreateFolder_(companyFolder, "McDonalds");
  const nsnFolder = getOrCreateFolder_(clientFolder, "NSN #" + nsn);
  return getOrCreateFolder_(nsnFolder, woNumber);
}

function getOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function getField_(namedValues, names) {
  for (let i = 0; i < names.length; i++) {
    const key = names[i];
    if (namedValues[key] && namedValues[key][0] !== undefined) {
      return String(namedValues[key][0]).trim();
    }
  }
  return "";
}

function detectFormLanguage_(namedValues) {
  if (namedValues["REPORTED PROBLEM"] || namedValues["ORDER PRIORITY"]) return "EN";
  if (namedValues["PROBLEMA REPORTADO"] || namedValues["PRIORIDAD DE LA ORDEN"]) return "ES";
  return "EN";
}

function safeTranslate_(text, from, to) {
  if (!text) return "";
  try {
    return LanguageApp.translate(text, from, to);
  } catch (err) {
    return text;
  }
}

function normalizePriorityEn_(value) {
  const v = String(value || "").toUpperCase();
  if (v.includes("EMERGENCY") || v.includes("EMERGENCIA")) return "Emergency";
  if (v.includes("URGENT") || v.includes("URGENTE")) return "Urgent";
  return "Normal";
}

function normalizePriorityEs_(value) {
  const v = String(value || "").toUpperCase();
  if (v.includes("EMERGENCY") || v.includes("EMERGENCIA")) return "Emergencia";
  if (v.includes("URGENT") || v.includes("URGENTE")) return "Urgente";
  return "Normal";
}

function rowToObject_(headers, rowData) {
  const obj = {};
  headers.forEach(function(h, i) {
    obj[String(h).trim()] = rowData[i];
  });
  return obj;
}

function setCellByHeader_(sheet, row, headers, headerName, value) {
  const col = headers.indexOf(headerName) + 1;
  if (col > 0) sheet.getRange(row, col).setValue(value);
}

function getCellByHeader_(sheet, row, headers, headerName) {
  const col = headers.indexOf(headerName) + 1;
  if (col > 0) return sheet.getRange(row, col).getValue();
  return "";
}

function getNamesFromText_(text) {
  return String(text || "").split(",").map(function(n) { return n.trim(); }).filter(Boolean);
}

function getTechEmails_(technicianText) {
  const names = getNamesFromText_(technicianText);
  const emails = [];

  names.forEach(function(name) {
    const email = CFG.TECHS[name];
    if (!email) throw new Error("No se encontró email para el técnico: " + name);
    emails.push(email);
  });

  return emails.join(",");
}
