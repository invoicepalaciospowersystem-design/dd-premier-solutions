// =====================================================
// FILE: 99_Helpers.gs
// =====================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function addLog_(companyId, woNumber, action, oldStatus, newStatus, user, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WO_LOG);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WO_LOG);

  sh.appendRow([companyId || CFG.DEFAULT_COMPANY_ID, new Date(), woNumber, action, oldStatus, newStatus, user, notes]);
}

function normalizeIdentity_(value) {
  return String(value || "").trim().toLowerCase();
}

function requireNamedSession_(sessionToken, allowedRoles, companyId, requestedName, label) {
  const session = requireSession_(sessionToken, allowedRoles, companyId);
  const role = String(session.role || "").trim().toUpperCase();
  const target = normalizeIdentity_(requestedName);

  if (target && role !== "OWNER" && role !== "ADMIN" && normalizeIdentity_(session.name) !== target) {
    throw new Error("No autorizado para ver este " + (label || "portal") + ".");
  }

  return session;
}

function requireWorkOrderSession_(sessionToken, row, headers, allowedRoles, actionName) {
  if (isSoftDeletedRow_(row, headers)) {
    throw new Error("Esta orden fue eliminada o desactivada.");
  }

  const companyId = getRowValue_(row, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
  const session = requireSession_(sessionToken, allowedRoles, companyId);
  const role = String(session.role || "").trim().toUpperCase();

  if (role === "TECH") {
    const assignedTech = String(getRowValue_(row, headers, "TECHNICIAN") || "").trim().toLowerCase();
    if (!assignedTech.includes(String(session.name || "").trim().toLowerCase())) {
      throw new Error("No autorizado para " + (actionName || "usar") + " esta orden.");
    }
  }

  if (role === "SUPERVISOR") {
    const nsn = normalizeNSN_(getRowValue_(row, headers, "NSN"));
    const stores = getSupervisorAllowedStores_(session.name, companyId);
    if (stores.indexOf(nsn) === -1) {
      throw new Error("No autorizado para " + (actionName || "usar") + " esta orden.");
    }
  }

  return session;
}

function getSessionActorLabel_(session) {
  session = session || {};
  return String(session.email || session.name || session.role || "SYSTEM").trim() || "SYSTEM";
}

function ensureSheetColumns_(sh, requiredColumns) {
  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  requiredColumns.forEach(function(columnName) {
    const exists = headers.some(function(h) {
      return String(h || "").trim().toUpperCase() === String(columnName || "").trim().toUpperCase();
    });

    if (!exists) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(columnName);
      headers.push(columnName);
    }
  });

  return sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });
}

function isSoftDeletedRow_(row, headers) {
  const normalized = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  });

  const idxActive = normalized.indexOf("ACTIVE");
  const idxDeletedAt = normalized.indexOf("DELETED_AT");
  const idxStatus = normalized.indexOf("STATUS");

  const active = idxActive >= 0 ? String(row[idxActive] || "").trim().toUpperCase() : "YES";
  const deletedAt = idxDeletedAt >= 0 ? String(row[idxDeletedAt] || "").trim() : "";
  const status = idxStatus >= 0 ? String(row[idxStatus] || "").trim().toUpperCase() : "";

  return active === "NO" || !!deletedAt || status === "DELETED";
}

function isSoftDeletedObject_(obj) {
  obj = obj || {};
  const active = String(obj.ACTIVE || "YES").trim().toUpperCase();
  const deletedAt = String(obj.DELETED_AT || "").trim();
  const status = String(obj.STATUS || "").trim().toUpperCase();

  return active === "NO" || !!deletedAt || status === "DELETED";
}

function addAuditLog_(moduleName, action, companyId, entityType, entityId, actor, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CFG.SHEET_AUDIT_LOGS || "AUDIT_LOGS";
    let sh = ss.getSheetByName(sheetName);

    if (!sh) {
      sh = ss.insertSheet(sheetName);
    }

    if (sh.getLastRow() === 0) {
      sh.appendRow([
        "TIMESTAMP",
        "MODULE",
        "ACTION",
        "COMPANY_ID",
        "ENTITY_TYPE",
        "ENTITY_ID",
        "ACTOR_EMAIL",
        "ACTOR_NAME",
        "ACTOR_ROLE",
        "DETAILS"
      ]);
    }

    actor = actor || {};

    sh.appendRow([
      new Date(),
      moduleName || "",
      action || "",
      companyId || CFG.DEFAULT_COMPANY_ID,
      entityType || "",
      entityId || "",
      actor.email || "",
      actor.name || "",
      actor.role || "",
      safeStringifySystemError_(details || {})
    ]);
  } catch (err) {
    Logger.log("ERROR addAuditLog_: " + err);
  }
}

function getRequiredScriptProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error("Falta " + key + " en Script Properties.");
  return value;
}

function validateSecurityProperties() {
  const missing = [];
  ["TWILIO_SID", "TWILIO_TOKEN", "TWILIO_FROM"].forEach(function(key) {
    if (!PropertiesService.getScriptProperties().getProperty(key)) missing.push(key);
  });

  return {
    environment: CFG.APP_ENV || "PRODUCTION",
    missing: missing,
    ok: missing.length === 0
  };
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

function hardenGeneratedPdfFile_(file) {
  if (!file) return file;

  try {
    file.setShareableByEditors(false);
  } catch (err) {
    Logger.log("WARN hardenGeneratedPdfFile_ setShareableByEditors: " + err);
  }

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log("WARN hardenGeneratedPdfFile_ setSharing: " + err);
  }

  try {
    file.getEditors().forEach(function(editor) {
      try {
        file.removeEditor(editor);
      } catch (removeErr) {
        Logger.log("WARN hardenGeneratedPdfFile_ removeEditor: " + removeErr);
      }
    });
  } catch (err) {
    Logger.log("WARN hardenGeneratedPdfFile_ getEditors: " + err);
  }

  return file;
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
