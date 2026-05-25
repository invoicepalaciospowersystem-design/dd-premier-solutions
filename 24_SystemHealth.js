// =====================================================
// FILE: 24_SystemHealth.gs
// =====================================================

const SYSTEM_HEALTH_BACKUP_PREFIX = "System Backup - ";
const SYSTEM_HEALTH_TRIGGER_HANDLERS = [
  "runSystemHealthCheck",
  "runDailySystemBackup"
];

function getSystemHealthDashboard(sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER"]);
  const report = buildSystemHealthReport_();

  addAuditLog_("SYSTEM_HEALTH", "SYSTEM_HEALTH_VIEWED", "ALL", "SYSTEM", "HEALTH", session, {
    status: report.status,
    openErrors: report.summary.openErrors,
    criticalChecks: report.summary.criticalChecks
  });

  return report;
}

function runSystemHealthCheck() {
  try {
    const report = buildSystemHealthReport_();
    const props = PropertiesService.getScriptProperties();

    props.setProperty("SYSTEM_HEALTH_LAST_STATUS", report.status);
    props.setProperty("SYSTEM_HEALTH_LAST_RUN_AT", String(Date.now()));

    if (report.status !== "OK") {
      sendSystemHealthAlert_(report);
    }

    return report;
  } catch (err) {
    notifySystemError_("SYSTEM_HEALTH_CHECK_ERROR", err, {
      module: "SYSTEM_HEALTH"
    });
    throw err;
  }
}

function installSystemMaintenanceTriggers(sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER"]);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (SYSTEM_HEALTH_TRIGGER_HANDLERS.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("runSystemHealthCheck")
    .timeBased()
    .everyMinutes(Number(CFG.SYSTEM_HEALTH_TRIGGER_MINUTES || 15))
    .create();

  ScriptApp.newTrigger("runDailySystemBackup")
    .timeBased()
    .everyDays(1)
    .atHour(Number(CFG.SYSTEM_BACKUP_HOUR || 2))
    .create();

  addAuditLog_("SYSTEM_HEALTH", "SYSTEM_MAINTENANCE_TRIGGERS_INSTALLED", "ALL", "SYSTEM", "TRIGGERS", session, {
    healthMinutes: Number(CFG.SYSTEM_HEALTH_TRIGGER_MINUTES || 15),
    backupHour: Number(CFG.SYSTEM_BACKUP_HOUR || 2)
  });

  return getSystemTriggerStatus_();
}

function createSystemBackup(sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER"]);
  const result = createSystemBackup_();

  addAuditLog_("SYSTEM_HEALTH", "SYSTEM_BACKUP_CREATED", "ALL", "SYSTEM", result.fileId || "", session, {
    name: result.name,
    url: result.url
  });

  return result;
}

function runDailySystemBackup() {
  try {
    return createSystemBackup_();
  } catch (err) {
    notifySystemError_("SYSTEM_BACKUP_ERROR", err, {
      module: "SYSTEM_HEALTH"
    });
    throw err;
  }
}

function markSystemErrorResolved(rowNumber, sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER"]);
  rowNumber = Number(rowNumber || 0);

  if (!rowNumber || rowNumber < 2) {
    throw new Error("Fila de error invalida.");
  }

  const sh = getSystemErrorSheet_();
  const headers = ensureSystemErrorLogHeaders_(sh);
  const rowValues = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  const context = getSystemHealthValue_(rowValues, headers, ["CONTEXT"]);
  const message = getSystemHealthValue_(rowValues, headers, ["MESSAGE"]);

  setSystemHealthCell_(sh, rowNumber, headers, "STATUS", "RESOLVED");
  setSystemHealthCell_(sh, rowNumber, headers, "RESOLVED_AT", new Date());
  setSystemHealthCell_(sh, rowNumber, headers, "RESOLVED_BY", session.email || session.name || "OWNER");

  addAuditLog_("SYSTEM_HEALTH", "SYSTEM_ERROR_RESOLVED", "ALL", "ERROR_LOG", String(rowNumber), session, {
    context: context,
    message: message
  });

  return true;
}

function buildSystemHealthReport_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = [];
  const summary = {
    openErrors: 0,
    openErrors24h: 0,
    criticalChecks: 0,
    warningChecks: 0,
    duplicateWorkOrders: 0,
    missingPdfLinks: 0,
    failedDocumentEmails: 0,
    missingSecrets: 0
  };
  const removedSessions = typeof cleanupExpiredSessions_ === "function" ? cleanupExpiredSessions_() : 0;

  checkRequiredSheets_(ss, checks);
  checkSecurityProperties_(checks, summary);
  addSystemCheck_(checks, "SESSION_CLEANUP", "Sesiones expiradas", "OK",
    removedSessions ? removedSessions + " sesiones expiradas limpiadas." : "No habia sesiones expiradas pendientes.", removedSessions);
  const errors = collectSystemErrors_(Number(CFG.SYSTEM_HEALTH_ERROR_LIMIT || 25), summary);
  const duplicates = findDuplicateWorkOrders_(ss);
  const documentIssues = findDocumentIssues_(ss);
  const emailIssues = findDocumentEmailIssues_(ss);
  const backupStatus = getSystemBackupStatus_();
  const triggerStatus = getSystemTriggerStatus_();

  summary.duplicateWorkOrders = duplicates.length;
  summary.missingPdfLinks = documentIssues.length;
  summary.failedDocumentEmails = emailIssues.length;

  addSystemCheck_(checks, "ERROR_LOGS", "Errores abiertos", summary.openErrors ? "WARNING" : "OK",
    summary.openErrors ? summary.openErrors + " errores sin resolver." : "No hay errores abiertos.", summary.openErrors);

  addSystemCheck_(checks, "DUPLICATE_WORK_ORDERS", "Ordenes duplicadas", duplicates.length ? "CRITICAL" : "OK",
    duplicates.length ? duplicates.length + " WO_NUMBER duplicados detectados." : "No se detectaron ordenes duplicadas.", duplicates.length);

  addSystemCheck_(checks, "PDF_LINKS", "PDF links", documentIssues.length ? "WARNING" : "OK",
    documentIssues.length ? documentIssues.length + " documentos marcados como generados sin URL." : "PDF links principales estan completos.", documentIssues.length);

  addSystemCheck_(checks, "CLIENT_EMAILS", "Emails de documentos", emailIssues.length ? "WARNING" : "OK",
    emailIssues.length ? emailIssues.length + " documentos con email fallido o pendiente." : "No hay fallos recientes de emails de documentos.", emailIssues.length);

  addSystemCheck_(checks, "BACKUPS", "Backups", backupStatus.status, backupStatus.detail, backupStatus.ageHours || 0);
  addSystemCheck_(checks, "TRIGGERS", "Vigilancia automatica", triggerStatus.status, triggerStatus.detail, triggerStatus.installedCount);

  checks.forEach(function(check) {
    if (check.status === "CRITICAL") summary.criticalChecks++;
    if (check.status === "WARNING") summary.warningChecks++;
  });

  const status = summary.criticalChecks ? "CRITICAL" : (summary.warningChecks ? "WARNING" : "OK");

  return {
    generatedAt: Utilities.formatDate(new Date(), CFG.TIMEZONE, "MM/dd/yyyy hh:mm a"),
    status: status,
    summary: summary,
    checks: checks,
    errors: errors,
    duplicateWorkOrders: duplicates.slice(0, 20),
    documentIssues: documentIssues.slice(0, 20),
    emailIssues: emailIssues.slice(0, 20),
    backup: backupStatus,
    triggers: triggerStatus
  };
}

function checkRequiredSheets_(ss, checks) {
  const required = [
    CFG.SHEET_COMPANIES,
    CFG.SHEET_USERS,
    CFG.SHEET_WORK_ORDERS,
    CFG.SHEET_STORES,
    CFG.SHEET_ECONOMY,
    CFG.SHEET_AUDIT_LOGS || "AUDIT_LOGS",
    CFG.SHEET_ERROR_LOGS || "ERROR_LOGS"
  ].filter(Boolean);

  const missing = required.filter(function(name) {
    return !ss.getSheetByName(name);
  });

  addSystemCheck_(checks, "REQUIRED_SHEETS", "Hojas requeridas", missing.length ? "CRITICAL" : "OK",
    missing.length ? "Faltan hojas: " + missing.join(", ") : "Todas las hojas principales existen.", missing.length);
}

function checkSecurityProperties_(checks, summary) {
  const security = validateSecurityProperties();
  const missing = security.missing || [];
  summary.missingSecrets = missing.length;

  addSystemCheck_(checks, "SECRETS", "Script Properties", missing.length ? "CRITICAL" : "OK",
    missing.length ? "Faltan secretos: " + missing.join(", ") : "Secretos requeridos configurados.", missing.length);
}

function collectSystemErrors_(limit, summary) {
  const sh = getSystemErrorSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];

  const headers = ensureSystemErrorLogHeaders_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const since = Date.now() - (Number(CFG.SYSTEM_HEALTH_LOOKBACK_HOURS || 24) * 60 * 60 * 1000);
  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const errors = [];

  rows.forEach(function(row, index) {
    const status = String(getSystemHealthValue_(row, headers, ["STATUS"]) || "OPEN").trim().toUpperCase();
    const timestamp = parseSystemHealthDate_(getSystemHealthValue_(row, headers, ["TIMESTAMP"]));

    if (status !== "RESOLVED" && status !== "IGNORED") {
      summary.openErrors++;
      if (timestamp && timestamp.getTime() >= since) summary.openErrors24h++;
    }

    errors.push({
      rowNumber: index + 2,
      timestamp: timestamp ? Utilities.formatDate(timestamp, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a") : String(getSystemHealthValue_(row, headers, ["TIMESTAMP"]) || ""),
      context: String(getSystemHealthValue_(row, headers, ["CONTEXT"]) || ""),
      message: String(getSystemHealthValue_(row, headers, ["MESSAGE"]) || ""),
      details: String(getSystemHealthValue_(row, headers, ["DETAILS"]) || ""),
      userEmail: String(getSystemHealthValue_(row, headers, ["USER_EMAIL"]) || ""),
      status: status || "OPEN",
      resolvedAt: String(getSystemHealthValue_(row, headers, ["RESOLVED_AT"]) || ""),
      resolvedBy: String(getSystemHealthValue_(row, headers, ["RESOLVED_BY"]) || ""),
      fingerprint: String(getSystemHealthValue_(row, headers, ["FINGERPRINT"]) || "")
    });
  });

  return errors.reverse().slice(0, limit || 25);
}

function findDuplicateWorkOrders_(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });
  const seen = {};
  const duplicates = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const wo = String(getSystemHealthValue_(row, headers, ["WO_NUMBER"]) || "").trim();
    if (!wo) continue;

    if (!seen[wo]) {
      seen[wo] = [i + 1];
      continue;
    }

    seen[wo].push(i + 1);
  }

  Object.keys(seen).forEach(function(wo) {
    if (seen[wo].length > 1) {
      duplicates.push({
        woNumber: wo,
        rows: seen[wo].join(", ")
      });
    }
  });

  return duplicates;
}

function findDocumentIssues_(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });
  const issues = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const wo = String(getSystemHealthValue_(row, headers, ["WO_NUMBER"]) || "").trim();
    const invoiceFlag = String(getSystemHealthValue_(row, headers, ["HAS_INVOICE_PDF"]) || "").trim().toUpperCase();
    const invoiceUrl = String(getSystemHealthValue_(row, headers, ["PDF_EN_URL", "PDF_ES_URL"]) || "").trim();
    const pmFlag = String(getSystemHealthValue_(row, headers, ["HAS_PM_REPORT_PDF", "PM_REPORT_STATUS"]) || "").trim().toUpperCase();
    const pmUrl = String(getSystemHealthValue_(row, headers, ["PM_REPORT_EN_URL", "PM_REPORT_ES_URL"]) || "").trim();

    if (invoiceFlag === "YES" && !invoiceUrl) {
      issues.push({ rowNumber: i + 1, woNumber: wo, type: "INVOICE_PDF", issue: "Marcado con invoice PDF pero sin URL." });
    }

    if ((pmFlag === "YES" || pmFlag === "GENERATED") && !pmUrl) {
      issues.push({ rowNumber: i + 1, woNumber: wo, type: "PM_REPORT", issue: "Marcado con PM report generado pero sin URL." });
    }
  }

  return issues;
}

function findDocumentEmailIssues_(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });
  const normalized = headers.map(function(h) { return h.toUpperCase(); });
  const statusIndexes = normalized.map(function(h, index) {
    return h.indexOf("_EMAIL_STATUS") !== -1 ? index : -1;
  }).filter(function(index) { return index >= 0; });
  const issues = [];

  if (!statusIndexes.length) return issues;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    statusIndexes.forEach(function(idx) {
      const status = String(row[idx] || "").trim().toUpperCase();
      if (!status || status === "SENT" || status === "SKIPPED_TEST_MODE") return;

      const sentHeader = String(headers[idx] || "").replace(/_STATUS$/i, "_SENT");
      const sentValue = String(getSystemHealthValue_(row, headers, [sentHeader]) || "").trim().toUpperCase();
      if (sentValue === "YES") return;

      issues.push({
        rowNumber: i + 1,
        woNumber: String(getSystemHealthValue_(row, headers, ["WO_NUMBER"]) || ""),
        type: headers[idx],
        issue: status
      });
    });
  }

  return issues;
}

function getSystemBackupStatus_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("SYSTEM_LAST_BACKUP_JSON") || "";

  if (!raw) {
    return {
      status: "WARNING",
      detail: "No hay backup registrado todavia.",
      lastBackupAt: "",
      url: "",
      ageHours: 0
    };
  }

  try {
    const parsed = JSON.parse(raw);
    const createdAt = Number(parsed.createdAt || 0);
    const ageHours = createdAt ? Math.round((Date.now() - createdAt) / 360000) / 10 : 0;
    const maxAge = Number(CFG.SYSTEM_BACKUP_MAX_AGE_HOURS || 30);

    return {
      status: createdAt && ageHours <= maxAge ? "OK" : "WARNING",
      detail: createdAt ? "Ultimo backup hace " + ageHours + " horas." : "Backup sin fecha valida.",
      lastBackupAt: parsed.createdAtLabel || "",
      url: parsed.url || "",
      name: parsed.name || "",
      ageHours: ageHours
    };
  } catch (err) {
    return {
      status: "WARNING",
      detail: "No se pudo leer el ultimo backup registrado.",
      lastBackupAt: "",
      url: "",
      ageHours: 0
    };
  }
}

function getSystemTriggerStatus_() {
  const triggers = ScriptApp.getProjectTriggers().map(function(trigger) {
    return {
      handler: trigger.getHandlerFunction(),
      source: String(trigger.getTriggerSource && trigger.getTriggerSource() || ""),
      eventType: String(trigger.getEventType && trigger.getEventType() || "")
    };
  });

  const installed = SYSTEM_HEALTH_TRIGGER_HANDLERS.filter(function(handler) {
    return triggers.some(function(trigger) {
      return trigger.handler === handler;
    });
  });

  return {
    status: installed.length === SYSTEM_HEALTH_TRIGGER_HANDLERS.length ? "OK" : "WARNING",
    detail: installed.length === SYSTEM_HEALTH_TRIGGER_HANDLERS.length
      ? "Health check y backup diario instalados."
      : "Faltan triggers: " + SYSTEM_HEALTH_TRIGGER_HANDLERS.filter(function(handler) {
          return installed.indexOf(handler) === -1;
        }).join(", "),
    installedCount: installed.length,
    requiredCount: SYSTEM_HEALTH_TRIGGER_HANDLERS.length,
    triggers: triggers
  };
}

function createSystemBackup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceFile = DriveApp.getFileById(ss.getId());
  const folder = getSystemBackupFolder_();
  const stamp = Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyyMMdd-HHmmss");
  const name = SYSTEM_HEALTH_BACKUP_PREFIX + stamp + " - " + (CFG.APP_NAME || "App");
  const copy = sourceFile.makeCopy(name, folder);

  try {
    copy.setShareableByEditors(false);
    copy.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  } catch (err) {
    Logger.log("WARN createSystemBackup_ sharing: " + err);
  }

  pruneSystemBackups_(folder, Number(CFG.SYSTEM_BACKUP_MAX_FILES || 14));

  const result = {
    fileId: copy.getId(),
    name: copy.getName(),
    url: copy.getUrl(),
    createdAt: Date.now(),
    createdAtLabel: Utilities.formatDate(new Date(), CFG.TIMEZONE, "MM/dd/yyyy hh:mm a")
  };

  PropertiesService.getScriptProperties().setProperty("SYSTEM_LAST_BACKUP_JSON", JSON.stringify(result));
  return result;
}

function getSystemBackupFolder_() {
  if (CFG.SYSTEM_BACKUP_FOLDER_ID) {
    return DriveApp.getFolderById(CFG.SYSTEM_BACKUP_FOLDER_ID);
  }

  const name = CFG.SYSTEM_BACKUP_FOLDER_NAME || "D&D Premier System Backups";
  const folders = DriveApp.getRootFolder().getFoldersByName(name);
  if (folders.hasNext()) return folders.next();

  return DriveApp.getRootFolder().createFolder(name);
}

function pruneSystemBackups_(folder, maxFiles) {
  maxFiles = Math.max(Number(maxFiles || 14), 1);
  const files = [];
  const iterator = folder.getFiles();

  while (iterator.hasNext()) {
    const file = iterator.next();
    if (String(file.getName() || "").indexOf(SYSTEM_HEALTH_BACKUP_PREFIX) !== 0) continue;
    files.push(file);
  }

  files.sort(function(a, b) {
    return b.getDateCreated().getTime() - a.getDateCreated().getTime();
  });

  files.slice(maxFiles).forEach(function(file) {
    try {
      file.setTrashed(true);
    } catch (err) {
      Logger.log("WARN pruneSystemBackups_: " + err);
    }
  });
}

function sendSystemHealthAlert_(report) {
  if (CFG.SYSTEM_ALERTS_ENABLED === false || !CFG.SYSTEM_ALERT_EMAIL) return;

  const minutes = Number(CFG.SYSTEM_HEALTH_ALERT_THROTTLE_MINUTES || 60);
  const key = "SYSTEM_HEALTH_ALERT_SENT_" + String(report.status || "UNKNOWN");
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const last = Number(props.getProperty(key) || 0);

  if (last && now - last < minutes * 60 * 1000) return;

  props.setProperty(key, String(now));

  const topChecks = (report.checks || [])
    .filter(function(check) { return check.status !== "OK"; })
    .slice(0, 8)
    .map(function(check) {
      return "<li><b>" + escapeHtmlForEmail_(check.label) + ":</b> " + escapeHtmlForEmail_(check.detail) + "</li>";
    }).join("");

  MailApp.sendEmail({
    to: CFG.SYSTEM_ALERT_EMAIL,
    subject: "APP HEALTH - " + report.status,
    htmlBody:
      "<h2>System health alert</h2>" +
      "<p><b>Status:</b> " + escapeHtmlForEmail_(report.status) + "</p>" +
      "<p><b>Generated:</b> " + escapeHtmlForEmail_(report.generatedAt) + "</p>" +
      "<ul>" + topChecks + "</ul>"
  });
}

function getSystemErrorSheet_(createIfMissing) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CFG.SHEET_ERROR_LOGS || "ERROR_LOGS";
  let sh = ss.getSheetByName(sheetName);

  if (!sh && createIfMissing !== false) {
    sh = ss.insertSheet(sheetName);
  }

  return sh;
}

function addSystemCheck_(checks, key, label, status, detail, count) {
  checks.push({
    key: key,
    label: label,
    status: status,
    detail: detail,
    count: Number(count || 0)
  });
}

function getSystemHealthValue_(row, headers, possibleNames) {
  const normalized = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  });

  for (let i = 0; i < possibleNames.length; i++) {
    const idx = normalized.indexOf(String(possibleNames[i] || "").trim().toUpperCase());
    if (idx >= 0) return row[idx];
  }

  return "";
}

function setSystemHealthCell_(sh, rowNumber, headers, headerName, value) {
  const normalized = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  });
  const idx = normalized.indexOf(String(headerName || "").trim().toUpperCase());
  if (idx >= 0) sh.getRange(rowNumber, idx + 1).setValue(value);
}

function parseSystemHealthDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}
