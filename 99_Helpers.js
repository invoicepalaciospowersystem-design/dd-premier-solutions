// =====================================================
// FILE: 99_Helpers.gs
// =====================================================

function addLog_(companyId, woNumber, action, oldStatus, newStatus, user, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WO_LOG);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WO_LOG);

  sh.appendRow([companyId || CFG.DEFAULT_COMPANY_ID, new Date(), woNumber, action, oldStatus, newStatus, user, notes]);
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
