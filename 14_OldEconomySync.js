function syncInvoicesToOldLog() {
  const oldSS = SpreadsheetApp.openById(CFG.OLD_ECONOMY_SPREADSHEET_ID);
  const oldLog = oldSS.getSheetByName("LOG");
  if (!oldLog) throw new Error("No existe la hoja LOG en el sistema viejo.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shInv = ss.getSheetByName("INVOICES");
  const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);

  if (!shInv) throw new Error("No existe la hoja INVOICES.");
  if (!shEco) throw new Error("No existe la hoja ECONOMY.");

  const invData = shInv.getDataRange().getValues();
  const ecoData = shEco.getDataRange().getValues();

  if (invData.length < 2) return { success: true, added: 0 };

  const invHeaders = invData[0].map(h => String(h).trim());
  const ecoHeaders = ecoData[0].map(h => String(h).trim());

  const oldData = oldLog.getDataRange().getValues();
  const oldHeaders = oldData[0].map(h => String(h).trim().toUpperCase());

  const oldInvoiceCol = oldHeaders.indexOf("INVOICE");
  if (oldInvoiceCol === -1) throw new Error("LOG viejo debe tener columna Invoice.");

  const existing = {};
  for (let i = 1; i < oldData.length; i++) {
    const inv = String(oldData[i][oldInvoiceCol] || "").trim();
    if (inv) existing[inv] = true;
  }

  function getByHeader(headers, row, name) {
    const idx = headers.indexOf(name);
    return idx >= 0 ? row[idx] : "";
  }

  const ecoMap = {};
  const ecoWO = ecoHeaders.indexOf("WO_NUMBER");

  if (ecoWO === -1) {
    throw new Error("ECONOMY debe tener WO_NUMBER.");
  }

  for (let i = 1; i < ecoData.length; i++) {
    const wo = String(ecoData[i][ecoWO] || "").trim();
    if (wo) ecoMap[wo] = ecoData[i];
  }

  const rowsToAdd = [];

  for (let i = 1; i < invData.length; i++) {
    const invRow = invData[i];

    const invoice = String(getByHeader(invHeaders, invRow, "Invoice") || "").trim();
    if (!invoice || existing[invoice]) continue;

    const woNumber = String(getByHeader(invHeaders, invRow, "WO_NUMBER") || "").trim();
    const ecoRow = ecoMap[woNumber] || [];

    const invSource =
      getByHeader(ecoHeaders, ecoRow, "INV_SOURCE") ||
      getByHeader(invHeaders, invRow, "INVERSION") ||
      "";

    const cost =
      getByHeader(ecoHeaders, ecoRow, "COST") ||
      getByHeader(invHeaders, invRow, "INVERSION") ||
      "";

    const technicians =
      getByHeader(ecoHeaders, ecoRow, "TECHNICIANS") ||
      getByHeader(invHeaders, invRow, "TECHNICIAN NAME") ||
      "";

    rowsToAdd.push([
      getByHeader(invHeaders, invRow, "Timestamp"),
      invoice,
      getByHeader(invHeaders, invRow, "NS"),
      getByHeader(invHeaders, invRow, "Source"),
      getByHeader(invHeaders, invRow, "TECHNICIAN EMAIL"),
      technicians,
      "RECIBIDO",
      getByHeader(invHeaders, invRow, "HORAS"),
      getByHeader(invHeaders, invRow, "PARTS"),
      getByHeader(invHeaders, invRow, "LABOR"),
      getByHeader(invHeaders, invRow, "TAX"),
      getByHeader(invHeaders, invRow, "TOTAL"),
      getByHeader(invHeaders, invRow, "CLIENTE"),
      cost
    ]);
  }

  if (rowsToAdd.length) {
    oldLog
      .getRange(oldLog.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length)
      .setValues(rowsToAdd);
  }

  return {
    success: true,
    added: rowsToAdd.length
  };
}

const ECONOMY_HISTORY_SHEET_NAME = "ECONOMY_HISTORY";
const ECONOMY_HISTORY_HEADERS = [
  "COMPANY_ID",
  "SOURCE_TYPE",
  "SOURCE_SHEET_ROW",
  "DATE_INVOICE",
  "PERIOD_YEAR",
  "PERIOD_MONTH",
  "PERIOD_LABEL",
  "INVOICE_NUMBER",
  "WO_NUMBER",
  "CLIENT",
  "NSN",
  "STATUS",
  "HORAS",
  "LABOR_HOURS",
  "LABOR_BILLED",
  "PARTS_BILLED",
  "TAX",
  "AMOUNT",
  "COST",
  "INV_SOURCE",
  "TECHNICIANS",
  "TECHNICIAN_EMAIL",
  "NOTES",
  "READ_ONLY",
  "IMPORTED_AT",
  "ACTIVE"
];

function syncOldEconomyLogToHistory(companyId, sessionToken) {
  companyId = String(companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);

  const oldSS = SpreadsheetApp.openById(CFG.OLD_ECONOMY_SPREADSHEET_ID);
  const oldLog = oldSS.getSheetByName("LOG");
  if (!oldLog) throw new Error("No existe la hoja LOG en el sistema viejo.");

  const oldData = oldLog.getDataRange().getValues();
  if (oldData.length < 2) {
    return { success: true, imported: 0, updated: 0, skippedActive: 0, skippedNoInvoice: 0 };
  }

  const oldHeaders = oldData[0].map(function(h) {
    return String(h || "").trim();
  });

  const history = ensureEconomyHistorySheet_();
  const historyHeaders = ensureEconomyHistoryHeaders_(history);
  const activeKeys = getActiveEconomyInvoiceKeySet_(companyId);
  const existingHistory = getEconomyHistoryInvoiceRowMap_(history, historyHeaders, companyId);
  const now = new Date();

  let imported = 0;
  let updated = 0;
  let skippedActive = 0;
  let skippedNoInvoice = 0;
  const importedInvoices = [];
  const updatedInvoices = [];
  const skippedActiveInvoices = [];

  for (let i = 1; i < oldData.length; i++) {
    const oldRow = oldData[i];
    const normalized = buildEconomyHistoryRowFromOldLog_(oldRow, oldHeaders, companyId, i + 1, now);
    const invoiceKey = getEconomyHistoryInvoiceKey_(companyId, normalized.INVOICE_NUMBER);

    if (!invoiceKey) {
      skippedNoInvoice++;
      continue;
    }

    if (activeKeys[invoiceKey]) {
      skippedActive++;
      skippedActiveInvoices.push(normalized.INVOICE_NUMBER);
      continue;
    }

    const values = historyHeaders.map(function(h) {
      return normalized[h] !== undefined ? normalized[h] : "";
    });

    if (existingHistory[invoiceKey]) {
      history.getRange(existingHistory[invoiceKey], 1, 1, historyHeaders.length).setValues([values]);
      updated++;
      updatedInvoices.push(normalized.INVOICE_NUMBER);
    } else {
      history.appendRow(values);
      existingHistory[invoiceKey] = history.getLastRow();
      imported++;
      importedInvoices.push(normalized.INVOICE_NUMBER);
    }
  }

  addAuditLog_("ECONOMY", "OLD_ECONOMY_HISTORY_SYNCED", companyId, "ECONOMY_HISTORY", "LOG", session, {
    imported: imported,
    updated: updated,
    skippedActive: skippedActive,
    skippedNoInvoice: skippedNoInvoice,
    importedInvoices: importedInvoices,
    updatedInvoices: updatedInvoices,
    skippedActiveInvoices: skippedActiveInvoices
  });

  return {
    success: true,
    imported: imported,
    updated: updated,
    skippedActive: skippedActive,
    skippedNoInvoice: skippedNoInvoice,
    importedInvoices: importedInvoices,
    updatedInvoices: updatedInvoices,
    skippedActiveInvoices: skippedActiveInvoices
  };
}

function buildEconomyHistoryRowFromOldLog_(oldRow, oldHeaders, companyId, sourceRow, importedAt) {
  const invoiceDate = parseOldEconomyDate_(getOldEconomyValue_(oldRow, oldHeaders, [
    "DATE_INVOICE",
    "INVOICE_DATE",
    "TIMESTAMP",
    "DATE",
    "FECHA"
  ]));
  const period = getEconomyHistoryPeriodFromDate_(invoiceDate);
  const rawSource = String(getOldEconomyValue_(oldRow, oldHeaders, [
    "INV_SOURCE",
    "INVERSION_SOURCE",
    "INVESTOR",
    "INVERSION"
  ]) || "").trim().toUpperCase();
  const invSource = normalizeOldEconomyInvSource_(rawSource);

  return {
    COMPANY_ID: companyId,
    SOURCE_TYPE: "OLD_LOG",
    SOURCE_SHEET_ROW: sourceRow,
    DATE_INVOICE: invoiceDate || "",
    PERIOD_YEAR: period.year,
    PERIOD_MONTH: period.month,
    PERIOD_LABEL: period.label,
    INVOICE_NUMBER: String(getOldEconomyValue_(oldRow, oldHeaders, [
      "INVOICE_NUMBER",
      "INVOICE",
      "INVOICE #",
      "Invoice"
    ]) || "").trim(),
    WO_NUMBER: String(getOldEconomyValue_(oldRow, oldHeaders, [
      "WO_NUMBER",
      "WO",
      "WORK_ORDER",
      "ORDER"
    ]) || "").trim(),
    CLIENT: getOldEconomyValue_(oldRow, oldHeaders, ["CLIENT", "CLIENTE", "CUSTOMER"]) || "",
    NSN: getOldEconomyValue_(oldRow, oldHeaders, ["NSN", "NSN #", "NS"]) || "",
    STATUS: normalizeOldEconomyStatus_(getOldEconomyValue_(oldRow, oldHeaders, ["STATUS", "ESTATUS", "ESTADO"])),
    HORAS: parseMoneyFlexible_(getOldEconomyValue_(oldRow, oldHeaders, ["HORAS", "HOURS", "LABOR_HOURS"])),
    LABOR_HOURS: parseMoneyFlexible_(getOldEconomyValue_(oldRow, oldHeaders, ["LABOR_HOURS", "HORAS", "HOURS"])),
    LABOR_BILLED: parseMoneyFlexible_(getOldEconomyValue_(oldRow, oldHeaders, [
      "LABOR_BILLED",
      "LABOR_AMOUNT",
      "LABOR",
      "COBRO_HORAS",
      "TOTAL_COBRO_HORAS"
    ])),
    PARTS_BILLED: parseMoneyFlexible_(getOldEconomyValue_(oldRow, oldHeaders, [
      "PARTS_BILLED",
      "PARTS_AMOUNT",
      "PARTS",
      "PART",
      "COBRO_PARTS",
      "TOTAL_COBRO_PARTS"
    ])),
    TAX: parseMoneyFlexible_(getOldEconomyValue_(oldRow, oldHeaders, ["TAX", "TAX_AMOUNT"])),
    AMOUNT: parseMoneyFlexible_(getOldEconomyValue_(oldRow, oldHeaders, ["AMOUNT", "TOTAL", "INVOICE_TOTAL"])),
    COST: parseMoneyFlexible_(getOldEconomyValue_(oldRow, oldHeaders, [
      "COST",
      "MATERIAL_COST",
      "INVERSION",
      "COMPRA",
      "PARTS_COST"
    ])),
    INV_SOURCE: invSource,
    TECHNICIANS: getOldEconomyValue_(oldRow, oldHeaders, [
      "TECHNICIANS",
      "TECHNICIAN",
      "TECHNICIAN NAME",
      "TECNICO",
      "TECNICOS"
    ]) || "",
    TECHNICIAN_EMAIL: getOldEconomyValue_(oldRow, oldHeaders, ["TECHNICIAN_EMAIL", "TECHNICIAN EMAIL", "EMAIL"]) || "",
    NOTES: getOldEconomyValue_(oldRow, oldHeaders, ["NOTES", "NOTE", "NOTAS"]) || "",
    READ_ONLY: "YES",
    IMPORTED_AT: importedAt,
    ACTIVE: "YES"
  };
}

function ensureEconomyHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ECONOMY_HISTORY_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(ECONOMY_HISTORY_SHEET_NAME);
    sh.getRange(1, 1, 1, ECONOMY_HISTORY_HEADERS.length).setValues([ECONOMY_HISTORY_HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }

  ensureEconomyHistoryHeaders_(sh);
  return sh;
}

function ensureEconomyHistoryHeaders_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });

  if (headers.length === 1 && !headers[0]) {
    headers = [];
  }

  ECONOMY_HISTORY_HEADERS.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
    }
  });

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return headers;
}

function getEconomyHistoryObjectsForCompany_(companyId, existingInvoiceKeys) {
  companyId = String(companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  existingInvoiceKeys = existingInvoiceKeys || {};

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ECONOMY_HISTORY_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) {
    return String(h || "").trim();
  });
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = String(getHeaderValueFlexible_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID)
      .trim()
      .toUpperCase();
    if (rowCompany !== companyId) continue;

    const invoiceNumber = String(getHeaderValueFlexible_(row, headers, ["INVOICE_NUMBER", "INVOICE"]) || "").trim();
    const key = getEconomyHistoryInvoiceKey_(companyId, invoiceNumber);
    if (key && existingInvoiceKeys[key]) continue;

    const obj = {};
    headers.forEach(function(h, c) {
      let value = row[c];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy");
      }
      obj[h] = value;
    });

    obj.ROW_NUMBER = i + 2;
    obj.SOURCE_TYPE = obj.SOURCE_TYPE || "OLD_LOG";
    obj.READ_ONLY = "YES";
    rows.push(obj);
  }

  return rows.reverse();
}

function getActiveEconomyInvoiceKeySet_(companyId) {
  companyId = String(companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  const keys = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_ECONOMY);
  if (!sh || sh.getLastRow() < 2) return keys;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) {
    return String(h || "").trim();
  });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = String(getHeaderValueFlexible_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID)
      .trim()
      .toUpperCase();
    if (rowCompany !== companyId) continue;

    const invoiceNumber = String(getHeaderValueFlexible_(row, headers, ["INVOICE_NUMBER", "INVOICE"]) || "").trim();
    const key = getEconomyHistoryInvoiceKey_(companyId, invoiceNumber);
    if (key) keys[key] = true;
  }

  return keys;
}

function getEconomyHistoryInvoiceRowMap_(sh, headers, companyId) {
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  headers = headers.map(function(h) {
    return String(h || "").trim();
  });

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowCompany = String(getHeaderValueFlexible_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID)
      .trim()
      .toUpperCase();
    if (rowCompany !== companyId) continue;

    const invoiceNumber = String(getHeaderValueFlexible_(row, headers, ["INVOICE_NUMBER", "INVOICE"]) || "").trim();
    const key = getEconomyHistoryInvoiceKey_(companyId, invoiceNumber);
    if (key) map[key] = i + 2;
  }

  return map;
}

function getEconomyHistoryInvoiceKey_(companyId, invoiceNumber) {
  companyId = String(companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  invoiceNumber = String(invoiceNumber || "").trim().toUpperCase();
  if (!companyId || !invoiceNumber) return "";
  return companyId + "::" + invoiceNumber;
}

function getOldEconomyValue_(row, headers, possibleNames) {
  const normalized = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  });

  for (let i = 0; i < possibleNames.length; i++) {
    const wanted = String(possibleNames[i] || "").trim().toUpperCase();
    const idx = normalized.indexOf(wanted);
    if (idx >= 0) return row[idx];
  }

  return "";
}

function parseOldEconomyDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) return parsed;

  const text = String(value || "").trim();
  const match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
  if (!match) return null;

  let year = Number(match[3]);
  if (year < 100) year += 2000;

  const date = new Date(year, Number(match[1]) - 1, Number(match[2]));
  return isNaN(date.getTime()) ? null : date;
}

function getEconomyHistoryPeriodFromDate_(dateValue) {
  if (!dateValue) return { year: "", month: "", label: "" };

  const year = dateValue.getFullYear();
  const month = dateValue.getMonth() + 1;
  return {
    year: year,
    month: month,
    label: year + "-" + String(month).padStart(2, "0")
  };
}

function normalizeOldEconomyInvSource_(value) {
  value = String(value || "").trim().toUpperCase();
  if (value === "DAVID" || value === "YOEL" || value === "MISCELANEAS") return value;
  if (value === "MISC" || value === "MISCELLANEOUS" || value === "MISCELANEA") return "MISCELANEAS";
  return "";
}

function normalizeOldEconomyStatus_(value) {
  value = String(value || "").trim().toUpperCase();
  if (!value) return "INVOICED";
  if (value === "RECIBIDO" || value === "RECEIVED") return "INVOICED";
  if (value === "PAGADO" || value === "PAID") return "PAID";
  if (value === "FACTURADO") return "INVOICED";
  if (value === "PENDIENTE") return "PENDING";
  return value;
}
