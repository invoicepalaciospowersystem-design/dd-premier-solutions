// =====================================================
// FILE: 08_Economy.gs
// =====================================================

function createEconomyFromWO(companyId, woNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);

  if (!shWO) throw new Error("No existe " + CFG.SHEET_WORK_ORDERS);
  if (!shEco) throw new Error("No existe " + CFG.SHEET_ECONOMY);

  ensureSheetColumns_(shEco, [
    "PERIOD_MONTH",
    "PERIOD_YEAR",
    "PERIOD_LABEL",
    "ACCOUNTING_PERIOD",
    "CLOSED_PERIOD",
    "MONTH_CLOSED_AT",
    "MONTH_CLOSED_BY",
    "MONTH_CLOSED_BY_EMAIL"
  ]);

  companyId = String(companyId || "").trim().toUpperCase();
  woNumber = String(woNumber || "").trim();

  const ecoLastRow = shEco.getLastRow();
  const ecoLastCol = shEco.getLastColumn();

  if (ecoLastRow < 1 || ecoLastCol < 1) {
    throw new Error("ECONOMY no tiene headers.");
  }

  const ecoData = shEco.getRange(1, 1, ecoLastRow, ecoLastCol).getValues();
  const ecoHeaders = ecoData[0].map(function(h) {
    return String(h).trim();
  });

  const ecoWO = ecoHeaders.indexOf("WO_NUMBER");
  const ecoCompany = ecoHeaders.indexOf("COMPANY_ID");

  if (ecoWO === -1 || ecoCompany === -1) {
    throw new Error("ECONOMY debe tener COMPANY_ID y WO_NUMBER.");
  }

  for (let i = 1; i < ecoData.length; i++) {
    const existingWO = String(ecoData[i][ecoWO] || "").trim();
    const existingCompany = String(ecoData[i][ecoCompany] || "").trim().toUpperCase();

    if (existingWO === woNumber && existingCompany === companyId) {
      return true;
    }
  }

  const woData = shWO.getDataRange().getValues();
  const woHeaders = woData[0].map(function(h) {
    return String(h).trim();
  });

  const idxWO = woHeaders.indexOf("WO_NUMBER");
  const idxCompany = woHeaders.indexOf("COMPANY_ID");
  const idxClient = woHeaders.indexOf("CLIENT");
  const idxNSN = woHeaders.indexOf("NSN");
  const idxCompleted = woHeaders.indexOf("DATE_COMPLETED");
  const idxWOType = woHeaders.indexOf("WO_TYPE");
const idxWOTypo = woHeaders.indexOf("WO_TYPO");
const idxPMType = woHeaders.indexOf("PM_TYPE");

  for (let i = 1; i < woData.length; i++) {
    if (
      String(woData[i][idxWO] || "").trim() === woNumber &&
      String(woData[i][idxCompany] || "").trim().toUpperCase() === companyId
    ) {

      const woType = idxWOType >= 0 ? String(woData[i][idxWOType] || "").trim().toUpperCase() : "";
const woTypo = idxWOTypo >= 0 ? String(woData[i][idxWOTypo] || "").trim().toUpperCase() : "";
const pmType = idxPMType >= 0 ? String(woData[i][idxPMType] || "").trim().toUpperCase() : "";

if (
  woType === "PM_FORM" ||
  woTypo === "PM_FORM" ||
  pmType
) {
  return true;
}

      let period = null;
      try {
        period = getCurrentEconomyPeriod();
      } catch (err) {
        period = null;
      }

      const rowObj = {
        COMPANY_ID: companyId,
        WO_NUMBER: woNumber,
        CLIENT: woData[i][idxClient] || "",
        NSN: woData[i][idxNSN] || "",
        DATE_COMPLETED: woData[i][idxCompleted] || "",
        AMOUNT: "",
        TAX: "",
        COST: "",
        PROFIT: "",
        STATUS: "PENDING",
        INVOICE_NUMBER: "",
        DATE_INVOICE: "",
        DATE_PAID: "",
        HORAS: "",
        LABOR_HOURS: "",
        TECHNICIANS: "",
        TECH_PAY_DETAIL: "",
        TECH_LABOR_COST: "",
        TECH_LABOR_PAY: "",
        NOTES: "",
        INV_SOURCE: "",
        PERIOD_MONTH: period ? period.month : "",
        PERIOD_YEAR: period ? period.year : "",
        PERIOD_LABEL: period ? period.label : "",
        ACCOUNTING_PERIOD: period ? period.label : ""
      };

      const newRow = ecoHeaders.map(function(h) {
        return rowObj[h] !== undefined ? rowObj[h] : "";
      });

      shEco.appendRow(newRow);
      return true;
    }
  }

  return false;
}

function getEconomyData(companyId, role, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_ECONOMY);
  if (!sh) throw new Error("No existe la hoja ECONOMY.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    throw new Error("La hoja ECONOMY no tiene headers.");
  }

  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  const idxCompany = headers.indexOf("COMPANY_ID");
  if (idxCompany === -1) {
    throw new Error("ECONOMY debe tener columna COMPANY_ID.");
  }

  const activeInvoiceKeys = {};
  const activeRows = data.slice(1).map(function(row, i) {
    if (isSoftDeletedRow_(row, headers)) return null;

    const obj = {};

    headers.forEach(function(h, c) {
      let value = row[c];

      if (value instanceof Date) {
        value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy");
      }

      obj[h] = value;
    });

    obj.ROW_NUMBER = i + 2;
    const invoiceKey = getEconomyHistoryInvoiceKey_(obj.COMPANY_ID || companyId, obj.INVOICE_NUMBER || obj.INVOICE || "");
    if (invoiceKey) activeInvoiceKeys[invoiceKey] = true;
    return obj;
  }).filter(function(o) {
    if (!o) return false;
    return String(o.COMPANY_ID || "").trim().toUpperCase() === companyId;
  }).reverse();

  const historyRows = getEconomyHistoryObjectsForCompany_(companyId, activeInvoiceKeys);
  return activeRows.concat(historyRows);
}

function updateEconomyRow(rowNumber, updates, sessionToken) {
  rowNumber = Number(rowNumber);
  updates = updates || {};
  if (!rowNumber || rowNumber < 2) throw new Error("Fila inválida.");

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_ECONOMY);
  if (!sh) throw new Error("No existe la hoja ECONOMY.");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const companyId = getCellByHeader_(sh, rowNumber, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  updates = restrictEconomyClosedRowUpdates_(sh, rowNumber, headers, updates);
  const updateKeys = Object.keys(updates);

  updateKeys.forEach(function(key) {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sh.getRange(rowNumber, col).setValue(updates[key]);
  });

  const amountCol = headers.indexOf("AMOUNT") + 1;
  const costCol = headers.indexOf("COST") + 1;
  const profitCol = headers.indexOf("PROFIT") + 1;

  if (amountCol > 0 && costCol > 0 && profitCol > 0) {
    const amount = parseMoneyFlexible_(sh.getRange(rowNumber, amountCol).getValue());
    const tax = getEconomyRowNumberByHeader_(sh, rowNumber, headers, ["TAX", "TAX_AMOUNT"]);
    const rawCost = parseMoneyFlexible_(sh.getRange(rowNumber, costCol).getValue());
    const techLabor = getEconomyRowNumberByHeader_(sh, rowNumber, headers, ["TECH_LABOR_COST", "TECH_LABOR_PAY"]);
    const invSource = getEconomyRowValueByHeader_(sh, rowNumber, headers, ["INV_SOURCE"]);
    const realCost = calculateEconomyRealPartsCost_(invSource, rawCost);

    sh.getRange(rowNumber, profitCol).setValue(amount - tax - realCost - techLabor);
  }

  const statusCol = headers.indexOf("STATUS") + 1;
  const dateInvoiceCol = headers.indexOf("DATE_INVOICE") + 1;
  const datePaidCol = headers.indexOf("DATE_PAID") + 1;

  const status = statusCol > 0
    ? String(sh.getRange(rowNumber, statusCol).getValue() || "").toUpperCase()
    : "";
  const statusWasUpdated = updateKeys.map(function(key) {
    return String(key || "").trim().toUpperCase();
  }).indexOf("STATUS") >= 0;

  if (statusWasUpdated && status === "INVOICED" && dateInvoiceCol > 0 && !sh.getRange(rowNumber, dateInvoiceCol).getValue()) {
    sh.getRange(rowNumber, dateInvoiceCol).setValue(new Date());
  }

  if (statusWasUpdated && status === "PAID" && datePaidCol > 0 && !sh.getRange(rowNumber, datePaidCol).getValue()) {
    sh.getRange(rowNumber, datePaidCol).setValue(new Date());
  }

  addAuditLog_("ECONOMY", "ECONOMY_ROW_UPDATED", companyId, "ECONOMY", String(rowNumber), session, {
    rowNumber: rowNumber,
    fields: Object.keys(updates || {})
  });

  return true;
}

function restrictEconomyClosedRowUpdates_(sh, rowNumber, headers, updates) {
  updates = updates || {};
  if (!isEconomyRowClosedForManualEdit_(sh, rowNumber, headers)) return updates;

  const allowed = {};
  const blocked = [];

  Object.keys(updates).forEach(function(key) {
    const normalizedKey = String(key || "").trim().toUpperCase();
    if (normalizedKey === "INV_SOURCE") {
      allowed[key] = updates[key];
    } else {
      blocked.push(key);
    }
  });

  if (blocked.length || !Object.keys(allowed).length) {
    throw new Error("Mes contable cerrado. Solo se permite actualizar INV_SOURCE.");
  }

  return allowed;
}

function isEconomyRowClosedForManualEdit_(sh, rowNumber, headers) {
  const row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  const closedPeriod = String(getHeaderValueFromRow_(row, headers, "CLOSED_PERIOD") || "").trim();
  const closedAt = getHeaderValueFromRow_(row, headers, "MONTH_CLOSED_AT");
  if (closedPeriod || closedAt) return true;

  const rowPeriod = getEconomyRowPeriodLabelFromData_(row, headers);
  let currentPeriodLabel = "";
  try {
    currentPeriodLabel = getCurrentEconomyPeriod().label;
  } catch (err) {
    currentPeriodLabel = "";
  }

  return !!(rowPeriod && currentPeriodLabel && rowPeriod !== currentPeriodLabel);
}

function previewMcDonaldsPaymentText(companyId, remittanceText, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  const parsed = parseMcDonaldsRemittanceText_(remittanceText);
  const preview = buildMcDonaldsPaymentPreview_(companyId, parsed);

  addAuditLog_("ECONOMY", "MCD_PAYMENT_PREVIEWED", companyId, "ECONOMY", parsed.paymentNumber || "", session, {
    paymentNumber: parsed.paymentNumber || "",
    invoiceCount: parsed.invoices.length,
    matchedCount: preview.matchedCount,
    unmatchedCount: preview.unmatchedCount
  });

  return preview;
}

function applyMcDonaldsPaymentText(companyId, remittanceText, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    const parsed = parseMcDonaldsRemittanceText_(remittanceText);
    if (!parsed.invoices.length) {
      throw new Error("No se encontraron invoices en el remittance.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);
    if (!shEco) throw new Error("No existe la hoja ECONOMY.");

    const result = applyMcDonaldsPaymentToEconomy_(shEco, companyId, parsed, session);
    const invoiceUpdate = applyMcDonaldsPaymentToInvoices_(ss, companyId, parsed, session);

    result.invoiceRowsUpdated = invoiceUpdate.updatedRows;
    result.invoiceRowsMissing = invoiceUpdate.missingInvoices;
    result.paymentNumber = parsed.paymentNumber;
    result.paymentDate = parsed.paymentDateText;
    result.paymentAmount = parsed.paymentAmount;

    addAuditLog_("ECONOMY", "MCD_PAYMENT_IMPORTED", companyId, "ECONOMY", parsed.paymentNumber || "", session, {
      paymentNumber: parsed.paymentNumber || "",
      paymentDate: parsed.paymentDateText || "",
      paymentAmount: parsed.paymentAmount,
      updatedRows: result.updatedRows,
      alreadyPaidRows: result.alreadyPaidRows,
      unmatchedInvoices: result.unmatchedInvoices,
      invoiceRowsUpdated: result.invoiceRowsUpdated
    });

    return result;
  } catch (err) {
    notifySystemError_("MCD_PAYMENT_IMPORT_ERROR", err, {
      module: "ECONOMY",
      companyId: companyId,
      actorEmail: session.email || "",
      actorRole: session.role || ""
    });
    throw err;
  } finally {
    if (locked) lock.releaseLock();
  }
}

function parseMcDonaldsRemittanceText_(remittanceText) {
  const text = String(remittanceText || "").replace(/\r/g, "\n");
  if (!text.trim()) throw new Error("Pegue el texto del remittance primero.");

  const paymentNumberMatch = text.match(/Payment\s+Number\s+([A-Za-z0-9-]+)/i);
  const paymentDateMatch = text.match(/Payment\s+Date\s+(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i);
  const paymentAmountMatch = text.match(/Payment\s+Amount\s+([\d,]+(?:\.\d{2})?)/i);

  const parsed = {
    paymentNumber: paymentNumberMatch ? paymentNumberMatch[1] : "",
    paymentDateText: paymentDateMatch ? paymentDateMatch[1] : "",
    paymentDate: paymentDateMatch ? parseMcDonaldsPaymentDate_(paymentDateMatch[1]) : new Date(),
    paymentAmount: paymentAmountMatch ? parseMoneyFlexible_(paymentAmountMatch[1]) : 0,
    invoices: []
  };

  const lines = text.split(/\n+/);
  const seen = {};

  lines.forEach(function(line) {
    const entry = parseMcDonaldsRemittanceInvoiceLine_(line);
    if (!entry || seen[entry.invoiceNumber]) return;

    seen[entry.invoiceNumber] = true;
    parsed.invoices.push(entry);
  });

  if (!parsed.invoices.length) {
    throw new Error("No se encontraron lineas de invoices en el texto pegado.");
  }

  return parsed;
}

function parseMcDonaldsRemittanceInvoiceLine_(line) {
  line = String(line || "").trim();
  if (!line) return null;

  const tokens = line.split(/\s+/);
  if (!/^\d{4,}$/.test(tokens[0] || "")) return null;

  const usdIndex = tokens.map(function(t) {
    return String(t || "").toUpperCase();
  }).indexOf("USD");
  if (usdIndex < 1) return null;

  const moneyTokens = [];
  for (let i = usdIndex + 1; i < tokens.length; i++) {
    if (/^[\d,]+(?:\.\d{2})?$/.test(tokens[i])) moneyTokens.push(tokens[i]);
  }

  if (moneyTokens.length < 1) return null;

  let invoiceDateText = "";
  for (let i = usdIndex - 1; i >= 1; i--) {
    if (/^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/.test(tokens[i])) {
      invoiceDateText = tokens[i];
      break;
    }
  }

  return {
    invoiceNumber: String(tokens[0] || "").trim(),
    invoiceDateText: invoiceDateText,
    currency: "USD",
    invoiceAmount: parseMoneyFlexible_(moneyTokens[0]),
    amountWithheld: parseMoneyFlexible_(moneyTokens[1] || 0),
    discountTaken: parseMoneyFlexible_(moneyTokens[2] || 0),
    paidAmount: parseMoneyFlexible_(moneyTokens[moneyTokens.length - 1] || moneyTokens[0])
  };
}

function parseMcDonaldsPaymentDate_(dateText) {
  const parts = String(dateText || "").trim().split(/[-\/]/).map(Number);
  if (parts.length !== 3) return new Date();

  let year = parts[2];
  if (year < 100) year += 2000;

  return new Date(year, parts[1] - 1, parts[0]);
}

function buildMcDonaldsPaymentPreview_(companyId, parsed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);
  if (!shEco) throw new Error("No existe la hoja ECONOMY.");

  const lookup = buildEconomyInvoiceLookup_(shEco, companyId);
  const rows = parsed.invoices.map(function(inv) {
    const match = lookup[normalizeInvoiceNumber_(inv.invoiceNumber)] || null;
    return {
      invoiceNumber: inv.invoiceNumber,
      paidAmount: inv.paidAmount,
      invoiceAmount: inv.invoiceAmount,
      found: !!match,
      rowNumber: match ? match.rowNumber : "",
      currentStatus: match ? match.status : "",
      currentAmount: match ? match.amount : 0,
      woNumber: match ? match.woNumber : "",
      client: match ? match.client : ""
    };
  });

  const matchedCount = rows.filter(function(row) { return row.found; }).length;

  return {
    paymentNumber: parsed.paymentNumber,
    paymentDate: parsed.paymentDateText,
    paymentAmount: parsed.paymentAmount,
    invoiceCount: rows.length,
    matchedCount: matchedCount,
    unmatchedCount: rows.length - matchedCount,
    rows: rows
  };
}

function applyMcDonaldsPaymentToEconomy_(sh, companyId, parsed, session) {
  const headers = ensureSheetColumns_(sh, [
    "DATE_PAID",
    "PAID_AMOUNT",
    "PAYMENT_NUMBER",
    "PAYMENT_DATE",
    "PAYMENT_AMOUNT",
    "PAYMENT_SOURCE",
    "PAYMENT_IMPORTED_AT",
    "PAYMENT_IMPORTED_BY",
    "PAYMENT_IMPORTED_BY_EMAIL",
    "PAYMENT_NOTES"
  ]);
  const lookup = buildEconomyInvoiceLookup_(sh, companyId);
  const importedAt = new Date();
  const actor = getSessionActorLabel_(session);
  const actorEmail = String(session && session.email || "").trim();
  const unmatchedInvoices = [];
  let updatedRows = 0;
  let alreadyPaidRows = 0;

  parsed.invoices.forEach(function(inv) {
    const match = lookup[normalizeInvoiceNumber_(inv.invoiceNumber)];
    if (!match) {
      unmatchedInvoices.push(inv.invoiceNumber);
      return;
    }

    const rowNumber = match.rowNumber;
    const currentStatus = String(getCellByHeader_(sh, rowNumber, headers, "STATUS") || "").trim().toUpperCase();
    if (currentStatus === "PAID") alreadyPaidRows++;

    setCellByHeader_(sh, rowNumber, headers, "STATUS", "PAID");
    setCellByHeader_(sh, rowNumber, headers, "DATE_PAID", parsed.paymentDate || new Date());
    setCellByHeader_(sh, rowNumber, headers, "PAID_AMOUNT", inv.paidAmount);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_NUMBER", parsed.paymentNumber);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_DATE", parsed.paymentDate || new Date());
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_AMOUNT", parsed.paymentAmount);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_SOURCE", "MCDONALDS_REMITTANCE");
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_IMPORTED_AT", importedAt);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_IMPORTED_BY", actor);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_IMPORTED_BY_EMAIL", actorEmail);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_NOTES", buildMcDonaldsPaymentNote_(inv, parsed));
    updatedRows++;
  });

  return {
    updatedRows: updatedRows,
    alreadyPaidRows: alreadyPaidRows,
    unmatchedInvoices: unmatchedInvoices,
    rows: buildMcDonaldsPaymentPreview_(companyId, parsed).rows
  };
}

function applyMcDonaldsPaymentToInvoices_(ss, companyId, parsed, session) {
  const sh = ss.getSheetByName("INVOICES");
  if (!sh || sh.getLastRow() < 2) {
    return {
      updatedRows: 0,
      missingInvoices: parsed.invoices.map(function(inv) { return inv.invoiceNumber; })
    };
  }

  const headers = ensureSheetColumns_(sh, [
    "PAYMENT_STATUS",
    "DATE_PAID",
    "PAID_AMOUNT",
    "PAYMENT_NUMBER",
    "PAYMENT_DATE",
    "PAYMENT_SOURCE",
    "PAYMENT_IMPORTED_AT",
    "PAYMENT_IMPORTED_BY"
  ]);
  const data = sh.getDataRange().getValues();
  const lookup = {};
  const importedAt = new Date();
  const actor = getSessionActorLabel_(session);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = String(getHeaderValueFromRow_(row, headers, "COMPANY_ID") || companyId || CFG.DEFAULT_COMPANY_ID)
      .trim()
      .toUpperCase();
    if (rowCompany !== companyId) continue;

    const invoiceNumber = getInvoiceNumberFromRow_(row, headers);
    const key = normalizeInvoiceNumber_(invoiceNumber);
    if (key && !lookup[key]) lookup[key] = i + 1;
  }

  let updatedRows = 0;
  const missingInvoices = [];

  parsed.invoices.forEach(function(inv) {
    const rowNumber = lookup[normalizeInvoiceNumber_(inv.invoiceNumber)];
    if (!rowNumber) {
      missingInvoices.push(inv.invoiceNumber);
      return;
    }

    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_STATUS", "PAID");
    setCellByHeader_(sh, rowNumber, headers, "DATE_PAID", parsed.paymentDate || new Date());
    setCellByHeader_(sh, rowNumber, headers, "PAID_AMOUNT", inv.paidAmount);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_NUMBER", parsed.paymentNumber);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_DATE", parsed.paymentDate || new Date());
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_SOURCE", "MCDONALDS_REMITTANCE");
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_IMPORTED_AT", importedAt);
    setCellByHeader_(sh, rowNumber, headers, "PAYMENT_IMPORTED_BY", actor);
    updatedRows++;
  });

  return {
    updatedRows: updatedRows,
    missingInvoices: missingInvoices
  };
}

function buildEconomyInvoiceLookup_(sh, companyId) {
  const lookup = {};
  if (!sh || sh.getLastRow() < 2) return lookup;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = String(getHeaderValueFromRow_(row, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID)
      .trim()
      .toUpperCase();
    if (rowCompany !== companyId) continue;

    const invoiceNumber = getInvoiceNumberFromRow_(row, headers);
    const key = normalizeInvoiceNumber_(invoiceNumber);
    if (!key || lookup[key]) continue;

    lookup[key] = {
      rowNumber: i + 1,
      invoiceNumber: invoiceNumber,
      status: String(getHeaderValueFromRow_(row, headers, "STATUS") || "").trim().toUpperCase(),
      amount: parseMoneyFlexible_(getHeaderValueFromRow_(row, headers, "AMOUNT")),
      woNumber: String(getHeaderValueFromRow_(row, headers, "WO_NUMBER") || "").trim(),
      client: String(getHeaderValueFromRow_(row, headers, "CLIENT") || "").trim()
    };
  }

  return lookup;
}

function getInvoiceNumberFromRow_(row, headers) {
  const possibleNames = ["INVOICE_NUMBER", "INVOICE #", "INVOICE", "INVOICE_NO"];

  for (let i = 0; i < possibleNames.length; i++) {
    const value = getHeaderValueFromRow_(row, headers, possibleNames[i]);
    if (value !== "" && value !== null && value !== undefined) return value;
  }

  return "";
}

function normalizeInvoiceNumber_(value) {
  return String(value || "").replace(/[^\dA-Za-z-]/g, "").trim().toUpperCase();
}

function buildMcDonaldsPaymentNote_(invoiceEntry, parsed) {
  return [
    "McDonalds remittance",
    "Payment " + (parsed.paymentNumber || ""),
    "Invoice " + invoiceEntry.invoiceNumber,
    "Paid $" + Number(invoiceEntry.paidAmount || 0).toFixed(2)
  ].join(" | ");
}

function getEconomyRowValueByHeader_(sh, rowNumber, headers, possibleNames) {
  const normalized = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  });

  for (let i = 0; i < possibleNames.length; i++) {
    const idx = normalized.indexOf(String(possibleNames[i] || "").trim().toUpperCase());
    if (idx >= 0) return sh.getRange(rowNumber, idx + 1).getValue();
  }

  return "";
}

function getEconomyRowNumberByHeader_(sh, rowNumber, headers, possibleNames) {
  return parseMoneyFlexible_(getEconomyRowValueByHeader_(sh, rowNumber, headers, possibleNames));
}

function calculateEconomyRealPartsCost_(source, cost) {
  source = String(source || "").trim().toUpperCase();
  cost = Number(cost || 0);

  return source === "YOEL" ? cost : 0;
}

function syncInvoicesToEconomy(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  let session = null;
  if (sessionToken) {
    session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shInv = ss.getSheetByName("INVOICES");
  const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);
  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  const shUsers = ss.getSheetByName(CFG.SHEET_USERS);

  if (!shInv || !shEco || !shWO || !shUsers) {
    throw new Error("Faltan hojas INVOICES, ECONOMY, WORK_ORDERS o USERS.");
  }

  ensureSheetColumns_(shEco, [
    "PERIOD_MONTH",
    "PERIOD_YEAR",
    "PERIOD_LABEL",
    "ACCOUNTING_PERIOD",
    "CLOSED_PERIOD",
    "MONTH_CLOSED_AT",
    "MONTH_CLOSED_BY",
    "MONTH_CLOSED_BY_EMAIL"
  ]);

  const invData = shInv.getDataRange().getValues();
  const woData = shWO.getDataRange().getValues();
  const userData = shUsers.getDataRange().getValues();

  let ecoData = shEco.getDataRange().getValues();

  if (invData.length < 2) return true;

  const invHeaders = invData[0].map(h => String(h).trim());
  let ecoHeaders = ecoData[0].map(h => String(h).trim());
  const woHeaders = woData[0].map(h => String(h).trim());
  const userHeaders = userData[0].map(h => String(h).trim());

  const idxInvWO = invHeaders.indexOf("WO_NUMBER");

  if (idxInvWO === -1) {
    throw new Error("INVOICES debe tener WO_NUMBER.");
  }

  const ecoWO = ecoHeaders.indexOf("WO_NUMBER");
  const ecoCompany = ecoHeaders.indexOf("COMPANY_ID");

  if (ecoWO === -1 || ecoCompany === -1) {
    throw new Error("ECONOMY debe tener WO_NUMBER y COMPANY_ID.");
  }

  const woWO = woHeaders.indexOf("WO_NUMBER");
  const woCompany = woHeaders.indexOf("COMPANY_ID");
  const woTech = woHeaders.indexOf("TECHNICIAN");
  const woClient = woHeaders.indexOf("CLIENT");
  const woNSN = woHeaders.indexOf("NSN");
  const woCompleted = woHeaders.indexOf("DATE_COMPLETED");
  const woStatus = woHeaders.indexOf("STATUS");
  const woClosed = woHeaders.indexOf("DATE_CLOSED");

  const userName = userHeaders.indexOf("NAME");
  const userCompany = userHeaders.indexOf("COMPANY_ID");
  const userRate = userHeaders.indexOf("HOURLY_RATE");

  if (woWO === -1 || woCompany === -1 || woTech === -1) {
    throw new Error("WORK_ORDERS debe tener WO_NUMBER, COMPANY_ID y TECHNICIAN.");
  }

  if (userName === -1 || userCompany === -1 || userRate === -1) {
    throw new Error("USERS debe tener NAME, COMPANY_ID y HOURLY_RATE.");
  }

  const woMap = {};

  for (let i = 1; i < woData.length; i++) {
    const wo = String(woData[i][woWO] || "").trim();
    const comp = String(woData[i][woCompany] || "").trim().toUpperCase();

    if (!wo || comp !== companyId) continue;

    woMap[wo] = {
      rowNumber: i + 1,
      companyId: comp,
      technician: String(woData[i][woTech] || "").trim(),
      client: woClient >= 0 ? woData[i][woClient] : "",
      nsn: woNSN >= 0 ? woData[i][woNSN] : "",
      dateCompleted: woCompleted >= 0 ? woData[i][woCompleted] : "",
      status: woStatus >= 0 ? String(woData[i][woStatus] || "").trim().toUpperCase() : "",
      dateClosed: woClosed >= 0 ? woData[i][woClosed] : ""
    };
  }

  const rateMap = {};

  for (let i = 1; i < userData.length; i++) {
    const name = String(userData[i][userName] || "").trim();
    const comp = String(userData[i][userCompany] || "").trim().toUpperCase();
    const rate = Number(userData[i][userRate] || 0);

    if (name && comp === companyId) {
      rateMap[name.toLowerCase()] = rate;
    }
  }

  const period = getCurrentEconomyPeriod();
  let syncedRows = 0;
  let createdRows = 0;
  let skippedClosedRows = 0;

  for (let i = 1; i < invData.length; i++) {
    const invRow = invData[i];
    const wo = String(invRow[idxInvWO] || "").trim();

    const idxInvoiceType = invHeaders.indexOf("INVOICE_TYPE");
const idxWOType = invHeaders.indexOf("WO_TYPE");
const idxWOTypo = invHeaders.indexOf("WO_TYPO");
const idxPMType = invHeaders.indexOf("PM_TYPE");

const invoiceType = idxInvoiceType >= 0 ? String(invRow[idxInvoiceType] || "").trim().toUpperCase() : "";
const woType = idxWOType >= 0 ? String(invRow[idxWOType] || "").trim().toUpperCase() : "";
const woTypo = idxWOTypo >= 0 ? String(invRow[idxWOTypo] || "").trim().toUpperCase() : "";
const pmType = idxPMType >= 0 ? String(invRow[idxPMType] || "").trim().toUpperCase() : "";

if (
  invoiceType === "PM" ||
  woType === "PM_FORM" ||
  woTypo === "PM_FORM" ||
  pmType
) {
  continue;
}

    const invType = String(
  invRow[invHeaders.indexOf("INVOICE_TYPE")] || ""
).trim().toUpperCase();

if (invType === "PM") continue;

    if (!wo) continue;
    if (!woMap[wo]) continue;

    const total = getInvoiceTotalValue_(invRow, invHeaders);
    const materialCost = getMaterialCostValue_(invRow, invHeaders);
    const horas = getHoursValue_(invRow, invHeaders);
    const tax = getTaxAmountValue_(invRow, invHeaders);
    const invoiceNumber = getInvoiceNumberValue_(invRow, invHeaders);
    const invoiceDate = getInvoiceDateValue_(invRow, invHeaders);

    const techniciansText = woMap[wo].technician || "";

    const technicians = techniciansText
      .split(",")
      .map(t => t.trim())
      .filter(Boolean);

    const hoursPerTech = technicians.length ? horas / technicians.length : 0;

    let laborTotal = 0;

    const detail = technicians.map(function(name) {
      const rate = Number(rateMap[name.toLowerCase()] || 0);
      const pay = hoursPerTech * rate;
      laborTotal += pay;

      return name + ": " + hoursPerTech.toFixed(2) + "h x $" + rate.toFixed(2) + " = $" + pay.toFixed(2);
    }).join(" | ");

    ecoData = shEco.getDataRange().getValues();
    ecoHeaders = ecoData[0].map(h => String(h).trim());

    let targetRow = -1;
    let foundClosedMatch = false;

    for (let j = 1; j < ecoData.length; j++) {
      if (isSoftDeletedRow_(ecoData[j], ecoHeaders)) continue;

      const ecoWo = String(ecoData[j][ecoWO] || "").trim();
      const ecoComp = String(ecoData[j][ecoCompany] || "").trim().toUpperCase();

      if (ecoWo === wo && ecoComp === companyId) {
        if (isEconomyRowClosedForSync_(ecoData[j], ecoHeaders, period.label)) {
          foundClosedMatch = true;
          continue;
        }

        targetRow = j + 1;
        break;
      }
    }

    if (foundClosedMatch && targetRow === -1) {
      skippedClosedRows++;
      if (!shouldCreateEconomyRowForClosedMatch_(invoiceDate, period)) {
        continue;
      }
    }

    if (targetRow === -1) {
      const rowObj = {
        COMPANY_ID: companyId,
        WO_NUMBER: wo,
        CLIENT: woMap[wo].client || "",
        NSN: woMap[wo].nsn || "",
        DATE_COMPLETED: woMap[wo].dateCompleted || "",
        STATUS: "PENDING"
      };

      const newRow = ecoHeaders.map(function(h) {
        return rowObj[h] !== undefined ? rowObj[h] : "";
      });

      shEco.appendRow(newRow);
      targetRow = shEco.getLastRow();
      createdRows++;
    }

    setEcoValue_(shEco, targetRow, ecoHeaders, ["AMOUNT", "INVOICE_TOTAL"], total);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["TAX", "TAX_AMOUNT"], tax);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["COST", "MATERIAL_COST"], materialCost);

    const existingInvSource = getCellByHeader_(shEco, targetRow, ecoHeaders, "INV_SOURCE");
    const realMaterialCost = calculateEconomyRealPartsCost_(existingInvSource, materialCost);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["PROFIT"], total - tax - realMaterialCost - laborTotal);

    setEcoValue_(shEco, targetRow, ecoHeaders, ["INVOICE_NUMBER"], invoiceNumber);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["DATE_INVOICE"], invoiceDate);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["STATUS"], "INVOICED");
    closeWorkOrderAfterEconomyInvoiceSync_(shWO, woHeaders, woMap[wo], invoiceNumber, invoiceDate);

    setEcoValue_(shEco, targetRow, ecoHeaders, ["HORAS", "LABOR_HOURS"], horas);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["TECHNICIANS"], techniciansText);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["TECH_PAY_DETAIL"], detail);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["TECH_LABOR_COST", "TECH_LABOR_PAY"], laborTotal);

    setEcoValue_(shEco, targetRow, ecoHeaders, ["PERIOD_MONTH"], period.month);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["PERIOD_YEAR"], period.year);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["PERIOD_LABEL"], period.label);
    setEcoValue_(shEco, targetRow, ecoHeaders, ["ACCOUNTING_PERIOD"], period.label);
    syncedRows++;
  }

  const repairResult = repairEconomyRowsAfterClose_(shEco, companyId, session, period);

  if (session) {
    addAuditLog_("ECONOMY", "INVOICES_SYNCED_TO_ECONOMY", companyId, "ECONOMY", period.label, session, {
      syncedRows: syncedRows,
      createdRows: createdRows,
      skippedClosedRows: skippedClosedRows,
      repairedRows: repairResult.repairedRows,
      duplicateRowsRemoved: repairResult.duplicateRowsRemoved,
      period: period.label
    });
  }

  return {
    syncedRows: syncedRows,
    createdRows: createdRows,
    skippedClosedRows: skippedClosedRows,
    repairedRows: repairResult.repairedRows,
    duplicateRowsRemoved: repairResult.duplicateRowsRemoved,
    period: period.label
  };
}

function isEconomyRowClosedForSync_(row, headers, currentPeriodLabel) {
  const closedPeriod = String(getHeaderValueFromRow_(row, headers, "CLOSED_PERIOD") || "").trim();
  const closedAt = getHeaderValueFromRow_(row, headers, "MONTH_CLOSED_AT");
  if (closedPeriod || closedAt) return true;

  const rowPeriod = getEconomyRowPeriodLabelFromData_(row, headers);
  currentPeriodLabel = String(currentPeriodLabel || "").trim();
  return !!(rowPeriod && currentPeriodLabel && rowPeriod !== currentPeriodLabel);
}

function repairEconomyDuplicatesAfterClose(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_ECONOMY);
  if (!sh) throw new Error("No existe la hoja ECONOMY.");

  const period = getCurrentEconomyPeriod();
  const result = repairEconomyRowsAfterClose_(sh, companyId, session, period);

  addAuditLog_("ECONOMY", "ECONOMY_DUPLICATES_REPAIRED", companyId, "ECONOMY", period.label, session, result);
  return result;
}

function repairEconomyRowsAfterClose_(sh, companyId, session, period) {
  const emptyResult = {
    repairedRows: 0,
    duplicateRowsRemoved: 0,
    oldCopiesRemoved: 0,
    period: period && period.label || ""
  };

  period = period || {};
  const currentPeriodLabel = String(period.label || "").trim();
  const lastClosedAt = period.lastClosedAt;
  const lastClosedPeriod = String(period.lastClosedPeriod || "").trim();

  if (!currentPeriodLabel || !lastClosedAt || sh.getLastRow() < 2) return emptyResult;

  let headers = ensureSheetColumns_(sh, [
    "ACTIVE",
    "DELETED_AT",
    "DELETED_BY",
    "DELETED_REASON"
  ]);

  const data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  headers = data[0].map(function(h) {
    return String(h || "").trim();
  });

  const closedKeys = {};
  const currentRows = [];
  const rowsToDelete = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = String(getHeaderValueFromRow_(row, headers, "COMPANY_ID") || companyId).trim().toUpperCase();
    if (rowCompany !== companyId) continue;

    const key = getEconomyDedupeKeyFromRow_(row, headers, companyId);
    if (!key) continue;

    const rowNumber = i + 1;
    const rowPeriod = getEconomyRowPeriodLabelFromData_(row, headers);
    const closedPeriod = String(getHeaderValueFromRow_(row, headers, "CLOSED_PERIOD") || "").trim();
    const closedAt = getHeaderValueFromRow_(row, headers, "MONTH_CLOSED_AT");
    const isClosedRow = !!(closedPeriod || closedAt || (lastClosedPeriod && rowPeriod === lastClosedPeriod));

    if (isClosedRow && rowPeriod !== currentPeriodLabel) {
      closedKeys[key] = true;
      continue;
    }

    if (rowPeriod === currentPeriodLabel) {
      currentRows.push({
        rowNumber: rowNumber,
        row: row,
        key: key
      });
    }
  }

  currentRows.forEach(function(item) {
    const invoiceDate = getHeaderValueFromRow_(item.row, headers, "DATE_INVOICE");
    if (closedKeys[item.key] && !shouldCreateEconomyRowForClosedMatch_(invoiceDate, period)) {
      rowsToDelete[item.rowNumber] = "Old invoice copied into current closed period sync";
    }
  });

  const groups = {};
  currentRows.forEach(function(item) {
    if (rowsToDelete[item.rowNumber]) return;
    if (!groups[item.key]) groups[item.key] = [];
    groups[item.key].push(item);
  });

  Object.keys(groups).forEach(function(key) {
    const group = groups[key];
    if (group.length <= 1) return;

    const keep = chooseEconomyRowToKeep_(group, headers);
    group.forEach(function(item) {
      if (item.rowNumber !== keep.rowNumber) {
        rowsToDelete[item.rowNumber] = "Duplicate current economy row";
      }
    });
  });

  const actor = getSessionActorLabel_(session);
  const rowNumbers = Object.keys(rowsToDelete).map(Number).sort(function(a, b) { return a - b; });

  rowNumbers.forEach(function(rowNumber) {
    setCellByHeader_(sh, rowNumber, headers, "ACTIVE", "NO");
    setCellByHeader_(sh, rowNumber, headers, "DELETED_AT", new Date());
    setCellByHeader_(sh, rowNumber, headers, "DELETED_BY", actor || "repairEconomyRowsAfterClose");
    setCellByHeader_(sh, rowNumber, headers, "DELETED_REASON", rowsToDelete[rowNumber]);
  });

  const oldCopiesRemoved = rowNumbers.filter(function(rowNumber) {
    return String(rowsToDelete[rowNumber] || "").indexOf("Old invoice") === 0;
  }).length;

  return {
    repairedRows: rowNumbers.length,
    duplicateRowsRemoved: rowNumbers.length - oldCopiesRemoved,
    oldCopiesRemoved: oldCopiesRemoved,
    period: currentPeriodLabel
  };
}

function shouldCreateEconomyRowForClosedMatch_(invoiceDate, period) {
  period = period || {};
  const closeDate = normalizeEconomyDateForCompare_(period.lastClosedAt);
  if (!closeDate) return false;

  const invDate = normalizeEconomyDateForCompare_(invoiceDate);
  if (!invDate) return false;

  return invDate.getTime() > closeDate.getTime();
}

function getEconomyDedupeKeyFromRow_(row, headers, fallbackCompanyId) {
  const companyId = String(getHeaderValueFromRow_(row, headers, "COMPANY_ID") || fallbackCompanyId || "")
    .trim()
    .toUpperCase();
  const invoiceNumber = normalizeInvoiceNumber_(getHeaderValueFromRow_(row, headers, "INVOICE_NUMBER"));
  const woNumber = String(getHeaderValueFromRow_(row, headers, "WO_NUMBER") || "").trim().toUpperCase();

  if (invoiceNumber) return companyId + "|INV|" + invoiceNumber;
  if (woNumber) return companyId + "|WO|" + woNumber;
  return "";
}

function chooseEconomyRowToKeep_(group, headers) {
  let best = group[0];
  let bestScore = scoreEconomyRowForKeep_(best, headers);

  for (let i = 1; i < group.length; i++) {
    const score = scoreEconomyRowForKeep_(group[i], headers);
    if (score > bestScore) {
      best = group[i];
      bestScore = score;
    }
  }

  return best;
}

function scoreEconomyRowForKeep_(item, headers) {
  const row = item.row;
  let score = Number(item.rowNumber || 0);

  const invSource = String(getHeaderValueFromRow_(row, headers, "INV_SOURCE") || "").trim();
  if (invSource) score += 1000000;

  const status = String(getHeaderValueFromRow_(row, headers, "STATUS") || "").trim().toUpperCase();
  if (status === "PAID") score += 200000;
  if (status === "INVOICED") score += 100000;

  return score;
}

function normalizeEconomyDateForCompare_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value;
  }

  const text = String(value || "").trim();
  if (!text) return null;

  const direct = new Date(text);
  if (!isNaN(direct.getTime())) return direct;

  const match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
  if (!match) return null;

  let first = Number(match[1]);
  let second = Number(match[2]);
  let year = Number(match[3]);
  if (year < 100) year += 2000;

  const month = first > 12 ? second : first;
  const day = first > 12 ? first : second;
  const parsed = new Date(year, month - 1, day);

  return isNaN(parsed.getTime()) ? null : parsed;
}

function getHeaderValueFromRow_(row, headers, headerName) {
  const normalizedHeaderName = String(headerName || "").trim().toUpperCase();
  const idx = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf(normalizedHeaderName);

  return idx >= 0 ? row[idx] : "";
}

function closeWorkOrderAfterEconomyInvoiceSync_(shWO, woHeaders, woInfo, invoiceNumber, invoiceDate) {
  if (!shWO || !woInfo || !woInfo.rowNumber || !invoiceNumber) return;
  if (woInfo.status === "CLOSED") return;

  const now = new Date();
  const closedDate = woInfo.dateClosed || invoiceDate || now;

  setWorkOrderValueByHeader_(shWO, woInfo.rowNumber, woHeaders, "STATUS", "CLOSED");
  setWorkOrderValueByHeader_(shWO, woInfo.rowNumber, woHeaders, "DATE_CLOSED", closedDate);
  setWorkOrderValueByHeader_(shWO, woInfo.rowNumber, woHeaders, "INVOICE_NUMBER", invoiceNumber);
  setWorkOrderValueByHeader_(shWO, woInfo.rowNumber, woHeaders, "DATE_INVOICE", invoiceDate || now);

  addLog_(
    woInfo.companyId,
    woHeaders.indexOf("WO_NUMBER") >= 0
      ? shWO.getRange(woInfo.rowNumber, woHeaders.indexOf("WO_NUMBER") + 1).getValue()
      : "",
    "ORDER CLOSED AFTER ECONOMY INVOICE SYNC",
    woInfo.status || "",
    "CLOSED",
    Session.getActiveUser().getEmail() || "Economy Sync",
    "Invoice " + invoiceNumber
  );

  woInfo.status = "CLOSED";
  woInfo.dateClosed = closedDate;
}

function setWorkOrderValueByHeader_(sh, rowNumber, headers, headerName, value) {
  const col = headers.indexOf(headerName) + 1;
  if (col > 0) {
    sh.getRange(rowNumber, col).setValue(value);
  }
}

function setEcoValue_(sh, rowNumber, headers, possibleNames, value) {
  for (let i = 0; i < possibleNames.length; i++) {
    const col = headers.indexOf(possibleNames[i]) + 1;
    if (col > 0) {
      sh.getRange(rowNumber, col).setValue(value);
    }
  }
}
function getCurrentEconomyPeriod() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_SETTINGS");
  if (!sh) throw new Error("No existe la hoja APP_SETTINGS.");

  const data = sh.getDataRange().getValues();

  let month = 0;
  let year = 0;
  let lastClosedPeriod = "";
  let lastClosedAt = "";

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || "").trim();
    const value = String(data[i][1] || "").trim();

    if (key === "ECONOMY_MONTH") month = Number(value);
    if (key === "ECONOMY_YEAR") year = Number(value);
    if (key === "ECONOMY_LAST_CLOSED_PERIOD") lastClosedPeriod = value;
    if (key === "ECONOMY_LAST_CLOSED_AT") lastClosedAt = data[i][1];
  }

  if (!month || !year) {
    throw new Error("APP_SETTINGS no tiene ECONOMY_MONTH / ECONOMY_YEAR");
  }

  return normalizeOpenEconomyPeriod_(month, year, lastClosedPeriod, lastClosedAt);
}

function normalizeOpenEconomyPeriod_(month, year, lastClosedPeriod, lastClosedAt) {
  month = Number(month || 0);
  year = Number(year || 0);
  let label = buildEconomyPeriodLabel_(year, month);
  lastClosedPeriod = String(lastClosedPeriod || "").trim();

  if (lastClosedPeriod && label && label <= lastClosedPeriod) {
    const next = getNextEconomyPeriodAfterLabel_(lastClosedPeriod);
    if (next) {
      next.lastClosedPeriod = lastClosedPeriod;
      next.lastClosedAt = lastClosedAt || "";
      return next;
    }
  }

  return {
    month: month,
    year: year,
    label: label,
    lastClosedPeriod: lastClosedPeriod,
    lastClosedAt: lastClosedAt || ""
  };
}

function buildEconomyPeriodLabel_(year, month) {
  year = Number(year || 0);
  month = Number(month || 0);
  if (!year || !month) return "";
  return year + "-" + String(month).padStart(2, "0");
}

function getNextEconomyPeriodAfterLabel_(label) {
  label = String(label || "").trim();
  const match = label.match(/^(\d{4})-(\d{1,2})$/);
  if (!match) return null;

  let year = Number(match[1]);
  let month = Number(match[2]) + 1;

  if (month > 12) {
    month = 1;
    year++;
  }

  return {
    month: month,
    year: year,
    label: buildEconomyPeriodLabel_(year, month)
  };
}

function closeEconomyMonth(sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], getDashboardCompanyIdFromSession_(sessionToken));
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_SETTINGS");
    if (!sh) throw new Error("No existe la hoja APP_SETTINGS.");

    const data = sh.getDataRange().getValues();
    let monthRow = -1;
    let yearRow = -1;
    let lastClosedPeriod = "";
    let lastClosedAt = "";

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || "").trim();

      if (key === "ECONOMY_MONTH") monthRow = i + 1;
      if (key === "ECONOMY_YEAR") yearRow = i + 1;
      if (key === "ECONOMY_LAST_CLOSED_PERIOD") lastClosedPeriod = String(data[i][1] || "").trim();
      if (key === "ECONOMY_LAST_CLOSED_AT") lastClosedAt = data[i][1];
    }

    if (monthRow === -1 || yearRow === -1) {
      throw new Error("No se encontro configuracion ECONOMY_MONTH / YEAR.");
    }

    let month = Number(sh.getRange(monthRow, 2).getValue());
    let year = Number(sh.getRange(yearRow, 2).getValue());

    if (!month || !year || month < 1 || month > 12) {
      throw new Error("Periodo economico invalido en APP_SETTINGS.");
    }

    const openPeriod = normalizeOpenEconomyPeriod_(month, year, lastClosedPeriod, lastClosedAt);
    month = openPeriod.month;
    year = openPeriod.year;

    const previousLabel = buildEconomyPeriodLabel_(year, month);
    const closedAt = new Date();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);
    if (!shEco) throw new Error("No existe la hoja ECONOMY.");

    const closedRows = closeEconomyRowsForPeriod_(shEco, previousLabel, closedAt, session);
    const economyMarkerRow = appendMonthCloseMarkerRow_(shEco, previousLabel, closedAt, session);

    const shInv = ss.getSheetByName("INVOICES");
    const invoiceMarkerRow = shInv
      ? appendMonthCloseMarkerRow_(shInv, previousLabel, closedAt, session)
      : 0;

    month++;

    if (month > 12) {
      month = 1;
      year++;
    }

    const newLabel = year + "-" + String(month).padStart(2, "0");

    sh.getRange(monthRow, 2).setValue(month);
    sh.getRange(yearRow, 2).setValue(year);
    setAppSettingValue_(sh, "ECONOMY_LAST_CLOSED_PERIOD", previousLabel);
    setAppSettingValue_(sh, "ECONOMY_LAST_CLOSED_AT", closedAt);
    setAppSettingValue_(sh, "ECONOMY_LAST_CLOSED_BY", getSessionActorLabel_(session));
    SpreadsheetApp.flush();

    addAuditLog_("ECONOMY", "ECONOMY_MONTH_CLOSED", session.companyId || CFG.DEFAULT_COMPANY_ID, "APP_SETTINGS", "ECONOMY_PERIOD", session, {
      previousLabel: previousLabel,
      newMonth: month,
      newYear: year,
      label: newLabel,
      closedRows: closedRows,
      economyMarkerRow: economyMarkerRow,
      invoiceMarkerRow: invoiceMarkerRow
    });

    return {
      month: month,
      year: year,
      previousLabel: previousLabel,
      label: newLabel,
      closedRows: closedRows,
      economyMarkerRow: economyMarkerRow,
      invoiceMarkerRow: invoiceMarkerRow
    };
  } catch (err) {
    notifySystemError_("ECONOMY_CLOSE_MONTH_ERROR", err, {
      module: "ECONOMY",
      actorEmail: session.email || "",
      actorRole: session.role || ""
    });
    throw err;
  } finally {
    if (locked) lock.releaseLock();
  }
}

function appendMonthCloseMarkerRow_(sh, periodLabel, closedAt, session) {
  if (!sh || sh.getLastColumn() < 1) return 0;

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  if (lastRow >= 2) {
    const lastValues = sh.getRange(lastRow, 1, 1, lastCol).getValues()[0];
    if (isMonthCloseMarkerRow_(lastValues)) return lastRow;
  }

  const marker = getMonthCloseMarkerText_();
  const values = Array(lastCol).fill(marker);
  sh.appendRow(values);

  const rowNumber = sh.getLastRow();
  const range = sh.getRange(rowNumber, 1, 1, lastCol);
  range
    .setBackground("#111827")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange(rowNumber, 1).setNote([
    "Period: " + String(periodLabel || ""),
    "Closed at: " + String(closedAt || ""),
    "Closed by: " + getSessionActorLabel_(session)
  ].join("\n"));

  return rowNumber;
}

function closeEconomyRowsForPeriod_(sh, periodLabel, closedAt, session) {
  periodLabel = String(periodLabel || "").trim();
  if (!periodLabel) return 0;

  const headers = ensureSheetColumns_(sh, [
    "ACCOUNTING_PERIOD",
    "CLOSED_PERIOD",
    "MONTH_CLOSED_AT",
    "MONTH_CLOSED_BY",
    "MONTH_CLOSED_BY_EMAIL"
  ]);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const data = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  const actor = getSessionActorLabel_(session);
  const actorEmail = String(session && session.email || "").trim();
  let closed = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowPeriod = getEconomyRowPeriodLabelFromData_(row, headers);
    if (rowPeriod !== periodLabel) continue;

    const rowNumber = i + 1;
    const existingAccountingPeriod = getCellByHeader_(sh, rowNumber, headers, "ACCOUNTING_PERIOD");
    if (!existingAccountingPeriod) {
      setCellByHeader_(sh, rowNumber, headers, "ACCOUNTING_PERIOD", periodLabel);
    }

    setCellByHeader_(sh, rowNumber, headers, "CLOSED_PERIOD", periodLabel);
    setCellByHeader_(sh, rowNumber, headers, "MONTH_CLOSED_AT", closedAt);
    setCellByHeader_(sh, rowNumber, headers, "MONTH_CLOSED_BY", actor);
    setCellByHeader_(sh, rowNumber, headers, "MONTH_CLOSED_BY_EMAIL", actorEmail);
    closed++;
  }

  return closed;
}

function getEconomyRowPeriodLabelFromData_(row, headers) {
  const idxLabel = headers.indexOf("PERIOD_LABEL");
  const label = idxLabel >= 0 ? String(row[idxLabel] || "").trim() : "";
  if (label) return label;

  const idxAccountingPeriod = headers.indexOf("ACCOUNTING_PERIOD");
  const accountingPeriod = idxAccountingPeriod >= 0 ? String(row[idxAccountingPeriod] || "").trim() : "";
  if (accountingPeriod) return accountingPeriod;

  const idxYear = headers.indexOf("PERIOD_YEAR");
  const idxMonth = headers.indexOf("PERIOD_MONTH");
  const year = idxYear >= 0 ? Number(row[idxYear] || 0) : 0;
  const month = idxMonth >= 0 ? Number(row[idxMonth] || 0) : 0;

  if (year && month) {
    return year + "-" + String(month).padStart(2, "0");
  }

  return "";
}

function setAppSettingValue_(sh, key, value) {
  key = String(key || "").trim();
  if (!key) return;

  const lastRow = Math.max(sh.getLastRow(), 1);
  const data = sh.getRange(1, 1, lastRow, Math.max(sh.getLastColumn(), 2)).getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  sh.appendRow([key, value]);
}

function removePMFromRegularEconomy(sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_ECONOMY);
  if (!sh) throw new Error("No existe ECONOMY.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return 0;

  let headers = ensureSheetColumns_(sh, ["ACTIVE", "DELETED_AT", "DELETED_BY"]);

  const idxInvSource = headers.indexOf("INV_SOURCE");
  const idxNotes = headers.indexOf("NOTES");

  let removed = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const invSource = idxInvSource >= 0 ? String(data[i][idxInvSource] || "").toUpperCase() : "";
    const notes = idxNotes >= 0 ? String(data[i][idxNotes] || "").toUpperCase() : "";

    if (
      invSource === "PM" ||
      notes.includes("PM INVOICE TOTAL") ||
      notes.includes("PAID TO D&D")
    ) {
      setCellByHeader_(sh, i + 1, headers, "ACTIVE", "NO");
      setCellByHeader_(sh, i + 1, headers, "DELETED_AT", new Date());
      setCellByHeader_(sh, i + 1, headers, "DELETED_BY", "removePMFromRegularEconomy");
      removed++;
    }
  }

  addAuditLog_("ECONOMY", "PM_ROWS_SOFT_DELETED_FROM_ECONOMY", session.companyId || CFG.DEFAULT_COMPANY_ID, "ECONOMY", "PM_ROWS", session, {
    removed: removed
  });

  return removed;
}

function getDashboardCompanyIdFromSession_(sessionToken) {
  const session = getSession_(sessionToken);
  if (!session) return "";
  return String(session.companyId || "").trim().toUpperCase();
}
