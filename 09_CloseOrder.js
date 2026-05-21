// =====================================================
// FILE: 09_CloseOrder.gs
// =====================================================

function getCloseOrderData(rowNumber, woNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe WORK_ORDERS.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error("WORK_ORDERS vacío.");

  const headers = data[0].map(h => String(h).trim());

  function clean(v) {
    return String(v || "").replace(/\u2013|\u2014/g, "-").trim().toUpperCase();
  }

  function safeValue(v) {
    if (v instanceof Date) {
      return Utilities.formatDate(v, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a");
    }
    return v === undefined || v === null ? "" : String(v);
  }

  const targetWO = clean(woNumber);
  let foundRow = null;
  let realRowNumber = null;

  for (let i = 1; i < data.length; i++) {
    for (let c = 0; c < data[i].length; c++) {
      if (clean(data[i][c]) === targetWO) {
        foundRow = data[i];
        realRowNumber = i + 1;
        break;
      }
    }
    if (foundRow) break;
  }

  if (!foundRow) {
    throw new Error("No se encontró la orden con WO: " + woNumber);
  }

  const obj = {};
  headers.forEach(function(h, i) {
    obj[h] = safeValue(foundRow[i]);
  });

  obj.WO_NUMBER = obj.WO_NUMBER || woNumber;
  obj.ROW_NUMBER = realRowNumber;
  obj.WO_TYPE = obj.WO_TYPE || "REPAIR_FORM";
  obj.PM_TYPE = obj.PM_TYPE || "";

  enrichWorkOrderObject_(obj);

  const normalizedAddress = normalizeCloseOrderAddress_(obj);

  obj.STORE_ADDRESS = normalizedAddress.full;
  obj.STORE_STREET = normalizedAddress.street;
  obj.STORE_CITY = normalizedAddress.city;
  obj.STORE_STATE = normalizedAddress.state;
  obj.STORE_ZIP = normalizedAddress.zip;

  Object.keys(obj).forEach(function(k) {
    obj[k] = safeValue(obj[k]);
  });

  return obj;
}

function saveCloseOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shInv = ss.getSheetByName("INVOICES");
  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  if (!shInv) throw new Error("No existe la hoja INVOICES.");
  if (!shWO) throw new Error("No existe la hoja WORK_ORDERS.");

  const rowNumber = Number(data.ROW_NUMBER);
  if (!rowNumber || rowNumber < 2) throw new Error("Fila inválida.");

  const companyId = String(data.COMPANY_ID || "").trim().toUpperCase();
  const woNumber = String(data.WO_NUMBER || "").trim();
  const woType = String(data.WO_TYPE || "REPAIR_FORM").trim().toUpperCase();

  if (woType === "PM_FORM") {
    const result = createPMInvoiceFromCloseOrder_(data);

    return {
      success: true,
      invoiceNumber: result.invoiceNumber || "",
      woNumber: woNumber,
      invoiceType: "PM",
      message: "PM invoice created and returned from saveCloseOrder"
    };
  }

  const hours = Number(data.LABOR_HOURS || data.HORAS || data.HOURS || 0);
  const partsTotal = Number(data.MATERIAL_COST || data.PARTS_TOTAL || data.INVERSION || data.COMPRA || data.PARTS || 0);

  const normalizedAddress = normalizeCloseOrderAddress_(data);

  const firstHourRate = 200;
  const additionalHourRate = 130;
  const taxRate = 0.07;

  let laborTotal = 0;
  let labor1Qty = 0;
  let labor2Qty = 0;
  let labor1Amount = 0;
  let labor2Amount = 0;

  if (hours > 0) {
    labor1Qty = 1;
    labor2Qty = Math.max(hours - 1, 0);
    labor1Amount = firstHourRate;
    labor2Amount = labor2Qty * additionalHourRate;
    laborTotal = labor1Amount + labor2Amount;
  }

  const subTotal = partsTotal + laborTotal;
  const tax = subTotal * taxRate;
  const total = subTotal + tax;

  const invoiceNumber = generateInvoiceNumber_(companyId);

  const invHeaders = shInv
    .getRange(1, 1, 1, shInv.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  const invoiceRow = {
    WO_NUMBER: woNumber,

    WO_TYPE: woType,
    WO_TYPO: woType,
    INVOICE_TYPE: "REPAIR",

    INVOICE_NUMBER: invoiceNumber,
    Invoice: invoiceNumber,

    DATE_INVOICE: new Date(),
    Timestamp: new Date(),

    INVOICE_TOTAL: total,
    TOTAL: total,

    LABOR_HOURS: hours,
    HORAS: hours,

    NS: data.NSN || "",

    Source: data.PURCHASE_LOCATION || "",
    MATERIAL_COST: partsTotal,
    PARTS: partsTotal,
    COMPRA: partsTotal,
    INVERSION: partsTotal,

    "TECHNICIAN EMAIL": data.TECHNICIAN_EMAIL || "",
    "TECHNICIAN NAME": data.TECHNICIAN || "",

    PROCESO: "RECIBIDO",
    WORK_PERFORMED: data.WORK_PERFORMED || "",

    LABOR_AMOUNT: laborTotal,
    LABOR: laborTotal,

    TAX_AMOUNT: tax,
    TAX: tax,

    I1_QTY: data.I1_QTY || "",
    I1_PART: data.I1_PART || "",
    I1_DESC: data.I1_DESC || "",
    I1_UNIT: data.I1_UNIT || "",
    I1_AMOUNT: data.I1_AMOUNT || "",

    I2_QTY: data.I2_QTY || "",
    I2_PART: data.I2_PART || "",
    I2_DESC: data.I2_DESC || "",
    I2_UNIT: data.I2_UNIT || "",
    I2_AMOUNT: data.I2_AMOUNT || "",

    I3_QTY: data.I3_QTY || "",
    I3_PART: data.I3_PART || "",
    I3_DESC: data.I3_DESC || "",
    I3_UNIT: data.I3_UNIT || "",
    I3_AMOUNT: data.I3_AMOUNT || "",

    I4_QTY: data.I4_QTY || "",
    I4_PART: data.I4_PART || "",
    I4_DESC: data.I4_DESC || "",
    I4_UNIT: data.I4_UNIT || "",
    I4_AMOUNT: data.I4_AMOUNT || "",

    LABOR_1_QTY: labor1Qty,
    LABOR_1_RATE: firstHourRate,
    LABOR_1_AMOUNT: labor1Amount,
    LABOR_2_QTY: labor2Qty,
    LABOR_2_RATE: additionalHourRate,
    LABOR_2_AMOUNT: labor2Amount,

    SUB_TOTAL: subTotal,
    GRAND_TOTAL: total,

    CLIENTE: data.CLIENT || "",
    "VALE DE COMPRA": data.PURCHASE_RECEIPT || "",
    COMPANY_ID: companyId,

    STORE_ADDRESS: normalizedAddress.full,
    STORE_STREET: normalizedAddress.street,
    STORE_CITY: normalizedAddress.city,
    STORE_STATE: normalizedAddress.state,
    STORE_ZIP: normalizedAddress.zip,

    REPORTED_PROBLEM: data.REPORTED_PROBLEM || "",
    EQUIPMENT_MAKE: data.EQUIPMENT_MAKE || "",
    EQUIPMENT_MODEL: data.EQUIPMENT_MODEL || "",
    EQUIPMENT_SERIAL: data.EQUIPMENT_SERIAL || "",
    NOTES: data.NOTES || "",
    SIGNATURE: data.SIGNATURE || ""
  };

  try {
    const pdfs = generateInvoicePDFs_(invoiceRow);

    invoiceRow.PDF_EN_URL = pdfs.pdfEnUrl;
    invoiceRow.PDF_ES_URL = pdfs.pdfEsUrl;
    invoiceRow.DOC_EN_URL = pdfs.docEnUrl;
    invoiceRow.DOC_ES_URL = pdfs.docEsUrl;
  } catch (pdfErr) {
    Logger.log("ERROR generando PDF invoice: " + pdfErr);
  }

  try {
    const pdfLinks = generatePdfFromCloseOrder_(invoiceRow);
    invoiceRow.PDF_ES_URL = pdfLinks.PDF_ES_URL;
    invoiceRow.PDF_EN_URL = pdfLinks.PDF_EN_URL;
  } catch (pdfErr) {
    Logger.log("ERROR generando PDFs desde CloseOrder: " + pdfErr);
  }

  const rowValues = invHeaders.map(function(h) {
    return invoiceRow[h] !== undefined ? invoiceRow[h] : "";
  });

  shInv.appendRow(rowValues);

  let oldSyncResult = null;

  try {
    oldSyncResult = syncCloseOrderToOldSystem_(invoiceRow);
  } catch (syncErr) {
    Logger.log("ERROR sincronizando con sistema viejo: " + syncErr);
  }

  try {
    updateTechOrderStatus(rowNumber, "COMPLETED");
  } catch (statusErr) {
    Logger.log("ERROR updateTechOrderStatus desde saveCloseOrder: " + statusErr);
    throw statusErr;
  }

  return {
    success: true,
    invoiceNumber: invoiceNumber,
    woNumber: woNumber,
    invoiceType: "REPAIR",
    oldSync: oldSyncResult
  };
}

function normalizeCloseOrderAddress_(data) {
  let street = String(data.STORE_STREET || "").trim();
  let city = String(data.STORE_CITY || "").trim();
  let state = String(data.STORE_STATE || "").trim().toUpperCase();
  let zip = String(data.STORE_ZIP || "").trim();
  let full = String(data.STORE_ADDRESS || "").trim();

  if ((!street || !city || !state || !zip) && full) {
    const m = full.match(/^(.*?),\s*([^,]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/i);

    if (m) {
      if (!street) street = m[1].trim();
      if (!city) city = m[2].trim().toUpperCase();
      if (!state) state = m[3].trim().toUpperCase();
      if (!zip) zip = m[4].trim();
    }
  }

  const cityStateZip = [
    city,
    [state, zip].filter(Boolean).join(" ")
  ].filter(Boolean).join(", ");

  full = [street, cityStateZip].filter(Boolean).join(", ");

  return {
    street: street,
    city: city,
    state: state,
    zip: zip,
    full: full
  };
}

function syncCloseOrderToOldSystem_(invoiceRow) {
  const result = {
    oldLog: false,
    errors: []
  };

  try {
    syncInvoiceToOldLog_(invoiceRow);
    result.oldLog = true;
  } catch (err) {
    result.errors.push("OLD LOG: " + err.message);
  }

  try {
    logOldSync_(invoiceRow, result);
  } catch (logErr) {
    Logger.log("ERROR guardando sync log: " + logErr);
  }

  return result;
}

function syncInvoiceToOldLog_(invoiceRow) {
  const ss = getOldSystemSpreadsheet_();

  const sheetName = "LOG";
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  const requiredHeaders = getOldLogHeaders_();

  sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);

  const oldRow = buildOldLogRow_(invoiceRow);
  const existingRow = findOldLogRow_(sh, requiredHeaders, oldRow.Invoice);

  const values = requiredHeaders.map(function(h) {
    return oldRow[h] !== undefined ? oldRow[h] : "";
  });

  if (existingRow > 1) {
    sh.getRange(existingRow, 1, 1, requiredHeaders.length).setValues([values]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, requiredHeaders.length).setValues([values]);
  }

  return true;
}

function getOldSystemSpreadsheet_() {
  if (CFG.OLD_ECONOMY_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CFG.OLD_ECONOMY_SPREADSHEET_ID);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOldLogHeaders_() {
  return [
    "Timestamp",
    "Invoice",
    "NS",
    "WO_TYPO",
    "TECHNICIAN EMAIL",
    "TECHNICIAN NAME",
    "PROCESO",
    "HORAS",
    "PARTS",
    "LABOR",
    "TAX",
    "TOTAL",
    "CLIENTE",
    "COMPRA",
    "VALE DE COMPRA"
  ];
}

function buildOldLogRow_(invoiceRow) {
  return {
    Timestamp: invoiceRow.Timestamp || invoiceRow.DATE_INVOICE || new Date(),
    Invoice: invoiceRow.Invoice || invoiceRow.INVOICE_NUMBER || "",
    NS: invoiceRow.NS || "",
    WO_TYPO: invoiceRow.WO_TYPO || invoiceRow.WO_TYPE || "REPAIR_FORM",
    "TECHNICIAN EMAIL": invoiceRow["TECHNICIAN EMAIL"] || "",
    "TECHNICIAN NAME": invoiceRow["TECHNICIAN NAME"] || "",
    PROCESO: invoiceRow.PROCESO || "",
    HORAS: invoiceRow.HORAS || invoiceRow.LABOR_HOURS || 0,
    PARTS: invoiceRow.PARTS || invoiceRow.MATERIAL_COST || invoiceRow.INVERSION || 0,
    LABOR: invoiceRow.LABOR || invoiceRow.LABOR_AMOUNT || 0,
    TAX: invoiceRow.TAX || invoiceRow.TAX_AMOUNT || 0,
    TOTAL: invoiceRow.GRAND_TOTAL || invoiceRow.INVOICE_TOTAL || invoiceRow.TOTAL || 0,
    CLIENTE: invoiceRow.CLIENTE || "",
    COMPRA: invoiceRow.COMPRA || invoiceRow.MATERIAL_COST || "",
    "VALE DE COMPRA": invoiceRow["VALE DE COMPRA"] || ""
  };
}

function findOldLogRow_(sh, headers, invoiceNumber) {
  const idxInv = headers.indexOf("Invoice");
  if (idxInv === -1) return -1;

  const data = sh.getDataRange().getValues();
  const targetInv = String(invoiceNumber || "").trim().toUpperCase();

  if (!targetInv) return -1;

  for (let i = 1; i < data.length; i++) {
    const rowInv = String(data[i][idxInv] || "").trim().toUpperCase();

    if (rowInv === targetInv) {
      return i + 1;
    }
  }

  return -1;
}

function logOldSync_(invoiceRow, result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "SYNC_LOG";
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow([
      "TIMESTAMP",
      "WO_NUMBER",
      "INVOICE_NUMBER",
      "COMPANY_ID",
      "OLD_LOG_SYNC",
      "ERRORS"
    ]);
  }

  sh.appendRow([
    new Date(),
    invoiceRow.WO_NUMBER || "",
    invoiceRow.INVOICE_NUMBER || invoiceRow.Invoice || "",
    invoiceRow.COMPANY_ID || "",
    result.oldLog === true ? "YES" : "NO",
    result.errors && result.errors.length ? result.errors.join(" | ") : ""
  ]);
}

function generateInvoiceNumber_(companyId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    const key = "LAST_INVOICE_NUMBER_" + String(companyId || "DEFAULT").toUpperCase();

    let last = Number(props.getProperty(key) || 10000);
    last++;
    props.setProperty(key, String(last));

    return String(last);
  } finally {
    lock.releaseLock();
  }
}
