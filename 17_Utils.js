function getMaterialCostValue_(row, headers) {
  return parseMoneyFlexible_(
    getHeaderValueFlexible_(row, headers, [
      "MATERIAL_COST",
      "INVERSION",
      "PARTS",
      "PART",
      "COMPRA",
      "COST"
    ])
  );
}

function getLaborAmountValue_(row, headers) {
  return parseMoneyFlexible_(
    getHeaderValueFlexible_(row, headers, [
      "LABOR_AMOUNT",
      "LABOR",
      "LABOR_COST"
    ])
  );
}

function getHoursValue_(row, headers) {
  return parseMoneyFlexible_(
    getHeaderValueFlexible_(row, headers, [
      "LABOR_HOURS",
      "HORAS",
      "HOURS"
    ])
  );
}

function getInvoiceNumberValue_(row, headers) {
  return String(
    getHeaderValueFlexible_(row, headers, [
      "INVOICE_NUMBER",
      "Invoice",
      "INVOICE"
    ]) || ""
  ).trim();
}

function getInvoiceDateValue_(row, headers) {
  return getHeaderValueFlexible_(row, headers, [
    "DATE_INVOICE",
    "Timestamp",
    "INVOICE_DATE"
  ]);
}

function getInvoiceTotalValue_(row, headers) {
  return parseMoneyFlexible_(
    getHeaderValueFlexible_(row, headers, [
      "INVOICE_TOTAL",
      "TOTAL",
      "AMOUNT"
    ])
  );
}

function getTaxAmountValue_(row, headers) {
  return parseMoneyFlexible_(
    getHeaderValueFlexible_(row, headers, [
      "TAX_AMOUNT",
      "TAX"
    ])
  );
}

function getHeaderValueFlexible_(row, headers, possibleNames) {
  headers = headers.map(function(h) {
    return String(h || "").trim();
  });

  for (let i = 0; i < possibleNames.length; i++) {
    const wanted = String(possibleNames[i] || "").trim();

    const idx = headers.indexOf(wanted);

    if (idx >= 0) {
      return row[idx];
    }
  }

  return "";
}

function parseMoneyFlexible_(v) {
  if (v === null || v === undefined) return 0;

  const cleaned = String(v).replace(/[^0-9.-]/g, "");
  const n = parseFloat(cleaned);

  return isNaN(n) ? 0 : n;
}

function normalizeAllSheetHeadersManual() {
  const sheetsToNormalize = [
    "INVOICES",
    CFG.SHEET_ECONOMY
  ];

  sheetsToNormalize.forEach(function(sheetName) {
    normalizeSheetHeaders_(sheetName);
  });

  return "Headers normalizados correctamente.";
}

function normalizeSheetHeaders_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);

  if (!sh) {
    Logger.log("No existe hoja: " + sheetName);
    return;
  }

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  const range = sh.getRange(1, 1, 1, lastCol);
  const headers = range.getValues()[0];

  const map = {
    // Invoice
    "Invoice": "INVOICE_NUMBER",
    "INVOICE": "INVOICE_NUMBER",
    "Timestamp": "DATE_INVOICE",

    // Totals
    "TOTAL": "INVOICE_TOTAL",
    "AMOUNT": "INVOICE_TOTAL",

    // Labor
    "HORAS": "LABOR_HOURS",
    "HOURS": "LABOR_HOURS",
    "LABOR": "LABOR_AMOUNT",

    // Materials
    "INVERSION": "MATERIAL_COST",
    "COMPRA": "MATERIAL_COST",
    "PART": "MATERIAL_COST",
    "PARTS": "MATERIAL_COST",
    "COST": "MATERIAL_COST",

    // Tax
    "TAX": "TAX_AMOUNT",

    // Tech pay
    "TECH_LABOR_COST": "TECH_LABOR_PAY"
  };

  const used = {};
  const newHeaders = headers.map(function(h) {
    const oldName = String(h || "").trim();
    const newName = map[oldName] || oldName;

    if (!newName) return oldName;

    // Evita columnas duplicadas con el mismo nombre
    if (used[newName]) {
      used[newName]++;
      return newName + "_OLD_" + used[newName];
    }

    used[newName] = 1;
    return newName;
  });

  range.setValues([newHeaders]);

  Logger.log("Headers normalizados en hoja: " + sheetName);
}