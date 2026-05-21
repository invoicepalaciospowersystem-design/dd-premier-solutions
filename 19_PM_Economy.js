// =====================================================
// FILE: 19_PM_Economy.gs
// =====================================================

const PM_ECO_CFG = {
  SHEET_NAME: "PM_ECONOMY",

  COMPANY_ID: "PPS",
  CLIENT: "McDonald's",

  // Invoice real hacia McDonald's
  INVOICE_SUBTOTAL: 391.67,
  TAX_RATE: 0.07,

  // Economía real D&D
  DD_PAYMENT: 180,
  TECH_PAY: 80,
  DD_PROFIT: 100
};

// =====================================================
// CREA / ACTUALIZA REGISTRO PM EN PM_ECONOMY
// =====================================================

function savePMEconomy(pmData) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(PM_ECO_CFG.SHEET_NAME);

  if (!sh) {
    throw new Error("No existe la hoja PM_ECONOMY");
  }

  const headers = sh.getRange(1,1,1,sh.getLastColumn())
    .getValues()[0]
    .map(String);

  const woNumber = String(
    pmData.WO_NUMBER || pmData.woNumber || ""
  ).trim();

  if (!woNumber) {
    throw new Error("WO_NUMBER requerido.");
  }

  // =====================================================
  // CALCULOS
  // =====================================================

  const subtotal = PM_ECO_CFG.INVOICE_SUBTOTAL;

  const tax = roundPMMoney_(
    subtotal * PM_ECO_CFG.TAX_RATE
  );

  const invoiceTotal = roundPMMoney_(
    subtotal + tax
  );

  const amount = PM_ECO_CFG.DD_PAYMENT;

  const techPay = PM_ECO_CFG.TECH_PAY;

  const profit = PM_ECO_CFG.DD_PROFIT;

  const completedDate =
    pmData.DATE_COMPLETED ||
    pmData.dateCompleted ||
    new Date();

  const period = getCurrentEconomyPeriod();

  // =====================================================
  // DATA
  // =====================================================

  const rowObj = {

    COMPANY_ID:
      pmData.COMPANY_ID ||
      PM_ECO_CFG.COMPANY_ID,

    WO_NUMBER: woNumber,

    CLIENT:
      pmData.CLIENT ||
      PM_ECO_CFG.CLIENT,

    NSN:
      pmData.NSN || "",

    DATE_COMPLETED:
      completedDate,

    // ECONOMIA REAL D&D
    AMOUNT:
      amount,

    TAX:
      0,

    COST:
      techPay,

    PROFIT:
      profit,

    STATUS:
      pmData.STATUS || "INVOICED",

    INVOICE_NUMBER:
      pmData.INVOICE_NUMBER || "",

    DATE_INVOICE:
      new Date(),

    DATE_PAID:
      "",

    HORAS:
      pmData.HORAS || "",

    TECHNICIANS:
      pmData.TECHNICIANS ||
      pmData.techName ||
      "",

    TECH_PAY_DETAIL:
      "Technician: $80",

    TECH_LABOR_COST:
      techPay,

    NOTES:
      "PM Invoice Total: $" + invoiceTotal.toFixed(2) +
      " | Subtotal: $" + subtotal.toFixed(2) +
      " | Tax: $" + tax.toFixed(2) +
      " | Paid to D&D: $180 | Tech Pay: $80 | D&D Profit: $100",

    INV_SOURCE:
      "PM",

    PERIOD_MONTH:
      period.month,

    PERIOD_YEAR:
      period.year,

    PERIOD_LABEL:
      period.label
  };

  // =====================================================
  // BUSCAR SI YA EXISTE
  // =====================================================

  const data = sh.getDataRange().getValues();

  const idxWO = headers.indexOf("WO_NUMBER");
  const idxCompany = headers.indexOf("COMPANY_ID");

  let targetRow = -1;

  for (let i = 1; i < data.length; i++) {

    const existingWO =
      String(data[i][idxWO] || "").trim();

    const existingCompany =
      String(data[i][idxCompany] || "")
      .trim()
      .toUpperCase();

    if (
      existingWO === woNumber &&
      existingCompany === String(rowObj.COMPANY_ID).toUpperCase()
    ) {
      targetRow = i + 1;
      break;
    }
  }

  // =====================================================
  // CREAR NUEVO
  // =====================================================

  const rowValues = headers.map(function(h) {
    return rowObj[h] !== undefined
      ? rowObj[h]
      : "";
  });

  if (targetRow === -1) {

    sh.appendRow(rowValues);

  } else {

    sh.getRange(
      targetRow,
      1,
      1,
      rowValues.length
    ).setValues([rowValues]);
  }

  return {
    success: true,
    woNumber: woNumber,
    amount: amount,
    techPay: techPay,
    profit: profit
  };
}

// =====================================================
// OBTENER DATA PM ECONOMY
// =====================================================

function getPMEconomyData() {

  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(PM_ECO_CFG.SHEET_NAME);

  if (!sh) {
    throw new Error("No existe PM_ECONOMY");
  }

  const data = sh.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0].map(String);

  return data.slice(1).map(function(row, i) {

    const obj = {};

    headers.forEach(function(h, c) {

      let value = row[c];

      if (value instanceof Date) {
        value = Utilities.formatDate(
          value,
          Session.getScriptTimeZone(),
          "MM/dd/yyyy"
        );
      }

      obj[h] = value;
    });

    obj.ROW_NUMBER = i + 2;

    return obj;

  }).reverse();
}

// =====================================================
// UPDATE PM ECONOMY
// =====================================================

function updatePMEconomyRow(rowNumber, updates) {

  rowNumber = Number(rowNumber);

  if (!rowNumber || rowNumber < 2) {
    throw new Error("Fila inválida.");
  }

  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(PM_ECO_CFG.SHEET_NAME);

  if (!sh) {
    throw new Error("No existe PM_ECONOMY");
  }

  const headers = sh.getRange(1,1,1,sh.getLastColumn())
    .getValues()[0]
    .map(String);

  Object.keys(updates).forEach(function(key) {

    const col = headers.indexOf(key) + 1;

    if (col > 0) {
      sh.getRange(rowNumber, col)
        .setValue(updates[key]);
    }
  });

  // DATE_PAID automatico
  const statusCol = headers.indexOf("STATUS") + 1;
  const datePaidCol = headers.indexOf("DATE_PAID") + 1;

  if (statusCol > 0 && datePaidCol > 0) {

    const status = String(
      sh.getRange(rowNumber, statusCol).getValue() || ""
    ).toUpperCase();

    if (
      status === "PAID" &&
      !sh.getRange(rowNumber, datePaidCol).getValue()
    ) {
      sh.getRange(rowNumber, datePaidCol)
        .setValue(new Date());
    }
  }

  return true;
}

// =====================================================
// HELPERS
// =====================================================

function roundPMMoney_(n) {
  return Math.round(Number(n || 0) * 100) / 100;
}