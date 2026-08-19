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

  // Economía real D&D: los $100 del técnico salen de los $225 de D&D.
  DD_PAYMENT: 225,
  TECH_PAY: 100,
  PALACIOS_PAYMENT: 125,
  SUPPLIES_RESERVE: 41.67,
  DD_PROFIT: 125
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

  let headers = sh.getRange(1,1,1,sh.getLastColumn())
    .getValues()[0]
    .map(String);

  headers = ensurePMEconomyHeaders_(sh, headers);

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

  const ddPayment = roundPMMoney_(PM_ECO_CFG.DD_PAYMENT);

  const techPay = roundPMMoney_(PM_ECO_CFG.TECH_PAY);

  const palaciosPayment = roundPMMoney_(PM_ECO_CFG.PALACIOS_PAYMENT);

  const suppliesReserve = roundPMMoney_(PM_ECO_CFG.SUPPLIES_RESERVE);

  const amount = roundPMMoney_(
    ddPayment +
    palaciosPayment +
    suppliesReserve
  );

  if (Math.abs(amount - subtotal) > 0.01) {
    throw new Error(
      "El desglose PM no cuadra con el subtotal. " +
      "Subtotal: $" + subtotal.toFixed(2) +
      " | Desglose: $" + amount.toFixed(2)
    );
  }

  const profit = roundPMMoney_(ddPayment - techPay);

  if (Math.abs(profit - roundPMMoney_(PM_ECO_CFG.DD_PROFIT)) > 0.01) {
    throw new Error(
      "La ganancia PM de D&D no cuadra. " +
      "D&D recibe: $" + ddPayment.toFixed(2) +
      " | Técnico: $" + techPay.toFixed(2) +
      " | Ganancia esperada: $" + PM_ECO_CFG.DD_PROFIT.toFixed(2)
    );
  }

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

    DD_PAYMENT:
      ddPayment,

    PALACIOS_PAYMENT:
      palaciosPayment,

    SUPPLIES_RESERVE:
      suppliesReserve,

    PM_TOTAL_AMOUNT:
      amount,

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
      "Technician: $" + techPay.toFixed(2),

    TECH_LABOR_COST:
      techPay,

    NOTES:
      buildPMEconomyNotes_(
        invoiceTotal,
        subtotal,
        tax,
        amount,
        ddPayment,
        techPay,
        palaciosPayment,
        suppliesReserve,
        profit
      ),

    INV_SOURCE:
      "PM",

    PERIOD_MONTH:
      period.month,

    PERIOD_YEAR:
      period.year,

    PERIOD_LABEL:
      period.label,

    ACCOUNTING_PERIOD:
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

  touchAppCacheVersion_();

  return {
    success: true,
    woNumber: woNumber,
    amount: amount,
    ddPayment: ddPayment,
    techPay: techPay,
    palaciosPayment: palaciosPayment,
    suppliesReserve: suppliesReserve,
    profit: profit
  };
}

// =====================================================
// OBTENER DATA PM ECONOMY
// =====================================================

function getPMEconomyData(companyId, role, sessionToken) {
  companyId = String(companyId || PM_ECO_CFG.COMPANY_ID || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);

  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(PM_ECO_CFG.SHEET_NAME);

  if (!sh) {
    throw new Error("No existe PM_ECONOMY");
  }

  let data = sh.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  let headers = data[0].map(function(h) {
    return String(h || "").trim();
  });
  headers = ensurePMEconomyHeaders_(sh, headers);
  repairCurrentPeriodPMLegacyRows_(sh, headers, companyId, session);
  data = sh.getDataRange().getValues();
  headers = data[0].map(function(h) {
    return String(h || "").trim();
  });

  return data.slice(1).map(function(row, i) {
    if (isMonthCloseMarkerRow_(row, headers)) return null;

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
    normalizePMEconomyObjectPeriodFields_(obj);

    return obj;

  }).filter(function(o) {
    if (!o) return false;
    return String(o.COMPANY_ID || "").trim().toUpperCase() === companyId;
  }).reverse();
}

// =====================================================
// UPDATE PM ECONOMY
// =====================================================

function updatePMEconomyRow(rowNumber, updates, sessionToken) {

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

  let headers = sh.getRange(1,1,1,sh.getLastColumn())
    .getValues()[0]
    .map(String);
  headers = ensurePMEconomyHeaders_(sh, headers);
  const companyId = getCellByHeader_(sh, rowNumber, headers, "COMPANY_ID") ||
    PM_ECO_CFG.COMPANY_ID ||
    CFG.DEFAULT_COMPANY_ID;
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);

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

  addAuditLog_("PM_ECONOMY", "PM_ECONOMY_ROW_UPDATED", companyId, "PM_ECONOMY", String(rowNumber), session, {
    rowNumber: rowNumber,
    fields: Object.keys(updates || {})
  });

  return true;
}

// =====================================================
// HELPERS
// =====================================================

function roundPMMoney_(n) {
  return Math.round(Number(n || 0) * 100) / 100;
}

function buildPMEconomyNotes_(
  invoiceTotal,
  subtotal,
  tax,
  amount,
  ddPayment,
  techPay,
  palaciosPayment,
  suppliesReserve,
  profit
) {
  return "PM Invoice Total: $" + roundPMMoney_(invoiceTotal).toFixed(2) +
    " | Subtotal: $" + roundPMMoney_(subtotal).toFixed(2) +
    " | Tax: $" + roundPMMoney_(tax).toFixed(2) +
    " | PM Total Amount: $" + roundPMMoney_(amount).toFixed(2) +
    " | D&D Premier: $" + roundPMMoney_(ddPayment).toFixed(2) +
    " | Tech Pay: $" + roundPMMoney_(techPay).toFixed(2) +
    " | Palacios Power System: $" + roundPMMoney_(palaciosPayment).toFixed(2) +
    " | PM Supplies Reserve: $" + roundPMMoney_(suppliesReserve).toFixed(2) +
    " | D&D Profit: $" + roundPMMoney_(profit).toFixed(2) +
    " | Split Check: $" + roundPMMoney_(amount).toFixed(2);
}

function repairCurrentPeriodPMLegacyRows_(sh, headers, companyId, actor) {
  if (!sh || sh.getLastRow() < 2) {
    return { repaired: 0, rows: [], invoices: [] };
  }

  headers = (headers || []).map(function(header) {
    return String(header || "").trim();
  });

  const required = [
    "COMPANY_ID",
    "AMOUNT",
    "COST",
    "PROFIT",
    "DD_PAYMENT",
    "PALACIOS_PAYMENT",
    "SUPPLIES_RESERVE",
    "PM_TOTAL_AMOUNT",
    "TECH_LABOR_COST",
    "NOTES"
  ];
  if (required.some(function(header) { return headers.indexOf(header) === -1; })) {
    return { repaired: 0, rows: [], invoices: [] };
  }

  const period = getCurrentEconomyPeriod();
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  const repairedRows = [];
  const repairedInvoices = [];

  values.forEach(function(row, index) {
    if (isMonthCloseMarkerRow_(row, headers) || isSoftDeletedRow_(row, headers)) return;

    const rowCompany = String(
      getHeaderValueFlexible_(row, headers, ["COMPANY_ID"]) || PM_ECO_CFG.COMPANY_ID
    ).trim().toUpperCase();
    if (rowCompany !== companyId) return;

    const obj = {};
    headers.forEach(function(header, columnIndex) {
      obj[header] = row[columnIndex];
    });
    normalizePMEconomyObjectPeriodFields_(obj);
    if (String(obj.PERIOD_LABEL || obj.ACCOUNTING_PERIOD || "") !== period.label) return;

    const amount = roundPMMoney_(obj.PM_TOTAL_AMOUNT || obj.AMOUNT);
    const ddPayment = roundPMMoney_(obj.DD_PAYMENT);
    const techPay = roundPMMoney_(obj.TECH_LABOR_COST || obj.COST);
    const profit = roundPMMoney_(obj.PROFIT);
    const palaciosPayment = roundPMMoney_(obj.PALACIOS_PAYMENT);
    const suppliesReserve = roundPMMoney_(obj.SUPPLIES_RESERVE);
    const isLegacyNetClassification =
      Math.abs(amount - PM_ECO_CFG.INVOICE_SUBTOTAL) <= 0.01 &&
      Math.abs(ddPayment - PM_ECO_CFG.DD_PROFIT) <= 0.01 &&
      Math.abs(techPay - PM_ECO_CFG.TECH_PAY) <= 0.01 &&
      Math.abs(profit - PM_ECO_CFG.DD_PROFIT) <= 0.01 &&
      Math.abs(palaciosPayment - PM_ECO_CFG.PALACIOS_PAYMENT) <= 0.01 &&
      Math.abs(suppliesReserve - PM_ECO_CFG.SUPPLIES_RESERVE) <= 0.01;

    if (!isLegacyNetClassification) return;

    const tax = roundPMMoney_(PM_ECO_CFG.INVOICE_SUBTOTAL * PM_ECO_CFG.TAX_RATE);
    const invoiceTotal = roundPMMoney_(PM_ECO_CFG.INVOICE_SUBTOTAL + tax);
    const rowNumber = index + 2;

    obj.AMOUNT = PM_ECO_CFG.INVOICE_SUBTOTAL;
    obj.COST = PM_ECO_CFG.TECH_PAY;
    obj.PROFIT = PM_ECO_CFG.DD_PROFIT;
    obj.DD_PAYMENT = PM_ECO_CFG.DD_PAYMENT;
    obj.PALACIOS_PAYMENT = PM_ECO_CFG.PALACIOS_PAYMENT;
    obj.SUPPLIES_RESERVE = PM_ECO_CFG.SUPPLIES_RESERVE;
    obj.PM_TOTAL_AMOUNT = PM_ECO_CFG.INVOICE_SUBTOTAL;
    obj.TECH_LABOR_COST = PM_ECO_CFG.TECH_PAY;
    obj.NOTES = buildPMEconomyNotes_(
      invoiceTotal,
      PM_ECO_CFG.INVOICE_SUBTOTAL,
      tax,
      PM_ECO_CFG.INVOICE_SUBTOTAL,
      PM_ECO_CFG.DD_PAYMENT,
      PM_ECO_CFG.TECH_PAY,
      PM_ECO_CFG.PALACIOS_PAYMENT,
      PM_ECO_CFG.SUPPLIES_RESERVE,
      PM_ECO_CFG.DD_PROFIT
    );

    sh.getRange(rowNumber, 1, 1, headers.length).setValues([
      headers.map(function(header) {
        return obj[header] !== undefined ? obj[header] : "";
      })
    ]);
    repairedRows.push(rowNumber);
    repairedInvoices.push(String(obj.INVOICE_NUMBER || obj.WO_NUMBER || rowNumber));
  });

  if (repairedRows.length) {
    addAuditLog_(
      "PM_ECONOMY",
      "PM_GROSS_CLASSIFICATION_REPAIRED",
      companyId,
      "PM_ECONOMY",
      period.label,
      actor,
      {
        period: period.label,
        rows: repairedRows,
        invoices: repairedInvoices,
        ddGrossPayment: PM_ECO_CFG.DD_PAYMENT,
        technicianExpense: PM_ECO_CFG.TECH_PAY,
        ddNetProfit: PM_ECO_CFG.DD_PROFIT
      }
    );
    touchAppCacheVersion_();
  }

  return {
    repaired: repairedRows.length,
    rows: repairedRows,
    invoices: repairedInvoices
  };
}

function ensurePMEconomyHeaders_(sh, headers) {
  const required = [
    "COMPANY_ID",
    "WO_NUMBER",
    "CLIENT",
    "NSN",
    "DATE_COMPLETED",
    "AMOUNT",
    "TAX",
    "COST",
    "PROFIT",
    "DD_PAYMENT",
    "PALACIOS_PAYMENT",
    "SUPPLIES_RESERVE",
    "PM_TOTAL_AMOUNT",
    "STATUS",
    "INVOICE_NUMBER",
    "DATE_INVOICE",
    "DATE_PAID",
    "HORAS",
    "TECHNICIANS",
    "TECH_PAY_DETAIL",
    "TECH_LABOR_COST",
    "NOTES",
    "INV_SOURCE",
    "PERIOD_MONTH",
    "PERIOD_YEAR",
    "PERIOD_LABEL",
    "ACCOUNTING_PERIOD",
    "CLOSED_PERIOD",
    "MONTH_CLOSED_AT",
    "MONTH_CLOSED_BY",
    "MONTH_CLOSED_BY_EMAIL"
  ];

  let changed = false;

  required.forEach(function(header) {
    const exists = headers.some(function(existingHeader) {
      return String(existingHeader || "").trim().toUpperCase() === String(header || "").trim().toUpperCase();
    });
    if (exists) return;

    sh.getRange(1, sh.getLastColumn() + 1)
      .setValue(header);
    changed = true;
  });

  if (!changed) {
    return headers;
  }

  return sh.getRange(1,1,1,sh.getLastColumn())
    .getValues()[0]
    .map(String);
}

function normalizePMEconomyObjectPeriodFields_(obj) {
  obj = obj || {};

  const label = getEconomyObjectPeriodLabel_(obj);
  if (!label) return obj;

  obj.PERIOD_LABEL = label;
  if (!obj.ACCOUNTING_PERIOD) obj.ACCOUNTING_PERIOD = label;

  const parts = label.split("-");
  if (!obj.PERIOD_YEAR) obj.PERIOD_YEAR = Number(parts[0]);
  if (!obj.PERIOD_MONTH) obj.PERIOD_MONTH = Number(parts[1]);

  return obj;
}
