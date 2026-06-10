// =====================================================
// FILE: 27_TechManualHours.gs
// Manual technician labor adjustments that do not create invoices.
// =====================================================

const TECH_MANUAL_HOURS_HEADERS_ = [
  "TIMESTAMP",
  "COMPANY_ID",
  "PERIOD_MONTH",
  "PERIOD_YEAR",
  "PERIOD_LABEL",
  "DATE_WORKED",
  "TECHNICIAN",
  "HOURS",
  "RATE",
  "LABOR_PAY",
  "REASON",
  "WO_NUMBER",
  "NOTES",
  "CREATED_BY_EMAIL",
  "CREATED_BY_NAME",
  "ACTIVE",
  "DELETED_AT",
  "DELETED_BY"
];

function getTechOptionsForManualHours(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_USERS);
  if (!sh) throw new Error("No existe USERS.");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  const idxName = headers.indexOf("NAME");
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxRole = headers.indexOf("ROLE");
  const idxRate = headers.indexOf("HOURLY_RATE");
  const idxActive = headers.indexOf("ACTIVE");

  if ([idxName, idxCompany, idxRole, idxRate].indexOf(-1) >= 0) {
    throw new Error("USERS debe tener NAME, COMPANY_ID, ROLE y HOURLY_RATE.");
  }

  const seen = {};
  const options = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const name = String(row[idxName] || "").trim();
    const rowCompany = String(row[idxCompany] || "").trim().toUpperCase();
    const role = String(row[idxRole] || "").trim().toUpperCase();
    const active = idxActive >= 0 ? String(row[idxActive] || "YES").trim().toUpperCase() : "YES";

    if (!name || rowCompany !== companyId || role !== "TECH" || active === "NO") continue;

    const key = name.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;

    options.push({
      name: name,
      rate: Number(row[idxRate] || 0)
    });
  }

  return options.sort(function(a, b) {
    return String(a.name || "").localeCompare(String(b.name || ""));
  });
}

function getTechManualHoursData(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);

  const sh = getOrCreateTechManualHoursSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const obj = techManualHoursRowToObject_(row, headers, i + 1);
    if (String(obj.COMPANY_ID || "").trim().toUpperCase() !== companyId) continue;
    result.push(obj);
  }

  return result.reverse();
}

function saveTechManualHours(payload, sessionToken) {
  payload = payload || {};
  const companyId = String(payload.companyId || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  if (!companyId) throw new Error("Falta la compania.");

  const technician = String(payload.technician || "").trim();
  if (!technician) throw new Error("Seleccione el tecnico.");

  const hours = Number(payload.hours || 0);
  if (!hours || hours <= 0) throw new Error("Las horas deben ser mayor que 0.");
  if (hours > 24) throw new Error("Revise las horas. No se permiten mas de 24 horas en un ajuste.");

  const dateWorked = parseManualHoursDate_(payload.dateWorked);
  if (!dateWorked) throw new Error("Seleccione la fecha trabajada.");

  const reason = String(payload.reason || "").trim();
  if (!reason) throw new Error("Escriba el motivo del ajuste.");

  let rate = Number(payload.rate || 0);
  if (rate < 0) throw new Error("El rate no puede ser negativo.");
  if (!rate) rate = getTechHourlyRate_(companyId, technician);

  const laborPay = roundMoney_(hours * rate);
  const period = getCurrentEconomyPeriod();
  const sh = getOrCreateTechManualHoursSheet_();
  const headers = ensureSheetColumns_(sh, TECH_MANUAL_HOURS_HEADERS_);

  const rowObject = {
    TIMESTAMP: new Date(),
    COMPANY_ID: companyId,
    PERIOD_MONTH: period.month,
    PERIOD_YEAR: period.year,
    PERIOD_LABEL: period.label,
    DATE_WORKED: dateWorked,
    TECHNICIAN: technician,
    HOURS: hours,
    RATE: rate,
    LABOR_PAY: laborPay,
    REASON: reason,
    WO_NUMBER: String(payload.woNumber || "").trim(),
    NOTES: String(payload.notes || "").trim(),
    CREATED_BY_EMAIL: session.email || "",
    CREATED_BY_NAME: session.name || "",
    ACTIVE: "YES",
    DELETED_AT: "",
    DELETED_BY: ""
  };

  sh.appendRow(headers.map(function(header) {
    return rowObject.hasOwnProperty(header) ? rowObject[header] : "";
  }));

  const rowNumber = sh.getLastRow();

  addAuditLog_("ECONOMY", "TECH_MANUAL_HOURS_CREATED", companyId, "TECH_MANUAL_HOURS", rowNumber, session, {
    technician: technician,
    hours: hours,
    rate: rate,
    laborPay: laborPay,
    periodLabel: period.label,
    reason: reason,
    woNumber: rowObject.WO_NUMBER
  });

  return {
    success: true,
    rowNumber: rowNumber,
    technician: technician,
    hours: hours,
    rate: rate,
    laborPay: laborPay,
    periodLabel: period.label
  };
}

function deleteTechManualHours(rowNumber, sessionToken) {
  rowNumber = Number(rowNumber || 0);
  if (!rowNumber || rowNumber < 2) throw new Error("Fila invalida.");

  const sh = getOrCreateTechManualHoursSheet_();
  const headers = ensureSheetColumns_(sh, TECH_MANUAL_HOURS_HEADERS_);
  const row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  if (isSoftDeletedRow_(row, headers)) return true;

  const companyId = getManualHoursCell_(row, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);

  setTechManualHoursCell_(sh, rowNumber, headers, "ACTIVE", "NO");
  setTechManualHoursCell_(sh, rowNumber, headers, "DELETED_AT", new Date());
  setTechManualHoursCell_(sh, rowNumber, headers, "DELETED_BY", getSessionActorLabel_(session));

  addAuditLog_("ECONOMY", "TECH_MANUAL_HOURS_DELETED", companyId, "TECH_MANUAL_HOURS", rowNumber, session, {
    technician: getManualHoursCell_(row, headers, "TECHNICIAN"),
    hours: getManualHoursCell_(row, headers, "HOURS"),
    laborPay: getManualHoursCell_(row, headers, "LABOR_PAY")
  });

  return true;
}

function addManualTechHoursToSummary_(summary, companyId, periodMode, month, year) {
  const rows = getTechManualHoursRowsForPeriod_(companyId, periodMode, month, year);

  rows.forEach(function(row) {
    const name = String(row.TECHNICIAN || "").trim();
    if (!name) return;

    const key = name.toLowerCase();
    const hours = Number(row.HOURS || 0);
    const rate = Number(row.RATE || 0);
    const laborPay = Number(row.LABOR_PAY || (hours * rate) || 0);
    const label = String(row.WO_NUMBER || "").trim()
      ? String(row.WO_NUMBER || "").trim() + " manual"
      : "Horas manuales: " + String(row.REASON || "").trim();

    if (!summary[key]) {
      summary[key] = {
        name: name,
        rate: rate,
        orders: 0,
        hours: 0,
        materialCost: 0,
        laborPay: 0,
        materialBonus: 0,
        totalPay: 0,
        workOrders: []
      };
    }

    if (!summary[key].rate && rate) summary[key].rate = rate;
    summary[key].hours += hours;
    summary[key].laborPay += laborPay;
    summary[key].totalPay += laborPay;
    summary[key].workOrders.push(label);
  });
}

function addManualTechHoursToInvoiceRows_(rows, companyId, techName, month, year) {
  const target = normalizeIdentity_(techName);
  const manualRows = getTechManualHoursRowsForPeriod_(companyId, "monthly", month, year);

  manualRows.forEach(function(row) {
    if (normalizeIdentity_(row.TECHNICIAN) !== target) return;

    const hours = Number(row.HOURS || 0);
    const laborPay = Number(row.LABOR_PAY || 0);

    rows.push({
      woNumber: String(row.WO_NUMBER || "").trim() || "MANUAL HOURS",
      storeNumber: "",
      invoiceNumber: "",
      equipment: "Manual labor adjustment",
      problem: String(row.REASON || "").trim(),
      hours: hours,
      laborPay: laborPay,
      materialBase: 0,
      materialBonus: 0,
      totalEarned: laborPay,
      pdfUrl: "",
      manualHours: true
    });
  });
}

function getTechManualHoursRowsForPeriod_(companyId, periodMode, month, year) {
  companyId = String(companyId || "").trim().toUpperCase();
  periodMode = String(periodMode || "monthly").trim().toLowerCase();
  periodMode = periodMode === "annual" || periodMode === "yearly" ? "annual" : "monthly";
  month = Number(month || 0);
  year = Number(year || 0);

  const sh = getOrCreateTechManualHoursSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const obj = techManualHoursRowToObject_(row, headers, i + 1);
    if (String(obj.COMPANY_ID || "").trim().toUpperCase() !== companyId) continue;
    if (!techManualHoursMatchesPeriod_(obj, periodMode, month, year)) continue;
    rows.push(obj);
  }

  return rows;
}

function techManualHoursMatchesPeriod_(obj, periodMode, month, year) {
  const rowYear = Number(obj.PERIOD_YEAR || 0);
  const rowMonth = Number(obj.PERIOD_MONTH || 0);
  const label = String(obj.PERIOD_LABEL || "").trim();

  if (String(periodMode || "monthly") === "annual") {
    if (rowYear && rowYear === Number(year)) return true;
    if (label && label.indexOf(String(year)) === 0) return true;
    const date = parseManualHoursDate_(obj.DATE_WORKED);
    return !!date && date.getFullYear() === Number(year);
  }

  const wantedLabel = Number(year) + "-" + String(Number(month)).padStart(2, "0");
  if (label && label === wantedLabel) return true;
  if (rowYear && rowMonth && rowYear === Number(year) && rowMonth === Number(month)) return true;

  const date = parseManualHoursDate_(obj.DATE_WORKED);
  return !!date && date.getFullYear() === Number(year) && (date.getMonth() + 1) === Number(month);
}

function getOrCreateTechManualHoursSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CFG.SHEET_TECH_MANUAL_HOURS || "TECH_MANUAL_HOURS";
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(TECH_MANUAL_HOURS_HEADERS_);
    return sh;
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(TECH_MANUAL_HOURS_HEADERS_);
  } else {
    ensureSheetColumns_(sh, TECH_MANUAL_HOURS_HEADERS_);
  }

  return sh;
}

function techManualHoursRowToObject_(row, headers, rowNumber) {
  const obj = {};

  headers.forEach(function(header, index) {
    let value = row[index];
    if (value instanceof Date) {
      value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy");
    }
    obj[header] = value;
  });

  obj.ROW_NUMBER = rowNumber;
  return obj;
}

function getTechHourlyRate_(companyId, techName) {
  companyId = String(companyId || "").trim().toUpperCase();
  const target = normalizeIdentity_(techName);
  if (!target) return 0;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  if (!sh) return 0;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return 0;

  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  const idxName = headers.indexOf("NAME");
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxRate = headers.indexOf("HOURLY_RATE");
  if ([idxName, idxCompany, idxRate].indexOf(-1) >= 0) return 0;

  for (let i = 1; i < values.length; i++) {
    const rowCompany = String(values[i][idxCompany] || "").trim().toUpperCase();
    const name = normalizeIdentity_(values[i][idxName]);
    if (rowCompany === companyId && name === target) {
      return Number(values[i][idxRate] || 0);
    }
  }

  return 0;
}

function parseManualHoursDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  value = String(value || "").trim();
  if (!value) return null;

  let date;
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    date = new Date(value + "T12:00:00");
  } else {
    date = new Date(value);
  }

  return isNaN(date.getTime()) ? null : date;
}

function roundMoney_(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}

function getManualHoursCell_(row, headers, headerName) {
  const idx = headers.indexOf(headerName);
  return idx >= 0 ? row[idx] : "";
}

function setTechManualHoursCell_(sh, rowNumber, headers, headerName, value) {
  const idx = headers.indexOf(headerName);
  if (idx >= 0) sh.getRange(rowNumber, idx + 1).setValue(value);
}
