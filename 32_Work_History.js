// =====================================================
// FILE: 32_Work_History.gs
// Repair and service history grouped by store/equipment.
// =====================================================

function getWorkHistory(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireSession_(
    sessionToken,
    ["TECH", "ADMIN", "ORDENES", "OWNER"],
    companyId
  );

  const role = String(session.role || "").trim().toUpperCase();
  if (!companyId && role !== "OWNER") {
    companyId = String(session.companyId || "").trim().toUpperCase();
  }
  if (!companyId) {
    companyId = String(CFG.DEFAULT_COMPANY_ID || "PPS").trim().toUpperCase();
  }

  return withAppCache_(["work-history", companyId], 90, function() {
    return buildWorkHistory_(companyId);
  });
}

function buildWorkHistory_(companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return emptyWorkHistoryResult_(companyId);
  }

  const headers = data[0].map(function(header) {
    return String(header || "").trim();
  });
  const idxCompany = workHistoryHeaderIndex_(headers, ["COMPANY_ID"]);
  const invoiceMap = buildWorkHistoryInvoiceMap_(ss, companyId);
  const draftMap = buildWorkHistoryDraftMap_(ss, companyId);
  const storeMap = getStoreMapByCompany_(companyId);
  const records = [];
  let totalMatchingRows = 0;

  for (let index = 1; index < data.length; index++) {
    const row = data[index];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = idxCompany >= 0
      ? String(row[idxCompany] || "").trim().toUpperCase()
      : String(CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
    if (rowCompany !== companyId) continue;

    totalMatchingRows++;
    const order = workHistoryRowObject_(headers, row);
    const woNumber = String(order.WO_NUMBER || "").trim();
    if (!woNumber) continue;

    const invoice = invoiceMap[woNumber.toUpperCase()] || {};
    const draft = draftMap[woNumber.toUpperCase()] || {};
    const nsn = normalizeNSN_(order.NSN);
    const store = storeMap[nsn] || {};
    const equipmentType = workHistoryFirstText_([
      order.REPORTED_EQUIPMENT,
      order["REPORTED EQUIPMENT"],
      order.REPORTED_EQUIPMENT_EN,
      draft.EQUIPMENT_TYPE,
      invoice.EQUIPMENT_TYPE
    ]) || "Equipo no especificado";
    const equipmentMake = workHistoryFirstText_([
      order.EQUIPMENT_MAKE,
      draft.EQUIPMENT_MAKE,
      invoice.EQUIPMENT_MAKE
    ]);
    const equipmentModel = workHistoryFirstText_([
      order.EQUIPMENT_MODEL,
      draft.EQUIPMENT_MODEL,
      invoice.EQUIPMENT_MODEL
    ]);
    const equipmentSerial = workHistoryFirstText_([
      order.EQUIPMENT_SERIAL,
      draft.EQUIPMENT_SERIAL,
      invoice.EQUIPMENT_SERIAL
    ]);
    const completionResult = workHistoryFirstText_([
      order.WORK_COMPLETION_RESULT,
      draft.WORK_COMPLETION_RESULT,
      invoice.WORK_COMPLETION_RESULT
    ]);
    const workPerformed = workHistoryFirstText_([
      order.WORK_PERFORMED,
      draft.WORK_PERFORMED,
      invoice.WORK_PERFORMED,
      invoice.TRABAJO_REALIZADO
    ]);
    const materialsUsed = workHistoryMaterialsText_([order, draft, invoice]);
    const dateValue = workHistoryFirstValue_([
      order.DATE_COMPLETED,
      draft.INVOICE_DATE,
      invoice.DATE_INVOICE,
      invoice.Timestamp,
      order.DATE_WORK_STARTED,
      order.DATE_CREATED,
      order.CREATED_AT
    ]);
    const dateInfo = workHistoryDateInfo_(dateValue);
    const status = String(order.STATUS || "").trim();
    const storeLabel = "NSN " + (nsn || "SIN NSN") +
      (store.client || order.CLIENT ? " - " + String(store.client || order.CLIENT) : "");
    const equipmentLabel = [
      equipmentType,
      equipmentMake,
      equipmentModel ? "Modelo " + equipmentModel : "Modelo no registrado",
      equipmentSerial ? "Serie " + equipmentSerial : ""
    ].filter(Boolean).join(" | ");

    records.push({
      rowNumber: index + 1,
      companyId: rowCompany,
      woNumber: woNumber,
      woType: String(order.WO_TYPE || "").trim(),
      status: status,
      nsn: nsn,
      storeLabel: storeLabel,
      storeAddress: workHistoryFirstText_([
        order.STORE_ADDRESS,
        store.fullAddress,
        [store.address, store.city, store.state, store.zip].filter(Boolean).join(", ")
      ]),
      client: String(store.client || order.CLIENT || "").trim(),
      technician: String(order.TECHNICIAN || draft.TECHNICIAN || invoice["TECHNICIAN NAME"] || "").trim(),
      equipmentType: equipmentType,
      equipmentMake: equipmentMake,
      equipmentModel: equipmentModel,
      equipmentSerial: equipmentSerial,
      equipmentLabel: equipmentLabel,
      reportedProblem: workHistoryFirstText_([
        order.REPORTED_PROBLEM_ES,
        order.REPORTED_PROBLEM_ORIGINAL,
        order.REPORTED_PROBLEM_EN,
        draft.REPORTED_PROBLEM,
        invoice.REPORTED_PROBLEM
      ]),
      problemFound: String(order.TECH_PROBLEM_FOUND || draft.TECH_PROBLEM_FOUND || invoice.TECH_PROBLEM_FOUND || "").trim(),
      proposedSolution: String(order.TECH_PROPOSED_SOLUTION || draft.TECH_PROPOSED_SOLUTION || invoice.TECH_PROPOSED_SOLUTION || "").trim(),
      completionResult: completionResult,
      workPerformed: workPerformed,
      materialsUsed: materialsUsed,
      notes: workHistoryFirstText_([order.NOTES, draft.NOTES, invoice.NOTES]),
      hours: workHistoryFirstText_([draft.TECH_HOURS, invoice.LABOR_HOURS, invoice.HORAS]),
      invoiceNumber: workHistoryFirstText_([draft.INVOICE_NUMBER, invoice.INVOICE_NUMBER, invoice.Invoice]),
      attachmentCount: Math.max(0, Number(order.ATTACHMENT_COUNT || 0)),
      date: dateInfo.text,
      dateMs: dateInfo.ms,
      completed: workHistoryIsCompletedStatus_(status) || !!String(order.DATE_COMPLETED || "").trim()
    });
  }

  records.sort(function(a, b) {
    return Number(b.dateMs || 0) - Number(a.dateMs || 0);
  });

  const maxRecords = 1500;
  const limitedRecords = records.slice(0, maxRecords);
  const storeKeys = {};
  const equipmentKeys = {};
  let completedCount = 0;

  limitedRecords.forEach(function(record) {
    storeKeys[record.nsn || record.storeLabel] = true;
    equipmentKeys[(record.nsn || "") + "|" + record.equipmentLabel] = true;
    if (record.completed) completedCount++;
    delete record.dateMs;
  });

  return {
    companyId: companyId,
    companyName: getCompanyName(companyId) || companyId,
    records: limitedRecords,
    totals: {
      records: limitedRecords.length,
      stores: Object.keys(storeKeys).length,
      equipment: Object.keys(equipmentKeys).length,
      completed: completedCount,
      sourceRows: totalMatchingRows
    },
    truncated: records.length > maxRecords,
    availableRecords: records.length
  };
}

function buildWorkHistoryInvoiceMap_(ss, companyId) {
  return buildWorkHistorySupplementMap_(ss.getSheetByName("INVOICES"), companyId);
}

function buildWorkHistoryDraftMap_(ss, companyId) {
  return buildWorkHistorySupplementMap_(
    ss.getSheetByName(CFG.SHEET_INVOICE_DRAFTS || "INVOICE_DRAFTS"),
    companyId
  );
}

function buildWorkHistorySupplementMap_(sheet, companyId) {
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(header) {
    return String(header || "").trim();
  });
  const idxCompany = workHistoryHeaderIndex_(headers, ["COMPANY_ID"]);
  const idxWO = workHistoryHeaderIndex_(headers, ["WO_NUMBER", "WO"]);
  if (idxWO === -1) return map;

  for (let index = 1; index < values.length; index++) {
    const rowCompany = idxCompany >= 0
      ? String(values[index][idxCompany] || "").trim().toUpperCase()
      : companyId;
    if (rowCompany && rowCompany !== companyId) continue;

    const woNumber = String(values[index][idxWO] || "").trim().toUpperCase();
    if (!woNumber) continue;
    map[woNumber] = workHistoryRowObject_(headers, values[index]);
  }

  return map;
}

function workHistoryRowObject_(headers, row) {
  const obj = {};
  headers.forEach(function(header, index) {
    obj[header] = row[index];
  });
  return obj;
}

function workHistoryHeaderIndex_(headers, candidates) {
  const normalized = headers.map(function(header) {
    return String(header || "").trim().toUpperCase();
  });

  for (let index = 0; index < candidates.length; index++) {
    const found = normalized.indexOf(String(candidates[index] || "").trim().toUpperCase());
    if (found !== -1) return found;
  }
  return -1;
}

function workHistoryFirstValue_(values) {
  for (let index = 0; index < values.length; index++) {
    const value = values[index];
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    if (String(value === undefined || value === null ? "" : value).trim()) return value;
  }
  return "";
}

function workHistoryFirstText_(values) {
  const value = workHistoryFirstValue_(values || []);
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a");
  }
  return String(value === undefined || value === null ? "" : value).trim();
}

function workHistoryDateInfo_(value) {
  let date = null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    date = value;
  } else if (String(value || "").trim()) {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) date = parsed;
  }

  return date
    ? {
        text: Utilities.formatDate(date, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a"),
        ms: date.getTime()
      }
    : { text: String(value || "").trim(), ms: 0 };
}

function workHistoryMaterialsText_(sources) {
  const lines = [];
  sources = sources || [];

  for (let itemNumber = 1; itemNumber <= 4; itemNumber++) {
    const qtyValues = [];
    const partValues = [];
    const descValues = [];

    sources.forEach(function(source) {
      source = source || {};
      qtyValues.push(source["I" + itemNumber + "_QTY"]);
      partValues.push(source["I" + itemNumber + "_PART"]);
      descValues.push(source["I" + itemNumber + "_DESC"]);
    });

    const qty = workHistoryFirstText_(qtyValues);
    const part = workHistoryFirstText_(partValues);
    const desc = workHistoryFirstText_(descValues);
    if (!qty && !part && !desc) continue;

    lines.push([
      qty ? "Qty " + qty : "",
      part ? "Part # " + part : "",
      desc
    ].filter(Boolean).join(" - "));
  }

  return lines.join("\n");
}

function workHistoryIsCompletedStatus_(status) {
  status = String(status || "").trim().toUpperCase();
  return ["READY FOR BILLING", "COMPLETED", "CLOSED", "INVOICED"].some(function(value) {
    return status.indexOf(value) !== -1;
  });
}

function emptyWorkHistoryResult_(companyId) {
  return {
    companyId: companyId,
    companyName: getCompanyName(companyId) || companyId,
    records: [],
    totals: { records: 0, stores: 0, equipment: 0, completed: 0, sourceRows: 0 },
    truncated: false,
    availableRecords: 0
  };
}
