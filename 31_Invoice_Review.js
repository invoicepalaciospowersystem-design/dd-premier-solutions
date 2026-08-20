// =====================================================
// FILE: 31_Invoice_Review.gs
// Reporte tecnico pendiente y aprobacion final del invoice.
// =====================================================

const INVOICE_DRAFT_HEADERS = [
  "DRAFT_ID",
  "COMPANY_ID",
  "WO_NUMBER",
  "ROW_NUMBER",
  "WO_TYPE",
  "STATUS",
  "INVOICE_DATE",
  "TECH_HOURS",
  "FINAL_HOURS",
  "TECH_PROBLEM_FOUND",
  "TECH_PROPOSED_SOLUTION",
  "WORK_COMPLETION_RESULT",
  "WORK_PERFORMED",
  "EQUIPMENT_MAKE",
  "EQUIPMENT_MODEL",
  "EQUIPMENT_SERIAL",
  "NOTES",
  "SIGNATURE",
  "TECHNICIAN",
  "TECHNICIAN_EMAIL",
  "NSN",
  "CLIENT",
  "STORE_ADDRESS",
  "STORE_STREET",
  "STORE_CITY",
  "STORE_STATE",
  "STORE_ZIP",
  "REPORTED_PROBLEM",
  "I1_QTY",
  "I1_PART",
  "I1_DESC",
  "I2_QTY",
  "I2_PART",
  "I2_DESC",
  "I3_QTY",
  "I3_PART",
  "I3_DESC",
  "I4_QTY",
  "I4_PART",
  "I4_DESC",
  "SUBMITTED_AT",
  "SUBMITTED_BY",
  "UPDATED_AT",
  "UPDATED_BY",
  "INVOICE_NUMBER",
  "FINALIZED_AT",
  "FINALIZED_BY"
];

function ensureInvoiceDraftSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CFG.SHEET_INVOICE_DRAFTS || "INVOICE_DRAFTS";
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, INVOICE_DRAFT_HEADERS.length).setValues([INVOICE_DRAFT_HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, INVOICE_DRAFT_HEADERS.length).setValues([INVOICE_DRAFT_HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }

  ensureSheetColumns_(sh, INVOICE_DRAFT_HEADERS);
  sh.setFrozenRows(1);
  return sh;
}

function saveRepairInvoiceDraft_(data, session, rowWO, headersWO) {
  data = data || {};
  session = session || {};

  const rowNumber = Number(data.ROW_NUMBER || 0);
  const companyId = String(
    getRowValue_(rowWO, headersWO, "COMPANY_ID") ||
    data.COMPANY_ID ||
    CFG.DEFAULT_COMPANY_ID
  ).trim().toUpperCase();
  const woNumber = String(
    getRowValue_(rowWO, headersWO, "WO_NUMBER") ||
    data.WO_NUMBER ||
    ""
  ).trim();
  const techProblemFound = String(
    getRowValue_(rowWO, headersWO, "TECH_PROBLEM_FOUND") ||
    data.TECH_PROBLEM_FOUND ||
    ""
  ).trim();
  const techProposedSolution = String(
    getRowValue_(rowWO, headersWO, "TECH_PROPOSED_SOLUTION") ||
    data.TECH_PROPOSED_SOLUTION ||
    ""
  ).trim();
  const workCompletionResult = String(
    data.WORK_COMPLETION_RESULT ||
    data.WORK_PERFORMED ||
    ""
  ).trim();
  const workPerformed = buildRepairCompletionReport_(
    techProblemFound,
    techProposedSolution,
    workCompletionResult,
    data.WORK_PERFORMED
  );
  const reportedProblem = String(
    getRowValue_(rowWO, headersWO, "REPORTED_PROBLEM_ES") ||
    getRowValue_(rowWO, headersWO, "REPORTED_PROBLEM_EN") ||
    getRowValue_(rowWO, headersWO, "REPORTED_PROBLEM_ORIGINAL") ||
    data.REPORTED_PROBLEM ||
    ""
  ).trim();
  const signature = String(data.SIGNATURE || "").trim();
  const rawHours = data.HORAS !== undefined && data.HORAS !== null && data.HORAS !== ""
    ? data.HORAS
    : data.LABOR_HOURS;
  const hours = validateInvoiceDraftHours_(rawHours);
  const rawInvoiceDate = data.INVOICE_DATE_MANUAL || data.MANUAL_INVOICE_DATE || "";
  if (!String(rawInvoiceDate || "").trim()) {
    throw new Error("La fecha del servicio / invoice es obligatoria.");
  }
  const invoiceDate = getCloseOrderInvoiceDate_(data);
  const technicalMaterials = [];
  for (let itemNumber = 1; itemNumber <= 4; itemNumber++) {
    technicalMaterials.push(normalizeTechnicalMaterial_(data, itemNumber));
  }

  if (!woNumber) throw new Error("La orden no tiene WO_NUMBER.");
  if (!workPerformed) throw new Error("Trabajo realizado es obligatorio.");
  if (!signature) throw new Error("La firma o nombre del cliente es obligatorio.");

  const existingInvoice = findInvoiceNumberForWorkOrder_(companyId, woNumber);
  if (existingInvoice) {
    throw new Error("Esta orden ya tiene el invoice " + existingInvoice + ".");
  }

  const actor = getSessionActorLabel_(session);
  const now = new Date();
  const draftLock = LockService.getScriptLock();
  draftLock.waitLock(30000);

  let draftId = "";
  let draftRowNumber = 0;

  try {
    const sh = ensureInvoiceDraftSheet_();
    const headers = getInvoiceDraftHeaders_(sh);
    const existing = findInvoiceDraft_(sh, headers, "", companyId, woNumber);
    const existingStatus = String(existing && existing.data.STATUS || "").trim().toUpperCase();

    if (existingStatus === "INVOICED") {
      throw new Error(
        "Este reporte ya fue convertido en invoice" +
        (existing.data.INVOICE_NUMBER ? " " + existing.data.INVOICE_NUMBER : "") +
        "."
      );
    }

    if (existingStatus === "PROCESSING") {
      throw new Error("Este reporte esta siendo procesado por Administracion.");
    }

    draftId = existing ? String(existing.data.DRAFT_ID || "") : Utilities.getUuid();
    draftRowNumber = existing ? existing.rowNumber : 0;

    const draft = Object.assign({}, existing ? existing.data : {}, {
      DRAFT_ID: draftId,
      COMPANY_ID: companyId,
      WO_NUMBER: woNumber,
      ROW_NUMBER: rowNumber,
      WO_TYPE: "REPAIR_FORM",
      STATUS: "READY FOR BILLING",
      INVOICE_DATE: invoiceDate,
      TECH_HOURS: hours,
      FINAL_HOURS: hours,
      TECH_PROBLEM_FOUND: techProblemFound,
      TECH_PROPOSED_SOLUTION: techProposedSolution,
      WORK_COMPLETION_RESULT: workCompletionResult,
      WORK_PERFORMED: workPerformed,
      EQUIPMENT_MAKE: String(data.EQUIPMENT_MAKE || "").trim(),
      EQUIPMENT_MODEL: String(data.EQUIPMENT_MODEL || "").trim(),
      EQUIPMENT_SERIAL: String(data.EQUIPMENT_SERIAL || "").trim(),
      NOTES: String(data.NOTES || "").trim(),
      SIGNATURE: signature,
      TECHNICIAN: String(
        getRowValue_(rowWO, headersWO, "TECHNICIAN") ||
        data.TECHNICIAN ||
        session.name ||
        ""
      ).trim(),
      TECHNICIAN_EMAIL: String(
        getRowValue_(rowWO, headersWO, "TECH_EMAIL") ||
        data.TECHNICIAN_EMAIL ||
        session.email ||
        ""
      ).trim(),
      NSN: String(getRowValue_(rowWO, headersWO, "NSN") || data.NSN || "").trim(),
      CLIENT: String(getRowValue_(rowWO, headersWO, "CLIENT") || data.CLIENT || "").trim(),
      STORE_ADDRESS: String(data.STORE_ADDRESS || "").trim(),
      STORE_STREET: String(data.STORE_STREET || "").trim(),
      STORE_CITY: String(data.STORE_CITY || "").trim(),
      STORE_STATE: String(data.STORE_STATE || "").trim(),
      STORE_ZIP: String(data.STORE_ZIP || "").trim(),
      REPORTED_PROBLEM: reportedProblem,
      I1_QTY: technicalMaterials[0].qty,
      I1_PART: technicalMaterials[0].part,
      I1_DESC: technicalMaterials[0].desc,
      I2_QTY: technicalMaterials[1].qty,
      I2_PART: technicalMaterials[1].part,
      I2_DESC: technicalMaterials[1].desc,
      I3_QTY: technicalMaterials[2].qty,
      I3_PART: technicalMaterials[2].part,
      I3_DESC: technicalMaterials[2].desc,
      I4_QTY: technicalMaterials[3].qty,
      I4_PART: technicalMaterials[3].part,
      I4_DESC: technicalMaterials[3].desc,
      SUBMITTED_AT: existing && existing.data.SUBMITTED_AT ? existing.data.SUBMITTED_AT : now,
      SUBMITTED_BY: existing && existing.data.SUBMITTED_BY ? existing.data.SUBMITTED_BY : actor,
      UPDATED_AT: now,
      UPDATED_BY: actor,
      INVOICE_NUMBER: "",
      FINALIZED_AT: "",
      FINALIZED_BY: ""
    });

    draftRowNumber = writeInvoiceDraft_(sh, headers, draftRowNumber, draft);
  } finally {
    draftLock.releaseLock();
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  const currentHeaders = ensureSheetColumns_(shWO, [
    "WORK_COMPLETION_RESULT",
    "WORK_PERFORMED",
    "TECH_COMPLETION_UPDATED_AT",
    "TECH_COMPLETION_BY"
  ]);
  const oldStatus = String(getCellByHeader_(shWO, rowNumber, currentHeaders, "STATUS") || "");

  setCellByHeader_(shWO, rowNumber, currentHeaders, "STATUS", "READY FOR BILLING");
  setCellByHeader_(shWO, rowNumber, currentHeaders, "WORK_COMPLETION_RESULT", workCompletionResult);
  setCellByHeader_(shWO, rowNumber, currentHeaders, "WORK_PERFORMED", workPerformed);
  setCellByHeader_(shWO, rowNumber, currentHeaders, "TECH_COMPLETION_UPDATED_AT", now);
  setCellByHeader_(shWO, rowNumber, currentHeaders, "TECH_COMPLETION_BY", actor);
  if (!getCellByHeader_(shWO, rowNumber, currentHeaders, "DATE_COMPLETED")) {
    setCellByHeader_(shWO, rowNumber, currentHeaders, "DATE_COMPLETED", invoiceDate);
  }

  addLog_(
    companyId,
    woNumber,
    "TECHNICAL REPORT SUBMITTED",
    oldStatus,
    "READY FOR BILLING",
    actor,
    ""
  );

  addAuditLog_("INVOICE", "TECHNICAL_REPORT_SUBMITTED", companyId, "INVOICE_DRAFT", draftId, session, {
    woNumber: woNumber,
    rowNumber: rowNumber,
    draftRowNumber: draftRowNumber,
    technicianHours: hours,
    invoiceDate: Utilities.formatDate(invoiceDate, CFG.TIMEZONE, "yyyy-MM-dd")
  });

  addNotification_(
    companyId,
    "ADMIN",
    woNumber,
    "BILLING",
    "Invoice pendiente de revision: " + woNumber
  );

  touchAppCacheVersion_();

  return {
    success: true,
    pending: true,
    draftId: draftId,
    woNumber: woNumber,
    status: "READY FOR BILLING",
    message: "Reporte enviado a Facturacion pendiente."
  };
}

function getPendingInvoiceDrafts(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);
  const role = String(session.role || "").trim().toUpperCase();
  const effectiveCompanyId = companyId ||
    (role === "OWNER" ? "" : String(session.companyId || "").trim().toUpperCase());

  const sh = ensureInvoiceDraftSheet_();
  if (sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(function(h) {
    return String(h || "").trim();
  });

  return values.slice(1).map(function(row, index) {
    const obj = invoiceDraftRowToObject_(headers, row);
    obj.SHEET_ROW = index + 2;
    return obj;
  }).filter(function(draft) {
    const status = String(draft.STATUS || "").trim().toUpperCase();
    const draftCompany = String(draft.COMPANY_ID || "").trim().toUpperCase();
    if (status !== "READY FOR BILLING") return false;
    return !effectiveCompanyId || draftCompany === effectiveCompanyId;
  }).sort(function(a, b) {
    const timeA = a.SUBMITTED_AT instanceof Date ? a.SUBMITTED_AT.getTime() : 0;
    const timeB = b.SUBMITTED_AT instanceof Date ? b.SUBMITTED_AT.getTime() : 0;
    return timeB - timeA;
  }).map(serializeInvoiceDraftForClient_);
}

function getPendingInvoiceDraftCount(companyId, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);
  const role = String(session.role || "").trim().toUpperCase();
  const effectiveCompanyId = companyId ||
    (role === "OWNER" ? "" : String(session.companyId || "").trim().toUpperCase());

  const sh = ensureInvoiceDraftSheet_();
  if (sh.getLastRow() < 2) return 0;

  const headers = getInvoiceDraftHeaders_(sh);
  const idxStatus = headers.indexOf("STATUS");
  const idxCompany = headers.indexOf("COMPANY_ID");
  if (idxStatus === -1) return 0;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  return values.reduce(function(total, row) {
    const status = String(row[idxStatus] || "").trim().toUpperCase();
    const draftCompany = idxCompany >= 0
      ? String(row[idxCompany] || "").trim().toUpperCase()
      : "";

    if (status !== "READY FOR BILLING") return total;
    if (effectiveCompanyId && draftCompany !== effectiveCompanyId) return total;
    return total + 1;
  }, 0);
}

function finalizeInvoiceDraft(data, sessionToken) {
  data = data || {};
  const draftId = String(data.DRAFT_ID || "").trim();
  if (!draftId) throw new Error("No llego el reporte pendiente.");

  const finalHours = validateInvoiceDraftHours_(data.FINAL_HOURS);
  const items = [];
  let partsTotal = 0;

  for (let i = 1; i <= 4; i++) {
    const item = normalizeInvoiceDraftItem_(data, i);
    items.push(item);
    partsTotal += item.amount;
  }
  partsTotal = roundCloseOrderMoney_(partsTotal);

  const processingLock = LockService.getScriptLock();
  processingLock.waitLock(30000);

  let draft = null;
  let session = null;
  let originalWorkPerformed = "";
  let approvedWorkPerformed = "";
  let workReportChanged = false;
  let originalInvoiceDetails = null;
  let approvedInvoiceDetails = null;
  let invoiceDetailsChanged = false;

  try {
    const sh = ensureInvoiceDraftSheet_();
    const headers = getInvoiceDraftHeaders_(sh);
    const found = findInvoiceDraft_(sh, headers, draftId, "", "");
    if (!found) throw new Error("No se encontro el reporte pendiente.");

    draft = found.data;
    const companyId = String(draft.COMPANY_ID || "").trim().toUpperCase();
    session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);

    const existingInvoice = findInvoiceNumberForWorkOrder_(companyId, draft.WO_NUMBER);
    if (existingInvoice) {
      updateInvoiceDraftStatus_(
        sh,
        headers,
        found.rowNumber,
        "INVOICED",
        session,
        existingInvoice,
        finalHours
      );

      return {
        success: true,
        alreadyCreated: true,
        invoiceNumber: existingInvoice,
        woNumber: draft.WO_NUMBER
      };
    }

    const status = String(draft.STATUS || "").trim().toUpperCase();
    if (status !== "READY FOR BILLING") {
      throw new Error(
        status === "PROCESSING"
          ? "Otro usuario esta creando este invoice."
          : "Este reporte ya no esta pendiente."
      );
    }

    const hasEditedWorkReport = Object.prototype.hasOwnProperty.call(data, "WORK_PERFORMED");
    originalWorkPerformed = String(draft.WORK_PERFORMED || "").trim();
    approvedWorkPerformed = hasEditedWorkReport
      ? String(data.WORK_PERFORMED || "").trim()
      : originalWorkPerformed;

    if (!approvedWorkPerformed) {
      throw new Error("El reporte completo para el invoice no puede quedar vacio.");
    }
    if (approvedWorkPerformed.length > 12000) {
      throw new Error("El reporte completo no puede exceder 12,000 caracteres.");
    }
    workReportChanged = approvedWorkPerformed !== originalWorkPerformed;

    originalInvoiceDetails = {
      equipmentMake: String(draft.EQUIPMENT_MAKE || "").trim(),
      equipmentModel: String(draft.EQUIPMENT_MODEL || "").trim(),
      equipmentSerial: String(draft.EQUIPMENT_SERIAL || "").trim(),
      signature: String(draft.SIGNATURE || "").trim()
    };
    approvedInvoiceDetails = {
      equipmentMake: normalizeInvoiceDraftEditableText_(
        data,
        "EQUIPMENT_MAKE",
        originalInvoiceDetails.equipmentMake,
        200,
        "La marca del equipo"
      ),
      equipmentModel: normalizeInvoiceDraftEditableText_(
        data,
        "EQUIPMENT_MODEL",
        originalInvoiceDetails.equipmentModel,
        200,
        "El modelo del equipo"
      ),
      equipmentSerial: normalizeInvoiceDraftEditableText_(
        data,
        "EQUIPMENT_SERIAL",
        originalInvoiceDetails.equipmentSerial,
        250,
        "El numero de serie"
      ),
      signature: normalizeInvoiceDraftEditableText_(
        data,
        "SIGNATURE",
        originalInvoiceDetails.signature,
        300,
        "El nombre o la firma del manager",
        true
      )
    };
    invoiceDetailsChanged = Object.keys(originalInvoiceDetails).some(function(key) {
      return originalInvoiceDetails[key] !== approvedInvoiceDetails[key];
    });

    const techHours = Number(draft.TECH_HOURS || 0);
    const hoursChanged = Math.abs(finalHours - techHours) > 0.0001;
    const changeReason = String(data.HOURS_CHANGE_REASON || "").trim();
    if (hoursChanged && !changeReason) {
      throw new Error("Escribe el motivo del cambio de horas.");
    }

    draft.STATUS = "PROCESSING";
    draft.FINAL_HOURS = finalHours;
    draft.WORK_PERFORMED = approvedWorkPerformed;
    draft.EQUIPMENT_MAKE = approvedInvoiceDetails.equipmentMake;
    draft.EQUIPMENT_MODEL = approvedInvoiceDetails.equipmentModel;
    draft.EQUIPMENT_SERIAL = approvedInvoiceDetails.equipmentSerial;
    draft.SIGNATURE = approvedInvoiceDetails.signature;
    draft.UPDATED_AT = new Date();
    draft.UPDATED_BY = getSessionActorLabel_(session);
    writeInvoiceDraft_(sh, headers, found.rowNumber, draft);
  } finally {
    processingLock.releaseLock();
  }

  try {
    const workOrder = findWorkOrderForInvoiceDraft_(draft.COMPANY_ID, draft.WO_NUMBER);
    const payload = buildFinalInvoicePayload_(draft, data, items, partsTotal, finalHours, workOrder);
    payload.sessionToken = sessionToken;

    const result = saveCloseOrder_(payload);

    const completionLock = LockService.getScriptLock();
    completionLock.waitLock(30000);
    try {
      const sh = ensureInvoiceDraftSheet_();
      const headers = getInvoiceDraftHeaders_(sh);
      const found = findInvoiceDraft_(sh, headers, draftId, "", "");
      if (found) {
        updateInvoiceDraftStatus_(
          sh,
          headers,
          found.rowNumber,
          "INVOICED",
          session,
          result.invoiceNumber || "",
          finalHours
        );
      }
    } finally {
      completionLock.releaseLock();
    }

    const techHours = Number(draft.TECH_HOURS || 0);
    if (Math.abs(finalHours - techHours) > 0.0001) {
      addAuditLog_(
        "INVOICE",
        "INVOICE_HOURS_ADJUSTED",
        draft.COMPANY_ID,
        "INVOICE_DRAFT",
        draftId,
        session,
        {
          woNumber: draft.WO_NUMBER,
          originalHours: techHours,
          finalHours: finalHours,
          reason: String(data.HOURS_CHANGE_REASON || "").trim()
        }
      );
    }

    if (workReportChanged) {
      addAuditLog_(
        "INVOICE",
        "INVOICE_REPORT_EDITED",
        draft.COMPANY_ID,
        "INVOICE_DRAFT",
        draftId,
        session,
        {
          woNumber: draft.WO_NUMBER,
          previousReport: originalWorkPerformed,
          approvedReport: approvedWorkPerformed
        }
      );
    }

    if (invoiceDetailsChanged) {
      addAuditLog_(
        "INVOICE",
        "INVOICE_DETAILS_EDITED",
        draft.COMPANY_ID,
        "INVOICE_DRAFT",
        draftId,
        session,
        {
          woNumber: draft.WO_NUMBER,
          previousDetails: originalInvoiceDetails,
          approvedDetails: approvedInvoiceDetails
        }
      );
    }

    addAuditLog_(
      "INVOICE",
      "INVOICE_DRAFT_FINALIZED",
      draft.COMPANY_ID,
      "INVOICE_DRAFT",
      draftId,
      session,
      {
        woNumber: draft.WO_NUMBER,
        invoiceNumber: result.invoiceNumber || "",
        technicianHours: Number(draft.TECH_HOURS || 0),
        finalHours: finalHours,
        partsTotal: partsTotal
      }
    );

    touchAppCacheVersion_();
    return result;
  } catch (err) {
    const recoveryLock = LockService.getScriptLock();
    recoveryLock.waitLock(30000);
    try {
      const sh = ensureInvoiceDraftSheet_();
      const headers = getInvoiceDraftHeaders_(sh);
      const found = findInvoiceDraft_(sh, headers, draftId, "", "");
      if (found && String(found.data.STATUS || "").trim().toUpperCase() === "PROCESSING") {
        found.data.STATUS = "READY FOR BILLING";
        found.data.UPDATED_AT = new Date();
        found.data.UPDATED_BY = getSessionActorLabel_(session);
        writeInvoiceDraft_(sh, headers, found.rowNumber, found.data);
      }
    } finally {
      recoveryLock.releaseLock();
    }

    notifySystemError_("INVOICE_DRAFT_FINALIZE_ERROR", err, {
      module: "INVOICE",
      companyId: draft && draft.COMPANY_ID,
      woNumber: draft && draft.WO_NUMBER,
      draftId: draftId
    });
    throw err;
  }
}

function buildFinalInvoicePayload_(draft, data, items, partsTotal, finalHours, workOrder) {
  const payload = {
    FINALIZE_INVOICE: "YES",
    ROW_NUMBER: workOrder.rowNumber,
    WO_NUMBER: draft.WO_NUMBER,
    COMPANY_ID: draft.COMPANY_ID,
    WO_TYPE: "REPAIR_FORM",
    NSN: draft.NSN,
    CLIENT: draft.CLIENT,
    INVOICE_DATE_MANUAL: draft.INVOICE_DATE,
    STORE_ADDRESS: draft.STORE_ADDRESS,
    STORE_STREET: draft.STORE_STREET,
    STORE_CITY: draft.STORE_CITY,
    STORE_STATE: draft.STORE_STATE,
    STORE_ZIP: draft.STORE_ZIP,
    TECHNICIAN: draft.TECHNICIAN,
    TECHNICIAN_EMAIL: draft.TECHNICIAN_EMAIL,
    REPORTED_PROBLEM: draft.REPORTED_PROBLEM,
    TECH_PROBLEM_FOUND: draft.TECH_PROBLEM_FOUND,
    TECH_PROPOSED_SOLUTION: draft.TECH_PROPOSED_SOLUTION,
    WORK_COMPLETION_RESULT: draft.WORK_COMPLETION_RESULT,
    HORAS: finalHours,
    PARTS_TOTAL: partsTotal,
    PURCHASE_RECEIPT: String(data.PURCHASE_RECEIPT || "").trim(),
    PURCHASE_LOCATION: String(data.PURCHASE_LOCATION || "").trim(),
    WORK_PERFORMED: draft.WORK_PERFORMED,
    APPROVED_WORK_PERFORMED: draft.WORK_PERFORMED,
    EQUIPMENT_MAKE: draft.EQUIPMENT_MAKE,
    EQUIPMENT_MODEL: draft.EQUIPMENT_MODEL,
    EQUIPMENT_SERIAL: draft.EQUIPMENT_SERIAL,
    NOTES: draft.NOTES,
    SIGNATURE: draft.SIGNATURE,
    SPECIAL_INVOICE: String(data.SPECIAL_INVOICE || "").trim().toUpperCase() === "YES" ? "YES" : "",
    MANUAL_PARTS_TOTAL: Number(data.MANUAL_PARTS_TOTAL || 0),
    MANUAL_SUB_TOTAL: Number(data.MANUAL_SUB_TOTAL || 0)
  };

  items.forEach(function(item, index) {
    const n = index + 1;
    payload["I" + n + "_QTY"] = item.qty || "";
    payload["I" + n + "_PART"] = item.part;
    payload["I" + n + "_DESC"] = item.desc;
    payload["I" + n + "_UNIT"] = item.unit;
    payload["I" + n + "_AMOUNT"] = item.amount;
  });

  return payload;
}

function normalizeInvoiceDraftItem_(data, number) {
  const prefix = "I" + number + "_";
  const rawQty = data[prefix + "QTY"];
  const rawUnit = data[prefix + "UNIT"];
  const part = String(data[prefix + "PART"] || "").trim();
  const desc = String(data[prefix + "DESC"] || "").trim();
  const hasValues = String(rawQty || "").trim() ||
    String(rawUnit || "").trim() ||
    part ||
    desc;

  if (!hasValues) {
    return { qty: 0, part: "", desc: "", unit: 0, amount: 0 };
  }

  const qty = Number(rawQty);
  const unit = Number(rawUnit);
  if (!isFinite(qty) || qty <= 0) {
    throw new Error("La cantidad del item " + number + " debe ser mayor que 0.");
  }
  if (!isFinite(unit) || unit < 0) {
    throw new Error("El precio del item " + number + " no es valido.");
  }
  if (!part && !desc) {
    throw new Error("Escribe el numero de parte o la descripcion del item " + number + ".");
  }

  return {
    qty: qty,
    part: part,
    desc: desc,
    unit: roundCloseOrderMoney_(unit),
    amount: roundCloseOrderMoney_(qty * unit)
  };
}

function normalizeInvoiceDraftEditableText_(data, key, fallback, maxLength, label, required) {
  const hasValue = Object.prototype.hasOwnProperty.call(data, key);
  const value = String(hasValue ? data[key] : (fallback || "")).trim();

  if (required && !value) {
    throw new Error(label + " es obligatorio.");
  }
  if (value.length > maxLength) {
    throw new Error(label + " no puede exceder " + maxLength + " caracteres.");
  }
  return value;
}

function validateInvoiceDraftHours_(value) {
  const hours = Number(value);
  if (!isFinite(hours) || hours < 0) {
    throw new Error("Las horas trabajadas no pueden ser negativas.");
  }
  if (Math.abs(hours * 4 - Math.round(hours * 4)) > 0.0001) {
    throw new Error("Las horas deben escribirse en incrementos de 0.25.");
  }
  return Math.round(hours * 100) / 100;
}

function normalizeInvoiceDraftMaterialQty_(value, itemNumber) {
  if (value === "" || value === undefined || value === null) return "";

  const qty = Number(value);
  if (!isFinite(qty) || qty <= 0) {
    throw new Error("La cantidad del material " + itemNumber + " debe ser mayor que 0.");
  }
  return Math.round(qty * 100) / 100;
}

function normalizeTechnicalMaterial_(data, itemNumber) {
  const prefix = "I" + itemNumber + "_";
  const rawQty = data[prefix + "QTY"];
  const part = String(data[prefix + "PART"] || "").trim();
  const desc = String(data[prefix + "DESC"] || "").trim();
  const hasMaterial = String(rawQty || "").trim() || part || desc;

  if (!hasMaterial) {
    return { qty: "", part: "", desc: "" };
  }

  if (!String(rawQty || "").trim()) {
    throw new Error("Escriba la cantidad del material " + itemNumber + ".");
  }

  const qty = normalizeInvoiceDraftMaterialQty_(rawQty, itemNumber);
  if (!part && !desc) {
    throw new Error(
      "Escriba el numero de parte o la descripcion del material " + itemNumber + "."
    );
  }

  return { qty: qty, part: part, desc: desc };
}

function findWorkOrderForInvoiceDraft_(companyId, woNumber) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe WORK_ORDERS.");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("No se encontro la orden.");

  const headers = values[0].map(String);
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxWO = headers.indexOf("WO_NUMBER");
  const targetCompany = String(companyId || "").trim().toUpperCase();
  const targetWO = String(woNumber || "").trim().toUpperCase();

  for (let i = 1; i < values.length; i++) {
    const rowCompany = idxCompany >= 0
      ? String(values[i][idxCompany] || "").trim().toUpperCase()
      : String(CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
    const rowWO = idxWO >= 0
      ? String(values[i][idxWO] || "").trim().toUpperCase()
      : "";

    if (rowCompany === targetCompany && rowWO === targetWO && !isSoftDeletedRow_(values[i], headers)) {
      return {
        rowNumber: i + 1,
        rowData: values[i],
        headers: headers
      };
    }
  }

  throw new Error("No se encontro la orden " + woNumber + " para crear el invoice.");
}

function findInvoiceNumberForWorkOrder_(companyId, woNumber) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INVOICES");
  if (!sh || sh.getLastRow() < 2) return "";

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(function(h) {
    return String(h || "").trim().toUpperCase();
  });
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxWO = headers.indexOf("WO_NUMBER");
  let idxInvoice = headers.indexOf("INVOICE_NUMBER");
  if (idxInvoice === -1) idxInvoice = headers.indexOf("INVOICE");

  const targetCompany = String(companyId || "").trim().toUpperCase();
  const targetWO = String(woNumber || "").trim().toUpperCase();

  for (let i = 1; i < values.length; i++) {
    const rowCompany = idxCompany >= 0
      ? String(values[i][idxCompany] || "").trim().toUpperCase()
      : targetCompany;
    const rowWO = idxWO >= 0 ? String(values[i][idxWO] || "").trim().toUpperCase() : "";
    if (rowCompany === targetCompany && rowWO === targetWO) {
      return idxInvoice >= 0 ? String(values[i][idxInvoice] || "").trim() : "EXISTING";
    }
  }

  return "";
}

function getInvoiceDraftHeaders_(sh) {
  return sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });
}

function findInvoiceDraft_(sh, headers, draftId, companyId, woNumber) {
  if (sh.getLastRow() < 2) return null;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  const idxId = headers.indexOf("DRAFT_ID");
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxWO = headers.indexOf("WO_NUMBER");
  const targetId = String(draftId || "").trim();
  const targetCompany = String(companyId || "").trim().toUpperCase();
  const targetWO = String(woNumber || "").trim().toUpperCase();

  for (let i = 0; i < values.length; i++) {
    const rowId = idxId >= 0 ? String(values[i][idxId] || "").trim() : "";
    const rowCompany = idxCompany >= 0
      ? String(values[i][idxCompany] || "").trim().toUpperCase()
      : "";
    const rowWO = idxWO >= 0 ? String(values[i][idxWO] || "").trim().toUpperCase() : "";
    const matchesId = targetId && rowId === targetId;
    const matchesOrder = !targetId && rowCompany === targetCompany && rowWO === targetWO;

    if (matchesId || matchesOrder) {
      return {
        rowNumber: i + 2,
        data: invoiceDraftRowToObject_(headers, values[i])
      };
    }
  }

  return null;
}

function invoiceDraftRowToObject_(headers, row) {
  const obj = {};
  headers.forEach(function(header, index) {
    obj[header] = row[index];
  });
  return obj;
}

function writeInvoiceDraft_(sh, headers, rowNumber, draft) {
  const values = headers.map(function(header) {
    return draft[header] !== undefined ? draft[header] : "";
  });

  if (rowNumber && rowNumber >= 2) {
    sh.getRange(rowNumber, 1, 1, headers.length).setValues([values]);
    return rowNumber;
  }

  sh.getRange(sh.getLastRow() + 1, 1, 1, headers.length).setValues([values]);
  return sh.getLastRow();
}

function updateInvoiceDraftStatus_(sh, headers, rowNumber, status, session, invoiceNumber, finalHours) {
  const current = invoiceDraftRowToObject_(
    headers,
    sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0]
  );
  current.STATUS = status;
  current.FINAL_HOURS = finalHours;
  current.INVOICE_NUMBER = invoiceNumber || current.INVOICE_NUMBER || "";
  current.UPDATED_AT = new Date();
  current.UPDATED_BY = getSessionActorLabel_(session);

  if (status === "INVOICED") {
    current.FINALIZED_AT = new Date();
    current.FINALIZED_BY = getSessionActorLabel_(session);
  }

  writeInvoiceDraft_(sh, headers, rowNumber, current);
}

function serializeInvoiceDraftForClient_(draft) {
  const result = {};

  Object.keys(draft || {}).forEach(function(key) {
    const value = draft[key];
    result[key] = value instanceof Date
      ? Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a")
      : value === undefined || value === null ? "" : String(value);
  });

  const rawInvoiceDate = draft && draft.INVOICE_DATE;
  if (rawInvoiceDate instanceof Date && !isNaN(rawInvoiceDate.getTime())) {
    result.INVOICE_DATE_ISO = Utilities.formatDate(rawInvoiceDate, CFG.TIMEZONE, "yyyy-MM-dd");
  } else {
    result.INVOICE_DATE_ISO = String(rawInvoiceDate || "");
  }

  return result;
}
