// =====================================================
// FILE: 30_Equipment_Manuals.gs
// Technical manuals library
// =====================================================

const EQUIPMENT_MANUAL_COLUMNS_ = [
  "ID",
  "COMPANY_ID",
  "EQUIPMENT_TYPE",
  "BRAND",
  "MODEL",
  "SERIAL",
  "TITLE",
  "MANUAL_URL",
  "SOURCE",
  "NOTES",
  "ACTIVE",
  "CREATED_AT",
  "CREATED_BY",
  "UPDATED_AT",
  "UPDATED_BY"
];

function getEquipmentManuals(companyId, sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "TECH", "SYSTEM"], companyId);
  const activeCompanyId = resolveEquipmentManualCompanyId_(companyId, session);
  const ctx = getEquipmentManualSheetContext_();
  const lastRow = ctx.sheet.getLastRow();
  const rows = lastRow > 1
    ? ctx.sheet.getRange(2, 1, lastRow - 1, ctx.headers.length).getValues()
    : [];

  const manuals = rows.map(function(row, index) {
    return toEquipmentManualObject_(row, ctx.idx, index + 2);
  }).filter(function(manual) {
    const manualCompanyId = String(manual.companyId || "").trim().toUpperCase();
    return manual.active !== "NO" &&
      (!manualCompanyId || manualCompanyId === "ALL" || manualCompanyId === activeCompanyId);
  }).sort(function(a, b) {
    return String(a.title || "").localeCompare(String(b.title || ""));
  });

  return {
    companyId: activeCompanyId,
    companyName: getEquipmentManualCompanyName_(activeCompanyId),
    canEdit: isEquipmentManualManager_(session),
    manuals: manuals
  };
}

function saveEquipmentManual(payload, sessionToken) {
  const raw = payload || {};
  const requestedCompany = raw.companyId || raw.COMPANY_ID || "";
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "SYSTEM"], requestedCompany);
  const activeCompanyId = resolveEquipmentManualCompanyId_(requestedCompany, session);
  const manual = sanitizeEquipmentManualPayload_(raw, activeCompanyId);
  const actor = getSessionActorLabel_(session);
  const now = new Date();
  const ctx = getEquipmentManualSheetContext_();
  const existingRow = manual.id ? findEquipmentManualRow_(ctx, manual.id, activeCompanyId) : 0;

  if (existingRow) {
    const current = ctx.sheet.getRange(existingRow, 1, 1, ctx.headers.length).getValues()[0];
    manual.createdAt = current[ctx.idx.CREATED_AT] || now;
    manual.createdBy = current[ctx.idx.CREATED_BY] || actor;
    manual.updatedAt = now;
    manual.updatedBy = actor;
    ctx.sheet.getRange(existingRow, 1, 1, ctx.headers.length).setValues([
      buildEquipmentManualRow_(ctx.headers, manual)
    ]);
    addAuditLog_("EQUIPMENT_MANUALS", "UPDATE", activeCompanyId, "MANUAL", manual.id, session, manual);
  } else {
    manual.id = "MAN-" + Utilities.getUuid().replace(/-/g, "").substring(0, 10).toUpperCase();
    manual.createdAt = now;
    manual.createdBy = actor;
    manual.updatedAt = now;
    manual.updatedBy = actor;
    ctx.sheet.appendRow(buildEquipmentManualRow_(ctx.headers, manual));
    addAuditLog_("EQUIPMENT_MANUALS", "CREATE", activeCompanyId, "MANUAL", manual.id, session, manual);
  }

  touchAppCacheVersion_();
  return getEquipmentManuals(activeCompanyId, sessionToken);
}

function deleteEquipmentManual(manualId, companyId, sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "SYSTEM"], companyId);
  const activeCompanyId = resolveEquipmentManualCompanyId_(companyId, session);
  const ctx = getEquipmentManualSheetContext_();
  const id = String(manualId || "").trim();
  const rowNumber = findEquipmentManualRow_(ctx, id, activeCompanyId);

  if (!rowNumber) {
    throw new Error("Manual no encontrado.");
  }

  ctx.sheet.getRange(rowNumber, ctx.idx.ACTIVE + 1).setValue("NO");
  ctx.sheet.getRange(rowNumber, ctx.idx.UPDATED_AT + 1).setValue(new Date());
  ctx.sheet.getRange(rowNumber, ctx.idx.UPDATED_BY + 1).setValue(getSessionActorLabel_(session));

  addAuditLog_("EQUIPMENT_MANUALS", "DEACTIVATE", activeCompanyId, "MANUAL", id, session, {});
  touchAppCacheVersion_();

  return getEquipmentManuals(activeCompanyId, sessionToken);
}

function getEquipmentManualSheetContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CFG.SHEET_EQUIPMENT_MANUALS || "EQUIPMENT_MANUALS";
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  const headers = ensureSheetColumns_(sh, EQUIPMENT_MANUAL_COLUMNS_);
  const idx = buildEquipmentManualIndex_(headers);

  return { spreadsheet: ss, sheet: sh, headers: headers, idx: idx };
}

function buildEquipmentManualIndex_(headers) {
  const idx = {};
  headers.forEach(function(header, index) {
    idx[String(header || "").trim().toUpperCase()] = index;
  });
  return idx;
}

function resolveEquipmentManualCompanyId_(companyId, session) {
  const requested = String(companyId || "").trim().toUpperCase();
  const sessionCompany = String((session || {}).companyId || "").trim().toUpperCase();
  if (requested) return requested;
  if (sessionCompany) return sessionCompany;
  return String(CFG.DEFAULT_COMPANY_ID || "PPS").trim().toUpperCase();
}

function getEquipmentManualCompanyName_(companyId) {
  const id = String(companyId || "").trim().toUpperCase();
  const branding = CFG.COMPANY_BRANDING && CFG.COMPANY_BRANDING[id];
  if (branding) return branding.companyName || branding.name || id;
  if (id === String(CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase()) return CFG.APP_NAME || id;
  return id;
}

function isEquipmentManualManager_(session) {
  const role = String((session || {}).role || "").trim().toUpperCase();
  return role === "OWNER" || role === "ADMIN" || role === "SYSTEM";
}

function sanitizeEquipmentManualPayload_(payload, companyId) {
  const manual = {
    id: String(payload.id || payload.ID || "").trim(),
    companyId: String(companyId || payload.companyId || payload.COMPANY_ID || "").trim().toUpperCase(),
    equipmentType: String(payload.equipmentType || payload.EQUIPMENT_TYPE || "").trim(),
    brand: String(payload.brand || payload.BRAND || "").trim(),
    model: String(payload.model || payload.MODEL || "").trim(),
    serial: String(payload.serial || payload.SERIAL || "").trim(),
    title: String(payload.title || payload.TITLE || "").trim(),
    manualUrl: String(payload.manualUrl || payload.MANUAL_URL || "").trim(),
    source: String(payload.source || payload.SOURCE || "").trim(),
    notes: String(payload.notes || payload.NOTES || "").trim(),
    active: "YES"
  };

  if (!manual.companyId) {
    manual.companyId = String(CFG.DEFAULT_COMPANY_ID || "PPS").trim().toUpperCase();
  }
  if (!manual.title) {
    throw new Error("Debe escribir el titulo del manual.");
  }
  if (!manual.manualUrl) {
    throw new Error("Debe escribir el URL del manual.");
  }
  if (!/^https?:\/\//i.test(manual.manualUrl)) {
    throw new Error("El URL del manual debe comenzar con http:// o https://.");
  }

  return manual;
}

function findEquipmentManualRow_(ctx, id, companyId) {
  id = String(id || "").trim();
  companyId = String(companyId || "").trim().toUpperCase();
  if (!id) return 0;

  const lastRow = ctx.sheet.getLastRow();
  if (lastRow < 2) return 0;

  const rows = ctx.sheet.getRange(2, 1, lastRow - 1, ctx.headers.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    const rowId = String(rows[i][ctx.idx.ID] || "").trim();
    const rowCompanyId = String(rows[i][ctx.idx.COMPANY_ID] || "").trim().toUpperCase();
    if (rowId === id && (!rowCompanyId || rowCompanyId === companyId)) {
      return i + 2;
    }
  }

  return 0;
}

function toEquipmentManualObject_(row, idx, rowNumber) {
  return {
    rowNumber: rowNumber,
    id: String(row[idx.ID] || "").trim(),
    companyId: String(row[idx.COMPANY_ID] || "").trim().toUpperCase(),
    equipmentType: String(row[idx.EQUIPMENT_TYPE] || "").trim(),
    brand: String(row[idx.BRAND] || "").trim(),
    model: String(row[idx.MODEL] || "").trim(),
    serial: String(row[idx.SERIAL] || "").trim(),
    title: String(row[idx.TITLE] || "").trim(),
    manualUrl: String(row[idx.MANUAL_URL] || "").trim(),
    source: String(row[idx.SOURCE] || "").trim(),
    notes: String(row[idx.NOTES] || "").trim(),
    active: String(row[idx.ACTIVE] || "YES").trim().toUpperCase() || "YES",
    createdAt: formatEquipmentManualDate_(row[idx.CREATED_AT]),
    createdBy: String(row[idx.CREATED_BY] || "").trim(),
    updatedAt: formatEquipmentManualDate_(row[idx.UPDATED_AT]),
    updatedBy: String(row[idx.UPDATED_BY] || "").trim()
  };
}

function buildEquipmentManualRow_(headers, manual) {
  return headers.map(function(header) {
    switch (String(header || "").trim().toUpperCase()) {
      case "ID": return manual.id || "";
      case "COMPANY_ID": return manual.companyId || "";
      case "EQUIPMENT_TYPE": return manual.equipmentType || "";
      case "BRAND": return manual.brand || "";
      case "MODEL": return manual.model || "";
      case "SERIAL": return manual.serial || "";
      case "TITLE": return manual.title || "";
      case "MANUAL_URL": return manual.manualUrl || "";
      case "SOURCE": return manual.source || "";
      case "NOTES": return manual.notes || "";
      case "ACTIVE": return manual.active || "YES";
      case "CREATED_AT": return manual.createdAt || "";
      case "CREATED_BY": return manual.createdBy || "";
      case "UPDATED_AT": return manual.updatedAt || "";
      case "UPDATED_BY": return manual.updatedBy || "";
      default: return "";
    }
  });
}

function formatEquipmentManualDate_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
  }
  return String(value || "").trim();
}
