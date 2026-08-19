// =====================================================
// FILE: 22_Companies.gs
// =====================================================

const COMPANY_ADMIN_COLUMNS = [
  "COMPANY_ID",
  "COMPANY_NAME",
  "LEGAL_NAME",
  "EIN",
  "WEBSITE_URL",
  "MAIN_EMAIL",
  "PHONE",
  "ADDRESS",
  "CITY",
  "STATE",
  "ZIP",
  "LOGIN_TITLE",
  "LOGIN_SUBTITLE",
  "LOGIN_OWNER_NAME",
  "LOGIN_PRIMARY_COLOR",
  "LOGIN_ACCENT_COLOR",
  "LOGIN_LOGO_FILE_ID",
  "LOGIN_LOGO_URL",
  "LOGIN_BACKGROUND_FILE_ID",
  "LOGIN_BACKGROUND_URL",
  "NOTES",
  "ACTIVE",
  "DELETED_AT",
  "DELETED_BY"
];

function getCompaniesAdmin(sessionToken) {
  requireCompaniesOwner_(sessionToken);

  const sh = getCompaniesSheet_();
  const headers = ensureCompanyAdminColumns_(sh);
  const data = sh.getDataRange().getValues();

  if (data.length < 2) return [];

  return data.slice(1).map(function(row, i) {
    const obj = {};
    headers.forEach(function(h, c) {
      obj[h] = row[c];
    });
    obj.ROW_NUMBER = i + 2;
    return obj;
  });
}

function saveCompanyAdmin(rowNumber, data, sessionToken) {
  const session = requireCompaniesOwner_(sessionToken);

  data = data || {};

  const sh = getCompaniesSheet_();
  const headers = ensureCompanyAdminColumns_(sh);
  const companyId = String(data.COMPANY_ID || "").trim().toUpperCase();
  const companyName = String(data.COMPANY_NAME || "").trim();

  if (!companyId) throw new Error("COMPANY_ID es obligatorio.");
  if (!companyName) throw new Error("COMPANY_NAME es obligatorio.");

  rowNumber = Number(rowNumber || 0);
  const isExisting = rowNumber && rowNumber > 1;

  assertCompanyIdIsUnique_(sh, headers, companyId, isExisting ? rowNumber : 0);

  data.COMPANY_ID = companyId;
  data.COMPANY_NAME = companyName;
  data.ACTIVE = String(data.ACTIVE || "YES").trim().toUpperCase();

  const values = headers.map(function(h) {
    return data[h] !== undefined ? data[h] : "";
  });

  if (isExisting) {
    sh.getRange(rowNumber, 1, 1, headers.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }

  addAuditLog_("COMPANIES", isExisting ? "COMPANY_UPDATED" : "COMPANY_CREATED", companyId, "COMPANY", companyId, session, {
    rowNumber: isExisting ? rowNumber : sh.getLastRow(),
    companyName: companyName
  });

  touchAppCacheVersion_();
  return true;
}

function toggleCompanyActiveAdmin(rowNumber, sessionToken) {
  const session = requireCompaniesOwner_(sessionToken);

  rowNumber = Number(rowNumber || 0);
  if (!rowNumber || rowNumber < 2) throw new Error("Fila invalida.");

  const sh = getCompaniesSheet_();
  const headers = ensureCompanyAdminColumns_(sh);
  const idxActive = headers.indexOf("ACTIVE");

  if (idxActive === -1) throw new Error("COMPANIES debe tener ACTIVE.");

  const current = String(sh.getRange(rowNumber, idxActive + 1).getValue() || "").trim().toUpperCase();
  const next = current === "YES" ? "NO" : "YES";
  sh.getRange(rowNumber, idxActive + 1).setValue(next);

  const idxCompany = headers.indexOf("COMPANY_ID");
  const companyId = idxCompany >= 0 ? String(sh.getRange(rowNumber, idxCompany + 1).getValue() || "").trim().toUpperCase() : "";

  if (next === "NO") {
    setCellByHeader_(sh, rowNumber, headers, "DELETED_AT", new Date());
    setCellByHeader_(sh, rowNumber, headers, "DELETED_BY", getSessionActorLabel_(session));
  } else {
    setCellByHeader_(sh, rowNumber, headers, "DELETED_AT", "");
    setCellByHeader_(sh, rowNumber, headers, "DELETED_BY", "");
  }

  addAuditLog_("COMPANIES", "COMPANY_ACTIVE_TOGGLED", companyId, "COMPANY", companyId, session, {
    rowNumber: rowNumber,
    active: next
  });

  touchAppCacheVersion_();
  return true;
}

function getCompaniesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_COMPANIES);

  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_COMPANIES);
  }

  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, COMPANY_ADMIN_COLUMNS.length).setValues([COMPANY_ADMIN_COLUMNS]);
  }

  return sh;
}

function ensureCompanyAdminColumns_(sh) {
  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  if (!headers.some(Boolean)) {
    sh.getRange(1, 1, 1, COMPANY_ADMIN_COLUMNS.length).setValues([COMPANY_ADMIN_COLUMNS]);
    return COMPANY_ADMIN_COLUMNS.slice();
  }

  COMPANY_ADMIN_COLUMNS.forEach(function(columnName) {
    if (headers.indexOf(columnName) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(columnName);
      headers.push(columnName);
    }
  });

  return headers;
}

function assertCompanyIdIsUnique_(sh, headers, companyId, currentRowNumber) {
  const idxId = headers.indexOf("COMPANY_ID");
  if (idxId === -1) throw new Error("COMPANIES debe tener COMPANY_ID.");

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1;
    const existing = String(data[i][idxId] || "").trim().toUpperCase();

    if (existing === companyId && rowNumber !== Number(currentRowNumber || 0)) {
      throw new Error("Ya existe una empresa con COMPANY_ID " + companyId + ".");
    }
  }
}

function requireCompaniesOwner_(sessionToken) {
  return requireSession_(sessionToken, ["OWNER"]);
}
