// =====================================================
// FILE: 10_Users.gs
// =====================================================

function getUsers(companyId, role) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  const data = sh.getDataRange().getValues();

  const headers = data[0].map(h => String(h).trim());
  const idxCompany = headers.indexOf("COMPANY_ID");

  return data.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, c) => {
      if (String(h).trim().toUpperCase() !== "PASSWORD") {
        obj[h] = row[c];
      }
    });
    obj.ROW_NUMBER = i + 2;
    return obj;
  }).filter(u => {
    if (role === "OWNER") return true;
    return String(u.COMPANY_ID || "").toUpperCase() === companyId;
  });
}

function saveUser(rowNumber, data) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const isExisting = rowNumber && Number(rowNumber) > 1;
  const existingRow = isExisting
    ? sh.getRange(Number(rowNumber), 1, 1, headers.length).getValues()[0]
    : null;

  const values = headers.map((h, i) => {
    if (String(h).trim().toUpperCase() === "PASSWORD" && isExisting && !data[h]) {
      return existingRow[i] || "";
    }

    return data[h] || "";
  });

  if (isExisting) {
    sh.getRange(Number(rowNumber),1,1,headers.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }

  return true;
}

function deleteUser(rowNumber) {
  rowNumber = Number(rowNumber);

  if (!rowNumber || rowNumber < 2) {
    throw new Error("Fila inválida.");
  }

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  if (!sh) throw new Error("No existe la hoja USERS.");

  sh.deleteRow(rowNumber);

  return true;
}

function getCompaniesForUser(companyId, role) {
  companyId = String(companyId || "").trim().toUpperCase();
  role = String(role || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_COMPANIES);
  if (!sh) throw new Error("No existe la hoja COMPANIES.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(String);
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxActive = headers.indexOf("ACTIVE");

  return data.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i];
    });
    return obj;
  }).filter(function(c) {
    const active = idxActive >= 0 ? String(c.ACTIVE || "").toUpperCase() : "YES";
    if (active !== "YES") return false;

    if (role === "OWNER") return true;

    return String(c.COMPANY_ID || "").trim().toUpperCase() === companyId;
  });
}

function getFormConfig_(e) {
  const source = e.source;

  if (!source) {
    return {
      companyId: CFG.DEFAULT_COMPANY_ID,
      lang: "EN"
    };
  }

  const formId = source.getId();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FORM_CONFIG");
  if (!sh) {
    return {
      companyId: CFG.DEFAULT_COMPANY_ID,
      lang: "EN"
    };
  }

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);

  const idxForm = headers.indexOf("FORM_ID");
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxLang = headers.indexOf("LANG");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxForm] || "") === formId) {
      return {
        companyId: String(data[i][idxCompany] || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase(),
        lang: String(data[i][idxLang] || "EN").trim().toUpperCase()
      };
    }
  }

  return {
    companyId: CFG.DEFAULT_COMPANY_ID,
    lang: "EN"
  };
}

function getFormUrls(companyId) {
  companyId = String(companyId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FORM_CONFIG");
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(String);

  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxLang = headers.indexOf("LANG");
  const idxUrl = headers.indexOf("FORM_URL");

  if (idxCompany === -1 || idxLang === -1 || idxUrl === -1) {
    throw new Error("FORM_CONFIG debe tener COMPANY_ID, LANG y FORM_URL.");
  }

  return data.slice(1)
    .filter(r => String(r[idxCompany] || "").trim().toUpperCase() === companyId)
    .map(r => ({
      lang: String(r[idxLang] || "").trim().toUpperCase(),
      url: String(r[idxUrl] || "").trim()
    }))
    .filter(f => f.url);
}

function getCompanyName(companyId) {
  companyId = String(companyId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_COMPANIES);
  if (!sh) return companyId;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return companyId;

  const headers = data[0].map(String);

  const idxId = headers.indexOf("COMPANY_ID");
  const idxName = headers.indexOf("COMPANY_NAME");

  if (idxId === -1 || idxName === -1) return companyId;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId] || "").trim().toUpperCase() === companyId) {
      return data[i][idxName] || companyId;
    }
  }

  return companyId;
}
