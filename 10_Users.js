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

function deleteUserLegacyDisabled_(rowNumber) {
  rowNumber = Number(rowNumber);

  if (!rowNumber || rowNumber < 2) {
    throw new Error("Fila inválida.");
  }

  throw new Error("deleteUser requiere sesion segura.");
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

// Secure admin-user functions. These later declarations intentionally replace
// the legacy browser-role versions above while preserving the rest of this file.
function getUsers(companyId, role, sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN"], companyId);
  companyId = String(companyId || session.companyId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  if (!sh) throw new Error("No existe la hoja USERS.");

  ensureUserSecurityColumns_(sh);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());

  return data.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, c) => {
      if (!isSensitiveUserHeader_(h)) {
        obj[h] = row[c];
      }
    });
    obj.ROW_NUMBER = i + 2;
    return obj;
  }).filter(u => {
    if (session.role === "OWNER") return true;
    return String(u.COMPANY_ID || "").toUpperCase() === String(session.companyId || companyId).toUpperCase();
  });
}

function saveUser(rowNumber, data, sessionToken) {
  data = data || {};

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  if (!sh) throw new Error("No existe la hoja USERS.");

  const headers = ensureUserSecurityColumns_(sh);
  const normalizedHeaders = headers.map(function(h) {
    return String(h).trim().toUpperCase();
  });

  const isExisting = rowNumber && Number(rowNumber) > 1;
  const existingRow = isExisting
    ? sh.getRange(Number(rowNumber), 1, 1, headers.length).getValues()[0]
    : null;

  const idxCompany = normalizedHeaders.indexOf("COMPANY_ID");
  const idxRole = normalizedHeaders.indexOf("ROLE");
  const idxPassword = normalizedHeaders.indexOf("PASSWORD");
  const idxPasswordHash = normalizedHeaders.indexOf("PASSWORD_HASH");

  const targetCompany = String(
    data.COMPANY_ID ||
    (existingRow && idxCompany >= 0 ? existingRow[idxCompany] : "") ||
    ""
  ).trim().toUpperCase();

  const targetRole = String(
    data.ROLE ||
    (existingRow && idxRole >= 0 ? existingRow[idxRole] : "") ||
    ""
  ).trim().toUpperCase();

  const session = requireSession_(sessionToken, ["OWNER", "ADMIN"], targetCompany);

  if (session.role !== "OWNER" && targetRole === "OWNER") {
    throw new Error("Solo OWNER puede crear o modificar usuarios OWNER.");
  }

  if (session.role !== "OWNER" && targetCompany !== session.companyId) {
    throw new Error("No autorizado para esta compania.");
  }

  const newPassword = String(data.PASSWORD || "").trim();

  if (!isExisting && !newPassword) {
    throw new Error("Password requerido para usuario nuevo.");
  }

  const existingHash = existingRow && idxPasswordHash >= 0 ? String(existingRow[idxPasswordHash] || "").trim() : "";
  const existingPlainPassword = existingRow && idxPassword >= 0 ? String(existingRow[idxPassword] || "").trim() : "";
  const migratedHash = !newPassword && !existingHash && existingPlainPassword
    ? hashPassword_(existingPlainPassword)
    : "";
  const passwordHash = newPassword ? hashPassword_(newPassword) : (existingHash || migratedHash);
  const passwordChanged = !!(newPassword || migratedHash);

  const values = headers.map((h, i) => {
    const upper = String(h).trim().toUpperCase();

    if (upper === "PASSWORD") return "";
    if (upper === "PASSWORD_HASH") return passwordHash || "";

    if (upper === "PASSWORD_UPDATED_AT") {
      if (passwordChanged) return new Date();
      return isExisting && existingRow ? existingRow[i] || "" : "";
    }

    if (upper === "COMPANY_ID") return targetCompany;
    if (upper === "ROLE") return targetRole;

    return data[h] !== undefined ? data[h] : (isExisting && existingRow ? existingRow[i] || "" : "");
  });

  if (isExisting) {
    sh.getRange(Number(rowNumber), 1, 1, headers.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }

  addAuditLog_("USERS", isExisting ? "USER_UPDATED" : "USER_CREATED", targetCompany, "USER", data.EMAIL || "", session, {
    rowNumber: isExisting ? Number(rowNumber) : sh.getLastRow(),
    role: targetRole
  });

  touchAppCacheVersion_();
  return true;
}

function deleteUser(rowNumber, sessionToken) {
  rowNumber = Number(rowNumber);

  if (!rowNumber || rowNumber < 2) {
    throw new Error("Fila invalida.");
  }

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  if (!sh) throw new Error("No existe la hoja USERS.");

  const headers = ensureUserSecurityColumns_(sh);
  const normalizedHeaders = headers.map(function(h) {
    return String(h).trim().toUpperCase();
  });
  const row = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const idxCompany = normalizedHeaders.indexOf("COMPANY_ID");
  const idxRole = normalizedHeaders.indexOf("ROLE");
  const idxEmail = normalizedHeaders.indexOf("EMAIL");
  const targetCompany = idxCompany >= 0 ? String(row[idxCompany] || "").trim().toUpperCase() : "";
  const targetRole = idxRole >= 0 ? String(row[idxRole] || "").trim().toUpperCase() : "";
  const targetEmail = idxEmail >= 0 ? String(row[idxEmail] || "").trim() : "";
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN"], targetCompany);

  if (session.role !== "OWNER" && targetRole === "OWNER") {
    throw new Error("Solo OWNER puede eliminar usuarios OWNER.");
  }

  setCellByHeader_(sh, rowNumber, headers, "ACTIVE", "NO");
  setCellByHeader_(sh, rowNumber, headers, "DELETED_AT", new Date());
  setCellByHeader_(sh, rowNumber, headers, "DELETED_BY", getSessionActorLabel_(session));

  addAuditLog_("USERS", "USER_SOFT_DELETED", targetCompany, "USER", targetEmail, session, {
    rowNumber: rowNumber,
    role: targetRole
  });

  touchAppCacheVersion_();
  return true;
}

function getCompaniesForUser(companyId, role, sessionToken) {
  const session = requireSession_(sessionToken, [], companyId);
  companyId = String(session.companyId || companyId || "").trim().toUpperCase();
  role = String(session.role || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_COMPANIES);
  if (!sh) throw new Error("No existe la hoja COMPANIES.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(String);
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

function isSensitiveUserHeader_(headerName) {
  const header = String(headerName || "").trim().toUpperCase();
  return [
    "PASSWORD",
    "PASSWORD_HASH",
    "PASSWORD_SALT",
    "SESSION_TOKEN"
  ].indexOf(header) !== -1;
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

  return withAppCache_(["company-name", companyId], 300, function() {
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
  });
}

function getCompanyBranding(companyId) {
  companyId = String(companyId || "").trim().toUpperCase();

  const base = Object.assign(
    {},
    (CFG.COMPANY_BRANDING && CFG.COMPANY_BRANDING.DEFAULT) || {}
  );

  const configured = companyId && CFG.COMPANY_BRANDING && CFG.COMPANY_BRANDING[companyId]
    ? Object.assign({}, CFG.COMPANY_BRANDING[companyId])
    : {};

  const sheetBranding = getCompanyBrandingFromSheet_(companyId);
  const branding = Object.assign(base, configured, sheetBranding);

  branding.companyId = companyId || "";
  branding.companyName = branding.companyName || getCompanyName(companyId) || base.companyName || CFG.APP_NAME;
  branding.loginTitle = branding.loginTitle || branding.companyName;
  branding.loginSubtitle = branding.loginSubtitle || "";
  branding.ownerName = branding.ownerName || CFG.APP_NAME;
  branding.primaryColor = branding.primaryColor || "#111827";
  branding.accentColor = branding.accentColor || "#dc2626";
  branding.backgroundImageUrl = branding.backgroundImageUrl || driveImageUrl_(branding.backgroundFileId, 1800);
  branding.logoImageUrl = branding.logoImageUrl || driveImageDataUrl_(branding.logoFileId) || getInlineCompanyLogo_(companyId) || driveImageUrl_(branding.logoFileId, 600);

  return branding;
}

function getCompanyBrandingFromSheet_(companyId) {
  if (!companyId) return {};

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_COMPANIES);
  if (!sh) return {};

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return {};

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  const idxId = headers.indexOf("COMPANY_ID");
  if (idxId === -1) return {};

  const map = {
    COMPANY_NAME: "companyName",
    LOGIN_TITLE: "loginTitle",
    LOGIN_SUBTITLE: "loginSubtitle",
    LOGIN_OWNER_NAME: "ownerName",
    LOGIN_PRIMARY_COLOR: "primaryColor",
    LOGIN_ACCENT_COLOR: "accentColor",
    LOGIN_BACKGROUND_URL: "backgroundImageUrl",
    LOGIN_BACKGROUND_FILE_ID: "backgroundFileId",
    LOGIN_LOGO_URL: "logoImageUrl",
    LOGIN_LOGO_FILE_ID: "logoFileId"
  };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId] || "").trim().toUpperCase() !== companyId) continue;

    const branding = {};
    Object.keys(map).forEach(function(header) {
      const idx = headers.indexOf(header);
      if (idx >= 0 && data[i][idx] !== "") {
        branding[map[header]] = String(data[i][idx] || "").trim();
      }
    });

    return branding;
  }

  return {};
}

function getInlineCompanyLogo_(companyId) {
  companyId = String(companyId || "").trim().toUpperCase();
  if (companyId !== "PPS") return "";

  return "data:image/jpeg;base64," + [
    "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAoHBwkHBgoJCAkLCwoMDxkQDw4ODx4WFxIZJCAmJSMgIyIoLTkwKCo2KyIjMkQyNjs9QEBAJjBGS0U+Sjk/QD3/2wBDAQsLCw8NDx0QEB09KSMpPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT3/wAARCAF1AjADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD2AmjP0/Kg9aSgBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc",
    "/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qM/T8qSigBc/T8qAaSlHWgAPWkpT1pKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiimSMwwE+8aAH0UwLP3A/Sl2y+lADqKbsm9P5UbJvQUAOopNkvoKTZL6frQA6imGOfH/16Qx3Hr+ooAkoqLyrj1P5ijyrj1/UUAS0VD5dz/kikK3Y6AH8qAJ6KrGW6T70OfwpBfAHEkbKaALVFRx3EUn3XGfQ8GpKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAClHWkpR1oAD1pKU9aSgAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiorm5gs4TLdTRwxjqzsAKAJaillSBJbiU4jhQsx9gMmoIdQa6GbW3kI6h5h5QI9QD8xH4Vi/EXURpPw81aZW+eSHyFPqXIU/oTQB47a69faj515cahqAM8ruqpcuAAT0xmrI1Cc/8xDUv/At/wDGsWyXyrSJDxtQf41bVqYGiL6c9dQ1P/wLf/GnC8m/6COp/wDgW/8AjVBTUgNAFz7ZN/0EdT/8C3/xpRdT/wDQR1L/AMCn/wAap1Dd3i2aoArSzynbFCvLOeg4oAt3eqy2kaj7dqck0h2xRLcyFpGPQAZr",
    "fu9Gn8PaLFc6/rOptq1yu6KwgvXVIh6uc5OPbGTx71c0Hw9D4F03/hJvE4WfXJhi1tSciEkdB7+p7Dgc9eR1HUrnVr+W8vZPMnlOWPYegHoBXNiK3IrLc9vJ8s+tT9pUXuL8fL/Mupr92qgG7vDjj/j4f/Gnf8JBc/8AP1e/+BDf41j0teZZ9z7X6vS/lRr/APCQXP8Az9Xv/gQ3+NOTxLeRnKXt8pPpcN/jWNSU7PuxPD0nvFHV2vjvVrcgLqU5HpMiuP5ZrpdM+JBfCataRyRHrNb9vqp/ofwry+lSVo23IxBq41akNpHHXyrCVlZwS9NPyPeoEsdXtBc6bcI6N3U5GfQjqDUTy3WnNiT7nY9VP+FeS6D4hutKv1ubNwsvR4j9yUeh/wA8V6zba4usaINQsYTPGMrcWhHzj+8B/tDqPWu+hiFU0ejPkcyyueDfMtYvr+jLtvqEM/BO1/Q1arjry4jtLOLVbOY3ejzDd9oj5aAf7Q6lfU9V7jvWvp2rBkTc4kjYZVgcgj1BroueY4tK/Q2qKRWV1DKcg0tBIUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFKOtJSjrQAHrSUp60lABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRQTgZPSqN3qsNqjHcPlGSzHAA9fp70AXSQoyxAHqaytW8S6do1s093cRxRjjc5wCfQdyfYVweqfEC71i/GmeFrZr+8c4Eu392vuB3A/vHC/Wqkllp/h+8W48QzHxF4kJAjtQd0Nux6AjufbH4DrSlJRV2aUqM60uSCuzoj4n1jX7Z7rT1j0fR1Hzanfrt3D/pnHnn2JOKvaHpDXbrdW4udp/5iWoDdcS/9ckIxGvvgfTvUui+Gr3UbiLVfFbia4X5reyH+qtvQ7ehb+XvXXUk29SqkYU3yp3ffp8u5WhtYbKBhEpyfvOxLMx9STya8z+ON8Y9A",
    "0nTFbDXVz5jD1VR/i4/KvUZ/ugHua8L+MeofbfiBbWanKWFsNw9GbLH9CtUYnMK+amRqpRtVhGpgXENSKarI1JcXgtgiohluJCFiiUZLE8Dgf5NAE15eraIgVDLcSHbFEoyXPQcD3r0Twd4Mg8JWMvinxUQ+pBN6oeRbA9FA7uentnA7mrHw++Hw0JTr3iIo+qMu9VcjbaLj8t2Op7dB3Jzdc1G++I+vjStFwun2xLGV8hT23t/JR+P0zqT5VpudWEw/tpXk7RW78v8AN9Dl/EGu3XiLVXvLo4H3YoweI07Af1Pes0Cu6/4VFqv/AEF7T/v01H/Co9V7ava/9+jXA6FRu7Pr6WcYClBQg7JeRw2KMV3X/CotV/6C9r/36aj/AIVHqv8A0FrT/v01L6vU7Gn9u4P+b8DhTSGuxuPhT4giBMN1YTgdBuZSfzGP1rlNS06/0a6+zapaS20p+7uHDD1BHB/CplSnHdHTQzLDV3ywlqQE00mgmo2kArM6pTS3JA21sg4Irtfh7rzWXiCGJmxBenyZB23/AMDf0/GvP3mbtxVmwvHtxLKjEPEVlQ+hU5/pVxTjJSODFOGIpyovqn/wD0fWtYPw78fkyLu0DWh5s8WMiGTOHdR+IJHcE+lM11X8C6nFcWgM3h++O5UU58hjydh9O4HQjPpmrHxkt0vfCGn6ioBMNwjAj+7IpB/XbVDwTMfGHwy1HQZzvutPGLcnk7eWj/Igr9K9Oaurrc+Gw9RQlyz1i9/68jtNI1dJI45I5BJBKoZWXow9a31YMoKnIPSvDfAviB7O+Gl3LHyJ2/dZ/wCWcnp9D0+uK9g0i78xfKJ9xRCfMrhisO6FRwe3Q06KKKs5wooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApR1pKUdaAA9aSlPWkoAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAqG5u4rVcyNzjIXuar3uoiGT7P",
    "b4efGWPaMep/wrhvGXi+Dw6hTi51KUZjhY8KP7746D0Hf6c0Aa3iXxfb6PaeffylEbPlQp/rJT6KP6ngV55bDXvibevGrLY6NC26ZyT5UYHPzHjzHx26D2qn4f8ADt5401G41jXbxotMh5urxztzj/lmnYfhwPqav+I/FSXtsuj6JD9h0SD5UiQYMv8AtN/PH4nnpnUqKmrs7MHgqmLnyx26su3evWWhWZ0TwYjKJCFnv8ZmuW6fKfT6fhjqe18DeB10RF1DUlD6i4yqnkQg/wA29T+H1z/hv4MW1hi1rUUzO43W0bD7in+M+57eg/T0TNZ04OT55nXjMRToxeFw232n1b/yFoooroPIIn+adF7DmvmPxJqH9q+Ndavs5V7lkQ/7KnaP0Ar6N1zURpOi6nqJP/Hrbu49yFJA/PFfLVpnyQzcsxLE+tAF5DU6NVRDUjziEKApklc7Y41GSxpgWnuTEUjjjaa4lIWKJQSXJ4HAr1/4efDhdBC6xrgWbWJBuAPK2w9B/tY6nt0HqYvhp8OToarreuKH1aVcpG3ItlPb/ex19OnrVDx947a9eTSdIZ/soJW4nUH96f7qn+76nv8ATrE5qKuzow2HliKihHTz7E/irxDd+MdVTw54eO63Zv30w+64HUk/3B+p/XufD3h618M6QtpZLuf70kpGGlf1P+HYV4Pa311ZMxtJLmBnwGMTMhbHTOKsnXtUGN2oagM9M3Dj+tcca9nzSWp9JVyj2kVRo1Uorp1b7s94W2uCS0jAsTzzUgtnA+8K8E/t7Uz01G+/8CX/AMa6TwfYav4ilN5dapfwaRAf3kxuXHmkdVXnp6n8Bz01hiOd2UTgxOS/Voe0qVV9z1PWRC/94UGF/UfnRaP5kKsqlYsARhupHqanrpPCKpSVBnJ/A1R1fSbXxJpU2n3yA7hlHxyjdmHuK2KqwLm4LD7q5FDVxxk4u63PnC4hltLmW2nGJYXaNx/tKcH+VQO1afiWdLnx",
    "Pqk0JzG93Kyn1G7r+lY8jV5bj7zSPu4VZSpRlLdpDXanRSbbe6J/554/M1Azc1paBYNqms6fYKMm6ukU/wC4Dlj+WfyrRRvocc63Jeb6J/keqfE1RB8OIoWPKLbj8QyiuS+DF20PjOaAfduLNsj3VlI/mfzrd+NepCPTLazU8z3AyP8AZQEn9WWuc+DMbTeO2dfuw2Uhb8WUCvQPjzB8Tw/2f4s1OKJtvlXchQjt8xIr1zw7qhu7axvBwZ41dh/td/1BryXxlMs3jLV3Q5U3UgH4Niu78BylvDViW/hlZPw8z/69YUviaPVxyvRhJ7/8A9UByMjoaKZAcwR5/uin1ueUFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFKOtJSjrQAHrSUp60lABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAVn6lfvEVtrX5rmTgf7PuasX12tnbs7HBxn6e9cvrutw+EtFm1S9XfeTfJFCTyzHkJn/AMeY/wD1qAMzxh4og8H6aLa2ZZ9VuAWQMM4/6aP7eg7/AJ15z4a8OXHi/Vp7i+uXS0iPnX97I3IHXGT/ABEfkOfSs+KHUvFfiFULme/vZMs7dF9SfRVHb0FdR4n1C306yj8KaKxFnanN7MOs8vcE/Xr74H8NTOSgrs3w9CWIqKnHqReJvE6apFFpukxfZdFtPlghAx5mP42/oPxPNXPh/wCFv+Eg1jzblM2NqQ0uejt2T+p9vrXLQQPcTJFEhaSRgqKO5PAAr6C8MaHH4e0K3sUwZAN0zD+Jz1P9PoBXHTTqz5pH0eNqRy/CqlS3f9NmoeOBwBSc5p5BzSY5ruPlRATmnZ4z6UbaZMdsLH2oA4L4v6l9h+HVxGGw97MkI9xncf0U14VEu2NV9AK9N+Od9un0LSwfuq9w4+pAH8mrzWPk0ABlEeAFLyMcIijJY17P8NPhqdIKa5r8YfU3GYYGGRbD1P8Atfy+teP2",
    "sEcVwtxG80c6NuWRJCrKfUEdK2Tqt/HE0sut6sqKMk/bpP8AGgD6ToryT4YaFrGs3C67qmoaoumoc2ttLdyN55/vtk/d9u59uvpOva7Z+HtNe8vXwBwiD70jegpNpK7KhCU5KMVdsNe1608O6a95evx0jjH3pG9B/nivDdc1678Qai95evyeEQH5Y19B/nml1/xBeeJNSa6vGwBxHEp+WNfQf1PeqNrBE95EbyN5bVWy8SPsaT2z2FedWre0dr2R9nluWvBU3Vceaf5eX+bNzwj4VPiR2vb9zbaJAf3spO0zkfwr7ep/DrXpcFxbXjw26LHa6ZBgRwghRgdOP6dqwIviNbxWkdqnh+JbeIAJH567VA6YG2nj4iQdtAhH/bZf/ia3p1qFNWTPIxmDzHF1OecPTVf5nfpeWxUbJUwPSni4iPRwa4AfEmMcDRYx/wBtx/8AE1DL8Tp0B8vTbVPQtKT/ACAq/rdLucqyXGv7H4r/ADPRWZnG1Aee5rj/ABz4xg8PafJYWUgbUZVxlT/qQf4j7+g/GuO1b4l6zeRGKGWK1U8E26kMf+BHJH4YriLidpZGd2LMxySTkk+tZzxKatA78NkcqclPEtadF+pDI1V3anyNUDtWMUejWqByzADqeK9I+D+jrNqd3r04xbWMZhhY9C5GWP4D/wBCrg9F0e713VoNNsVzcXBxntGndj7Y/wA817B4mv7L4e+Co7Cx2nyFCID1mmPIz693b2AHcV1Uoa3PCx+ItH2a3e55p8Tta/tXxa8StmOzXy/+Bk7n/Lgf8Brqfg7Cuk6Dr3iS5GIkXy0J4yEBZvzJA/CvKYornUb2OKJXnu7mTao6l3Y/1Jr1rx3LD4P8B6b4TtHBmdA9yw7gHJP/AAJ/0WtpOyuebRpupNRPMrmdri5lmk+/I5dvqTk16n4IiaLw5pSHIMrmT85D/QV5RHG80ixxgs7sFUDuScCvdNA09Yrm0s05S1hEf4gYz/M1lSWtzvzC",
    "fuqJ2kAxBGP9kU+iitzywooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApR1pKUdaAA9aSlPWkoAKKKKACiiigAooooAKKKKACiiigAooooAKRmCIWY4AGTS1marcsWW2gwZGIGPc9M/zNAEYb7ZdPcTELb25zljgbhzk+yjn6/SvC/G3ihvFGvvcIx+w2+Y7VT/AHe7n3Y8/TArvfit4hGj6HDoNlIRNeL+9IPKwg8593OfwzXnPhDRRrviGC2kH+ixfvrgnpsXt+JwPxNAHSab/wAUV4QbU2UDWtWXZaqfvQxdc/XofqVHY1zMcJUBCSWBy565Y9a0dd1g+IPElxeg5trcbLde20cA/icmoLdAASa87E1buyPscjwPJDnluztPhdoQvNbk1CVcxWS/JnoZG6fkMn8q9b2nNc74F09dK8K2gK4luR57/wDAun6Yrow6muuhDlgj57NMR7fFSa2Wi+QCloyKWtjzwqG4+YIn941NUP3rr2RaAPnn4sX/ANu+JN4oOUtIkgX/AL5yf1Y1y0ZrotX8F+LtX8QahqA0K8/0m4eQblA4LEjqfSqF34J8U6bZy3d5pU0FvEu55JGQBR+dAFWFgMknAHJPpXYeAPBcnjW/W9vkZNBtX6Hg3Ljt9PU/h1zjI8AeC7vxvqJ85mi0m3YfaZRwXPXYvv8AyHPpX0Xa2cGnWMdrYwJFDCm2KJeFAHQUAU9b1uw8MaSbm6KxxINkUSAAuccKo/zivDvEHiG88R6k13eNgDiOIH5Y19B/U967XxH4D8WeJtTN3eXmmKq5WKJZJNsa+g+Xr6nvWX/wqLxB/wA/emf9/H/+JrkrRqT0S0PoMsq4PCrnnK8/yOKzTg1aHiLw7c+GbqO2vLyzmuGGTHbszFB23ZAxn0rKziuOUHF2Z9NRxMaseeGxOHpfMxVbfSFj6Gp5TX25YMnNMaWoC7ehqNmb0P5VSgYzxJM8tV5JOaY7EdcD6nFR",
    "B/NlEUIaaVvuxxKWJ/Kto0+x51fFpL3mDvzVjTNMvNY1KOw063a4u5OiDog/vMegArqPD3wv1vWWWbUU/siy6s8ozMw9l7fU4/GvSdOg0Twdo86aSIra2iGbrUJjnP1P8Teijj0FdMKXc8LEZgtoDPD3h/Tvh5oU8ss8TXzR+Ze3rj5UX2/2Qeg6sf08U8Y+KZfFOtNPh0tIsrbRseQpPLN/tMeT+A6Crvjnx3P4nm+y2vmQ6XG25Uc/PO399/f0HQVP4C8Ef2/L/amrHyNGtyWdmO3z8dQD2Udz+A56dGx5Lbk7s3/hnolvoOlz+NNcXbFEhWyRhy2eC49z90fia4zXtaufEGsXGoXZ/eTNkKOijso9gK3PHPjE+IrpLSyHk6Va/LBGBtDY43Y7ccAdh+Nc1ZWU+o3kVrapulkOAOw9SfQDrXPOXM7I9fDUVRjzS3/I6DwNpn2nVTqEq/uLL5hnoZD938uv4CvZfC1mRC93IOX+7/n6fzrkdC0VEFrpNl8yLzI5H3z/ABMf89ABXpcUaWsCRKcKgwM9/etoR5VY87EVfazciSimNMirknj17VjXvi3TbUyIlzbPLGpJTzh+XFUYG5RWFYeLrC/A8v5uOfKcSY/LmtSPUbSU4WdAfRvlP60AWaKAQwyCCPUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABSjrSUo60AB60lKetJQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQA2WVYYnkb7qjJrnre5DTy3crbdpK7j0HGWP4DA/OrfiK7ENskOfv8AzN/uiuE8dam+j+DZ41bE9wotwe+ZMlz+W4UAeZ+JdcbxF4jvNSbOyV8RA/wxjhB+XP1JroNPk/4R34e3F2DtvNYfy4z3EQyP5bj+IrjbO2e9vILSL788ixjHbJxXTeOruM6pb6fb4FvYxBFUdiQP/ZQtTJ2RrRhzTSKmmxY08E8ea/6Cta1tBPJFChJeV1jGPc4r",
    "MIMNlaR+iZP41ueC1+0+LtKiJ4+0B8f7oLf0rynFzn8z7+FVYfDPuo3/AAue6rCsSJGgwqKFH0HFPCD3qvJKRMw3AY96BP8A7Qr1z88epa2e5pQpHeq4uPfP4VIs2aBEmG9a5/UZ/EgvZV06wthBnCyPICz++MjFb4kBqGa9SO4S2jHmXDjdsB+6v95vQfz7UAchqGo+LbG38y5MEQZgi7AjEk9ABzXKfGvV54tO0jw8bgvNIv2i6bpuxwucdidx/AV6Rqqm81/TbTqkebhx9On6j9a8A+Ier/2z8QdUuFbdFA/2aP6JwcfiCfxoA2NG+Juu6NpsNhptpo8FtCMKohfn1JO7knqTWi3xe8TpC0sg0hEUZJ8mT8vvV57C3BycAck+ld38OPAx8XXi6pqiMNGtX/dRnj7S49f9kd/y9aAPSfh9q3iPXtLOqa+lrBBOAbWGKIqzL/fJJPB7D8fSovHvj6HwzbmzsmWTVJF4XqIQf4m9/Qf5K+PfHcHhSz+yWex9TlT92n8MK/3mH8h/SvEZZZp7iS5upGluJGLsznJyepPvWFWry6I9TAYF1nzz+EfJNNPO9xcyNJPISzMxycnqT700tUO7rUUk/Zfzrj5W2fSutGnGyJJpcRuVPI4/GvRn+F2k2VnaPq3iG4tLieIO0ZReDgbgPYE1x/grSP7c8W6ZZMu6PzPPm/3F55+uAPxr1q70+HxF4suZLlS8ELLbRgEjpy3T3zXXRpq2p8/mOLm5pRdjjz4A8LA8+LLj/v2P8KevgHwYB+98S3zn/ZAH/slekDwZoYx/oIOPWV/8aePCGhj/AJh0X4lj/WtuSPY8x4iq95M89i8L/Dyxbc63d6R3ld8H8toro9Lvbe3Ty/C3h5YlPHmLEEH4t3/E1oare+EvCEfmXwsLaTGVjCB5W+i8mvOfE/xnvr0Nb+Hrf7DD0+0SgNKR7L0X9T9Kqxk5N7nU+Jteg8P2wm8TX3mzsMxabat8z/7x7D3/",
    "ACzXjvibxbqHiidTdlYbSI/uLSLiOIfTu3ueaTS9E1jxbqMn2KGa8nY7p55G4X3dz/8Arr0bQvCXhzwkBcarqtjPqSnIJPmCI/7KDv7nn6UCOW8NeA0MH9r+KG+x6agDCFzteX0z3APp1Pt1qXxP4tm1tV0zS4mt9NTCxwRrgyY6ZA7eiiuvvrjSNSlWV9L1XWmX7hnb7Pbj6Dj+pqa2vtTt1QadaaXpELHb/otuJX78FzgdveolGUtOh1UatKj71ry/BHAaZ8P/ABJqu1odNlijbnzLj92uPXnn9K7PSvDej+FoCmoa1aLdyD96IAZZD/sqB0H4c1Nre+Cwa91e7vb6NAGaN5yFJLBcbV2jqayJfEZsdItLrStOt7X7RLLGQo5+QLjJUA87vU9KIwSJrYqdVWex1tjrEen5/sfSL2RpBjz72QQKR7Drj6CrBbxHqOd10lmh7WsABH/A5efyWqfhK4mvNJs7u52/aJBNvYKAT84AH5CulQ81ZzmM3hZZmiGoTy3RkfaTcTvNj5WPT5V7elWhpNpZx38UcKBIoMr8irglGz0AFacj/vbT/rqf/QHqG5O5dVI/54f+yNQB4xcaR/a8irp+wXMZ3OVJBUeuQPX3rf0LRdXtZUik12+C4Y7SVkHA4GGz3pNDGdTnz/zwX/0I10louLxf91v5UAcbp/xLkjYbru3Dd98Lwn80JX9K6nT/AImNKQCyy/8AXOVJf0+U15inh1bmFWgmhl45Cyqx/LNQTeE7peREf++TQB7pa+P7CTAuCsTHs+6M/wDjwx+tbMHiCwuEDCXaPXqPzGa+bHs9Z0yNnimnRFGSFkJH5Vo+G73Ub+dfMnQIcgsEw2fquKAPpGKaOeMPC6uh/iU5FPrjvAusZA0iRCZFi+0ebuznJ5BrsaACiiigAooooAKKKKACiiigAooooAKUdaSlHWgAPWkpT1pKACiiigAooooAKKKKACiiigAooooAKKKbLIIonkPRFLfl",
    "QByWry/bdadM5jQ7PwHJ/rXlnxR1Q3Oo2dkDxEjTuP8AaY4H6D9a9GhZmNxMepXH4sf8M14x4suxe+K9RlH3Vl8pfogC/wBKALXgmASeI0mf7ttG0pPofuj+efwrMvbltQv7i5frNIX+gJ4/Stfw0wtdE1q7yA3liJfyP9SKwU4FZ1HoduDjdtnRX/ymEDsuK2PALY8caVn/AJ6n/wBAasi9G+OJx9P0q74WmFp4o0udiAFuo8k9gTg/zrzqT1R9jjYNwqJdv0PdJxi5fj0/lSLnNT3ibZA3rxUC5zXqnwRKoOanQCoFzmo9U1W20PSbnUb1tsFuhdsdT6Ae5OAPrQIqeI/EP9jC2s7KJbnV75vLtLcngnu7+iL1JrR0jS10u1IeVri7lO+4uXHzTP6+wHQDoBgVw/w1jufEF/qHjHVlH2i6Y29onUQxKeQv48Z9j616KXCoWbgKMmgDmNZ1RdIs/EGut/y6QmKLPdgOn/fRAr5niZn3O5JZ23EnqTXsvxh1M2Hgex04NibUrgzyD1UfMf1Zfyryvw14fvfFGsQaXp6/O/MkhHyxIOrH2H68CgDc8A+DZ/GuseU++PS7chrqUcbvRFPqf0HPpXr3jLxlY+BtKj03TI4vtvlhYLdR8sK9AzD+Q71S1rXNO+GXh230LQ41m1Fl+RMZO49ZHx1JPQd/oKj8E/D2YXQ1/wAUFrjUZW81IZefLP8Aef1b0HQfyiTb0R0U6cYr2lTbou//AADh73RbjTNGfXvEZaTU9SYi1hlPzDPLSuPYcAdsj6VyTT5Ymu0+LWrm+8XTQK2Y7KNYVH+0fmY/qB+FcFmuaUU2e3QqyjTTe71/yJGlLfSmr87hfU4pmach2q7+gwPqf8miwnNt6nq3wgsVtrPWPEUq/Ki+RDn0A3N/7IK7XTp7Hw9ZpNrF7bWjkGRvPlVSWPsTmuO8WTP4J+D+naVbyNDe3pVXZDhgT+8kOfyX8a8c+eebJ3SysepyzE/z",
    "rqirKx4Vap7Sbke9az8aPD9gGTTVn1KYdDGuyP8A76b+gNee618V/EusbkhuE06A/wAFqMNj3c8/liucs/Dt1cEGZ4rVf+mp+b/vkc/niul03w3pkGGkR7x/Wb5U/BR/UmrsZHL2WmX+tXLtbQzXMjH95KTkA+rOePzNdTY+CIbeMy38i3M4IAgUlYgT/ebq30GK6y2RvsbFFCwx4VQowoOegAqKQ43Z/wCeooEQLaSPbiCa5k8hSAtvD+7iH/ARxTdQij0jRLu7sI4opok3qxTP8QH16GrQ7fUVBr2G8MaiCf8Alj/7MKAOb8R6xfpf4W6aNDawvtUkZLRqzHj3Jr0ODB0i2d8nDgk9z8rV5Z4mm2amAOR9itx6Z/crXq1rxpFl2+df/QWpDMrxcwPhm5kVfl2IQCCP+WoxXFXUjSeF9NC4Aa7uM/lH6YruPGxA8M3WTjMac/8AbQVwly4/4RXTDuH/AB9XPXPTEdAj0Lwawi8M2RYltqTscck4k7V0dncC5top1R0Eihgsgww9iPWuZ8EsD4Z04qONs3bH/LWunVvWgZHdmWbbHt25kZYyr7S2YW5z25JH4UsMbx6beRytukSzVXOc5YRkHnvzRM37+2/66H/0B6kY/utT/wCuH/sjUAefaIuNWmH/AE7j/wBDNdHartu1Psf5Vz+iD/ieTA/8+y/+hmultgPta/Q/ypiPFpPDUoG4xnJ5+6aiFjfWv+omuIyP7rsKni1HX4UHlX9zsycBjuA/Opk8Ta9Bw7wSj/ppAh/lSGU21fWrRMNePIv92VQ4P/fQNbnhe+n1C5R50tl9PKgCE/UisqfxRPIP9M0rT5h3wjIf0atbw3cwzXEXkWAtAxONs7MOh7Ef1oA9D8IKY/F0GOjWZB/8eNei1534a+XxTYHPWAj/ANDr0SgAooooAKKKKACiiigAooooAKKKKAClHWkpR1oAD1pKU9aSgAooooAKKKKACiiigAooooAKKK5fxD42",
    "i0W4kgiiRzEP3ju2FU0AdRVLWJPL0qb/AGgF/OvOv+FxoG/5csD03f41WvfivDfQiKQ2ipnPylgTQB0EzrbWgckANIWP0Uf/AK68GLNcSvM3LSMXP4nNeqzeNvDd/YpDdzzRMqup8pgw+YcnNc2lt4DRMG71Hr2kX/CgDGt2MHhK7U8GWcfl8v8Agayk6V3o/wCFeS2qwzXuqBQc4V06/lSC0+GY/wCX3WP++0/wqJxb2OvDV401qYkbefpsZ7hR+lRKzKQVOCDkH0rq4JvhzbR7Y7/V9pOcEof6Uhn+HJ/5fNX/ADT/AAri+qzTdj6T+3MM4Rune1np/wAE9ZstRTVdDsr5DkTxK59jjkfgc08MK4PRfHfg7QtNNja32oPBvLqJVVipPUD271d/4Wd4U5xdXf8A37X/ABrvV7anydTl5nybdDsfMGeteUfGbxC0lxbaJA+Y4FFxcAHq5HyKfoOf+BCul/4WZ4UJ/wCPq8/79r/jXj2v6ouq63f37Pn7RO0ignOE6KPyApkH0T4Z09dJ8N6bZIMCG2QH/eIy36k1q3ALWzJ/z0wn58fypLQLJbRSIQVZFKkdwRRf3UOn2M15cttht0aVz6AAmgDwP4t3k/iL4ippWnxtM1qiWsUaDO5z8xx+YH4V19nHD8LPD8OmafEt94q1PBZUG7B7f8BHOB3OT06QaHp7eGDJrN1brdeMNZLzW9s54tEY5LP/AHRzyf8AgI711PhPQ4dPmm1C6la91a5OZ7xxyf8AZUfwr/h+FJ32RcOVayI/BvgE6XdHWdel+261Kd5ZjuWEn09W9+3QV2V3dR2VnNcznbFDG0jn0AGTUiA7RXF/FnVxpvgyW3VgJb5xAB329W/QY/GlpFFJyrTSfU8L1S+k1K/nu5j+8uJWlb6sc1SqRtrNyT+FJtj/ALz/AJCuZHuST6Eea3/BukHWvFelWBXdG8wllH+wvJ/QfrWHtj/vP+Q/xr1D4O2UcDax4iuRtgtY",
    "fKQnsANz/kAv51cVdnPXlyU2zK+MmrjUvGUdijbodOiCkdvMf5m/TaK5O3lYcKcZ9OKq3l++q6ndahcuBLcytM+exY5x/SrVtc20bKBJGG7tIeB9FHJ/Eiug8Y2rCLdgkgAdSa7fRPDc11GkswaC3P8AERhnH+yD/P8AnWP4d17wdpKJLdzT3V5wS8salEP+yucfjXRyfE/w6+f9ImPuUH/xVO4i/qsMVtp/lxKFjQqFA+tc5I4Jb/fzTtQ8d6BeW5jFxOMkHIVe341lnxH4eYn/AEu45Of4f8aANLfwPrVXX5MeF9Q/64kH8xUX/CQ+H+P9Nnx9F/xqO713w3dWUlrJcztFLw/IBx7H8KAOU8Qs0t8jCIn/AEO3GQAP+WS163CcaRY/9dF/9BauFuZ/Bl4d8897nYsfyuAdqqFHf0FdCnjXw2II4jPKEjbcvT3A7+9IYvjqRx4cuNuT+7jwB/10FcLcNI3hTSzsAxd3OcgHtF6muy1TxL4U1exe1ubudI2xnYOTg5Hf1rLaXwNJZQWj3d4YoHeROx3NjcSd3+yKAOm8FZ/4RjTzjHyzHpj/AJaV0gbFcfp/jDwtptpFa21xciCIMF3LuPJyckn1q0fH/hw9LyX/AL90AdFNJ+/tcf8APQ/+gNT7dVhsL+NGZlS32gs24nCHqe5rlZvHXh93gZdRkj2PvOYicjaRjj608fEDw5Hb3SJftmZCoBhYAHaR6UAZ+jN/xPZv+vVf/QzXTW5U3SD1B/lXA2viTTrTUnnW7iZWhEfRhghifT3rUt/HOlpcK73EQUAj+P0/3aYjjbfxjcQoUbS7eUA9VdlNSN4wtZDi40J/fZcf4rV63bw0kapJqIOOP9Wf/iae48KE/wDIRA/7Zn/4mkMw7jWdAuVIm06+hz3SRG/mBWn4dn0yW+to7GW9O04AniXHQ/xA/wBKjutO8KXA41hk+kf/ANjUmlp4f0q6imi1oyCNshWTH/stAHoGhEL4m0w/",
    "7A/UuK9Hrx228WaLb3trdJfx77cDauGwcEnnj3roF+K9g3S4tP8Ax7/CgD0KiuR0jx7bavOYbWW0nlAzsRyGI9getdRa3KXlussecHgg9j6UATUUUUAFFFFABRRRQAUUUUAFKOtJSjrQAHrSUp60lABRRRQAUUUUAFFFFABRRRQAZxz6V454vgXUfGWgaew3Le3u+RT0K7lBz+G6vYJ22wSH0U15SsYvvjbpEPVbO2Mp/wC+Xb/2ZaAPQ7/SfD2mafNd3Oj6esEClmxaITj6Yrm/+Eo8C/8AQLt//Bev+FaHxIvTb+D5ow2DcTJF+Gdx/Ra8RabLE56muWtWlCVonvZZl1DEUXUrX3srP/gHr/8AwlPgX/oF2/8A4L1/wpf+Eo8Cf9Ay2/8ABev+FeZaRoeoa1FPNaCFLeAgST3EyxIpPQZPerw8IXv/AEEdEP01FKlVK7V0jeeCyuEuWc2n6/8AAO/HifwJ/wBA62H/AHD1/wAKu6fqngfUphFDbaakjHAWazVM/QlcV5Nq+iahoZhN6kflzqWililEiOB1wR9aoJKd4HUE4NQ8RVi7NG8cmwVaHNTm9dndP9D6H/4R3Rf+gRp//gMn+FB8PaIoJOk6cAOSTbJ/hWP4K1aS50XS47lyZJLQtljyQrlQfxGKx/iX4tFnbto9nJiWRf8ASWU/dU9E+p7+31rrlUUYczPnqODnWxHsI99/JbsfL4q8DxSvH/ZVu4ViNyWKFW9x7Uz/AIS3wP8A9AeL/wAAI68nV3nlVEBZ3YKqjuTwBTrhJbS5lt51KSxOUdT1DA4Irj+sVdz6X+x8Dflu7+q/yPctCHhfxFbyTWOk2W2N9jCS0RTnGfStT/hG9F/6A+nf+Aqf4V5h8MdTaLU7izVj/pEJZB6yRnIH4gtXr0UizRJIn3XAYV10ZucLs+dzHCxwuIdOG2jXzK1/e22i6Y9zMuy3gUDCKOBkAAD8q5qX4l+H23xSpcsAcMDCCDj8ah+K",
    "d/5HhuK3B5nuAD/uqCx/XFeNb+axrVpQlaJ6WWZZRxFD2ta+re3yPedE8R6J4jubmS0g/eRIplllhVSQcgDPfoa30iiwCkaD6LXlvgS2K6Bk8G+1FIh7pGNx/XNeqgYGBXRTblBNnk42lTpYiVOnstBQMVXu9Os7/b9ttLe4CZ2+dGH2564yKsUVZymd/wAI5o3/AECNP/8AAZP8KoazaeGtC09ry+0uwWMEKAtqhZiewGK3pJFijaSRgqICzMTgADqa8P8AG/itvEGqnymIs4crAp7ju59z/Ksa1RU4+Z6WW4KWMq2k7RW7/T1Z2H/CWeBx/wAweL/wAjqxF498JwWj2sNk8dtJnfClqoRs9cgHBzXkKCSVZGjRnWJd7kDIVcgZPtkgfjTFm2uD6GuT6xVPoP7HwLdrv7/+AfQVroPh69tIbmHR9NaOZA6H7KnIIyO1SjwxoY6aNp3/AIDJ/hXMfD/WGn0CzjLZ+zztauPY/Mh/XFd0TgV3wlzRTPk8RSdGrKm+jaMa70Tw1YW7XF3pulQwr1eSCMD+Vc7Lr3gCJyosrJ8d0sMj/wBBrjfiH4ik1TxDNCsh+zWrGKJQeMjhm+pOR9BXJxtJNKkcSs8jsFVV5LEnAArkniZc1oI9/C5LRdJTxEndq9l0/M9c/wCEi8A/9A+1/wDBeP8A4mk/4SLwB/0DrT/wXj/4muAk8KX0D+XdX+jW0w+9DNfqHQ+hAzg0sPhG/upBHZ3+j3Mzfdihv1LsfQA0/aV+xH1XKf8An4/vX+R6Amv+AHYL9hs1z3bT+P8A0GtvTLHwnrETSadYaVOq8Nttkyv1BGRXgzO8cjI4KupIYHsRXUfDu8nTxlYLE5AkZ0kGfvJsJOfyB/ClDETckpI0xWT4aNGVSlJ3Svra35I9N1q28KeH4I59R0qwRJG2Lts1Yk4z0A9qxv8AhIvAH/QOtf8AwXj/AOJrJ+LGp+cdMhXgeQ0+P94gD+RrzdXJYDPU06te",
    "cZNRIwOU4etQjUqt3fZrvbse76La+FdftpLjT9JsXjjfYxazVecA9x7ipNSs/CWjoG1Cy0mDd0DW6bj9ABk1zHhO/bRPB8FxkhBDPfypj75LbIlz6HGfwrzbUdUuNQvpbm5laSV2JZiev/1varnWcIrTVnPhssp4itO0moRdvNnqjeIfACsR9gtDjuNPHP8A47R/wkfgH/oH2n/gvH/xNeX6Xpd7rMsq2aIVhXfNLLII44l9WY9KvL4Znbpq2h/+BwP9KzVWvJXSOqpgcrpS5Z1Gn6r/ACPQv+Ej8A/9A+0/8F4/+Jp8WueAJjg2dgn+/YY/9lrzm68L6hbabJfpNY3VrEwWR7W4Emwk4GR161i+cVPXBFKVerF2aRtSynL60eanOTXqv8j6AtdC8M39us9ppelTRN0dLeMg/pS3Hh/w3aW7zXOlaVFEgyzvbxgD8cV518LtYlt9UvICWa3Nq0zqOm5WUA/XDGj4pa9LPrZ05JP9HtQPlHQuRkk/QED8629v+757anlrKr4x4dS0Wt/L/PodHNrPw+hkKixsJPdNPBH57aiOu/D49dMs/wDwXD/4mvKbOG41G+htbVTJPO4jRc4yT71tHw5bIzJL4l0gSISrhPMcAjqMhcGs41a0tkjurYDLaFlUnK/y/wAju/7b+Hn/AEC7P/wXD/4mnJq/w7kbB02xX3bTxj/0GuBi8OW1zMkFt4l0iSaRgqIfNXcT0GdtYl/DPpmoz2V1hZ7dzG4DZGR6GiVWtHdIVHA5ZWdqc5P5r/I9107RvB+rQmTT9N0i4QcNst0yv1GMirZ8H+Hv+gFpn/gKn+FeN+A9WmtPF2niKQjzp1hcZ+8rcEH+f4V7n9sRLOSeVgFhDlz6bc5/lW9Ko5xu0eTmGEjhavLCV01c8XtFgvvjlHbWEEVvZ2d0wWKFAiDyozuOB6tmvW9DTbp5b+/Ixrxr4Wlr/wAaajqcv3ltJ5yT2aRh/ia9s0xPL06Ef7OfzrU4",
    "S1RRRQAUUUUAFFFFABRRRQAUo60lKOtAAetJSnrSUAFFFFABRRRQAUUUUAFFFFAFbUX8uwlPtivNvBai9+LfiO8IytpAIFPoTtX/ANlNeh6wT9lRB/G4FeefCk+bJ4q1M8me92hsdgWb/wBmFAEvxXv9tjptvuyWaWdh9MKv8zXlQfmuz+K18ZPEa22eLa2ijI9zlz/MVwgfBriqK82z6bBz9nh4R+f3s9S0Pw5c6j4F0iBbfzILi8lu7gbgOANqZyQeetdlD4F8NxRKr6bESBy3mn8/vV8//b26fJx7Ufb29U/KtFVaVrHHUwNOpNzdTfy/4J3XxD1HTjc2um6QyGzsIzGuw5XezZbB79ua46FmeQLGCzn5VA7seAPzNVkM97KFhjkmc8BY0LH8AK9D8E+Cjpt6mteJ2W0iskN0lo/+sO3o7j+EA9AeSf1z9m5yuzteLp4WkoU3tt3b/wCHOqv7yHwbpou22PcQ20en2MR6FkHzufYMT+XvXj99qEt9dyTzSNI7sWZmPLE9TWj4v8Tz+ItWkuZMohG2KLP+rj7D6nqawYXjWZDKoZAQWUkjcPTjmio+d+QYSn9Vhb7T3/y/z8ztPC1suiWKeI7yASzyMY9Ktm6Syd5T/sr/AD/CqvxAiEPidrpF2x6hBHdqB6svzf8AjwNWZviFFcT280ui6aXtoxFDgygRqOgADYFY3iXxKNfjslFpBbi0Ro0ERc5UnODuJ6HP51T5OTlRhTWIWIdadrbboXwzq39la5Z3hOFgmV2/3c4b9Ca+hdOZVSWAH/UuQOf4Tyv6Gvl+GQCQZ6Hg/jxXvPhLVmvdAsLvcC89n5D/APXWM7f1zmqoaXRjm3vqFT5fqv1OU+Leo+ZfWNuG/wBXbmUj3duP0WvOUfLrn1rofiNf/afGGoBT8kUiwL9EUA/rmuatEe4nWGP/AFkhEa/VjtH86xmuaTZ6WFqKlQhHsr/qe1eD7bybXw3alcMttJeP9ZDkfpmu",
    "+rmdBgT+371o8mGzjjs4z/uKM/qTXSbxXclZWPlqk3OTk+o6ik3Cud8a+KU8M6OXjZTez5S3Q+vdj7D+eKJNRV2FKnKrNQgtWct8T/F6xq2i2cnTBumU/lH/AFP4D1ryjzHnmVEVnkkYKqqMkk8ACmXt491Ozu7OxYsWY5LMepNXfDetroGrJqP2eC4miB8pZgSqt/e47j+tcL9+XNI+qp2w1L2NHp+L/rY9A0HTYNN8zwsyRSXV/A66jcdfKlZCYolP+z1Pua8yYtG7JICHUlWB7Eda7CH4jx285nh0PSUlMnmlxHJkvnO7Oeua5DVr8ajqt1eCNIzcSNKUjztBJycZ960qcsklE48Iq9KUpVGtfO53Pw21Mi8vLHPNxb+bGPWSI5A/FSfyr1q61NI9NS8HKeU05+ioW/wr558Lat/Y/iCxvS2FgnVn/wBw/K/6E16/4puBp/grVY9xDQBrVB/syOuP/HTWlJ2g/I48fT9piYtfat/keMXlw0twzucseWPueT/Oug+H8QfxIt46hksLea7OemVXC/8AjzD8q5OaTdM5HTca7fwLbMPDmuXIBLTtb2SEDoGfLfpisaUPeR6eYYi1Kduuh3PhnwNpF1o0dxqVr59xL87sZGHJ5PQ+9GtaTovhSS01DTdGkmvIpN6BHcjGCDknIHX0rrbCMQWEUYGABUWr6vDoWkT6hchjHEB8qnBYk4AH512N21Z8zGLlJRjuzyGT+x2kLN4MvSzHJJv5uT/3zVzSdQsdK1KK4sfCN1BOT5SyPeSMF3/KeCuO9dC/xi0sEg6fd5H/AE0T/Gt3RfE9t4q037VDBLBDFdBG81gc7RvPSs4unJ6HZWp4unH95e2255b8Sr9bjxXcxxn93b7bdfoo5/UmuTWQ/MRyQpI+vapNYvmv9SnuWOTNI8p/4ExNWvC1n/afiTTrQ8rNdRq3+6Dub9BXLy8zv3Pe9r7Gmor7K/Jf5no3iv8A4lXhB7MEAI1rY4Hp",
    "HH5jfq1eU+bn8a7vx/qaz6TYNGTi6lurvnuDJsT/AMdFefK3zDPTNXVV5swwE+TDx823+n6Ho3hWwjm8JLFKu4ajqXzD+9HCmf8A0KvR4vBegiFQ2mxElRn5m5P51zPhTT2Sz8MwFcRx2LXB92lfP8hXoYPNdUFaKR4GIqOpVlPuzgfGenaboXhzUhp9okDXHkQsVJOfnLdz6CvG5Jcux9STXqfxZ1DyrK2tt3zSzyTlf9lQEX+teSA7mC+pxXLWV5nvZdJ08Kl3bf5L9D0DwDIljYalfztsUy2lmGP+1IGf9MVieOGmj8X6olwpR/tTsAe6k/KfoRius8PaBPdeBNMJSJbS4vJLy6kmkCKqj5EBJ9RVq9sTdqkF7/wjmq28XyQtc6gI5o0/uiReSPrWrppwSOGONlTxM6iV09P6+48y07U5tMv4bu3bbNCwdGwDtI+tddbfEiNWDXehaTM2cllh8sn64yKm1bw34c0+zjub3T723gmcxrNp2oJdIGAzj5gKxW0LwxOCbfxDe22fu/atPLD8Sh/pUqlOPws2njsPV/ixv6r/ACOsT4maFcRBLnw7CnIIMLJlSDkEcCkl8U+G9UummmtreKWQ5ea40iObn1JV8n8q4nV/Cw0/RF1ay1W01GxM/wBnLRK6Mr7d2CrD0/pWAJCvKnBHcUnKotGVChhKi5opr0b/AFuewLp2sWz299o2meGr2Nh5kFzaWAU8dx82QRV3XL2+0r4V6nLqEbQ3csTREHjLStjj/vo1S+Ft5I2lNGzNsGohUHYEwksB+WaX42X/AJXh7T7AfeubrzD/ALqL/iy10Qd4pnj4iHs6soXvZ2MH4S2wXT9duscsYbZT+JY/0r2WJdkKL6KB+leZ/C612eCY22/Nd6g7Z9QoC/0Nend6oxCiiigAooooAKKKKACiiigApR1pKUdaAA9aSlPWkoAKKKKACiiigAooooAKKKKAMPxRdC1sHkJ4iikl/JSf6VyXwotW",
    "i+H8bEfPeXcjE+vIQfyNaPxOu/s/hzUiDgi12D6sQv8AWq3hqG/svh5oa6ZHG1yIRMBI21csxbn86APL/Guof2j4t1ScHKtdOF+i/KP0FUNBsjqfiDT7IDPn3McZ+hYZ/TNekP4d1Z3LP4U8MMxOSTv5/wDHqnstF1yxukubPwz4at7iPmOVN25DjGR83vWXstbnoPGrk5UuliOPWPEXiDxLqlvo17aQ21vM4jV7ePAQNtHOwk5xmtAaR4376np3/fqP/wCN1N4H8M3mgi7e+EYlmK/cfd0zn9TXYcgYrU884WbTPGpXyzrSruONsBVCT/wFBXKeLtVi0q1bQbOczyCQSajcMxZrif8Au7jyVT+f0r1PU11BY9+mxQST8gCZiF5GO2K5A+E9XY5HhPwy3uVf/wCKqZxclY3w9WNKfO1drY8ntre41G+itbaNpridwiIvVmNdbN4f8J2ExtbvU9Xnuofkme0hjMRcfeCk8kA5H4V10fhvxFaiQ2OhaDp00iGP7RbgrKinrtJJwccZrf8AC3hZ9E0/y3hjaZ8FyvzAe1TGmlua1sZOb93Q8uOmeDAP+PzxD/34i/xpt14e0G70HUbzQrvU5LmwRJXiu40UMhbBI2+nWvbGsdnWFB9UH+FZWqabfBHaxtbSR5UMTx3Kfu3Q9QcY9BVOEexlHE1E7tnz3ur1f4U36XGm3ttNMIxYzpfjP9zaQ4/NR+dTN4W1Xr/wiPhjHrsfH/oVWLLQtfslma08NeHYfOiaGUxhwHQ9VPzdDURp8rub1sWqsHBo8k1K+bUL6a5f700jyn6sc0lhetY3cdxESJImV0YY+Vgcg8+9eqnwhq56eEPDP5P/APFUn/CH6sOvhHwwPwfn/wAeqfY+ZssxS+ycj/ws3xH/ANBSf/viP/4mlHxN8R/9BSf/AL4j/wDia65fB+rk/wDIoeGcf7r/APxVL/wiGrA8+EvDA/Bv/iqPZPuCx9P/AJ9r7l/kQeCfHOqX",
    "Nze32tahK+m2Vq0koKLyxYBAMAcnnFcL4o8S3PiPVZby5O3d8qRg8RoOij+vqa9IXwxry2jWy+F/Dgt3cSNGC4DMAQCRu5xk/nUX/CH6t38JeGB+Df8AxVN0m1a4o46EZuahZvTTT+rnmfh7Q5vEOrJaRMI4lBkuJ2+7DEPvOf8APWt4aT4NLELc+ImAJAYRQ4YeorsW8J689nJZQaVo2m21wym4+xsVMoHQMSTx7V2uhaDFo+lxWwiRnHLttzk/4VUaaW5z1cZUm7xdkeM/2R4Oz/r/ABH/AN+of8azfFOhWWl2mmX2kzXUtlfJJg3IUOjo2GB28elfRJtUPHlJ/wB8CuR8QeHtSuZPKtNM0i+tRIZVS9QkoxA3YwR1xTdNNaE08VUjK8ndHgqMN2D0Iwa9hvFbxX8MIZbMmTUJI45GiHJleD5XA98c4o/4Q/V88+FfC4/4A3/xVWB4a8TG0jtLfTdJ06JJxOr2bFXjfGNykkgHHHTkUo07XNauM5+Vpap3PFicMQeoPIrX0bxTqWgh/wCzrua3MmA+zBDY6ZB4zXqN74a8Q3krNd6R4d1Jz1mntwJD9SuM1V/4QzV/+hS8Mfk//wAVUex8zZZirWcbnHf8LN8R/wDQUuP+/cf+FZ2q+L9T1kL/AGheT3GzlVcgKD67RxmvQf8AhDNXPTwj4Y/J/wD4qnr4Q1qLlPCvhYfWIt/NqHRb3Y45lGDvGCT9F/keRRLJcTrFCjyyucKiKWZj7Ada9VsmPhrwZd2EjD7baWU1xdKrZEU07BI1J/vBc5H0rQh8O+LY0ZLO30fS1dcMbCFY3P8AwLGR+BpYfCer2Wmy2cGmaZeJOwef7WxbeQcjOCO5z9auNPlOetjHVautEzxeV8yNjp0FdV8Oj5Ou3F/xjTrGe5yf723av/oVdmfB+qA/N4S8Mj8H/wDiqli8Ma2lvPbQaHollFcqEne1JV2UHOOSe4qY0rM0rY72kWktzmPE2nXN94P0",
    "q5t43lfTITbXkajLIpO5JMf3SM81wIfJyDX0EfC158k1rLJaXUa7VlifnHoRjBHsax7vwtrkrsbjRvDmoMesk1qEdvqVxTlTu7kUcbKnFQauked6b481vSrNLWz1C4jgjGFTCsF9hkHAq0fib4j/AOgpP/37j/wrsB4N1Tv4O8N/gz//ABVB8HamP+ZO8Of99N/8VU+yfc2/tCD1dNfcv8jy7VdavNYuWnvp5JpWwC0jZOB0HsPYVN4f0G98Q33kWahY15nuX4jgTuzN04HbvXpqeE9YjP7rwj4aQ/7UZf8AmxqWXwn4i1OOO11OaCG0Ug/ZbXEcX/fKgZ/GmqK6kVMwk17qsVL++Gs+ENW0nSlZrWyig+wxbfmmiiYB3x1JPXHpivKWnYlsHgmvfH8JyRW0Isg8M0PKSI2xlPqDj9OlY9/4Z1m4lL3eiaBqEhOTLNCI5G9yVIzTnTUncjDY2VCLh0PO9D8cX+h2slrD5T20jb3gmhEiFsYz6joKvN8Qt5JfRNCYnrmw/wDr10UvhG+3fN4I0hveO5dR/wCh0w+ErkD/AJEKw/8AA6T/AOLpKlJaJmssbSm+adNNnG634rl1nT47Ew2lraRyGUQWluIkL4xuPqcViWlrcahdx2tnBJPPIcJHGuWavT4/Cl+rDZ4I0hP+ulw7j9ZK1bfw34jdGgiXStFtpBiQWEQR2HoSMk/99ChUn1YpY+KVoQsP8EaadMuLbS1dZG08PPeuhypuZBtCA99iAg+9cn8Zr4XHiu0tQ3y2loCfZnYn+SivVdD0S30OzS1tQSM5Zj1Y+teD+Pr06l421qVTuxOYEx6IAg/UGtUrHnSk5O7PXPAtqbbwx4ctmGD9n+0MPdyWrtKxdKtfs15FbjGLS1jhHthQK2qYgooooAKKKKACiiigAooooAKUdaSlHWgAPWkpT1pKACiiigAooooAKKKKACiilHJFAHlPxgu8aLdxgjLTxRgfT5v/AGWuv0yMW2m2",
    "VqB/qYI48fRQK85+Ikp1DVNOtBz9q1M8eoBC/wDs1ejK/wA5IPfigCm/jXw9DK8UuqQrIjFWUq/BBwR0pV8deGu+rwf98P8A/E046FpEjs8ml2TOxJJMCkknv0qVPDuiH/mEWH/gOv8AhU+8bfuez/D/ACI18d+GR/zF4P8Avh/8Kd/wnnhn/oMQf98P/wDE1YXw3oZ/5g9h/wCA6/4VIvhnQ/8AoD6f/wCA6/4Ue8O9Ds/vX+RYsNUsNTsje2c6zWqlsyAED5eT1FeJ+GxZ+KtV1W71/wASz6VEZPMiAughcszHADdgMdPWvYtW0x18L3+n6HBBBLNC8cKLiNAzDBPtwTXJeDvhdY2OkOniWwtbu9aYkESMwVMAAAgjvk1SMXa+mxzfg3U5NM+IF8llq95qGhWNvPNK8kpKyRqmQSOn3sAEVY8KabqfxQuNR1LWdZv7W2hdUiitX2qCQTtA6AAY9zmvRJvCWmp4c1LStJtYNPF9A0TSRJzkjgk9SPb61xHh/wAM/ETwlazWOlNpLW8khkJd1b5sAZGQD0A4NAjP0GO98PfGKLQLLVL27s1k2SrNISGUxbmBHTj19qT4fa5/wj3ibxKNVupntrK3lYrJIWyY5QABnuc4/Gur8C+BdR0nX7zxB4juYp9TuNwVYjuClj8zE4Az2AHAFc54h+FOtax4xv7uB7aHTry53s/nfMEJBJ29znJx60AZfgzxFqE3xHs9T1i4lit9QE8xDOfLCbX6DsAUI/4DWr4XvLnxb4u1bxZfSzpo+lhpY4d5CHaMouOnAG4+5HrV7x78NdT1rVLA+H0tY7K0sUtUWSbYRgt7ehH15ro7zwhPYfDKXw5oQia5liCPJI2wOzEGRifcZA9sUAePaNqdjfzXlx4k1nW4HkcPGLIlskkls5PHbFdnENO0j4ca5rmj6nrFwblRYxm+bBViwyVA9j19qdoPhD4i+GbN7XSZtLihkfzGDFHJbAHUrnoBWv4o8L+L",
    "/Engiw066ks5dQFy01ywdY0AGQgGBg9c0AcVolloM/hgX2seMb621Aq7fZIbgkjBIUYxnJwPzrd+HeqeI4vAXiLUI3urswoBZLJmTEgB3lc5JAyDjpxWx4l+FcV74Q0210e2tIdVtAgll4Tzvlw+5gOeeRmtY+HvE7/Dmw0qyvotN1e0CqzxP8kiLkAbgOMjB6dRQB5Rp1/pOpWs0nijxFr9vqbSHaU3OgHY465znjivVPhVbsmkXkieIjrVq0irHuV1MDAfMMPyM5XisO+0P4k6roraVfRaPJE8flNcOQZCOhO7198ZrsvAPhE+DvDv2KWZZ7mWQzTOgIXcQAAM9gAKAJ/G1/faT4O1C80qN3vUjAj2JuKZIBbHfAJP4V4dptzo2p2Ur+JfE2uWuqs7YOHkTHbPc+/SvbPHOla/qukRJ4a1H7FdRyb2+bZ5i4Ixuwcc81xeoaJ8SdX0htMvodGeORPLediDIR0zn19wKAIdI8PT6z8N9WSw8Szas0biW0KeZG0MiAlk+Y7vmBxjp0NYel+O9Y1XwdaeE9JjuJdYnkaE3G7kQ9fvdQcEgnsB716f4I8Iz+DPCs9oskdxqEzNM5XIQvtwqgntwOfrWL8LvAmp+Fr7UrzWUg8+4VEjaKTdxkl+3HO2gDnvG1pqvgLwfp+nabeXskl4zPf3ys5JIAwgb+BeT6Zx9a5tx4amgiez8aaxbXgIMjXUMjKOOcbORz7mvT/F+i+NpPEMeoeGdTj+yBFBs5ZNq7h1yMYIP1Fc1q/gTxr4zu7ZNcXSLKCEkmS3Hzc4ycDJJ9iQKANbxT4Se78JW+qXnia9cabp27zIDhbggFt5OcnPyj8K5L4f+B7rxnpNzfXeuajapHN5KCOQtuwoJ6n3FemeMfDl7eeAToOgpHu2xQDzX2gRqRnn/gIH41N4E8PXHhbwdb2FwqG8UvJKFbKlyxIGfpgUAeP+JdTgu/iReW95qGo2+l2x+zZs2LOBGgXgdOWH",
    "JrpPA1h4fvvEST2Or+I5msEN0y3mFiIXjn15PT2qrpHgP4gaFq11qWnPp8V1dbhI7SK+QzbiOV9a6yw0/wAe3Gha3b63PZSzz2xhs1j2KAzZDEkKOgNAHjd9qOq6gl/rqX10sRvduBMwwz73XHOAAFr0P4ieNZ9WsbbRtFkYO1qt9fzRsR5aBA+3I6dif+AjvUlv8LdVi+Gl3pLJb/2pLfJcKPN+TaoCj5sem786NP8Ahdq2m+BNStYltm1rUnSOQmTCxwqwO0Njvjn6j0oAu/BK2uJdF1LULmaaXzp1hTzJC2AgycZ92/Stn4peJ7jwv4XVtPk8q8u5RDHIBkouCWYe+Bj8a1vAugS+GfCNnp1yE+0pueYocgszE9fxFYvxR8E6h4xsLEaZLCJbR3JjlYqHDADg4PIx+tAGDpfwyvNa8O22p3viPV/7RuYBOqecdqlhlQc8+marzXfir4eeAdQ/tm+33dzNHDYEzec0WQS7ZPsOB2NWZdH+KF5o40iSXTYLZoxC0iOFcIBjqBnp6CjXfhLqM3g7T9OsL9Lm8t5pJ5zO7BZWdQPlPOMBQBn3oAh0f4azat4Zg1fWPEmrJcTwfaCFmJVFI3DOck8cnpWB4Z8TapD8NfFMlxf3EqKIoLcySFmR5CQ21jyOOa6V9H+KFxo39kPNpcdq0XkF1YB9mMYyB6ccCm6p8LtUt/ANtoWkSW89w939qvJXbYpIXChfYcfl70AcroNloV14b+2az4zvbO+Ic/Zo7gkgAkKMHkk4zjPerPhzXdUtfhX4kurq8uZI5HjtbRpZCxV24faTyOCOnpXXeJvhZDeeDbC00a0s4dVthH5kuAnnfLh8tjnnn8KZ4h8G6/q3w40bQ4ILOK7tHBuEWUKhChgGBA5Jzk+9AHG6Jpmi3PhT+0dW8ZXdrfbZGFpHdAsMEhRt65OOnvXW/BS71O+0zVHvri4ntlljWEzOXw+DvAJ9ttbul/DLw3b6VaR3ukW012kK",
    "CaQljvfA3Hr65rqbKyttOtUtrG3it4I/uRxKFUfgKAMe58beHbS5lt7jVYo5onKOpV/lYHBHSov+E/8ADH/QYh/74f8AwrRl8NaLNK8sukWLyOxZmaBSST1J4pn/AAi2hf8AQG0//wAB1/wqPeN06HZ/h/kUP+E+8Mf9BiH/AL4f/wCJo/4T7wx/0GIf++H/APiavHwtoX/QG0//AMB1/wAKafC2hf8AQG0//wAB1/wo94d6HZ/ev8iSz1/Tr/T5tQsblJ7e3DF3UEAFV3Ecj0r530KJ9Z8UadHJy15fIz55+8+4/wBa9s8cG18P/DzVhYwRWySReUqRIFG6QhScDvgn8q8q+GdusvjvT3cZS2SW4P8AwFDj9SKpeZhK1/d2Pd9MIlu72YdDJtH51o1l+HlP9liQ9ZGLVqUxBRRRQAUUUUAFFFFABRRRQAUo60lKOtAAetJSnrSUAFFFFABRRRQAUUUUAFH0oooA8T8SaBeavcW0tlex2moWUj4jnJQliQchugPHfH1rmNVufHOhjdqF1qsMfaUSloz9HXI/WvobUNHstTAN1ArOOkg4YfjWTL4fubCN2sbndHglo34JH8jQB89/8Jn4iHTW7/8A7/GnxeMPE80qxw6zqTyOdqqkrEsfQDvXrGoeHPDuu2sk1xplus4HMsA8ps++3g/iK8o8MPbWnxB0thII7WLUEO+RgAqh+pPTpQBbbXPHMTxo93rqNIdqKwkBY4zgcc8VDbeKvGF7M0VrqmsTyKMskTuzAe4Fem3WqR/8LE0e6lvrdbFb2c+YdaFwhBjfafLPEQ7fjisTwRY/8Iz4r1Z7zUdNP2zT5zCYtRRQW3rhfMH3Cex69+1AHIN4k8aJeLatqOti5YZEJaQOR67etSLrnjp3dEvNfZ4yA6jzSVJ9eOK7aC/Y3XiC1t9UttP1q8tIPsVxLqv2nCKzb0E56E+nuKzvB2r6rpXxKt7bW/EcdxFJCzTyLe74iRG20M2cFhxQ",
    "Bza6347eR0W78QF0xuUCXK56ZGOKrHxb4uEUsn9r6x5cLBZW8x8Rnphj2P1rs/BfiuMeGdRm1/VrwzTajBGZY7zZOE4G7OclR3A7Zq1qmtWfm+LJtQksjE2q2MhiglVxNCjLyMfeJUAn3zQBw03ifxnb2y3E+p61FA+NsjtIqt9CeKjk8W+LoY4pJdX1eNJhmJmkcCQf7J7/AIV1njSHU9c1qTHiqwm0LULuJYYhfKQiEjB8v+HbySeOlaXi/VPD+veHZ7XTtSjkk0O7hNvG6iPEQxGyoSfnHy7sj2oA4Cfxj4stZmhuNa1WGVfvJJK6sPqDUf8AwnHibH/If1L/AMCGroPinpJuPEmpa9b6hps9nK0YRYrtHkPyKv3Rz1FcVe2D2VvayPNBILqLzVEUgYoMkYYdjx0oA3h4n8ZmWOIanrZklTzETfJl0/vAdx70n/CV+MN0KnVtZ3XAzCPMfMgP931/CvSrDXtBTXvDqSrC15HoKqL77aAkP7tgYynTd16nPNYvgvXNMk8I2epancwi+8MG4NvDI4DTK6fIAOpw3HtigDkB4q8YFZ2/tXWitucTHe+Ij6N6fjT7XxN4zvo2ez1PW7hFOGaJ5HA+pFdx4U1TRLDw/ZaZq+sRifXhNNfKFDhjL8q73z8hGM89z2qlA97/AMIhpel+HPEljpl3plxPHfKb0QiU7/lk3fxrj69fagDi/wDhOPEwyG1/UwR63DVcu/EfjWwgimvdR123imGY3leRFf6E9ay/D1vp974utoNevRFZPOfPuFbg9T97sCcc9s5r0Lx1bWevaPeapcahH/aqpvTTrfVFmiADAeaoJ/uAggDOeaAOPtPE/jK/VmtNT1u4CnDGF5HAPvilt/E3jK6lkjt9T1yWSI4kSN5GKfUDpW98Ota0/S/Burf2hfXFssl7bgm1uPKmCkgFh3IHfHaul0/WEbxh40leeySSdLdbcRakIFmUAgFZh0OOTjoeKAPPo/EXjWW4",
    "kgj1DXnmix5kamUsmemR1FOh17xxOm+G98QSISRuTzSMjg9BXa+HdTnTUdatp7ixS0nvI2mf+3itzFhF+cTdZFA7eoIrK0rxPJaWnjaG18RXMsESFtPkluMMzGRiWTpyc5JA560Acz/wlXi/bO39q61ttziY+ZJiI+jen40g8V+LiICNX1ki4OIT5kmJT0wv97n0rvPC2qaLZaFY6Xq+tIbjXEmmvwVEgkaX5V3yZ+QqBnnue1VNLht7uw8IRjVtMhbw5ezfblmulQgCUMGX+8CF4I9aAOUi8QeN5y/k3+vyeWxR9hlO1h1Bx0PtUEPizxbcz+Rb6xrEs3P7uOV2bjrwOa9F0bxTo4s76aXUp4Yb3xQ7I1tciJwjKNruOpjOBn/61Z1hqdnpWueKvEepaja2N3Nci0tTahbggZDM6qCMgqF+bjkn6UAcbF4p8X3DyJBqutStECZFSSRimOuQOn40sXifxjOIvI1TW5POBMex5G3467cdcd8V3sWsaf4e8SeJ9d0m8s5Yr6wivrdDIvzsWy6EdQxOeOo3Vl+NNastPsfC0nhXUvJhMlxIBBLteBJWRzGwByACSMH09qAOWuvE3jKxVWvNT1u3VjhTM8iAn2zS23iTxneoZLTUtduEB2lomkcA+mRVn4p63NqXja+hTUGutPiZGgRZt8S5jXJXBx1z0rovh3qEcPw+uLdLiIXB1Mv5R1X7C+zy1G7cOSM9u/4UAcnB4m8ZXXmfZtT1ybyjiTy3kbYfQ46VF/wmHiv7Mbga1q/kBthl859gb0z0z7V3GlX1xLpGiQ6L4jsNNawu5m1WOW8CNKxlzuJ/5aAqDz3yPwZrcdprun+JdDsruwsro62LpIrmdYlaPywCynocnJ49aAOOk8T+Mo/MMmqa2nlBTJueQbA3Qn0z2z1p9z4j8bWcXm3Oo69BHkDfIZFGfTJrr/Emu6deReN/st/BKGg0+KNlkH71kb5tn94D1FP+IXjuzju9V0i0",
    "e4vDfRwI0jXSvbRABSTGo6N689eaAOMbxJ40S5S2fUtdW4kGUiLSB2HqB1PQ/lTo/EHjecyLDfa/I0bbHCGU7G9DgcH2r03UPF2kHXpb83tq2pWV2NPtmDqQ0MzRkyZ9FHmDPvVew1a28rxOLe8t3kk1tpI1XVhZl02L8wkGSR/P8KAPNZPFPjGKFppNV1pIlfy2dnkChhwVJ9c9qmXW/Hblgt54hYqcMAJeDjODx6EV2Wn65psXhOPRdZnt3g1LU7qC5b7QJWhYtlJd2ckBgDuPXrVgeJ2h+LOqxDW1GnNYlxtuQIWlEKDI5xuyPrxQB53ceMPFlpMYbnWdWhlXGUkldWH4GpR4n8ZHT/t41bVmtNxQzLMzKrDsSOh+tZVpqWo6hrqXMsR1e+kwu26Qzl+MDI6nAH4V694S1EaBG82v/wDCJ6Msi7ZIrYATuPRgrEfz+lAHHaPH8RPEFt9o0nV7q6i7smpLlfYgtkfiKtXOlfEyyQvdapcQqOpfVY1/m9dPffED4e2V2Lq3tRPeIciaxtTE2f8Af+XNcX4u8eeHPEhkKeE0E7A4ujP5cmeeSEGD175oAyrrxX4pspNk3iK6ZvSK+En/AKCTUH/CbeJT/wAx7Uv/AAIaudU4NeweBfD3hx/Ctpql3psd1cuDveZi4DAkcL07elAHCLe+J/FaC0FxqepxhgxjLs6AjoT2H412XgrQ7jwxdXl1qbQrcT2pt4oY33sm4gksRwOBjGSa7qwtpdZdraKSO1toQD5caYAHsBxW5Y+HNPsZBKIjLMORJKc4PsOgoAtaXEYdLt0YYbYCQe2eatUUUAFFFFABRRRQAUUUUAFFFFABSjrSUo60AB60lKetJQAUUUUAFFFFABRRRQAUUUUAFU9Yn+zaPdyekZA+p4/rVyqeq6f/AGnp8lt5pi3EEMBnkeo9KAPPbmY2fhi/uI1LMiMRjnJCk/zxXisFrbTHH2iXd6CHP9a9xvrDVdInzc5aLPys",
    "h+X8MdPoRWRqHhPR/EH714DaXfXz7bCNn1Zeh+v60mrlwkou7in63/Ro8rSxsHOBftn08k/41YXQY3B2TXDYODi2PFdDrPgzUtLVpL60OqWY5+2WfyzIPVl7/iD/AL1N0LWNa01hPoGpNqMSfegLFJ1HptPX9RUckv5vyOhYil1pL75f5mF/wjgPR7n/AMBmp3/CNf7d1/4CtXsXhj4l22qkW11K8F2DgxSfK+fp3/Dn2rtIb7z03Rzb19Q1HJL+b8ivrND/AJ8r75f5nzT/AMIyf791/wCArUf8Iyf711/4CtX04J3P8bfnS+a/99vzo5JfzfkH1mh/z5X3y/zPmL/hGG/vXX/gK1H/AAjDet3/AOArV9OGV/7zfnVS/wBVi02AS3UzKGYIiqCzyMeioo5ZvYUuSX835B9Zof8APlffL/M+bv8AhGH/AOnv/wABWpf+EXf/AKfP/ARq9p1Hx5NbSmJ5rOwI6pMXuZh/vJGQqn2L5qpb/ECR2x/bGnuc/duLKaBT/wADV3x+Iqdnbn/I3S5o8yw2n/b3+Z5D/wAIu/8A09/+AjUf8Iu//T5/4CNX0BYeJo7qWGC53Wk0/EBMokhuP+uco4Y/7Jw3tW0pl7u351XJL+b8jD6xRX/LlffL/M+Zv+EWk/6fP/ARqP8AhFZP+nz/AMBGr6eUv/fb86eN/wDeb86OSX835B9Zof8APlffL/M+Xv8AhFZP+nz/AMBGo/4RWT/p8/8AARq+o8sP4j+dG5vU/nRyS/m/IPrND/nyvvl/mfLv/CKyf9Pn/gI1H/CKSf8AT5/4CNX1GC3qaXJ9T+dHJL+b8g+s0P8Anyvvl/mfLf8Awikn/T5/4CNR/wAIpJ/09/8AgI1fQd940tYEZrJRcRKSpupJlht8jqBI33z/ALgasRviQA+PtWkDJ4wblh/315f9KpUqj2b+4HisOt6K++X+Z4v/AMIpJ6Xn/gI1H/CKSel5/wCAjV73YeOUuss1uk8Q",
    "5aTT5/tBQerRkLIB9FNdFaXsF/aJc2dwk8EgyskbZBpOnNfa/IPrND/nyvvl/mfMP/CKyf8AT3/4CNR/wisn/T3/AOArV9R5b+8aMt6n86XJL+b8g+s0P+fK++X+Z8uf8ItJ/wBPf/gK1H/CLSf9Pf8A4CtX1H8/XcfzpCX/ALx/Ojkl/N+QfWaH/PlffL/M+Xf+EXf/AKe//AVqT/hF3/6e/wDwFavqHc/94/nSgt3Y/nRyS/m/IPrND/nyvvl/mfLv/CLv63f/AICtR/wi7+t1/wCArV9H6n4ig06ZraMSXV4q72giYDy1/vSOSFRfdjz2BrmLj4hSJKQ2oaVB/sQxTXZH1YbFP4ZpNOO8/wAjWnOFV2hh7+jl/meL/wDCLv63X/gK1H/CMN/euv8AwFavarT4gSySbRfaRdf7DedZsfoX3Ln6kfWun0zxDb6lcNakS2t8q72tbjAcr/eUgkOv+0pIoSb2n+QqlSnSdp4dL1cv8z5u/wCEYb+9df8AgK1B8MH+9df+ArV9Sb29T+dG8juafJL+b8jP6zQ/58r75f5ny1/wjJ/v3X/gK1H/AAjJ/v3P/gK1fUDXJBPNRm7Oepo5JfzfkH1mh/z5X3y/zPmP/hHSM4luhng4tmpP+EcA/juP/AZq+mJNQ2DJfH1NYeteM7PR7UzT3aqvYs5AP07t9AKfJL+b8g+s0P8Anyvvl/meAnQFGcy3H/gM1V5NOtYiPMvWTIyN0BGa9C1Xxv4h8RRudM3WlgOt5dN5SAeoHT/0I/SuXtNHbWbxksI7vX73PzzuTHbR/Vicn8Sv40ckv5vyJeIpdKS++X+ZgGzhIYw3XmbeuIiAPqe1en/DeY3Hge6hdtohuGVS3AORnA/M02z8B2untHN4luI7uZeUsbcbYU/AYLfoPc1vfar66Mdrp1tFHEo2xwpEDgewxgfh+dWlbqc85qTuopel/wBWze8LSBNTK9PMi5+vBrra5zQNBvLWdLu/mQOB",
    "8sSKOPqa6OmQFFFFABRRRQAUUUUAFFFFABRRRQAUo60lKOtAAetJSnrSUAFFFFABRRRQAUUUUAFFFFABRRRQAjosiFHUMjDBVhkGsO88MxbjJYkRn/nmen4HtW7RQByqJNaS7XDKy9uhH+f8msfWvAujeIXNx5bWN/8AeF1bfI+fVl6H69feu/lgjnXEqBsdD3FYWpOunTKkwxE/3JOwPv6GgDyzWvDur6RGf7f05Nb05Ol9ar++iHqw6j+XvWj4c1W6VPM0HUU1i3UZa1uJPLuUHoHPB/4GMe9dxNrAsstI6/LzncA3/wBeuP1i08Ja/diaOaXTNVLfu7qzhdSW9wBg/Xg+9AHX6T4lt7+UW7LLb3eMm2uE8uX6gHhh7qT+FbKzrJkKwJHUdxXk1v4ou7HXI/Deurba4rOixzQrtcM2MZBA55GehHrXo+m2yrPNEZpZkhCgeadxQnnAf7x4x1z1HNAGrkAFmYKoGST0A7mvKPEniaaRI72Nnju9RjLwnobWzJIRV9GkxuZuuCBXf+JJdnhTWfIkYyCxn2Y9dhryLxewj1yMx/6hrG1aDHTy/JXGPyNZVm1DQ78tpwniFz7LUi0jS7rWtShsbGPfNKfwUd2PoBW745tbLRJ7LQrEBvskfmXEpHzSSvjk/gBx2zXX+Crzwr4W0kCTWrB76YBriUSf+Oj2H69am8S+HvCEFnNr+p+dILk+YrLO2ZmYZAUZ/wD1CsFR9zzPVnmnNiVdPlWy7s8y0jV1sWe3vFM+mXGFubfPUf319HXqCOeK9g8I6lNdRXem38wnvdNkWNp/+fiJhuil+rL19wa8KuJo5J5HijEMZYlYwxbaOwyeT9a9X8Ds/wDwlsanO5fD1oJx6PuJXPvtIq8O3qjDOoQvGqlZvc9B24pWIiRnk+VFBLE9gKccYrgfiBoOmafocuowyXFrNuWNYo5m8uQscYKHjpnpiuk8I0vA2qSalFfGUsWml+2Rhj0SQngewK/r",
    "XU4I614/Dpd5Nq+iW/iK2u9Ps7lRBHLC5jLdditj7pJwMcHmvUtN0iw0eIx2EAiB+8SzMzfUkkmgC9g1w3jjxAgkurFyfsNjGjXiK2DcyyZ8q3z2UgFnx/CAO5rud3FeK+NpnNtqZYn5vEUokPssChB+WaqCUpJMUnZXIrOw1Txdc3E6Zmkt4GbIT93GFHyxIBwuegArX8KWHhvUNB1i7vVvFeC3BnUsrbF+8GjOOpK9D06dDVfw14v1rStFUaTpd1fQh1hjjETvGT/Edw5VuRx05r0HTPDlhNp97LdaedPuNWhH261WYEJ1zgjpnJyR/OumpUautvQxhC+p5R/Zs0+ny63ZRSW9pDcCKImTMinBOdwxyMAEjua3/DXiqS0aXUZG+aFl/tSMDC3ELHaLkAdJEJAbH3lOTzVvXNb8T6TDdWo0WysNLsI8rIE3xbOi4Y8EnI4xnJ5rivDc5utcuY5DlJ7C7E3AAKmFienA5Apz9+Db6BH3ZJI+g+CMjketGM9Aax/CdzJc+D9FmmOZHsYWY+p2CrGq6NZa1Esd7HISv3HjlaN1+hUiuM3Od17XmtfHOlwJJtgtV/0kZ4PnHaM/TAP412W0+leJ3mnXUd1ra6TBd3tnDIYHuHzIV2kc5HUgg9OgrtvA/h/SJNFsdWVri6uZU3M89wzhHBwQFztGCD2zQB2u2sbxNqs2mWMMNjsOo30wtrXcMqrHlnI/uqoLH6e9bGa4fxyZT4k0sJncdN1HyMf89fKHT3xmga1Z5xrWuC8kezs5ZDp0chYM5y90/eaQ/wATN1HoMAYq74H0pta8VWkQXMMLefKewVTnH4nA/GsfwZo6+IvEdnpzsyQyEtIVOCEUZOPfjH416w66X8M9CM8UCC6vJ1jAdyx6929FXJOO/wBa44wc5c8tj6WrjIYah9Wor3mvz6nD/EPRk0PxNJ5Khbe6Xz41HRSThh+fP41V8PasZZYNLu7l4oi+bO5B+eymP3XU/wB0",
    "nhl6EGvWdc8PjWPEuiXTorW9mJmlyMhshQq/nk/hXiXizTRoHiS/0+JiUhfMRzztIDKPqAQKJwdOXPEWHxcMXh/q9XVpb/18j2/wzrMms6QJLtFivreV7a8jXosyHDY9jwR7GtViK5DwvLKPEvihSPlE1qzf9dTAPM/UCuoeTYm5ufpXYfNEUzYJPasfVNVg06ES3VwkCMcLuyWc+iqMlj7AGr19M0tpIUm+zomTIyIGfb7Z4H1IP0rgfGHieHwXeR2+n6cr6hcx7/t94xfjJBy33mIx0yAPQ0AT6nfapdQPcIsekWQ63mpY8zH+zDnC/wDAjn2rkIIk1XUWPh6xutevs4fUb/8A1Uf0BwBj3xWxdaXZRzJeeLtRn1e8OGjiCMtsoPI2gYyPyHsa04/Er3MSWmlxrawqMKsafN+AAwv4fnQBDB4GtInju/GmqNqFwOUs4iViX2wMFvwwPrW5LqbJbJZ6VaLZW68IsagH8AOB/P3rPsY5JrlURGe4lbbls7ifc16BpujW+mqGUeZPjmRv6egoA5zS/CU1yRNekwo3JB5kb8+n411dlp9tp8ey2iCZ6t1ZvqasUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFKOtJSjrQAHrSUp60lABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABWV4nWE+H7prgHCqChHUNnAx+datct48udmnWtqp+aeXJHsP8A65oA4HVikMSbeJZTu/D1qC3Mem6Tdazcj93CCsanjeen6kgfiasX1q+p+JVsYOq7Yh7YGSfwyfyrm/iVqkaXFvolm3+j2qhnx3bHGfwJP/AqALXwts21TxTea3ftuW0RpWc/89Hzz+A3H8BXsmnQtHp6tINss5Msg9C3b8BgfhXFeANGGneE7C3cfvdQf7XcD/Y4wPxGwf8AAjXemTI5NAEcsashQruVgVYHoQeoryfXvDk7iHRck6lYKw09nOPt9pkkIp6eZHkjb3HSvXMg",
    "iqmp6RY6zZ/ZtRt1mizuXnDI3ZlYcqfcUmk1Zl06kqclKO6PEPDFpbT6+o1dxb2lmDPdLKMNtT+DB5JJwMdeaseLvF9x4q1MzyAxWsWVt4M8Rr6n3Pf8u1ehal4J1C4YAXenarGBhBq9sTMo9POjIZvxqpD8P72KVTDpnhe2PXzG+03JX3Cudv51j7LSyZ6McwXP7SUdVscFo2kpNAdV1ctBo0LfM/8AFct2iiH8THpkcDkmvYfBGl3Vrb3mr6rEIdS1aQSyQgf6iJRiOL8F/wA8UzTPB0FvqEWpatdy6tqMIxDJMoSKD/rnEPlX9a6UGtIQUVZHHiMTOvLmkWN/FQzwQXDRNNFHI0L+ZGXXOxsYyPQ4J5pQeaXNWc4k0cdzEY540kQkEq4yMg5B/PmlPNVrnUbe0urO3mZhLeSNHCAuQWCljk9uAatUAICQK858faGoursyuIbDWfLzcN9y1vIxiNn9EdflJ7GvSKjuLSG8tpLe5iSaCVSskbjKsD2Ipp21QHiei+OdT8EQXOi3GmW1nJDC5A8o+ZJMcbXZskMD1yOMAYre8Daab3wvq0mrarENV8RxMsAnlBkZAGAbBOeSTx6AVr6r8PJ/IW3057O/09P9XYaoGJgHpFOvzqPY5Fc4/wAMpC+T4TlBz/yy11dv4bo81pzxa7E2Mux+I1zo/gv+x5Yo7u7S4aNxeKZUEOPu4J5w3GPSk0LQrrbLZpEINY1uPy1hUHFjZEgySMDyu4YVVPOPrXS6V8O9SgnV7PTtM0dgf+PqaZr65X3QEBFPvjNdzofhuy8P28iW3mSzzNuuLqdt807erN/TpROonpHqCj3L9pBHaWsVvAu2GFFjjX0UDA/QVMc9qAvpTwMCsiitb28dpCsVtEsUa5wqDABJyT9SaSC0htt/kRJEJHMjBBgFj1OPU4pkWpwXGrXenIH8+0jjkkyPlw+7bg9/umrZoAaTXP8Ai3S7rULC3u9NAOp6bMLq1UnAlIGG",
    "jPsykj8q36QigDw6yePwr4ms/EunpK+iNMVkXb89oTkPDIv8LLk4z1AGK2fi5dQ6tDpWpadew3dgFeM+U4YI5wQSO2QMc+ldxrHhKC/u5L+wupdN1KRdsk8KhknHpLGflcfXmuSufh3ctMTPoegXZPPnW1zNZlvqgBUfhUOGjSOqniWpxnLdE3w68ew2/hm9j16+QJp20ws7fvHQg4UD+IgjA+orlfMfUtcl8W63bOtrLPus7TH7y9kGBHGg6lRhdzdO3eultfAN5BKDbaT4fsCP+W88st86+6q2Ez9RXTaX4WtdNvf7RurifUtUI2/a7oglB6RqOEH0oUdFcUq6UpOmrXH+E9IuNI0pjqDB9SvZnu7xh081/wCEeyjA/A1tuN6FfUVGpGMml3cZzVnMUEZRcGN1ysgKMD39q4T4maKdQ8IC5A3XWkSbWOOWiOAT+WxvzrubpfmZwff6VXu0iuTtuEDW1/EbeZffBx+Y3D8BQB5j4duG8R+CpLUtuvdL+76mPsfyBH/AR61Z8J36W+r24lGI5W8px6BuP54rltCvJvA3j0w3RxHFKba4z0aMnhvp0auo8QaX/Y/iCQQfLbzgTQH09vwI/SgD0vw1psUOrX8pyZImCRg/wg8kiulrm9DvFnurS7X7t9bDP+8OcfzFdJQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFKOtJSjrQAHrSUp60lABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABXB+KbgXfi+3tyf3drGC3t1Y/wBK7z615Lq175tzq19/FOxhi99xx/KgBdDuYtP07VvEl593D+X688nH5qv415bpdrN4r8WwRTMS95cbpW9F6t+QBrrPiNqJ07SNO8PRNyFE1xjuc9Pzz+QrK+HF3ZWWrXU09zDBdm3MdoJm2oXbrlug6Y59aAPZLSRDPJMowoxFGB2ReOPxz+AFc54zvp117SbePUPskEsEzPu1A2aMQVxl8Hnr",
    "gYrTtrlIQkDq8LbQEWUYLAdwejfUE1aPlTAebFFJjpvQNj86AKeka3bweLdRsrrVYzEYrIWiT3IPmbo+SmfvEkgkjrkVy+i6xqn9g6nfRam9xqMME5hhbUjLICHxuNuRxtXJByegOOa6nXbq5s4Leew0+KYq+JZha+c9ugHDBAQTzgcHgVUk1y5tnsm061sLl9WhUWN1DD5aiXjzN/8AFt2kuO/ykGgB1rBpt5o17/ZvivVL2dbT7TIyX5LIyqSD0+TJ4K/4VmSLqC+HPCpt7/WLq41V/MnRb8o7nyd21GP3RnnFbMOq38PiCXTZo7OxW4aRLRmsmKXZCcP5itjOeShGQOM5qpJr2tWXh7VNQl/syR9PvDZwRx2jDDCRELgbu4Y4Ax25oAq6ZrmrG30KWG+mu5ZrXUpUhdy7MyLmKOTgbnU+3etHw1rOl2ekjWbnxNdX98ti093aTXYYBgAWHlY+QhvlH1pbrxFd6NNo8l7bwLHPNM13I9kbeSKIFF3hSxI5fJOeQKgtPEeq6npupXdpolrLKt5GgVLbzJBbOm8OyZBkbBU4yOvtQBT0rWtZtNH1221WbUoru40p7+2e6GwpIqt5giOT8oyhHfg13XhjWLTVdDtPs9/DeTxW0X2gpKHZXKjO7HQ5B6+9Jod5DrWi2l8Xt7rzEYGVYSgPJVgFbJXpgjPY1pwW0FsD9nghi3dfLjC5+uKAMjX+fEXhb/r9m/8ASeSugHArn9f/AORh8Lf9f0v/AKTyV0HagBc8Vyvi2/1O18Q+GotJKvLPNcBreSYxxzARZwxAPTqODzV/xRqGradpsUui2f2mVpQspEfmtEmDlljDKXPTgGsK58YXkVnpeoW8FtqsOoRtb2rxW7Rsl7kgAhjlUODnuNp5oAzF1XW0sNX1a5uJEl0jXFa4t4p2dPs+1RJGOmVGd3Tsaiu9b1HUdPVra61Itr2pSSW62ZJlhsoeMoMjBYgH8a3pdZ1aw8QxabqUdlbW",
    "ty8cIuPsLPFeSsvzfOG+Q7sgKwOQOTSaVqOt/btbS5fSdmiIYx5NmyF8xbxg7ztAOMjvigDD/wCEmudT03wza6pqdxpds0k9rq1wknlSLPEAFR3/AIN3U+ua7/w8bQ6PGNP1OXU7dXcLcyzeazc8jf3x0rjIvEury+BLzVpLW0aeZbd44pdNaKKR5HUdS58wYPUY7VZ0zx3cX+vaNYWNnbpZzQMl0u3aY7kRs5iXsNu0Z/3qAO+HFBbFcr4S8R6jql9c2etCK2vY4hL9i+yvE8YLYyHLFZF6DcMc107GgDndOP8AxX3iD/r1s/8A2pVD4iz3iWGjxWE1xHJc6nHCVt5/JeQFW+Xf2zgVd0//AJH7X/8Ar0s//albskaOV3orbTuXcoO0+o9DQB5FLq2qf2BpYm1i4iMur3MLrcagYXgRVwIpZsdQRnOP4hV7XNWOnCwnk1w3FlDZozW1rrO2feXOZFYjE/TGPbpW7r+qXi+MotFsrS1eGS1W7djpxuW3mQoScMMDAHzGq2q63eQrrd7Z22kpZ6BN5X2eeA+bJgKSVYH5M7vlwDnFADYYY7LxrfyXOsax/Z9lYR6l5Ut223JZiwZf7uB92sjR9d1thfrfHUkk1nT57myEy7RHMoZlSHnp5bKe3Ira1PxLPDpXiK7WC1Z7K4ggiWWPJeOQIcPzz98+1Qap4ovYJSsYtYymrXVkjtavMVSKPcpCqd24ngkdjQBBF4k/tFPC8VhqjTzLp9xJfrFISVItuDJ6EP696zvC+t3+mul7qt5di2XRzfPHdXfnC6OcBoxj5ORgjOeRV+Txld6bfRtqWmw2dvLYLNcqseJFuGEhRCe4by8YPOWFV4vEurXGnWCTWNmL17+4snSKzM2wRoGwqbgeuc89vagCGw1rVrLR9btdTn1BLybTGv7Z7kbGSQKfMWM5PyglSPoal0/VdSGsaHoWpXty0u2aRp1cr9qgkg3RsT/eVtw9ioNW9S8S3WnapPBe",
    "21vJbxWaE3DREGOeRGKhlJOEYrtI7EgE1JHqWsXGo6VGJNMWO8szcgtasXiAVNyg7u+/j6UALodm6a7rPmanqtwlhcJFDHNdl0KtFk7hj5jk8VvITdWstsrYl+/ET2YHI/8AHgPzqCaUKWwANxyfU/X1qqt/FazpJI4UZxju3sB3oA8++K+mK93Ya3Cm2O8i8qTjo6jjPvjj/gJrY0O4bxd8PQM79T0c4x/E6Y/qB+an1p/jnVdLfw3qFheyGOd5BPaRsAZA/B5XqoJL9cferjPh74jbQPFEDu+22uSIJx7E8N+B/TNAHqnha/36CrKctYXIcf7jc/1avRQQwDKcgjI+lec2lmNG8V3ViABaahE3lAdAwO4L9R8w+hFdtoVybnR4C3348xN9V4/ligDQooooAKKKKACiiigAooooAKKKKACiiigAooooAKUdaSlHWgAPWkpT1pKACiiigAooooAKKKKACiiigAooooAKKKKAI7lHltZY4mCSOjKrHoCRwa8t1Pw5eWMXlahFKsIORNCcgH1yOlerUHkEHkHqDQB4n/YYlUomy8hPJgmBb8gef++Tmsu5+HljqIY6bctY3I6w3GXiJ9Nw+ZfxB+te033hjT7wllj8hz1MfAP1HSsmbw1d2zbk2XSDpn7w/r+RoA8YYeLfBC+XPFIbDOdsiie2f6HkD8MGt/SPiDYTqEud+nyH++WmgP0P30/8eFem290sDFJ4mQkYZXHUeh45/EVg638OfDmvEywwNYTn/lpa4UE+6fdP4YNAFG51TQ7+wWTWI45beNsq4Rp0Ge4ZAcA++PpUkfjLwgotgLuAC1OYALeQCLjHy/LxxxXG33w28SaBO0+jT/alHe3YxyEe6Hr+BNc7c36rMYde0lorgdZIl8iX6lSNp/IUtSo8vU9Sg8T+BrfUGv4ZbZLxiWMwtpN2T1P3eCe5FTSeLvBU9ncWs9zA9vcyGWaM28m2RyQSx+XrkA/hXkX9nxXWDpl/",
    "BKT/AMsbjEMn05O0/gaoXS3lhKYrqBoZP7rpj/8AXS940So9Wz2qHxL4ChgWFJLcRKsihPs0hADgBxyvQgDP0qS48UeAryNo7mS3lV2V2BtpOWVdqnhey8fSvC/tknov5U5buZuFVT9Fpe8Vah3Z9BWvxB8IWlvHBb3yRQxjakcdtIFUegG2rA+JXhX/AKCZ/wC/En/xNfPIluh/yxP/AH7NPWW87QN/36NF5lKOG6tn0DJ8Q/CE0kMkt8ryQMXiZraQlGIIJHy8HBI/Gph8SvC5/wCYmf8AwHk/+Jr59RtQP3bSQ/SFqmUar2sJj/27t/hU3qdkaKOD6yl+B7lqXjPwTrNstvqVwlzErb1WS2k+VumRxweTTh4z8E+RaQCaIRWbiS3QWsgWJgCAVG3g8n868PVdY/6B0/8A4Ct/hTlGsj/mHXH/AICv/hSvV7ItQwHWUvuR7QfE3gJtW/tQmE3+7f55tZC24DGemM471ZXxt4MVrsrcqDef8fJFtJ+942/Nx6cV4bNPqttEZZ7OSKPON8luyjPpkimWt9qN9dJb2kImmcnYiR5JwMn9AaL1eyK5Mu/ml9yPZ7fW/h3aQyw26QRxTbfMRbaXDbTuXt2IBqwPFfgVJhKrwiQSyT7hayZ8yQYdvu9WHBrwsazdMRgRnPT5Ktk61/0Dp/8AwEf/AAovV7IPZ5d/PL7kez6Z4p8B6I0jaY8Nq0gAcx2sgJA6DOOnt0q8fiT4X/6CR/8AAeT/AOJrwc/2znnT5/8AwFb/AApjNq2ebGb/AMBm/wAKL1eyJcMB0lL7ke6L8QPCEd1LcJeqs8qqskgtpNzhc7QTt7ZP508/Erwv/wBBM/8AfiT/AOJrwUvqmf8Ajyl/8B2qNpNRH3rRx9YGp3qdkQ4YPpKX3I9p1LxN4B1e5W41FoLiZU2CR7aXIXOcZA6ZJqLUvEXw+1a48/UBa3ExAUyPaybiB0BIXnHvXi/n3x/5d2/79GjzL7r9mf8A",
    "79GneZPLhf5n+B7RP4m8CXOoR39w1rJeR7dsrWshI29P4ecdielP/wCEx8HJOJkuoVlWV5w4t5MiRxhn+71I4NeHNeTg4ZVH1XFNN7L6J+VF5kuOH7s9uufF/g+6d3muoJHfZuL28h3bCSmfl7Ekj61VuvEXgm7iMdxLayIZmnINvJgyN95vu9T3ryC3+23sgS1t2mb0jjJrQ/swWwzql/bWx/54x/vZfyHA/E0/eJaod2eof8JP4O8iSHz7fy5YlhkU28hDov3VPy8gdqsw6/o1yiS2H7020ZjjdYGURocAqGYBQOB37V5MNRs7dgmn2DTzE/LJd/Oc+0Y4/PNblr4K8X+KQjXMckFueVa7bykA9kHP5CqV+plLk+zc39Y8eadbEqLjzn/552mGP4yH5R/wEH61yU3ifWtauGtdEtZIDJwVtg0kzD/ak+9j6YFdzpXwj0uy2vqU0uozDkxpmOL9PmI/EVttFa2NubW2MFrAP+WNqgOfrjgn3JzTIPMLTwDP/rdYuRExOTDEQ7/8Cb7o/U+1dBpeii0BOk2KxAfeuXIJ/wC+26fhj6V2Vpodze4a1sjjtLcYI/AdP0Nblv4OjkKvqV1JMw/gQ4A/H/DFAHHaba7byNjPPdTowZY4QWG76n+g/GvQdAtLm1t5zdIIzLJvVM5K8c5q9aWVtYx+XawpEvfaOT9T1NT0AFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABSjrSUo60AB60lKetJQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFADJYY51xLGrj/AGhms6fQYHyYJHgY+hyK1KKAObl0zVrY/ujHcIO2ev51mXr290ht9a00Mh42yxhx+R5H4V29I6JKu2RVdfRhkUAeRah8L/Dur5fS7p7GU9FU+Ymf91uR+dc7e/DvxhosLR2ix6pZKc+WhEg/79vyD9K9suvDenXPPk+U3YxnGPwqi2galaHdp+pFwOkc",
    "o/8A10AfON5DbxytFe2NzptyOqhSV/74b5h+Z+lU18yBibef/gUbEH/GvpK8M9xEIfEGhw3sI43GMOPwznH6VzGpfDnwdq4Z7RrjS5j2RsqP+Atn9CKAPILfXtYtcfZ9WvosHOFuGH9a0YPHfieDG3Wrpsf89CH/APQga6HUfg7q8G59Kv7PUYx0Xd5b/keP1rktS8OaxozEalpl1bgfxlCU/wC+hx+tAG/b/FXxNCuGntpv9+AD/wBBxWhB8ZNaQDz7Kzk91aRP/ZjXnwIPQinYoA9PtvjVNvUXOkkA9WjuyB+RWuqtPiPZXc4hgubYyNwoa7xn6ZArwbFJtB60Aeh/F1NQuLzT9QuJZWtGUxCEt8sTjnI/3h39jWv8Nbbw8scms6fb3sN5bjyZGupVdFJGSVIA7cc9M+9eX/2nfHTm09rqVrMkN5LtuVSOhAPQ/Snvq94dFTSUYR2ayGV1QYMrHux74wMDpxQB0/j/AE7w3oty1vp9vf8A2+Yedkyr5CKxPQbcnuMcY9a7PwDq2rWPgyGXULppYiWaHznx5UQ4ALHtkEjPQYryG81C71GG1iu5PNFqhjjZh8wTOdpPcA5x6Zp+oane6oyfbLh5EjULHH0SNQMAKo4AxQB7DqnxWg0+zdoxBdXAYBYorvOR3JIBxXNT/GfUnB8nS7dM9N87tj8sV5yBRQB203xb8QyE+WllFkdo2b+bVnXHxJ8UTHjURF/1zhQf0rmCQOpFNUmSQJErO54CqMk0AbNz4w8RXQxLrd/j0WYoP0xVB9S1K4O2S+vJc8YaZj/WtXTvAfiTVArRabJBG3/LS5IiH/j3P6V0tj8JdgDa1rcMK944F3H82wP0NAHCLZIfnvL2OEf3QTI5/AcfmRVyyjgmlEWk6RcajccfNMC4/wC/af1Jr1HTfCngnTGGy0m1KYdDKTJk/wC6ML+ldfZTXqxCHStDFtD0G8CNR/wEYoA8ss/h9411yNY7po9Otf8Ank7i",
    "NQP+uaD+ddNpfwT0mz2vrOoy3BHVEIiQ/wAyfzFd0mmatcf8fN+kCnqkK/1q3BodtFzI0szert/hQBk6fpfh/QF2aVYxxtjG6GP52/4EeT+dWy9/cnFtZCJT/HL/AIcVsxQRQDEUaIP9kYp9AGGPDslwc3947r/zzj4X/P4VftdGsLMgw2ybh0ZvmP61dooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKUdaSlHWgAPWkpT1pKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAM4qGazt7j/AF0Eb+5Xn86mooAzJfD9m5zEZIW7FWzj86jOlXcQIhvAy+kg61r0UAcTq3grTdR3NqOgW0jHrLbDy2P/AHzXH3/wn0aXJsNSvbBz0S4j8xR+Iwf517NSMquMOoYe4zQB87X/AMLfENrlrE2mpx9jbygP/wB8tg/zrmr7StS0lsalpt3a+8sTKPzIr6jm0iymHMIU+qcVXbRXVCtvezKp/gk+dfyNAHy0rxt0YfjxUgjzXv2qfD+z1AsbrR9MuCerxKYH/NMf1rj9T+FOnxsxh/tSz/2AVmX8OAaAPMzGBySAPeojJHuCqS7HoFGa9d0v4a6dHho9Eu79/wDnpfTFV/75Xb/M112neE7q0AFsmmaWvpaWy7/++sf1oA8N0/wZ4j1UB7XSJ0iP/LWceUv1y2K27b4YMmG1nXbSD1jtlMzfnwP517QPCdrK26+ubu7bvvkwD+Aq/baNp1pjyLKBSO5XcfzNAHlWmeAvDsRBt9M1HVpB/FMxVPyXH611th4c1KBNthp+n6XGf7iqG/QZrtRwMDgelFAHMR+D55W3X2qyvnqsS4/U/wCFX7fwnpFudxtvOb+9Mxb/AOtWxRQBHDbw264ghjiH+woH8qkoooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKA",
    "CiiigAooooAKUdaSlHWgAPWkpT1pKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAClyRSUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUo60lKOtAAce9HHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nRx6H86KKADj0P50ceh/OiigA49D+dHHofzoooAOPQ/nQMUUUAf/9k="
  ].join("");
}
function driveImageUrl_(fileId, size) {
  fileId = String(fileId || "").trim();
  if (!fileId) return "";

  return "https://drive.google.com/thumbnail?id=" +
    encodeURIComponent(fileId) +
    "&sz=w" +
    encodeURIComponent(String(size || 1200));
}

function driveImageDataUrl_(fileId) {
  fileId = String(fileId || "").trim();
  if (!fileId) return "";

  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    const contentType = String(blob.getContentType() || "");

    if (contentType.indexOf("image/") !== 0) return "";

    return "data:" +
      contentType +
      ";base64," +
      Utilities.base64Encode(blob.getBytes());
  } catch (err) {
    Logger.log("driveImageDataUrl_ error: " + err);
    return "";
  }
}
