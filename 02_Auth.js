// =====================================================
// FILE: 02_Auth.gs
// =====================================================

function loginUser(email, password) {
  email = String(email || "").trim().toLowerCase();
  password = String(password || "").trim();

  if (!email || !password) {
    throw new Error("Debe escribir email y password.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_USERS);

  if (!sh) throw new Error("No existe la hoja USERS.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error("La hoja USERS no tiene usuarios.");

  const headers = data[0].map(function(h) {
    return String(h).trim().toUpperCase();
  });

  const idxEmail = headers.indexOf("EMAIL");
  const idxName = headers.indexOf("NAME");
  const idxPassword = headers.indexOf("PASSWORD");
  const idxRole = headers.indexOf("ROLE");
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxActive = headers.indexOf("ACTIVE");

  if ([idxEmail, idxName, idxPassword, idxRole, idxCompany, idxActive].indexOf(-1) !== -1) {
    throw new Error("USERS debe tener EMAIL, NAME, PASSWORD, ROLE, COMPANY_ID y ACTIVE.");
  }

  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][idxEmail] || "").trim().toLowerCase();
    const rowPassword = String(data[i][idxPassword] || "").trim();
    const active = String(data[i][idxActive] || "").trim().toUpperCase();

    if (rowEmail === email && rowPassword === password) {
      if (active !== "YES") throw new Error("Usuario inactivo.");

      return {
        success: true,
        email: rowEmail,
        name: String(data[i][idxName] || "").trim(),
        role: String(data[i][idxRole] || "").trim().toUpperCase(),
        companyId: String(data[i][idxCompany] || "").trim().toUpperCase()
      };
    }
  }

  throw new Error("Email o password incorrecto.");
}

function loginUserForCompany(email, password, requestedCompanyId) {
  const user = loginUser(email, password);
  const requested = String(requestedCompanyId || "").trim().toUpperCase();

  if (
    requested &&
    String(user.role || "").trim().toUpperCase() !== "OWNER" &&
    String(user.companyId || "").trim().toUpperCase() !== requested
  ) {
    throw new Error("Este login pertenece a otra compania.");
  }

  return user;
}
