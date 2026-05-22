// =====================================================
// FILE: 05_Customer.gs
// =====================================================

function getCustomerOrdersBySupervisor(supervisorName, sessionToken, companyId) {
  supervisorName = String(supervisorName || "").trim();
  companyId = String(companyId || "").trim().toUpperCase();
  const session = requireNamedSession_(sessionToken, ["SUPERVISOR", "OWNER", "ADMIN"], companyId, supervisorName, "supervisor");
  const role = String(session.role || "").trim().toUpperCase();

  if (role === "SUPERVISOR") {
    supervisorName = String(session.name || "").trim();
    companyId = String(session.companyId || companyId || "").trim().toUpperCase();
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  if (!sh) {
    throw new Error("Missing sheet: " + CFG.SHEET_WORK_ORDERS);
  }

  const allowedStores = getSupervisorAllowedStores_(supervisorName, companyId);

  if (allowedStores.length === 0) {
    throw new Error("Supervisor not found or has no assigned stores: " + supervisorName);
  }

  const allowedSet = new Set(allowedStores);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function(h) {
    return String(h).trim();
  });

  const idxNSN = headers.indexOf("NSN");
  const idxCompany = headers.indexOf("COMPANY_ID");

  if (idxNSN === -1) {
    throw new Error("Missing NSN column in WORK_ORDERS.");
  }

  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const nsn = normalizeNSN_(row[idxNSN]);
    const rowCompany = idxCompany >= 0 ? String(row[idxCompany] || "").trim().toUpperCase() : "";

    if (!allowedSet.has(nsn)) continue;
    if (companyId && rowCompany && rowCompany !== companyId) continue;

    const obj = {};

    headers.forEach(function(h, c) {
      let value = row[c];

      if (value instanceof Date) {
        value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a");
      }

      obj[h] = value;
    });

    obj.ROW_NUMBER = i + 1;

    enrichWorkOrderObject_(obj);

    obj.PDF_EN_URL = getInvoicePdfByWO(obj.WO_NUMBER);
    obj.HAS_INVOICE_PDF = obj.PDF_EN_URL ? "YES" : "NO";

    obj.QUOTE_EN_URL = obj.QUOTE_EN_URL || "";
    obj.QUOTE_ES_URL = obj.QUOTE_ES_URL || "";
    obj.QUOTE_STATUS = obj.QUOTE_STATUS || "";
    obj.PM_REPORT_ES_URL = obj.PM_REPORT_ES_URL || "";
    obj.PM_REPORT_EN_URL = obj.PM_REPORT_EN_URL || "";
    obj.HAS_PM_REPORT_PDF = (obj.PM_REPORT_ES_URL || obj.PM_REPORT_EN_URL) ? "YES" : "NO";

    result.push(obj);
  }

  return result.reverse();
}

function getSupervisorAllowedStores_(supervisorName, companyId) {
  const fromSheet = getSupervisorStoresFromSheet_(supervisorName, companyId);
  const fromConstant = getSupervisorStoresFromConstant_(supervisorName);

  const all = fromSheet.concat(fromConstant)
    .map(function(n) {
      return normalizeNSN_(n);
    })
    .filter(Boolean);

  return Array.from(new Set(all));
}

function getSupervisorStoresFromConstant_(supervisorName) {
  supervisorName = String(supervisorName || "").trim();

  if (typeof SUPERVISORS === "undefined") {
    return [];
  }

  return SUPERVISORS[supervisorName] || [];
}

function getSupervisorStoresFromSheet_(supervisorName, companyId) {
  supervisorName = String(supervisorName || "").trim().toUpperCase();
  companyId = String(companyId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_STORES);
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  const idxNSN = headers.indexOf("NSN #");
  const idxSupervisor = headers.indexOf("SUPERVISOR_NAME");
  const idxActive = headers.indexOf("ACTIVE");
  const idxCompany = headers.indexOf("COMPANY_ID");

  if (idxNSN === -1 || idxSupervisor === -1) return [];

  const result = [];

  for (let i = 1; i < data.length; i++) {
    const rowSupervisor = String(data[i][idxSupervisor] || "").trim().toUpperCase();
    const rowCompany = idxCompany >= 0 ? String(data[i][idxCompany] || "").trim().toUpperCase() : "";

    const active = idxActive >= 0
      ? String(data[i][idxActive] || "YES").trim().toUpperCase()
      : "YES";

    if (rowSupervisor !== supervisorName) continue;
    if (companyId && rowCompany && rowCompany !== companyId) continue;
    if (active === "NO") continue;

    result.push(normalizeNSN_(data[i][idxNSN]));
  }

  return result;
}

function getInvoicePdfByWO(woNumber) {
  woNumber = String(woNumber || "").trim();

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INVOICES");
  if (!sh) return "";

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return "";

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  const idxWO = headers.indexOf("WO_NUMBER");
  const idxPdfEn = headers.indexOf("PDF_EN_URL");
  const idxPdfEs = headers.indexOf("PDF_ES_URL");

  if (idxWO === -1) return "";

  for (let i = data.length - 1; i >= 1; i--) {
    const rowWO = String(data[i][idxWO] || "").trim();

    if (rowWO === woNumber) {
      const pdfEn = idxPdfEn >= 0 ? String(data[i][idxPdfEn] || "").trim() : "";
      const pdfEs = idxPdfEs >= 0 ? String(data[i][idxPdfEs] || "").trim() : "";

      return pdfEn || pdfEs || "";
    }
  }

  return "";
}

function getSupervisorCompany(supervisorName) {
  supervisorName = String(supervisorName || "").trim();

  const supervisorClients = {
    "Zuelem Santiago": "McDonald's"
  };

  return supervisorClients[supervisorName] || "McDonald's";
}

function sendSupervisorOrderMessage(rowNumber, supervisorName, message, sessionToken) {
  if (!rowNumber || isNaN(rowNumber)) {
    throw new Error("Fila inválida.");
  }

  message = String(message || "").trim();

  if (!message) {
    throw new Error("El mensaje no puede estar vacío.");
  }

  const to = "invoice.palaciospowersystem@gmail.com";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("Missing sheet: " + CFG.SHEET_WORK_ORDERS);

  rowNumber = Number(rowNumber);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  const woNumber = getCellByHeader_(sh, rowNumber, headers, "WO_NUMBER");
  const companyId = getCellByHeader_(sh, rowNumber, headers, "COMPANY_ID");
  const client = getCellByHeader_(sh, rowNumber, headers, "CLIENT");
  const nsn = getCellByHeader_(sh, rowNumber, headers, "NSN");
  const status = getCellByHeader_(sh, rowNumber, headers, "STATUS");
  const equipment = getCellByHeader_(sh, rowNumber, headers, "REPORTED_EQUIPMENT");
  const problem = getCellByHeader_(sh, rowNumber, headers, "REPORTED_PROBLEM_ORIGINAL");

  const session = requireNamedSession_(sessionToken, ["SUPERVISOR", "OWNER", "ADMIN"], companyId, supervisorName, "supervisor");
  const sentBy = getSessionActorLabel_(session);
  const allowedStores = getSupervisorAllowedStores_(
    String(session.role || "").toUpperCase() === "SUPERVISOR" ? session.name : supervisorName,
    companyId
  );

  if (String(session.role || "").toUpperCase() === "SUPERVISOR" &&
      allowedStores.indexOf(normalizeNSN_(nsn)) === -1) {
    throw new Error("No autorizado para enviar mensajes de esta orden.");
  }

  if (!woNumber) {
    throw new Error("No se encontró WO_NUMBER para esta orden.");
  }

  MailApp.sendEmail({
      to: to,
      subject: "Supervisor Message - " + woNumber,
      htmlBody:
        "<h2>Supervisor Message</h2>" +
        "<p>A supervisor sent a message about this Work Order.</p>" +
        "<hr>" +
        "<p><b>Supervisor:</b> " + sentBy + "</p>" +
        "<p><b>Company:</b> " + (companyId || "") + "</p>" +
        "<p><b>Work Order:</b> " + (woNumber || "") + "</p>" +
        "<p><b>Customer:</b> " + (client || "") + "</p>" +
        "<p><b>NSN:</b> " + (nsn || "") + "</p>" +
        "<p><b>Status:</b> " + (status || "") + "</p>" +
        "<p><b>Equipment:</b> " + (equipment || "") + "</p>" +
        "<p><b>Problem:</b><br>" + (problem || "") + "</p>" +
        "<hr>" +
        "<p><b>Message:</b></p>" +
        "<p style='font-size:16px; background:#f3f4f6; padding:12px; border-radius:8px;'>" +
        message +
        "</p>"
    });
    

  addNotification_(
    companyId || CFG.DEFAULT_COMPANY_ID,
    "ADMIN",
    woNumber,
    "SUPERVISOR MESSAGE",
    "💬 Supervisor message for " + woNumber
  );

  addLog_(
    companyId || CFG.DEFAULT_COMPANY_ID,
    woNumber,
    "SUPERVISOR MESSAGE SENT",
    "",
    "EMAIL SENT",
    sentBy,
    message
  );

  return true;
}

function testSupervisorEmail() {
  MailApp.sendEmail({
    to: "invoice.palaciospowersystem@gmail.com",
    subject: "TEST EMAIL FROM APP SCRIPT",
    htmlBody: "<h2>Test OK</h2><p>If you receive this, MailApp works.</p>"
  });
}
