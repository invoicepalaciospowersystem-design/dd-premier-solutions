// =====================================================
// FILE: 16_Quote.gs
// =====================================================

function generateQuoteNumber(companyId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    const year = Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy");

    const company = String(companyId || CFG.DEFAULT_COMPANY_ID || "GENERAL")
      .trim()
      .toUpperCase();

    const key = "LAST_QUOTE_NUMBER_" + company + "_" + year;

    let last = Number(props.getProperty(key) || 0);
    last++;

    props.setProperty(key, String(last));

    return "Q-" + year + "-" + ("0000" + last).slice(-4);

  } finally {
    lock.releaseLock();
  }
}

function getWorkOrderForQuoteByWO(woNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  const shStores = ss.getSheetByName("STORES");

  if (!shWO) throw new Error("No existe la hoja WORK_ORDERS");
  if (!shStores) throw new Error("No existe la hoja STORES");

  const woValues = shWO.getDataRange().getValues();
  if (woValues.length < 2) return null;

  const woHeaders = woValues[0].map(h => String(h).trim().toUpperCase());

  const idxWO = woHeaders.indexOf("WO_NUMBER");
  if (idxWO === -1) throw new Error("No existe columna WO_NUMBER en WORK_ORDERS");

  const targetWO = String(woNumber || "").trim();
  let found = null;

  for (let i = 1; i < woValues.length; i++) {
    const row = woValues[i];
    const currentWO = String(row[idxWO] || "").trim();

    if (currentWO === targetWO) {
      const obj = {};
      woHeaders.forEach((h, c) => obj[h] = row[c]);
      found = obj;
      break;
    }
  }

  if (!found) return null;

  const nsn = String(
    found["NSN #"] ||
    found["NSN"] ||
    found["NS"] ||
    ""
  ).trim();

  const storeInfo = getStoreInfoForQuote_(ss, nsn);

  return {
    woNumber: found["WO_NUMBER"] || "",
    client: storeInfo.client || found["CLIENT"] || found["CLIENTE"] || found["CUSTOMER"] || "",
    address: storeInfo.address || "",
    nsn: nsn,
    problemEn: found["PROBLEM_EN"] || found["PROBLEM"] || "",
    problemEs: found["PROBLEM_ES"] || "",
    equipment: found["REPORTED_EQUIPMENT"] || ""
  };
}

function getStoreInfoForQuote_(ss, nsn) {
  const sh = ss.getSheetByName("STORES");
  if (!sh || !nsn) return { client: "", address: "" };

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { client: "", address: "" };

  const headers = values[0].map(h => String(h).trim().toUpperCase());

  const idxNS = headers.indexOf("NSN #");
  const idxClient = headers.indexOf("CLIENTE");
  const idxAddress = headers.indexOf("ADDRESS");
  const idxCity = headers.indexOf("CITY");
  const idxState = headers.indexOf("STATE");
  const idxZip = headers.indexOf("ZIP");

  if (idxNS === -1) throw new Error("No existe columna NSN # en STORES");

  const target = String(nsn || "").trim();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const current = String(row[idxNS] || "").trim();

    if (current === target) {
      const street = String(row[idxAddress] || "").trim();
      const city = String(row[idxCity] || "").trim();
      const state = String(row[idxState] || "").trim();
      const zip = String(row[idxZip] || "").trim();
      const client = String(row[idxClient] || "").trim();

      return {
        client: client,
        address: [street, city, [state, zip].filter(Boolean).join(" ")]
          .filter(Boolean)
          .join(", ")
      };
    }
  }

  return { client: "", address: "" };
}

function createQuote(data) {
  try {
    return createQuote_(data);
  } catch (err) {
    notifySystemError_("QUOTE_CREATE_ERROR", err, {
      module: "QUOTE",
      companyId: data && data.companyId,
      woNumber: data && data.wo_number,
      quoteNumber: data && data.quote_number,
      nsn: data && data.ns_number
    });
    throw err;
  }
}

function createQuote_(data) {
  if (!data) throw new Error("No quote data received.");

  setupQuotesModule();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("QUOTES");

  const address = normalizeQuoteAddressForTemplate_(data);

  data.direccion = address.street;
  data.ciudad = address.city;
  data.estado = address.state;
  data.zip_code = address.zip;

  const targetFolder = getQuoteTargetFolder_(data);

  const pdfEs = generateQuotePdfFromTemplate_(data, "ES", targetFolder);
  const pdfEn = generateQuotePdfFromTemplate_(data, "EN", targetFolder);

  const quoteHeaders = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  const quoteRow = {
    CREATED_AT: new Date(),
    COMPANY_ID: data.companyId || CFG.DEFAULT_COMPANY_ID || "",
    ROW_NUMBER: data.rowNumber || "",
    QUOTE_NUMBER: data.quote_number || "",
    WO_NUMBER: data.wo_number || "",
    CLIENT: data.nombre_cliente || "",
    STATUS: "CREATED",
    PDF_EN_URL: pdfEn || "",
    PDF_ES_URL: pdfEs || "",
    APPROVED_BY: "",
    APPROVED_AT: "",
    DATA_JSON: JSON.stringify(data)
  };

  sh.appendRow(quoteHeaders.map(function(h) {
    return quoteRow[h] !== undefined ? quoteRow[h] : "";
  }));

  const quoteSheetRow = sh.getLastRow();

  updateWorkOrderQuoteLinks_(data.rowNumber, pdfEn, pdfEs);

  let emailResult = null;
  try {
    emailResult = sendQuoteCreatedClientEmail_(data, quoteRow, quoteSheetRow);
  } catch (emailErr) {
    Logger.log("ERROR sendQuoteCreatedClientEmail_: " + emailErr);
    notifySystemError_("QUOTE_EMAIL_ERROR", emailErr, {
      module: "QUOTE",
      companyId: data.companyId || CFG.DEFAULT_COMPANY_ID || "",
      woNumber: data.wo_number || "",
      quoteNumber: data.quote_number || "",
      nsn: data.ns_number || ""
    });
    emailResult = {
      sent: false,
      status: "ERROR: " + emailErr
    };
  }

  addNotification_(data.companyId, "ADMIN", data.wo_number, "QUOTE", "📄 Quote created for " + data.wo_number);
  addNotification_(data.companyId, "SYSTEM", data.wo_number, "QUOTE", "📄 Quote created for " + data.wo_number);
  addNotification_(data.companyId, "ORDENES", data.wo_number, "QUOTE", "📄 Quote created for " + data.wo_number);

  return {
    success: true,
    pdfEnUrl: pdfEn,
    pdfEsUrl: pdfEs,
    emailResult: emailResult
  };
}

function setupQuotesModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("QUOTES");
  if (!sh) sh = ss.insertSheet("QUOTES");

  const headers = [
    "CREATED_AT",
    "COMPANY_ID",
    "ROW_NUMBER",
    "QUOTE_NUMBER",
    "WO_NUMBER",
    "CLIENT",
    "STATUS",
    "PDF_EN_URL",
    "PDF_ES_URL",
    "APPROVED_BY",
    "APPROVED_AT",
    "DATA_JSON"
  ];

  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const normalizedCurrent = current.map(function(h) {
      return String(h || "").trim().toUpperCase();
    });

    headers.forEach(function(h) {
      if (normalizedCurrent.indexOf(h) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
        normalizedCurrent.push(h);
      }
    });
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, sh.getLastColumn());
}

function normalizeQuoteAddressForTemplate_(data) {
  const street = String(data.direccion || "").trim();
  const city = String(data.ciudad || "").trim();
  const state = String(data.estado || "").trim();
  const zip = String(data.zip_code || "").trim();

  return {
    street: street,
    city: city,
    state: state,
    zip: zip,
    full: [street, city, [state, zip].filter(Boolean).join(" ")]
      .filter(Boolean)
      .join(", ")
  };
}

function getQuoteTargetFolder_(data) {
  const rootId = CFG.QUOTES_FOLDER_ID || CFG.INVOICES_FOLDER_ID;
  if (!rootId) throw new Error("Falta CFG.QUOTES_FOLDER_ID o CFG.INVOICES_FOLDER_ID");

  const root = DriveApp.getFolderById(rootId);

  const company = String(data.companyId || CFG.DEFAULT_COMPANY_ID || "GENERAL")
    .trim()
    .toUpperCase();

  const year = Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy");

  const nsn = String(
    data.ns_number ||
    data.nsn ||
    data.NS ||
    "SIN_NSN"
  ).trim();

  const companyFolder = getOrCreateFolder_(root, safeFolderName_(company));
  const yearFolder = getOrCreateFolder_(companyFolder, year);
  const storeFolder = getOrCreateFolder_(yearFolder, safeFolderName_(nsn));

  return storeFolder;
}

function getOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function generateQuotePdfFromTemplate_(data, lang, targetFolder) {
  const templateFile = DriveApp.getFileById(CFG.QUOTE_TEMPLATE_ES_DOC_ID);

  const quoteNumber = data.quote_number || "QUOTE";
  const woNumber = data.wo_number || "";
  const fileName = ("Quote_" + quoteNumber + "_" + lang + "_" + woNumber + ".pdf")
    .replace(/[\\/:*?"<>|]/g, "-");

  const docCopy = templateFile.makeCopy("TMP_" + fileName.replace(".pdf", ""), targetFolder);
  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  let map = buildPlaceholderMapFromQuote_(data);

  if (lang === "EN") {
    map = translateQuoteMapToEnglish_(map);
  }

  Object.keys(map).forEach(function(key) {
    body.replaceText("\\{\\{" + escapeRegex_(key) + "\\}\\}", String(map[key] ?? ""));
  });

  doc.saveAndClose();

  Utilities.sleep(3000);

  const blob = docCopy.getAs(MimeType.PDF).setName(fileName);
  const pdfFile = targetFolder.createFile(blob);
  hardenGeneratedPdfFile_(pdfFile);

  docCopy.setTrashed(true);

  return pdfFile.getUrl();
}

function buildQuoteItemsText_(items) {
  if (!items || !items.length) return "";

  return items.map(function(item, i) {
    return (
      (i + 1) + ". " +
      "QTY: " + (item.qty || "") + " | " +
      "PART #: " + (item.part || "") + " | " +
      "DESC: " + (item.desc || "") + " | " +
      "UNIT: " + money_(item.unit) + " | " +
      "LABOR HOURS: " + (item.labor || "") + " | " +
      "RATE: " + (item.rate || "") + " | " +
      "AMOUNT: " + money_(item.amount)
    );
  }).join("\n");
}

function money_(value) {
  const n = Number(value || 0);
  return "$" + n.toFixed(2);
}

function escapeRegExp_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function updateWorkOrderQuoteLinks_(rowNumber, pdfEn, pdfEs) {
  if (!rowNumber) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe WORK_ORDERS");

  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  ensureColumnExists_(sh, headers, "QUOTE_EN_URL");
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  ensureColumnExists_(sh, headers, "QUOTE_ES_URL");
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  ensureColumnExists_(sh, headers, "QUOTE_STATUS");
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  setCellByHeader_(sh, rowNumber, headers, "QUOTE_EN_URL", pdfEn);
  setCellByHeader_(sh, rowNumber, headers, "QUOTE_ES_URL", pdfEs);
  setCellByHeader_(sh, rowNumber, headers, "QUOTE_STATUS", "CREATED");
}

function approveQuote(rowNumber, sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe WORK_ORDERS");

  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  ensureColumnExists_(sh, headers, "QUOTE_STATUS");
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  const woNumber = getCellByHeader_(sh, rowNumber, headers, "WO_NUMBER");
  const companyId = getCellByHeader_(sh, rowNumber, headers, "COMPANY_ID");
  const client = getCellByHeader_(sh, rowNumber, headers, "CLIENT");
  const nsn = getCellByHeader_(sh, rowNumber, headers, "NSN");
  const equipment = getCellByHeader_(sh, rowNumber, headers, "REPORTED_EQUIPMENT");
  const quoteEnUrl = getCellByHeader_(sh, rowNumber, headers, "QUOTE_EN_URL");
  const quoteEsUrl = getCellByHeader_(sh, rowNumber, headers, "QUOTE_ES_URL");

  const session = requireSession_(sessionToken, ["SUPERVISOR", "OWNER", "ADMIN"], companyId);
  if (String(session.role || "").toUpperCase() === "SUPERVISOR") {
    const allowedStores = getSupervisorAllowedStores_(session.name, companyId);
    if (allowedStores.indexOf(normalizeNSN_(nsn)) === -1) {
      throw new Error("No autorizado para aprobar este quote.");
    }
  }

  const approvedBy = getSessionActorLabel_(session);

  setCellByHeader_(sh, rowNumber, headers, "QUOTE_STATUS", "APPROVED");

  updateQuoteSheetApproval_(woNumber);

  addNotification_(companyId, "ADMIN", woNumber, "QUOTE APPROVED", "✅ Quote Approved " + woNumber);
  addNotification_(companyId, "SYSTEM", woNumber, "QUOTE APPROVED", "✅ Quote Approved " + woNumber);
  addNotification_(companyId, "ORDENES", woNumber, "QUOTE APPROVED", "✅ Quote Approved " + woNumber);

  sendQuoteApprovedEmail_({
    companyId: companyId,
    woNumber: woNumber,
    client: client,
    nsn: nsn,
    equipment: equipment,
    approvedBy: approvedBy,
    quoteEnUrl: quoteEnUrl,
    quoteEsUrl: quoteEsUrl
  });

  addAuditLog_("SUPERVISOR", "QUOTE_APPROVED", companyId, "WORK_ORDER", woNumber, session, {
    rowNumber: rowNumber,
    nsn: nsn
  });

  return true;
}

function updateQuoteSheetApproval_(woNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("QUOTES");
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  let headers = data[0].map(String);

  ensureColumnExists_(sh, headers, "APPROVED_BY");
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  ensureColumnExists_(sh, headers, "APPROVED_AT");
  headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  const idxWO = headers.indexOf("WO_NUMBER");
  const idxStatus = headers.indexOf("STATUS");
  const idxBy = headers.indexOf("APPROVED_BY");
  const idxAt = headers.indexOf("APPROVED_AT");

  if (idxWO === -1) return;

  for (let i = sh.getLastRow(); i >= 2; i--) {
    const rowWO = String(sh.getRange(i, idxWO + 1).getValue() || "").trim();

    if (rowWO === String(woNumber || "").trim()) {
      if (idxStatus >= 0) sh.getRange(i, idxStatus + 1).setValue("APPROVED");
      if (idxBy >= 0) sh.getRange(i, idxBy + 1).setValue(Session.getActiveUser().getEmail() || "Supervisor");
      if (idxAt >= 0) sh.getRange(i, idxAt + 1).setValue(new Date());
      break;
    }
  }
}

function sendQuoteApprovedEmail_(info) {
  const to = "invoice.palaciospowersystem@gmail.com";

  MailApp.sendEmail({
    to: to,
    subject: "✅ Quote Approved - " + (info.woNumber || ""),
    htmlBody:
      "<h2>✅ Quote Approved</h2>" +
      "<p>A quote has been approved and is ready for invoice processing.</p>" +
      "<hr>" +
      "<p><b>Company ID:</b> " + (info.companyId || "") + "</p>" +
      "<p><b>Work Order:</b> " + (info.woNumber || "") + "</p>" +
      "<p><b>Customer:</b> " + (info.client || "") + "</p>" +
      "<p><b>NSN:</b> " + (info.nsn || "") + "</p>" +
      "<p><b>Equipment:</b> " + (info.equipment || "") + "</p>" +
      "<p><b>Approved By:</b> " + (info.approvedBy || "") + "</p>" +
      "<p><b>Approved Date:</b> " +
      Utilities.formatDate(new Date(), CFG.TIMEZONE, "MM/dd/yyyy hh:mm a") +
      "</p>" +
      "<hr>" +
      "<p><b>Quote EN:</b> " +
      (info.quoteEnUrl ? "<a href='" + info.quoteEnUrl + "'>Open English Quote</a>" : "No file") +
      "</p>" +
      "<p><b>Quote ES:</b> " +
      (info.quoteEsUrl ? "<a href='" + info.quoteEsUrl + "'>Open Spanish Quote</a>" : "No file") +
      "</p>"
  });
}

function addNotification_(companyId, role, ref, type, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("NOTIFICATIONS");

  if (!sh) {
    sh = ss.insertSheet("NOTIFICATIONS");
    sh.appendRow(["TIMESTAMP", "COMPANY_ID", "ROLE", "REF", "TYPE", "MESSAGE", "READ"]);
  }

  sh.appendRow([
    new Date(),
    companyId || "",
    role || "",
    ref || "",
    type || "",
    message || "",
    false
  ]);
}
function ensureColumnExists_(sheet, headers, columnName) {
  const target = String(columnName || "").trim();

  let normalizedHeaders = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  });

  if (normalizedHeaders.indexOf(target.toUpperCase()) !== -1) {
    return;
  }

  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue(target);
}

function getCellByHeader_(sheet, rowNumber, headers, headerName) {
  const idx = headers.map(String).map(h => h.trim().toUpperCase())
    .indexOf(String(headerName).trim().toUpperCase());

  if (idx === -1) return "";
  return sheet.getRange(Number(rowNumber), idx + 1).getValue();
}

function setCellByHeader_(sheet, rowNumber, headers, headerName, value) {
  const idx = headers.map(String).map(h => h.trim().toUpperCase())
    .indexOf(String(headerName).trim().toUpperCase());

  if (idx === -1) throw new Error("No existe la columna " + headerName);

  sheet.getRange(Number(rowNumber), idx + 1).setValue(value);
}

function buildPlaceholderMapFromQuote_(data) {
  const items = data.items || [];

  const i1 = items[0] || {};
  const i2 = items[1] || {};
  const i3 = items[2] || {};
  const i4 = items[3] || {};

  const tm = splitMoneyParts_(data.total_material);
  const tl = splitMoneyParts_(data.total_labor);
  const st = splitMoneyParts_(data.sub_total);
  const tx = splitMoneyParts_(data.tax);
  const gt = splitMoneyParts_(data.total);

  return {
    quote_number: data.quote_number || "",
    wo_number: data.wo_number || "",
    quote_date: data.quote_date || "",

    nombre_cliente: data.nombre_cliente || "",
    ns_number: data.ns_number || "",

    direccion: data.direccion || "",
    ciudad: data.ciudad || "",
    estado: data.estado || "",
    zip_code: data.zip_code || "",

    marca: data.marca || "",
    modelo: data.modelo || "",
    numero_serie: data.numero_serie || "",

    servicio_solicitado: data.servicio_solicitado || "",
    trabajo_propuesto: data.trabajo_propuesto || "",

    i1_qty: i1.qty || "",
    i1_part: i1.part || "",
    i1_desc: i1.desc || "",
    i1_unit: i1.unit || "",
    i1_labor: i1.labor || "",
    i1_rate: i1.rate || "",
    i1_amount: i1.amount || "",

    i2_qty: i2.qty || "",
    i2_part: i2.part || "",
    i2_desc: i2.desc || "",
    i2_unit: i2.unit || "",
    i2_labor: i2.labor || "",
    i2_rate: i2.rate || "",
    i2_amount: i2.amount || "",

    i3_qty: i3.qty || "",
    i3_part: i3.part || "",
    i3_desc: i3.desc || "",
    i3_unit: i3.unit || "",
    i3_labor: i3.labor || "",
    i3_rate: i3.rate || "",
    i3_amount: i3.amount || "",

    i4_qty: i4.qty || "",
    i4_part: i4.part || "",
    i4_desc: i4.desc || "",
    i4_unit: i4.unit || "",
    i4_labor: i4.labor || "",
    i4_rate: i4.rate || "",
    i4_amount: i4.amount || "",

    total_material_d: tm.d,
    total_material_c: tm.c,

    total_labor_d: tl.d,
    total_labor_c: tl.c,

    sub_total_d: st.d,
    sub_total_c: st.c,

    tax_d: tx.d,
    tax_c: tx.c,

    total_d: gt.d,
    total_c: gt.c
  };
}

function translateQuoteMapToEnglish_(map) {
  const m = Object.assign({}, map);

  m.servicio_solicitado = safeTranslateToEn_(m.servicio_solicitado);
  m.trabajo_propuesto = safeTranslateToEn_(m.trabajo_propuesto);

  m.i1_desc = safeTranslateToEn_(m.i1_desc);
  m.i2_desc = safeTranslateToEn_(m.i2_desc);
  m.i3_desc = safeTranslateToEn_(m.i3_desc);
  m.i4_desc = safeTranslateToEn_(m.i4_desc);

  return m;
}

function safeTranslateToEn_(text) {
  if (!text) return "";

  const t = String(text).trim();
  if (!t) return "";

  try {
    const out = LanguageApp.translate(t, "es", "en");
    return out || t;
  } catch (e) {
    return t;
  }
}

function splitMoneyParts_(amount) {
  const rounded = Math.round(Number(amount || 0) * 100);
  const dollars = Math.floor(Math.abs(rounded) / 100);
  const cents = Math.abs(rounded) % 100;

  return {
    d: String(dollars),
    c: String(cents).padStart(2, "0")
  };
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function safeFolderName_(s) {
  return String(s || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "-") || "SIN_NOMBRE";
}
