const LABOR_RATE_1_MCD = 200;
const LABOR_RATE_2_MCD = 130;

function generatePdfFromCloseOrder_(invoiceRow) {
  const invoicesRoot = DriveApp.getFolderById(CFG.INVOICES_FOLDER_ID);

  const clientName = invoiceRow.CLIENTE || "SIN_CLIENTE";
  const nsNumber = invoiceRow.NS || "SIN_NS";
  const submittedAt = invoiceRow.Timestamp || new Date();

  const monthFolderName = formatMonthFolder_(submittedAt);
  const clientFolder = getOrCreateFolder_(invoicesRoot, safeFolderName_(clientName));
  const monthFolder = getOrCreateFolder_(clientFolder, monthFolderName);
  const nsFolder = getOrCreateFolder_(monthFolder, safeFolderName_(nsNumber));

  const pdfES = generateOnePdfFromTemplate_(invoiceRow, nsFolder, "ES", false);
  const pdfEN = generateOnePdfFromTemplate_(invoiceRow, nsFolder, "EN", true);

  return {
    PDF_ES_URL: pdfES,
    PDF_EN_URL: pdfEN
  };
}

function regenerateInvoiceEnglishPdf_(invoiceRow) {
  const invoicesRoot = DriveApp.getFolderById(CFG.INVOICES_FOLDER_ID);
  const normalizedRow = normalizeInvoiceRowForPdfRepair_(invoiceRow);

  const clientName = normalizedRow.CLIENTE || "SIN_CLIENTE";
  const nsNumber = normalizedRow.NS || "SIN_NS";
  const submittedAt = normalizedRow.Timestamp || new Date();

  const monthFolderName = formatMonthFolder_(submittedAt);
  const clientFolder = getOrCreateFolder_(invoicesRoot, safeFolderName_(clientName));
  const monthFolder = getOrCreateFolder_(clientFolder, monthFolderName);
  const nsFolder = getOrCreateFolder_(monthFolder, safeFolderName_(nsNumber));

  return generateOnePdfFromTemplate_(normalizedRow, nsFolder, "EN", true);
}

function generateOnePdfFromTemplate_(invoiceRow, targetFolder, langTag, doTranslate) {
  const templateFile = DriveApp.getFileById(CFG.TEMPLATE_DOC_ID);

  const invoiceNo = invoiceRow.Invoice || "NO-INVOICE";
  const clientName = invoiceRow.CLIENTE || "Cliente";
  const nsNumber = invoiceRow.NS || "NS";
  const submittedAt = invoiceRow.Timestamp || new Date();

  const stamp = Utilities.formatDate(new Date(submittedAt), CFG.TIMEZONE, "yyyyMMdd-HHmm");

  const pdfName = ("Invoice_" + invoiceNo + "_" + langTag + "_" + clientName + "_NS" + nsNumber + "_" + stamp + ".pdf")
    .replace(/[\\/:*?"<>|]/g, "-");

  const docCopy = templateFile.makeCopy("TMP_" + pdfName.replace(".pdf", ""), targetFolder);
  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  let map = buildPlaceholderMapFromInvoice_(invoiceRow);

  if (doTranslate) {
    map = translateMapToEnglish_(map);
  }

  Object.keys(map).forEach(function(key) {
    body.replaceText("\\{\\{" + escapeRegex_(key) + "\\}\\}", String(map[key] ?? ""));
  });

  if (doTranslate) {
    replaceStaticInvoiceLabelsToEnglish_(body);
  }

  doc.saveAndClose();

  Utilities.sleep(3000);

  let pdfFile = null;
  let lastErr = "";

  for (let i = 0; i < 5; i++) {
    try {
      const blob = docCopy.getAs(MimeType.PDF).setName(pdfName);
      pdfFile = targetFolder.createFile(blob);
      hardenGeneratedPdfFile_(pdfFile);
      break;
    } catch (err) {
      lastErr = err;
      Utilities.sleep(3000 * (i + 1));
    }
  }

  if (!pdfFile) {
    throw new Error("No se pudo crear el PDF. Último error: " + lastErr);
  }

  docCopy.setTrashed(true);

  return pdfFile.getUrl();
}

function repairExistingInvoiceEnglishPdfs(companyId, invoiceNumbers, sessionToken, options) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  return repairExistingInvoiceEnglishPdfs_(companyId, invoiceNumbers, session, options);
}

function repairLatestPpsInvoiceEnglishPdfsFromEditor() {
  throw new Error(
    "Esta reparacion masiva quedo desactivada porque puede crear PDFs incompletos si la fila INVOICES no contiene todos los datos del template. " +
    "Use restoreInvoiceEnglishPdfUrlsFromPreviousFromEditor() para regresar a los PDFs anteriores."
  );
}

function restoreInvoiceEnglishPdfUrlsFromPreviousFromEditor() {
  return restoreInvoiceEnglishPdfUrlsFromPrevious_("PPS", {
    limit: 200,
    actor: {
      email: Session.getActiveUser().getEmail() || "SCRIPT_EDITOR",
      name: "SCRIPT_EDITOR",
      role: "OWNER"
    }
  });
}

function restoreInvoiceEnglishPdfUrlsFromPrevious(companyId, sessionToken, options) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"], companyId);
  options = options || {};
  options.actor = session;
  return restoreInvoiceEnglishPdfUrlsFromPrevious_(companyId, options);
}

function restoreInvoiceEnglishPdfUrlsFromPrevious_(companyId, options) {
  options = options || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("INVOICES");
  if (!sh) throw new Error("No existe la hoja INVOICES.");

  companyId = String(companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  const actor = options.actor || {};
  const limit = Math.max(1, Number(options.limit || 200));

  let headers = ensureSheetColumns_(sh, [
    "PDF_EN_URL",
    "PDF_EN_PREVIOUS_URL",
    "PDF_EN_REPAIR_BAD_URL",
    "PDF_EN_RESTORED_AT",
    "PDF_EN_RESTORED_BY"
  ]);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return {
      ok: true,
      companyId: companyId,
      restored: [],
      skipped: []
    };
  }

  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const restored = [];
  const skipped = [];

  for (let i = data.length - 1; i >= 0; i--) {
    if (restored.length >= limit) break;

    const sheetRow = i + 2;
    const row = data[i];

    if (isMonthCloseMarkerRow_(row, headers) || isSoftDeletedRow_(row, headers)) continue;

    const invoiceObj = rowToObject_(headers, row);
    const rowCompany = String(invoiceObj.COMPANY_ID || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
    if (companyId && rowCompany !== companyId) continue;

    const currentUrl = String(invoiceObj.PDF_EN_URL || "").trim();
    const previousUrl = String(invoiceObj.PDF_EN_PREVIOUS_URL || "").trim();
    const invoiceNumber = normalizeInvoiceRepairNumber_(
      invoiceObj.INVOICE_NUMBER || invoiceObj.Invoice || invoiceObj.INVOICE
    );

    if (!previousUrl) {
      skipped.push({ row: sheetRow, invoiceNumber: invoiceNumber, reason: "NO_PREVIOUS_URL" });
      continue;
    }

    if (currentUrl === previousUrl) {
      skipped.push({ row: sheetRow, invoiceNumber: invoiceNumber, reason: "ALREADY_RESTORED" });
      continue;
    }

    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) {
      return String(h || "").trim();
    });

    setCellByHeader_(sh, sheetRow, headers, "PDF_EN_REPAIR_BAD_URL", currentUrl);
    setCellByHeader_(sh, sheetRow, headers, "PDF_EN_URL", previousUrl);
    setCellByHeader_(sh, sheetRow, headers, "PDF_EN_RESTORED_AT", new Date());
    setCellByHeader_(sh, sheetRow, headers, "PDF_EN_RESTORED_BY", getSessionActorLabel_(actor));

    restored.push({
      row: sheetRow,
      invoiceNumber: invoiceNumber,
      restoredUrl: previousUrl,
      badUrl: currentUrl
    });
  }

  addAuditLog_("INVOICE", "ENGLISH_PDF_URLS_RESTORED", companyId, "INVOICE", "BATCH", actor, {
    limit: limit,
    restoredCount: restored.length,
    skippedCount: skipped.length
  });

  return {
    ok: true,
    companyId: companyId,
    restored: restored,
    skipped: skipped
  };
}

function repairExistingInvoiceEnglishPdfs_(companyId, invoiceNumbers, actor, options) {
  options = options || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("INVOICES");
  if (!sh) throw new Error("No existe la hoja INVOICES.");

  companyId = String(companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  const wantedInvoices = normalizeInvoiceRepairList_(invoiceNumbers);
  const limit = Math.max(1, Number(options.limit || 25));
  const latestFirst = options.latestFirst !== false;

  let headers = ensureSheetColumns_(sh, [
    "PDF_EN_URL",
    "PDF_EN_REPAIRED_AT",
    "PDF_EN_REPAIRED_BY",
    "PDF_EN_PREVIOUS_URL"
  ]);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return {
      ok: true,
      companyId: companyId,
      processed: 0,
      updated: [],
      skipped: [],
      errors: [],
      remainingEstimate: 0
    };
  }

  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const rowIndexes = data.map(function(_, i) { return i; });
  if (latestFirst) rowIndexes.reverse();

  const updated = [];
  const skipped = [];
  const errors = [];
  let remainingEstimate = 0;
  let attempted = 0;

  for (let n = 0; n < rowIndexes.length; n++) {
    const i = rowIndexes[n];
    const sheetRow = i + 2;
    const row = data[i];

    if (isMonthCloseMarkerRow_(row, headers) || isSoftDeletedRow_(row, headers)) continue;

    const invoiceObj = rowToObject_(headers, row);
    const rowCompany = String(invoiceObj.COMPANY_ID || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
    if (companyId && rowCompany !== companyId) continue;

    const invoiceNumber = normalizeInvoiceRepairNumber_(
      invoiceObj.INVOICE_NUMBER || invoiceObj.Invoice || invoiceObj.INVOICE
    );

    if (!invoiceNumber) {
      skipped.push({ row: sheetRow, reason: "NO_INVOICE_NUMBER" });
      continue;
    }

    if (wantedInvoices.length && wantedInvoices.indexOf(invoiceNumber) === -1) continue;

    if (attempted >= limit) {
      remainingEstimate++;
      continue;
    }

    attempted++;

    try {
      const oldUrl = String(invoiceObj.PDF_EN_URL || "").trim();
      const newUrl = regenerateInvoiceEnglishPdf_(invoiceObj);
      validateGeneratedInvoicePdfUrl_(newUrl, invoiceNumber);

      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) {
        return String(h || "").trim();
      });

      setCellByHeader_(sh, sheetRow, headers, "PDF_EN_URL", newUrl);
      setCellByHeader_(sh, sheetRow, headers, "PDF_EN_REPAIRED_AT", new Date());
      setCellByHeader_(sh, sheetRow, headers, "PDF_EN_REPAIRED_BY", getSessionActorLabel_(actor));
      setCellByHeader_(sh, sheetRow, headers, "PDF_EN_PREVIOUS_URL", oldUrl);

      updated.push({
        row: sheetRow,
        invoiceNumber: invoiceNumber,
        oldUrl: oldUrl,
        newUrl: newUrl
      });
    } catch (err) {
      errors.push({
        row: sheetRow,
        invoiceNumber: invoiceNumber,
        error: err && err.message ? err.message : String(err)
      });
    }
  }

  addAuditLog_("INVOICE", "ENGLISH_PDFS_REPAIRED", companyId, "INVOICE", "BATCH", actor, {
    requestedInvoices: wantedInvoices,
    limit: limit,
    updatedCount: updated.length,
    errorCount: errors.length,
    remainingEstimate: remainingEstimate
  });

  return {
    ok: errors.length === 0,
    companyId: companyId,
    processed: attempted,
    updated: updated,
    skipped: skipped,
    errors: errors,
    remainingEstimate: remainingEstimate
  };
}

function normalizeInvoiceRowForPdfRepair_(invoiceRow) {
  const r = Object.assign({}, invoiceRow || {});

  r.INVOICE_NUMBER = r.INVOICE_NUMBER || r.Invoice || r.INVOICE || "";
  r.Invoice = r.Invoice || r.INVOICE_NUMBER || "";
  r.Timestamp = r.Timestamp || r.DATE_INVOICE || r.INVOICE_DATE || new Date();

  r.CLIENTE = r.CLIENTE || r.CLIENT || r.CUSTOMER || CFG.CLIENT_DEFAULT || "";
  r.NS = r.NS || r.NSN || r.STORE_NUMBER || "";

  r.PARTS = r.PARTS || r.MATERIAL_COST || r.COMPRA || r.INVERSION || 0;
  r.LABOR = r.LABOR || r.LABOR_AMOUNT || 0;
  r.TAX = r.TAX || r.TAX_AMOUNT || 0;
  r.TOTAL = r.TOTAL || r.INVOICE_TOTAL || r.GRAND_TOTAL || 0;
  r.GRAND_TOTAL = r.GRAND_TOTAL || r.TOTAL || 0;
  r.SUB_TOTAL = r.SUB_TOTAL || Math.max(0, parseMoney_(r.TOTAL) - parseMoney_(r.TAX));

  r.EQUIPMENT_MAKE = r.EQUIPMENT_MAKE || r.MAKE || "";
  r.EQUIPMENT_MODEL = r.EQUIPMENT_MODEL || r.MODEL || "";
  r.EQUIPMENT_SERIAL = r.EQUIPMENT_SERIAL || r.SERIAL_NUMBER || r.SERIAL || "";

  return r;
}

function normalizeInvoiceRepairList_(invoiceNumbers) {
  if (!invoiceNumbers) return [];

  if (typeof invoiceNumbers === "string") {
    invoiceNumbers = invoiceNumbers.split(/[,\n;]+/);
  }

  if (!Array.isArray(invoiceNumbers)) return [];

  return invoiceNumbers.map(normalizeInvoiceRepairNumber_).filter(Boolean);
}

function normalizeInvoiceRepairNumber_(value) {
  return String(value || "").trim().replace(/\.0$/, "").toUpperCase();
}

function validateGeneratedInvoicePdfUrl_(pdfUrl, invoiceNumber) {
  const fileId = extractInvoicePdfFileId_(pdfUrl);
  if (!fileId) throw new Error("No se pudo validar el PDF generado para invoice " + invoiceNumber + ".");

  const file = DriveApp.getFileById(fileId);
  const size = Number(file.getSize() || 0);

  if (size < 12000) {
    try {
      file.setTrashed(true);
    } catch (trashErr) {
      Logger.log("No se pudo enviar a trash PDF sospechoso: " + trashErr);
    }

    throw new Error(
      "PDF generado parece estar vacio o incompleto para invoice " + invoiceNumber +
      " (" + size + " bytes). Se conserva el PDF anterior."
    );
  }

  return true;
}

function extractInvoicePdfFileId_(urlOrId) {
  const s = String(urlOrId || "").trim();
  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;

  let m = s.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];

  m = s.match(/\/file\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];

  return "";
}

function buildPlaceholderMapFromInvoice_(r) {
  const totalMat = parseMoney_(r.PARTS);
  const totalLab = parseMoney_(r.LABOR);
  const subTotal = parseMoney_(r.SUB_TOTAL || (totalMat + totalLab));
  const taxAmt = parseMoney_(r.TAX);
  const grandTot = parseMoney_(r.TOTAL || r.GRAND_TOTAL);

  const tm = splitMoneyParts_(totalMat);
  const tl = splitMoneyParts_(totalLab);
  const st = splitMoneyParts_(subTotal);
  const tx = splitMoneyParts_(taxAmt);
  const gt = splitMoneyParts_(grandTot);

  const r1 = splitMoneyParts_(r.LABOR_1_RATE);
  const r2 = splitMoneyParts_(r.LABOR_2_RATE);
  const a1 = splitMoneyParts_(r.LABOR_1_AMOUNT);
  const a2 = splitMoneyParts_(r.LABOR_2_AMOUNT);

  const address = normalizeInvoiceAddressForTemplate_(r);

  return {
    invoice_number: r.Invoice || "",
    wo_number: r.WO_NUMBER || "",
    invoice_date: Utilities.formatDate(new Date(r.Timestamp || new Date()), CFG.TIMEZONE, "MM/dd/yyyy"),

    nombre_cliente: r.CLIENTE || "",
    vendor_id: r.VENDOR_ID || "",
    ns_number: r.NS || "",

    direccion: address.street,
    ciudad: address.city,
    estado: address.state,
    zip_code: address.zip,

    marca: r.EQUIPMENT_MAKE || "",
    modelo: r.EQUIPMENT_MODEL || "",
    numero_serie: r.EQUIPMENT_SERIAL || "",

    problema_reportado: r.REPORTED_PROBLEM || "",
    trabajo_realizado: r.WORK_PERFORMED || r.TRABAJO_REALIZADO || "",

    i1_qty: r.I1_QTY || "",
    i1_part: r.I1_PART || "",
    i1_desc: r.I1_DESC || "",
    i1_unit: r.I1_UNIT || "",
    i1_labor: r.LABOR_1_QTY || "",
    i1_rate_d: r1.d,
    i1_rate_c: r1.c,
    i1_amount_d: a1.d,
    i1_amount_c: a1.c,

    i2_qty: r.I2_QTY || "",
    i2_part: r.I2_PART || "",
    i2_desc: r.I2_DESC || "",
    i2_unit: r.I2_UNIT || "",
    i2_labor: r.LABOR_2_QTY || "",
    i2_rate_d: r2.d,
    i2_rate_c: r2.c,
    i2_amount_d: a2.d,
    i2_amount_c: a2.c,

    i3_qty: r.I3_QTY || "",
    i3_part: r.I3_PART || "",
    i3_desc: r.I3_DESC || "",
    i3_unit: r.I3_UNIT || "",
    i3_labor: "",
    i3_rate: "",
    i3_amount: r.I3_AMOUNT || "",

    i4_qty: r.I4_QTY || "",
    i4_part: r.I4_PART || "",
    i4_desc: r.I4_DESC || "",
    i4_unit: r.I4_UNIT || "",
    i4_labor: "",
    i4_rate: "",
    i4_amount: r.I4_AMOUNT || "",

    total_material_d: tm.d,
    total_material_c: tm.c,

    total_labor_d: tl.d,
    total_labor_c: tl.c,

    sub_total_d: st.d,
    sub_total_c: st.c,

    tax_d: tx.d,
    tax_c: tx.c,

    total_d: gt.d,
    total_c: gt.c,

    firma_cliente: r.SIGNATURE || ""
  };
}

function normalizeInvoiceAddressForTemplate_(r) {
  let street = String(r.STORE_STREET || "").trim();
  let city = String(r.STORE_CITY || r.CITY || "").trim().toUpperCase();
  let state = String(r.STORE_STATE || r.STATE || "").trim().toUpperCase();
  let zip = String(r.STORE_ZIP || r.ZIP || "").trim();
  const full = String(r.STORE_ADDRESS || "").trim();

  if ((!street || !city || !state || !zip) && full) {
    const m = full.match(/^(.*?),\s*([^,]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/i);

    if (m) {
      if (!street) street = m[1].trim();
      if (!city) city = m[2].trim().toUpperCase();
      if (!state) state = m[3].trim().toUpperCase();
      if (!zip) zip = m[4].trim();
    }
  }

  return {
    street: street,
    city: city,
    state: state,
    zip: zip
  };
}

function safeTranslateToEn_(text) {
  if (!text) return "";

  const t = String(text).trim();
  if (!t) return "";
  if (/^\d+(\.\d+)?$/.test(t)) return t;
  if (/^[A-Za-z0-9\-_.\/]+$/.test(t) && t.length <= 25) return t;

  try {
    const out = LanguageApp.translate(t, "es", "en");
    if (!out || out.trim() === "") return t;
    return out;
  } catch (e) {
    return t;
  }
}

function translateMapToEnglish_(map) {
  const m = Object.assign({}, map);

  m.problema_reportado = safeTranslateToEn_(m.problema_reportado);
  m.trabajo_realizado = safeTranslateToEn_(m.trabajo_realizado);

  m.i1_desc = safeTranslateToEn_(m.i1_desc);
  m.i2_desc = safeTranslateToEn_(m.i2_desc);
  m.i3_desc = safeTranslateToEn_(m.i3_desc);
  m.i4_desc = safeTranslateToEn_(m.i4_desc);

  return m;
}

function replaceStaticInvoiceLabelsToEnglish_(body) {
  const replacements = [
    { pattern: "\\bMARCA\\b", value: "MAKE" },
    { pattern: "\\bMarca\\b", value: "Make" },
    { pattern: "\\bMODELO\\b", value: "MODEL" },
    { pattern: "\\bModelo\\b", value: "Model" },
    { pattern: "\\bN.?MERO\\s+DE\\s+SERIE\\b", value: "SERIAL NUMBER" },
    { pattern: "\\bN.?mero\\s+de\\s+Serie\\b", value: "Serial Number" },
    { pattern: "\\bN.?mero\\s+de\\s+serie\\b", value: "Serial Number" },
    { pattern: "\\bNUMERO\\s+DE\\s+SERIE\\b", value: "SERIAL NUMBER" },
    { pattern: "\\bNumero\\s+de\\s+Serie\\b", value: "Serial Number" },
    { pattern: "\\bNumero\\s+de\\s+serie\\b", value: "Serial Number" },
    { pattern: "\\bN.?MERO\\s+SERIE\\b", value: "SERIAL NUMBER" },
    { pattern: "\\bNumero\\s+Serie\\b", value: "Serial Number" },
    { pattern: "\\bN\\.\\s*SERIE\\b", value: "SERIAL NUMBER" },
    { pattern: "\\bNo\\.\\s*Serie\\b", value: "Serial Number" },
    { pattern: "\\bNO\\.\\s*SERIE\\b", value: "SERIAL NUMBER" }
  ];

  replacements.forEach(function(item) {
    body.replaceText(item.pattern, item.value);
  });
}

function getOrCreateFolder_(parentFolder, folderName) {
  const it = parentFolder.getFoldersByName(folderName);
  return it.hasNext() ? it.next() : parentFolder.createFolder(folderName);
}

function formatMonthFolder_(dateObj) {
  const d = dateObj instanceof Date ? dateObj : new Date(dateObj);
  const y = Utilities.formatDate(d, CFG.TIMEZONE, "yyyy");
  const mNum = Utilities.formatDate(d, CFG.TIMEZONE, "MM");
  const mName = monthNameEs_(parseInt(mNum, 10));
  return y + "-" + mNum + "_" + mName;
}

function monthNameEs_(m) {
  const names = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  return names[m - 1] || "Mes";
}

function safeFolderName_(s) {
  return String(s || "").trim().replace(/[\\/:*?"<>|]/g, "-") || "SIN_NOMBRE";
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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

function parseMoney_(v) {
  if (v === null || v === undefined) return 0;
  const s = String(v).trim();
  if (!s) return 0;
  const cleaned = s.replace(/[^0-9.-]/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function testGeneratePdfFromCloseOrder(sessionToken) {
  requireSession_(sessionToken, ["OWNER", "ADMIN"]);
  const testInvoice = {
    Invoice: "TEST-1001",
    WO_NUMBER: "WO-TEST",
    Timestamp: new Date(),

    CLIENTE: "PALACIOS POWER SYSTEMS CORP",
    NS: "1155",
    VENDOR_ID: "TEST-VENDOR",

    STORE_ADDRESS: "123 Test St, Miami, FL 33055",
    STORE_STREET: "123 Test St",
    STORE_CITY: "MIAMI",
    STORE_STATE: "FL",
    STORE_ZIP: "33055",

    EQUIPMENT_MAKE: "Test Make",
    EQUIPMENT_MODEL: "Test Model",
    EQUIPMENT_SERIAL: "12345",

    REPORTED_PROBLEM: "Equipo no enfría",
    PROCESO: "Se revisó el sistema y se reparó correctamente",

    I1_QTY: 1,
    I1_PART: "TEST PART",
    I1_DESC: "Motor de prueba",
    I1_UNIT: 50,
    I1_AMOUNT: 50,

    I2_QTY: "",
    I2_PART: "",
    I2_DESC: "",
    I2_UNIT: "",
    I2_AMOUNT: "",

    I3_QTY: "",
    I3_PART: "",
    I3_DESC: "",
    I3_UNIT: "",
    I3_AMOUNT: "",

    I4_QTY: "",
    I4_PART: "",
    I4_DESC: "",
    I4_UNIT: "",
    I4_AMOUNT: "",

    PARTS: 50,
    LABOR: 200,

    LABOR_1_QTY: 1,
    LABOR_1_RATE: 200,
    LABOR_1_AMOUNT: 200,

    LABOR_2_QTY: 0,
    LABOR_2_RATE: 130,
    LABOR_2_AMOUNT: 0,

    SUB_TOTAL: 250,
    TAX: 17.5,
    GRAND_TOTAL: 267.5,
    TOTAL: 267.5,

    SIGNATURE: "Cliente Test"
  };

  const result = generatePdfFromCloseOrder_(testInvoice);
  Logger.log(JSON.stringify(result));
}
