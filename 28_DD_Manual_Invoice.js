// =====================================================
// FILE: 28_DD_Manual_Invoice.gs
// Manual D&D Premier invoice generator for OWNER users.
// =====================================================

const DD_MANUAL_INVOICE_COLUMNS = [
  "COMPANY_ID",
  "INVOICE_BRAND_COMPANY_ID",
  "INVOICE_TYPE",
  "SPECIAL_INVOICE",
  "INVOICE_NUMBER",
  "Invoice",
  "INVOICE",
  "WO_NUMBER",
  "DATE_INVOICE",
  "INVOICE_DATE",
  "Timestamp",
  "CLIENTE",
  "CUSTOMER_EMAIL",
  "VENDOR_ID",
  "NS",
  "STORE_ADDRESS",
  "STORE_STREET",
  "STORE_CITY",
  "STORE_STATE",
  "STORE_ZIP",
  "REPORTED_PROBLEM",
  "WORK_PERFORMED",
  "DESCRIPTION",
  "MATERIALS_USED",
  "MATERIALS_LIST",
  "MATERIALS_JSON",
  "NOTES",
  "PARTS",
  "MATERIAL_COST",
  "LABOR",
  "SUB_TOTAL",
  "TAX",
  "TOTAL",
  "GRAND_TOTAL",
  "INVOICE_TOTAL",
  "BILLING_COMPANY_NAME",
  "BILLING_COMPANY_ADDRESS",
  "BILLING_COMPANY_PHONE",
  "BILLING_COMPANY_EMAIL",
  "BILLING_COMPANY_WEBSITE",
  "BILLING_COMPANY_LOGO_FILE_ID",
  "PDF_ES_URL",
  "PDF_EN_URL",
  "CLIENT_EMAIL_SENT",
  "CREATED_BY_EMAIL",
  "CREATED_BY_NAME"
];

function createDDPremierManualInvoice(data, sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER"]);
  data = data || {};

  const invoiceDate = parseDDManualInvoiceDate_(data.invoiceDate) || new Date();
  const partsTotal = roundDDManualInvoiceMoney_(data.partsTotal);
  const laborTotal = roundDDManualInvoiceMoney_(data.laborTotal);
  const subtotal = hasDDManualValue_(data.subtotal)
    ? roundDDManualInvoiceMoney_(data.subtotal)
    : roundDDManualInvoiceMoney_(partsTotal + laborTotal);
  const taxAmount = hasDDManualValue_(data.taxAmount)
    ? roundDDManualInvoiceMoney_(data.taxAmount)
    : roundDDManualInvoiceMoney_(subtotal * (Number(data.taxRate || 0) / 100));
  const total = hasDDManualValue_(data.total)
    ? roundDDManualInvoiceMoney_(data.total)
    : roundDDManualInvoiceMoney_(subtotal + taxAmount);

  const customerName = String(data.customerName || "").trim();
  if (!customerName) {
    throw new Error("Debe indicar el cliente del invoice manual.");
  }

  const billing = getDDManualInvoiceCompanyDefaults_(data);
  const materialItems = normalizeDDManualMaterialItems_(data.materialsItems, data.materialsUsed);
  const materialsText = formatDDManualMaterialItemsText_(materialItems, data.materialsUsed);
  const invoiceNumber = generateInvoiceNumber_("DD");
  const invoiceNumberText = String(invoiceNumber || "").trim();
  const now = new Date();

  const row = {
    COMPANY_ID: "DD",
    INVOICE_BRAND_COMPANY_ID: "DD",
    INVOICE_TYPE: "DD_MANUAL",
    SPECIAL_INVOICE: "YES",
    INVOICE_NUMBER: invoiceNumberText,
    Invoice: invoiceNumberText,
    INVOICE: invoiceNumberText,
    WO_NUMBER: String(data.woNumber || "DD-MANUAL-" + invoiceNumberText).trim(),
    DATE_INVOICE: invoiceDate,
    INVOICE_DATE: invoiceDate,
    Timestamp: now,
    CLIENTE: customerName,
    CUSTOMER_EMAIL: String(data.customerEmail || "").trim(),
    VENDOR_ID: String(data.vendorId || "").trim(),
    NS: String(data.nsn || "").trim(),
    STORE_ADDRESS: buildDDManualCustomerAddress_(data),
    STORE_STREET: String(data.customerStreet || "").trim(),
    STORE_CITY: String(data.customerCity || "").trim(),
    STORE_STATE: String(data.customerState || "").trim(),
    STORE_ZIP: String(data.customerZip || "").trim(),
    REPORTED_PROBLEM: String(data.reportedProblem || data.problemReported || "").trim(),
    WORK_PERFORMED: String(data.workPerformed || data.description || "").trim(),
    DESCRIPTION: String(data.description || "").trim(),
    MATERIALS_USED: materialsText,
    MATERIALS_LIST: materialsText,
    MATERIALS_JSON: JSON.stringify(materialItems),
    NOTES: String(data.notes || "").trim(),
    PARTS: partsTotal,
    MATERIAL_COST: partsTotal,
    LABOR: laborTotal,
    SUB_TOTAL: subtotal,
    TAX: taxAmount,
    TOTAL: total,
    GRAND_TOTAL: total,
    INVOICE_TOTAL: total,
    BILLING_COMPANY_NAME: billing.name,
    BILLING_COMPANY_ADDRESS: billing.address,
    BILLING_COMPANY_PHONE: billing.phone,
    BILLING_COMPANY_EMAIL: billing.email,
    BILLING_COMPANY_WEBSITE: billing.website,
    BILLING_COMPANY_LOGO_FILE_ID: billing.logoFileId,
    PDF_ES_URL: "",
    PDF_EN_URL: "",
    CLIENT_EMAIL_SENT: "NO",
    CREATED_BY_EMAIL: String(session.email || "").trim(),
    CREATED_BY_NAME: String(session.name || "").trim()
  };

  const pdfs = generateDDPremierManualInvoicePdfs_(row);
  row.PDF_ES_URL = pdfs.PDF_ES_URL || "";
  row.PDF_EN_URL = pdfs.PDF_EN_URL || "";

  appendDDManualInvoiceRow_(row);

  addAuditLog_("INVOICE", "DD_MANUAL_INVOICE_CREATED", "DD", "INVOICE", invoiceNumberText, session, {
    customer: customerName,
    total: total,
    pdfEsUrl: row.PDF_ES_URL,
    pdfEnUrl: row.PDF_EN_URL
  });

  return {
    ok: true,
    invoiceNumber: invoiceNumberText,
    total: total,
    pdfEsUrl: row.PDF_ES_URL,
    pdfEnUrl: row.PDF_EN_URL
  };
}

function appendDDManualInvoiceRow_(row) {
  const ss = getDDManualInvoiceSpreadsheet_();
  const sh = ss.getSheetByName("INVOICES") || ss.insertSheet("INVOICES");
  const headers = ensureSheetColumns_(sh, DD_MANUAL_INVOICE_COLUMNS);
  sh.appendRow(headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : "";
  }));
}

function getDDManualInvoiceSpreadsheet_() {
  const id = String(CFG.DD_PREMIER_INVOICE_SPREADSHEET_ID || "").trim();
  if (id) {
    return SpreadsheetApp.openById(id);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

function generateDDPremierManualInvoicePdfs_(row) {
  const rootFolderId = String(CFG.DD_PREMIER_INVOICE_FOLDER_ID || CFG.INVOICES_FOLDER_ID || "").trim();
  if (!rootFolderId) {
    throw new Error("No esta configurada la carpeta para guardar los invoices manuales de D&D.");
  }

  const root = DriveApp.getFolderById(rootFolderId);
  const monthFolder = getOrCreateFolder_(root, formatMonthFolder_(row.DATE_INVOICE || new Date()));
  const invoiceFolder = getOrCreateFolder_(monthFolder, safeFolderName_("Invoice_" + row.Invoice));

  return {
    PDF_ES_URL: generateDDPremierManualInvoicePdf_(row, invoiceFolder, "ES"),
    PDF_EN_URL: generateDDPremierManualInvoicePdf_(row, invoiceFolder, "EN")
  };
}

function generateDDPremierManualInvoicePdf_(row, folder, lang) {
  const invoiceNumber = String(row.Invoice || row.INVOICE_NUMBER || "").trim();
  const isSpanish = String(lang || "").toUpperCase() === "ES";
  const docName = safeFolderName_("DD_Manual_Invoice_" + invoiceNumber + "_" + String(lang || "EN").toUpperCase());
  const doc = DocumentApp.create(docName);
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(folder);

  const body = doc.getBody();
  body.clear();
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(42);
  body.setMarginRight(42);

  appendDDManualHeader_(body, row);
  appendDDManualTitle_(body, isSpanish ? "FACTURA" : "INVOICE");
  appendDDManualInfoTable_(body, row, isSpanish);
  appendDDManualDescription_(body, row, isSpanish);
  appendDDManualTotals_(body, row, isSpanish);
  appendDDManualFooter_(body, row, isSpanish);

  doc.saveAndClose();

  const pdfBlob = docFile.getBlob()
    .getAs(MimeType.PDF)
    .setName(docName + ".pdf");
  const pdfFile = folder.createFile(pdfBlob);
  hardenGeneratedPdfFile_(pdfFile);
  docFile.setTrashed(true);

  return pdfFile.getUrl();
}

function appendDDManualHeader_(body, row) {
  const headerTable = body.appendTable([["", ""]]);
  headerTable.setBorderColor("#ffffff");
  const headerRow = headerTable.getRow(0);
  const leftCell = headerRow.getCell(0);
  const rightCell = headerRow.getCell(1);
  leftCell.setText("");
  rightCell.setText("");

  const logoId = String(row.BILLING_COMPANY_LOGO_FILE_ID || "").trim();
  if (logoId) {
    try {
      const logo = leftCell.appendImage(DriveApp.getFileById(logoId).getBlob());
      const maxWidth = 170;
      const width = logo.getWidth();
      const height = logo.getHeight();
      if (width > maxWidth) {
        logo.setWidth(maxWidth);
        logo.setHeight(Math.round(height * (maxWidth / width)));
      }
    } catch (err) {
      Logger.log("WARN DD manual invoice logo: " + err);
    }
  }

  const name = rightCell.appendParagraph(String(row.BILLING_COMPANY_NAME || "D&D Premier Solutions Corp"));
  name.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  name.editAsText().setBold(true).setFontSize(14).setForegroundColor("#111827");

  const address = String(row.BILLING_COMPANY_ADDRESS || "").trim();
  if (address) {
    const addressParagraph = rightCell.appendParagraph(address);
    addressParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    addressParagraph.editAsText().setFontSize(9).setForegroundColor("#374151");
  }

  const contact = buildDDManualCompanyContactLine_(row);
  if (contact) {
    const contactParagraph = rightCell.appendParagraph(contact);
    contactParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    contactParagraph.editAsText().setFontSize(9).setForegroundColor("#374151");
  }

  body.appendHorizontalRule();
}

function appendDDManualTitle_(body, titleText) {
  const title = body.appendParagraph(titleText);
  title.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  title.editAsText().setBold(true).setFontSize(24).setForegroundColor("#dc2626");
}

function appendDDManualInfoTable_(body, row, isSpanish) {
  const invoiceDate = formatDDManualDate_(row.DATE_INVOICE || new Date());
  const customerAddress = String(row.STORE_ADDRESS || "").trim();
  const data = [
    [isSpanish ? "Facturar a" : "Bill To", "", isSpanish ? "Detalles" : "Details", ""],
    [isSpanish ? "Cliente" : "Customer", String(row.CLIENTE || ""), isSpanish ? "Factura #" : "Invoice #", String(row.Invoice || row.INVOICE_NUMBER || "")],
    [isSpanish ? "Direccion" : "Address", customerAddress, isSpanish ? "Fecha" : "Date", invoiceDate],
    [isSpanish ? "Vendor ID" : "Vendor ID", String(row.VENDOR_ID || ""), isSpanish ? "Referencia / WO" : "Reference / WO", String(row.WO_NUMBER || "")],
    [isSpanish ? "NSN / Tienda" : "NSN / Store", String(row.NS || ""), "", ""]
  ];

  const table = body.appendTable(data);
  table.setBorderColor("#d1d5db");
  for (let r = 0; r < table.getNumRows(); r++) {
    const rowObj = table.getRow(r);
    for (let c = 0; c < rowObj.getNumCells(); c++) {
      rowObj.getCell(c).editAsText().setFontSize(9).setForegroundColor("#111827");
    }

    if (r === 0) {
      rowObj.getCell(0).setBackgroundColor("#111827");
      rowObj.getCell(1).setBackgroundColor("#111827");
      rowObj.getCell(2).setBackgroundColor("#111827");
      rowObj.getCell(3).setBackgroundColor("#111827");
      rowObj.getCell(0).editAsText().setBold(true).setForegroundColor("#ffffff");
      rowObj.getCell(2).editAsText().setBold(true).setForegroundColor("#ffffff");
    } else {
      rowObj.getCell(0).setBackgroundColor("#f3f4f6");
      rowObj.getCell(2).setBackgroundColor("#f3f4f6");
      rowObj.getCell(0).editAsText().setBold(true);
      rowObj.getCell(2).editAsText().setBold(true);
    }
  }
  body.appendParagraph("");
}

function appendDDManualDescription_(body, row, isSpanish) {
  const reportedProblem = String(row.REPORTED_PROBLEM || "").trim();
  if (reportedProblem) {
    const problemHeading = body.appendParagraph(isSpanish ? "Problema reportado" : "Reported Problem");
    problemHeading.editAsText().setBold(true).setFontSize(12).setForegroundColor("#111827");

    body.appendParagraph(translateDDManualTextForLang_(reportedProblem, isSpanish))
      .editAsText()
      .setFontSize(10)
      .setForegroundColor("#374151");

    body.appendParagraph("");
  }

  const heading = body.appendParagraph(isSpanish ? "Descripcion / Trabajo" : "Description / Work Performed");
  heading.editAsText().setBold(true).setFontSize(12).setForegroundColor("#111827");

  const description = String(row.DESCRIPTION || row.WORK_PERFORMED || "").trim();
  body.appendParagraph(description
      ? translateDDManualTextForLang_(description, isSpanish)
      : (isSpanish ? "Servicio segun acuerdo." : "Service per agreement."))
    .editAsText()
    .setFontSize(10)
    .setForegroundColor("#374151");

  const materialItems = getDDManualMaterialItemsFromRow_(row);
  if (materialItems.length) {
    const materialsHeading = body.appendParagraph(isSpanish ? "Materiales usados" : "Materials Used");
    materialsHeading.editAsText().setBold(true).setFontSize(10).setForegroundColor("#111827");

    const tableData = [[isSpanish ? "QTY" : "QTY", isSpanish ? "Descripcion del material" : "Material Description"]].concat(
      materialItems.map(function(item) {
        return [
          String(item.qty || ""),
          translateDDManualTextForLang_(String(item.description || ""), isSpanish)
        ];
      })
    );
    const table = body.appendTable(tableData);
    table.setBorderColor("#e5e7eb");
    for (let r = 0; r < table.getNumRows(); r++) {
      const rowObj = table.getRow(r);
      for (let c = 0; c < rowObj.getNumCells(); c++) {
        rowObj.getCell(c).editAsText().setFontSize(9).setForegroundColor("#374151");
        if (r === 0) {
          rowObj.getCell(c).setBackgroundColor("#f3f4f6");
          rowObj.getCell(c).editAsText().setBold(true).setForegroundColor("#111827");
        }
      }
    }
  }

  const notes = String(row.NOTES || "").trim();
  if (notes) {
    const notesHeading = body.appendParagraph(isSpanish ? "Notas" : "Notes");
    notesHeading.editAsText().setBold(true).setFontSize(10).setForegroundColor("#111827");
    body.appendParagraph(translateDDManualTextForLang_(notes, isSpanish)).editAsText().setFontSize(9).setForegroundColor("#374151");
  }

  body.appendParagraph("");
}

function appendDDManualTotals_(body, row, isSpanish) {
  const chargeRows = [
    [isSpanish ? "Concepto" : "Item", isSpanish ? "Monto" : "Amount"],
    [isSpanish ? "Materiales / Parts" : "Parts / Materials", formatDDManualMoney_(row.PARTS)],
    [isSpanish ? "Labor / Servicios" : "Labor / Services", formatDDManualMoney_(row.LABOR)]
  ];

  const chargeTable = body.appendTable(chargeRows);
  chargeTable.setBorderColor("#d1d5db");
  for (let r = 0; r < chargeTable.getNumRows(); r++) {
    const rowObj = chargeTable.getRow(r);
    rowObj.getCell(0).editAsText().setFontSize(10).setForegroundColor("#111827");
    rowObj.getCell(1).editAsText().setFontSize(10).setForegroundColor("#111827");
    if (r === 0) {
      rowObj.getCell(0).setBackgroundColor("#f3f4f6");
      rowObj.getCell(1).setBackgroundColor("#f3f4f6");
      rowObj.getCell(0).editAsText().setBold(true);
      rowObj.getCell(1).editAsText().setBold(true);
    }
  }

  body.appendParagraph("");

  const totalsData = [
    [isSpanish ? "Subtotal" : "Subtotal", formatDDManualMoney_(row.SUB_TOTAL)],
    [isSpanish ? "Tax" : "Tax", formatDDManualMoney_(row.TAX)],
    [isSpanish ? "Total" : "Total", formatDDManualMoney_(row.TOTAL)]
  ];

  const totalsTable = body.appendTable(totalsData);
  totalsTable.setBorderColor("#d1d5db");
  for (let r = 0; r < totalsTable.getNumRows(); r++) {
    const rowObj = totalsTable.getRow(r);
    rowObj.getCell(0).editAsText().setFontSize(10);
    rowObj.getCell(1).editAsText().setFontSize(10);
    if (r === totalsData.length - 1) {
      rowObj.getCell(0).setBackgroundColor("#111827");
      rowObj.getCell(1).setBackgroundColor("#111827");
      rowObj.getCell(0).editAsText().setBold(true).setForegroundColor("#ffffff");
      rowObj.getCell(1).editAsText().setBold(true).setForegroundColor("#ffffff");
    }
  }
}

function appendDDManualFooter_(body, row, isSpanish) {
  body.appendParagraph("");
  body.appendHorizontalRule();
  const footer = body.appendParagraph(isSpanish
    ? "Gracias por su negocio."
    : "Thank you for your business.");
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setFontSize(9).setForegroundColor("#6b7280");
}

function getDDManualInvoiceCompanyDefaults_(data) {
  data = data || {};
  const branding = (CFG.COMPANY_BRANDING && CFG.COMPANY_BRANDING.DEFAULT) || {};
  return {
    name: String(data.billingCompanyName || branding.companyName || "D&D Premier Solutions Corp").trim(),
    address: String(data.billingCompanyAddress || "").trim(),
    phone: String(data.billingCompanyPhone || "").trim(),
    email: String(data.billingCompanyEmail || "ddpremiersolutions.corp@gmail.com").trim(),
    website: String(data.billingCompanyWebsite || CFG.OWNER_WEB_APP_URL || "https://ddpremiersolutionscorp.com").trim(),
    logoFileId: String(data.billingLogoFileId || CFG.DD_PREMIER_INVOICE_LOGO_FILE_ID || branding.logoFileId || "").trim()
  };
}

function translateDDManualTextForLang_(text, isSpanish) {
  const value = String(text || "").trim();
  if (!value || isSpanish) return value;

  try {
    if (typeof safeTranslateToEn_ === "function") {
      return safeTranslateToEn_(value);
    }

    const translated = LanguageApp.translate(value, "es", "en");
    return translated && translated.trim() ? translated : value;
  } catch (err) {
    return value;
  }
}

function normalizeDDManualMaterialItems_(items, fallbackText) {
  let rawItems = [];
  if (Array.isArray(items)) {
    rawItems = items;
  } else if (typeof items === "string" && items.trim()) {
    try {
      const parsed = JSON.parse(items);
      if (Array.isArray(parsed)) rawItems = parsed;
    } catch (err) {
      rawItems = [];
    }
  }

  rawItems = rawItems.map(function(item) {
    item = item || {};
    return {
      qty: String(item.qty || item.QTY || item.quantity || item.Quantity || "").trim(),
      description: String(item.description || item.DESCRIPTION || item.material || item.Material || "").trim()
    };
  }).filter(function(item) {
    return item.qty || item.description;
  });

  if (rawItems.length) return rawItems;
  return parseDDManualMaterialLines_(fallbackText);
}

function parseDDManualMaterialLines_(text) {
  return String(text || "").split(/\r?\n/).map(function(line) {
    line = String(line || "").trim();
    if (!line) return null;

    const pipeParts = line.split("|");
    if (pipeParts.length >= 2) {
      return {
        qty: String(pipeParts.shift() || "").trim(),
        description: pipeParts.join("|").trim()
      };
    }

    const dashParts = line.split(/\s+-\s+/);
    if (dashParts.length >= 2) {
      return {
        qty: String(dashParts.shift() || "").trim(),
        description: dashParts.join(" - ").trim()
      };
    }

    return {
      qty: "",
      description: line
    };
  }).filter(function(item) {
    return item && (item.qty || item.description);
  });
}

function formatDDManualMaterialItemsText_(items, fallbackText) {
  if (!items || !items.length) return String(fallbackText || "").trim();
  return items.map(function(item) {
    const qty = String(item.qty || "").trim();
    const description = String(item.description || "").trim();
    return (qty ? qty + " | " : "") + description;
  }).join("\n");
}

function getDDManualMaterialItemsFromRow_(row) {
  const json = String(row.MATERIALS_JSON || "").trim();
  if (json) {
    try {
      const parsed = JSON.parse(json);
      const items = normalizeDDManualMaterialItems_(parsed, "");
      if (items.length) return items;
    } catch (err) {
      // Fall back to the readable text columns below.
    }
  }
  return normalizeDDManualMaterialItems_([], row.MATERIALS_USED || row.MATERIALS_LIST || "");
}

function buildDDManualCustomerAddress_(data) {
  const street = String(data.customerStreet || "").trim();
  const city = String(data.customerCity || "").trim();
  const state = String(data.customerState || "").trim();
  const zip = String(data.customerZip || "").trim();
  const cityLine = [city, state, zip].filter(Boolean).join(" ");
  return [street, cityLine].filter(Boolean).join(", ");
}

function buildDDManualCompanyContactLine_(row) {
  return [
    String(row.BILLING_COMPANY_PHONE || "").trim(),
    String(row.BILLING_COMPANY_EMAIL || "").trim(),
    String(row.BILLING_COMPANY_WEBSITE || "").trim()
  ].filter(Boolean).join(" | ");
}

function parseDDManualInvoiceDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value;
  }

  const text = String(value || "").trim();
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDDManualDate_(dateValue) {
  const dateObj = parseDDManualInvoiceDate_(dateValue) || new Date();
  return Utilities.formatDate(dateObj, CFG.TIMEZONE || Session.getScriptTimeZone(), "MM/dd/yyyy");
}

function hasDDManualValue_(value) {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function roundDDManualInvoiceMoney_(value) {
  const parsed = Number(String(value || 0).replace(/[^0-9.-]/g, ""));
  if (!isFinite(parsed)) return 0;
  return Math.round(parsed * 100) / 100;
}

function formatDDManualMoney_(value) {
  const amount = roundDDManualInvoiceMoney_(value);
  return "$" + amount.toFixed(2);
}
