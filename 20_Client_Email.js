// =====================================================
// FILE: 20_Client_Email.gs
// =====================================================

function sendPMReportClientEmail_(payload, report) {
  const ctx = getWorkOrderEmailContext_(payload.row, payload.woNumber);

  const result = trySendClientDocumentEmail_({
    docType: "PM_REPORT",
    companyId: ctx.companyId || payload.companyId || CFG.DEFAULT_COMPANY_ID,
    nsn: ctx.nsn || payload.storeNumber || "",
    woNumber: ctx.woNumber || payload.woNumber || "",
    client: ctx.client || CFG.CLIENT_DEFAULT || "",
    address: ctx.address || payload.address || "",
    title: "PM Report",
    subject: "PM Report - Store " + (ctx.nsn || payload.storeNumber || "") + " - WO " + (ctx.woNumber || payload.woNumber || ""),
    intro: "The PM report has been generated.",
    links: [
      { label: "PM Report English", url: report.pdfEnUrl || "" },
      { label: "PM Report Spanish", url: report.pdfEsUrl || "" }
    ]
  });

  markWorkOrderDocumentEmailResult_(payload.row, payload.woNumber, "PM_REPORT", result);
  return result;
}

function sendQuoteCreatedClientEmail_(data, quoteRow, quoteSheetRow) {
  const ctx = getWorkOrderEmailContext_(data.rowNumber, data.wo_number);

  const result = trySendClientDocumentEmail_({
    docType: "QUOTE",
    companyId: ctx.companyId || data.companyId || CFG.DEFAULT_COMPANY_ID,
    nsn: ctx.nsn || data.ns_number || "",
    woNumber: ctx.woNumber || data.wo_number || "",
    client: ctx.client || data.nombre_cliente || "",
    address: ctx.address || normalizeQuoteAddressForTemplate_(data).full || "",
    title: "Quote",
    subject: "Quote " + (data.quote_number || "") + " - Store " + (ctx.nsn || data.ns_number || "") + " - WO " + (ctx.woNumber || data.wo_number || ""),
    intro: "The quote has been created.",
    links: [
      { label: "Quote English", url: quoteRow.PDF_EN_URL || "" },
      { label: "Quote Spanish", url: quoteRow.PDF_ES_URL || "" }
    ]
  });

  markWorkOrderDocumentEmailResult_(data.rowNumber, data.wo_number, "QUOTE", result);
  markSheetRowEmailResult_("QUOTES", quoteSheetRow, result);
  return result;
}

function sendInvoiceCreatedClientEmail_(invoiceRow, workOrderRow, invoiceSheetRow) {
  const ctx = getWorkOrderEmailContext_(workOrderRow, invoiceRow.WO_NUMBER);

  const nsn = ctx.nsn || invoiceRow.NS || "";
  const invoiceNumber = invoiceRow.INVOICE_NUMBER || invoiceRow.Invoice || "";

  const result = trySendClientDocumentEmail_({
    docType: "INVOICE",
    companyId: ctx.companyId || invoiceRow.COMPANY_ID || CFG.DEFAULT_COMPANY_ID,
    nsn: nsn,
    woNumber: ctx.woNumber || invoiceRow.WO_NUMBER || "",
    client: ctx.client || invoiceRow.CLIENTE || CFG.CLIENT_DEFAULT || "",
    address: ctx.address || invoiceRow.STORE_ADDRESS || "",
    title: "Invoice",
    subject: "Invoice " + invoiceNumber + " - Store " + nsn + " - WO " + (ctx.woNumber || invoiceRow.WO_NUMBER || ""),
    intro: "The invoice has been created.",
    links: [
      { label: "Invoice English", url: invoiceRow.PDF_EN_URL || "" },
      { label: "Invoice Spanish", url: invoiceRow.PDF_ES_URL || "" }
    ]
  });

  markWorkOrderDocumentEmailResult_(workOrderRow, invoiceRow.WO_NUMBER, "INVOICE", result);
  markSheetRowEmailResult_("INVOICES", invoiceSheetRow, result);
  return result;
}

function trySendClientDocumentEmail_(info) {
  try {
    return sendClientDocumentEmail_(info);
  } catch (err) {
    Logger.log("ERROR trySendClientDocumentEmail_: " + err);
    notifySystemError_("CLIENT_DOCUMENT_EMAIL_ERROR", err, {
      module: "CLIENT_EMAIL",
      docType: info && info.docType,
      companyId: info && info.companyId,
      woNumber: info && info.woNumber,
      nsn: info && info.nsn,
      title: info && info.title
    });
    return {
      sent: false,
      to: "",
      actualTo: "",
      status: "ERROR: " + (err && err.message ? err.message : err)
    };
  }
}

function sendClientDocumentEmail_(info) {
  if (CFG.AUTO_SEND_CLIENT_EMAILS === false) {
    return {
      sent: false,
      to: "",
      actualTo: "",
      status: "DISABLED"
    };
  }

  const recipients = resolveClientEmailRecipients_(info.docType, info.companyId, info.nsn);

  if (!recipients.actualTo.length) {
    notifySystemError_("CLIENT_EMAIL_NO_RECIPIENTS", new Error("No client email recipients configured."), {
      module: "CLIENT_EMAIL",
      docType: info.docType || "",
      companyId: info.companyId || "",
      woNumber: info.woNumber || "",
      nsn: info.nsn || ""
    });
    return {
      sent: false,
      to: "",
      actualTo: "",
      status: "NO_RECIPIENTS"
    };
  }

  if (!recipients.to.length) {
    notifySystemError_("CLIENT_EMAIL_TEST_RECIPIENT_MISSING", new Error("Client email test recipient is missing."), {
      module: "CLIENT_EMAIL",
      docType: info.docType || "",
      companyId: info.companyId || "",
      woNumber: info.woNumber || "",
      nsn: info.nsn || "",
      actualTo: recipients.actualTo.join(", ")
    });
    return {
      sent: false,
      to: "",
      actualTo: recipients.actualTo.join(", "),
      status: "TEST_RECIPIENT_MISSING"
    };
  }

  const attachmentInfo = buildClientEmailAttachments_(info.links || []);
  const isTest = isClientEmailTestMode_();
  const subject = (isTest ? "[TEST] " : "") + String(info.subject || info.title || "Document");
  const htmlBody = buildClientDocumentEmailBody_(info, recipients, attachmentInfo);

  MailApp.sendEmail({
    to: recipients.to.join(","),
    subject: subject,
    htmlBody: htmlBody,
    attachments: attachmentInfo.attachments
  });

  return {
    sent: true,
    to: recipients.to.join(", "),
    actualTo: recipients.actualTo.join(", "),
    status: isTest
      ? "TEST_SENT" + (attachmentInfo.omitted ? "_LINKS_ONLY" : "")
      : "SENT" + (attachmentInfo.omitted ? "_LINKS_ONLY" : "")
  };
}

function resolveClientEmailRecipients_(docType, companyId, nsn) {
  const store = getStoreByNSN_(nsn, companyId) || {};
  const type = String(docType || "").toUpperCase();
  let fields = [];

  if (type === "PM_REPORT") {
    fields = [
      store.storeEmails,
      store.supervisorEmail
    ];
  }

  if (type === "QUOTE") {
    fields = [
      store.supervisorEmail
    ];
  }

  if (type === "INVOICE") {
    fields = [
      store.storeEmails,
      store.supervisorEmail
    ];
  }

  const actualTo = parseEmailList_(fields.join(","));
  const to = isClientEmailTestMode_()
    ? parseEmailList_(CFG.CLIENT_EMAIL_TEST_TO || "")
    : actualTo;

  return {
    to: to,
    actualTo: actualTo,
    store: store
  };
}

function buildClientDocumentEmailBody_(info, recipients, attachmentInfo) {
  const links = (info.links || []).filter(function(link) {
    return link && link.url;
  });

  const linkHtml = links.length
    ? "<ul>" + links.map(function(link) {
        return "<li><a href='" + escapeHtmlForEmail_(link.url) + "'>" +
          escapeHtmlForEmail_(link.label || "Open PDF") +
          "</a></li>";
      }).join("") + "</ul>"
    : "<p>No PDF links were available.</p>";

  const testBanner = isClientEmailTestMode_()
    ? "<div style='background:#fff7ed;border:1px solid #fed7aa;padding:12px;border-radius:8px;margin-bottom:12px;'>" +
      "<b>TEST MODE:</b> This email was sent only to " + escapeHtmlForEmail_(recipients.to.join(", ")) +
      ". Real recipients would be: " + escapeHtmlForEmail_(recipients.actualTo.join(", ")) +
      "</div>"
    : "";

  const attachmentNote = attachmentInfo.omitted
    ? "<p><b>Note:</b> PDF attachments were omitted because the total size was too large. Use the links below.</p>"
    : "<p>The PDF files are attached when size allows. Links are included below as backup.</p>";

  return testBanner +
    "<h2>" + escapeHtmlForEmail_(info.title || "Document") + "</h2>" +
    "<p>" + escapeHtmlForEmail_(info.intro || "A document has been generated.") + "</p>" +
    "<hr>" +
    "<p><b>Company:</b> " + escapeHtmlForEmail_(info.companyId || "") + "</p>" +
    "<p><b>Client:</b> " + escapeHtmlForEmail_(info.client || "") + "</p>" +
    "<p><b>Store / NSN:</b> " + escapeHtmlForEmail_(info.nsn || "") + "</p>" +
    "<p><b>Work Order:</b> " + escapeHtmlForEmail_(info.woNumber || "") + "</p>" +
    "<p><b>Address:</b> " + escapeHtmlForEmail_(info.address || "") + "</p>" +
    "<p><b>Date:</b> " + escapeHtmlForEmail_(Utilities.formatDate(new Date(), CFG.TIMEZONE, "MM/dd/yyyy hh:mm a")) + "</p>" +
    "<hr>" +
    attachmentNote +
    linkHtml;
}

function isClientEmailTestMode_() {
  return CFG.CLIENT_EMAIL_TEST_MODE === true;
}

function buildClientEmailAttachments_(links) {
  const maxBytes = 20 * 1024 * 1024;
  const attachments = [];
  let totalBytes = 0;
  let omitted = false;

  (links || []).forEach(function(link) {
    if (!link || !link.url || omitted) return;

    try {
      const id = extractDriveFileIdForEmail_(link.url);
      const file = DriveApp.getFileById(id);
      const blob = file.getBlob().setName(file.getName());
      const bytes = blob.getBytes().length;

      if (totalBytes + bytes > maxBytes) {
        omitted = true;
        return;
      }

      totalBytes += bytes;
      attachments.push(blob);
    } catch (err) {
      Logger.log("Client email attachment skipped: " + err);
    }
  });

  return {
    attachments: omitted ? [] : attachments,
    omitted: omitted
  };
}

function getWorkOrderEmailContext_(rowNumber, woNumber) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {};

  const headers = sh.getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  let row = Number(rowNumber || 0);
  const targetWO = String(woNumber || "").trim();

  if (row < 2 || row > lastRow || (targetWO && String(getEmailCell_(sh, row, headers, "WO_NUMBER") || "").trim() !== targetWO)) {
    row = findWorkOrderRowForEmail_(sh, headers, targetWO);
  }

  if (!row) return {};

  const values = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const obj = {};

  headers.forEach(function(h, i) {
    obj[h] = values[i];
  });

  const companyId = String(obj.COMPANY_ID || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  const nsn = String(obj.NSN || obj["NSN #"] || obj.NS || "").trim();
  const store = getStoreByNSN_(nsn, companyId) || {};

  return {
    row: row,
    woNumber: obj.WO_NUMBER || targetWO || "",
    companyId: companyId,
    client: obj.CLIENT || obj.CLIENTE || store.client || CFG.CLIENT_DEFAULT || "",
    nsn: nsn || store.nsn || "",
    address: store.fullAddress || obj.STORE_ADDRESS || obj.ADDRESS || "",
    store: store
  };
}

function findWorkOrderRowForEmail_(sh, headers, woNumber) {
  if (!woNumber) return 0;

  const idx = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf("WO_NUMBER");

  if (idx === -1 || sh.getLastRow() < 2) return 0;

  const values = sh.getRange(2, idx + 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === String(woNumber || "").trim()) {
      return i + 2;
    }
  }

  return 0;
}

function markWorkOrderDocumentEmailResult_(rowNumber, woNumber, prefix, result) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) return false;

  ensureEmailResultColumns_(sh, prefix);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  let row = Number(rowNumber || 0);
  const targetWO = String(woNumber || "").trim();

  if (row < 2 || row > sh.getLastRow() || (targetWO && String(getEmailCell_(sh, row, headers, "WO_NUMBER") || "").trim() !== targetWO)) {
    row = findWorkOrderRowForEmail_(sh, headers, targetWO);
  }

  if (!row) return false;

  setEmailCell_(sh, row, headers, prefix + "_EMAIL_SENT", result && result.sent ? "YES" : "NO");
  setEmailCell_(sh, row, headers, prefix + "_EMAIL_SENT_AT", new Date());
  setEmailCell_(sh, row, headers, prefix + "_EMAIL_TO", result && result.actualTo ? result.actualTo : "");
  setEmailCell_(sh, row, headers, prefix + "_EMAIL_STATUS", result && result.status ? result.status : "");

  return true;
}

function markSheetRowEmailResult_(sheetName, rowNumber, result) {
  if (!rowNumber) return false;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return false;

  ensureEmailResultColumns_(sh, "CLIENT");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  setEmailCell_(sh, rowNumber, headers, "CLIENT_EMAIL_SENT", result && result.sent ? "YES" : "NO");
  setEmailCell_(sh, rowNumber, headers, "CLIENT_EMAIL_SENT_AT", new Date());
  setEmailCell_(sh, rowNumber, headers, "CLIENT_EMAIL_TO", result && result.actualTo ? result.actualTo : "");
  setEmailCell_(sh, rowNumber, headers, "CLIENT_EMAIL_STATUS", result && result.status ? result.status : "");

  return true;
}

function ensureEmailResultColumns_(sh, prefix) {
  const required = [
    prefix + "_EMAIL_SENT",
    prefix + "_EMAIL_SENT_AT",
    prefix + "_EMAIL_TO",
    prefix + "_EMAIL_STATUS"
  ];

  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim().toUpperCase();
    });

  required.forEach(function(h) {
    if (headers.indexOf(h.toUpperCase()) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      headers.push(h.toUpperCase());
    }
  });
}

function getEmailCell_(sh, rowNumber, headers, headerName) {
  const idx = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf(String(headerName || "").trim().toUpperCase());

  if (idx === -1) return "";
  return sh.getRange(Number(rowNumber), idx + 1).getValue();
}

function setEmailCell_(sh, rowNumber, headers, headerName, value) {
  const idx = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf(String(headerName || "").trim().toUpperCase());

  if (idx === -1) return false;

  sh.getRange(Number(rowNumber), idx + 1).setValue(value);
  return true;
}

function parseEmailList_(value) {
  const seen = {};
  const emails = [];

  String(value || "")
    .split(/[,\n;]/)
    .map(function(x) {
      return String(x || "").trim();
    })
    .filter(Boolean)
    .forEach(function(email) {
      const key = email.toLowerCase();

      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
      if (seen[key]) return;

      seen[key] = true;
      emails.push(email);
    });

  return emails;
}

function extractDriveFileIdForEmail_(urlOrId) {
  const s = String(urlOrId || "").trim();

  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;

  let m = s.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];

  m = s.match(/[-\w]{25,}/);
  if (m && m[0]) return m[0];

  throw new Error("Could not extract Drive file ID.");
}

function escapeHtmlForEmail_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
