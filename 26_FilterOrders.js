const FILTER_ORDER_FOLDER_NAME_ = "PEDIDO DE FILTROS";
const FILTER_ORDER_UNITED_ACCOUNT_ = "11437479";

const FILTER_ORDER_CATALOG_ = [
  { sku: "20X24X2AB40", size: "20x24x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 5.87 },
  { sku: "24X24X2AB40", size: "24x24x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 6.23 },
  { sku: "16X20X2AB40", size: "16x20x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 4.84 },
  { sku: "20X30X2AB40", size: "20x30x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 10.96 },
  { sku: "14X30X2AB40", size: "14x30x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 12.69 },
  { sku: "20X25X2AB40", size: "20x25x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 5.93 },
  { sku: "16X25X2AB40", size: "16x25x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 5.37 },
  { sku: "18X24X2AB40", size: "18x24x2", description: "MERV 8 PLEATED FILTER STANDARD CAPACITY", price: 5.70 }
];

function createFilterOrderPdf(payload, sessionToken) {
  payload = payload || {};

  const companyId = String(payload.companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES", "SYSTEM"], companyId);
  const items = sanitizeFilterOrderItems_(payload.items);

  if (!items.length) {
    throw new Error("Seleccione al menos un filtro con cantidad.");
  }

  const timezone = CFG.TIMEZONE || "America/New_York";
  const now = new Date();
  const stamp = Utilities.formatDate(now, timezone, "yyyyMMdd-HHmmss");
  const displayDate = Utilities.formatDate(now, timezone, "MM/dd/yyyy HH:mm");
  const companyName = getFilterOrderCompanyName_(companyId);
  const requestedBy = String(payload.requestedBy || session.name || session.email || "").trim();
  const notes = String(payload.notes || "").trim().slice(0, 2000);
  const fileName = ("Pedido_Filtros_" + companyId + "_" + stamp + ".pdf").replace(/[\\/:*?"<>|]/g, "-");

  const folder = getOrCreateFilterOrderFolder_();
  const doc = DocumentApp.create("TMP_" + fileName.replace(".pdf", ""));
  const body = doc.getBody();

  body.appendParagraph(companyName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("PEDIDO DE FILTROS").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendTable([
    ["Fecha", displayDate],
    ["Compania", companyName + " (" + companyId + ")"],
    ["Solicitado por", requestedBy || ""],
    ["Cuenta United Refrigeration", FILTER_ORDER_UNITED_ACCOUNT_]
  ]);

  body.appendParagraph("");

  const tableRows = [[
    "Filtro",
    "SKU",
    "Descripcion",
    "Cantidad",
    "Precio ref.",
    "Total ref."
  ]];

  let estimatedTotal = 0;

  items.forEach(function(item) {
    const lineTotal = item.qty * item.price;
    estimatedTotal += lineTotal;

    tableRows.push([
      item.size,
      item.sku,
      item.description,
      String(item.qty),
      formatFilterOrderMoney_(item.price),
      formatFilterOrderMoney_(lineTotal)
    ]);
  });

  body.appendTable(tableRows);

  const totalParagraph = body.appendParagraph("Total referencia: " + formatFilterOrderMoney_(estimatedTotal));
  totalParagraph.editAsText().setBold(true);

  if (notes) {
    body.appendParagraph("");
    body.appendParagraph("Notas").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(notes);
  }

  body.appendParagraph("");
  const footer = body.appendParagraph("Precios de referencia tomados del ultimo documento de United Refrigeration. Verificar disponibilidad y precio final al ordenar.");
  footer.editAsText().setFontSize(9).setForegroundColor("#64748b");

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = getFilterOrderPdfBlob_(docFile, fileName);
  const pdfFile = folder.createFile(pdfBlob);

  hardenGeneratedPdfFile_(pdfFile);
  docFile.setTrashed(true);

  addAuditLog_("FILTER_ORDER", "FILTER_ORDER_PDF_CREATED", companyId, "DRIVE_FILE", pdfFile.getId(), session, {
    folderName: FILTER_ORDER_FOLDER_NAME_,
    fileName: fileName,
    itemCount: items.length,
    totalQty: items.reduce(function(sum, item) { return sum + item.qty; }, 0),
    estimatedTotal: estimatedTotal
  });

  return {
    success: true,
    name: fileName,
    fileId: pdfFile.getId(),
    url: pdfFile.getUrl(),
    folderUrl: folder.getUrl()
  };
}

function sanitizeFilterOrderItems_(items) {
  const bySku = {};
  FILTER_ORDER_CATALOG_.forEach(function(item) {
    bySku[item.sku] = item;
  });

  const totals = {};
  (items || []).forEach(function(item) {
    const sku = String(item && item.sku || "").trim().toUpperCase();
    const qty = Math.floor(Number(item && item.qty || 0));

    if (!bySku[sku] || qty <= 0) return;
    totals[sku] = (totals[sku] || 0) + qty;
  });

  return Object.keys(totals).map(function(sku) {
    const base = bySku[sku];
    return {
      sku: base.sku,
      size: base.size,
      description: base.description,
      price: base.price,
      qty: totals[sku]
    };
  });
}

function getOrCreateFilterOrderFolder_() {
  if (CFG.FILTER_ORDERS_FOLDER_ID) {
    return DriveApp.getFolderById(CFG.FILTER_ORDERS_FOLDER_ID);
  }

  const parent = CFG.FILTER_ORDERS_PARENT_FOLDER_ID
    ? DriveApp.getFolderById(CFG.FILTER_ORDERS_PARENT_FOLDER_ID)
    : DriveApp.getRootFolder();

  const folders = parent.getFoldersByName(FILTER_ORDER_FOLDER_NAME_);
  return folders.hasNext() ? folders.next() : parent.createFolder(FILTER_ORDER_FOLDER_NAME_);
}

function getFilterOrderPdfBlob_(docFile, fileName) {
  let lastErr = "";

  for (let i = 0; i < 5; i++) {
    try {
      return docFile.getAs(MimeType.PDF).setName(fileName);
    } catch (err) {
      lastErr = err;
      Utilities.sleep(1500 * (i + 1));
    }
  }

  throw new Error("No se pudo crear el PDF de filtros. Ultimo error: " + lastErr);
}

function getFilterOrderCompanyName_(companyId) {
  try {
    return getCompanyName(companyId) || companyId;
  } catch (err) {
    return companyId || CFG.APP_NAME;
  }
}

function formatFilterOrderMoney_(value) {
  return "$" + Number(value || 0).toFixed(2);
}
