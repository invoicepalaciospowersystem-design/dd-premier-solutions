// =====================================================
// FILE: 18_PM_Invoice.gs
// =====================================================

const PM_CFG = {
  PM_ROOT_FOLDER_ID: "1iJm3u71GXZE-D-q6BgeSYcJRQ7ePqbY8",

  TEMPLATE_DOC_ID_ES: "1JS16ewBRhO1vZIGANjY_YQP1AIXpkQtRuG-lPXmAUAY",
  TEMPLATE_DOC_ID_EN: "1KtgBVI8NifTEkY82_kljbL1loV78ZWeniyc-yMz_ZHA",

  CLIENT_NAME: "McDonald's",
  SUBTOTAL_FIXED: 391.67,
  TAX_RATE: 0.07,

  TROUBLE_ES: "Mantenimiento preventivo a equipos de HVAC.",
  TROUBLE_EN: "Scheduled preventive maintenance on HVAC.",

  MAKE_ES: "MANTENIMIENTO PREVENTIVO",
  MODEL_ES: "Varios equipos",
  SERIAL_ES: "PM",

  MAKE_EN: "PREVENTIVE MAINTENANCE",
  MODEL_EN: "Multiple units",
  SERIAL_EN: "PM",

  PM_TEXTS: {
    MENSUAL: {
      ES: "Mantenimiento preventivo mensual a sistemas de aire acondicionado, incluyendo:\n\n• Limpieza de tuberías de drenaje\n• Aplicación de tabletas anti-algas en bandejas de condensado\n• Limpieza de condensadores\n• Medición de temperaturas del aire en suministro en las áreas del restaurante\n\nCualquier anomalía detectada durante el servicio será incluida en el reporte correspondiente para su evaluación y recomendación de reparación.",
      EN: "Monthly preventive maintenance on air conditioning systems, including:\n\n• Drain line cleaning\n• Anti-algae tablets applied in condensate drain pans\n• Condenser cleaning\n• Supply air temperature measurements in restaurant areas\n\nAny abnormal condition found during the service will be included in the corresponding report for evaluation and repair recommendation."
    },

    TRIMESTRAL: {
      ES: "Mantenimiento preventivo trimestral a sistemas de aire acondicionado, incluyendo:\n\n• Limpieza de tuberías de drenaje\n• Aplicación de tabletas anti-algas en bandejas de condensado\n• Limpieza de condensadores\n• Cambio de filtros de aire\n• Medición de temperaturas del aire en suministro en las áreas del restaurante\n\nCualquier anomalía detectada durante el servicio será incluida en el reporte correspondiente para su evaluación y recomendación de reparación.",
      EN: "Quarterly preventive maintenance on air conditioning systems, including:\n\n• Drain line cleaning\n• Anti-algae tablets applied in condensate drain pans\n• Condenser cleaning\n• Air filter replacement\n• Supply air temperature measurements in restaurant areas\n\nAny abnormal condition found during the service will be included in the corresponding report for evaluation and repair recommendation."
    },

    ANNUAL: {
      ES: "Mantenimiento preventivo anual a sistemas de aire acondicionado, incluyendo:\n\n• Limpieza de tuberías de drenaje\n• Aplicación de tabletas anti-algas en bandejas de condensado\n• Limpieza de condensadores\n• Cambio de filtros de aire\n• Limpieza de evaporadores\n• Limpieza de blowers\n• Verificación de presiones de trabajo del refrigerante\n• Medición de temperaturas del aire en suministro en las áreas del restaurante\n\nCualquier anomalía detectada durante el servicio será incluida en el reporte correspondiente para su evaluación y recomendación de reparación.",
      EN: "Annual preventive maintenance on air conditioning systems, including:\n\n• Drain line cleaning\n• Anti-algae tablets applied in condensate drain pans\n• Condenser cleaning\n• Air filter replacement\n• Evaporator cleaning\n• Blower cleaning\n• Refrigerant operating pressure verification\n• Supply air temperature measurements in restaurant areas\n\nAny abnormal condition found during the service will be included in the corresponding report for evaluation and repair recommendation."
    }
  }
};

function createPMInvoiceFromCloseOrder_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shInv = ss.getSheetByName("INVOICES");

  if (!shInv) throw new Error("No existe la hoja INVOICES.");

  const companyId = String(data.COMPANY_ID || "").trim().toUpperCase();
  const woNumber = String(data.WO_NUMBER || "").trim();
  const rowNumber = Number(data.ROW_NUMBER || 0);

  const pmType = normalizePMType_(data.PM_TYPE || data.TIPO_MANTENIMIENTO || "MENSUAL");

  if (!PM_CFG.PM_TEXTS[pmType]) {
    throw new Error("PM_TYPE inválido: " + pmType);
  }

  const invoiceNumber = generateInvoiceNumber_(companyId);
  const now = new Date();

  const store = getPMStoreData_(data);
  const folder = getPMInvoiceFolder_(now, store.nsn || data.NSN || "NO_NSN");

  let f1 = data.I1_DESC || "";
  let q1 = data.I1_QTY || "";
  let f2 = data.I2_DESC || "";
  let q2 = data.I2_QTY || "";
  let f3 = data.I3_DESC || "";
  let q3 = data.I3_QTY || "";
  let f4 = data.I4_DESC || "";
  let q4 = data.I4_QTY || "";

  if (pmType === "MENSUAL") {
    f1 = ""; q1 = "";
    f2 = ""; q2 = "";
    f3 = ""; q3 = "";
    f4 = ""; q4 = "";
  }

  const subTotal = round2PM_(PM_CFG.SUBTOTAL_FIXED);
  const tax = round2PM_(subTotal * PM_CFG.TAX_RATE);
  const total = round2PM_(subTotal + tax);

  const pdfs = generatePMPdfs_({
    invoiceNumber: invoiceNumber,
    invoiceDate: now,
    pmType: pmType,
    store: store,
    folder: folder,
    signature: data.SIGNATURE || "",
    f1: f1, q1: q1,
    f2: f2, q2: q2,
    f3: f3, q3: q3,
    f4: f4, q4: q4,
    subTotal: subTotal,
    tax: tax,
    total: total
  });

  const invHeaders = shInv
    .getRange(1, 1, 1, shInv.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  const invoiceRow = {
    WO_NUMBER: woNumber,
    WO_TYPE: "PM_FORM",
    WO_TYPO: "PM_FORM",
    PM_TYPE: pmType,
    INVOICE_TYPE: "PM",

    INVOICE_NUMBER: invoiceNumber,
    Invoice: invoiceNumber,

    DATE_INVOICE: now,
    Timestamp: now,

    INVOICE_TOTAL: total,
    TOTAL: total,
    GRAND_TOTAL: total,
    SUB_TOTAL: subTotal,

    MATERIAL_COST: 0,
    PARTS: 0,
    COMPRA: 0,
    INVERSION: 0,

    LABOR_AMOUNT: subTotal,
    LABOR: subTotal,

    TAX_AMOUNT: tax,
    TAX: tax,

    LABOR_HOURS: 0,
    HORAS: 0,

    COMPANY_ID: companyId,
    CLIENTE: data.CLIENT || PM_CFG.CLIENT_NAME,
    NS: store.nsn || data.NSN || "",

    "TECHNICIAN EMAIL": data.TECHNICIAN_EMAIL || "",
    "TECHNICIAN NAME": data.TECHNICIAN || "",

    PROCESO: "RECIBIDO",
    WORK_PERFORMED: PM_CFG.PM_TEXTS[pmType].EN,

    I1_QTY: q1,
    I1_DESC: f1,
    I2_QTY: q2,
    I2_DESC: f2,
    I3_QTY: q3,
    I3_DESC: f3,
    I4_QTY: q4,
    I4_DESC: f4,

    STORE_ADDRESS: store.address || "",
    STORE_CITY: store.city || "",
    STORE_STATE: store.state || "",
    STORE_ZIP: store.zip || "",

    REPORTED_PROBLEM: PM_CFG.TROUBLE_EN,
    EQUIPMENT_MAKE: PM_CFG.MAKE_EN,
    EQUIPMENT_MODEL: PM_CFG.MODEL_EN,
    EQUIPMENT_SERIAL: PM_CFG.SERIAL_EN,

    PDF_ES_URL: pdfs.pdfEsUrl,
    PDF_EN_URL: pdfs.pdfEnUrl,
    DOC_ES_URL: pdfs.docEsUrl,
    DOC_EN_URL: pdfs.docEnUrl,

    SIGNATURE: data.SIGNATURE || "",
    NOTES: data.NOTES || ""
  };

  const rowValues = invHeaders.map(function(h) {
    return invoiceRow[h] !== undefined ? invoiceRow[h] : "";
  });

  shInv.appendRow(rowValues);
  const invoiceSheetRow = shInv.getLastRow();

  let emailResult = null;
  try {
    emailResult = sendInvoiceCreatedClientEmail_(invoiceRow, rowNumber, invoiceSheetRow);
  } catch (emailErr) {
    Logger.log("ERROR sendInvoiceCreatedClientEmail_ PM: " + emailErr);
    notifySystemError_("PM_INVOICE_EMAIL_ERROR", emailErr, {
      module: "PM_INVOICE",
      companyId: companyId,
      woNumber: woNumber,
      invoiceNumber: invoiceNumber,
      pmType: pmType
    });
    emailResult = {
      sent: false,
      status: "ERROR: " + emailErr
    };
  }

    try {
    savePMEconomy({
      COMPANY_ID: "PPS",
      WO_NUMBER: woNumber,
      CLIENT: data.CLIENT || PM_CFG.CLIENT_NAME,
      NSN: store.nsn || data.NSN || "",
      DATE_COMPLETED: now,
      INVOICE_NUMBER: invoiceNumber,
      TECHNICIANS: data.TECHNICIAN || "",
      HORAS: 0,
      STATUS: "INVOICED"
    });
  } catch (err) {
    Logger.log("ERROR save PM economy: " + err);
    notifySystemError_("PM_ECONOMY_SAVE_ERROR", err, {
      module: "PM_INVOICE",
      companyId: companyId,
      woNumber: woNumber,
      invoiceNumber: invoiceNumber,
      pmType: pmType
    });
  }

  try {
    syncCloseOrderToOldSystem_(invoiceRow);
  } catch (err) {
    Logger.log("ERROR sync PM old system: " + err);
    notifySystemError_("PM_OLD_SYSTEM_SYNC_ERROR", err, {
      module: "PM_INVOICE",
      companyId: companyId,
      woNumber: woNumber,
      invoiceNumber: invoiceNumber,
      pmType: pmType
    });
  }

  try {
    updateTechOrderStatusInternal_(rowNumber, "COMPLETED", "PM Invoice");
  } catch (err) {
    Logger.log("ERROR updateTechOrderStatus PM: " + err);
    notifySystemError_("PM_ORDER_STATUS_UPDATE_ERROR", err, {
      module: "PM_INVOICE",
      companyId: companyId,
      woNumber: woNumber,
      rowNumber: rowNumber,
      targetStatus: "COMPLETED"
    });
  }

  return {
    success: true,
    invoiceNumber: invoiceNumber,
    woNumber: woNumber,
    invoiceType: "PM",
    pmType: pmType,
    pdfEnUrl: pdfs.pdfEnUrl,
    pdfEsUrl: pdfs.pdfEsUrl,
    emailResult: emailResult
  };
}

function normalizePMType_(value) {
  const v = String(value || "").trim().toUpperCase();

  if (v === "MONTHLY") return "MENSUAL";
  if (v === "MENSUAL") return "MENSUAL";

  if (v === "QUARTERLY") return "TRIMESTRAL";
  if (v === "TRIMESTRAL") return "TRIMESTRAL";

  if (v === "ANNUAL") return "ANNUAL";
  if (v === "ANUAL") return "ANNUAL";

  return v || "MENSUAL";
}

function generatePMPdfs_(info) {
  const mapES = buildPMMap_(info, "ES");
  const mapEN = buildPMMap_(info, "EN");

  const fileBaseEs = info.invoiceNumber + " - " + info.store.nsn + " - PM - ES";
  const fileBaseEn = info.invoiceNumber + " - " + info.store.nsn + " - PM - EN";

  const es = buildPMDocAndPdf_({
    templateId: PM_CFG.TEMPLATE_DOC_ID_ES,
    folder: info.folder,
    fileBase: fileBaseEs,
    replacements: mapES
  });

  const en = buildPMDocAndPdf_({
    templateId: PM_CFG.TEMPLATE_DOC_ID_EN,
    folder: info.folder,
    fileBase: fileBaseEn,
    replacements: mapEN
  });

  return {
    pdfEsUrl: es.pdfUrl,
    docEsUrl: es.docUrl,
    pdfEnUrl: en.pdfUrl,
    docEnUrl: en.docUrl
  };
}

function buildPMMap_(info, lang) {
  const isES = lang === "ES";

  const sub = splitMoneyPM_(info.subTotal);
  const tx = splitMoneyPM_(info.tax);
  const tot = splitMoneyPM_(info.total);

  return {
    "{{invoice_number}}": info.invoiceNumber,
    "{{invoice_date}}": Utilities.formatDate(info.invoiceDate, CFG.TIMEZONE, "MM/dd/yyyy"),

    "{{nombre_cliente}}": PM_CFG.CLIENT_NAME,
    "{{ns_number}}": info.store.nsn || "",
    "{{direccion}}": info.store.address || "",
    "{{ciudad}}": info.store.city || "",
    "{{estado}}": info.store.state || "",
    "{{zip_code}}": info.store.zip || "",
    "{{vendor_id}}": info.store.vendor_id || "",

    "{{trabajo_realizado}}": isES
      ? PM_CFG.PM_TEXTS[info.pmType].ES
      : PM_CFG.PM_TEXTS[info.pmType].EN,

    "{{problema_reportado}}": isES ? PM_CFG.TROUBLE_ES : PM_CFG.TROUBLE_EN,
    "{{marca}}": isES ? PM_CFG.MAKE_ES : PM_CFG.MAKE_EN,
    "{{modelo}}": isES ? PM_CFG.MODEL_ES : PM_CFG.MODEL_EN,
    "{{numero_serie}}": isES ? PM_CFG.SERIAL_ES : PM_CFG.SERIAL_EN,

    "{{i1_qty}}": info.q1 || "",
    "{{i1_desc}}": info.f1 || "",
    "{{i2_qty}}": info.q2 || "",
    "{{i2_desc}}": info.f2 || "",
    "{{i3_qty}}": info.q3 || "",
    "{{i3_desc}}": info.f3 || "",
    "{{i4_qty}}": info.q4 || "",
    "{{i4_desc}}": info.f4 || "",

    "{{i1_part}}": "",
    "{{i1_unit}}": "",
    "{{i1_labor}}": "",
    "{{i1_rate}}": "",
    "{{i1_amount}}": "",

    "{{i2_part}}": "",
    "{{i2_unit}}": "",
    "{{i2_labor}}": "",
    "{{i2_rate}}": "",
    "{{i2_amount}}": "",

    "{{i3_part}}": "",
    "{{i3_unit}}": "",
    "{{i3_labor}}": "",
    "{{i3_rate}}": "",
    "{{i3_amount}}": "",

    "{{i4_part}}": "",
    "{{i4_unit}}": "",
    "{{i4_labor}}": "",
    "{{i4_rate}}": "",
    "{{i4_amount}}": "",

    "{{total_material_d}}": "",
    "{{total_material_c}}": "",
    "{{total_labor_d}}": "",
    "{{total_labor_c}}": "",

    "{{sub_total_d}}": sub.d,
    "{{sub_total_c}}": sub.c,
    "{{tax_d}}": tx.d,
    "{{tax_c}}": tx.c,
    "{{total_d}}": tot.d,
    "{{total_c}}": tot.c,

    "{{firma_cliente}}": info.signature || ""
  };
}

function buildPMDocAndPdf_(cfg) {
  const templateFile = DriveApp.getFileById(cfg.templateId);
  const docCopy = templateFile.makeCopy(cfg.fileBase + " (TEMP)", cfg.folder);
  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  Object.keys(cfg.replacements).forEach(function(key) {
    body.replaceText(escapePMRegex_(key), String(cfg.replacements[key] || ""));
  });

  doc.saveAndClose();

  const pdfBlob = docCopy.getBlob()
    .getAs("application/pdf")
    .setName(cfg.fileBase + ".pdf");

  const pdfFile = cfg.folder.createFile(pdfBlob);
  const docUrl = docCopy.getUrl();

  docCopy.setTrashed(true);

  return {
    pdfUrl: pdfFile.getUrl(),
    docUrl: docUrl
  };
}

function getPMInvoiceFolder_(dateValue, nsn) {
  const root = DriveApp.getFolderById(PM_CFG.PM_ROOT_FOLDER_ID);
  const year = Utilities.formatDate(dateValue, CFG.TIMEZONE, "yyyy");
  const day = Utilities.formatDate(dateValue, CFG.TIMEZONE, "yyyy-MM-dd");

  let folder = root;

  [year, day, String(nsn || "NO_NSN")].forEach(function(name) {
    const it = folder.getFoldersByName(name);
    folder = it.hasNext() ? it.next() : folder.createFolder(name);
  });

  return folder;
}

function getPMStoreData_(data) {
  return {
    nsn: data.NSN || data.NS || "",
    address: data.STORE_ADDRESS || "",
    city: data.STORE_CITY || "",
    state: data.STORE_STATE || "",
    zip: data.STORE_ZIP || "",
    vendor_id: data.VENDOR_ID || data.VendorID || ""
  };
}

function round2PM_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function splitMoneyPM_(n) {
  const fixed = round2PM_(n).toFixed(2);
  const parts = fixed.split(".");
  return {
    d: parts[0],
    c: parts[1]
  };
}

function escapePMRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
