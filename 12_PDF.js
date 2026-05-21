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

  doc.saveAndClose();

  Utilities.sleep(3000);

  let pdfFile = null;
  let lastErr = "";

  for (let i = 0; i < 5; i++) {
    try {
      const blob = docCopy.getAs(MimeType.PDF).setName(pdfName);
      pdfFile = targetFolder.createFile(blob);
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

function testGeneratePdfFromCloseOrder() {
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