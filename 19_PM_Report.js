// =====================================================
// FILE: 19_PM_Report.gs
// =====================================================

function getPMOrderData(row) {
  row = Number(row);
  if (!row || row < 2) throw new Error("Fila inválida.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!shWO) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

  const woValues = shWO.getDataRange().getValues();
  if (woValues.length < 2) throw new Error("No hay órdenes.");

  const woHeaders = woValues[0].map(h => String(h).trim());
  const dataRow = woValues[row - 1];

  const obj = {};
  woHeaders.forEach(function(h, i) {
    let value = dataRow[i];
    if (value instanceof Date) {
      value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy");
    }
    obj[h] = value;
  });

  const nsn = String(obj.NSN || "").trim();
  const companyId = String(obj.COMPANY_ID || "").trim().toUpperCase();
  const store = getStoreByNSN_(nsn, companyId) || {};

  const address =
    store.address ||
    store.ADDRESS ||
    store.STORE_ADDRESS ||
    store.Store_Address ||
    store.storeAddress ||
    store.direccion ||
    store.DIRECCION ||
    obj.STORE_ADDRESS ||
    obj.ADDRESS ||
    "";

  return {
    row,
    woNumber: obj.WO_NUMBER || "",
    woType: obj.WO_TYPE || "",
    companyId,
    nsn,
    storeAddress: address,
    storeName: store.name || store.STORE_NAME || store.NAME || store.storeName || "",
    technician: obj.TECHNICIAN || "",
    equipment: obj.REPORTED_EQUIPMENT || obj.REPORTED_EQUIPMENT_EN || obj["REPORTED EQUIPMENT"] || "",
    problem: obj.REPORTED_PROBLEM_ES || obj.REPORTED_PROBLEM_EN || obj.REPORTED_PROBLEM_ORIGINAL || obj["REPORTED PROBLEM"] || "",
    status: obj.STATUS || ""
  };
}

function getPMTexts() {
  return PM_TEXTS;
}

function uploadPMFile(fileObj) {
  if (!fileObj || !fileObj.data) throw new Error("Archivo inválido.");
  if (!fileObj.folderId) throw new Error("Falta folderId para subir archivo.");

  const folder = DriveApp.getFolderById(fileObj.folderId);
  const bytes = Utilities.base64Decode(fileObj.data);

  const blob = Utilities.newBlob(
    bytes,
    fileObj.mimeType || "application/octet-stream",
    fileObj.name || "archivo"
  );

  const file = folder.createFile(blob);

  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName(),
    mimeType: file.getMimeType()
  };
}

function createPMReportFolder(data) {
  const root = DriveApp.getFolderById(CFG.PM_REPORTS_FOLDER_ID);
  const freq = String(data.pmFrequency || "MENSUAL").toUpperCase();
  const store = safeName_(data.storeNumber || "NO_STORE");
  const wo = safeName_(data.woNumber || "NO_WO");

  const freqFolder = getOrCreateFolder_(root, freq);
  const storeFolder = getOrCreateFolder_(freqFolder, "Store " + store);
  const woFolder = getOrCreateFolder_(storeFolder, wo);

  const pdfFolder = getOrCreateFolder_(woFolder, "PDF");
  const filesFolder = getOrCreateFolder_(woFolder, "ARCHIVOS");

  return {
    id: woFolder.getId(),
    url: woFolder.getUrl(),
    pdfFolderId: pdfFolder.getId(),
    filesFolderId: filesFolder.getId()
  };
}

function generatePMReportPDF(payload) {
  if (!payload) throw new Error("No llegó payload.");
  if (!payload.folderId) throw new Error("Falta folderId.");
  if (!CFG.PM_REPORT_TEMPLATE_ID) throw new Error("Falta CFG.PM_REPORT_TEMPLATE_ID.");

  const pdfFolder = DriveApp.getFolderById(payload.pdfFolderId || payload.folderId);
  const mainFolder = DriveApp.getFolderById(payload.folderId);

  const baseName =
    "PM_Report_" +
    safeName_(payload.storeNumber || "NO_STORE") +
    "_" +
    safeName_(payload.woNumber || "NO_WO");

  const docFileEs = DriveApp
    .getFileById(CFG.PM_REPORT_TEMPLATE_ID)
    .makeCopy(baseName + "_ES", pdfFolder);

  fillPMTemplate_(docFileEs.getId(), payload, "ES");

  const pdfEs = exportPMPdf_(docFileEs.getId(), pdfFolder, baseName + "_ES");

  const docFileEn = DriveApp
    .getFileById(CFG.PM_REPORT_TEMPLATE_ID)
    .makeCopy(baseName + "_EN", pdfFolder);

  fillPMTemplate_(docFileEn.getId(), payload, "EN");

  const pdfEn = exportPMPdf_(docFileEn.getId(), pdfFolder, baseName + "_EN");

  DriveApp.getFileById(docFileEs.getId()).setTrashed(true);
  DriveApp.getFileById(docFileEn.getId()).setTrashed(true);

  return {
    pdfEsId: pdfEs.id,
    pdfEsUrl: pdfEs.url,
    pdfEnId: pdfEn.id,
    pdfEnUrl: pdfEn.url,
    folderUrl: mainFolder.getUrl()
  };
}

function fillPMTemplate_(docId, p, lang) {
  lang = String(lang || "ES").toUpperCase();

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  replacePMTitles_(body, lang);

  const serviceType = lang === "EN" ? "PREVENTIVE (PM)" : p.serviceType;
  const pmFrequency = formatPMFrequency_(p.pmFrequency, lang);

  replacePM_(body, "{{WO_NUMBER}}", p.woNumber);
  replacePM_(body, "{{STORE}}", p.storeNumber);
  replacePM_(body, "{{ADDRESS}}", p.address);
  replacePM_(body, "{{TECHNICIAN}}", p.technician);
  replacePM_(body, "{{SERVICE_DATE}}", p.serviceDate);
  replacePM_(body, "{{SERVICE_TYPE}}", serviceType);
  replacePM_(body, "{{PM_FREQUENCY}}", pmFrequency);

  replacePM_(body, "{{ROOF_RECOM}}",
    lang === "EN" ? safeTranslatePMToEn_(p.roofRecommendations) : p.roofRecommendations
  );

  replacePM_(body, "{{FIRMA_TECNICO_TEXTO}}", p.techSignature);
  replacePM_(body, "{{FIRMA_CLIENTE_TEXTO}}", p.clientSignature);

  const equipmentQty = Number(p.equipmentQty || 5);
  const rtus = p.rtus || [];

  for (let i = 5; i > equipmentQty; i--) {
    removeRTUSection_(body, i);
  }

  for (let i = 1; i <= equipmentQty; i++) {
    const rtu = rtus.find(r => Number(r.number) === i) || {};

    replacePM_(body, "{{TITLE_RTU" + i + "}}", "🔹 RTU " + i);

    replacePM_(body, "{{TITLE_RTU" + i + "_PHOTOS}}",
      lang === "EN" ? "RTU " + i + " PHOTOS:" : "FOTOS RTU " + i + ":"
    );

    replacePM_(body, "{{TITLE_RTU" + i + "_PROBLEM_PHOTOS}}",
      lang === "EN" ? "RTU " + i + " ISSUE PHOTOS:" : "FOTOS DE PROBLEMAS ENCONTRADOS:"
    );

    replacePM_(body, "{{RTU" + i + "_WORK}}",
      getRTUWorkByLang_(rtu.work || "", p.pmFrequency, lang)
    );

    replacePM_(body, "{{RTU" + i + "_ISSUES}}",
      lang === "EN" ? safeTranslatePMToEn_(rtu.issues || "") : rtu.issues || ""
    );

    insertPMFilesAtMarker_(
      body,
      "[[FOTOS_RTU" + i + "]]",
      rtu.photos || [],
      lang === "EN" ? "RTU " + i + " PHOTOS" : "FOTOS RTU " + i
    );

    insertPMFilesAtMarker_(
      body,
      "[[FOTOS_RTU" + i + "_PROB]]",
      rtu.problemPhotos || [],
      lang === "EN" ? "RTU " + i + " ISSUE PHOTOS" : "FOTOS PROBLEMAS RTU " + i
    );
  }

  insertPMFilesAtMarker_(
    body,
    "[[TEMP_COCINA_FOTOS]]",
    p.kitchenTemps || [],
    lang === "EN" ? "KITCHEN TEMPERATURES" : "TEMPERATURAS COCINA"
  );

  insertPMFilesAtMarker_(
    body,
    "[[TEMP_DINNER_FOTOS]]",
    p.dinnerTemps || [],
    lang === "EN" ? "DINING AREA TEMPERATURES" : "TEMPERATURAS DINNER"
  );

  doc.saveAndClose();
}

function replacePMTitles_(body, lang) {
  const en = lang === "EN";

  const map = {
    "{{TITLE_GENERAL_INFO}}": en ? "🔹 GENERAL INFORMATION" : "🔹 INFORMACIÓN GENERAL",
    "{{TITLE_STORE}}": "NATIONAL STORE:",
    "{{TITLE_ADDRESS}}": en ? "ADDRESS:" : "DIRECCION:",
    "{{TITLE_SERVICE_DATE}}": en ? "SERVICE DATE:" : "FECHA DEL SERVICIO:",
    "{{TITLE_SERVICE_TYPE}}": en ? "SERVICE TYPE:" : "TIPO DE SERVICIO:",
    "{{TITLE_PM_FREQUENCY}}": en ? "PM FREQUENCY:" : "FRECUENCIA PM:",
    "{{TITLE_ROOF_AREA}}": en ? "🏗️ ROOF AREA" : "🏗️ AREA ROOF",
    "{{TITLE_HVAC_WORK}}": en ? "HVAC WORK" : "TRABAJOS EN HVAC",
    "{{TITLE_WORK_DONE}}": en ? "WORK PERFORMED:" : "TRABAJOS REALIZADOS:",
    "{{TITLE_ISSUES_FOUND}}": en ? "ISSUES FOUND:" : "PROBLEMAS ENCONTRADOS:",
    "{{TITLE_ROOF_RECOMMENDATIONS}}": en ? "📝 ROOF RECOMMENDATIONS" : "📝 RECOMENDACIONES AREA ROOF",
    "{{TITLE_TEMPERATURES}}": en ? "🌡️ TEMPERATURE MEASUREMENTS" : "🌡️ MEDICIÓN DE TEMPERATURAS",
    "{{TITLE_KITCHEN}}": en ? "KITCHEN:" : "COCINA:",
    "{{TITLE_DINING}}": en ? "DINING AREA:" : "DINNER:",
    "{{TITLE_FINALIZATION}}": en ? "✍️ FINALIZATION" : "✍️ FINALIZACION",
    "{{TITLE_TECH_SIGNATURE}}": en ? "TECHNICIAN SIGNATURE:" : "FIRMA DEL TECNICO:",
    "{{TITLE_CLIENT_SIGNATURE}}": en ? "CLIENT / MANAGER SIGNATURE:" : "FIRMA DEL CLIENTE / MANAGER:",
    "{{TITLE_COMPANY}}": "🏢 COMPANY"
  };

  Object.keys(map).forEach(function(k) {
    replacePM_(body, k, map[k]);
  });
}

function exportPMPdf_(docId, folder, baseName) {
  DocumentApp.openById(docId).saveAndClose();
  Utilities.sleep(1500);

  const pdfBlob = DriveApp
    .getFileById(docId)
    .getBlob()
    .getAs(MimeType.PDF)
    .setName(baseName + ".pdf");

  const pdfFile = folder.createFile(pdfBlob);

  return {
    id: pdfFile.getId(),
    url: pdfFile.getUrl()
  };
}

function getRTUWorkByLang_(originalText, pmFrequency, lang) {
  if (String(lang || "ES").toUpperCase() === "ES") return originalText || "";

  const freq = String(pmFrequency || "MENSUAL").toUpperCase();

  if (typeof PM_TEXTS !== "undefined" && PM_TEXTS[freq] && PM_TEXTS[freq].EN) {
    const esDefault = String(PM_TEXTS[freq].ES || "").trim();
    const original = String(originalText || "").trim();

    if (!original || original === esDefault) {
      return PM_TEXTS[freq].EN;
    }
  }

  return safeTranslatePMToEn_(originalText);
}

function formatPMFrequency_(freq, lang) {
  freq = String(freq || "MENSUAL").toUpperCase();
  lang = String(lang || "ES").toUpperCase();

  if (lang === "ES") {
    if (freq === "MENSUAL") return "Mensual";
    if (freq === "TRIMESTRAL") return "Trimestral";
    if (freq === "ANNUAL") return "Anual";
    return freq;
  }

  if (freq === "MENSUAL") return "Monthly";
  if (freq === "TRIMESTRAL") return "Quarterly";
  if (freq === "ANNUAL") return "Annual";
  return freq;
}

function safeTranslatePMToEn_(text) {
  if (!text) return "";

  const t = String(text).trim();
  if (!t) return "";

  if (/^\d+(\.\d+)?$/.test(t)) return t;
  if (/^[A-Za-z0-9\-_.\/#]+$/.test(t) && t.length <= 35) return t;

  try {
    const out = LanguageApp.translate(t, "es", "en");
    if (!out || !out.trim()) return t;
    return out;
  } catch (e) {
    Logger.log("PM translate error: " + e);
    return t;
  }
}

function replacePM_(body, marker, value) {
  body.replaceText(
    marker.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
    String(value || "")
  );
}

function insertPMFilesAtMarker_(body, marker, urls, title) {
  const found = body.findText(marker.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  if (!found) return;

  const el = found.getElement();
  const par = el.getParent().asParagraph();
  const idx = body.getChildIndex(par);

  par.setText(" ");
  body.insertParagraph(idx + 1, title || "Archivos").setBold(true);

  if (!urls || !urls.length) {
    body.insertParagraph(idx + 2, "(Sin archivos)");
    return;
  }

  const table = body.insertTable(idx + 2, []);
  table.setBorderWidth(0);

  let row = table.appendTableRow();
  let col = 0;

  urls.forEach(function(url) {
    if (col === 2) {
      row = table.appendTableRow();
      col = 0;
    }

    const cell = row.appendTableCell(" ");
    cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(4);

    try {
      const fileId = extractDriveIdPM_(url);
      const file = DriveApp.getFileById(fileId);
      const mime = String(file.getMimeType() || "").toLowerCase();

      if (mime.indexOf("image/") === 0) {
        const img = cell.appendParagraph(" ").appendInlineImage(file.getBlob());
        img.setWidth(250);
        img.setLinkUrl(file.getUrl());
      } else {
        const link = cell.appendParagraph("🎬 Archivo / Video");
        link.setLinkUrl(file.getUrl());
      }
    } catch (e) {
      cell.appendParagraph("Archivo: " + url);
    }

    col++;
  });

  if (col === 1) row.appendTableCell(" ");
}

function removeRTUSection_(body, rtuNum) {
  const startText = "🔹 RTU " + rtuNum;
  const nextText = rtuNum < 5 ? "🔹 RTU " + (rtuNum + 1) : "📝 RECOMENDACIONES AREA ROOF";

  let startIndex = -1;
  let endIndex = -1;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const txt = child.getText ? child.getText() : "";

    if (startIndex === -1 && txt.indexOf(startText) !== -1) {
      startIndex = i;
      continue;
    }

    if (startIndex !== -1 && txt.indexOf(nextText) !== -1) {
      endIndex = i;
      break;
    }
  }

  if (startIndex === -1 || endIndex === -1) return;

  for (let i = endIndex - 1; i >= startIndex; i--) {
    body.removeChild(body.getChild(i));
  }
}

function extractDriveIdPM_(urlOrId) {
  const s = String(urlOrId || "").trim();

  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;

  let m = s.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];

  m = s.match(/[-\w]{25,}/);
  if (m && m[0]) return m[0];

  throw new Error("No se pudo extraer ID Drive: " + s);
}

function safeName_(s) {
  return String(s || "")
    .trim()
    .replace(/[\\/:*?"<>|#%{}[\]^~`]/g, "-") || "SIN_NOMBRE";
}