// =====================================================
// FILE: 19_PM_Report.gs
// =====================================================

function getPMOrderData(row, woNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!shWO) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

  const woValues = shWO.getDataRange().getValues();
  if (woValues.length < 2) throw new Error("No hay órdenes.");

  const woHeaders = woValues[0].map(h => String(h).trim());
  const rowNumber = resolvePMOrderDataRow_(woValues, woHeaders, row, woNumber);
  if (!rowNumber) throw new Error("No se encontró la orden PM: " + (woNumber || row || ""));

  const dataRow = woValues[rowNumber - 1];

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
    row: rowNumber,
    woNumber: obj.WO_NUMBER || "",
    woType: obj.WO_TYPE || "",
    pmFrequency: normalizePMReportFrequency_(obj.PM_TYPE || obj.PM_FREQUENCY || obj.TIPO_MANTENIMIENTO || ""),
    companyId,
    nsn,
    storeAddress: address,
    storeName: store.name || store.STORE_NAME || store.NAME || store.storeName || "",
    technician: obj.TECHNICIAN || "",
    equipment: obj.REPORTED_EQUIPMENT || obj.REPORTED_EQUIPMENT_EN || obj["REPORTED EQUIPMENT"] || "",
    problem: obj.REPORTED_PROBLEM_ES || obj.REPORTED_PROBLEM_EN || obj.REPORTED_PROBLEM_ORIGINAL || obj["REPORTED PROBLEM"] || "",
    status: obj.STATUS || "",
    pmReportEsUrl: obj.PM_REPORT_ES_URL || "",
    pmReportEnUrl: obj.PM_REPORT_EN_URL || "",
    pmReportStatus: obj.PM_REPORT_STATUS || ""
  };
}

function resolvePMOrderDataRow_(values, headers, row, woNumber) {
  const rowNumber = Number(row || 0);
  const targetWO = String(woNumber || "").trim();

  if (rowNumber >= 2 && rowNumber <= values.length) {
    if (!targetWO) return rowNumber;

    const idxWOForRow = headers.indexOf("WO_NUMBER");
    const rowWO = idxWOForRow >= 0
      ? String(values[rowNumber - 1][idxWOForRow] || "").trim()
      : "";

    if (rowWO === targetWO) return rowNumber;
  }

  if (!targetWO) return 0;

  const idxWO = headers.indexOf("WO_NUMBER");
  if (idxWO === -1) return 0;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxWO] || "").trim() === targetWO) {
      return i + 1;
    }
  }

  return 0;
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

  const folderUrl = mainFolder.getUrl();

  const result = {
    pdfEsId: pdfEs.id,
    pdfEsUrl: pdfEs.url,
    pdfEnId: pdfEn.id,
    pdfEnUrl: pdfEn.url,
    savedToWorkOrder: false
  };

  result.savedToWorkOrder = updatePMReportLinks_(
    payload,
    Object.assign({ folderUrl: folderUrl }, result)
  );

  return result;
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

  const equipmentQty = normalizePMEquipmentQty_(p.equipmentQty);
  const rtus = p.rtus || [];

  for (let i = equipmentQty + 1; i <= 5; i++) {
    removeRTUSection_(body, i, lang);
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

  cleanupRemainingPMMarkers_(body);
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
  const doc = DocumentApp.openById(docId);
  const tabId = getPMFirstTabId_(doc);
  doc.saveAndClose();
  Utilities.sleep(1500);

  const pdfBlob = exportPMPdfBlob_(docId, tabId, baseName + ".pdf");

  const pdfFile = folder.createFile(pdfBlob);

  return {
    id: pdfFile.getId(),
    url: pdfFile.getUrl()
  };
}

function getPMFirstTabId_(doc) {
  if (!doc || typeof doc.getTabs !== "function") return "";

  const tabs = doc.getTabs();
  if (!tabs || !tabs.length) return "";

  return tabs[0].getId();
}

function normalizePMEquipmentQty_(value) {
  const qty = Number(value || 5);
  if (!isFinite(qty)) return 5;

  return Math.max(1, Math.min(5, Math.floor(qty)));
}

function exportPMPdfBlob_(docId, tabId, fileName) {
  if (!tabId) {
    return DriveApp
      .getFileById(docId)
      .getBlob()
      .getAs(MimeType.PDF)
      .setName(fileName);
  }

  const url =
    "https://docs.google.com/document/d/" +
    encodeURIComponent(docId) +
    "/export?format=pdf&tab=" +
    encodeURIComponent(tabId);

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("No se pudo exportar el PDF del reporte PM. HTTP " + code);
  }

  return response.getBlob().setName(fileName);
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

  if (!urls || !urls.length) {
    body.insertParagraph(idx + 1, "(Sin archivos)");
    return;
  }

  const table = body.insertTable(idx + 1, []);
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
        normalizePMImageSize_(img);
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

function normalizePMImageSize_(img) {
  const maxWidth = 180;
  const maxHeight = 150;
  const originalWidth = Number(img.getWidth() || 0);
  const originalHeight = Number(img.getHeight() || 0);

  if (!originalWidth || !originalHeight) {
    img.setWidth(maxWidth);
    img.setHeight(maxHeight);
    return;
  }

  const ratio = Math.min(maxWidth / originalWidth, maxHeight / originalHeight);
  const width = Math.max(1, Math.round(originalWidth * ratio));
  const height = Math.max(1, Math.round(originalHeight * ratio));

  img.setWidth(width);
  img.setHeight(height);
}

function cleanupRemainingPMMarkers_(body) {
  body.replaceText("\\{\\{[^}]+\\}\\}", "");
  body.replaceText("\\[\\[[^\\]]+\\]\\]", "");
}

function updatePMReportLinks_(payload, report) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
    if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

    let headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });

    [
      "PM_REPORT_ES_URL",
      "PM_REPORT_EN_URL",
      "PM_REPORT_FOLDER_URL",
      "PM_REPORT_DATE",
      "PM_REPORT_STATUS"
    ].forEach(function(h) {
      ensurePMReportColumn_(sh, headers, h);
      headers = sh.getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map(function(x) {
          return String(x).trim();
        });
    });

    const rowNumber = resolvePMReportRow_(sh, headers, payload.row, payload.woNumber);
    if (!rowNumber) {
      throw new Error("No se pudo encontrar la orden PM para guardar links: " + (payload.woNumber || ""));
    }

    const oldPdfUrls = [
      getPMReportCell_(sh, rowNumber, headers, "PM_REPORT_ES_URL"),
      getPMReportCell_(sh, rowNumber, headers, "PM_REPORT_EN_URL")
    ];

    setPMReportCell_(sh, rowNumber, headers, "PM_REPORT_ES_URL", report.pdfEsUrl || "");
    setPMReportCell_(sh, rowNumber, headers, "PM_REPORT_EN_URL", report.pdfEnUrl || "");
    setPMReportCell_(sh, rowNumber, headers, "PM_REPORT_FOLDER_URL", report.folderUrl || "");
    setPMReportCell_(sh, rowNumber, headers, "PM_REPORT_DATE", new Date());
    setPMReportCell_(sh, rowNumber, headers, "PM_REPORT_STATUS", "GENERATED");

    trashOldPMReportPdfs_(oldPdfUrls, [
      report.pdfEsUrl,
      report.pdfEnUrl
    ]);

    try {
      const companyId = getPMReportCell_(sh, rowNumber, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
      addLog_(companyId, payload.woNumber || "", "PM REPORT GENERATED", "", "GENERATED", payload.technician || "PM Report", report.folderUrl || "");
    } catch (logErr) {
      Logger.log("PM report log error: " + logErr);
    }

    return true;
  } catch (err) {
    Logger.log("ERROR updatePMReportLinks_: " + err);
    return false;
  }
}

function trashOldPMReportPdfs_(oldUrls, newUrls) {
  const keepIds = {};

  (newUrls || []).forEach(function(url) {
    try {
      const id = extractDriveIdPM_(url);
      if (id) keepIds[id] = true;
    } catch (err) {
      Logger.log("PM report keep-id parse error: " + err);
    }
  });

  (oldUrls || []).forEach(function(url) {
    try {
      if (!url) return;

      const id = extractDriveIdPM_(url);
      if (!id || keepIds[id]) return;

      DriveApp.getFileById(id).setTrashed(true);
    } catch (err) {
      Logger.log("ERROR trashOldPMReportPdfs_: " + err);
    }
  });
}

function ensurePMReportColumn_(sheet, headers, columnName) {
  const target = String(columnName || "").trim();
  const exists = headers.some(function(h) {
    return String(h || "").trim().toUpperCase() === target.toUpperCase();
  });

  if (!exists) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(target);
  }
}

function resolvePMReportRow_(sheet, headers, rowNumber, woNumber) {
  const row = Number(rowNumber || 0);
  const targetWO = String(woNumber || "").trim();
  const lastRow = sheet.getLastRow();

  if (row >= 2 && row <= lastRow) {
    if (!targetWO) return row;

    const rowWO = String(getPMReportCell_(sheet, row, headers, "WO_NUMBER") || "").trim();
    if (rowWO === targetWO) return row;
  }

  if (!targetWO) return 0;

  const idxWO = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf("WO_NUMBER");

  if (idxWO === -1 || lastRow < 2) return 0;

  const values = sheet.getRange(2, idxWO + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === targetWO) {
      return i + 2;
    }
  }

  return 0;
}

function setPMReportCell_(sheet, rowNumber, headers, headerName, value) {
  const idx = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf(String(headerName || "").trim().toUpperCase());

  if (idx === -1) throw new Error("No existe la columna " + headerName);
  sheet.getRange(Number(rowNumber), idx + 1).setValue(value);
}

function getPMReportCell_(sheet, rowNumber, headers, headerName) {
  const idx = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf(String(headerName || "").trim().toUpperCase());

  if (idx === -1) return "";
  return sheet.getRange(Number(rowNumber), idx + 1).getValue();
}

function normalizePMReportFrequency_(value) {
  const v = String(value || "").trim().toUpperCase();

  if (v === "MONTHLY" || v === "MENSUAL") return "MENSUAL";
  if (v === "QUARTERLY" || v === "TRIMESTRAL") return "TRIMESTRAL";
  if (v === "ANNUAL" || v === "ANUAL") return "ANNUAL";

  return v || "MENSUAL";
}

function removeRTUSection_(body, rtuNum, lang) {
  const isEn = String(lang || "ES").toUpperCase() === "EN";
  const startTokens = [
    "{{TITLE_RTU" + rtuNum + "}}",
    "RTU " + rtuNum
  ];
  const nextTokens = rtuNum < 5
    ? [
        "{{TITLE_RTU" + (rtuNum + 1) + "}}",
        "RTU " + (rtuNum + 1)
      ]
    : [
        "{{TITLE_ROOF_RECOMMENDATIONS}}",
        isEn ? "ROOF RECOMMENDATIONS" : "RECOMENDACIONES AREA ROOF"
      ];

  let startIndex = -1;
  let endIndex = -1;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const txt = child.getText ? child.getText() : "";

    if (startIndex === -1 && textContainsAnyPM_(txt, startTokens)) {
      startIndex = i;
      continue;
    }

    if (startIndex !== -1 && textContainsAnyPM_(txt, nextTokens)) {
      endIndex = i;
      break;
    }
  }

  if (startIndex === -1 || endIndex === -1) return;

  for (let i = endIndex - 1; i >= startIndex; i--) {
    body.removeChild(body.getChild(i));
  }
}

function textContainsAnyPM_(text, tokens) {
  const value = String(text || "").toUpperCase();

  return tokens.some(function(token) {
    return value.indexOf(String(token || "").toUpperCase()) !== -1;
  });
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
