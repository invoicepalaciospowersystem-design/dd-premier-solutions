// =====================================================
// FILE: 03_Orders.gs
// =====================================================

function saveWorkOrderRow_(o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  if (!sh) {
    throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  const rowObj = {
    COMPANY_ID: o.companyId || CFG.DEFAULT_COMPANY_ID,
    WO_NUMBER: o.woNumber,
    DATE_CREATED: o.dateCreated,
    FORM_LANGUAGE: o.formLang,
    CLIENT: o.client,

    WO_TYPE: o.woType || "REPAIR_FORM",
    PM_TYPE: o.pmType || "",

    NSN: o.nsn,
    STATUS: o.status,

    TECHNICIAN: "",
    TECH_EMAIL: "",
    EMAIL_SENT: "NO",
    EMAIL_SENT_DATE: "",
    DATE_SENT_TO_TECH: "",

    REPORTED_EQUIPMENT: o.equipment,
    REPORTED_EQUIPMENT_EN: o.equipmentEn,

    ORDER_PRIORITY: o.priorityOriginal,

    REPORTED_PROBLEM_ORIGINAL: o.problemOriginal,
    REPORTED_PROBLEM_EN: o.problemEn,
    REPORTED_PROBLEM_ES: o.problemEs,

    MANAGER_NAME: o.manager,
    MANAGER_EMAIL: o.managerEmail,

    PDF_EN_URL: o.pdfEnUrl,
    PDF_ES_URL: o.pdfEsUrl,
    FOLDER_URL: o.folderUrl,

    NOTES: "",
    DATE_COMPLETED: "",
    DATE_CLOSED: "",

    QUOTE_EN_URL: "",
    QUOTE_ES_URL: "",
    QUOTE_STATUS: ""
  };

  const rowValues = headers.map(function(h) {
    return rowObj[h] !== undefined ? rowObj[h] : "";
  });

  sh.appendRow(rowValues);
  touchAppCacheVersion_();
}

function getDashboardData(companyId, role, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  role = String(role || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES", "ECONOMIA"], companyId);

  return withAppCache_(["orders-dashboard", companyId, role], 90, function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  if (!sh) {
    throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers = values[0].map(String);
  const idxCompany = headers.indexOf("COMPANY_ID");

  if (idxCompany === -1) {
    throw new Error("WORK_ORDERS debe tener la columna COMPANY_ID.");
  }

  const data = [];
  const storeMap = getStoreMapByCompany_(companyId);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = String(row[idxCompany] || "").trim().toUpperCase();

    if (rowCompany !== companyId) continue;

    const obj = {};

    headers.forEach(function(h, c) {
      let value = row[c];

      if (value instanceof Date) {
        value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a");
      }

      obj[h] = value;
    });

    obj.ROW_NUMBER = i + 1;

    enrichWorkOrderObjectFromStoreMap_(obj, storeMap);

    data.push(obj);
  }

  return data.reverse();
  });
}

function updateWorkOrderFromDashboard(rowNumber, updates, sessionToken) {
  if (!rowNumber || isNaN(rowNumber)) {
    throw new Error("Fila inválida recibida: " + rowNumber);
  }

  if (!updates) {
    throw new Error("No llegaron cambios desde el dashboard.");
  }

  rowNumber = Number(rowNumber);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(String);

  const companyId = getCellByHeader_(sh, rowNumber, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
  const woNumber = getCellByHeader_(sh, rowNumber, headers, "WO_NUMBER");
  const oldStatus = getCellByHeader_(sh, rowNumber, headers, "STATUS");
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);

  Object.keys(updates).forEach(function(key) {
    setCellByHeader_(sh, rowNumber, headers, key, updates[key]);
  });

  if (updates.STATUS) {

    if (updates.STATUS === "COMPLETED") {
      setCellByHeader_(sh, rowNumber, headers, "DATE_COMPLETED", new Date());

      createEconomyFromWO(companyId, woNumber);

      addNotification_(companyId, "David Dominguez", woNumber, "DONE", "✅ Completed " + woNumber);

      addNotification_(companyId, "Sarahi", woNumber, "BILLING", "💰 Ready for billing " + woNumber);
    }

    if (updates.STATUS === "CLOSED") {
      setCellByHeader_(sh, rowNumber, headers, "DATE_CLOSED", new Date());
    }

    if (updates.STATUS === "REQUEST PARTS") {
      addNotification_(companyId, "Dayre", woNumber, "PARTS", "🔧 Parts requested for " + woNumber);
    }

    if (updates.STATUS === "PARTS IN TRANSIT") {
      addNotification_(companyId, "Dayre", woNumber, "TRANSIT", "🚚 Parts in transit for " + woNumber);
    }

    addLog_(
      companyId,
      woNumber,
      "STATUS UPDATED FROM DASHBOARD",
      oldStatus,
      updates.STATUS,
      getSessionActorLabel_(session),
      ""
    );
  }

  addAuditLog_("ORDERS", "WORK_ORDER_UPDATED", companyId, "WORK_ORDER", woNumber, session, {
    rowNumber: rowNumber,
    updates: updates
  });

  touchAppCacheVersion_();
  return true;
}

function sendWorkOrderFromDashboard(rowNumber, sessionToken) {
  if (!rowNumber || isNaN(rowNumber)) {
    throw new Error("No llegó número de fila desde el dashboard. Valor recibido: " + rowNumber);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const companyId = getCellByHeader_(sh, Number(rowNumber), headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);
  const woNumber = getCellByHeader_(sh, Number(rowNumber), headers, "WO_NUMBER");

  dispatchRowToTech_(sh, Number(rowNumber));

  addAuditLog_("ORDERS", "WORK_ORDER_SENT_TO_TECH", companyId, "WORK_ORDER", woNumber, session, {
    rowNumber: Number(rowNumber)
  });

  touchAppCacheVersion_();
  return true;
}

function deleteWorkOrder(rowNumber, sessionToken) {
  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < 2) throw new Error("Fila invalida.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

  let headers = ensureSheetColumns_(sh, ["ACTIVE", "DELETED_AT", "DELETED_BY"]);
  const companyId = getCellByHeader_(sh, rowNumber, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
  const woNumber = getCellByHeader_(sh, rowNumber, headers, "WO_NUMBER");
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);

  setCellByHeader_(sh, rowNumber, headers, "ACTIVE", "NO");
  setCellByHeader_(sh, rowNumber, headers, "DELETED_AT", new Date());
  setCellByHeader_(sh, rowNumber, headers, "DELETED_BY", getSessionActorLabel_(session));

  const idxStatus = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  }).indexOf("STATUS");

  if (idxStatus >= 0) {
    setCellByHeader_(sh, rowNumber, headers, "STATUS", "DELETED");
  }

  addAuditLog_("ORDERS", "WORK_ORDER_SOFT_DELETED", companyId, "WORK_ORDER", woNumber, session, {
    rowNumber: rowNumber
  });

  touchAppCacheVersion_();
  return true;
}

function dispatchRowToTech_(sh, row) {
  if (!row || isNaN(row)) throw new Error("Fila inválida: " + row);
  row = Number(row);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rowData = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const obj = rowToObject_(headers, rowData);
  enrichWorkOrderObject_(obj);

  const companyId = obj.COMPANY_ID || CFG.DEFAULT_COMPANY_ID;

  if (!obj.WO_NUMBER) throw new Error("No se encontró WO_NUMBER en la fila: " + row);

  const technician = String(obj.TECHNICIAN || "").trim();
  if (!technician) throw new Error("Debes seleccionar al menos un técnico antes de enviar.");

  const techEmail = getTechEmails_(technician);

  const publicBaseUrl = getWebAppBaseUrl_(companyId);
  const attachmentPortalUrl = obj.FOLDER_URL
    ? publicBaseUrl +
      "?view=orderFiles" +
      "&row=" + encodeURIComponent(String(row)) +
      "&wo=" + encodeURIComponent(String(obj.WO_NUMBER || "")) +
      "&companyId=" + encodeURIComponent(String(companyId || ""))
    : "";

  if (isTechEmailTestMode_()) {
    Logger.log("TEST MODE: Email NO enviado. Técnicos: " + technician + " / " + techEmail);
  } else {
    MailApp.sendEmail({
      to: techEmail,
      subject: "🚨 WORK ORDER " + obj.WO_NUMBER + " / NUEVA ORDEN DE TRABAJO",
      htmlBody:
        "<h2>🚨 NEW WORK ORDER / NUEVA ORDEN DE TRABAJO</h2>" +
        "<p><b>WO#:</b> " + obj.WO_NUMBER + "</p>" +
        "<p><b>Type:</b> " + (obj.WO_TYPE || "REPAIR_FORM") + "</p>" +
        "<p><b>PM Type:</b> " + (obj.PM_TYPE || "") + "</p>" +
        "<p><b>Client:</b> " + obj.CLIENT + "</p>" +
        "<p><b>NSN #:</b> " + obj.NSN + "</p>" +
        "<p><b>Address:</b> " + (obj.STORE_ADDRESS || "") + "</p>" +
        "<p><b>Equipment:</b> " + (obj.REPORTED_EQUIPMENT_EN || obj.REPORTED_EQUIPMENT || "") + "</p>" +
        "<p><b>Priority:</b> " + obj.ORDER_PRIORITY + "</p>" +
        "<p><b>Issue:</b><br>" + obj.REPORTED_PROBLEM_EN + "</p>" +
        (attachmentPortalUrl
          ? "<p><b>Photos / Videos:</b> <a href='" +
            escapeHtmlForEmail_(attachmentPortalUrl) +
            "'>Open photos and videos in the work order system</a></p>"
          : "") +
        "<hr>" +
        "<p><b>Dirección:</b> " + (obj.STORE_ADDRESS || "") + "</p>" +
        "<p><b>Problema:</b><br>" + obj.REPORTED_PROBLEM_ES + "</p>" +
        "<hr>" +
        "<p><b>Portal tecnico:</b> <a href='" + publicBaseUrl + "'>" + publicBaseUrl + "</a></p>" +
        "<p>Actualice los estados solamente desde la Web App / Update statuses only from the Web App.</p>" +
        "<p>Order details are included in this email / Los detalles estan incluidos en este correo.</p>"
    });
  }

  setCellByHeader_(sh, row, headers, "STATUS", "SENT TO TECH");
  setCellByHeader_(sh, row, headers, "EMAIL_SENT", "YES");
  setCellByHeader_(sh, row, headers, "EMAIL_SENT_DATE", new Date());
  setCellByHeader_(sh, row, headers, "TECH_EMAIL", techEmail);
  setCellByHeader_(sh, row, headers, "DATE_SENT_TO_TECH", new Date());

  getNamesFromText_(technician).forEach(function(techName) {
    addNotification_(
      companyId,
      techName,
      obj.WO_NUMBER,
      "ASSIGNED",
      "🚨 Assigned Work Order " + obj.WO_NUMBER + " / NSN " + obj.NSN
    );
  });

  addLog_(
    companyId,
    obj.WO_NUMBER,
    String(obj.EMAIL_SENT || "").toUpperCase() === "YES" ? "EMAIL RESENT TO TECH" : "EMAIL SENT TO TECH",
    obj.STATUS || "",
    "SENT TO TECH",
    Session.getActiveUser().getEmail() || "Dashboard",
    technician + " / " + techEmail
  );
}

function isTechEmailTestMode_() {
  if (CFG.TECH_EMAIL_TEST_MODE !== undefined) {
    return CFG.TECH_EMAIL_TEST_MODE === true;
  }

  return CFG.TEST_MODE === true;
}

function handleWorkOrderAction_(e) {
  return HtmlService.createHtmlOutput(
    "<h2>Acciones por email desactivadas</h2>" +
    "<p>Por seguridad, los estados de las ordenes se actualizan solamente desde la Web App.</p>"
  )
    .setTitle("Work Order Action")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildWorkOrderActionLink_(publicBaseUrl, companyId, woNumber, action) {
  const token = getWorkOrderActionToken_(companyId, woNumber, action);

  return publicBaseUrl +
    "?action=" + encodeURIComponent(action) +
    "&wo=" + encodeURIComponent(woNumber) +
    "&companyId=" + encodeURIComponent(companyId || CFG.DEFAULT_COMPANY_ID) +
    "&token=" + encodeURIComponent(token);
}

function verifyWorkOrderActionToken_(companyId, woNumber, action, token) {
  return constantTimeEquals_(
    String(token || ""),
    getWorkOrderActionToken_(companyId, woNumber, action)
  );
}

function getWorkOrderActionToken_(companyId, woNumber, action) {
  const secret = getWorkOrderActionSecret_();
  const payload = [
    String(companyId || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase(),
    String(woNumber || "").trim(),
    String(action || "").trim().toUpperCase()
  ].join("|");

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    secret + "|" + payload,
    Utilities.Charset.UTF_8
  );

  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, "");
}

function getWorkOrderActionSecret_() {
  const props = PropertiesService.getScriptProperties();
  const key = "WO_ACTION_SECRET";
  let secret = props.getProperty(key);

  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(key, secret);
  }

  return secret;
}

function generateWONumber_(companyId) {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();

    const year = Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy");

    const company = String(companyId || CFG.DEFAULT_COMPANY_ID)
      .trim()
      .toUpperCase();

    const key = "LAST_WO_NUMBER_" + company + "_" + year;

    let last = Number(props.getProperty(key) || 0);

    last++;

    props.setProperty(key, String(last));

    return "WO-" + year + "-" + ("0000" + last).slice(-4);

  } finally {
    lock.releaseLock();
  }
}

function createWorkOrderFromApp(data, sessionToken) {
  const companyId = String(data.COMPANY_ID || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);

  const woType = String(data.WO_TYPE || "REPAIR_FORM").trim();

  const pmType = String(data.PM_TYPE || "").trim();

  const nsn = String(data.NSN || "").trim();

  const equipmentOriginal = String(data.REPORTED_EQUIPMENT || "").trim();

  const priorityOriginal = String(data.ORDER_PRIORITY || "").trim();

  const problemOriginal = String(data.REPORTED_PROBLEM || "").trim();

  const manager = String(data.MANAGER || "").trim();

  if (!companyId) throw new Error("No llegó COMPANY_ID.");

  if (!nsn) throw new Error("NSN # es obligatorio.");

  if (!equipmentOriginal) throw new Error("Reported Equipment es obligatorio.");

  if (!priorityOriginal) throw new Error("Order Priority es obligatoria.");

  if (!problemOriginal) throw new Error("Reported Problem es obligatorio.");

  if (!manager) throw new Error("Manager Name es obligatorio.");

  if (woType === "PM_FORM" && !pmType) {
    throw new Error("PM_TYPE es obligatorio para PM_FORM.");
  }

  const store = getStoreByNSN_(nsn, companyId);

  if (!store || !store.fullAddress) {
    throw new Error(
      "No se encontró tienda/dirección para NSN: " +
      nsn +
      " en empresa " +
      companyId
    );
  }

  const woNumber = generateWONumber_(companyId);

  const equipmentEn = String(data._EQUIPMENT_EN || "").trim() || safeTranslate_(equipmentOriginal, "auto", "en");

  const problemEn = String(data._PROBLEM_EN || "").trim() || safeTranslate_(problemOriginal, "auto", "en");

  const statusEn = "ORDER RECEIVED";

  const clientName = store.client || CFG.CLIENT_DEFAULT;

  saveWorkOrderRow_({
    companyId: companyId,
    woNumber: woNumber,
    dateCreated: new Date(),
    formLang: "APP",

    client: clientName,

    woType: woType,
    pmType: pmType,

    nsn: nsn,

    equipment: equipmentOriginal,
    equipmentEn: equipmentEn,

    priorityOriginal: normalizePriorityEn_(priorityOriginal),

    problemOriginal: problemOriginal,
    problemEn: problemEn,
    problemEs: problemOriginal,

    manager: manager,
    managerEmail: "",

    status: statusEn,

    pdfEnUrl: "",
    pdfEsUrl: "",
    folderUrl: ""
  });

  try {
    addLog_(
      companyId,
      woNumber,
      "ORDER CREATED FROM APP",
      "",
      statusEn,
      getSessionActorLabel_(session),
      "Created manually from CreateOrder"
    );
  } catch (err) {
    Logger.log("LOG ERROR createWorkOrderFromApp: " + err);
    notifySystemError_("CREATE_ORDER_APP_LOG_ERROR", err, {
      module: "CREATE_ORDER",
      companyId: companyId || "",
      woNumber: woNumber || "",
      nsn: nsn || ""
    });
  }

  addNotification_(
    companyId,
    "ADMIN",
    woNumber,
    "NEW",
    "🚨 New Work Order " + woNumber + " / NSN " + nsn
  );

  addNotification_(
    companyId,
    "David Dominguez",
    woNumber,
    "NEW",
    "🚨 New Work Order " + woNumber + " / NSN " + nsn
  );

  addNotification_(
    companyId,
    "Dayre",
    woNumber,
    "NEW",
    "🚨 New Work Order " + woNumber + " / NSN " + nsn
  );

  addAuditLog_("ORDERS", "WORK_ORDER_CREATED_FROM_APP", companyId, "WORK_ORDER", woNumber, session, {
    nsn: nsn,
    woType: woType,
    pmType: pmType,
    manager: manager
  });

  return {
    success: true,
    woNumber: woNumber
  };
}

function uploadWorkOrderAttachment(data, sessionToken) {
  try {
    return uploadWorkOrderAttachment_(data, sessionToken);
  } catch (err) {
    notifySystemError_("WORK_ORDER_ATTACHMENT_UPLOAD_ERROR", err, {
      module: "CREATE_ORDER",
      companyId: data && data.companyId,
      woNumber: data && data.woNumber,
      fileName: data && data.name,
      mimeType: data && data.mimeType
    });
    throw err;
  }
}

function uploadWorkOrderAttachment_(data, sessionToken) {
  data = data || {};

  const requestedCompanyId = String(data.companyId || "").trim().toUpperCase();
  const woNumber = String(data.woNumber || "").trim();
  const uploadId = String(data.uploadId || "").trim().substring(0, 180);
  const rawBase64 = String(data.data || "").trim();
  const maxFiles = Number(CFG.WORK_ORDER_ATTACHMENT_MAX_FILES || 10);
  const maxBytes = Number(CFG.WORK_ORDER_ATTACHMENT_MAX_BYTES || 26214400);

  if (!requestedCompanyId) throw new Error("Falta COMPANY_ID para subir el archivo.");
  if (!woNumber) throw new Error("Falta WO_NUMBER para subir el archivo.");
  if (!uploadId) throw new Error("Falta el identificador del archivo.");
  if (!rawBase64) throw new Error("El archivo esta vacio.");

  const estimatedBytes = Math.floor(rawBase64.length * 3 / 4);
  if (estimatedBytes > maxBytes + 3) {
    throw new Error(
      "El archivo excede el limite de " +
      Math.round(maxBytes / 1024 / 1024) +
      " MB."
    );
  }

  const fileName = sanitizeWorkOrderAttachmentName_(data.name);
  const mimeType = normalizeWorkOrderAttachmentMimeType_(data.mimeType, fileName);
  if (mimeType.indexOf("image/") !== 0 && mimeType.indexOf("video/") !== 0) {
    throw new Error("Solo se permiten archivos de fotos y videos.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

  let headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(header) {
      return String(header || "").trim();
    });

  const values = sh.getDataRange().getValues();
  const idxWO = headers.indexOf("WO_NUMBER");
  const idxCompany = headers.indexOf("COMPANY_ID");
  if (idxWO === -1) throw new Error("WORK_ORDERS no tiene la columna WO_NUMBER.");

  let rowNumber = 0;
  let rowData = null;

  for (let i = 1; i < values.length; i++) {
    const currentWO = String(values[i][idxWO] || "").trim();
    const currentCompanyId = idxCompany >= 0
      ? String(values[i][idxCompany] || "").trim().toUpperCase()
      : String(CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();

    if (currentWO === woNumber && currentCompanyId === requestedCompanyId) {
      rowNumber = i + 1;
      rowData = values[i];
      break;
    }
  }

  if (!rowNumber || !rowData) {
    throw new Error("No se encontro la orden " + woNumber + ".");
  }

  const session = requireWorkOrderSession_(
    sessionToken,
    rowData,
    headers,
    ["TECH", "OWNER", "ADMIN", "ORDENES"],
    "subir archivos a"
  );

  const nsn = String(getRowValue_(rowData, headers, "NSN") || "").trim();
  if (!nsn) throw new Error("La orden no tiene NSN.");

  const orderFolder = createOrderFolders_(requestedCompanyId, nsn, woNumber);
  const attachmentFolder = getOrCreateFolder_(orderFolder, "Order Attachments");

  const description = "CREATE_ORDER_UPLOAD:" + uploadId;
  const existingFiles = attachmentFolder.getFiles();
  let existingFile = null;
  let fileCount = 0;

  while (existingFiles.hasNext()) {
    const currentFile = existingFiles.next();
    fileCount++;

    if (String(currentFile.getDescription() || "") === description) {
      existingFile = currentFile;
    }
  }

  headers = ensureSheetColumns_(sh, ["FOLDER_URL", "ATTACHMENT_COUNT"]);

  if (existingFile) {
    setCellByHeader_(sh, rowNumber, headers, "FOLDER_URL", orderFolder.getUrl());
    setCellByHeader_(sh, rowNumber, headers, "ATTACHMENT_COUNT", fileCount);

    return {
      success: true,
      reused: true,
      id: existingFile.getId(),
      url: existingFile.getUrl(),
      name: existingFile.getName(),
      mimeType: existingFile.getMimeType(),
      attachmentCount: fileCount
    };
  }

  if (fileCount >= maxFiles) {
    throw new Error("Esta orden ya tiene el maximo de " + maxFiles + " archivos.");
  }

  let bytes;
  try {
    bytes = Utilities.base64Decode(rawBase64);
  } catch (err) {
    throw new Error("No se pudo leer el archivo seleccionado.");
  }

  if (bytes.length > maxBytes) {
    throw new Error(
      "El archivo excede el limite de " +
      Math.round(maxBytes / 1024 / 1024) +
      " MB."
    );
  }

  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = attachmentFolder.createFile(blob);
  file.setDescription(description);

  setCellByHeader_(sh, rowNumber, headers, "FOLDER_URL", orderFolder.getUrl());
  setCellByHeader_(sh, rowNumber, headers, "ATTACHMENT_COUNT", fileCount + 1);

  addLog_(
    requestedCompanyId,
    woNumber,
    "ORDER ATTACHMENT UPLOADED",
    "",
    "",
    getSessionActorLabel_(session),
    fileName
  );

  addAuditLog_(
    "ORDERS",
    "WORK_ORDER_ATTACHMENT_UPLOADED",
    requestedCompanyId,
    "WORK_ORDER",
    woNumber,
    session,
    {
      fileId: file.getId(),
      fileName: file.getName(),
      mimeType: file.getMimeType(),
      size: bytes.length
    }
  );

  touchAppCacheVersion_();

  return {
    success: true,
    reused: false,
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    attachmentCount: fileCount + 1
  };
}

function sanitizeWorkOrderAttachmentName_(value) {
  const safeName = String(value || "archivo")
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, "-")
    .trim()
    .substring(0, 180);

  return safeName || "archivo";
}

function normalizeWorkOrderAttachmentMimeType_(mimeType, fileName) {
  const normalized = String(mimeType || "").trim().toLowerCase();
  if (normalized) return normalized;

  const extension = String(fileName || "").toLowerCase().split(".").pop();
  const byExtension = {
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    gif: "image/gif",
    webp: "image/webp",
    heic: "image/heic",
    heif: "image/heif",
    bmp: "image/bmp",
    tif: "image/tiff",
    tiff: "image/tiff",
    mp4: "video/mp4",
    mov: "video/quicktime",
    m4v: "video/x-m4v",
    webm: "video/webm",
    avi: "video/x-msvideo",
    mpg: "video/mpeg",
    mpeg: "video/mpeg",
    "3gp": "video/3gpp"
  };

  return byExtension[extension] || "application/octet-stream";
}

function getWorkOrderAttachmentGallery(rowNumber, woNumber, companyId, sessionToken) {
  const context = resolveWorkOrderAttachmentContext_(
    rowNumber,
    woNumber,
    companyId,
    sessionToken,
    "ver las fotos y videos de"
  );

  const filesFolder = getWorkOrderAttachmentFilesFolder_(context);
  const files = [];

  if (filesFolder) {
    const iterator = filesFolder.getFiles();

    while (iterator.hasNext()) {
      const file = iterator.next();
      const mimeType = String(file.getMimeType() || "").toLowerCase();
      if (mimeType.indexOf("image/") !== 0 && mimeType.indexOf("video/") !== 0) {
        continue;
      }

      let thumbnailDataUrl = "";
      try {
        const thumbnail = file.getThumbnail();
        if (thumbnail) {
          thumbnailDataUrl =
            "data:" +
            (thumbnail.getContentType() || "image/png") +
            ";base64," +
            Utilities.base64Encode(thumbnail.getBytes());
        }
      } catch (err) {
        Logger.log("WARN getWorkOrderAttachmentGallery thumbnail: " + err);
      }

      files.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: mimeType,
        kind: mimeType.indexOf("video/") === 0 ? "video" : "image",
        size: Number(file.getSize() || 0),
        createdAtMs: file.getDateCreated().getTime(),
        createdAt: Utilities.formatDate(
          file.getDateCreated(),
          CFG.TIMEZONE,
          "MM/dd/yyyy hh:mm a"
        ),
        thumbnailDataUrl: thumbnailDataUrl
      });
    }
  }

  files.sort(function(a, b) {
    return Number(b.createdAtMs || 0) - Number(a.createdAtMs || 0);
  });

  files.forEach(function(file) {
    delete file.createdAtMs;
  });

  return {
    success: true,
    rowNumber: context.rowNumber,
    woNumber: context.woNumber,
    companyId: context.companyId,
    nsn: context.nsn,
    attachmentCount: files.length,
    files: files
  };
}

function getWorkOrderAttachmentContent(fileId, rowNumber, woNumber, companyId, sessionToken) {
  fileId = String(fileId || "").trim();
  if (!/^[a-zA-Z0-9_-]{20,}$/.test(fileId)) {
    throw new Error("El archivo solicitado no es valido.");
  }

  const context = resolveWorkOrderAttachmentContext_(
    rowNumber,
    woNumber,
    companyId,
    sessionToken,
    "abrir las fotos y videos de"
  );

  const filesFolder = getWorkOrderAttachmentFilesFolder_(context);
  if (!filesFolder) throw new Error("Esta orden no tiene archivos.");

  const iterator = filesFolder.getFiles();
  let selectedFile = null;

  while (iterator.hasNext()) {
    const currentFile = iterator.next();
    if (currentFile.getId() === fileId) {
      selectedFile = currentFile;
      break;
    }
  }

  if (!selectedFile) {
    throw new Error("El archivo no pertenece a esta orden.");
  }

  const mimeType = String(selectedFile.getMimeType() || "").toLowerCase();
  if (mimeType.indexOf("image/") !== 0 && mimeType.indexOf("video/") !== 0) {
    throw new Error("Solo se pueden abrir fotos y videos.");
  }

  const maxBytes = Number(CFG.WORK_ORDER_ATTACHMENT_MAX_BYTES || 26214400);
  const fileSize = Number(selectedFile.getSize() || 0);
  if (fileSize > maxBytes) {
    throw new Error(
      "El archivo excede el limite de vista de " +
      Math.round(maxBytes / 1024 / 1024) +
      " MB."
    );
  }

  const blob = selectedFile.getBlob();

  return {
    success: true,
    id: selectedFile.getId(),
    name: selectedFile.getName(),
    mimeType: mimeType,
    size: fileSize,
    data: Utilities.base64Encode(blob.getBytes())
  };
}

function resolveWorkOrderAttachmentContext_(
  rowNumber,
  woNumber,
  requestedCompanyId,
  sessionToken,
  actionName
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("No hay ordenes disponibles.");

  const headers = values[0].map(function(header) {
    return String(header || "").trim();
  });
  const idxWO = headers.indexOf("WO_NUMBER");
  const idxCompany = headers.indexOf("COMPANY_ID");
  if (idxWO === -1) throw new Error("WORK_ORDERS no tiene la columna WO_NUMBER.");

  const requestedWO = String(woNumber || "").trim();
  requestedCompanyId = String(requestedCompanyId || "").trim().toUpperCase();
  let resolvedRow = Number(rowNumber || 0);

  function rowMatchesRequest(row) {
    const rowWO = String(row[idxWO] || "").trim();
    const rowCompanyId = idxCompany >= 0
      ? String(row[idxCompany] || "").trim().toUpperCase()
      : String(CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();

    return (!requestedWO || rowWO === requestedWO) &&
      (!requestedCompanyId || rowCompanyId === requestedCompanyId);
  }

  if (
    resolvedRow < 2 ||
    resolvedRow > values.length ||
    !rowMatchesRequest(values[resolvedRow - 1])
  ) {
    resolvedRow = 0;

    for (let i = 1; i < values.length; i++) {
      if (rowMatchesRequest(values[i])) {
        resolvedRow = i + 1;
        break;
      }
    }
  }

  if (!resolvedRow) {
    throw new Error(
      "No se encontro la orden " +
      requestedWO +
      (requestedCompanyId ? " para " + requestedCompanyId : "") +
      "."
    );
  }

  const rowData = values[resolvedRow - 1];
  requireWorkOrderSession_(
    sessionToken,
    rowData,
    headers,
    ["TECH", "OWNER", "ADMIN", "ORDENES", "SUPERVISOR"],
    actionName
  );

  return {
    sheet: sh,
    headers: headers,
    rowData: rowData,
    rowNumber: resolvedRow,
    woNumber: String(getRowValue_(rowData, headers, "WO_NUMBER") || "").trim(),
    companyId: String(
      getRowValue_(rowData, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID
    ).trim().toUpperCase(),
    nsn: String(getRowValue_(rowData, headers, "NSN") || "").trim(),
    folderUrl: String(getRowValue_(rowData, headers, "FOLDER_URL") || "").trim()
  };
}

function getWorkOrderAttachmentFilesFolder_(context) {
  if (!context || !context.folderUrl) return null;

  const folderId = extractDriveFileIdForPortal_(context.folderUrl);
  const orderFolder = DriveApp.getFolderById(folderId);

  if (String(orderFolder.getName() || "").trim() !== String(context.woNumber || "").trim()) {
    throw new Error("La carpeta de archivos no corresponde a esta orden.");
  }

  const folders = orderFolder.getFoldersByName("Order Attachments");
  return folders.hasNext() ? folders.next() : null;
}

function createPMWorkOrdersBatchFromApp(data, sessionToken) {
  data = data || {};

  const companyId = String(data.COMPANY_ID || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ORDENES"], companyId);
  const pmType = String(data.PM_TYPE || "").trim();
  const equipmentOriginal = String(data.REPORTED_EQUIPMENT || "").trim();
  const priorityOriginal = String(data.ORDER_PRIORITY || "").trim();
  const problemOriginal = String(data.REPORTED_PROBLEM || "").trim();
  const manager = String(data.MANAGER || "").trim();

  if (!companyId) throw new Error("No llego COMPANY_ID.");
  if (!pmType) throw new Error("PM_TYPE es obligatorio para PM_FORM.");
  if (!equipmentOriginal) throw new Error("Reported Equipment es obligatorio.");
  if (!priorityOriginal) throw new Error("Order Priority es obligatoria.");
  if (!problemOriginal) throw new Error("Reported Problem es obligatorio.");
  if (!manager) throw new Error("Manager Name es obligatorio.");

  const rawNsns = Array.isArray(data.STORE_NSNS)
    ? data.STORE_NSNS
    : String(data.STORE_NSNS || "").split(/[,\n;]+/);

  const nsns = [];
  const seen = {};

  rawNsns.forEach(function(value) {
    const nsn = normalizeNSN_(value);
    if (!nsn || seen[nsn]) return;
    seen[nsn] = true;
    nsns.push(nsn);
  });

  if (!nsns.length) {
    throw new Error("Selecciona al menos una tienda para crear los PM.");
  }

  if (nsns.length > 150) {
    throw new Error("Selecciona 150 tiendas o menos por lote.");
  }

  const missingStores = nsns.filter(function(nsn) {
    const store = getStoreByNSN_(nsn, companyId);
    return !store || !store.fullAddress;
  });

  if (missingStores.length) {
    throw new Error("No se encontro tienda/direccion para NSN: " + missingStores.join(", "));
  }

  const created = [];
  const equipmentEn = safeTranslate_(equipmentOriginal, "auto", "en");
  const problemEn = safeTranslate_(problemOriginal, "auto", "en");

  nsns.forEach(function(nsn) {
    const result = createWorkOrderFromApp({
      COMPANY_ID: companyId,
      WO_TYPE: "PM_FORM",
      PM_TYPE: pmType,
      NSN: nsn,
      REPORTED_EQUIPMENT: equipmentOriginal,
      ORDER_PRIORITY: priorityOriginal,
      REPORTED_PROBLEM: problemOriginal,
      _EQUIPMENT_EN: equipmentEn,
      _PROBLEM_EN: problemEn,
      MANAGER: manager
    }, sessionToken);

    created.push(result.woNumber);
  });

  addAuditLog_("ORDERS", "PM_WORK_ORDERS_BATCH_CREATED", companyId, "WORK_ORDER_BATCH", created.join(","), session, {
    count: created.length,
    pmType: pmType,
    nsns: nsns,
    woNumbers: created
  });

  return {
    success: true,
    count: created.length,
    woNumbers: created,
    nsns: nsns
  };
}
