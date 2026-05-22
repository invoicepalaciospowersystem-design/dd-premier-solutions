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
}

function getDashboardData(companyId, role) {
  companyId = String(companyId || "").trim().toUpperCase();
  role = String(role || "").trim().toUpperCase();

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

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
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

    enrichWorkOrderObject_(obj);

    data.push(obj);
  }

  return data.reverse();
}

function updateWorkOrderFromDashboard(rowNumber, updates) {
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
      Session.getActiveUser().getEmail() || "Dashboard",
      ""
    );
  }

  return true;
}

function sendWorkOrderFromDashboard(rowNumber) {
  if (!rowNumber || isNaN(rowNumber)) {
    throw new Error("No llegó número de fila desde el dashboard. Valor recibido: " + rowNumber);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  dispatchRowToTech_(sh, Number(rowNumber));

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
  const companyParam = companyId ? "&companyId=" + encodeURIComponent(companyId) : "";
  const startJobLink = publicBaseUrl + "?action=START&wo=" + encodeURIComponent(obj.WO_NUMBER) + companyParam;
  const waitingPartsLink = publicBaseUrl + "?action=WAITING&wo=" + encodeURIComponent(obj.WO_NUMBER) + companyParam;
  const completedLink = publicBaseUrl + "?action=COMPLETED&wo=" + encodeURIComponent(obj.WO_NUMBER) + companyParam;

  if (CFG.TEST_MODE) {
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
        "<hr>" +
        "<p><b>Dirección:</b> " + (obj.STORE_ADDRESS || "") + "</p>" +
        "<p><b>Problema:</b><br>" + obj.REPORTED_PROBLEM_ES + "</p>" +
        "<br>" +
        "<a href='" + startJobLink + "' style='background:#2563eb;color:white;padding:12px 20px;text-decoration:none;border-radius:8px;font-weight:bold;display:inline-block;margin-right:10px;'>START JOB / IN PROGRESS</a>" +
        "<a href='" + waitingPartsLink + "' style='background:#f59e0b;color:white;padding:12px 20px;text-decoration:none;border-radius:8px;font-weight:bold;display:inline-block;margin-right:10px;'>PARTS IN TRANSIT / PIEZAS EN ENVÍO</a>" +
        "<a href='" + completedLink + "' style='background:#16a34a;color:white;padding:12px 20px;text-decoration:none;border-radius:8px;font-weight:bold;display:inline-block;'>COMPLETED / TERMINADO</a>" +
        "<br><br><p>Order details are included in this email / Los detalles están incluidos en este correo.</p>"
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

function createWorkOrderFromApp(data) {
  const companyId = String(data.COMPANY_ID || "").trim().toUpperCase();

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

  const equipmentEn = safeTranslate_(equipmentOriginal, "auto", "en");

  const problemEn = safeTranslate_(problemOriginal, "auto", "en");

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
      "App",
      "Created manually from CreateOrder"
    );
  } catch (err) {
    Logger.log("LOG ERROR createWorkOrderFromApp: " + err);
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

  return {
    success: true,
    woNumber: woNumber
  };
}
