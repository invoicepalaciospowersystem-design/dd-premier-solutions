// =====================================================
// FILE: 01_Router.gs
// =====================================================

function doGet(e) {
  try {
    e = e || { parameter: {} };
    const p = e.parameter || {};
    const ownerPortal = isOwnerPortalRequest_(p);
    const baseUrl = getWebAppBaseUrl_(p.companyId, ownerPortal);

    if (p.action && p.wo) {
      return handleWorkOrderAction_(e);
    }

    if (p.view === "tech") {
      const template = HtmlService.createTemplateFromFile("Tech");
      applyTemplateDefaults_(template, p, baseUrl);
      template.techName = p.tech || "";
      template.baseUrl = baseUrl;
      template.unitedRefrigerationAccount = CFG.UNITED_REFRIGERATION_ACCOUNT || "";

      return template.evaluate()
        .setTitle("Technician Work Orders")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "customer") {
      const template = HtmlService.createTemplateFromFile("Customer");
      applyTemplateDefaults_(template, p, baseUrl);
      template.supervisor = p.supervisor || "";
      template.baseUrl = baseUrl;

      return template.evaluate()
        .setTitle("Supervisor Dashboard")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "techAdmin") {
      const template = HtmlService.createTemplateFromFile("TechAdmin");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;
      template.companyId = p.companyId || "";
      template.ownerOnly = ownerPortal ? "true" : "";

      return template.evaluate()
        .setTitle("Admin Technician Portal")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "supervisorAdmin") {
      const template = HtmlService.createTemplateFromFile("SupervisorAdmin");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;
      template.companyId = p.companyId || "";
      template.ownerOnly = ownerPortal ? "true" : "";

      return template.evaluate()
        .setTitle("Admin Supervisor Portal")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "orders") {
      const template = HtmlService.createTemplateFromFile("Ordenes");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;
      template.companyId = p.companyId || "";
      template.ownerOnly = ownerPortal ? "true" : "";

      return template.evaluate()
        .setTitle("Work Orders")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "economy") {
      const template = HtmlService.createTemplateFromFile("Economy");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;
      template.companyId = p.companyId || "";
      template.ownerOnly = ownerPortal ? "true" : "";

      return template.evaluate()
        .setTitle("Economy")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "accounting") {
      const template = HtmlService.createTemplateFromFile("Contabilidad");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;
      template.companyId = p.companyId || "";
      template.ownerOnly = ownerPortal ? "true" : "";

      return template.evaluate()
        .setTitle("Contabilidad")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "stores") {
      const template = HtmlService.createTemplateFromFile("Stores");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;
      template.companyId = p.companyId || "";
      template.ownerOnly = ownerPortal ? "true" : "";

      return template.evaluate()
        .setTitle("Stores")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "companies") {
      const template = HtmlService.createTemplateFromFile("Companies");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;

      return template.evaluate()
        .setTitle("Companies")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "users") {
      const template = HtmlService.createTemplateFromFile("Users");
      applyTemplateDefaults_(template, p, baseUrl);
      template.baseUrl = baseUrl;
      template.companyId = p.companyId || "";
      template.ownerOnly = ownerPortal ? "true" : "";

      return template.evaluate()
        .setTitle("Users")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "close") {
      const template = HtmlService.createTemplateFromFile("CloseOrder");
      applyTemplateDefaults_(template, p, baseUrl);
      template.row = p.row || "";
      template.wo = p.wo || "";
      template.techName = p.tech || "";
      template.baseUrl = baseUrl;
      template.returnTo = p.returnTo || "";
      template.specialInvoice = p.specialInvoice || "";

      return template.evaluate()
        .setTitle("Close Work Order")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

        if (p.view === "pm_report") {
      const template = HtmlService.createTemplateFromFile("PM_Report");
      applyTemplateDefaults_(template, p, baseUrl);
      template.row = p.row || "";
      template.wo = p.wo || "";
      template.techName = p.tech || "";
      template.baseUrl = baseUrl;
      template.returnTo = p.returnTo || "";

      return template.evaluate()
        .setTitle("PM Report")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "createOrder") {
      const template = HtmlService.createTemplateFromFile("CreateOrder");
      applyTemplateDefaults_(template, p, baseUrl);
      template.companyId = p.companyId || "";
      template.baseUrl = baseUrl;
      template.returnTo = p.returnTo || "";

      return template.evaluate()
        .setTitle("Create Work Order")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (p.view === "quote") {
  const template = HtmlService.createTemplateFromFile("Quote");
  applyTemplateDefaults_(template, p, baseUrl);
  template.row = p.row || "";
  template.wo = p.wo || "";
  template.companyId = p.companyId || "";
  template.baseUrl = baseUrl;
  template.returnTo = p.returnTo || "";

  return template.evaluate()
    .setTitle("Create Quote")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

    const template = HtmlService.createTemplateFromFile("Main");
    applyTemplateDefaults_(template, p, baseUrl);
    template.baseUrl = baseUrl;
    template.companyId = p.companyId || "";
    template.ownerOnly = ownerPortal ? "true" : "";
    template.unitedRefrigerationAccount = CFG.UNITED_REFRIGERATION_ACCOUNT || "";

    return template.evaluate()
      .setTitle(CFG.APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    notifySystemError_("WEB_APP_DOGET_ERROR", err, {
      module: "ROUTER",
      parameters: e && e.parameter ? e.parameter : {}
    });
    return HtmlService.createHtmlOutput(
      "<h2>Error</h2><p>" + err.message + "</p>"
    );
  }
}

function isOwnerPortalRequest_(params) {
  params = params || {};
  return params.ownerOnly === "1" ||
    params.ownerOnly === "true" ||
    params.portal === "owner";
}

function getWebAppBaseUrl_(companyId, ownerPortal) {
  if (ownerPortal) {
    return CFG.OWNER_WEB_APP_URL || CFG.WEB_APP_URL;
  }

  companyId = String(companyId || "").trim().toUpperCase();

  if (companyId && CFG.PUBLIC_WEB_APP_URLS && CFG.PUBLIC_WEB_APP_URLS[companyId]) {
    return CFG.PUBLIC_WEB_APP_URLS[companyId];
  }

  return CFG.WEB_APP_URL;
}

function applyTemplateDefaults_(template, params, baseUrl) {
  params = params || {};
  template.baseUrl = baseUrl;
  template.companyId = params.companyId || "";
  template.ownerOnly = isOwnerPortalRequest_(params) ? "true" : "";
  template.returnTo = params.returnTo || "";
  template.watermarkLogoCss = getCompanyWatermarkLogoCss_(params.companyId || "");
}

function getCompanyWatermarkLogoCss_(companyId) {
  companyId = String(companyId || "").trim().toUpperCase();
  if (!companyId) return "none";

  if (companyId === "PPS") {
    return "var(--pps-dashboard-bg);" +
      "--company-watermark-repeat:no-repeat;" +
      "--company-watermark-position:center bottom;" +
      "--company-watermark-size:92vw auto;" +
      "--company-watermark-opacity:.28;" +
      "--company-watermark-filter:grayscale(100%) brightness(1.10)";
  }

  const branding = getCompanyBranding(companyId || "");
  const logoUrl = String(branding.logoImageUrl || "").trim();

  if (!logoUrl) return "none";

  return "url('" + logoUrl.replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "')";
}

function onFormSubmit(e) {
  try {
    if (!e || !e.namedValues) {
      throw new Error("Este script debe ejecutarse desde un trigger de formulario.");
    }

    const sheetName = e.range ? e.range.getSheet().getName() : "";

    if (sheetName === "INVOICES") {
      onInvoiceFormSubmit(e);
      return;
    }

    const nv = e.namedValues;
    const formConfig = getFormConfig_(e);

    const companyId = formConfig.companyId;
    const formLang = formConfig.lang || detectFormLanguage_(nv);

    const nsn = getField_(nv, ["NSN #"]);
    const store = getStoreByNSN_(nsn, companyId);

    const equipmentOriginal = getField_(nv, ["REPORTED EQUIPMENT", "EQUIPO REPORTADO"]);
    const priorityOriginal = getField_(nv, ["ORDER PRIORITY", "PRIORIDAD DE LA ORDEN"]);
    const problemOriginal = getField_(nv, ["REPORTED PROBLEM", "PROBLEMA REPORTADO"]);
    const manager = getField_(nv, ["MANAGER WHO MAKES THE REPORT", "MANAGER QUE HACE EL REPORTE"]);
    const managerEmail = getField_(nv, ["Email Address", "Dirección de correo electrónico", "Correo electrónico"]);

    const woNumber = generateWONumber_(companyId);

    let equipmentEn = "";
    let equipmentEs = "";
    let problemEn = "";
    let problemEs = "";

    if (formLang === "EN") {
      equipmentEn = equipmentOriginal;
      equipmentEs = safeTranslate_(equipmentOriginal, "en", "es");
      problemEn = problemOriginal;
      problemEs = safeTranslate_(problemOriginal, "en", "es");
    } else {
      equipmentEs = equipmentOriginal;
      equipmentEn = safeTranslate_(equipmentOriginal, "es", "en");
      problemEs = problemOriginal;
      problemEn = safeTranslate_(problemOriginal, "es", "en");
    }

    const statusEn = "ORDER RECEIVED";
    const clientName = store.client || CFG.CLIENT_DEFAULT;

    saveWorkOrderRow_({
      companyId: companyId,
      woNumber: woNumber,
      dateCreated: new Date(),
      formLang: formLang,
      client: clientName,
      nsn: nsn,
      equipment: equipmentOriginal,
      equipmentEn: equipmentEn,
      priorityOriginal: priorityOriginal,
      problemOriginal: problemOriginal,
      problemEn: problemEn,
      problemEs: problemEs,
      manager: manager,
      managerEmail: managerEmail,
      status: statusEn,
      pdfEnUrl: "",
      pdfEsUrl: "",
      folderUrl: ""
    });

    try {
      addLog_(companyId, woNumber, "ORDER CREATED", "", statusEn, "System", "Order received from form");
    } catch (logErr) {
      Logger.log("LOG ERROR: " + logErr);
      notifySystemError_("FORM_SUBMIT_LOG_ERROR", logErr, {
        module: "FORM_SUBMIT",
        companyId: companyId || "",
        woNumber: woNumber || "",
        nsn: nsn || ""
      });
    }

    addNotification_(companyId, "ADMIN", woNumber, "NEW", "🚨 New Work Order " + woNumber + " / NSN " + nsn);
    addNotification_(companyId, "David Dominguez", woNumber, "NEW", "🚨 New Work Order " + woNumber + " / NSN " + nsn);
    addNotification_(companyId, "Dayre", woNumber, "NEW", "🚨 New Work Order " + woNumber + " / NSN " + nsn);

    Logger.log("ORDER CREATED OK: " + woNumber);

  } catch (err) {
    Logger.log("ERROR onFormSubmit: " + err);
    notifySystemError_("FORM_SUBMIT_ERROR", err, {
      module: "FORM_SUBMIT",
      sheetName: e && e.range ? e.range.getSheet().getName() : "",
      namedValues: e && e.namedValues ? e.namedValues : {}
    });
    throw err;
  }
}

function onInvoiceFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CFG.SHEET_COMPANIES);

    if (!sh) throw new Error("No existe COMPANIES");

    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const idxCompany = headers.indexOf("COMPANY_ID");

    for (let i = 1; i < data.length; i++) {
      const companyId = String(data[i][idxCompany] || "").trim().toUpperCase();

      if (companyId) {
        syncInvoicesToEconomy(companyId);
      }
    }

  } catch (err) {
    Logger.log("ERROR onInvoiceFormSubmit: " + err);
    notifySystemError_("INVOICE_FORM_SUBMIT_ERROR", err, {
      module: "FORM_SUBMIT"
    });
  }
}
