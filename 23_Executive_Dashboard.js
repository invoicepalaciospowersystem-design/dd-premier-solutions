// =====================================================
// FILE: 23_Executive_Dashboard.gs
// =====================================================

function getOwnerExecutiveDashboard(sessionToken, companyId) {
  const session = requireSession_(sessionToken, ["OWNER"]);
  const requestedCompany = String(companyId || "").trim().toUpperCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const companies = getExecutiveCompanies_(ss);
  const companyMap = {};

  companies.forEach(function(c) {
    companyMap[c.companyId] = c.companyName || c.companyId;
  });

  const dashboard = {
    generatedAt: Utilities.formatDate(new Date(), CFG.TIMEZONE, "MM/dd/yyyy hh:mm a"),
    scopeCompanyId: requestedCompany,
    scopeLabel: requestedCompany ? (companyMap[requestedCompany] || requestedCompany) : "Todas las empresas",
    totals: {
      activeCompanies: requestedCompany ? 1 : companies.filter(function(c) { return c.active === "YES"; }).length,
      totalOrders: 0,
      openOrders: 0,
      completedPendingBilling: 0,
      closedOrders: 0,
      emergencyOrders: 0,
      ordersToday: 0,
      ordersLast7Days: 0,
      oldestOpenDays: 0,
      invoicesCreated: 0,
      economyPending: 0,
      economyInvoiced: 0,
      revenueTracked: 0,
      quotesOpen: 0,
      pmReportsGenerated: 0,
      unreadNotifications: 0
    },
    companies: [],
    alerts: []
  };

  const perCompany = {};
  companies.forEach(function(c) {
    if (requestedCompany && c.companyId !== requestedCompany) return;
    ensureExecutiveCompanySummary_(perCompany, c.companyId, c.companyName || c.companyId);
  });

  const woCompanyMap = collectExecutiveWorkOrders_(ss, dashboard, perCompany, companyMap, requestedCompany);
  collectExecutiveInvoices_(ss, dashboard, perCompany, companyMap, requestedCompany, woCompanyMap);
  collectExecutiveEconomy_(ss, dashboard, perCompany, companyMap, requestedCompany);
  collectExecutiveNotifications_(ss, dashboard, requestedCompany);

  dashboard.companies = Object.keys(perCompany)
    .map(function(key) { return perCompany[key]; })
    .filter(function(c) {
      return c.totalOrders || c.openOrders || c.invoicesCreated || c.quotesOpen || c.pmReportsGenerated || !requestedCompany;
    })
    .sort(function(a, b) {
      return Number(b.openOrders || 0) - Number(a.openOrders || 0) ||
        Number(b.emergencyOrders || 0) - Number(a.emergencyOrders || 0) ||
        String(a.companyName || "").localeCompare(String(b.companyName || ""));
    });

  dashboard.alerts = buildExecutiveAlerts_(dashboard.totals);

  addAuditLog_("DASHBOARD", "OWNER_EXECUTIVE_DASHBOARD_VIEWED", requestedCompany || "ALL", "DASHBOARD", requestedCompany || "ALL", session, {
    scopeLabel: dashboard.scopeLabel
  });

  return dashboard;
}

function collectExecutiveWorkOrders_(ss, dashboard, perCompany, companyMap, requestedCompany) {
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  const woCompanyMap = {};
  if (!sh || sh.getLastRow() < 2) return woCompanyMap;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });
  const now = new Date();
  const todayLabel = Utilities.formatDate(now, CFG.TIMEZONE, "yyyy-MM-dd");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const companyId = String(getExecutiveValue_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
    if (requestedCompany && companyId !== requestedCompany) continue;

    const company = ensureExecutiveCompanySummary_(perCompany, companyId, companyMap[companyId] || companyId);
    const woNumber = String(getExecutiveValue_(row, headers, ["WO_NUMBER"]) || "").trim();
    if (woNumber) woCompanyMap[woNumber] = companyId;

    const status = String(getExecutiveValue_(row, headers, ["STATUS"]) || "").trim().toUpperCase();
    const priority = String(getExecutiveValue_(row, headers, ["ORDER_PRIORITY", "PRIORITY"]) || "").trim().toUpperCase();
    const createdAt = parseExecutiveDate_(getExecutiveValue_(row, headers, ["DATE_CREATED", "CREATED_AT"]));
    const quoteStatus = String(getExecutiveValue_(row, headers, ["QUOTE_STATUS"]) || "").trim().toUpperCase();
    const quoteUrl = String(getExecutiveValue_(row, headers, ["QUOTE_EN_URL", "QUOTE_ES_URL"]) || "").trim();
    const pmStatus = String(getExecutiveValue_(row, headers, ["PM_REPORT_STATUS"]) || "").trim().toUpperCase();
    const pmUrl = String(getExecutiveValue_(row, headers, ["PM_REPORT_ES_URL", "PM_REPORT_EN_URL"]) || "").trim();
    const isClosed = status === "CLOSED";
    const isCompleted = status === "COMPLETED";
    const isOpen = status && status !== "CLOSED" && status !== "COMPLETED" && status !== "DELETED";
    const isEmergency = priority.indexOf("EMERGENCY") !== -1 || priority.indexOf("EMERGENCIA") !== -1 ||
      priority.indexOf("URGENT") !== -1 || priority.indexOf("URGENTE") !== -1;

    dashboard.totals.totalOrders++;
    company.totalOrders++;

    if (isOpen) {
      dashboard.totals.openOrders++;
      company.openOrders++;

      if (createdAt) {
        const ageDays = Math.max(0, Math.floor((now.getTime() - createdAt.getTime()) / 86400000));
        dashboard.totals.oldestOpenDays = Math.max(dashboard.totals.oldestOpenDays, ageDays);
        company.oldestOpenDays = Math.max(company.oldestOpenDays, ageDays);
      }
    }

    if (isCompleted) {
      dashboard.totals.completedPendingBilling++;
      company.completedPendingBilling++;
    }

    if (isClosed) {
      dashboard.totals.closedOrders++;
      company.closedOrders++;
    }

    if (isEmergency && !isClosed) {
      dashboard.totals.emergencyOrders++;
      company.emergencyOrders++;
    }

    if (createdAt) {
      const createdLabel = Utilities.formatDate(createdAt, CFG.TIMEZONE, "yyyy-MM-dd");
      if (createdLabel === todayLabel) {
        dashboard.totals.ordersToday++;
        company.ordersToday++;
      }

      if (now.getTime() - createdAt.getTime() <= 7 * 86400000) {
        dashboard.totals.ordersLast7Days++;
        company.ordersLast7Days++;
      }
    }

    if (quoteUrl && quoteStatus !== "APPROVED") {
      dashboard.totals.quotesOpen++;
      company.quotesOpen++;
    }

    if (pmUrl || pmStatus === "GENERATED") {
      dashboard.totals.pmReportsGenerated++;
      company.pmReportsGenerated++;
    }
  }

  return woCompanyMap;
}

function collectExecutiveInvoices_(ss, dashboard, perCompany, companyMap, requestedCompany, woCompanyMap) {
  const sh = ss.getSheetByName("INVOICES");
  if (!sh || sh.getLastRow() < 2) return;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const woNumber = String(getExecutiveValue_(row, headers, ["WO_NUMBER"]) || "").trim();
    const companyId = String(
      getExecutiveValue_(row, headers, ["COMPANY_ID"]) ||
      woCompanyMap[woNumber] ||
      CFG.DEFAULT_COMPANY_ID
    ).trim().toUpperCase();

    if (requestedCompany && companyId !== requestedCompany) continue;

    const company = ensureExecutiveCompanySummary_(perCompany, companyId, companyMap[companyId] || companyId);
    dashboard.totals.invoicesCreated++;
    company.invoicesCreated++;
  }
}

function collectExecutiveEconomy_(ss, dashboard, perCompany, companyMap, requestedCompany) {
  const sh = ss.getSheetByName(CFG.SHEET_ECONOMY);
  if (!sh || sh.getLastRow() < 2) return;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const companyId = String(getExecutiveValue_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
    if (requestedCompany && companyId !== requestedCompany) continue;

    const company = ensureExecutiveCompanySummary_(perCompany, companyId, companyMap[companyId] || companyId);
    const status = String(getExecutiveValue_(row, headers, ["STATUS"]) || "").trim().toUpperCase();
    const amount = Number(getExecutiveValue_(row, headers, ["AMOUNT", "INVOICE_TOTAL", "TOTAL"]) || 0);

    if (status === "INVOICED") {
      dashboard.totals.economyInvoiced++;
      company.economyInvoiced++;
    } else {
      dashboard.totals.economyPending++;
      company.economyPending++;
    }

    if (!isNaN(amount)) {
      dashboard.totals.revenueTracked += amount;
      company.revenueTracked += amount;
    }
  }
}

function collectExecutiveNotifications_(ss, dashboard, requestedCompany) {
  const sh = ss.getSheetByName(CFG.SHEET_NOTIFICATIONS);
  if (!sh || sh.getLastRow() < 2) return;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const companyId = String(getExecutiveValue_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
    if (requestedCompany && companyId !== requestedCompany) continue;

    const read = String(getExecutiveValue_(row, headers, ["READ"]) || "").trim().toUpperCase();
    if (read !== "YES" && read !== "TRUE") {
      dashboard.totals.unreadNotifications++;
    }
  }
}

function getExecutiveCompanies_(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_COMPANIES);
  if (!sh || sh.getLastRow() < 2) {
    return [{
      companyId: CFG.DEFAULT_COMPANY_ID,
      companyName: getCompanyName(CFG.DEFAULT_COMPANY_ID),
      active: "YES"
    }];
  }

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });
  const companies = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const companyId = String(getExecutiveValue_(row, headers, ["COMPANY_ID"]) || "").trim().toUpperCase();
    if (!companyId) continue;

    companies.push({
      companyId: companyId,
      companyName: String(getExecutiveValue_(row, headers, ["COMPANY_NAME", "LEGAL_NAME"]) || companyId).trim(),
      active: String(getExecutiveValue_(row, headers, ["ACTIVE"]) || "YES").trim().toUpperCase()
    });
  }

  return companies.length ? companies : [{
    companyId: CFG.DEFAULT_COMPANY_ID,
    companyName: getCompanyName(CFG.DEFAULT_COMPANY_ID),
    active: "YES"
  }];
}

function ensureExecutiveCompanySummary_(perCompany, companyId, companyName) {
  companyId = String(companyId || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();

  if (!perCompany[companyId]) {
    perCompany[companyId] = {
      companyId: companyId,
      companyName: companyName || companyId,
      totalOrders: 0,
      openOrders: 0,
      completedPendingBilling: 0,
      closedOrders: 0,
      emergencyOrders: 0,
      ordersToday: 0,
      ordersLast7Days: 0,
      oldestOpenDays: 0,
      invoicesCreated: 0,
      economyPending: 0,
      economyInvoiced: 0,
      revenueTracked: 0,
      quotesOpen: 0,
      pmReportsGenerated: 0
    };
  }

  return perCompany[companyId];
}

function buildExecutiveAlerts_(totals) {
  const alerts = [];

  if (totals.emergencyOrders > 0) {
    alerts.push({
      label: "Ordenes emergencia / urgentes",
      value: totals.emergencyOrders,
      tone: "danger"
    });
  }

  if (totals.completedPendingBilling > 0) {
    alerts.push({
      label: "Completadas esperando billing",
      value: totals.completedPendingBilling,
      tone: "warning"
    });
  }

  if (totals.economyPending > 0) {
    alerts.push({
      label: "Filas pendientes en economia",
      value: totals.economyPending,
      tone: "warning"
    });
  }

  if (totals.quotesOpen > 0) {
    alerts.push({
      label: "Quotes abiertos",
      value: totals.quotesOpen,
      tone: "info"
    });
  }

  if (!alerts.length) {
    alerts.push({
      label: "Sin alertas criticas",
      value: "OK",
      tone: "success"
    });
  }

  return alerts;
}

function getExecutiveValue_(row, headers, names) {
  const normalized = headers.map(function(h) {
    return String(h || "").trim().toUpperCase();
  });

  for (let i = 0; i < names.length; i++) {
    const idx = normalized.indexOf(String(names[i] || "").trim().toUpperCase());
    if (idx >= 0) return row[idx];
  }

  return "";
}

function parseExecutiveDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const d = new Date(value);
  if (isNaN(d.getTime())) return null;
  return d;
}
