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
    alerts: [],
    recentActivity: []
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
  dashboard.recentActivity = collectExecutiveRecentActivity_(ss, companyMap, requestedCompany, 12);

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

function getAdminCompanyDashboard(sessionToken, companyId) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN"], companyId);
  const role = String(session.role || "").trim().toUpperCase();
  const requestedCompany = String(companyId || session.companyId || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
  const companyScope = role === "OWNER"
    ? requestedCompany
    : String(session.companyId || requestedCompany || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();

  if (!companyScope) {
    throw new Error("No se pudo determinar la compania del admin.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const companies = getExecutiveCompanies_(ss);
  const companyMap = {};

  companies.forEach(function(c) {
    companyMap[c.companyId] = c.companyName || c.companyId;
  });

  const dashboard = {
    generatedAt: Utilities.formatDate(new Date(), CFG.TIMEZONE, "MM/dd/yyyy hh:mm a"),
    companyId: companyScope,
    companyName: companyMap[companyScope] || getCompanyName(companyScope) || companyScope,
    periodLabel: "",
    totals: {
      activeCompanies: 1,
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
      economyPaid: 0,
      revenueTracked: 0,
      quotesOpen: 0,
      pmReportsGenerated: 0,
      economyRows: 0,
      totalBilled: 0,
      taxTotal: 0,
      laborHours: 0,
      laborBilled: 0,
      partsBilled: 0,
      partsCost: 0,
      techLaborCost: 0,
      grossProfit: 0
    },
    investors: {
      davidPartsCost: 0,
      yoelPartsCost: 0,
      unassignedPartsCost: 0
    }
  };

  const perCompany = {};
  ensureExecutiveCompanySummary_(perCompany, companyScope, dashboard.companyName);

  const woCompanyMap = collectExecutiveWorkOrders_(ss, dashboard, perCompany, companyMap, companyScope);
  collectExecutiveInvoices_(ss, dashboard, perCompany, companyMap, companyScope, woCompanyMap);
  collectAdminEconomyFinancials_(ss, dashboard, companyScope);

  addAuditLog_("DASHBOARD", "ADMIN_COMPANY_DASHBOARD_VIEWED", companyScope, "DASHBOARD", companyScope, session, {
    periodLabel: dashboard.periodLabel
  });

  return dashboard;
}

function collectAdminEconomyFinancials_(ss, dashboard, companyId) {
  const sh = ss.getSheetByName(CFG.SHEET_ECONOMY);
  if (!sh || sh.getLastRow() < 2) return;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });
  const companyRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedRow_(row, headers)) continue;

    const rowCompany = String(getExecutiveValue_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
    if (rowCompany !== companyId) continue;

    companyRows.push(row);
  }

  let currentPeriod = "";
  try {
    const period = getCurrentEconomyPeriod();
    currentPeriod = period && period.label ? String(period.label || "").trim() : "";
  } catch (err) {
    currentPeriod = "";
  }

  const hasPeriodRows = companyRows.some(function(row) {
    return String(getExecutiveValue_(row, headers, ["PERIOD_LABEL"]) || "").trim();
  });

  const rows = hasPeriodRows && currentPeriod
    ? companyRows.filter(function(row) {
        return String(getExecutiveValue_(row, headers, ["PERIOD_LABEL"]) || "").trim() === currentPeriod;
      })
    : companyRows;

  dashboard.periodLabel = hasPeriodRows && currentPeriod ? currentPeriod : "All Periods";

  rows.forEach(function(row) {
    const status = String(getExecutiveValue_(row, headers, ["STATUS"]) || "").trim().toUpperCase();
    const amount = parseMoneyFlexible_(getExecutiveValue_(row, headers, ["AMOUNT", "INVOICE_TOTAL", "TOTAL"]));
    const tax = parseMoneyFlexible_(getExecutiveValue_(row, headers, ["TAX", "TAX_AMOUNT"]));
    const hours = parseMoneyFlexible_(getExecutiveValue_(row, headers, ["HORAS", "LABOR_HOURS", "HOURS"]));
    const laborBilled = getAdminLaborBilled_(row, headers, hours);
    const partsCost = getMaterialCostValue_(row, headers);
    const partsBilled = getAdminPartsBilled_(row, headers, amount, tax, laborBilled);
    const techLaborCost = parseMoneyFlexible_(getExecutiveValue_(row, headers, ["TECH_LABOR_COST", "TECH_LABOR_PAY"]));
    const invSource = String(getExecutiveValue_(row, headers, ["INV_SOURCE"]) || "").trim().toUpperCase();
    const split = getAdminInvestorSplit_(invSource, partsCost);

    dashboard.totals.economyRows++;
    dashboard.totals.totalBilled += amount;
    dashboard.totals.taxTotal += tax;
    dashboard.totals.laborHours += hours;
    dashboard.totals.laborBilled += laborBilled;
    dashboard.totals.partsBilled += partsBilled;
    dashboard.totals.partsCost += partsCost;
    dashboard.totals.techLaborCost += techLaborCost;

    if (status === "PAID") dashboard.totals.economyPaid++;
    if (status === "INVOICED") dashboard.totals.economyInvoiced++;
    if (!status || status === "PENDING") dashboard.totals.economyPending++;

    dashboard.investors.davidPartsCost += split.david;
    dashboard.investors.yoelPartsCost += split.yoel;
    dashboard.investors.unassignedPartsCost += split.unassigned;
  });

  dashboard.totals.grossProfit =
    dashboard.totals.totalBilled -
    dashboard.totals.taxTotal -
    dashboard.totals.partsCost -
    dashboard.totals.techLaborCost;
}

function getAdminLaborBilled_(row, headers, hours) {
  const explicit = parseMoneyFlexible_(getExecutiveValue_(row, headers, [
    "LABOR_AMOUNT",
    "LABOR",
    "LABOR_COST",
    "TOTAL_COBRO_HORAS"
  ]));

  if (explicit) return explicit;
  return calculateAdminLaborCharge_(hours);
}

function calculateAdminLaborCharge_(hours) {
  hours = Number(hours || 0);
  if (hours <= 0) return 0;
  if (hours <= 1) return 200;
  return 200 + ((hours - 1) * 130);
}

function getAdminPartsBilled_(row, headers, amount, tax, laborBilled) {
  const explicit = parseMoneyFlexible_(getExecutiveValue_(row, headers, [
    "PARTS_AMOUNT",
    "PARTS_CHARGE",
    "TOTAL_COBRO_PARTS",
    "COBRO_PARTS"
  ]));

  if (explicit) return explicit;
  if (!amount) return 0;

  return Math.max(0, amount - Number(tax || 0) - Number(laborBilled || 0));
}

function getAdminInvestorSplit_(source, cost) {
  source = String(source || "").trim().toUpperCase();
  cost = Number(cost || 0);

  // Regla PPS: quien pone la inversion principal cobra 75% de parts; el otro cobra 25%.
  if (source === "YOEL") {
    return { david: cost * 0.25, yoel: cost * 0.75, unassigned: 0 };
  }

  if (source === "DAVID") {
    return { david: cost * 0.75, yoel: cost * 0.25, unassigned: 0 };
  }

  if (source === "MISCELANEAS") {
    return { david: cost * 0.50, yoel: cost * 0.50, unassigned: 0 };
  }

  return { david: 0, yoel: 0, unassigned: cost };
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

function collectExecutiveRecentActivity_(ss, companyMap, requestedCompany, maxItems) {
  const items = [];
  collectExecutiveAuditActivity_(ss, companyMap, requestedCompany, items);
  collectExecutiveWorkOrderLogActivity_(ss, companyMap, requestedCompany, items);

  return items
    .filter(function(item) { return item.sortTime; })
    .sort(function(a, b) { return b.sortTime - a.sortTime; })
    .slice(0, maxItems || 12)
    .map(function(item) {
      delete item.sortTime;
      return item;
    });
}

function collectExecutiveAuditActivity_(ss, companyMap, requestedCompany, items) {
  const sh = ss.getSheetByName(CFG.SHEET_AUDIT_LOGS || "AUDIT_LOGS");
  if (!sh || sh.getLastRow() < 2) return;

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h || "").trim(); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const action = String(getExecutiveValue_(row, headers, ["ACTION"]) || "").trim();
    if (!action || action === "OWNER_EXECUTIVE_DASHBOARD_VIEWED") continue;

    const companyId = String(getExecutiveValue_(row, headers, ["COMPANY_ID"]) || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
    if (requestedCompany && companyId !== requestedCompany) continue;

    const timestamp = parseExecutiveDate_(getExecutiveValue_(row, headers, ["TIMESTAMP"]));
    if (!timestamp) continue;

    const entityType = String(getExecutiveValue_(row, headers, ["ENTITY_TYPE"]) || "").trim();
    const entityId = String(getExecutiveValue_(row, headers, ["ENTITY_ID"]) || "").trim();
    const actorName = String(getExecutiveValue_(row, headers, ["ACTOR_NAME"]) || "").trim();
    const actorEmail = String(getExecutiveValue_(row, headers, ["ACTOR_EMAIL"]) || "").trim();
    const moduleName = String(getExecutiveValue_(row, headers, ["MODULE"]) || "").trim();

    items.push({
      timestamp: formatExecutiveActivityDate_(timestamp),
      companyId: companyId,
      companyName: companyMap[companyId] || companyId || "ALL",
      title: formatExecutiveActionLabel_(action),
      meta: buildExecutiveActivityMeta_(entityType, entityId, actorName || actorEmail, moduleName),
      actor: actorName || actorEmail || "Sistema",
      tone: getExecutiveActivityTone_(action, moduleName),
      sortTime: timestamp.getTime()
    });
  }
}

function collectExecutiveWorkOrderLogActivity_(ss, companyMap, requestedCompany, items) {
  const sh = ss.getSheetByName(CFG.SHEET_WO_LOG);
  if (!sh || sh.getLastRow() < 1) return;

  const data = sh.getDataRange().getValues();
  if (!data.length) return;

  const firstRow = data[0].map(function(h) { return String(h || "").trim().toUpperCase(); });
  const hasHeaders = firstRow.indexOf("TIMESTAMP") !== -1 || firstRow.indexOf("ACTION") !== -1;
  const startIndex = hasHeaders ? 1 : 0;
  const headers = hasHeaders ? data[0].map(function(h) { return String(h || "").trim(); }) : [];

  for (let i = startIndex; i < data.length; i++) {
    const row = data[i];
    const rowCompany = hasHeaders ? getExecutiveValue_(row, headers, ["COMPANY_ID"]) : row[0];
    const companyId = String(rowCompany || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();
    if (requestedCompany && companyId !== requestedCompany) continue;

    const timestamp = parseExecutiveDate_(hasHeaders ? getExecutiveValue_(row, headers, ["TIMESTAMP"]) : row[1]);
    if (!timestamp) continue;

    const woNumber = String(hasHeaders ? getExecutiveValue_(row, headers, ["WO_NUMBER"]) : row[2] || "").trim();
    const action = String(hasHeaders ? getExecutiveValue_(row, headers, ["ACTION"]) : row[3] || "").trim();
    if (!action) continue;

    const oldStatus = String(hasHeaders ? getExecutiveValue_(row, headers, ["OLD_STATUS"]) : row[4] || "").trim();
    const newStatus = String(hasHeaders ? getExecutiveValue_(row, headers, ["NEW_STATUS"]) : row[5] || "").trim();
    const actor = String(hasHeaders ? getExecutiveValue_(row, headers, ["USER", "ACTOR"]) : row[6] || "").trim();

    items.push({
      timestamp: formatExecutiveActivityDate_(timestamp),
      companyId: companyId,
      companyName: companyMap[companyId] || companyId,
      title: formatExecutiveActionLabel_(action),
      meta: buildExecutiveWorkOrderActivityMeta_(woNumber, oldStatus, newStatus, actor),
      actor: actor || "Sistema",
      tone: getExecutiveActivityTone_(action, "WO_LOG"),
      sortTime: timestamp.getTime()
    });
  }
}

function buildExecutiveActivityMeta_(entityType, entityId, actor, moduleName) {
  const parts = [];
  if (entityType || entityId) parts.push([entityType, entityId].filter(Boolean).join(" "));
  if (actor) parts.push("por " + actor);
  if (moduleName) parts.push(moduleName);
  return parts.join(" | ");
}

function buildExecutiveWorkOrderActivityMeta_(woNumber, oldStatus, newStatus, actor) {
  const parts = [];
  if (woNumber) parts.push("WO " + woNumber);
  if (oldStatus || newStatus) parts.push([oldStatus || "-", newStatus || "-"].join(" -> "));
  if (actor) parts.push("por " + actor);
  return parts.join(" | ");
}

function formatExecutiveActionLabel_(action) {
  const key = String(action || "").trim().toUpperCase();
  const labels = {
    USER_CREATED: "Usuario creado",
    USER_UPDATED: "Usuario actualizado",
    USER_SOFT_DELETED: "Usuario desactivado",
    STORE_CREATED: "Tienda creada",
    STORE_UPDATED: "Tienda actualizada",
    SUPERVISOR_STORES_ASSIGNED: "Tiendas asignadas a supervisor",
    WORK_ORDER_UPDATED: "Orden actualizada",
    WORK_ORDER_SENT_TO_TECH: "Orden enviada a tecnico",
    WORK_ORDER_SOFT_DELETED: "Orden desactivada",
    STATUS_UPDATED_FROM_TECH_APP: "Estado actualizado por tecnico",
    ORDER_CREATED: "Orden creada",
    ORDER_CREATED_FROM_APP: "Orden creada desde app",
    PM_REPORT_GENERATED: "Reporte PM generado",
    COMPANY_CREATED: "Empresa creada",
    COMPANY_UPDATED: "Empresa actualizada"
  };

  if (labels[key]) return labels[key];

  return key
    .replace(/_/g, " ")
    .toLowerCase()
    .replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

function getExecutiveActivityTone_(action, moduleName) {
  const text = String(action || moduleName || "").toUpperCase();
  if (text.indexOf("DELETE") !== -1 || text.indexOf("DELETED") !== -1 || text.indexOf("ERROR") !== -1) return "danger";
  if (text.indexOf("COMPLETED") !== -1 || text.indexOf("GENERATED") !== -1 || text.indexOf("CREATED") !== -1) return "success";
  if (text.indexOf("SENT") !== -1 || text.indexOf("ASSIGNED") !== -1 || text.indexOf("UPDATED") !== -1) return "info";
  return "neutral";
}

function formatExecutiveActivityDate_(dateValue) {
  return Utilities.formatDate(dateValue, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a");
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
