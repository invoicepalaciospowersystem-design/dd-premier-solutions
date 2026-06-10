

// =====================================================
// FILE: 14_Tech_Summary.gs
// =====================================================

function getTechMonthlySummary(companyId, month, year, sessionToken) {
  return getTechSummaryByPeriod_(companyId, "monthly", month, year, sessionToken);
}

function getTechPeriodSummary(companyId, periodMode, month, year, sessionToken) {
  return getTechSummaryByPeriod_(companyId, periodMode, month, year, sessionToken);
}

function getTechSummaryByPeriod_(companyId, periodMode, month, year, sessionToken) {

  companyId = String(companyId || "").trim().toUpperCase();
  periodMode = String(periodMode || "monthly").trim().toLowerCase();
  periodMode = (periodMode === "annual" || periodMode === "yearly") ? "annual" : "monthly";
  month = Number(month);
  year = Number(year);
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA", "TECH"], companyId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);
  const shPM = ss.getSheetByName("PM_ECONOMY");
  const shUsers = ss.getSheetByName(CFG.SHEET_USERS);

  if (!shEco) throw new Error("No existe ECONOMY.");
  if (!shUsers) throw new Error("No existe USERS.");

  const ecoData = shEco.getDataRange().getValues();
  const pmData = shPM ? shPM.getDataRange().getValues() : [];
  const userData = shUsers.getDataRange().getValues();

  if (ecoData.length < 2) return [];

  const eh = ecoData[0].map(h => String(h).trim());
  const uh = userData[0].map(h => String(h).trim());

  const eCompany = eh.indexOf("COMPANY_ID");
  const eDate = eh.indexOf("DATE_COMPLETED");
  const eWO = eh.indexOf("WO_NUMBER");
  const eHours = eh.indexOf("HORAS");
  const eCost = eh.indexOf("COST");
  const eTechs = eh.indexOf("TECHNICIANS");
  const eInvSource = eh.indexOf("INV_SOURCE");

  const uName = uh.indexOf("NAME");
  const uCompany = uh.indexOf("COMPANY_ID");
  const uRate = uh.indexOf("HOURLY_RATE");

  if ([eCompany, eDate, eWO, eHours, eCost, eTechs].includes(-1)) {
    throw new Error("ECONOMY debe tener COMPANY_ID, DATE_COMPLETED, WO_NUMBER, HORAS, COST, TECHNICIANS.");
  }

  if ([uName, uCompany, uRate].includes(-1)) {
    throw new Error("USERS debe tener NAME, COMPANY_ID, HOURLY_RATE.");
  }

  // =====================================================
  // MAPA DE RATES
  // =====================================================

  const rates = {};

  for (let i = 1; i < userData.length; i++) {

    const name = String(userData[i][uName] || "").trim();
    const comp = String(userData[i][uCompany] || "").trim().toUpperCase();
    const rate = Number(userData[i][uRate] || 0);

    if (name && comp === companyId) {
      rates[name.toLowerCase()] = rate;
    }
  }

  // =====================================================
  // SUMMARY
  // =====================================================

  const summary = {};

  const ddKey = "d&d premier solutions corp";

  summary[ddKey] = {
    name: "D&D PREMIER SOLUTIONS CORP",
    rate: "",
    orders: 0,
    hours: 0,
    materialCost: 0,
    laborPay: 0,
    materialBonus: 0,
    totalPay: 0,
    workOrders: []
  };

  // =====================================================
  // ECONOMY NORMAL
  // =====================================================

  for (let i = 1; i < ecoData.length; i++) {

    const row = ecoData[i];
    if (isSoftDeletedRow_(row, eh)) continue;

    const comp = String(row[eCompany] || "").trim().toUpperCase();

    if (comp !== companyId) continue;

    const d = row[eDate] instanceof Date
      ? row[eDate]
      : new Date(row[eDate]);

    if (isNaN(d.getTime())) continue;

    if (!techSummaryDateMatches_(d, periodMode, month, year)) continue;

    const wo = String(row[eWO] || "").trim();

    const totalHours = Number(row[eHours] || 0);
    const totalCost = Number(row[eCost] || 0);

    const techs = String(row[eTechs] || "")
      .split(",")
      .map(t => t.trim())
      .filter(Boolean);

    // =====================================================
    // TECNICOS
    // =====================================================

    if (techs.length) {

      const hoursPerTech = totalHours / techs.length;
      const costPerTech = totalCost / techs.length;

      techs.forEach(function(name) {

        const key = name.toLowerCase();

        const rate = Number(rates[key] || 0);

        const laborPay = hoursPerTech * rate;
        const materialBonus = costPerTech * 0.10;
        const totalPay = laborPay + materialBonus;

        if (!summary[key]) {

          summary[key] = {
            name: name,
            rate: rate,
            orders: 0,
            hours: 0,
            materialCost: 0,
            laborPay: 0,
            materialBonus: 0,
            totalPay: 0,
            workOrders: []
          };
        }

        summary[key].orders += 1;
        summary[key].hours += hoursPerTech;
        summary[key].materialCost += costPerTech;
        summary[key].laborPay += laborPay;
        summary[key].materialBonus += materialBonus;
        summary[key].totalPay += totalPay;
        summary[key].workOrders.push(wo);
      });
    }

    // =====================================================
    // D&D PREMIER
    // =====================================================

    let ddHoursPay = 0;

    if (totalHours > 0) {
      ddHoursPay = 100 + (Math.max(totalHours - 1, 0) * 80);
    }

    const invSource = eInvSource >= 0
      ? String(row[eInvSource] || "").toUpperCase()
      : "";

    let partsRate = 0;

    if (invSource === "YOEL") partsRate = 0.25;
    if (invSource === "DAVID") partsRate = 0.75;
    if (invSource === "MISCELANEAS") partsRate = 0.50;

    const ddPartsPay = totalCost * partsRate;

    summary[ddKey].orders += 1;
    summary[ddKey].hours += totalHours;
    summary[ddKey].materialCost += totalCost;
    summary[ddKey].laborPay += ddHoursPay;
    summary[ddKey].materialBonus += ddPartsPay;
    summary[ddKey].totalPay += ddHoursPay + ddPartsPay;
    summary[ddKey].workOrders.push(wo);
  }

  // =====================================================
  // PM_ECONOMY
  // =====================================================

  if (pmData.length > 1) {

    const ph = pmData[0].map(h => String(h).trim());

    const pCompany = ph.indexOf("COMPANY_ID");
    const pDate = ph.indexOf("DATE_COMPLETED");
    const pWO = ph.indexOf("WO_NUMBER");
    const pTechs = ph.indexOf("TECHNICIANS");
    const pPay = ph.indexOf("TECH_LABOR_COST");

    for (let i = 1; i < pmData.length; i++) {

      const row = pmData[i];
      if (isSoftDeletedRow_(row, ph)) continue;

      const comp = String(row[pCompany] || "").trim().toUpperCase();

      if (comp !== companyId) continue;

      const d = row[pDate] instanceof Date
        ? row[pDate]
        : new Date(row[pDate]);

      if (isNaN(d.getTime())) continue;

      if (!techSummaryDateMatches_(d, periodMode, month, year)) continue;

      const techName = String(row[pTechs] || "").trim();

      if (!techName) continue;

      const pay = Number(row[pPay] || 0);
      const wo = String(row[pWO] || "").trim();

      const key = techName.toLowerCase();

      if (!summary[key]) {

        summary[key] = {
          name: techName,
          rate: "PM",
          orders: 0,
          hours: 0,
          materialCost: 0,
          laborPay: 0,
          materialBonus: 0,
          totalPay: 0,
          workOrders: []
        };
      }

      summary[key].orders += 1;
      summary[key].laborPay += pay;
      summary[key].totalPay += pay;
      summary[key].workOrders.push(wo + " PM");
    }
  }

  // =====================================================
  // HORAS MANUALES
  // =====================================================

  addManualTechHoursToSummary_(summary, companyId, periodMode, month, year);

  // =====================================================
  // HISTORICO VIEJO
  // =====================================================

  if (summary[ddKey].totalPay === 0) {

    const NEW_SYSTEM_START_LABEL = "2026-06";

    if (periodMode === "annual") {

      summary["d&d premier solutions corp"] = {
        name: "D&D PREMIER SOLUTIONS CORP",
        rate: "",
        orders: 0,
        hours: 0,
        materialCost: 0,
        laborPay: 0,
        materialBonus: 0,
        totalPay: 0,
        workOrders: []
      };

      for (let oldMonth = 1; oldMonth <= 12; oldMonth++) {
        const selectedLabel = year + "-" + String(oldMonth).padStart(2, "0");
        if (selectedLabel < NEW_SYSTEM_START_LABEL) {
          addOldDDMonthlyToSummary_(summary, oldMonth, year);
        }
      }
    } else {

      const selectedLabel = year + "-" + String(month).padStart(2, "0");

      if (selectedLabel < NEW_SYSTEM_START_LABEL) {

        summary["d&d premier solutions corp"] = {
          name: "D&D PREMIER SOLUTIONS CORP",
          rate: "",
          orders: 0,
          hours: 0,
          materialCost: 0,
          laborPay: 0,
          materialBonus: 0,
          totalPay: 0,
          workOrders: []
        };

        addOldDDMonthlyToSummary_(summary, month, year);
      }
    }
  }

  let rows = Object.keys(summary)
    .map(k => summary[k])
    .sort((a, b) => Number(b.totalPay || 0) - Number(a.totalPay || 0));

  if (String(session.role || "").trim().toUpperCase() === "TECH") {
    const ownName = normalizeIdentity_(session.name);
    rows = rows.filter(function(row) {
      return normalizeIdentity_(row.name) === ownName;
    });
  }

  return rows;
}

function techSummaryDateMatches_(dateValue, periodMode, month, year) {
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime())) return false;

  if (String(periodMode || "monthly") === "annual") {
    return dateValue.getFullYear() === Number(year);
  }

  return dateValue.getFullYear() === Number(year) &&
    (dateValue.getMonth() + 1) === Number(month);
}

// =====================================================
// HISTORICO VIEJO
// =====================================================

function addOldDDMonthlyToSummary_(summary, month, year) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("DD_OLD_MONTHLY");

  if (!sh) return;

  const data = sh.getDataRange().getValues();

  if (data.length < 2) return;

  const headers = data[0].map(h => String(h).trim());

  const idxMes = headers.indexOf("MES");
  const idxHoras = headers.indexOf("TOTAL_HORAS");
  const idxParts = headers.indexOf("TOTAL_PARTS");
  const idxCobroHoras = headers.indexOf("TOTAL_COBRO_HORAS");
  const idxCobroParts = headers.indexOf("TOTAL_COBRO_PARTS");
  const idxTotal = headers.indexOf("TOTAL_COMPANIA");

  if ([idxMes, idxHoras, idxParts, idxCobroHoras, idxCobroParts, idxTotal].includes(-1)) return;

  const targetLabel = year + "-" + String(month).padStart(2, "0");

  const ddKey = "d&d premier solutions corp";

  if (!summary[ddKey]) {

    summary[ddKey] = {
      name: "D&D PREMIER SOLUTIONS CORP",
      rate: "",
      orders: 0,
      hours: 0,
      materialCost: 0,
      laborPay: 0,
      materialBonus: 0,
      totalPay: 0,
      workOrders: []
    };
  }

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    let mes = row[idxMes];

    if (mes instanceof Date) {
      mes = Utilities.formatDate(mes, CFG.TIMEZONE, "yyyy-MM");
    } else {
      mes = String(mes || "").trim();
    }

    if (mes !== targetLabel) continue;

    summary[ddKey].hours += Number(row[idxHoras] || 0);
    summary[ddKey].materialCost += Number(row[idxParts] || 0);
    summary[ddKey].laborPay += Number(row[idxCobroHoras] || 0);
    summary[ddKey].materialBonus += Number(row[idxCobroParts] || 0);
    summary[ddKey].totalPay += Number(row[idxTotal] || 0);

    summary[ddKey].workOrders.push("Histórico viejo " + targetLabel);
  }
}

// =====================================================
// SOLO UN TECNICO
// =====================================================

function getMyTechMonthlySummary(companyId, techName, month, year, sessionToken) {
  const session = requireNamedSession_(sessionToken, ["TECH", "OWNER", "ADMIN"], companyId, techName, "tecnico");
  if (String(session.role || "").trim().toUpperCase() === "TECH") {
    techName = String(session.name || "").trim();
    companyId = String(session.companyId || companyId || "").trim().toUpperCase();
  }

  const all = getTechMonthlySummary(companyId, month, year, sessionToken);

  const target = String(techName || "").trim().toLowerCase();

  const found = all.find(function(t) {
    return String(t.name || "").trim().toLowerCase() === target;
  });

  return found || {
    name: techName,
    rate: 0,
    orders: 0,
    hours: 0,
    materialCost: 0,
    laborPay: 0,
    materialBonus: 0,
    totalPay: 0,
    workOrders: []
  };
}
