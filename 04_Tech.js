// =====================================================
// FILE: 04_Tech.gs
// =====================================================

function getTechnicianOrders(techName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_WORK_ORDERS);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(String);

  const data = values.slice(1).map(function(row, i) {
    const obj = {};
    headers.forEach(function(h, c) {
      let value = row[c];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, CFG.TIMEZONE, "MM/dd/yyyy hh:mm a");
      }
      obj[h] = value;
    });

    obj.ROW_NUMBER = i + 2;
    enrichWorkOrderObject_(obj);
    return obj;
  });

  return data.filter(function(o) {
    const assignedTech = String(o.TECHNICIAN || "").toLowerCase();
    const selectedTech = String(techName || "").toLowerCase();
    const status = String(o.STATUS || "").toUpperCase();

    return assignedTech.includes(selectedTech) && status !== "CLOSED" && status !== "COMPLETED";
  }).reverse();
}

function updateTechOrderStatus(rowNumber, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  const companyId = getCellByHeader_(sh, rowNumber, headers, "COMPANY_ID") || CFG.DEFAULT_COMPANY_ID;
  const woNumber = getCellByHeader_(sh, rowNumber, headers, "WO_NUMBER");
  const oldStatus = getCellByHeader_(sh, rowNumber, headers, "STATUS");

  setCellByHeader_(sh, rowNumber, headers, "STATUS", newStatus);

  if (newStatus === "COMPLETED") {
    setCellByHeader_(sh, rowNumber, headers, "DATE_COMPLETED", new Date());

    createEconomyFromWO(companyId, woNumber);

    addNotification_(companyId, "David Dominguez", woNumber, "DONE", "✅ Completed " + woNumber);
    addNotification_(companyId, "Sarahi", woNumber, "BILLING", "💰 Ready for billing " + woNumber);
  }

  if (newStatus === "REQUEST PARTS") {
    addNotification_(companyId, "Dayre", woNumber, "PARTS", "🔧 Parts requested for " + woNumber);
  }

  if (newStatus === "PARTS IN TRANSIT") {
    addNotification_(companyId, "Dayre", woNumber, "TRANSIT", "🚚 Parts in transit for " + woNumber);
  }

  addLog_(companyId, woNumber, "STATUS UPDATED FROM TECH APP", oldStatus, newStatus, "Technician App", "");
  return true;
}

function getMyTechInvoiceSummary(companyId, techName, month, year) {
  companyId = String(companyId || "").trim().toUpperCase();
  techName = String(techName || "").trim();
  month = Number(month);
  year = Number(year);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY || "ECONOMY");
  const shPM = ss.getSheetByName("PM_ECONOMY");
  const shInv = ss.getSheetByName("INVOICES");
  const shWO = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  if (!shEco) throw new Error("No existe la hoja ECONOMY.");
  if (!shInv) throw new Error("No existe la hoja INVOICES.");
  if (!shWO) throw new Error("No existe la hoja WORK_ORDERS.");

  const ecoData = shEco.getDataRange().getValues();
  const pmData = shPM ? shPM.getDataRange().getValues() : [];
  const invData = shInv.getDataRange().getValues();
  const woData = shWO.getDataRange().getValues();

  if (ecoData.length < 2 && pmData.length < 2) {
    return { rows: [], totals: emptyTechInvoiceTotals_() };
  }

  const eh = ecoData.length ? ecoData[0].map(String) : [];
  const ph = pmData.length ? pmData[0].map(String) : [];
  const ih = invData[0].map(String);
  const wh = woData[0].map(String);

  const invMap = {};
  const woMap = {};

  for (let i = 1; i < invData.length; i++) {
    const row = invData[i];
    const wo = String(getRowValue_(row, ih, "WO_NUMBER") || "").trim();
    if (!wo) continue;

    invMap[wo] = {
      invoiceNumber: getInvoiceNumberValue_(row, ih),
      dateInvoice: getInvoiceDateValue_(row, ih),
      materialCost: getMaterialCostValue_(row, ih),
      laborAmount: getLaborAmountValue_(row, ih),
      invoiceTotal: getInvoiceTotalValue_(row, ih),
      taxAmount: getTaxAmountValue_(row, ih),
      pdfEnUrl: getRowValue_(row, ih, "PDF_EN_URL"),
      pdfEsUrl: getRowValue_(row, ih, "PDF_ES_URL")
    };
  }

  for (let i = 1; i < woData.length; i++) {
    const row = woData[i];
    const wo = String(getRowValue_(row, wh, "WO_NUMBER") || "").trim();
    if (!wo) continue;

    woMap[wo] = {
      storeNumber: getRowValue_(row, wh, "NSN"),
      equipment:
        getRowValue_(row, wh, "REPORTED_EQUIPMENT") ||
        getRowValue_(row, wh, "REPORTED_EQUIPMENT_EN"),
      problem:
        getRowValue_(row, wh, "REPORTED_PROBLEM_ES") ||
        getRowValue_(row, wh, "REPORTED_PROBLEM_EN") ||
        getRowValue_(row, wh, "REPORTED_PROBLEM_ORIGINAL")
    };
  }

  const rows = [];

  // =====================================================
  // ECONOMY REGULAR
  // =====================================================

  if (ecoData.length > 1) {
    for (let i = 1; i < ecoData.length; i++) {
      const row = ecoData[i];

      const rowCompany = String(getRowValue_(row, eh, "COMPANY_ID") || "").trim().toUpperCase();
      const woNumber = String(getRowValue_(row, eh, "WO_NUMBER") || "").trim();
      const techniciansText = String(getRowValue_(row, eh, "TECHNICIANS") || "");
      const techniciansLower = techniciansText.toLowerCase();

      if (companyId && rowCompany !== companyId) continue;
      if (!techniciansLower.includes(techName.toLowerCase())) continue;

      const dateInvoiceRaw = getInvoiceDateValue_(row, eh) || getRowValue_(row, eh, "DATE_COMPLETED");
      const d = new Date(dateInvoiceRaw);

      if (!isNaN(d.getTime())) {
        const rowMonth = d.getMonth() + 1;
        const rowYear = d.getFullYear();

        if (month && rowMonth !== month) continue;
        if (year && rowYear !== year) continue;
      }

      const inv = invMap[woNumber] || {};
      const wo = woMap[woNumber] || {};

      const hours = getTechHoursFromPayDetail_(row, eh, techName);
      const laborPay = getTechLaborPayFromEconomy_(row, eh, techName);

      const techList = techniciansText
        .split(",")
        .map(function(t) {
          return t.trim();
        })
        .filter(Boolean);

      const materialBase = Number(inv.materialCost || getRowValue_(row, eh, "COST") || 0);

      const materialBonus = techList.length
        ? (materialBase * 0.10) / techList.length
        : 0;

      const totalEarned = laborPay + materialBonus;

      rows.push({
        woNumber: woNumber,
        storeNumber: wo.storeNumber || getRowValue_(row, eh, "NSN") || "",
        invoiceNumber: getInvoiceNumberValue_(row, eh) || inv.invoiceNumber || "",
        equipment: wo.equipment || "",
        problem: wo.problem || "",
        hours: hours,
        laborPay: laborPay,
        materialBase: materialBase,
        materialBonus: materialBonus,
        totalEarned: totalEarned,
        pdfUrl: inv.pdfEnUrl || inv.pdfEsUrl || ""
      });
    }
  }

  // =====================================================
  // PM_ECONOMY
  // =====================================================

  if (pmData.length > 1) {
    for (let i = 1; i < pmData.length; i++) {
      const row = pmData[i];

      const rowCompany = String(getRowValue_(row, ph, "COMPANY_ID") || "").trim().toUpperCase();
      const woNumber = String(getRowValue_(row, ph, "WO_NUMBER") || "").trim();
      const techniciansText = String(getRowValue_(row, ph, "TECHNICIANS") || "");
      const techniciansLower = techniciansText.toLowerCase();

      if (companyId && rowCompany !== companyId) continue;
      if (!techniciansLower.includes(techName.toLowerCase())) continue;

      const dateRaw = getRowValue_(row, ph, "DATE_COMPLETED") || getRowValue_(row, ph, "DATE_INVOICE");
      const d = new Date(dateRaw);

      if (!isNaN(d.getTime())) {
        const rowMonth = d.getMonth() + 1;
        const rowYear = d.getFullYear();

        if (month && rowMonth !== month) continue;
        if (year && rowYear !== year) continue;
      }

      const inv = invMap[woNumber] || {};
      const wo = woMap[woNumber] || {};

      const laborPay = Number(
        getRowValue_(row, ph, "TECH_LABOR_COST") ||
        getRowValue_(row, ph, "TECH_LABOR_PAY") ||
        80
      );

      rows.push({
        woNumber: woNumber + " PM",
        storeNumber: wo.storeNumber || getRowValue_(row, ph, "NSN") || "",
        invoiceNumber: getRowValue_(row, ph, "INVOICE_NUMBER") || inv.invoiceNumber || "",
        equipment: "PM",
        problem: "Preventive Maintenance",
        hours: 0,
        laborPay: laborPay,
        materialBase: 0,
        materialBonus: 0,
        totalEarned: laborPay,
        pdfUrl: inv.pdfEnUrl || inv.pdfEsUrl || ""
      });
    }
  }

  const totals = rows.reduce(function(acc, r) {
    acc.invoices++;
    acc.hours += Number(r.hours || 0);
    acc.laborPay += Number(r.laborPay || 0);
    acc.materialBonus += Number(r.materialBonus || 0);
    acc.totalEarned += Number(r.totalEarned || 0);
    return acc;
  }, emptyTechInvoiceTotals_());

  return {
    rows: rows,
    totals: totals
  };
}

function getTechLaborPayFromEconomy_(row, headers, techName) {
  const detail = String(getRowValue_(row, headers, "TECH_PAY_DETAIL") || "").trim();
  const totalLaborCost = parseMoneyFlexible_(
    getHeaderValueFlexible_(row, headers, [
      "TECH_LABOR_PAY",
      "TECH_LABOR_COST"
    ])
  );

  if (!detail) {
    return totalLaborCost;
  }

  const name = String(techName || "").trim().toLowerCase();

  const parts = detail.split(/[,;|]/).map(function(x) {
    return x.trim();
  }).filter(Boolean);

  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const low = p.toLowerCase();

    if (!low.includes(name)) continue;

    const nums = p.match(/-?\d+(?:\.\d+)?/g);
    if (nums && nums.length) {
      return Number(nums[nums.length - 1]) || 0;
    }
  }

  return totalLaborCost;
}

function getTechHoursFromPayDetail_(row, headers, techName) {
  const detail = String(getRowValue_(row, headers, "TECH_PAY_DETAIL") || "").trim();

  if (!detail) {
    return getHoursValue_(row, headers);
  }

  const name = String(techName || "").trim().toLowerCase();

  const parts = detail.split(/[,;|]/).map(function(x) {
    return x.trim();
  }).filter(Boolean);

  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const low = p.toLowerCase();

    if (!low.includes(name)) continue;

    const hourMatch = p.match(/(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hour|hours|hora|horas)/i);

    if (hourMatch) {
      return Number(hourMatch[1]) || 0;
    }
  }

  return getHoursValue_(row, headers);
}

function getRowValue_(row, headers, name) {
  const idx = headers.indexOf(name);
  return idx >= 0 ? row[idx] : "";
}

function formatDateTech_(v) {
  if (!v) return "";
  const d = new Date(v);
  if (isNaN(d.getTime())) return String(v || "");
  return Utilities.formatDate(d, CFG.TIMEZONE, "MM/dd/yyyy");
}

function emptyTechInvoiceTotals_() {
  return {
    invoices: 0,
    hours: 0,
    laborPay: 0,
    materialBonus: 0,
    totalEarned: 0
  };
}