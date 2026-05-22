// =====================================================
// FILE: 21_Admin_Portals.gs
// =====================================================

function getTechPortalAdminData(companyId, role, sessionToken) {
  companyId = normalizeAdminCompanyId_(companyId);
  role = String(role || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN"], companyId);
  role = session.role;

  if (!isAdminPortalRole_(role)) {
    throw new Error("No autorizado.");
  }

  const techs = getAdminUsersByRoles_(companyId, ["TECH"]);
  const orderStats = getTechOrderStats_(companyId, techs);

  techs.forEach(function(tech) {
    const key = normalizeAdminName_(tech.name);
    tech.stats = orderStats.byTech[key] || emptyTechAdminStats_();
  });

  return {
    companyId: companyId,
    companyName: getCompanyName(companyId),
    techs: techs,
    unassignedOrders: orderStats.unassignedOrders,
    totals: orderStats.totals
  };
}

function getSupervisorPortalAdminData(companyId, role, sessionToken) {
  companyId = normalizeAdminCompanyId_(companyId);
  role = String(role || "").trim().toUpperCase();
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN"], companyId);
  role = session.role;

  if (!isAdminPortalRole_(role)) {
    throw new Error("No autorizado.");
  }

  const supervisorUsers = getAdminUsersByRoles_(companyId, ["SUPERVISOR"]);
  const stores = getAdminStores_(companyId);
  const groups = {};
  const unassignedStores = [];

  supervisorUsers.forEach(function(user) {
    const key = normalizeAdminName_(user.name);
    if (!key) return;

    groups[key] = groups[key] || {
      name: user.name,
      email: user.email,
      userEmail: user.email,
      active: user.active,
      userRow: user.rowNumber,
      stores: []
    };
  });

  stores.forEach(function(store) {
    const name = String(store.SUPERVISOR_NAME || "").trim();
    const key = normalizeAdminName_(name);

    if (!key) {
      unassignedStores.push(store);
      return;
    }

    groups[key] = groups[key] || {
      name: name,
      email: String(store.SUPERVISOR_EMAIL || "").trim(),
      userEmail: "",
      active: "",
      userRow: "",
      stores: []
    };

    if (!groups[key].email && store.SUPERVISOR_EMAIL) {
      groups[key].email = String(store.SUPERVISOR_EMAIL || "").trim();
    }

    groups[key].stores.push(store);
  });

  const supervisors = Object.keys(groups)
    .sort()
    .map(function(key) {
      const s = groups[key];
      s.storeCount = s.stores.length;
      return s;
    });

  return {
    companyId: companyId,
    companyName: getCompanyName(companyId),
    supervisors: supervisors,
    unassignedStores: unassignedStores,
    totals: {
      supervisors: supervisors.length,
      stores: stores.length,
      unassignedStores: unassignedStores.length
    }
  };
}

function saveSupervisorStoreAssignment(data, sessionToken) {
  data = data || {};

  const companyId = normalizeAdminCompanyId_(data.companyId);
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN"], companyId);
  const rowNumbers = (data.rowNumbers || []).map(Number).filter(function(n) {
    return n && n > 1;
  });

  const supervisorName = String(data.supervisorName || "").trim();
  const supervisorEmail = String(data.supervisorEmail || "").trim();

  if (!companyId) throw new Error("COMPANY_ID requerido.");
  if (!rowNumbers.length) throw new Error("Seleccione al menos una tienda.");
  if (!supervisorName) throw new Error("Supervisor name requerido.");

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_STORES);
  if (!sh) throw new Error("No existe la hoja STORES.");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  const idxCompany = headers.indexOf("COMPANY_ID");
  if (idxCompany === -1) throw new Error("STORES debe tener COMPANY_ID.");

  ensureAdminStoreColumn_(sh, headers, "SUPERVISOR_NAME");
  ensureAdminStoreColumn_(sh, headers, "SUPERVISOR_EMAIL");

  const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  let updated = 0;

  rowNumbers.forEach(function(row) {
    const rowCompany = String(getCellByHeader_(sh, row, freshHeaders, "COMPANY_ID") || "")
      .trim()
      .toUpperCase();

    if (rowCompany !== companyId) return;

    setCellByHeader_(sh, row, freshHeaders, "SUPERVISOR_NAME", supervisorName);
    setCellByHeader_(sh, row, freshHeaders, "SUPERVISOR_EMAIL", supervisorEmail);
    updated++;
  });

  addAuditLog_("SUPERVISOR_ADMIN", "SUPERVISOR_STORES_ASSIGNED", companyId, "SUPERVISOR", supervisorName, session, {
    supervisorEmail: supervisorEmail,
    requestedRows: rowNumbers,
    updated: updated
  });

  return {
    success: true,
    updated: updated
  };
}

function getAdminUsersByRoles_(companyId, roles) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_USERS);
  if (!sh) throw new Error("No existe USERS.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(function(h) {
    return String(h).trim().toUpperCase();
  });

  const idxEmail = headers.indexOf("EMAIL");
  const idxName = headers.indexOf("NAME");
  const idxRole = headers.indexOf("ROLE");
  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxActive = headers.indexOf("ACTIVE");
  const idxRate = headers.indexOf("HOURLY_RATE");

  const allowedRoles = roles.map(function(r) {
    return String(r || "").trim().toUpperCase();
  });

  return data.slice(1).map(function(row, i) {
    return {
      rowNumber: i + 2,
      email: idxEmail >= 0 ? String(row[idxEmail] || "").trim() : "",
      name: idxName >= 0 ? String(row[idxName] || "").trim() : "",
      role: idxRole >= 0 ? String(row[idxRole] || "").trim().toUpperCase() : "",
      companyId: idxCompany >= 0 ? String(row[idxCompany] || "").trim().toUpperCase() : "",
      active: idxActive >= 0 ? String(row[idxActive] || "").trim().toUpperCase() : "",
      hourlyRate: idxRate >= 0 ? row[idxRate] || "" : ""
    };
  }).filter(function(user) {
    if (user.companyId !== companyId) return false;
    if (allowedRoles.indexOf(user.role) === -1) return false;
    return user.name || user.email;
  });
}

function getTechOrderStats_(companyId, techs) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_WORK_ORDERS);

  const result = {
    byTech: {},
    unassignedOrders: [],
    totals: {
      techs: techs.length,
      openOrders: 0,
      completedOrders: 0,
      unassignedOrders: 0
    }
  };

  techs.forEach(function(tech) {
    result.byTech[normalizeAdminName_(tech.name)] = emptyTechAdminStats_();
  });

  if (!sh) return result;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return result;

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  data.slice(1).forEach(function(row, i) {
    const obj = {};
    headers.forEach(function(h, c) {
      obj[h] = row[c];
    });

    const rowCompany = String(obj.COMPANY_ID || "").trim().toUpperCase();
    if (rowCompany !== companyId) return;

    const status = String(obj.STATUS || "").trim().toUpperCase();
    const isClosed = status === "CLOSED" || status === "COMPLETED";
    const assigned = String(obj.TECHNICIAN || "").trim();

    if (!assigned && !isClosed) {
      result.unassignedOrders.push({
        rowNumber: i + 2,
        woNumber: obj.WO_NUMBER || "",
        nsn: obj.NSN || "",
        status: obj.STATUS || "",
        priority: obj.ORDER_PRIORITY || "",
        problem: obj.REPORTED_PROBLEM_ES || obj.REPORTED_PROBLEM_EN || obj.REPORTED_PROBLEM_ORIGINAL || ""
      });
      result.totals.unassignedOrders++;
      return;
    }

    techs.forEach(function(tech) {
      const key = normalizeAdminName_(tech.name);
      const stats = result.byTech[key];

      if (!key || !stats) return;
      if (assigned.toLowerCase().indexOf(tech.name.toLowerCase()) === -1) return;

      stats.totalOrders++;

      if (isClosed) {
        stats.completedOrders++;
        result.totals.completedOrders++;
      } else {
        stats.openOrders++;
        result.totals.openOrders++;
      }

      if (String(obj.WO_TYPE || "").toUpperCase().indexOf("PM") !== -1) {
        stats.pmOrders++;
      }
    });
  });

  return result;
}

function getAdminStores_(companyId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_STORES);
  if (!sh) throw new Error("No existe STORES.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  return data.slice(1).map(function(row, i) {
    const obj = {};
    headers.forEach(function(h, c) {
      obj[h] = row[c];
    });
    obj.ROW_NUMBER = i + 2;
    return obj;
  }).filter(function(store) {
    return String(store.COMPANY_ID || "").trim().toUpperCase() === companyId;
  });
}

function emptyTechAdminStats_() {
  return {
    totalOrders: 0,
    openOrders: 0,
    completedOrders: 0,
    pmOrders: 0
  };
}

function normalizeAdminCompanyId_(companyId) {
  return String(companyId || CFG.DEFAULT_COMPANY_ID || "").trim().toUpperCase();
}

function normalizeAdminName_(name) {
  return String(name || "").trim().toUpperCase();
}

function isAdminPortalRole_(role) {
  role = String(role || "").trim().toUpperCase();
  return role === "OWNER" || role === "ADMIN";
}

function ensureAdminStoreColumn_(sheet, headers, columnName) {
  if (headers.indexOf(columnName) !== -1) return;
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue(columnName);
  headers.push(columnName);
}
