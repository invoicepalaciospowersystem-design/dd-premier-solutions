// =====================================================
// FILE: 06_Stores.gs
// =====================================================

function getStoreByNSN_(nsn, companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_STORES);
  if (!sh) return {};

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return {};

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  const idxCompany = headers.indexOf("COMPANY_ID");
  const idxNSN = headers.indexOf("NSN #");
  const idxClient = headers.indexOf("CLIENTE");
  const idxAddress = headers.indexOf("ADDRESS");
  const idxCity = headers.indexOf("CITY");
  const idxState = headers.indexOf("STATE");
  const idxZip = headers.indexOf("ZIP");
  const idxVendor = headers.indexOf("VendorID");
  const idxActive = headers.indexOf("ACTIVE");
  const idxStoreEmail = headers.indexOf("STORE_EMAIL");
  const idxStoreEmails = headers.indexOf("STORE_EMAILS");
  const idxSupervisorName = headers.indexOf("SUPERVISOR_NAME");
  const idxSupervisorEmail = headers.indexOf("SUPERVISOR_EMAIL");
  const idxBillingEmails = headers.indexOf("BILLING_EMAILS");
  const idxPmReportEmails = headers.indexOf("PM_REPORT_EMAILS");
  const idxQuoteEmails = headers.indexOf("QUOTE_EMAILS");
  const idxInvoiceEmails = headers.indexOf("INVOICE_EMAILS");

  if (idxNSN === -1) return {};

  const targetNSN = normalizeNSN_(nsn);
  const targetCompany = String(companyId || CFG.DEFAULT_COMPANY_ID).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const rowCompany = idxCompany >= 0
      ? String(data[i][idxCompany] || "").trim().toUpperCase()
      : targetCompany;

    const rowNSN = normalizeNSN_(data[i][idxNSN]);

    const active = idxActive >= 0
      ? String(data[i][idxActive] || "YES").trim().toUpperCase()
      : "YES";

    if (rowCompany === targetCompany && rowNSN === targetNSN && active !== "NO") {
      const normalized = normalizeStoreAddressFields_({
        ADDRESS: idxAddress >= 0 ? data[i][idxAddress] || "" : "",
        CITY: idxCity >= 0 ? data[i][idxCity] || "" : "",
        STATE: idxState >= 0 ? data[i][idxState] || "" : "",
        ZIP: idxZip >= 0 ? data[i][idxZip] || "" : ""
      });

      return {
        nsn: rowNSN,
        client: idxClient >= 0 ? data[i][idxClient] || "" : "",
        address: normalized.ADDRESS,
        city: normalized.CITY,
        state: normalized.STATE,
        zip: normalized.ZIP,
        vendorId: idxVendor >= 0 ? data[i][idxVendor] || "" : "",
        storeEmail: idxStoreEmail >= 0 ? data[i][idxStoreEmail] || "" : "",
        storeEmails: [
          idxStoreEmail >= 0 ? data[i][idxStoreEmail] || "" : "",
          idxStoreEmails >= 0 ? data[i][idxStoreEmails] || "" : ""
        ].filter(Boolean).join(","),
        supervisorName: idxSupervisorName >= 0 ? data[i][idxSupervisorName] || "" : "",
        supervisorEmail: idxSupervisorEmail >= 0 ? data[i][idxSupervisorEmail] || "" : "",
        supervisorEmails: idxSupervisorEmail >= 0 ? data[i][idxSupervisorEmail] || "" : "",
        billingEmails: idxBillingEmails >= 0 ? data[i][idxBillingEmails] || "" : "",
        pmReportEmails: idxPmReportEmails >= 0 ? data[i][idxPmReportEmails] || "" : "",
        quoteEmails: idxQuoteEmails >= 0 ? data[i][idxQuoteEmails] || "" : "",
        invoiceEmails: idxInvoiceEmails >= 0 ? data[i][idxInvoiceEmails] || "" : "",
        fullAddress: normalized.FULL_ADDRESS
      };
    }
  }

  return {};
}

function normalizeStoreAddressFields_(store) {
  let address = String(store.ADDRESS || "").trim();
  let city = String(store.CITY || "").trim();
  let state = String(store.STATE || "").trim().toUpperCase();
  let zip = String(store.ZIP || "").trim();

  if (address) {
    const parsed = parseStoreAddress_(address);

    if (parsed.street && parsed.city && parsed.state && parsed.zip) {
      address = parsed.street;
      city = city || parsed.city;
      state = state || parsed.state;
      zip = zip || parsed.zip;
    }
  }

  const fullAddress = buildFullStoreAddress_(address, city, state, zip);

  return {
    ADDRESS: address,
    CITY: city,
    STATE: state,
    ZIP: zip,
    FULL_ADDRESS: fullAddress
  };
}

function parseStoreAddress_(fullAddress) {
  const text = String(fullAddress || "").trim();

  if (!text) {
    return {
      street: "",
      city: "",
      state: "",
      zip: ""
    };
  }

  const match = text.match(/^(.*?),\s*([^,]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/i);

  if (!match) {
    return {
      street: text,
      city: "",
      state: "",
      zip: ""
    };
  }

  return {
    street: match[1].trim(),
    city: match[2].trim().toUpperCase(),
    state: match[3].trim().toUpperCase(),
    zip: match[4].trim()
  };
}

function buildFullStoreAddress_(address, city, state, zip) {
  address = String(address || "").trim();
  city = String(city || "").trim();
  state = String(state || "").trim().toUpperCase();
  zip = String(zip || "").trim();

  const cityStateZip = [
    city,
    [state, zip].filter(Boolean).join(" ")
  ].filter(Boolean).join(", ");

  return [address, cityStateZip].filter(Boolean).join(", ");
}

function normalizeNSN_(value) {
  return String(value || "")
    .replace(/NSN/gi, "")
    .replace(/#/g, "")
    .replace(/\.0$/g, "")
    .trim();
}

function enrichWorkOrderObject_(obj) {
  const store = getStoreByNSN_(obj.NSN, obj.COMPANY_ID || CFG.DEFAULT_COMPANY_ID);

  obj.STORE_ADDRESS = store.fullAddress || "";
  obj.STORE_STREET = store.address || "";
  obj.STORE_CITY = store.city || "";
  obj.STORE_STATE = store.state || "";
  obj.STORE_ZIP = store.zip || "";
  obj.VENDOR_ID = store.vendorId || "";

  if (store.client) obj.CLIENT = store.client;

  return obj;
}

function getStores(companyId, role, sessionToken) {
  companyId = String(companyId || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN"], companyId);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_STORES);
  if (!sh) throw new Error("No existe la hoja STORES.");

  ensureStoreEmailColumns_(sh);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  const idxCompany = headers.indexOf("COMPANY_ID");

  if (idxCompany === -1) {
    throw new Error("La hoja STORES debe tener la columna COMPANY_ID.");
  }

  return data.slice(1).map(function(row, i) {
    const obj = {};

    headers.forEach(function(h, c) {
      obj[h] = row[c];
    });

    const normalized = normalizeStoreAddressFields_(obj);
    obj.ADDRESS = normalized.ADDRESS;
    obj.CITY = normalized.CITY;
    obj.STATE = normalized.STATE;
    obj.ZIP = normalized.ZIP;
    obj.FULL_ADDRESS = normalized.FULL_ADDRESS;

    obj.ROW_NUMBER = i + 2;
    return obj;
  }).filter(function(o) {
    return String(o.COMPANY_ID || "").trim().toUpperCase() === companyId;
  });
}

function saveStore(store, sessionToken) {
  store = store || {};
  const targetCompany = String(store.COMPANY_ID || "").trim().toUpperCase();
  requireSession_(sessionToken, ["OWNER", "ADMIN"], targetCompany);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_STORES);
  if (!sh) throw new Error("No existe la hoja STORES.");

  ensureStoreEmailColumns_(sh);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) {
    return String(h).trim();
  });

  const normalized = normalizeStoreAddressFields_(store);

  store.ADDRESS = normalized.ADDRESS;
  store.CITY = normalized.CITY;
  store.STATE = normalized.STATE;
  store.ZIP = normalized.ZIP;

  const row = Number(store.ROW_NUMBER || 0);

  const values = headers.map(function(h) {
    return store[h] !== undefined ? store[h] : "";
  });

  if (row > 1) {
    sh.getRange(row, 1, 1, headers.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }

  return true;
}

function ensureStoreEmailColumns_(sh) {
  const required = [];

  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim().toUpperCase();
    });

  required.forEach(function(h) {
    if (headers.indexOf(h) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      headers.push(h);
    }
  });
}
