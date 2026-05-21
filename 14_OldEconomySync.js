function syncInvoicesToOldLog() {
  const oldSS = SpreadsheetApp.openById(CFG.OLD_ECONOMY_SPREADSHEET_ID);
  const oldLog = oldSS.getSheetByName("LOG");
  if (!oldLog) throw new Error("No existe la hoja LOG en el sistema viejo.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shInv = ss.getSheetByName("INVOICES");
  const shEco = ss.getSheetByName(CFG.SHEET_ECONOMY);

  if (!shInv) throw new Error("No existe la hoja INVOICES.");
  if (!shEco) throw new Error("No existe la hoja ECONOMY.");

  const invData = shInv.getDataRange().getValues();
  const ecoData = shEco.getDataRange().getValues();

  if (invData.length < 2) return { success: true, added: 0 };

  const invHeaders = invData[0].map(h => String(h).trim());
  const ecoHeaders = ecoData[0].map(h => String(h).trim());

  const oldData = oldLog.getDataRange().getValues();
  const oldHeaders = oldData[0].map(h => String(h).trim().toUpperCase());

  const oldInvoiceCol = oldHeaders.indexOf("INVOICE");
  if (oldInvoiceCol === -1) throw new Error("LOG viejo debe tener columna Invoice.");

  const existing = {};
  for (let i = 1; i < oldData.length; i++) {
    const inv = String(oldData[i][oldInvoiceCol] || "").trim();
    if (inv) existing[inv] = true;
  }

  function getByHeader(headers, row, name) {
    const idx = headers.indexOf(name);
    return idx >= 0 ? row[idx] : "";
  }

  const ecoMap = {};
  const ecoWO = ecoHeaders.indexOf("WO_NUMBER");

  if (ecoWO === -1) {
    throw new Error("ECONOMY debe tener WO_NUMBER.");
  }

  for (let i = 1; i < ecoData.length; i++) {
    const wo = String(ecoData[i][ecoWO] || "").trim();
    if (wo) ecoMap[wo] = ecoData[i];
  }

  const rowsToAdd = [];

  for (let i = 1; i < invData.length; i++) {
    const invRow = invData[i];

    const invoice = String(getByHeader(invHeaders, invRow, "Invoice") || "").trim();
    if (!invoice || existing[invoice]) continue;

    const woNumber = String(getByHeader(invHeaders, invRow, "WO_NUMBER") || "").trim();
    const ecoRow = ecoMap[woNumber] || [];

    const invSource =
      getByHeader(ecoHeaders, ecoRow, "INV_SOURCE") ||
      getByHeader(invHeaders, invRow, "INVERSION") ||
      "";

    const cost =
      getByHeader(ecoHeaders, ecoRow, "COST") ||
      getByHeader(invHeaders, invRow, "INVERSION") ||
      "";

    const technicians =
      getByHeader(ecoHeaders, ecoRow, "TECHNICIANS") ||
      getByHeader(invHeaders, invRow, "TECHNICIAN NAME") ||
      "";

    rowsToAdd.push([
      getByHeader(invHeaders, invRow, "Timestamp"),
      invoice,
      getByHeader(invHeaders, invRow, "NS"),
      getByHeader(invHeaders, invRow, "Source"),
      getByHeader(invHeaders, invRow, "TECHNICIAN EMAIL"),
      technicians,
      "RECIBIDO",
      getByHeader(invHeaders, invRow, "HORAS"),
      getByHeader(invHeaders, invRow, "PARTS"),
      getByHeader(invHeaders, invRow, "LABOR"),
      getByHeader(invHeaders, invRow, "TAX"),
      getByHeader(invHeaders, invRow, "TOTAL"),
      getByHeader(invHeaders, invRow, "CLIENTE"),
      cost
    ]);
  }

  if (rowsToAdd.length) {
    oldLog
      .getRange(oldLog.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length)
      .setValues(rowsToAdd);
  }

  return {
    success: true,
    added: rowsToAdd.length
  };
}
