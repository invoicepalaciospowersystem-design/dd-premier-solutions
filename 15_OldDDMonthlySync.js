function syncOldDDPremierMonthly(sessionToken) {
  const session = requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"]);
  const oldSS = SpreadsheetApp.openById(CFG.OLD_ECONOMY_SPREADSHEET_ID);
  const oldSh = oldSS.getSheetByName("RESUMEN  D&D-PREMIER");
  if (!oldSh) throw new Error("No existe RESUMEN  D&D-PREMIER.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newSh = ss.getSheetByName("DD_OLD_MONTHLY");
  if (!newSh) throw new Error("No existe DD_OLD_MONTHLY.");

  const data = oldSh.getDataRange().getValues();
  if (data.length < 2) return { success: true, rows: 0 };

  const output = [];

  output.push([
    "MES",
    "TOTAL_HORAS",
    "TOTAL_PARTS",
    "TOTAL_COBRO_HORAS",
    "TOTAL_COBRO_PARTS",
    "TOTAL_COMPANIA",
    "TOTAL_FACTURADO",
    "NOTAS"
  ]);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    let mes = row[0];
    if (mes instanceof Date) {
      mes = Utilities.formatDate(mes, CFG.TIMEZONE, "yyyy-MM");
    } else {
      mes = String(mes || "").trim();
    }

    output.push([
      mes,
      Number(row[1] || 0),
      Number(row[2] || 0),
      Number(row[3] || 0),
      Number(row[4] || 0),
      Number(row[5] || 0),
      Number(row[6] || 0),
      row[7] || ""
    ]);
  }

  newSh.clearContents();
  newSh.getRange(1, 1, output.length, output[0].length).setValues(output);

  newSh.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
  if (output.length > 1) {
    newSh.getRange(2, 2, output.length - 1, 1).setNumberFormat("0.00");
    newSh.getRange(2, 3, output.length - 1, 5).setNumberFormat("$#,##0.00");
  }

  addAuditLog_("ECONOMY", "OLD_DD_MONTHLY_SYNCED", session.companyId || CFG.DEFAULT_COMPANY_ID, "DD_OLD_MONTHLY", "SYNC", session, {
    rows: output.length - 1
  });

  return {
    success: true,
    rows: output.length - 1
  };
}
function testOldSheetData(sessionToken) {
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"]);
  const oldSS = SpreadsheetApp.openById("1kkOZMlZRcrmnf5fYT6GYJk0Ne8t_but2-1Enknoy7vw");
  const oldSh = oldSS.getSheetByName("RESUMEN  D&D-PREMIER");

  const data = oldSh.getDataRange().getValues();

  Logger.log(data);
}

function getOldDDPremierMonthlyData(sessionToken) {
  requireSession_(sessionToken, ["OWNER", "ADMIN", "ECONOMIA"]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("DD_OLD_MONTHLY");

  if (!sh) throw new Error("No existe DD_OLD_MONTHLY.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(function(h) {
    return String(h).trim();
  });

  return data.slice(1).map(function(row) {
    const obj = {};

    headers.forEach(function(h, i) {
      let value = row[i];

      if (value instanceof Date) {
        value = Utilities.formatDate(value, CFG.TIMEZONE, "yyyy-MM");
      }

      obj[h] = value;
    });

    return obj;
  });
}
