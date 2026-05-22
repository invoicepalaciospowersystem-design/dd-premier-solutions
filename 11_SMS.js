// =====================================================
// FILE: 11_SMS.gs
// =====================================================

function sendSMS_(to, message) {
  try {
    if (CFG.SMS_ENABLED === false) return null;
    if (!to || !message) return null;

    to = normalizeSmsPhone_(to);
    if (!to) return null;

    const props = PropertiesService.getScriptProperties();
    const ACCOUNT_SID = props.getProperty("TWILIO_SID");
    const AUTH_TOKEN = props.getProperty("TWILIO_TOKEN");
    const FROM = normalizeSmsPhone_(props.getProperty("TWILIO_FROM"));

    if (!ACCOUNT_SID || !AUTH_TOKEN || !FROM) {
      Logger.log("Twilio properties missing.");
      return null;
    }

    const url = "https://api.twilio.com/2010-04-01/Accounts/" + ACCOUNT_SID + "/Messages.json";

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      payload: { To: to, From: FROM, Body: message },
      headers: { Authorization: "Basic " + Utilities.base64Encode(ACCOUNT_SID + ":" + AUTH_TOKEN) },
      muteHttpExceptions: true
    });

    Logger.log("SMS CODE: " + response.getResponseCode());
    Logger.log("SMS RESPONSE: " + response.getContentText());

    return {
      phone: to,
      code: response.getResponseCode(),
      body: response.getContentText()
    };

  } catch (err) {
    Logger.log("ERROR SMS: " + err);
    notifySystemError_("SMS_SEND_ERROR", err, {
      to: to || "",
      messagePreview: String(message || "").slice(0, 120)
    });
    return null;
  }
}

function sendSMSToTechnicians_(technicianText, message, companyId) {
  const names = getNamesFromText_(technicianText);
  const results = [];

  names.forEach(function(name) {
    const phone = getTechnicianPhone_(name, companyId);
    if (phone) {
      results.push(sendSMS_(phone, message));
    } else {
      Logger.log("No SMS phone found for technician: " + name);
    }
  });

  return results.filter(Boolean);
}

function getTechnicianPhone_(technicianName, companyId) {
  const name = String(technicianName || "").trim();
  if (!name) return "";

  const fromSheet = getTechnicianPhoneFromUsers_(name, companyId);
  if (fromSheet) return fromSheet;

  return TECH_PHONES[name] || "";
}

function getTechnicianPhoneFromUsers_(technicianName, companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_USERS);
  if (!sh) return "";

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return "";

  const headers = data[0].map(function(h) {
    return String(h || "").trim().toUpperCase();
  });

  const idxName = headers.indexOf("NAME");
  const idxCompany = headers.indexOf("COMPANY_ID");
  let idxPhone = headers.indexOf("PHONE");
  if (idxPhone === -1) idxPhone = headers.indexOf("PHONES");

  if (idxName === -1 || idxPhone === -1) return "";

  const targetName = String(technicianName || "").trim().toLowerCase();
  const targetCompany = String(companyId || "").trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][idxName] || "").trim().toLowerCase();
    const rowCompany = idxCompany >= 0
      ? String(data[i][idxCompany] || "").trim().toUpperCase()
      : targetCompany;

    if (rowName !== targetName) continue;
    if (targetCompany && rowCompany && rowCompany !== targetCompany) continue;

    return normalizeSmsPhone_(data[i][idxPhone]);
  }

  return "";
}

function normalizeSmsPhone_(phone) {
  let value = String(phone || "").trim();
  if (!value) return "";

  value = value.replace(/[^\d+]/g, "");
  if (!value) return "";

  if (value.charAt(0) === "+") return value;
  if (value.length === 10) return "+1" + value;
  if (value.length === 11 && value.charAt(0) === "1") return "+" + value;

  return "+" + value;
}

function testSMS() {
  sendSMS_("+17869674478", "TEST SMS funcionando");
}
