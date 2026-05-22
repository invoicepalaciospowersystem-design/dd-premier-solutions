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
    const responseText = response.getContentText();
    const parsed = parseTwilioResponse_(responseText);

    if (parsed && parsed.sid) {
      props.setProperty("LAST_SMS_SID", parsed.sid);
    }

    Logger.log("SMS CODE: " + response.getResponseCode());
    Logger.log("SMS RESPONSE: " + responseText);

    return {
      phone: to,
      sid: parsed && parsed.sid ? parsed.sid : "",
      status: parsed && parsed.status ? parsed.status : "",
      errorCode: parsed && parsed.error_code ? parsed.error_code : "",
      errorMessage: parsed && parsed.error_message ? parsed.error_message : "",
      code: response.getResponseCode(),
      body: responseText
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

function checkLastSMSStatus() {
  const props = PropertiesService.getScriptProperties();
  const sid = props.getProperty("LAST_SMS_SID");

  if (!sid) {
    Logger.log("No hay LAST_SMS_SID guardado todavia.");
    return null;
  }

  return checkSMSStatus_(sid);
}

function checkSMSStatus_(messageSid) {
  const props = PropertiesService.getScriptProperties();
  const ACCOUNT_SID = props.getProperty("TWILIO_SID");
  const AUTH_TOKEN = props.getProperty("TWILIO_TOKEN");
  const sid = String(messageSid || "").trim();

  if (!ACCOUNT_SID || !AUTH_TOKEN) {
    throw new Error("Faltan TWILIO_SID o TWILIO_TOKEN en Script Properties.");
  }

  if (!sid) {
    throw new Error("Falta el Message SID del SMS.");
  }

  const url = "https://api.twilio.com/2010-04-01/Accounts/" +
    ACCOUNT_SID + "/Messages/" + encodeURIComponent(sid) + ".json";

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Basic " + Utilities.base64Encode(ACCOUNT_SID + ":" + AUTH_TOKEN) },
    muteHttpExceptions: true
  });

  const responseText = response.getContentText();
  const parsed = parseTwilioResponse_(responseText) || {};

  Logger.log("SMS STATUS CODE: " + response.getResponseCode());
  Logger.log("SMS SID: " + (parsed.sid || sid));
  Logger.log("SMS STATUS: " + (parsed.status || ""));
  Logger.log("SMS ERROR CODE: " + (parsed.error_code || ""));
  Logger.log("SMS ERROR MESSAGE: " + (parsed.error_message || ""));
  Logger.log("SMS TO: " + (parsed.to || ""));
  Logger.log("SMS FROM: " + (parsed.from || ""));
  Logger.log("SMS FULL RESPONSE: " + responseText);

  return parsed;
}

function parseTwilioResponse_(responseText) {
  try {
    return JSON.parse(responseText || "{}");
  } catch (err) {
    return null;
  }
}
