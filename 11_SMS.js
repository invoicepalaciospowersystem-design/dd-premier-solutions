// =====================================================
// FILE: 11_SMS.gs
// =====================================================

function sendSMS_(to, message) {
  try {
    if (!to || !message) return;

    const props = PropertiesService.getScriptProperties();
    const ACCOUNT_SID = props.getProperty("TWILIO_SID");
    const AUTH_TOKEN = props.getProperty("TWILIO_TOKEN");
    const FROM = props.getProperty("TWILIO_FROM");

    if (!ACCOUNT_SID || !AUTH_TOKEN || !FROM) {
      Logger.log("Twilio properties missing.");
      return;
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

  } catch (err) {
    Logger.log("ERROR SMS: " + err);
  }
}

function sendSMSToTechnicians_(technicianText, message) {
  const names = getNamesFromText_(technicianText);
  names.forEach(function(name) {
    if (TECH_PHONES[name]) sendSMS_(TECH_PHONES[name], message);
  });
}

function testSMS() {
  sendSMS_("+17869674478", "🔥 TEST SMS funcionando");
}