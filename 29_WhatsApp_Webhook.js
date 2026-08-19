// =====================================================
// FILE: 29_WhatsApp_Webhook.gs
// =====================================================

const WHATSAPP_WEBHOOK_ROUTE_ = "whatsappWebhook";
const WHATSAPP_VERIFY_TOKEN_PROPERTY_ = "WHATSAPP_VERIFY_TOKEN";
const WHATSAPP_DEFAULT_VERIFY_TOKEN_ = "DDP_WHATSAPP_20260724_A9M7K3";
const WHATSAPP_ACCESS_TOKEN_PROPERTY_ = "WHATSAPP_ACCESS_TOKEN";
const WHATSAPP_PHONE_NUMBER_ID_PROPERTY_ = "WHATSAPP_PHONE_NUMBER_ID";
const WHATSAPP_GRAPH_VERSION_PROPERTY_ = "WHATSAPP_GRAPH_API_VERSION";
const WHATSAPP_WEBHOOK_LOG_SHEET_ = "WHATSAPP_WEBHOOK_LOGS";
const WHATSAPP_WEBHOOK_LOG_HEADERS_ = [
  "TIMESTAMP",
  "EVENT_TYPE",
  "WABA_ID",
  "PHONE_NUMBER_ID",
  "DISPLAY_PHONE_NUMBER",
  "MESSAGE_ID",
  "FROM_PHONE",
  "TO_PHONE",
  "STATUS",
  "MESSAGE_TYPE",
  "TEXT",
  "ERROR",
  "RAW_JSON"
];

function isWhatsAppWebhookRequest_(params) {
  params = params || {};
  return String(params.route || "").trim() === WHATSAPP_WEBHOOK_ROUTE_ ||
    String(params.webhook || "").trim().toLowerCase() === "whatsapp" ||
    String(params.whatsappWebhook || "").trim() === "1";
}

function handleWhatsAppWebhookVerification_(e) {
  const p = e && e.parameter ? e.parameter : {};
  const mode = String(p["hub.mode"] || "").trim();
  const token = String(p["hub.verify_token"] || "").trim();
  const challenge = String(p["hub.challenge"] || "").trim();
  const expectedToken = getWhatsAppVerifyToken_();

  if (mode === "subscribe" && token === expectedToken && challenge) {
    return whatsappTextOutput_(challenge);
  }

  return whatsappTextOutput_("Forbidden");
}

function handleWhatsAppWebhookEvent_(e) {
  const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : "{}";

  try {
    const payload = JSON.parse(rawBody);
    appendWhatsAppWebhookLog_(payload);
    return whatsappTextOutput_("EVENT_RECEIVED");
  } catch (err) {
    notifySystemError_("WHATSAPP_WEBHOOK_EVENT_ERROR", err, {
      module: "WHATSAPP",
      rawBody: rawBody
    });

    return whatsappJsonOutput_({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function setupWhatsAppWebhook() {
  const props = PropertiesService.getScriptProperties();
  let verifyToken = props.getProperty(WHATSAPP_VERIFY_TOKEN_PROPERTY_);

  if (!verifyToken) {
    verifyToken = WHATSAPP_DEFAULT_VERIFY_TOKEN_;
    props.setProperty(WHATSAPP_VERIFY_TOKEN_PROPERTY_, verifyToken);
  }

  if (!props.getProperty(WHATSAPP_GRAPH_VERSION_PROPERTY_)) {
    props.setProperty(WHATSAPP_GRAPH_VERSION_PROPERTY_, "v25.0");
  }

  ensureWhatsAppWebhookLogSheet_();

  const info = {
    callbackUrl: getWhatsAppWebhookCallbackUrl_(),
    verifyToken: verifyToken,
    accessTokenSet: !!props.getProperty(WHATSAPP_ACCESS_TOKEN_PROPERTY_),
    phoneNumberId: props.getProperty(WHATSAPP_PHONE_NUMBER_ID_PROPERTY_) || "",
    graphVersion: props.getProperty(WHATSAPP_GRAPH_VERSION_PROPERTY_) || "v25.0",
    logSheet: WHATSAPP_WEBHOOK_LOG_SHEET_
  };

  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

function getWhatsAppWebhookSetupInfo() {
  return setupWhatsAppWebhook();
}

function getWhatsAppWebhookCallbackUrl_() {
  return String(CFG.WEB_APP_URL || "").replace(/\/+$/g, "") + "?route=" + WHATSAPP_WEBHOOK_ROUTE_;
}

function getWhatsAppVerifyToken_() {
  return PropertiesService.getScriptProperties()
    .getProperty(WHATSAPP_VERIFY_TOKEN_PROPERTY_) || WHATSAPP_DEFAULT_VERIFY_TOKEN_;
}

function appendWhatsAppWebhookLog_(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sh = ensureWhatsAppWebhookLogSheet_();
    const rows = extractWhatsAppWebhookRows_(payload);

    if (!rows.length) return;

    sh.getRange(sh.getLastRow() + 1, 1, rows.length, WHATSAPP_WEBHOOK_LOG_HEADERS_.length)
      .setValues(rows);
  } finally {
    lock.releaseLock();
  }
}

function ensureWhatsAppWebhookLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(WHATSAPP_WEBHOOK_LOG_SHEET_);

  if (!sh) {
    sh = ss.insertSheet(WHATSAPP_WEBHOOK_LOG_SHEET_);
  }

  ensureSheetColumns_(sh, WHATSAPP_WEBHOOK_LOG_HEADERS_);
  return sh;
}

function extractWhatsAppWebhookRows_(payload) {
  const rows = [];
  const now = new Date();
  const raw = safeStringifySystemError_(payload);
  const entries = Array.isArray(payload && payload.entry) ? payload.entry : [];

  entries.forEach(function(entry) {
    const wabaId = entry && entry.id ? String(entry.id) : "";
    const changes = Array.isArray(entry && entry.changes) ? entry.changes : [];

    changes.forEach(function(change) {
      const value = change && change.value ? change.value : {};
      const metadata = value.metadata || {};
      const phoneNumberId = String(metadata.phone_number_id || "");
      const displayPhoneNumber = String(metadata.display_phone_number || "");
      const messages = Array.isArray(value.messages) ? value.messages : [];
      const statuses = Array.isArray(value.statuses) ? value.statuses : [];

      messages.forEach(function(message) {
        rows.push(makeWhatsAppLogRow_({
          timestamp: now,
          eventType: "MESSAGE",
          wabaId: wabaId,
          phoneNumberId: phoneNumberId,
          displayPhoneNumber: displayPhoneNumber,
          messageId: message.id || "",
          fromPhone: message.from || "",
          toPhone: displayPhoneNumber,
          messageType: message.type || "",
          text: getWhatsAppMessageText_(message),
          rawJson: raw
        }));
      });

      statuses.forEach(function(status) {
        rows.push(makeWhatsAppLogRow_({
          timestamp: now,
          eventType: "STATUS",
          wabaId: wabaId,
          phoneNumberId: phoneNumberId,
          displayPhoneNumber: displayPhoneNumber,
          messageId: status.id || "",
          toPhone: status.recipient_id || "",
          status: status.status || "",
          error: safeStringifySystemError_(status.errors || []),
          rawJson: raw
        }));
      });
    });
  });

  if (!rows.length) {
    rows.push(makeWhatsAppLogRow_({
      timestamp: now,
      eventType: "UNKNOWN",
      rawJson: raw
    }));
  }

  return rows;
}

function makeWhatsAppLogRow_(data) {
  data = data || {};
  return [
    data.timestamp || new Date(),
    data.eventType || "",
    data.wabaId || "",
    data.phoneNumberId || "",
    data.displayPhoneNumber || "",
    data.messageId || "",
    data.fromPhone || "",
    data.toPhone || "",
    data.status || "",
    data.messageType || "",
    data.text || "",
    data.error || "",
    data.rawJson || ""
  ];
}

function getWhatsAppMessageText_(message) {
  message = message || {};

  if (message.text && message.text.body) {
    return String(message.text.body || "");
  }

  if (message.button && message.button.text) {
    return String(message.button.text || "");
  }

  if (message.interactive) {
    return safeStringifySystemError_(message.interactive);
  }

  return safeStringifySystemError_(message).slice(0, 5000);
}

function sendWhatsAppTextMessage_(toPhone, body) {
  const props = PropertiesService.getScriptProperties();
  const accessToken = getRequiredScriptProperty_(WHATSAPP_ACCESS_TOKEN_PROPERTY_);
  const phoneNumberId = getRequiredScriptProperty_(WHATSAPP_PHONE_NUMBER_ID_PROPERTY_);
  const graphVersion = props.getProperty(WHATSAPP_GRAPH_VERSION_PROPERTY_) || "v25.0";

  const url = "https://graph.facebook.com/" + graphVersion + "/" +
    encodeURIComponent(phoneNumberId) + "/messages";

  const payload = {
    messaging_product: "whatsapp",
    to: normalizeWhatsAppPhone_(toPhone),
    type: "text",
    text: {
      preview_url: false,
      body: String(body || "")
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    muteHttpExceptions: true,
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + accessToken
    },
    payload: JSON.stringify(payload)
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("WhatsApp send failed " + code + ": " + text);
  }

  return JSON.parse(text);
}

function testWhatsAppTextMessage(toPhone) {
  return sendWhatsAppTextMessage_(toPhone, "D&D Premier Work Orders test message.");
}

function normalizeWhatsAppPhone_(value) {
  const digits = String(value || "").replace(/[^\d]/g, "");
  if (digits.length === 10) return "1" + digits;
  return digits;
}

function whatsappTextOutput_(text) {
  return ContentService.createTextOutput(String(text || ""))
    .setMimeType(ContentService.MimeType.TEXT);
}

function whatsappJsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value || {}, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}
