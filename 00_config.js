// =====================================================
// D&D PREMIER SOLUTIONS CORP
// WORK ORDERS SYSTEM - BACKEND SEPARADO EN ARCHIVOS .GS
// Mantiene la misma lógica del Codigo.gs original.
// Copia cada sección en un archivo .gs con el nombre indicado.
// =====================================================


// =====================================================
// FILE: 00_Config.gs
// =====================================================

const CFG = {
  APP_NAME: "D&D Premier Solutions Corp",
  PM_REPORTS_FOLDER_ID: "1eCiT7if6s_bqpQPTIi8HLy9_Hl5wfgUi",
PM_REPORT_TEMPLATE_ID: "1wuLdLD_W4IAb4MoA8IHf0birHG6f1rJ0O7Umf5ahbMg",
  APP_ENV: "PRODUCTION",
  TEST_MODE: false,
  SESSION_TTL_SECONDS: 21600,
  LOGIN_LOCK_MAX_ATTEMPTS: 5,
  LOGIN_LOCK_MINUTES: 15,
  TECH_EMAIL_TEST_MODE: false,
  SMS_ENABLED: true,
  AUTO_SEND_CLIENT_EMAILS: true,
  CLIENT_EMAIL_TEST_MODE: false,
  CLIENT_EMAIL_TEST_TO: "ddpremiersolutions.corp@gmail.com",
  SYSTEM_ALERTS_ENABLED: true,
  SYSTEM_ALERT_EMAIL: "ddpremiersolutions.corp@gmail.com",
  SYSTEM_ALERT_THROTTLE_MINUTES: 10,
  INVOICES_FOLDER_ID: "1nHOt6jpoGPmTcRAkndkaDtNjsIal7WDy",
  QUOTES_FOLDER_ID: "1OHZ9Mxffsi1_P2wenpIwgZ95o0wubWop",
  QUOTE_TEMPLATE_ES_DOC_ID: "1buu6D95PW-XjEtBtkI2T4RW8mWd7IfFi6eqTCZyVlqc",
  TEMPLATE_DOC_ID: "1oBT9Ga08guEcPn3Lgfxrkw33EFJiTmWsYO63D10cR2s",
  OLD_ECONOMY_SPREADSHEET_ID: "1kkOZMlZRcrmnf5fYT6GYJk0Ne8t_but2-1Enknoy7vw",

  // IMPORTANTE: cambia esto cuando publiques la Web App V2.
  WEB_APP_URL: "https://script.google.com/macros/s/AKfycbwAiZ0Dh5BQoX-QTmMbZcRDEr974-X_nNWcW5x2XeYurC_CeXLNSrl1k-f3p1DDKppOCw/exec",
  PUBLIC_WEB_APP_URLS: {
    PPS: "https://app.palaciospowersystems.com"
  },

  SHEET_COMPANIES: "COMPANIES",
  SHEET_USERS: "USERS",
  SHEET_WORK_ORDERS: "WORK_ORDERS",
  SHEET_WO_LOG: "WO_LOG",
  SHEET_STORES: "STORES",
  SHEET_NOTIFICATIONS: "NOTIFICATIONS",
  SHEET_ECONOMY: "ECONOMY",
  SHEET_ERROR_LOGS: "ERROR_LOGS",
  SHEET_AUDIT_LOGS: "AUDIT_LOGS",

  TIMEZONE: "America/New_York",
  DEFAULT_COMPANY_ID: "PPS",
  CLIENT_DEFAULT: "McDonald's",

  COMPANY_BRANDING: {
    DEFAULT: {
      companyName: "D&D Premier Solutions Corp",
      loginTitle: "D&D Premier Solutions Corp",
      loginSubtitle: "Sistema de ordenes y operaciones",
      ownerName: "D&D Premier Solutions Corp",
      primaryColor: "#111827",
      accentColor: "#dc2626",
      backgroundImageUrl: "",
      backgroundFileId: "",
      logoImageUrl: "",
      logoFileId: "1nE27lzdKy_EPYbAf2CX0eiQIRU5tsSMN"
    },
    PPS: {
      companyName: "Palacios Power Systems Corp",
      loginTitle: "Palacios Power Systems Corp",
      loginSubtitle: "Electrical - HVAC & Restaurant Equipment",
      ownerName: "D&D Premier Solutions Corp",
      primaryColor: "#b91c1c",
      accentColor: "#111827",
      backgroundImageUrl: "",
      backgroundFileId: "",
      logoImageUrl: "",
      logoFileId: "1fgvNvOfimOPIgrN_dfM5uzHH-IpGIP6-"
    }
  },

  TECHS: {
    "David Dominguez": "ddpremiersolutions.corp@gmail.com",
    "Miguel Rodriguez": "minongo2005@gmail.com",
    "Osnier Huerta": "osnierj@gmail.com",
    "Alain Mora": "alainmora0392@gmail.com"
  }
};

const SUPERVISORS = {
  "Ivana Jaime": ["1155", "4801", "10571", "11242", "11472", "4143"]
};

const TECH_PHONES = {
  "David Dominguez": "+17869674478",
  "Miguel Rodriguez": "+17866414816",
  "Osnier": "+18135856487",
  "Sarahi": "+19542680157",
  "Dayre": "+17869674479"
};

const PM_TEXTS = {
  MENSUAL: {
    ES: `Mantenimiento preventivo mensual a sistemas de aire acondicionado, incluyendo:

• Limpieza de tuberías de drenaje
• Aplicación de tabletas anti-algas en bandejas de condensado
• Limpieza de condensadores
• Medición de temperaturas del aire en suministro en las áreas del restaurante.`,

    EN: `Monthly preventive maintenance on air conditioning systems, including:

• Drain line cleaning
• Anti-algae tablets applied in condensate drain pans
• Condenser cleaning
• Supply air temperature measurements in restaurant areas.`
  },

  TRIMESTRAL: {
    ES: `Mantenimiento preventivo trimestral a sistemas de aire acondicionado, incluyendo:

• Limpieza de tuberías de drenaje
• Aplicación de tabletas anti-algas en bandejas de condensado
• Limpieza de condensadores
• Cambio de filtros de aire
• Medición de temperaturas del aire en suministro en las áreas del restaurante.`,

    EN: `Quarterly preventive maintenance on air conditioning systems, including:

• Drain line cleaning
• Anti-algae tablets applied in condensate drain pans
• Condenser cleaning
• Air filter replacement
• Supply air temperature measurements in restaurant areas.`
  },

  ANNUAL: {
    ES: `Mantenimiento preventivo anual a sistemas de aire acondicionado, incluyendo:

• Limpieza de tuberías de drenaje
• Aplicación de tabletas anti-algas en bandejas de condensado
• Limpieza de condensadores
• Cambio de filtros de aire
• Limpieza de evaporadores
• Limpieza de blowers
• Verificación de presiones de trabajo del refrigerante
• Medición de temperaturas del aire en suministro en las áreas del restaurante.`,

    EN: `Annual preventive maintenance on air conditioning systems, including:

• Drain line cleaning
• Anti-algae tablets applied in condensate drain pans
• Condenser cleaning
• Air filter replacement
• Evaporator cleaning
• Blower cleaning
• Refrigerant operating pressure verification
• Supply air temperature measurements in restaurant areas.`
  }
};
