const APPS_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbwAiZ0Dh5BQoX-QTmMbZcRDEr974-X_nNWcW5x2XeYurC_CeXLNSrl1k-f3p1DDKppOCw/exec";

const DEFAULT_COMPANY_ID = "PPS";

export default {
  async fetch(request) {
    const requestUrl = new URL(request.url);
    const params = new URLSearchParams(requestUrl.search);

    if (!params.get("companyId")) {
      params.set("companyId", DEFAULT_COMPANY_ID);
    }

    const iframeUrl = APPS_SCRIPT_URL + "?" + params.toString();

    return new Response(renderAppShell(iframeUrl), {
      headers: {
        "content-type": "text/html; charset=UTF-8",
        "cache-control": "no-store"
      }
    });
  }
};

function renderAppShell(iframeUrl) {
  return `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Palacios Power Systems Corp</title>
  <style>
    html,
    body {
      width: 100%;
      height: 100%;
      margin: 0;
      overflow: hidden;
      background: #07090d;
    }

    iframe {
      width: 100%;
      height: 100%;
      border: 0;
      display: block;
      background: #07090d;
    }

    .loading {
      position: fixed;
      inset: 0;
      display: grid;
      place-items: center;
      color: #f9fafb;
      font: 600 15px Arial, sans-serif;
      background:
        radial-gradient(circle at 20% 20%, rgba(220, 38, 38, .22), transparent 32%),
        #07090d;
      z-index: 0;
    }
  </style>
</head>
<body>
  <div class="loading">Cargando Palacios Power Systems...</div>
  <iframe
    title="Palacios Power Systems"
    src="${escapeHtml(iframeUrl)}"
    allow="clipboard-read; clipboard-write"
  ></iframe>
</body>
</html>`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
