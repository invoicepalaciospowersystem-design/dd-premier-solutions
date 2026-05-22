const APPS_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbwAiZ0Dh5BQoX-QTmMbZcRDEr974-X_nNWcW5x2XeYurC_CeXLNSrl1k-f3p1DDKppOCw/exec";

export default {
  async fetch(request) {
    const requestUrl = new URL(request.url);
    const params = new URLSearchParams(requestUrl.search);
    const host = requestUrl.hostname.toLowerCase();
    const isPalacios = host === "app.palaciospowersystems.com";
    const isDdPremier = host === "ddpremiersolutionscorp.com" ||
      host === "www.ddpremiersolutionscorp.com";

    if (isPalacios) {
      params.set("companyId", "PPS");
      params.delete("ownerOnly");
    }

    if (isDdPremier) {
      params.delete("companyId");
      params.set("ownerOnly", "1");
    }

    const query = params.toString();
    const iframeUrl = APPS_SCRIPT_URL + (query ? "?" + query : "");
    const title = isDdPremier
      ? "D&D Premier Solutions Corp"
      : "Palacios Power Systems Corp";

    return new Response(renderAppShell(iframeUrl, title), {
      headers: {
        "content-type": "text/html; charset=UTF-8",
        "cache-control": "no-store"
      }
    });
  }
};

function renderAppShell(iframeUrl, title) {
  return `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(title)}</title>
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
      position: fixed;
      inset: 0;
      width: 100%;
      height: 100%;
      border: 0;
      display: block;
      background: #07090d;
    }
  </style>
</head>
<body>
  <iframe
    title="${escapeHtml(title)}"
    src="${escapeHtml(iframeUrl)}"
    allow="clipboard-read; clipboard-write"
  ></iframe>
  <script>
    window.addEventListener("message", function(event) {
      var data = event.data || {};
      if (!data || data.type !== "PPS_NAVIGATE") return;

      try {
        var target = new URL(String(data.url || ""), window.location.href);
        if (target.protocol !== "https:") return;

        if (target.hostname === "script.google.com") {
          var sameHost = new URL(window.location.href);
          sameHost.search = target.search;
          sameHost.hash = target.hash;
          window.location.href = sameHost.href;
          return;
        }

        var allowedHosts = {
          "app.palaciospowersystems.com": true,
          "ddpremiersolutionscorp.com": true,
          "www.ddpremiersolutionscorp.com": true
        };

        if (!allowedHosts[target.hostname]) return;
        window.location.href = target.href;
      } catch (err) {}
    });
  </script>
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
