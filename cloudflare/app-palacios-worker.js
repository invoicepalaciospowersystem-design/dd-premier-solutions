const APPS_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbwAiZ0Dh5BQoX-QTmMbZcRDEr974-X_nNWcW5x2XeYurC_CeXLNSrl1k-f3p1DDKppOCw/exec";

const SERVICE_WORKER_VERSION = "dd-premier-pwa-v1";

const BRANDS = {
  PPS: {
    companyId: "PPS",
    title: "Palacios Power Systems Corp",
    shortName: "Palacios",
    description: "Work orders, invoices, quotes and PM reports for Palacios Power Systems.",
    themeColor: "#111827",
    backgroundColor: "#f8fafc",
    lang: "en",
    iconLetters: "PPS",
    iconAccent: "#dc2626"
  },
  DD: {
    companyId: "",
    title: "D&D Premier Solutions Corp",
    shortName: "D&D Premier",
    description: "Multi-company operations, work orders and economy dashboards.",
    themeColor: "#111827",
    backgroundColor: "#080b10",
    lang: "en",
    iconLetters: "D&D",
    iconAccent: "#ef1717"
  }
};

export default {
  async fetch(request) {
    const requestUrl = new URL(request.url);
    const host = requestUrl.hostname.toLowerCase();
    const brand = getBrandForHost(host);

    if (request.method === "GET") {
      if (requestUrl.pathname === "/manifest.webmanifest") {
        return jsonResponse(buildManifest(brand));
      }

      if (requestUrl.pathname === "/service-worker.js") {
        return new Response(renderServiceWorker(), {
          headers: {
            "content-type": "application/javascript; charset=UTF-8",
            "cache-control": "no-cache"
          }
        });
      }

      if (requestUrl.pathname === "/offline") {
        return new Response(renderOfflinePage(brand), {
          headers: {
            "content-type": "text/html; charset=UTF-8",
            "cache-control": "no-store"
          }
        });
      }

      if (
        requestUrl.pathname === "/pwa-icon.svg" ||
        requestUrl.pathname === "/apple-touch-icon.svg" ||
        requestUrl.pathname === "/favicon.ico"
      ) {
        return new Response(renderIconSvg(brand), {
          headers: {
            "content-type": "image/svg+xml; charset=UTF-8",
            "cache-control": "public, max-age=86400"
          }
        });
      }
    }

    const params = new URLSearchParams(requestUrl.search);

    if (brand.companyId) {
      params.set("companyId", brand.companyId);
      params.delete("ownerOnly");
    } else {
      params.delete("companyId");
      params.set("ownerOnly", "1");
    }

    const query = params.toString();
    const iframeUrl = APPS_SCRIPT_URL + (query ? "?" + query : "");

    return new Response(renderAppShell(iframeUrl, brand), {
      headers: {
        "content-type": "text/html; charset=UTF-8",
        "cache-control": "no-store"
      }
    });
  }
};

function getBrandForHost(host) {
  if (host === "app.palaciospowersystems.com") {
    return BRANDS.PPS;
  }

  if (
    host === "app.ddpremiersolutionscorp.com" ||
    host === "ddpremiersolutionscorp.com" ||
    host === "www.ddpremiersolutionscorp.com"
  ) {
    return BRANDS.DD;
  }

  return BRANDS.DD;
}

function buildManifest(brand) {
  return {
    id: "/",
    name: brand.title,
    short_name: brand.shortName,
    description: brand.description,
    start_url: "/",
    scope: "/",
    display: "standalone",
    display_override: ["standalone", "window-controls-overlay", "browser"],
    orientation: "any",
    lang: brand.lang,
    theme_color: brand.themeColor,
    background_color: brand.backgroundColor,
    categories: ["business", "productivity"],
    icons: [
      {
        src: "/pwa-icon.svg",
        sizes: "any",
        type: "image/svg+xml",
        purpose: "any maskable"
      }
    ],
    shortcuts: buildShortcuts(brand)
  };
}

function buildShortcuts(brand) {
  if (brand.companyId === "PPS") {
    return [
      {
        name: "Technician Portal",
        short_name: "Tech",
        url: "/?companyId=PPS&view=tech",
        icons: [{ src: "/pwa-icon.svg", sizes: "any", type: "image/svg+xml" }]
      },
      {
        name: "Supervisor Portal",
        short_name: "Supervisor",
        url: "/?companyId=PPS&view=customer",
        icons: [{ src: "/pwa-icon.svg", sizes: "any", type: "image/svg+xml" }]
      }
    ];
  }

  return [
    {
      name: "Owner Dashboard",
      short_name: "Owner",
      url: "/?ownerOnly=1",
      icons: [{ src: "/pwa-icon.svg", sizes: "any", type: "image/svg+xml" }]
    }
  ];
}

function renderAppShell(iframeUrl, brand) {
  return `<!doctype html>
<html lang="${escapeHtml(brand.lang)}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <meta name="theme-color" content="${escapeHtml(brand.themeColor)}">
  <meta name="application-name" content="${escapeHtml(brand.title)}">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-title" content="${escapeHtml(brand.shortName)}">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <link rel="manifest" href="/manifest.webmanifest">
  <link rel="icon" href="/pwa-icon.svg" type="image/svg+xml">
  <link rel="apple-touch-icon" href="/apple-touch-icon.svg">
  <title>${escapeHtml(brand.title)}</title>
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

    .install-pwa {
      position: fixed;
      right: max(14px, env(safe-area-inset-right));
      bottom: max(14px, env(safe-area-inset-bottom));
      z-index: 2;
      display: none;
      align-items: center;
      gap: 8px;
      border: 1px solid rgba(255,255,255,.18);
      border-radius: 999px;
      padding: 10px 14px;
      color: #fff;
      background: linear-gradient(135deg, ${escapeHtml(brand.iconAccent)}, #991b1b);
      box-shadow: 0 14px 30px rgba(0,0,0,.28);
      font: 700 13px/1 Arial, sans-serif;
      cursor: pointer;
    }

    .install-pwa:hover {
      filter: brightness(1.06);
    }

    @media (display-mode: standalone) {
      .install-pwa {
        display: none !important;
      }
    }
  </style>
</head>
<body>
  <iframe
    title="${escapeHtml(brand.title)}"
    src="${escapeHtml(iframeUrl)}"
    allow="clipboard-read; clipboard-write"
  ></iframe>
  <button class="install-pwa" id="installPwaButton" type="button" aria-label="Install app">
    Install app
  </button>
  <script>
    (function() {
      var installPromptEvent = null;
      var installButton = document.getElementById("installPwaButton");

      if ("serviceWorker" in navigator && window.isSecureContext) {
        window.addEventListener("load", function() {
          navigator.serviceWorker.register("/service-worker.js").catch(function() {});
        });
      }

      window.addEventListener("beforeinstallprompt", function(event) {
        event.preventDefault();
        installPromptEvent = event;
        if (installButton) installButton.style.display = "inline-flex";
      });

      window.addEventListener("appinstalled", function() {
        installPromptEvent = null;
        if (installButton) installButton.style.display = "none";
      });

      if (installButton) {
        installButton.addEventListener("click", function() {
          if (!installPromptEvent) return;
          installPromptEvent.prompt();
          installPromptEvent.userChoice.finally(function() {
            installPromptEvent = null;
            installButton.style.display = "none";
          });
        });
      }

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
            "app.ddpremiersolutionscorp.com": true,
            "ddpremiersolutionscorp.com": true,
            "www.ddpremiersolutionscorp.com": true
          };

          if (!allowedHosts[target.hostname]) return;
          window.location.href = target.href;
        } catch (err) {}
      });
    })();
  </script>
</body>
</html>`;
}

function renderServiceWorker() {
  return `const CACHE_NAME = "${SERVICE_WORKER_VERSION}";
const OFFLINE_URL = "/offline";

self.addEventListener("install", function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) {
        return cache.addAll([OFFLINE_URL, "/pwa-icon.svg"]);
      })
      .then(function() {
        return self.skipWaiting();
      })
  );
});

self.addEventListener("activate", function(event) {
  event.waitUntil(
    caches.keys()
      .then(function(keys) {
        return Promise.all(keys.map(function(key) {
          if (key !== CACHE_NAME) return caches.delete(key);
          return Promise.resolve();
        }));
      })
      .then(function() {
        return self.clients.claim();
      })
  );
});

self.addEventListener("fetch", function(event) {
  var request = event.request;
  var url = new URL(request.url);

  if (url.origin !== self.location.origin) return;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request).catch(function() {
        return caches.match(OFFLINE_URL);
      })
    );
    return;
  }

  event.respondWith(
    fetch(request).catch(function() {
      return caches.match(request);
    })
  );
});`;
}

function renderOfflinePage(brand) {
  return `<!doctype html>
<html lang="${escapeHtml(brand.lang)}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="theme-color" content="${escapeHtml(brand.themeColor)}">
  <title>${escapeHtml(brand.title)} Offline</title>
  <style>
    body {
      min-height: 100vh;
      margin: 0;
      display: grid;
      place-items: center;
      background: radial-gradient(circle at 20% 20%, rgba(220,38,38,.22), transparent 34%), #080b10;
      color: #f8fafc;
      font-family: Arial, sans-serif;
    }
    main {
      width: min(420px, calc(100% - 32px));
      border: 1px solid rgba(255,255,255,.14);
      border-radius: 18px;
      padding: 28px;
      background: rgba(10,14,20,.86);
      box-shadow: 0 22px 60px rgba(0,0,0,.35);
    }
    h1 { margin: 0 0 10px; font-size: 24px; }
    p { margin: 0; color: #cbd5e1; line-height: 1.5; }
  </style>
</head>
<body>
  <main>
    <h1>${escapeHtml(brand.shortName)} is offline</h1>
    <p>No internet connection was detected. Reconnect and open the app again.</p>
  </main>
</body>
</html>`;
}

function renderIconSvg(brand) {
  const letters = escapeHtml(brand.iconLetters);
  const accent = escapeHtml(brand.iconAccent);
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#111827"/>
      <stop offset=".58" stop-color="#05070a"/>
      <stop offset="1" stop-color="${accent}"/>
    </linearGradient>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="18" stdDeviation="18" flood-color="#000" flood-opacity=".45"/>
    </filter>
  </defs>
  <rect width="512" height="512" rx="112" fill="url(#bg)"/>
  <path d="M82 356 430 156" stroke="${accent}" stroke-width="32" stroke-linecap="round" opacity=".9"/>
  <rect x="94" y="132" width="324" height="248" rx="56" fill="#0b0f16" stroke="#ffffff" stroke-opacity=".46" stroke-width="10" filter="url(#shadow)"/>
  <text x="256" y="286" text-anchor="middle" font-family="Arial, sans-serif" font-size="104" font-weight="900" fill="#fff">${letters}</text>
</svg>`;
}

function jsonResponse(value) {
  return new Response(JSON.stringify(value, null, 2), {
    headers: {
      "content-type": "application/manifest+json; charset=UTF-8",
      "cache-control": "no-cache"
    }
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
