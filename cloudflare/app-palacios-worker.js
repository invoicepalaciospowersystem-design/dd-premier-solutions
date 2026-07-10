const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwAiZ0Dh5BQoX-QTmMbZcRDEr974-X_nNWcW5x2XeYurC_CeXLNSrl1k-f3p1DDKppOCw/exec";
const GOOGLE_SCRIPT_ORIGIN = "https://script.google.com";
const GOOGLE_USERCONTENT_ORIGIN = "https://n-pdqbm3jiktl42ouuuxskcgzyi3matvmx4a26gha-0lu-script.googleusercontent.com";

const BRANDS = {
  PPS: {
    companyId: "PPS",
    name: "Palacios Power Systems Corp",
    shortName: "Palacios",
    theme: "#111827",
    bg: "#f8fafc",
    accent: "#dc2626",
    letters: "PPS"
  },
  DD: {
    companyId: "",
    name: "D&D Premier Solutions Corp",
    shortName: "D&D Premier",
    theme: "#111827",
    bg: "#080b10",
    accent: "#ef1717",
    letters: "D&D"
  }
};

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const brand = getBrand(url.hostname);

    if (request.method === "GET") {
      if (url.pathname === "/manifest.webmanifest") return manifestResponse(brand);
      if (url.pathname === "/service-worker.js") return jsResponse(serviceWorkerJs());
      if (url.pathname === "/offline") return htmlResponse(offlineHtml(brand), "no-store");
      if (url.pathname === "/pwa-icon.svg" || url.pathname === "/favicon.ico" || url.pathname === "/apple-touch-icon.svg") {
        return svgResponse(iconSvg(brand));
      }
    }

    return redirectToAppsScript(brand, url);
  }
};

function getBrand(hostname) {
  const host = String(hostname || "").toLowerCase();
  if (host === "app.palaciospowersystems.com") return BRANDS.PPS;
  return BRANDS.DD;
}

function manifestResponse(brand) {
  const shortcuts = brand.companyId
    ? [
        shortcut("Technician Portal", "/?companyId=PPS&view=tech"),
        shortcut("Supervisor Portal", "/?companyId=PPS&view=customer")
      ]
    : [shortcut("Owner Dashboard", "/?ownerOnly=1")];

  const body = {
    name: brand.name,
    short_name: brand.shortName,
    id: "/",
    start_url: "/",
    scope: "/",
    display: "standalone",
    orientation: "any",
    theme_color: brand.theme,
    background_color: brand.bg,
    categories: ["business", "productivity"],
    icons: [{ src: "/pwa-icon.svg", sizes: "any", type: "image/svg+xml", purpose: "any maskable" }],
    shortcuts
  };

  return new Response(JSON.stringify(body), {
    headers: {
      "content-type": "application/manifest+json; charset=UTF-8",
      "cache-control": "no-cache"
    }
  });
}

function shortcut(name, url) {
  return {
    name,
    short_name: name,
    url,
    icons: [{ src: "/pwa-icon.svg", sizes: "any", type: "image/svg+xml" }]
  };
}

async function proxyAppsScriptRequest(request, brand, requestUrl) {
  const upstreamUrl = buildAppsScriptUrl(brand, requestUrl.search);
  const upstreamResponse = await fetch(new Request(upstreamUrl, request));
  const contentType = upstreamResponse.headers.get("content-type") || "";

  if (contentType.toLowerCase().includes("text/html")) {
    const html = injectAppShellSupport(await upstreamResponse.text(), brand);
    return new Response(html, {
      status: upstreamResponse.status,
      statusText: upstreamResponse.statusText,
      headers: proxyHeaders(upstreamResponse.headers, "text/html; charset=UTF-8", "no-store")
    });
  }

  return new Response(upstreamResponse.body, {
    status: upstreamResponse.status,
    statusText: upstreamResponse.statusText,
    headers: proxyHeaders(upstreamResponse.headers, contentType, "no-store")
  });
}

function redirectToAppsScript(brand, requestUrl) {
  return Response.redirect(buildAppsScriptUrl(brand, requestUrl.search), 302);
}

async function proxyGoogleRequest(request, upstreamUrl) {
  const upstreamResponse = await fetch(new Request(upstreamUrl, request));
  const contentType = upstreamResponse.headers.get("content-type") || "application/octet-stream";

  return new Response(upstreamResponse.body, {
    status: upstreamResponse.status,
    statusText: upstreamResponse.statusText,
    headers: proxyHeaders(upstreamResponse.headers, contentType, "no-store")
  });
}

function buildAppsScriptUrl(brand, search) {
  const params = new URLSearchParams(search || "");

  if (brand.companyId) {
    params.set("companyId", brand.companyId);
    params.delete("ownerOnly");
  } else {
    params.delete("companyId");
    params.set("ownerOnly", "1");
  }

  const query = params.toString();
  return APPS_SCRIPT_URL + (query ? "?" + query : "");
}

function proxyHeaders(sourceHeaders, contentType, cacheControl) {
  const headers = new Headers(sourceHeaders);
  headers.set("content-type", contentType || "application/octet-stream");
  headers.set("cache-control", cacheControl || "no-store");
  headers.delete("content-security-policy");
  headers.delete("content-security-policy-report-only");
  headers.delete("x-frame-options");
  headers.delete("content-encoding");
  headers.delete("content-length");
  headers.delete("set-cookie");
  return headers;
}

function injectAppShellSupport(html, brand) {
  const headInject = `
  <meta name="theme-color" content="${esc(brand.theme)}">
  <meta name="application-name" content="${esc(brand.name)}">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-title" content="${esc(brand.shortName)}">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <link rel="manifest" href="/manifest.webmanifest">
  <link rel="icon" href="/pwa-icon.svg" type="image/svg+xml">
  <link rel="apple-touch-icon" href="/apple-touch-icon.svg">`;

  const bodyInject = `
  <script>
    (function(){
      try { sessionStorage.setItem("pps_wrapper_active_session", "1"); } catch (err) {}
      if ("serviceWorker" in navigator && window.isSecureContext) {
        window.addEventListener("load", function(){
          navigator.serviceWorker.register("/service-worker.js").catch(function(){});
        });
      }
      window.addEventListener("message", function(event) {
        var data = event.data || {};
        if (!data || data.type !== "PPS_NAVIGATE") return;
        try {
          var target = new URL(String(data.url || ""), window.location.href);
          if (target.protocol !== "https:") return;
          var ok = {
            "app.palaciospowersystems.com": true,
            "app.ddpremiersolutionscorp.com": true,
            "ddpremiersolutionscorp.com": true,
            "www.ddpremiersolutionscorp.com": true,
            "script.google.com": true
          };
          if (!ok[target.hostname]) return;
          var next = new URL(window.location.href);
          next.search = target.search;
          next.hash = target.hash;
          window.location.href = next.href;
        } catch (err) {}
      });
    })();
  </script>`;

  return String(html || "")
    .replace(/(<head[^>]*>)/i, "$1" + headInject)
    .replace(/<\/body>/i, bodyInject + "</body>");
}

function serviceWorkerJs() {
  return `const CACHE_NAME="dd-premier-pwa-v5";
const OFFLINE_URL="/offline";
self.addEventListener("install",event=>event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll([OFFLINE_URL,"/pwa-icon.svg"])).then(()=>self.skipWaiting())));
self.addEventListener("activate",event=>event.waitUntil(caches.keys().then(keys=>Promise.all(keys.map(key=>key===CACHE_NAME?null:caches.delete(key)))).then(()=>self.clients.claim())));
self.addEventListener("fetch",event=>{
  const req=event.request;
  const url=new URL(req.url);
  if(url.origin!==self.location.origin)return;
  if(req.mode==="navigate"){
    event.respondWith(fetch(req).catch(()=>caches.match(OFFLINE_URL)));
    return;
  }
  event.respondWith(fetch(req).catch(()=>caches.match(req)));
});`;
}

function offlineHtml(brand) {
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(brand.shortName)} Offline</title><style>body{min-height:100vh;margin:0;display:grid;place-items:center;background:#080b10;color:#f8fafc;font-family:Arial,sans-serif}main{max-width:420px;border:1px solid rgba(255,255,255,.16);border-radius:18px;padding:28px;background:#0b0f16}p{color:#cbd5e1}</style></head><body><main><h1>${esc(brand.shortName)} is offline</h1><p>No internet connection was detected. Reconnect and open the app again.</p></main></body></html>`;
}

function iconSvg(brand) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#111827"/><stop offset=".62" stop-color="#05070a"/><stop offset="1" stop-color="${esc(brand.accent)}"/></linearGradient></defs><rect width="512" height="512" rx="112" fill="url(#g)"/><path d="M82 356 430 156" stroke="${esc(brand.accent)}" stroke-width="32" stroke-linecap="round"/><rect x="92" y="132" width="328" height="248" rx="58" fill="#0b0f16" stroke="#fff" stroke-opacity=".50" stroke-width="10"/><text x="256" y="286" text-anchor="middle" font-family="Arial,sans-serif" font-size="104" font-weight="900" fill="#fff">${esc(brand.letters)}</text></svg>`;
}

function htmlResponse(body, cacheControl) {
  return new Response(body, {
    headers: {
      "content-type": "text/html; charset=UTF-8",
      "cache-control": cacheControl || "no-cache"
    }
  });
}

function jsResponse(body) {
  return new Response(body, {
    headers: {
      "content-type": "application/javascript; charset=UTF-8",
      "cache-control": "no-cache"
    }
  });
}

function svgResponse(body) {
  return new Response(body, {
    headers: {
      "content-type": "image/svg+xml; charset=UTF-8",
      "cache-control": "public, max-age=86400"
    }
  });
}

function esc(value) {
  return String(value || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
