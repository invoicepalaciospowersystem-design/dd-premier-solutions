# Cloudflare wrapper + PWA

Use `app-palacios-worker.js` as the Worker for:

`app.palaciospowersystems.com/*`
`ddpremiersolutionscorp.com/*`
`www.ddpremiersolutionscorp.com/*`
`app.ddpremiersolutionscorp.com/*`

This Worker keeps the browser on the public domain and proxies the Google Apps Script Web App through Cloudflare. It removes the Google frame headers that can produce `script.google.com refused to connect`, and it proxies the `/static` and `/macros` paths that Apps Script needs internally. It also serves the PWA files from the same domain:

- `/manifest.webmanifest`
- `/service-worker.js`
- `/pwa-icon.svg`
- `/offline`

Host routing:

- `app.palaciospowersystems.com` loads the app with `companyId=PPS`.
- `ddpremiersolutionscorp.com`, `www.ddpremiersolutionscorp.com`, and `app.ddpremiersolutionscorp.com` load the owner portal with `ownerOnly=1`.

Cloudflare steps:

1. Go to Workers & Pages.
2. Create a Worker.
3. Paste the contents of `app-palacios-worker.js`.
4. Deploy it.
5. In the Worker settings, add the routes for each domain/subdomain listed above.
6. Disable old Redirect Rules for those domains if they point directly to Apps Script.

Keep the DNS record for `app` proxied through Cloudflare.

PWA test:

1. Open `https://app.palaciospowersystems.com/manifest.webmanifest`.
2. Open `https://app.palaciospowersystems.com/service-worker.js`.
3. Open the app in Chrome or Edge and use the browser menu option `Install app`.
4. On iPhone/iPad, use Safari share button > `Add to Home Screen`.
