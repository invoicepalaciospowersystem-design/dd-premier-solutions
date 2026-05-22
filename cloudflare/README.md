# Cloudflare wrapper

Use `app-palacios-worker.js` as the Worker for:

`app.palaciospowersystems.com/*`

This Worker keeps the browser on `app.palaciospowersystems.com` and loads the Google Apps Script Web App inside an iframe with `companyId=PPS`.

Cloudflare steps:

1. Go to Workers & Pages.
2. Create a Worker.
3. Paste the contents of `app-palacios-worker.js`.
4. Deploy it.
5. In the Worker settings, add route `app.palaciospowersystems.com/*` for zone `palaciospowersystems.com`.
6. Disable the old Redirect Rule for `app.palaciospowersystems.com`.

Keep the DNS record for `app` proxied through Cloudflare.
