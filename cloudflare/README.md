# Cloudflare Worker + PWA

Este README explica como activar el Worker correcto en Cloudflare para que la app abra desde los dominios propios y no desde el link largo de Google Apps Script.

Archivo que debes copiar en Cloudflare:

`cloudflare/app-palacios-worker.js`

Ruta completa en esta PC:

`C:\Users\ddval\OneDrive\Documents\D&D PREMIER SOLUTIONS CORP\cloudflare\app-palacios-worker.js`

## Que hace este Worker

- Mantiene la app en `app.palaciospowersystems.com`.
- Mantiene el portal owner en los dominios de D&D Premier.
- Sirve la PWA desde el mismo dominio.
- Evita el error `script.google.com refused to connect`.
- Ya no usa un iframe directo hacia Google Apps Script.
- Hace proxy de Google Apps Script por Cloudflare.
- Hace proxy de rutas internas que Apps Script necesita, como `/static/` y `/macros/`.

## Dominios que debe manejar

El mismo Worker debe estar conectado a estos dominios/rutas:

- `app.palaciospowersystems.com/*`
- `ddpremiersolutionscorp.com/*`
- `www.ddpremiersolutionscorp.com/*`
- `app.ddpremiersolutionscorp.com/*`

## Como pegar el Worker en Cloudflare

1. Entra a Cloudflare.
2. Ve a `Workers & Pages`.
3. Abre el Worker `palacios-app-wrapper`.
4. Dale a `Edit code`.
5. Borra TODO el codigo viejo.
6. Abre en esta PC el archivo:

   `C:\Users\ddval\OneDrive\Documents\D&D PREMIER SOLUTIONS CORP\cloudflare\app-palacios-worker.js`

7. Copia TODO el contenido de ese archivo.
8. Pegalo completo en Cloudflare.
9. Dale a `Deploy`.

Importante: no pegues solo una parte. Si Cloudflare se queda con el codigo viejo, la app seguira usando iframe y volvera el error `script.google.com refused to connect`.

## Rutas en Cloudflare

Despues de desplegar el Worker, revisa que tenga estas rutas:

- `app.palaciospowersystems.com/*`
- `ddpremiersolutionscorp.com/*`
- `www.ddpremiersolutionscorp.com/*`
- `app.ddpremiersolutionscorp.com/*`

Si alguna ruta no esta, agregala en:

`Worker > Settings > Domains & Routes`

## DNS

En Cloudflare DNS, los subdominios deben estar proxied por Cloudflare, con la nube naranja activa.

Revisa especialmente:

- `app.palaciospowersystems.com`
- `app.ddpremiersolutionscorp.com`

## Redirect Rules

No debe quedar una regla vieja que redirija `app.palaciospowersystems.com` directo al link de Apps Script.

Si existe una regla vieja hacia:

`https://script.google.com/macros/s/.../exec`

desactivala o borrala.

El dominio debe pasar por el Worker, no por una Redirect Rule.

## Como probar que quedo bien

Abre:

- `https://app.palaciospowersystems.com/`
- `https://app.palaciospowersystems.com/manifest.webmanifest`
- `https://app.palaciospowersystems.com/service-worker.js`

Tambien prueba:

- `https://ddpremiersolutionscorp.com/`
- `https://app.ddpremiersolutionscorp.com/`

## Como saber si todavia esta mal

Todavia esta usando el Worker viejo si pasa cualquiera de estas cosas:

- Sale `script.google.com refused to connect`.
- La pagina abre en blanco.
- El login no carga y se queda en loading.
- En el codigo fuente aparece un `<iframe>` apuntando a `script.google.com`.

El Worker nuevo no debe depender de un iframe directo a `script.google.com`.

## PWA

Cuando el Worker nuevo este activo:

1. Abre `https://app.palaciospowersystems.com/` en Chrome o Edge.
2. Usa el menu del navegador.
3. Dale a `Install app` o `Instalar app`.

En iPhone:

1. Abre `https://app.palaciospowersystems.com/` en Safari.
2. Dale al boton de compartir.
3. Dale a `Add to Home Screen`.

## Nota importante

No hace falta hacer redeploy de Apps Script para este cambio. Este cambio es solamente de Cloudflare Worker.

Si despues de hacer `Deploy` sigue saliendo el error, casi seguro Cloudflare todavia esta sirviendo el Worker viejo o la ruta apunta a otro Worker.
