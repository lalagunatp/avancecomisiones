/* Mi Avance de Instalaciones — service worker
   Sube el número de versión cada vez que cambies index.html,
   así los teléfonos que ya la tienen instalada reciben la actualización. */
const CACHE = 'tp-avance-v9';
const ARCHIVOS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './icon-mask.png'];

/* Se guarda uno por uno: si falta un archivo, no se cae toda la instalación */
self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    for (const a of ARCHIVOS) { try { await c.add(a); } catch (err) {} }
    self.skipWaiting();
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const llaves = await caches.keys();
    await Promise.all(llaves.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  let url;
  try { url = new URL(req.url); } catch (err) { return; }

  /* Los datos de Google siempre van a la red, nunca al caché */
  if (url.hostname.indexOf('google') > -1) return;

  /* Al abrir la app: primero el caché, así arranca al instante */
  if (req.mode === 'navigate') {
    e.respondWith((async () => {
      const c = await caches.open(CACHE);
      const guardado = await c.match('./index.html');
      if (guardado) {
        fetch(req).then(r => { if (r && r.ok) c.put('./index.html', r.clone()); }).catch(() => {});
        return guardado;
      }
      try {
        const r = await fetch(req);
        if (r && r.ok) c.put('./index.html', r.clone());
        return r;
      } catch (err) {
        return new Response(
          '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
          + '<body style="background:#150E2E;color:#F2EEFF;font-family:system-ui;padding:32px;text-align:center">'
          + '<h1 style="font-size:1.4rem">Sin conexión</h1>'
          + '<p style="color:#A99CD4">Vuelve a abrir la app cuando tengas señal.</p></body>',
          { headers: { 'Content-Type': 'text/html;charset=utf-8' } }
        );
      }
    })());
    return;
  }

  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    const guardado = await c.match(req);
    if (guardado) return guardado;
    try {
      const r = await fetch(req);
      if (r && r.ok && url.origin === self.location.origin) c.put(req, r.clone());
      return r;
    } catch (err) {
      return new Response('', { status: 504 });
    }
  })());
});
