/* Calidad Samitex · funciona sin señal: guarda la app en el equipo y la abre desde ahí */
const VERSION = 'cs-v17';
const PROPIOS = ['./', './index.html', './config.js', './manifest.webmanifest', './icon-192.png', './icon-512.png', './admin.webmanifest', './icon-adm-192.png', './icon-adm-512.png'];
const EXTERNOS = ['https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(async c => {
    await c.addAll(PROPIOS);
    for (const u of EXTERNOS) { try { await c.add(new Request(u, { mode: 'no-cors' })); } catch (x) {} }
  }).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;                                   // los envíos al servidor nunca se guardan
  const u = new URL(r.url);
  if (u.hostname.endsWith('script.google.com') || u.hostname.endsWith('googleusercontent.com')) return;
  if (r.mode === 'navigate') {                                       // la página: primero la red, si no hay señal la copia guardada
    e.respondWith(fetch(r).then(res => { const cp = res.clone(); caches.open(VERSION).then(c => c.put('./index.html', cp)); return res; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(r).then(hit => {                          // archivos: copia guardada y se actualiza en segundo plano
    const net = fetch(r).then(res => { if (res && (res.ok || res.type === 'opaque')) { const cp = res.clone(); caches.open(VERSION).then(c => c.put(r, cp)); } return res; }).catch(() => hit);
    return hit || net;
  }));
});
