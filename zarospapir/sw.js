// Lipóti zárós papír – alkalmazás (PWA) háttérszolgáltatás.
// Mindig a weboldal legfrissebb változatát tölti (hálózat először), így az alkalmazás és a weboldal 100%-ban ugyanaz.
// A tárolt másolat csak akkor kell, ha épp nincs internet.
const TAR = 'zarospapir-v1';
const ALAP = ['./', './index.html', './kep.html', './manifest.webmanifest', './ikon/ikon-192.png', './ikon/ikon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(TAR).then(c => c.addAll(ALAP)).catch(() => {}).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(n => n !== TAR).map(n => caches.delete(n)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const k = e.request;
  if (k.method !== 'GET') return;
  const u = new URL(k.url);
  if (u.origin !== location.origin) return;   // a Google (súgók, archívum) kérései érintetlenül mennek
  e.respondWith((async () => {
    try {
      const v = await fetch(k, { cache: 'no-store' });
      if (v && v.ok && !u.searchParams.has('_')) {   // a változat-ellenőrző kérést nem tároljuk
        const kulcs = k.mode === 'navigate' ? new Request(u.origin + u.pathname) : k;
        const c = await caches.open(TAR); c.put(kulcs, v.clone()).catch(() => {});
      }
      return v;
    } catch (err) {
      const c = await caches.open(TAR);
      return (await c.match(k, { ignoreSearch: true })) || (k.mode === 'navigate' ? await c.match('./index.html') || await c.match('./') : undefined) || Response.error();
    }
  })());
});
