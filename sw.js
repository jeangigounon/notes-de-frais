// Service worker : l'application fonctionne hors ligne (sauf l'analyse IA et l'envoi).
// Fichiers du site : réseau d'abord (mises à jour immédiates), cache en secours.
// Bibliothèques externes (versions figées) : cache d'abord.

const VERSION = 'ndf-v4';
const SHELL = [
  './',
  'index.html',
  'manifest.webmanifest',
  'css/app.css',
  'icons/icon.svg',
  'icons/apple-touch-icon.png',
  'js/app.js',
  'js/ui.js',
  'js/lib.js',
  'js/fields.js',
  'js/store.js',
  'js/image.js',
  'js/ai.js',
  'js/xlsx.js',
  'js/mail.js',
  'js/sync.js',
  'js/capture.js',
  'js/transfer.js',
  'js/pdf.js',
  'js/fx.js',
  'js/templateAI.js',
  'js/views/home.js',
  'js/views/expenses.js',
  'js/views/expenseSheet.js',
  'js/views/sendSheet.js',
  'js/views/clients.js',
  'js/views/mapper.js',
  'js/views/settings.js',
];
const CDN_HOSTS = ['cdn.jsdelivr.net', 'cdnjs.cloudflare.com', 'cdn.sheetjs.com', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(VERSION).then(cache => cache.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  if (url.origin === self.location.origin) {
    // « no-cache » : on revalide auprès de GitHub Pages pour ne jamais mélanger d'anciens fichiers à une mise à jour.
    const fresh = req.mode === 'navigate' ? fetch(req.url, { cache: 'no-cache', credentials: 'same-origin' }) : fetch(req, { cache: 'no-cache' });
    event.respondWith(
      fresh
        .then(res => {
          if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
          return res;
        })
        .catch(async () => (await caches.match(req, { ignoreSearch: true })) || (req.mode === 'navigate' ? caches.match('index.html') : Response.error())),
    );
    return;
  }

  if (CDN_HOSTS.includes(url.hostname)) {
    event.respondWith(
      caches.match(req).then(hit => hit || fetch(req).then(res => {
        if (res.ok || res.type === 'opaque') { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
        return res;
      })),
    );
  }
});
