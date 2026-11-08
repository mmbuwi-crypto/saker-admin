// Offline shell for the school app: lets the app open (and the teacher sign in) with no internet.
// Only touches files from the school's own site. Supabase and other sites are never cached here.
const VERSION = "sbcc-shell-v1";

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(["/", "/index.html", "/manifest.webmanifest"])).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  if (new URL(req.url).origin !== self.location.origin) return;

  // Opening the app: try the network for 3.5 s (picks up new versions), else use the saved copy
  if (req.mode === "navigate") {
    e.respondWith((async () => {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 3500);
      try {
        const res = await fetch(req, { signal: ctrl.signal });
        clearTimeout(timer);
        if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put("/index.html", copy)); }
        return res;
      } catch {
        clearTimeout(timer);
        return (await caches.match("/index.html")) || fetch(req);
      }
    })());
    return;
  }

  // Scripts, styles, images: saved copy first, saved as they are used
  e.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
      return res;
    }))
  );
});
