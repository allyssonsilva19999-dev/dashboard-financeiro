// Mantém o Organizador funcionando offline quando instalado na tela de início.
// Rede primeiro (sempre a versão mais nova); sem rede, usa a cópia guardada.
const CACHE = "organizador-v1";
const ARQUIVOS = ["./", "./index.html", "./organizador.html", "./manifest.webmanifest", "./icone-180.png", "./icone-192.png", "./icone-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(chaves => Promise.all(chaves.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== location.origin) return;
  e.respondWith(fetch(req)
    .then(res => {
      if (res.ok) { const copia = res.clone(); caches.open(CACHE).then(c => c.put(req, copia)); }
      return res;
    })
    .catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match("./organizador.html"))));
});
