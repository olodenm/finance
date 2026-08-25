/* Service worker приложения «Финансы».
   Стратегия: stale-while-revalidate — мгновенная отдача из кэша,
   обновление подтягивается в фоне и применяется при следующем запуске.
   При правке index.html поднимай номер версии, иначе старая копия
   может задержаться у уже установленных клиентов на один запуск. */
const VERSION = "fin-v11";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png",
  /* Шрифты лежат на своём домене специально: обработчик fetch ниже
     пропускает всё чужое, поэтому через CDN они офлайн не работали. */
  "./fonts/inter-latin.woff2",
  "./fonts/inter-cyrillic.woff2",
  "./fonts/unbounded-latin.woff2",
  "./fonts/unbounded-cyrillic.woff2",
  "./fonts/mono-latin.woff2",
  "./fonts/mono-cyrillic.woff2"
];

self.addEventListener("install", e => {
  /* Каждый файл кладём в кэш отдельно. У addAll всё или ничего: одна
     опечатка в имени или один отсутствующий файл молча оставляли кэш
     пустым, и приложение переставало работать офлайн, ничего не сообщив. */
  e.waitUntil(
    caches.open(VERSION)
      .then(c => Promise.allSettled(ASSETS.map(a => c.add(a))))
      .then(res => {
        const lost = res.map((r, i) => r.status === "rejected" ? ASSETS[i] : null).filter(Boolean);
        if (lost.length) console.warn("sw: не попали в кэш —", lost.join(", "));
        return self.skipWaiting();
      })
      .catch(() => self.skipWaiting())
  );
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  if (new URL(req.url).origin !== self.location.origin) return;

  e.respondWith(
    caches.match(req).then(cached => {
      const network = fetch(req).then(res => {
        if (res && res.status === 200 && res.type === "basic") {
          const copy = res.clone();
          caches.open(VERSION).then(c => c.put(req, copy));
        }
        return res;
      }).catch(() => cached);
      return cached || network;
    })
  );
});
