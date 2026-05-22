// DHOOM SERVICE WORKER — Auto-update on every GitHub push
// Version is set to current timestamp at build time; GitHub Actions or manual edit triggers refresh

const CACHE_VERSION = "dhoom-v__BUILD__";
const CACHE_ASSETS = [
  "./",
  "./index.html",
  "./users.json",
  "./stockists_master.csv",
  "./manifest.json"
];

self.addEventListener("install", function(e) {
  e.waitUntil(
    caches.open(CACHE_VERSION).then(function(cache) {
      return cache.addAll(CACHE_ASSETS);
    }).then(function() {
      return self.skipWaiting(); // activate immediately
    })
  );
});

self.addEventListener("activate", function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_VERSION; })
            .map(function(k) { return caches.delete(k); })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

self.addEventListener("fetch", function(e) {
  // Network first for app files so updates are instant
  if (e.request.url.indexOf(self.location.origin) === 0) {
    e.respondWith(
      fetch(e.request).then(function(networkRes) {
        var cloned = networkRes.clone();
        caches.open(CACHE_VERSION).then(function(cache) { cache.put(e.request, cloned); });
        return networkRes;
      }).catch(function() {
        return caches.match(e.request);
      })
    );
  }
});
