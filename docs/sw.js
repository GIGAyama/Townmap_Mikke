/**
 * みっけ！ シェル用 Service Worker（最小構成）
 * キャッシュするのはシェル資産のみ。
 * GAS（script.google.com / googleusercontent.com）と gsi/client は絶対にキャッシュしない。
 */
/* 【重要】キャッシュの掃除は、かならず自アプリのぶんだけに限る。
 *
 * gigayama.github.io は数十本の学習アプリが同じドメインを共有している。
 * ブラウザのキャッシュはドメイン単位なので、caches.keys() はこのアプリのものだけでなく、
 * 同居する全アプリのキャッシュを返す。
 *
 * これまでは「CACHE_NAME 以外ぜんぶ」を消していたため、みっけ！を開いて
 * 新しい Service Worker が有効になった瞬間、その端末に入っていた
 * 児童むけアプリ（Qalc・KANJI_Town など）のオフライン用データまで消えていた。
 * 児童がオフラインで開いても起動せず、しかも原因がそのアプリ側に見えないため
 * 「たまに開かなくなる」という再現しにくい不具合になっていた。 */
const CACHE_PREFIX = 'mikke-shell-';
const CACHE_NAME = CACHE_PREFIX + 'v5';
const SHELL_ASSETS = ['./', './index.html', './config.js', './manifest.webmanifest', './icon.svg'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys
        .filter((k) => k.startsWith(CACHE_PREFIX) && k !== CACHE_NAME)
        .map((k) => caches.delete(k)))   // ← 自アプリ分だけ削除
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  // 同一オリジンのシェル資産のみ扱う。クロスオリジン（GAS / gsi 等）はブラウザに素通しする
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;
  // 診断ページは常に最新をネットワークから取得する（キャッシュ対象外）
  if (url.pathname.endsWith('/diag.html')) return;

  event.respondWith(
    caches.match(event.request, { ignoreSearch: url.pathname.endsWith('/') || url.pathname.endsWith('index.html') }).then((hit) => {
      const fetchAndUpdate = fetch(event.request).then((res) => {
        if (res && res.ok) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return res;
      }).catch(() => hit);
      // cache-first（裏で更新）
      return hit || fetchAndUpdate;
    })
  );
});
