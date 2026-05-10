// 簡易的 Service Worker 讓瀏覽器識別為 PWA
const CACHE_NAME = 'image-extractor-v1';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  // 對於此工具，主要是在本地處理，不需要複雜的快取邏輯
  // 但 Service Worker 必須存在才能觸發安裝圖示
  event.respondWith(fetch(event.request));
});
