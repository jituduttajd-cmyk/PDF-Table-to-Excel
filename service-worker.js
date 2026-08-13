const CACHE="pdf-to-excel-v2";
const SHELL=["./","./index.html","./app.js","./manifest.json","./icons/icon-192.png","./icons/icon-512.png"];

self.addEventListener("install",event=>{
  event.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)));
  self.skipWaiting();
});
self.addEventListener("activate",event=>{
  event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch",event=>{
  if(event.request.method!=="GET")return;
  const url=new URL(event.request.url);
  if(url.origin!==location.origin)return;
  event.respondWith(caches.match(event.request).then(cached=>{
    return cached||fetch(event.request).then(response=>{
      const copy=response.clone();
      caches.open(CACHE).then(c=>c.put(event.request,copy));
      return response;
    });
  }));
});
