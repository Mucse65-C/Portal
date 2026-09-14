const CACHE_NAME = 'cse65c-offline-pwa-v2';
const APP_SHELL = [
 './',
 './index.html',
 './manifest.webmanifest',
 './icons/icon-192.png',
 './icons/icon-512.png'
];

self.addEventListener('install', event=>{
 self.skipWaiting();
 event.waitUntil(caches.open(CACHE_NAME).then(c=>c.addAll(APP_SHELL).catch(()=>{})));
});

self.addEventListener('activate', event=>{
 event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', event=>{
 const req=event.request;
 if(req.method!=='GET') return;

 event.respondWith(
  fetch(req)
   .then(res=>{
     const copy=res.clone();
     caches.open(CACHE_NAME).then(c=>c.put(req,copy));
     return res;
   })
   .catch(()=>caches.match(req).then(r=>r || caches.match('./index.html')))
 );
});
