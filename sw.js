const CACHE_NAME='jh-assist-1.1.0-v1';
const APP_SHELL=[
  './','./index.html','./app.js','./sync-config.js','./manifest.webmanifest',
  './icon-192.png','./icon-512.png','./vendor/msal-browser.min.js'
];

self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll(APP_SHELL)).then(()=>self.skipWaiting()));
});

self.addEventListener('activate',event=>{
  event.waitUntil(
    caches.keys()
      .then(keys=>Promise.all(keys.filter(key=>key!==CACHE_NAME).map(key=>caches.delete(key))))
      .then(()=>self.clients.claim())
  );
});

self.addEventListener('fetch',event=>{
  const url=new URL(event.request.url);
  if(event.request.method!=='GET')return;
  if(url.origin!==self.location.origin)return;

  if(event.request.mode==='navigate'||url.pathname.endsWith('/index.html')||url.pathname.endsWith('/sync-config.js')){
    event.respondWith(
      fetch(event.request).then(response=>{
        const copy=response.clone();
        caches.open(CACHE_NAME).then(cache=>cache.put(event.request,copy));
        return response;
      }).catch(()=>caches.match(event.request).then(found=>found||caches.match('./index.html')))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached=>cached||fetch(event.request).then(response=>{
      const copy=response.clone();
      caches.open(CACHE_NAME).then(cache=>cache.put(event.request,copy));
      return response;
    }))
  );
});
