const C="pgen-v2",L=["./","index.html","manifest.json","icon.svg","pptxgen.bundle.js"];
self.addEventListener("install",e=>e.waitUntil(caches.open(C).then(c=>Promise.all(L.map(u=>c.add(u).catch(()=>{}))))));
self.addEventListener("fetch",e=>e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request))));
