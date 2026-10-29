const C="pgen-v1",L=["./","index.html","manifest.json","icon.svg","https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js"];
self.addEventListener("install",e=>e.waitUntil(caches.open(C).then(c=>Promise.all(L.map(u=>c.add(u).catch(()=>{}))))));
self.addEventListener("fetch",e=>e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request))));
