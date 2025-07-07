self.addEventListener('install', event => {
  event.waitUntil(
    caches.open('vbaber-cache').then(cache => {
      return cache.addAll([
        '/',
        '/index.html',
        '/style.css',
        '/app.js',
        '/config.js'
      ]);
    })
  );
});
