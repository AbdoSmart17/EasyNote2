// Service Worker لتطبيق إدارة نقاط التلاميذ
const CACHE_NAME = 'EasyNote-app-v2.0.0';
const urlsToCache = [
    '/',
    '/index.html',
    '/style.css',
    '/script.js',
    '/manifest.json',
    '/assets/favicon.ico',
    '/assets/icon-192.png',
    '/assets/icon-512.png',
    '/assets/hero-image.svg',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    'https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap'
];

// تثبيت Service Worker
self.addEventListener('install', function(event) {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(function(cache) {
                console.log('✅ تم فتح الكاش');
                return cache.addAll(urlsToCache);
            })
            .then(() => {
                console.log('✅ تم تثبيت جميع الموارد في الكاش');
                return self.skipWaiting();
            })
    );
});

// تفعيل Service Worker
self.addEventListener('activate', function(event) {
    event.waitUntil(
        caches.keys().then(function(cacheNames) {
            return Promise.all(
                cacheNames.map(function(cacheName) {
                    if (cacheName !== CACHE_NAME) {
                        console.log('🗑️ حذف الكاش القديم:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(() => {
            console.log('✅ Service Worker مفعل');
            return self.clients.claim();
        })
    );
});

// اعتراض الطلبات
self.addEventListener('fetch', function(event) {
    // تجاهل طلبات البيانات الخارجية غير الأساسية
    if (!event.request.url.startsWith(self.location.origin) && 
        !event.request.url.includes('xlsx.full.min.js') &&
        !event.request.url.includes('fonts.googleapis.com')) {
        return;
    }

    event.respondWith(
        caches.match(event.request)
            .then(function(response) {
                // إذا وجدت الاستجابة في الكاش، أرجعها
                if (response) {
                    return response;
                }

                // استنساخ الطلب
                const fetchRequest = event.request.clone();

                return fetch(fetchRequest).then(
                    function(response) {
                        // التحقق من أن الاستجابة صالحة
                        if(!response || response.status !== 200 || response.type !== 'basic') {
                            return response;
                        }

                        // استنساخ الاستجابة
                        const responseToCache = response.clone();

                        caches.open(CACHE_NAME)
                            .then(function(cache) {
                                cache.put(event.request, responseToCache);
                            });

                        return response;
                    }
                ).catch(function() {
                    // إذا فشل التحميل، يمكن إرجاع صفحة بديلة
                    if (event.request.destination === 'document') {
                        return caches.match('/');
                    }
                });
            })
    );
});

// التعامل مع الرسائل
self.addEventListener('message', function(event) {
    if (event.data && event.data.type === 'SKIP_WAITING') {
        self.skipWaiting();
    }
});