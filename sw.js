/* خدمة التثبيت — لا تخزّن شيئاً كي يظهر أي تحديث فوراً */
self.addEventListener("install", function () { self.skipWaiting(); });
self.addEventListener("activate", function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener("fetch", function () { /* مرور دائم للشبكة */ });
