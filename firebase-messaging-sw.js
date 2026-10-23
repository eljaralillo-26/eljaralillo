importScripts("https://www.gstatic.com/firebasejs/10.13.0/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/10.13.0/firebase-messaging-compat.js");

// Estos datos no son secretos (son los mismos que ya usa la web), hace
// falta repetirlos aquí porque este archivo se ejecuta aparte, en
// segundo plano, y no puede leer las variables de entorno de la app.
firebase.initializeApp({
  apiKey: "AIzaSyBmFPlT3Zq8A0gKyx5ivK94VCZu85MVHNI",
  authDomain: "eljaralillo-2f879.firebaseapp.com",
  projectId: "eljaralillo-2f879",
  storageBucket: "eljaralillo-2f879.firebasestorage.app",
  messagingSenderId: "41439645428",
  appId: "1:41439645428:web:7edc4d8d9ddeb2acb458a1",
});

const messaging = firebase.messaging();

messaging.onBackgroundMessage((payload) => {
  const titulo = payload.notification?.title || "El Jaralillo";
  const cuerpo = payload.notification?.body || "";
  self.registration.showNotification(titulo, {
    body: cuerpo,
    icon: "/icono.png",
    badge: "/icono.png",
  });
});
