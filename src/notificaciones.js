import { initializeApp } from "firebase/app";
import { getMessaging, getToken, isSupported } from "firebase/messaging";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: `${import.meta.env.VITE_FIREBASE_PROJECT_ID}.firebaseapp.com`,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

let appInstanciado = null;
function obtenerApp() {
  if (!appInstanciado) appInstanciado = initializeApp(firebaseConfig);
  return appInstanciado;
}

// Pide permiso al navegador y, si dice que sí, devuelve el "número de
// buzón" (token) único de este móvil/navegador — eso es lo que se guarda
// para poder mandarle avisos más adelante.
export async function pedirPermisoYObtenerToken() {
  try {
    if (!("Notification" in window)) {
      return { ok: false, motivo: "Este navegador no admite notificaciones." };
    }
    const soportado = await isSupported();
    if (!soportado) {
      return { ok: false, motivo: "Este navegador no admite notificaciones push." };
    }

    const permiso = await Notification.requestPermission();
    if (permiso !== "granted") {
      return { ok: false, motivo: "No has dado permiso para las notificaciones." };
    }

    const registro = await navigator.serviceWorker.register("/firebase-messaging-sw.js");
    const messaging = getMessaging(obtenerApp());
    const token = await getToken(messaging, {
      vapidKey: import.meta.env.VITE_FIREBASE_VAPID_KEY,
      serviceWorkerRegistration: registro,
    });

    if (!token) {
      return { ok: false, motivo: "No se ha podido generar el token de notificaciones." };
    }
    return { ok: true, token };
  } catch (err) {
    return { ok: false, motivo: String(err?.message || err) };
  }
}
