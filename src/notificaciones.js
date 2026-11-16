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

function conLimiteDeTiempo(promesa, ms, mensaje) {
  return Promise.race([
    promesa,
    new Promise((_, reject) => setTimeout(() => reject(new Error(mensaje)), ms)),
  ]);
}

// Pide permiso al navegador y, si dice que sí, devuelve el "número de
// buzón" (token) único de este móvil/navegador. Nunca se queda colgado:
// si algo tarda demasiado, avisa del motivo en vez de pensar para siempre.
export async function pedirPermisoYObtenerToken() {
  try {
    if (!("Notification" in window)) {
      return {
        ok: false,
        motivo:
          "Este navegador no admite notificaciones. En iPhone hay que añadir la web a la pantalla de inicio y abrirla desde ahí.",
      };
    }
    if (!import.meta.env.VITE_FIREBASE_VAPID_KEY) {
      return {
        ok: false,
        motivo: "Todavía falta terminar de configurar las notificaciones en Firebase (clave VAPID).",
      };
    }

    const soportado = await conLimiteDeTiempo(
      isSupported(),
      5000,
      "Ha tardado demasiado en comprobar si el navegador admite notificaciones."
    );
    if (!soportado) {
      return { ok: false, motivo: "Este navegador no admite notificaciones push." };
    }

    const permiso = await conLimiteDeTiempo(
      Notification.requestPermission(),
      30000,
      "No has respondido al permiso de notificaciones a tiempo."
    );
    if (permiso !== "granted") {
      return { ok: false, motivo: "No has dado permiso para las notificaciones." };
    }

    const registro = await conLimiteDeTiempo(
      navigator.serviceWorker.register("/firebase-messaging-sw.js"),
      10000,
      "Ha tardado demasiado en preparar el servicio de notificaciones."
    );
    const messaging = getMessaging(obtenerApp());
    const token = await conLimiteDeTiempo(
      getToken(messaging, {
        vapidKey: import.meta.env.VITE_FIREBASE_VAPID_KEY,
        serviceWorkerRegistration: registro,
      }),
      15000,
      "Ha tardado demasiado en generar el token de notificaciones."
    );

    if (!token) {
      return { ok: false, motivo: "No se ha podido generar el token de notificaciones." };
    }
    return { ok: true, token };
  } catch (err) {
    return { ok: false, motivo: String(err?.message || err) };
  }
}
