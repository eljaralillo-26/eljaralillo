import admin from "firebase-admin";

const responder = (statusCode, cuerpo) => ({
  statusCode,
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(cuerpo),
});

function iniciarFirebase() {
  if (!admin.apps.length) {
    const cuenta = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_JSON);
    admin.initializeApp({ credential: admin.credential.cert(cuenta) });
  }
}

export const handler = async (event) => {
  if (event.httpMethod !== "POST") return responder(405, { ok: false });
  try {
    iniciarFirebase();
    const { tokens, titulo, mensaje } = JSON.parse(event.body || "{}");
    if (!tokens || tokens.length === 0) return responder(200, { ok: true, enviados: 0 });
    const respuesta = await admin.messaging().sendEachForMulticast({
      tokens,
      notification: { title: titulo || "El Jaralillo", body: mensaje || "" },
    });
    return responder(200, { ok: true, enviados: respuesta.successCount });
  } catch (err) {
    return responder(500, { ok: false, error: String(err?.message || err) });
  }
};
