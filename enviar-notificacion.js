const admin = require("firebase-admin");

if (!admin.apps.length) {
  const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_JSON);
  admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
  });
}

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }
  try {
    const { tokens, titulo, mensaje } = JSON.parse(event.body || "{}");
    if (!tokens || tokens.length === 0) {
      return { statusCode: 200, body: JSON.stringify({ ok: true, enviados: 0 }) };
    }
    const respuesta = await admin.messaging().sendEachForMulticast({
      tokens,
      notification: { title: titulo || "El Jaralillo", body: mensaje || "" },
    });
    return {
      statusCode: 200,
      body: JSON.stringify({ ok: true, enviados: respuesta.successCount }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ ok: false, error: String(err?.message || err) }),
    };
  }
};
