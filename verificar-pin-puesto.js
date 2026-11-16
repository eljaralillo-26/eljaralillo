import { timingSafeEqual } from "node:crypto";

const responder = (statusCode, cuerpo) => ({
  statusCode,
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(cuerpo),
});

// El PIN del puesto vive SOLO aquí, en Netlify (variable PUESTO_PIN).
// Nunca viaja dentro de la web, así que nadie puede sacarlo mirando el código.
export const handler = async (event) => {
  if (event.httpMethod !== "POST") return responder(405, { ok: false });

  const correcto = process.env.PUESTO_PIN;
  if (!correcto) return responder(500, { ok: false, error: "no-configurado" });

  let pin = "";
  try {
    pin = String(JSON.parse(event.body || "{}").pin || "");
  } catch {
    return responder(400, { ok: false });
  }

  const a = Buffer.from(pin);
  const b = Buffer.from(String(correcto));
  const ok = a.length === b.length && timingSafeEqual(a, b);

  // Si falla, se hace esperar un segundo: frena a quien intente adivinar.
  if (!ok) await new Promise((r) => setTimeout(r, 1000));
  return responder(200, { ok });
};
