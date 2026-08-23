import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { Leaf, Stamp, Bell, LogOut, ShieldCheck, Clock, Coins, ChevronRight, Send, X, Loader2, Upload, Download, CheckCircle2, AlertTriangle, UserPlus, Scale, ArrowRight, FileText, PenTool, Eraser, RotateCcw, Mic, Package, MessageCircle, Pin, User, Droplet, KeyRound, Users, ClipboardList, RefreshCw } from "lucide-react";

// ---------------------------------------------------------------------------
// Datos de ejemplo (en la app real esto vendría del listado de la báscula)
// ---------------------------------------------------------------------------
const SEED_AGRICULTORES = {
  "12345678A": { nombre: "Antonio Ruiz Gómez", poblacion: "Jaralillo" },
  "23456789B": { nombre: "María Fernández López", poblacion: "Jaralillo" },
  "34567890C": { nombre: "José Manuel Torres", poblacion: "Villanueva" },
  "45678901D": { nombre: "Carmen Sánchez Ortega", poblacion: "Jaralillo" },
};

const SEED_ALBARANES = {
  "12345678A": [
    { n: 214, fecha: "2026-08-14", kilos: 820, precio: 0.62 },
    { n: 189, fecha: "2026-08-11", kilos: 1040, precio: 0.6 },
  ],
  "23456789B": [{ n: 221, fecha: "2026-08-15", kilos: 610, precio: 0.62 }],
  "34567890C": [
    { n: 205, fecha: "2026-08-13", kilos: 1320, precio: 0.61 },
  ],
  "45678901D": [],
};

const PUESTO_PIN = "1957";
const PUESTO_DNI = "PUESTO";
const PUESTO_NOMBRE = "El Jaralillo";

const LOGO_B64 =
  "data:image/png;base64,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";

// ---------------------------------------------------------------------------
// Utilidades de almacenamiento
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Guardado en la nube: Firestore (base de datos real de Google), hablando
// directamente con su API por internet — sin necesitar instalar nada.
// Cada "cajón" de la demo (albaranes, avisos, chat...) es ahora un documento
// real dentro de la colección "datos", con un único campo de texto que
// guarda toda la información en formato JSON, igual que antes.
// ---------------------------------------------------------------------------
const FIREBASE_PROJECT_ID = import.meta.env.VITE_FIREBASE_PROJECT_ID;
const FIREBASE_API_KEY = import.meta.env.VITE_FIREBASE_API_KEY;
const FIRESTORE_BASE = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/datos`;
const FIRESTORE_ROOT = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents`;

// Los mensajes del chat se guardan CADA UNO en su propia ficha (dentro de
// la colección "mensajes"), en vez de todos juntos en una sola — así no
// hay un límite de tamaño combinado entre todos los mensajes de todo el
// mundo, como sí pasaba antes (por eso fallaba el segundo audio).
async function listarMensajesRemotos() {
  try {
    const res = await fetch(`${FIRESTORE_ROOT}/mensajes?pageSize=300&key=${FIREBASE_API_KEY}`);
    if (!res.ok) return { hay: false };
    const data = await res.json();
    const docs = data.documents || [];
    const mensajes = docs
      .map((d) => {
        const raw = d.fields?.valor?.stringValue;
        return raw ? JSON.parse(raw) : null;
      })
      .filter(Boolean);
    return { hay: true, value: mensajes };
  } catch {
    return { hay: false };
  }
}

async function guardarMensajeRemoto(mensaje) {
  try {
    const cuerpo = { fields: { valor: { stringValue: JSON.stringify(mensaje) } } };
    const res = await fetch(
      `${FIRESTORE_ROOT}/mensajes/${encodeURIComponent(String(mensaje.id))}?key=${FIREBASE_API_KEY}`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(cuerpo),
      }
    );
    return res.ok;
  } catch {
    return false;
  }
}

async function borrarMensajeRemoto(id) {
  try {
    await fetch(`${FIRESTORE_ROOT}/mensajes/${encodeURIComponent(String(id))}?key=${FIREBASE_API_KEY}`, {
      method: "DELETE",
    });
    return true;
  } catch {
    return false;
  }
}

async function loadShared(key, fallback) {
  try {
    const res = await fetch(`${FIRESTORE_BASE}/${encodeURIComponent(key)}?key=${FIREBASE_API_KEY}`);
    if (!res.ok) return fallback;
    const doc = await res.json();
    const raw = doc?.fields?.valor?.stringValue;
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

// Para archivos grandes (como el PDF de una factura), un primer intento
// puede fallar por una lectura demasiado rápida o una conexión lenta.
// Reintenta unas cuantas veces, con una pequeña pausa, antes de rendirse.
async function loadSharedConReintentos(key, intentos = 6) {
  for (let i = 0; i < intentos; i++) {
    const valor = await loadShared(key, undefined);
    if (valor !== undefined) return valor;
    if (i < intentos - 1) {
      await new Promise((resolve) => setTimeout(resolve, 900));
    }
  }
  return null;
}

// Para el refresco periódico: si la lectura falla (por ejemplo, por exceso
// de peticiones), NUNCA queremos pisar los datos buenos que ya tenemos en
// pantalla con el valor de fábrica. Por eso esta versión avisa si ha
// fallado (hay:false) en vez de devolver un valor por defecto silencioso.
async function pollShared(key) {
  try {
    const res = await fetch(`${FIRESTORE_BASE}/${encodeURIComponent(key)}?key=${FIREBASE_API_KEY}`);
    if (!res.ok) return { hay: false };
    const doc = await res.json();
    const raw = doc?.fields?.valor?.stringValue;
    if (!raw) return { hay: false };
    return { hay: true, value: JSON.parse(raw) };
  } catch {
    return { hay: false };
  }
}

async function saveShared(key, value) {
  try {
    const cuerpo = { fields: { valor: { stringValue: JSON.stringify(value) } } };
    const res = await fetch(`${FIRESTORE_BASE}/${encodeURIComponent(key)}?key=${FIREBASE_API_KEY}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(cuerpo),
    });
    return res.ok;
  } catch {
    // demo: si falla el guardado, seguimos en memoria
    return false;
  }
}

function descargarDataUrl(dataUrl, nombreArchivo) {
  try {
    const partes = dataUrl.split(",");
    const mimeMatch = partes[0].match(/:(.*?);/);
    const mime = mimeMatch ? mimeMatch[1] : "application/pdf";
    const binario = atob(partes[1]);
    const bytes = new Uint8Array(binario.length);
    for (let i = 0; i < binario.length; i++) bytes[i] = binario.charCodeAt(i);
    const blob = new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    // Abrir en pestaña nueva es lo más fiable en móvil: el propio teléfono
    // muestra su visor de PDF, y desde ahí el agricultor puede guardarlo o
    // compartirlo con el botón nativo. El intento de "descarga forzada" con
    // el atributo download falla en silencio en muchos navegadores móviles
    // (sobre todo iPhone) con archivos grandes como un PDF.
    const ventana = window.open(url, "_blank");
    if (!ventana) {
      // si el navegador bloquea la ventana emergente, probamos con descarga directa
      const a = document.createElement("a");
      a.href = url;
      a.download = nombreArchivo;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch {
    // último recurso: abrir el data URL tal cual
    window.open(dataUrl, "_blank");
  }
}

// Reduce la foto que sube el agricultor a un cuadrado pequeño antes de
// guardarla, para que no ocupe demasiado.
function redimensionarFoto(file, maxLado = 220) {
  return new Promise((resolve, reject) => {
    const lector = new FileReader();
    lector.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const lado = Math.min(img.width, img.height);
        const canvas = document.createElement("canvas");
        canvas.width = maxLado;
        canvas.height = maxLado;
        const ctx = canvas.getContext("2d");
        const sx = (img.width - lado) / 2;
        const sy = (img.height - lado) / 2;
        ctx.drawImage(img, sx, sy, lado, lado, 0, 0, maxLado, maxLado);
        resolve(canvas.toDataURL("image/jpeg", 0.85));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    lector.onerror = reject;
    lector.readAsDataURL(file);
  });
}

function encontrarColumna(headers, pistas) {
  return headers.find((h) => pistas.some((p) => h.toLowerCase().includes(p)));
}

// Busca, en todas las hojas del archivo, la fila que tenga a la vez una
// columna de DNI y una de KILOS (puede haber varias filas de datos del
// camión/almazara por encima, como en la guía real). Prioriza hojas cuyo
// nombre contenga "guía"/"guia".
function localizarTablaGuia(wb) {
  const nombresOrdenados = [...wb.SheetNames].sort((a, b) => {
    const aGuia = /gu[ií]a/i.test(a) ? 0 : 1;
    const bGuia = /gu[ií]a/i.test(b) ? 0 : 1;
    return aGuia - bGuia;
  });

  for (const nombreHoja of nombresOrdenados) {
    const hoja = wb.Sheets[nombreHoja];
    const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "" });

    for (let i = 0; i < Math.min(filas.length, 25); i++) {
      const filaTexto = filas[i].map((c) => String(c || "").toLowerCase());
      const tieneDni = filaTexto.some((c) => c.includes("dni"));
      const tieneKilos = filaTexto.some((c) => c.includes("kilo"));
      if (!tieneDni || !tieneKilos) continue;

      const headersReales = filas[i].map((c) => String(c || "").trim());
      const filasDatos = filas
        .slice(i + 1)
        .filter((f) => f.some((c) => String(c || "").trim() !== ""))
        .map((f) => {
          const obj = {};
          headersReales.forEach((h, idx) => {
            if (h) obj[h] = f[idx];
          });
          return obj;
        });

      let precioGlobal = null;
      for (const fila of filas) {
        for (let c = 0; c < fila.length; c++) {
          const texto = String(fila[c] || "").toLowerCase();
          if (texto.includes("precio del día") || texto.includes("precio del dia")) {
            for (let k = c + 1; k < fila.length; k++) {
              const v = fila[k];
              if (v !== "" && v !== null && !isNaN(Number(v))) {
                precioGlobal = Number(v);
                break;
              }
            }
          }
        }
      }

      return { filasDatos, headersReales, nombreHoja, precioGlobal };
    }
  }
  return null;
}

const money = (v) => v.toLocaleString("es-ES", { style: "currency", currency: "EUR" });
const fecha = (iso) =>
  new Date(iso + "T00:00:00").toLocaleDateString("es-ES", { day: "2-digit", month: "short" });

// ---------------------------------------------------------------------------
// Componente principal
// ---------------------------------------------------------------------------
export default function App() {
  const [ready, setReady] = useState(false);
  const [agricultores, setAgricultores] = useState(SEED_AGRICULTORES);
  const [albaranes, setAlbaranes] = useState(SEED_ALBARANES);
  const [avisos, setAvisos] = useState([]);
  const [mensajes, setMensajes] = useState([]);
  const [perfiles, setPerfiles] = useState({});
  const [codigo, setCodigo] = useState("07");
  const [presencia, setPresencia] = useState({});
  const [precios, setPrecios] = useState({ garrafa: 15, caja: 48 });
  const [pedidos, setPedidos] = useState([]);
  const [tiempoEspera, setTiempoEspera] = useState(15);
  const [kgTotales, setKgTotales] = useState(0);
  const [pendientes, setPendientes] = useState([]);
  const [facturas, setFacturas] = useState([]);
  const [vista, setVista] = useState("portal"); // portal | agricultor | puesto
  const [sesion, setSesion] = useState(null);
  const [notiToast, setNotiToast] = useState(null);
  const avisosCountRef = useRef(0);
  const mensajesCountRef = useRef(0);
  const agricultoresCountRef = useRef(0);
  const ultimaEdicionRef = useRef(0);
  const vistaRef = useRef(vista);
  const sesionRef = useRef(sesion);
  useEffect(() => {
    vistaRef.current = vista;
    sesionRef.current = sesion;
  }, [vista, sesion]);

  useEffect(() => {
    (async () => {
      const ag = await loadShared("jaralillo-agricultores", SEED_AGRICULTORES);
      const a = await loadShared("jaralillo-albaranes", SEED_ALBARANES);
      const av = await loadShared("jaralillo-avisos", []);
      const ch = await listarMensajesRemotos();
      const chVal = ch.hay ? ch.value : [];
      const pf = await loadShared("jaralillo-perfiles", {});
      const co = await loadShared("jaralillo-codigo", "07");
      const pr = await loadShared("jaralillo-presencia", {});
      const pc = await loadShared("jaralillo-precios", { garrafa: 15, caja: 48 });
      const pd = await loadShared("jaralillo-pedidos", []);
      const te = await loadShared("jaralillo-espera", 15);
      const kg = await loadShared("jaralillo-kg-totales", 0);
      const pend = await loadShared("jaralillo-albaranes-pendientes", []);
      const fact = await loadShared("jaralillo-facturas-index", []);
      setAgricultores(ag);
      setAlbaranes(a);
      setAvisos(av);
      setMensajes(chVal);
      setPerfiles(pf);
      setCodigo(co);
      setPresencia(pr);
      setPrecios(pc);
      setPedidos(pd);
      setTiempoEspera(te);
      setKgTotales(kg);
      setPendientes(pend);
      setFacturas(fact);
      avisosCountRef.current = av.length;
      mensajesCountRef.current = chVal.length;
      agricultoresCountRef.current = Object.keys(ag).length;
      setReady(true);
    })();
  }, []);

  // Refresca la presencia, los avisos, el chat y el resto de datos cada
  // pocos segundos. Usa pollShared: si una lectura falla (por ejemplo por
  // exceso de peticiones), NUNCA se pisan los datos que ya había en
  // pantalla — se deja como estaba y se reintenta en el siguiente ciclo.
  useEffect(() => {
    const id = setInterval(async () => {
      const pr = await pollShared("jaralillo-presencia");
      if (pr.hay) setPresencia(pr.value);

      // La ventana de gracia dura más que el propio intervalo de refresco,
      // para asegurar que se salta AL MENOS un ciclo completo tras cualquier
      // cambio hecho desde el puesto (dar de alta, mandar albaranes, etc.)
      // y así nunca se pisa con una lectura que todavía no se ha actualizado.
      const editadoHaceNada = Date.now() - ultimaEdicionRef.current < 15000;

      if (!editadoHaceNada) {
        const ag = await pollShared("jaralillo-agricultores");
        if (ag.hay && Object.keys(ag.value).length !== agricultoresCountRef.current) {
          setAgricultores(ag.value);
          agricultoresCountRef.current = Object.keys(ag.value).length;
        }

        const te = await pollShared("jaralillo-espera");
        if (te.hay) setTiempoEspera(te.value);

        const kg = await pollShared("jaralillo-kg-totales");
        if (kg.hay) setKgTotales(kg.value);

        const pend = await pollShared("jaralillo-albaranes-pendientes");
        if (pend.hay) setPendientes(pend.value);

        const alb = await pollShared("jaralillo-albaranes");
        if (alb.hay) setAlbaranes(alb.value);

        const fact = await pollShared("jaralillo-facturas-index");
        if (fact.hay) setFacturas(fact.value);
      }

      const av = await pollShared("jaralillo-avisos");
      if (av.hay) {
        if (av.value.length > avisosCountRef.current) {
          const nuevos = av.value.slice(0, av.value.length - avisosCountRef.current);
          const s = sesionRef.current;
          if (vistaRef.current === "agricultor" && s) {
            const relevante = nuevos.find(
              (a) =>
                a.destinatarios === "todos" ||
                (Array.isArray(a.destinatarios) && a.destinatarios.includes(s))
            );
            if (relevante) {
              setNotiToast({ tipo: "aviso", texto: `Nuevo aviso: ${relevante.mensaje}` });
            }
          }
          setAvisos(av.value);
        }
        avisosCountRef.current = av.value.length;
      }

      const ch = await listarMensajesRemotos();
      if (ch.hay) {
        const ordenados = [...ch.value].sort((a, b) => (a.ts || a.id) - (b.ts || b.id));
        if (ordenados.length > mensajesCountRef.current) {
          const ultimo = ordenados[ordenados.length - 1];
          const s = sesionRef.current;
          const vAhora = vistaRef.current;
          const esAgricultorViendo = vAhora === "agricultor" && s && ultimo.dni !== s;
          const esPuestoViendo = vAhora === "puesto" && ultimo.dni !== PUESTO_DNI;
          if (esAgricultorViendo || esPuestoViendo) {
            setNotiToast({ tipo: "mensaje", texto: `${ultimo.nombre} (Whatsapillo): ${ultimo.texto}` });
          }
          setMensajes(ordenados);
        }
        mensajesCountRef.current = ordenados.length;
      }
    }, 10000);
    return () => clearInterval(id);
  }, []);

  useEffect(() => {
    if (!notiToast) return;
    const id = setTimeout(() => setNotiToast(null), 4500);
    return () => clearTimeout(id);
  }, [notiToast]);

  const marcarPresencia = useCallback(async (dniActivo) => {
    setPresencia((prev) => {
      const nuevo = { ...prev, [dniActivo]: Date.now() };
      saveShared("jaralillo-presencia", nuevo);
      return nuevo;
    });
  }, []);

  const publicarAviso = useCallback(
    async (aviso) => {
      const nuevo = [{ ...aviso, id: Date.now(), ts: Date.now() }, ...avisos];
      setAvisos(nuevo);
      avisosCountRef.current = nuevo.length;
      await saveShared("jaralillo-avisos", nuevo);
    },
    [avisos]
  );

  const eliminarAviso = useCallback(
    async (id) => {
      const nuevo = avisos.filter((a) => a.id !== id);
      setAvisos(nuevo);
      avisosCountRef.current = nuevo.length;
      await saveShared("jaralillo-avisos", nuevo);
    },
    [avisos]
  );

  const enviarMensaje = useCallback(
    async (autorDni, autorNombre, texto, audioUrl, duracion) => {
      const mensajeNuevo = {
        id: Date.now(),
        dni: autorDni,
        nombre: autorNombre,
        texto: texto || "",
        audioUrl: audioUrl || null,
        duracion: duracion || null,
        ts: Date.now(),
      };
      const nuevo = [...mensajes, mensajeNuevo];
      setMensajes(nuevo);
      mensajesCountRef.current = nuevo.length;
      const ok = await guardarMensajeRemoto(mensajeNuevo);
      return { ok };
    },
    [mensajes]
  );

  const eliminarMensaje = useCallback(
    async (id) => {
      const nuevo = mensajes.filter((m) => m.id !== id);
      setMensajes(nuevo);
      mensajesCountRef.current = nuevo.length;
      await borrarMensajeRemoto(id);
    },
    [mensajes]
  );

  const registrarPerfil = useCallback(
    async (dniReg, pin) => {
      const nuevo = { ...perfiles, [dniReg]: { pin } };
      setPerfiles(nuevo);
      await saveShared("jaralillo-perfiles", nuevo);
    },
    [perfiles]
  );

  const reiniciarPin = useCallback(
    async (dniReset) => {
      const nuevo = { ...perfiles };
      delete nuevo[dniReset];
      setPerfiles(nuevo);
      await saveShared("jaralillo-perfiles", nuevo);
    },
    [perfiles]
  );

  const actualizarPerfilPersonal = useCallback(
    async (dniP, cambios) => {
      const nuevo = { ...perfiles, [dniP]: { ...(perfiles[dniP] || {}), ...cambios } };
      setPerfiles(nuevo);
      await saveShared("jaralillo-perfiles", nuevo);
    },
    [perfiles]
  );

  const actualizarCodigo = useCallback(async (nuevoCodigo) => {
    setCodigo(nuevoCodigo);
    await saveShared("jaralillo-codigo", nuevoCodigo);
  }, []);

  const actualizarPrecios = useCallback(async (nuevo) => {
    setPrecios(nuevo);
    await saveShared("jaralillo-precios", nuevo);
  }, []);

  const actualizarTiempoEspera = useCallback((minutos) => {
    ultimaEdicionRef.current = Date.now();
    setTiempoEspera(minutos);
    saveShared("jaralillo-espera", minutos);
  }, []);

  const reiniciarKgTotales = useCallback(async () => {
    ultimaEdicionRef.current = Date.now();
    setKgTotales(0);
    await saveShared("jaralillo-kg-totales", 0);
  }, []);

  const vaciarFacturas = useCallback(async () => {
    ultimaEdicionRef.current = Date.now();
    setFacturas([]);
    await saveShared("jaralillo-facturas-index", []);
    ultimaEdicionRef.current = Date.now();
  }, []);

  const empezarCampañaNueva = useCallback(async () => {
    ultimaEdicionRef.current = Date.now();
    setAlbaranes({});
    setPedidos([]);
    setAvisos([]);
    setKgTotales(0);
    setPendientes([]);
    avisosCountRef.current = 0;
    await Promise.all([
      saveShared("jaralillo-albaranes", {}),
      saveShared("jaralillo-pedidos", []),
      saveShared("jaralillo-avisos", []),
      saveShared("jaralillo-kg-totales", 0),
      saveShared("jaralillo-albaranes-pendientes", []),
    ]);
    ultimaEdicionRef.current = Date.now();
  }, []);

  const alternarRevisado = useCallback(
    async (id) => {
      ultimaEdicionRef.current = Date.now();
      const nuevo = pendientes.map((p) =>
        p.id === id ? { ...p, revisado: !p.revisado } : p
      );
      setPendientes(nuevo);
      await saveShared("jaralillo-albaranes-pendientes", nuevo);
    },
    [pendientes]
  );

  const marcarTodosPendientes = useCallback(
    async (revisado) => {
      ultimaEdicionRef.current = Date.now();
      const nuevo = pendientes.map((p) => ({ ...p, revisado }));
      setPendientes(nuevo);
      await saveShared("jaralillo-albaranes-pendientes", nuevo);
      ultimaEdicionRef.current = Date.now();
    },
    [pendientes]
  );

  const enviarPendientesAAgricultores = useCallback(async () => {
    ultimaEdicionRef.current = Date.now();
    const aEnviar = pendientes.filter((p) => p.revisado);
    const siguenPendientes = pendientes.filter((p) => !p.revisado);
    const nuevo = { ...albaranes };
    aEnviar.forEach((p) => {
      const lista = nuevo[p.dni] ? [...nuevo[p.dni]] : [];
      const yaExiste = lista.some((a) => a.n === p.n);
      if (!yaExiste) {
        lista.push({ n: p.n, fecha: p.fecha, kilos: p.kilos, precio: p.precio });
      }
      nuevo[p.dni] = lista;
    });
    setAlbaranes(nuevo);
    await saveShared("jaralillo-albaranes", nuevo);
    setPendientes(siguenPendientes);
    await saveShared("jaralillo-albaranes-pendientes", siguenPendientes);
    ultimaEdicionRef.current = Date.now();
  }, [pendientes, albaranes]);

  const subirFactura = useCallback(
    async (dni, nombreArchivo, dataUrl) => {
      ultimaEdicionRef.current = Date.now();
      const nueva = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        dni,
        nombreArchivo,
        dataUrl,
        fechaSubida: Date.now(),
        firmada: false,
        fechaFirmada: null,
        dataUrlFirmada: null,
        nombreArchivoFirmado: null,
      };
      const nuevo = [nueva, ...facturas];
      setFacturas(nuevo);
      const ok = await saveShared("jaralillo-facturas-index", nuevo);
      ultimaEdicionRef.current = Date.now();
      return { ok };
    },
    [facturas]
  );

  const subirFacturaFirmada = useCallback(
    async (id, nombreArchivoFirmado, dataUrl) => {
      ultimaEdicionRef.current = Date.now();
      const nuevo = facturas.map((f) =>
        f.id === id
          ? { ...f, firmada: true, fechaFirmada: Date.now(), nombreArchivoFirmado, dataUrlFirmada: dataUrl }
          : f
      );
      setFacturas(nuevo);
      const ok = await saveShared("jaralillo-facturas-index", nuevo);
      ultimaEdicionRef.current = Date.now();
      return { ok };
    },
    [facturas]
  );

  const sincronizarGuia = useCallback(
    async (filas) => {
      // Clientes nuevos que traiga la guía.
      ultimaEdicionRef.current = Date.now();
      const agricultoresNuevo = { ...agricultores };
      let clientesNuevos = 0;
      filas
        .filter((f) => f.estado === "nuevo")
        .forEach((f) => {
          agricultoresNuevo[f.dni] = { nombre: f.nombreFila, poblacion: f.poblacionFila || "" };
          clientesNuevos++;
        });
      if (clientesNuevos > 0) {
        setAgricultores(agricultoresNuevo);
        await saveShared("jaralillo-agricultores", agricultoresNuevo);
      }

      // Los albaranes van a la carpeta de revisión, no directos.
      const nuevosPendientes = [...pendientes];
      let añadidos = 0;
      filas
        .filter((f) => f.estado === "cliente" || f.estado === "nuevo")
        .forEach((f) => {
          const yaExiste = nuevosPendientes.some((p) => p.dni === f.dni && p.n === f.n);
          if (!yaExiste) {
            nuevosPendientes.push({
              id: Date.now() + Math.random(),
              dni: f.dni,
              nombreAgricultor: agricultoresNuevo[f.dni]?.nombre || f.nombreFila,
              n: f.n,
              fecha: new Date().toISOString().slice(0, 10),
              kilos: f.kilos,
              precio: f.precio,
              revisado: false,
            });
            añadidos++;
          }
        });
      setPendientes(nuevosPendientes);
      await saveShared("jaralillo-albaranes-pendientes", nuevosPendientes);

      // Los kg sin agricultor se suman al contador total.
      const kgSueltosHoy = filas
        .filter((f) => f.estado === "soloKilos")
        .reduce((s, f) => s + f.kilos, 0);
      let kgTotalNuevo = kgTotales;
      if (kgSueltosHoy > 0) {
        kgTotalNuevo = kgTotales + kgSueltosHoy;
        setKgTotales(kgTotalNuevo);
        await saveShared("jaralillo-kg-totales", kgTotalNuevo);
      }

      ultimaEdicionRef.current = Date.now();
      return { añadidos, clientesNuevos, kgSueltosHoy, kgTotalNuevo };
    },
    [agricultores, pendientes, kgTotales]
  );

  const crearPedido = useCallback(
    async (autorDni, autorNombre, tipo, cantidad) => {
      const nuevo = [
        {
          id: Date.now(),
          dni: autorDni,
          nombre: autorNombre,
          tipo,
          cantidad,
          estado: "pendiente",
          ts: Date.now(),
        },
        ...pedidos,
      ];
      setPedidos(nuevo);
      await saveShared("jaralillo-pedidos", nuevo);
    },
    [pedidos]
  );

  const cambiarEstadoPedido = useCallback(
    async (id) => {
      const orden = ["pendiente", "preparado", "entregado"];
      const nuevo = pedidos.map((p) => {
        if (p.id !== id) return p;
        const siguiente = orden[(orden.indexOf(p.estado) + 1) % orden.length];
        return { ...p, estado: siguiente };
      });
      setPedidos(nuevo);
      await saveShared("jaralillo-pedidos", nuevo);
    },
    [pedidos]
  );

  if (!ready) {
    return (
      <div style={styles.loadingScreen}>
        <Loader2 className="animate-spin" size={28} color="#C9982B" />
      </div>
    );
  }

  return (
    <div style={styles.app}>
      <GlobalStyle />
      {notiToast && (
        <div style={styles.toast} onClick={() => setNotiToast(null)}>
          <Bell size={15} style={{ flexShrink: 0, marginTop: 1 }} />
          <p style={styles.toastTexto}>{notiToast.texto}</p>
        </div>
      )}
      {vista === "portal" && (
        <Portal
          onAgricultor={() => setVista("login")}
          onPuesto={() => setVista("puestoLogin")}
        />
      )}
      {vista === "login" && (
        <LoginAgricultor
          agricultores={agricultores}
          perfiles={perfiles}
          codigo={codigo}
          onRegistrar={registrarPerfil}
          onVolver={() => setVista("portal")}
          onEntrar={(d) => {
            setSesion(d);
            setVista("agricultor");
          }}
        />
      )}
      {vista === "agricultor" && sesion && !agricultores[sesion] && (
        <div style={styles.loadingScreen}>
          <Loader2 className="animate-spin" size={26} color="#C9982B" />
        </div>
      )}
      {vista === "agricultor" && sesion && agricultores[sesion] && (
        <VistaAgricultor
          dni={sesion}
          datos={agricultores[sesion]}
          perfil={perfiles[sesion]}
          tiempoEspera={tiempoEspera}
          onActualizarPerfil={(cambios) => actualizarPerfilPersonal(sesion, cambios)}
          albaranes={albaranes[sesion] || []}
          avisos={avisos}
          mensajes={mensajes}
          perfiles={perfiles}
          onEnviarMensaje={(texto) =>
            enviarMensaje(sesion, perfiles[sesion]?.nombre || agricultores[sesion].nombre, texto)
          }
          onEnviarAudio={(audioUrl, duracion) =>
            enviarMensaje(sesion, perfiles[sesion]?.nombre || agricultores[sesion].nombre, "", audioUrl, duracion)
          }
          onPresencia={() => marcarPresencia(sesion)}
          precios={precios}
          pedidos={pedidos.filter((p) => p.dni === sesion)}
          onCrearPedido={(tipo, cantidad) =>
            crearPedido(
              sesion,
              perfiles[sesion]?.nombre || agricultores[sesion].nombre,
              tipo,
              cantidad
            )
          }
          facturas={facturas.filter((f) => f.dni === sesion)}
          onSubirFacturaFirmada={subirFacturaFirmada}
          onSalir={() => {
            setSesion(null);
            setVista("portal");
          }}
        />
      )}
      {vista === "puestoLogin" && (
        <LoginPuesto onVolver={() => setVista("portal")} onEntrar={() => setVista("puesto")} />
      )}
      {vista === "puesto" && (
        <VistaPuesto
          agricultores={agricultores}
          avisos={avisos}
          onPublicar={publicarAviso}
          onEliminarAviso={eliminarAviso}
          tiempoEspera={tiempoEspera}
          onActualizarTiempoEspera={actualizarTiempoEspera}
          kgTotales={kgTotales}
          onReiniciarKgTotales={reiniciarKgTotales}
          onEmpezarCampañaNueva={empezarCampañaNueva}
          pendientes={pendientes}
          onAlternarRevisado={alternarRevisado}
          onMarcarTodosPendientes={marcarTodosPendientes}
          onEnviarPendientes={enviarPendientesAAgricultores}
          onSincronizarGuia={sincronizarGuia}
          facturas={facturas}
          onSubirFactura={subirFactura}
          onVaciarFacturas={vaciarFacturas}
          codigo={codigo}
          onActualizarCodigo={actualizarCodigo}
          mensajes={mensajes}
          onEliminarMensaje={eliminarMensaje}
          onEnviarMensaje={(texto) => enviarMensaje(PUESTO_DNI, PUESTO_NOMBRE, texto)}
          onEnviarAudio={(audioUrl, duracion) => enviarMensaje(PUESTO_DNI, PUESTO_NOMBRE, "", audioUrl, duracion)}
          perfiles={perfiles}
          onReiniciarPin={reiniciarPin}
          presencia={presencia}
          precios={precios}
          onActualizarPrecios={actualizarPrecios}
          pedidos={pedidos}
          onCambiarEstadoPedido={cambiarEstadoPedido}
          onSalir={() => setVista("portal")}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pantalla: Portal de entrada
// ---------------------------------------------------------------------------
function Portal({ onAgricultor, onPuesto }) {
  return (
    <div style={styles.portalWrap}>
      <img src={LOGO_B64} alt="El Jaralillo" style={styles.logoPortal} />
      <p style={styles.eyebrow}>Campaña 2026 · 2027</p>

      <button style={{ ...styles.primaryBtn, marginTop: 8 }} onClick={onAgricultor}>
        <Leaf size={18} /> Soy agricultor <ChevronRight size={16} style={{ marginLeft: "auto" }} />
      </button>
      <button style={styles.ghostBtn} onClick={onPuesto}>
        Acceso del puesto
      </button>

      <div style={styles.portalFoot}>
        Demo funcional · los datos son de ejemplo
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pantalla: login agricultor (DNI + PIN propio; registro con código del puesto)
// ---------------------------------------------------------------------------
function LoginAgricultor({ agricultores, perfiles, codigo, onRegistrar, onVolver, onEntrar }) {
  const [paso, setPaso] = useState("dni"); // dni | pin | registro
  const [dni, setDni] = useState("");
  const [pin, setPin] = useState("");
  const [codigoIntro, setCodigoIntro] = useState("");
  const [nuevoPin, setNuevoPin] = useState("");
  const [nuevoPinRepetir, setNuevoPinRepetir] = useState("");
  const [error, setError] = useState("");

  const comprobarDni = () => {
    const limpio = dni.trim().toUpperCase();
    if (!agricultores[limpio]) {
      setError("Ese DNI no está registrado en el puesto. Prueba con 12345678A para la demo.");
      return;
    }
    setDni(limpio);
    setError("");
    setPaso(perfiles[limpio] ? "pin" : "registro");
  };

  const entrarConPin = () => {
    if (perfiles[dni]?.pin === pin.trim()) {
      setError("");
      onEntrar(dni);
    } else {
      setError("PIN incorrecto.");
    }
  };

  const registrarse = () => {
    if (codigoIntro.trim() !== codigo) {
      setError("Ese número no es el que ha dado hoy el puesto.");
      return;
    }
    if (!/^\d{4,6}$/.test(nuevoPin)) {
      setError("Elige un PIN de entre 4 y 6 números.");
      return;
    }
    if (nuevoPin !== nuevoPinRepetir) {
      setError("Los dos PIN no coinciden.");
      return;
    }
    setError("");
    onRegistrar(dni, nuevoPin);
    onEntrar(dni);
  };

  return (
    <div style={styles.centerCard}>
      <button
        style={styles.backLink}
        onClick={() => (paso === "dni" ? onVolver() : setPaso("dni"))}
      >
        ← Volver
      </button>
      <div style={styles.sealSmall}>
        <img src={LOGO_B64} alt="El Jaralillo" style={styles.logoSello} />
      </div>

      {paso === "dni" && (
        <>
          <h2 style={styles.cardTitle}>Acceso agricultor</h2>
          <p style={styles.cardHint}>Introduce tu DNI para continuar.</p>
          <input
            style={styles.input}
            placeholder="12345678A"
            value={dni}
            onChange={(e) => setDni(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && comprobarDni()}
            autoFocus
          />
          {error && <p style={styles.errorText}>{error}</p>}
          <button style={styles.primaryBtn} onClick={comprobarDni}>
            Continuar <ChevronRight size={16} style={{ marginLeft: "auto" }} />
          </button>
          <p style={styles.tinyHint}>
            DNIs de prueba: 12345678A · 23456789B · 34567890C · 45678901D
          </p>
        </>
      )}

      {paso === "pin" && (
        <>
          <h2 style={styles.cardTitle}>Tu PIN</h2>
          <p style={styles.cardHint}>{agricultores[dni]?.nombre} — introduce tu PIN.</p>
          <input
            style={styles.input}
            type="password"
            inputMode="numeric"
            placeholder="PIN"
            value={pin}
            onChange={(e) => setPin(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && entrarConPin()}
            autoFocus
          />
          {error && <p style={styles.errorText}>{error}</p>}
          <button style={styles.primaryBtn} onClick={entrarConPin}>
            Entrar <ChevronRight size={16} style={{ marginLeft: "auto" }} />
          </button>
          <button
            style={styles.linkBtn}
            onClick={() => {
              setError("");
              setCodigoIntro("");
              setNuevoPin("");
              setNuevoPinRepetir("");
              setPaso("registro");
            }}
          >
            ¿Olvidaste tu PIN?
          </button>
        </>
      )}

      {paso === "registro" && (
        <>
          <h2 style={styles.cardTitle}>
            {perfiles[dni] ? "Recuperar PIN" : "Primer acceso"}
          </h2>
          <p style={styles.cardHint}>
            {agricultores[dni]?.nombre} — pregunta en el puesto el número de dos cifras de
            hoy para {perfiles[dni] ? "poder elegir un PIN nuevo" : "poder registrarte"}.
          </p>
          <input
            style={styles.input}
            inputMode="numeric"
            placeholder="Número de dos cifras del puesto"
            value={codigoIntro}
            onChange={(e) => setCodigoIntro(e.target.value)}
          />
          <input
            style={styles.input}
            type="password"
            inputMode="numeric"
            placeholder={perfiles[dni] ? "Elige tu nuevo PIN (4-6 números)" : "Elige tu PIN (4-6 números)"}
            value={nuevoPin}
            onChange={(e) => setNuevoPin(e.target.value)}
          />
          <input
            style={styles.input}
            type="password"
            inputMode="numeric"
            placeholder="Repite tu PIN"
            value={nuevoPinRepetir}
            onChange={(e) => setNuevoPinRepetir(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && registrarse()}
          />
          {error && <p style={styles.errorText}>{error}</p>}
          <button style={styles.primaryBtn} onClick={registrarse}>
            {perfiles[dni] ? "Guardar nuevo PIN" : "Registrarme"}{" "}
            <ChevronRight size={16} style={{ marginLeft: "auto" }} />
          </button>
          {perfiles[dni] && (
            <button style={styles.linkBtn} onClick={() => { setError(""); setPaso("pin"); }}>
              Ya me acuerdo, quiero entrar con mi PIN
            </button>
          )}
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pantalla: login del puesto (PIN)
// ---------------------------------------------------------------------------
function LoginPuesto({ onVolver, onEntrar }) {
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");

  const entrar = () => {
    if (pin === PUESTO_PIN) {
      setError("");
      onEntrar();
    } else {
      setError(`PIN incorrecto. Para la demo usa ${PUESTO_PIN}.`);
    }
  };

  return (
    <div style={styles.centerCard}>
      <button style={styles.backLink} onClick={onVolver}>
        ← Volver
      </button>
      <div style={styles.sealSmall}>
        <img src={LOGO_B64} alt="El Jaralillo" style={styles.logoSello} />
      </div>
      <h2 style={styles.cardTitle}>Acceso del puesto</h2>
      <p style={styles.cardHint}>Solo para el equipo de El Jaralillo.</p>
      <input
        style={styles.input}
        type="password"
        placeholder="PIN"
        value={pin}
        onChange={(e) => setPin(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && entrar()}
        autoFocus
      />
      {error && <p style={styles.errorText}>{error}</p>}
      <button style={styles.primaryBtn} onClick={entrar}>
        Entrar <ChevronRight size={16} style={{ marginLeft: "auto" }} />
      </button>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pantalla: vista del agricultor
// ---------------------------------------------------------------------------
function VistaAgricultor({
  dni,
  datos,
  perfil,
  onActualizarPerfil,
  albaranes,
  avisos,
  mensajes,
  perfiles,
  onEnviarMensaje,
  onEnviarAudio,
  onPresencia,
  tiempoEspera,
  precios,
  pedidos,
  onCrearPedido,
  facturas,
  onSubirFacturaFirmada,
  onSalir,
}) {
  const [tab, setTab] = useState("albaranes"); // albaranes | facturas | tablon | perfil | chat | pedidos
  const nombreMostrar = perfil?.nombre || datos.nombre;

  useEffect(() => {
    onPresencia();
    const id = setInterval(onPresencia, 45000);
    return () => clearInterval(id);
  }, [onPresencia]);
  const totalKilos = albaranes.reduce((s, a) => s + a.kilos, 0);
  const totalEuros = albaranes.reduce((s, a) => s + a.kilos * a.precio, 0);
  const paraMi = avisos
    .filter(
      (a) => a.destinatarios === "todos" || (Array.isArray(a.destinatarios) && a.destinatarios.includes(dni))
    )
    .sort((a, b) => (b.ts || b.id) - (a.ts || a.id));

  return (
    <div style={styles.screen}>
      <TopBar
        label={nombreMostrar}
        sub={`DNI ${dni} · ${datos.poblacion}`}
        onSalir={onSalir}
        avatarFoto={perfil?.foto}
        avatarNombre={nombreMostrar}
      />

      <div style={styles.tabRow}>
        <button
          style={{ ...styles.tabBtn, ...(tab === "albaranes" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("albaranes")}
        >
          <Package size={14} /> Mis albaranes
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "facturas" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("facturas")}
        >
          <FileText size={14} /> Facturas{facturas.some((f) => !f.firmada) ? ` (${facturas.filter((f) => !f.firmada).length})` : ""}
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "tablon" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("tablon")}
        >
          <Pin size={14} /> Tablón
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "perfil" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("perfil")}
        >
          <User size={14} /> Mi perfil
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "chat" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("chat")}
        >
          <MessageCircle size={14} /> Whatsapillo
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "pedidos" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("pedidos")}
        >
          <Droplet size={14} /> Pedir aceite
        </button>
      </div>

      {tab === "albaranes" && (
        <>
          <EsperaBar minutos={tiempoEspera} />

          <div style={{ ...styles.statsRow, marginTop: 20 }}>
            <div style={styles.statCard}>
              <p style={styles.statLabel}>Kilos entregados</p>
              <p style={styles.statValue}>{totalKilos.toLocaleString("es-ES")} kg</p>
            </div>
            <div style={styles.statCard}>
              <p style={styles.statLabel}>Total campaña</p>
              <p style={styles.statValue}>{money(totalEuros)}</p>
            </div>
          </div>

          <p style={styles.sectionLabel}>Tus albaranes</p>
          {albaranes.length === 0 ? (
            <p style={styles.emptyText}>Todavía no tienes albaranes esta campaña.</p>
          ) : (
            <div style={styles.albaranList}>
              {albaranes.map((a) => (
                <div style={styles.albaranRow} key={a.n}>
                  <div>
                    <p style={styles.albaranN}>Albarán nº {a.n}</p>
                    <p style={styles.albaranFecha}>{fecha(a.fecha)}</p>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <p style={styles.albaranKilos}>{a.kilos.toLocaleString("es-ES")} kg</p>
                    <p style={styles.albaranPrecio}>{a.precio.toFixed(2)} €/kg</p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}

      {tab === "pedidos" && (
        <PanelPedidosAgricultor precios={precios} pedidos={pedidos} onCrear={onCrearPedido} />
      )}

      {tab === "facturas" && (
        <PanelFacturasAgricultor facturas={facturas} onSubirFirmada={onSubirFacturaFirmada} />
      )}

      {tab === "tablon" && <PanelTablon avisos={paraMi} />}

      {tab === "perfil" && (
        <PanelPerfilAgricultor
          datosOficiales={datos}
          perfil={perfil}
          onGuardar={onActualizarPerfil}
        />
      )}

      {tab === "chat" && (
        <ChatAgricultores dni={dni} mensajes={mensajes} perfiles={perfiles} onEnviar={onEnviarMensaje} onEnviarAudio={onEnviarAudio} />
      )}
    </div>
  );
}

function PanelFacturasAgricultor({ facturas, onSubirFirmada }) {
  const ordenadas = [...facturas].sort((a, b) => b.fechaSubida - a.fechaSubida);

  return (
    <div style={{ marginTop: 20 }}>
      {ordenadas.length === 0 ? (
        <p style={styles.emptyText}>Todavía no tienes facturas subidas.</p>
      ) : (
        <div style={styles.albaranList}>
          {ordenadas.map((f) => (
            <FacturaCardAgricultor key={f.id} factura={f} onSubirFirmada={onSubirFirmada} />
          ))}
        </div>
      )}
    </div>
  );
}

function FacturaCardAgricultor({ factura: f, onSubirFirmada }) {
  const [subiendo, setSubiendo] = useState(false);
  const [error, setError] = useState("");

  const descargar = (tipo) => {
    setError("");
    const url = tipo === "firmada" ? f.dataUrlFirmada : f.dataUrl;
    if (url) {
      descargarDataUrl(url, tipo === "firmada" ? `firmada-${f.nombreArchivo}` : f.nombreArchivo);
    } else {
      setError("Este archivo no está disponible.");
    }
  };

  const subirFirmada = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    if (file.type !== "application/pdf") {
      setError("Solo se admiten archivos PDF.");
      e.target.value = "";
      return;
    }
    if (file.size > 500 * 1024) {
      setError("El PDF pesa demasiado (máx. 500 KB por ahora).");
      e.target.value = "";
      return;
    }
    setSubiendo(true);
    const lector = new FileReader();
    lector.onload = async (ev) => {
      const resultado = await onSubirFirmada(f.id, file.name, ev.target.result);
      setSubiendo(false);
      e.target.value = "";
      if (resultado && resultado.ok === false) {
        setError("No se ha podido guardar el archivo. Prueba con un PDF más ligero.");
      }
    };
    lector.readAsDataURL(file);
  };

  return (
    <div style={styles.facturaCard}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ display: "flex", gap: 10 }}>
          <FileText size={22} color="#8A5A1F" style={{ flexShrink: 0, marginTop: 2 }} />
          <div>
            <p style={styles.albaranN}>{f.nombreArchivo}</p>
            <p style={styles.albaranFecha}>
              {new Date(f.fechaSubida).toLocaleDateString("es-ES", { day: "2-digit", month: "short", year: "numeric" })}
            </p>
          </div>
        </div>
        <span
          style={{
            ...styles.estadoBadge,
            ...(f.firmada ? styles.estadoBadgeOk : styles.estadoBadgeAviso),
          }}
        >
          {f.firmada ? "Firmada" : "Sin firmar"}
        </span>
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
        <button style={{ ...styles.ghostBtn, flex: 1 }} onClick={() => descargar("original")}>
          <Download size={14} /> Descargar
        </button>
        {f.firmada && (
          <button style={{ ...styles.ghostBtn, flex: 1 }} onClick={() => descargar("firmada")}>
            <CheckCircle2 size={14} color="#4E7D3A" /> Ver firmada
          </button>
        )}
      </div>

      {!f.firmada && (
        <label style={{ ...styles.subirBtn, marginTop: 8, opacity: subiendo ? 0.6 : 1 }}>
          {subiendo ? <Loader2 size={15} className="animate-spin" /> : <Upload size={15} />}
          Subir ya firmada (PDF)
          <input type="file" accept="application/pdf" onChange={subirFirmada} style={{ display: "none" }} disabled={subiendo} />
        </label>
      )}

      {f.firmada && (
        <p style={styles.tinyHintLeft}>
          Firmada el{" "}
          {new Date(f.fechaFirmada).toLocaleDateString("es-ES", { day: "2-digit", month: "short", year: "numeric" })}
        </p>
      )}

      {error && (
        <p style={styles.errorText}>
          <AlertTriangle size={13} style={{ flexShrink: 0, marginTop: 1 }} /> {error}
        </p>
      )}
    </div>
  );
}

function PanelTablon({ avisos }) {
  return (
    <div style={styles.tablonFondo}>
      {avisos.length === 0 ? (
        <p style={styles.emptyText}>No hay avisos del puesto por ahora.</p>
      ) : (
        <div style={styles.tablonGrid}>
          {avisos.map((a) => (
            <div key={a.id} style={styles.notaTablon}>
              <AvisoCard aviso={a} />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function PanelPerfilAgricultor({ datosOficiales, perfil, onGuardar }) {
  const [nombre, setNombre] = useState(perfil?.nombre || datosOficiales.nombre);
  const [foto, setFoto] = useState(perfil?.foto || null);
  const [cargando, setCargando] = useState(false);
  const [guardado, setGuardado] = useState(false);

  const elegirFoto = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setCargando(true);
    try {
      const dataUrl = await redimensionarFoto(file);
      setFoto(dataUrl);
    } catch {
      // si falla la lectura de la imagen, simplemente no cambiamos la foto
    }
    setCargando(false);
  };

  const guardar = () => {
    onGuardar({ nombre: nombre.trim() || datosOficiales.nombre, foto });
    setGuardado(true);
    setTimeout(() => setGuardado(false), 1800);
  };

  return (
    <div style={{ marginTop: 20 }}>
      <div style={styles.perfilAvatarWrap}>
        <Avatar foto={foto} nombre={nombre || datosOficiales.nombre} size={88} />
        <label style={styles.perfilCambiarFotoBtn}>
          {cargando ? <Loader2 size={14} className="animate-spin" /> : "Cambiar foto"}
          <input
            type="file"
            accept="image/*"
            onChange={elegirFoto}
            style={{ display: "none" }}
          />
        </label>
        {foto && (
          <button style={styles.linkBtnMini} onClick={() => setFoto(null)}>
            Quitar foto
          </button>
        )}
      </div>

      <p style={{ ...styles.sectionLabel, marginTop: 22 }}>Nombre para mostrar</p>
      <input
        style={styles.input}
        value={nombre}
        onChange={(e) => setNombre(e.target.value)}
        placeholder={datosOficiales.nombre}
      />
      <p style={styles.tinyHintLeft}>
        Este es el nombre que verán los demás en Whatsapillo y en tus pedidos. Tu nombre y DNI
        oficiales del puesto ({datosOficiales.nombre}) no cambian.
      </p>

      <button style={{ ...styles.primaryBtn, marginTop: 6 }} onClick={guardar}>
        Guardar cambios
      </button>
      {guardado && <p style={styles.tinyHint}>Perfil actualizado.</p>}
    </div>
  );
}

function PanelPedidosAgricultor({ precios, pedidos, onCrear }) {
  const [tipo, setTipo] = useState("garrafa");
  const [cantidad, setCantidad] = useState(1);
  const [enviando, setEnviando] = useState(false);
  const [hecho, setHecho] = useState(false);

  const precioUnidad = tipo === "garrafa" ? precios.garrafa : precios.caja;
  const totalPedido = precioUnidad * cantidad;

  const pedir = async () => {
    if (cantidad < 1) return;
    setEnviando(true);
    await onCrear(tipo, cantidad);
    setEnviando(false);
    setCantidad(1);
    setHecho(true);
    setTimeout(() => setHecho(false), 2000);
  };

  const ordenados = [...pedidos].sort((a, b) => (b.ts || b.id) - (a.ts || a.id));

  return (
    <div style={{ marginTop: 20 }}>
      <p style={styles.sectionLabel}>Precio actual</p>
      <div style={styles.statsRow}>
        <div style={styles.statCard}>
          <p style={styles.statLabel}>Garrafa</p>
          <p style={styles.statValue}>{money(precios.garrafa)}</p>
        </div>
        <div style={styles.statCard}>
          <p style={styles.statLabel}>Caja</p>
          <p style={styles.statValue}>{money(precios.caja)}</p>
        </div>
      </div>

      <p style={styles.sectionLabel}>Hacer un pedido</p>
      <div style={styles.tipoRow}>
        <button
          style={{ ...styles.tipoBtn, ...(tipo === "garrafa" ? styles.tipoBtnActivo : {}) }}
          onClick={() => setTipo("garrafa")}
        >
          Garrafa
        </button>
        <button
          style={{ ...styles.tipoBtn, ...(tipo === "caja" ? styles.tipoBtnActivo : {}) }}
          onClick={() => setTipo("caja")}
        >
          Caja
        </button>
      </div>

      <div style={styles.cantidadRow}>
        <button
          style={styles.cantidadBtn}
          onClick={() => setCantidad((c) => Math.max(1, c - 1))}
        >
          −
        </button>
        <span style={styles.cantidadValor}>{cantidad}</span>
        <button style={styles.cantidadBtn} onClick={() => setCantidad((c) => c + 1)}>
          +
        </button>
      </div>

      <button style={{ ...styles.primaryBtn, marginTop: 4 }} onClick={pedir} disabled={enviando}>
        {enviando ? <Loader2 size={16} className="animate-spin" /> : <Send size={16} />}
        Pedir por {money(totalPedido)}
      </button>
      {hecho && <p style={styles.tinyHint}>Pedido enviado al puesto.</p>}

      <p style={{ ...styles.sectionLabel, marginTop: 26 }}>Tus pedidos</p>
      {ordenados.length === 0 ? (
        <p style={styles.emptyText}>Todavía no has pedido nada.</p>
      ) : (
        <div style={styles.albaranList}>
          {ordenados.map((p) => (
            <div style={styles.albaranRow} key={p.id}>
              <div>
                <p style={styles.albaranN}>
                  {p.cantidad} {p.tipo === "garrafa" ? "garrafa" : "caja"}
                  {p.cantidad === 1 ? "" : "s"}
                </p>
                <p style={styles.albaranFecha}>
                  {new Date(p.ts || p.id).toLocaleString("es-ES", {
                    day: "2-digit",
                    month: "short",
                    hour: "2-digit",
                    minute: "2-digit",
                  })}
                </p>
              </div>
              <span
                style={{
                  ...styles.estadoBadge,
                  ...(p.estado === "entregado"
                    ? styles.estadoBadgeOk
                    : p.estado === "preparado"
                    ? styles.estadoBadgePreparado
                    : styles.estadoBadgeAviso),
                }}
              >
                {p.estado === "pendiente" && "Pendiente"}
                {p.estado === "preparado" && "Preparado"}
                {p.estado === "entregado" && "Entregado"}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function HiloWhatsapillo({ mensajes, perfiles, propioDni, onEliminar }) {
  const ordenados = [...mensajes].sort((a, b) => (a.ts || a.id) - (b.ts || b.id));

  if (ordenados.length === 0) {
    return <p style={styles.emptyText}>Todavía no hay mensajes. ¡Escribe el primero!</p>;
  }

  return ordenados.map((m) => {
    const esPropio = m.dni === propioDni;
    const esPuesto = m.dni === PUESTO_DNI;
    const nombreMostrar = esPuesto ? PUESTO_NOMBRE : perfiles?.[m.dni]?.nombre || m.nombre;
    const hora = new Date(m.ts || m.id).toLocaleString("es-ES", {
      day: "2-digit",
      month: "short",
      hour: "2-digit",
      minute: "2-digit",
    });
    return (
      <div
        key={m.id}
        style={{
          display: "flex",
          alignItems: "flex-end",
          gap: 8,
          alignSelf: esPuesto ? "center" : esPropio ? "flex-end" : "flex-start",
          maxWidth: esPuesto ? "88%" : "82%",
        }}
      >
        {!esPropio && !esPuesto && (
          <Avatar foto={perfiles?.[m.dni]?.foto} nombre={nombreMostrar} size={26} />
        )}
        <div
          style={{
            ...styles.chatBurbuja,
            maxWidth: "100%",
            position: "relative",
            ...(esPuesto
              ? styles.chatBurbujaPuesto
              : esPropio
              ? styles.chatBurbujaMia
              : styles.chatBurbujaOtro),
          }}
        >
          {!esPropio && (
            <p style={esPuesto ? styles.chatAutorPuesto : styles.chatAutor}>
              {esPuesto ? `🫒 ${PUESTO_NOMBRE}` : nombreMostrar.split(" ")[0]}
            </p>
          )}
          {m.audioUrl ? (
            <div style={styles.audioMsgRow}>
              <Mic size={14} style={{ flexShrink: 0 }} />
              <audio controls src={m.audioUrl} style={styles.audioPlayer} />
              {m.duracion && <span style={styles.audioDuracion}>{m.duracion}s</span>}
            </div>
          ) : (
            <p style={styles.chatTexto}>{m.texto}</p>
          )}
          <p style={styles.chatHora}>{hora}</p>
          {onEliminar && (
            <button style={styles.chatBorrarMini} onClick={() => onEliminar(m.id)}>
              <X size={12} />
            </button>
          )}
        </div>
      </div>
    );
  });
}

function ChatAgricultores({ dni, mensajes, perfiles, onEnviar, onEnviarAudio }) {
  const [texto, setTexto] = useState("");

  const enviar = () => {
    if (!texto.trim()) return;
    onEnviar(texto.trim());
    setTexto("");
  };

  return (
    <div>
      <p style={{ ...styles.sectionLabel, marginTop: 20 }}>
        Whatsapillo — mensajes entre agricultores del puesto
      </p>
      <div style={styles.chatBox}>
        <HiloWhatsapillo mensajes={mensajes} perfiles={perfiles} propioDni={dni} />
      </div>
      <div style={styles.chatInputRow}>
        <input
          style={{ ...styles.input, marginBottom: 0, flex: 1 }}
          placeholder="Escribe un mensaje..."
          value={texto}
          onChange={(e) => setTexto(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && enviar()}
        />
        <BotonGrabar onGrabado={onEnviarAudio} />
        <button style={styles.chatEnviarBtn} onClick={enviar}>
          <Send size={17} />
        </button>
      </div>
    </div>
  );
}

function BotonGrabar({ onGrabado }) {
  const [grabando, setGrabando] = useState(false);
  const [segundos, setSegundos] = useState(0);
  const [error, setError] = useState("");
  const mediaRecorderRef = useRef(null);
  const streamRef = useRef(null);
  const chunksRef = useRef([]);
  const segundosRef = useRef(0);
  const intervaloRef = useRef(null);
  const MAX_SEG = 45;

  // Muy importante: si el agricultor/puesto cambia de pantalla con el
  // micrófono todavía grabando, hay que soltarlo aquí — si no, el
  // navegador lo deja "ocupado" y la siguiente grabación (en cualquier
  // zona) deja de funcionar.
  useEffect(() => {
    return () => {
      if (intervaloRef.current) clearInterval(intervaloRef.current);
      if (mediaRecorderRef.current && mediaRecorderRef.current.state !== "inactive") {
        try {
          mediaRecorderRef.current.stop();
        } catch {
          // ya estaba parado, no pasa nada
        }
      }
      if (streamRef.current) {
        streamRef.current.getTracks().forEach((t) => t.stop());
      }
    };
  }, []);

  const detener = () => {
    if (intervaloRef.current) clearInterval(intervaloRef.current);
    setGrabando(false);
    if (mediaRecorderRef.current && mediaRecorderRef.current.state !== "inactive") {
      mediaRecorderRef.current.stop();
    }
  };

  const empezar = async () => {
    setError("");
    if (!navigator.mediaDevices?.getUserMedia) {
      setError("Este navegador o esta ventana no permite grabar audio.");
      return;
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      streamRef.current = stream;
      chunksRef.current = [];
      segundosRef.current = 0;
      const mr = new MediaRecorder(stream);
      mediaRecorderRef.current = mr;
      mr.ondataavailable = (e) => {
        if (e.data.size > 0) chunksRef.current.push(e.data);
      };
      mr.onstop = () => {
        stream.getTracks().forEach((t) => t.stop());
        streamRef.current = null;
        if (chunksRef.current.length === 0) {
          setError("No se ha grabado nada — mantén pulsado un poco más antes de parar.");
          return;
        }
        const blob = new Blob(chunksRef.current, { type: mr.mimeType || "audio/webm" });
        const lector = new FileReader();
        lector.onload = async () => {
          const resultado = await onGrabado(lector.result, segundosRef.current);
          if (resultado && resultado.ok === false) {
            setError("No se ha podido enviar el audio. Vuelve a intentarlo.");
          }
        };
        lector.readAsDataURL(blob);
      };
      mr.start();
      setGrabando(true);
      setSegundos(0);
      intervaloRef.current = setInterval(() => {
        segundosRef.current += 1;
        setSegundos(segundosRef.current);
        if (segundosRef.current >= MAX_SEG) detener();
      }, 1000);
    } catch (err) {
      setError(
        err?.name === "NotAllowedError"
          ? "Permiso de micrófono denegado o bloqueado en esta vista."
          : "No se puede acceder al micrófono aquí."
      );
    }
  };

  return (
    <div style={{ position: "relative" }}>
      {grabando ? (
        <button style={styles.grabandoBtn} onClick={detener} title="Parar y enviar">
          <span style={styles.puntoGrabando} /> {segundos}s
        </button>
      ) : (
        <button style={styles.micBtn} onClick={empezar} title="Grabar audio">
          <Mic size={17} />
        </button>
      )}
      {error && <p style={styles.audioErrorTooltip}>{error}</p>}
    </div>
  );
}

function EsperaBar({ minutos }) {
  const TOPE = 150; // escala visual: hasta 150 min, luego se corta y se muestra el número real
  const pct = Math.min(minutos, TOPE) / TOPE * 100;
  const zona = minutos <= 30 ? "verde" : minutos <= 90 ? "naranja" : "rojo";
  const color = { verde: "#4E7D3A", naranja: "#C9982B", rojo: "#B4552F" }[zona];
  const fondo = { verde: "#E7EAD9", naranja: "#F1E4C4", rojo: "#F4E1D8" }[zona];
  const etiqueta = { verde: "Espera corta", naranja: "Espera media", rojo: "Espera larga" }[zona];

  return (
    <div style={{ ...styles.esperaCard, background: fondo }}>
      <div style={styles.esperaTop}>
        <Clock size={15} color={color} />
        <p style={styles.esperaTitulo}>Tiempo de espera estimado en el puesto</p>
      </div>
      <p style={{ ...styles.esperaMinutos, color }}>
        {minutos} min <span style={styles.esperaEtiqueta}>· {etiqueta}</span>
      </p>
      <div style={styles.esperaTrack}>
        <div style={{ ...styles.esperaFill, width: `${pct}%`, background: color }} />
        <div style={{ ...styles.esperaMarca, left: "20%" }} />
        <div style={{ ...styles.esperaMarca, left: "60%" }} />
      </div>
    </div>
  );
}

function AvisoCard({ aviso }) {
  const cfg = {
    precio: { icon: Coins, bg: "#F1E4C4" },
    espera: { icon: Clock, bg: "#EFE0D6" },
    general: { icon: Bell, bg: "#E7EAD9" },
  }[aviso.tipo] || { icon: Bell, bg: "#E7EAD9" };
  const Icon = cfg.icon;
  const horaTexto = aviso.ts
    ? new Date(aviso.ts).toLocaleString("es-ES", {
        day: "2-digit",
        month: "short",
        hour: "2-digit",
        minute: "2-digit",
      })
    : "";
  return (
    <div style={{ ...styles.avisoCard, background: cfg.bg }}>
      <Icon size={16} color="#2E3A1F" style={{ flexShrink: 0, marginTop: 2 }} />
      <div>
        <p style={styles.avisoText}>{aviso.mensaje}</p>
        {horaTexto && <p style={styles.avisoHora}>{horaTexto}</p>}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pantalla: vista del puesto
// ---------------------------------------------------------------------------
function VistaPuesto({
  agricultores,
  avisos,
  onPublicar,
  onEliminarAviso,
  tiempoEspera,
  onActualizarTiempoEspera,
  kgTotales,
  onReiniciarKgTotales,
  onEmpezarCampañaNueva,
  pendientes,
  onAlternarRevisado,
  onMarcarTodosPendientes,
  onEnviarPendientes,
  onSincronizarGuia,
  facturas,
  onSubirFactura,
  onVaciarFacturas,
  codigo,
  onActualizarCodigo,
  mensajes,
  onEliminarMensaje,
  onEnviarMensaje,
  onEnviarAudio,
  perfiles,
  onReiniciarPin,
  presencia,
  precios,
  onActualizarPrecios,
  pedidos,
  onCambiarEstadoPedido,
  onSalir,
}) {
  const [tab, setTab] = useState("avisos"); // avisos | codigo | chat | agricultores | pedidos | revisar
  const [tipo, setTipo] = useState("precio");
  const [mensaje, setMensaje] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [destino, setDestino] = useState("todos"); // todos | seleccionados
  const [seleccionados, setSeleccionados] = useState([]);

  const tipos = [
    { id: "precio", label: "Cambio de precio", icon: Coins },
    { id: "espera", label: "Tiempo de espera", icon: Clock },
    { id: "general", label: "Aviso general", icon: Bell },
  ];

  const agricultoresLista = Object.entries(agricultores); // [dni, datos]

  const toggleSeleccionado = (dni) => {
    setSeleccionados((prev) =>
      prev.includes(dni) ? prev.filter((d) => d !== dni) : [...prev, dni]
    );
  };

  const puedeEnviar =
    mensaje.trim() && (destino === "todos" || seleccionados.length > 0);

  const enviar = async () => {
    if (!puedeEnviar) return;
    setEnviando(true);
    await onPublicar({
      tipo,
      mensaje: mensaje.trim(),
      destinatarios: destino === "todos" ? "todos" : seleccionados,
    });
    setMensaje("");
    setSeleccionados([]);
    setDestino("todos");
    setEnviando(false);
  };

  const textoBoton =
    destino === "todos"
      ? "Publicar a todos los clientes"
      : `Publicar a ${seleccionados.length} agricultor${seleccionados.length === 1 ? "" : "es"}`;

  return (
    <div style={styles.screen}>
      <TopBar label="Panel del puesto" sub="El Jaralillo" onSalir={onSalir} />

      <div style={styles.tabRow}>
        <button
          style={{ ...styles.tabBtn, ...(tab === "avisos" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("avisos")}
        >
          <Bell size={14} /> Avisos
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "codigo" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("codigo")}
        >
          <KeyRound size={14} /> Código de registro
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "chat" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("chat")}
        >
          <MessageCircle size={14} /> Whatsapillo
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "agricultores" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("agricultores")}
        >
          <Users size={14} /> Agricultores
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "pedidos" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("pedidos")}
        >
          <Droplet size={14} /> Pedidos
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "revisar" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("revisar")}
        >
          <ClipboardList size={14} /> Por revisar{pendientes.length > 0 ? ` (${pendientes.length})` : ""}
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "sincronizar" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("sincronizar")}
        >
          <RefreshCw size={14} /> Sincronizar guía
        </button>
        <button
          style={{ ...styles.tabBtn, ...(tab === "facturas" ? styles.tabBtnActivo : {}) }}
          onClick={() => setTab("facturas")}
        >
          <FileText size={14} /> Facturas
        </button>
      </div>

      {tab === "codigo" && <PanelCodigo codigo={codigo} onActualizar={onActualizarCodigo} />}

      {tab === "facturas" && (
        <PanelFacturasPuesto agricultores={agricultores} facturas={facturas} onSubir={onSubirFactura} onVaciar={onVaciarFacturas} />
      )}

      {tab === "chat" && (
        <PanelChatPuesto mensajes={mensajes} perfiles={perfiles} onEliminar={onEliminarMensaje} onEnviar={onEnviarMensaje} onEnviarAudio={onEnviarAudio} />
      )}

      {tab === "agricultores" && (
        <PanelAgricultores agricultores={agricultores} perfiles={perfiles} presencia={presencia} onReiniciarPin={onReiniciarPin} />
      )}

      {tab === "revisar" && (
        <PanelRevisarAlbaranes
          kgTotales={kgTotales}
          onReiniciarKgTotales={onReiniciarKgTotales}
          pendientes={pendientes}
          onAlternar={onAlternarRevisado}
          onMarcarTodos={onMarcarTodosPendientes}
          onEnviar={onEnviarPendientes}
          onEmpezarCampañaNueva={onEmpezarCampañaNueva}
        />
      )}

      {tab === "sincronizar" && (
        <PanelSincronizarGuia agricultores={agricultores} onSincronizar={onSincronizarGuia} />
      )}

      {tab === "pedidos" && (
        <PanelPedidosPuesto
          precios={precios}
          onActualizarPrecios={onActualizarPrecios}
          pedidos={pedidos}
          onCambiarEstado={onCambiarEstadoPedido}
        />
      )}

      {tab === "avisos" && (
        <>
      <PanelTiempoEsperaPuesto tiempoEspera={tiempoEspera} onActualizar={onActualizarTiempoEspera} />

      <p style={{ ...styles.sectionLabel, marginTop: 24 }}>Tipo de aviso</p>
      <div style={styles.tipoRow}>
        {tipos.map((t) => {
          const Icon = t.icon;
          const activo = tipo === t.id;
          return (
            <button
              key={t.id}
              style={{ ...styles.tipoBtn, ...(activo ? styles.tipoBtnActivo : {}) }}
              onClick={() => setTipo(t.id)}
            >
              <Icon size={14} /> {t.label}
            </button>
          );
        })}
      </div>

      <p style={{ ...styles.sectionLabel, marginTop: 18 }}>¿Para quién es?</p>
      <div style={styles.tipoRow}>
        <button
          style={{ ...styles.tipoBtn, ...(destino === "todos" ? styles.tipoBtnActivo : {}) }}
          onClick={() => setDestino("todos")}
        >
          Todos los agricultores
        </button>
        <button
          style={{ ...styles.tipoBtn, ...(destino === "seleccionados" ? styles.tipoBtnActivo : {}) }}
          onClick={() => setDestino("seleccionados")}
        >
          Elegir agricultores
        </button>
      </div>

      {destino === "seleccionados" && (
        <div style={styles.agricultorList}>
          {agricultoresLista.map(([dniAg, datos]) => {
            const marcado = seleccionados.includes(dniAg);
            return (
              <label
                key={dniAg}
                style={{
                  ...styles.agricultorRow,
                  ...(marcado ? styles.agricultorRowActivo : {}),
                }}
              >
                <input
                  type="checkbox"
                  checked={marcado}
                  onChange={() => toggleSeleccionado(dniAg)}
                  style={styles.checkbox}
                />
                <div>
                  <p style={styles.agricultorNombre}>{datos.nombre}</p>
                  <p style={styles.agricultorDni}>{dniAg} · {datos.poblacion}</p>
                </div>
              </label>
            );
          })}
        </div>
      )}

      <textarea
        style={{ ...styles.textarea, marginTop: 16 }}
        placeholder="Ej: Hoy el precio sube a 0,64 €/kg a partir de las 16:00"
        value={mensaje}
        onChange={(e) => setMensaje(e.target.value)}
      />
      <button style={styles.primaryBtn} onClick={enviar} disabled={!puedeEnviar || enviando}>
        {enviando ? <Loader2 size={16} className="animate-spin" /> : <Send size={16} />}
        {textoBoton}
      </button>

      <p style={{ ...styles.sectionLabel, marginTop: 28 }}>
        Últimos avisos enviados
      </p>
      {avisos.length === 0 ? (
        <p style={styles.emptyText}>Aún no has mandado ningún aviso.</p>
      ) : (
        <div style={styles.albaranList}>
          {avisos.map((a) => (
            <div key={a.id} style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <div style={{ position: "relative" }}>
                <AvisoCard aviso={a} />
                <button style={styles.borrarAvisoBtn} onClick={() => onEliminarAviso(a.id)}>
                  <X size={13} />
                </button>
              </div>
              <p style={styles.destinoTag}>
                Para: {a.destinatarios === "todos"
                  ? "todos los agricultores"
                  : (a.destinatarios || [])
                      .map((d) => agricultores[d]?.nombre.split(" ")[0])
                      .join(", ")}
              </p>
            </div>
          ))}
        </div>
      )}
      </>
      )}
    </div>
  );
}

function PanelAgricultores({ agricultores, perfiles, presencia, onReiniciarPin }) {
  const CONECTADO_MS = 45000; // se considera "en línea" si avisó en los últimos 45s
  const ahora = Date.now();
  const [confirmando, setConfirmando] = useState(null);

  const ultimaVezTexto = (ts) => {
    if (!ts) return "Nunca ha entrado";
    const minutos = Math.floor((ahora - ts) / 60000);
    if (minutos < 1) return "Justo ahora";
    if (minutos < 60) return `Hace ${minutos} min`;
    const horas = Math.floor(minutos / 60);
    if (horas < 24) return `Hace ${horas} h`;
    return `Hace ${Math.floor(horas / 24)} d`;
  };

  return (
    <div style={{ marginTop: 20 }}>
      <p style={styles.cardHint}>
        Estado de cada agricultor: si ya se ha registrado con su PIN y si está conectado en
        la app ahora mismo. Si alguien se queda sin PIN y no puede recuperarlo con el código,
        puedes reiniciárselo aquí — tendrá que registrarse de nuevo con el código de dos cifras.
      </p>
      <div style={styles.albaranList}>
        {Object.entries(agricultores).map(([dniAg, datos]) => {
          const registrado = !!perfiles[dniAg];
          const ultimaActividad = presencia[dniAg];
          const conectado = ultimaActividad && ahora - ultimaActividad < CONECTADO_MS;
          return (
            <div key={dniAg} style={styles.albaranRow}>
              <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                <Avatar foto={perfiles[dniAg]?.foto} nombre={perfiles[dniAg]?.nombre || datos.nombre} size={36} />
                <div>
                  <p style={styles.albaranN}>
                    {datos.nombre}
                    {perfiles[dniAg]?.nombre && perfiles[dniAg].nombre !== datos.nombre
                      ? ` (${perfiles[dniAg].nombre})`
                      : ""}
                  </p>
                  <p style={styles.albaranFecha}>{dniAg} · {datos.poblacion}</p>
                </div>
              </div>
              <div style={{ textAlign: "right" }}>
                <span
                  style={{
                    ...styles.estadoBadge,
                    ...(registrado ? styles.estadoBadgeOk : styles.estadoBadgeAviso),
                  }}
                >
                  {registrado ? "Registrado" : "Sin registrar"}
                </span>
                <p style={styles.conexionTexto}>
                  <span
                    style={{
                      ...styles.puntoEstado,
                      background: conectado ? "#4E7D3A" : "#C9BFA5",
                    }}
                  />
                  {conectado ? "En línea ahora" : ultimaVezTexto(ultimaActividad)}
                </p>
                {registrado && (
                  confirmando === dniAg ? (
                    <div style={styles.confirmarRow}>
                      <button
                        style={styles.confirmarBtnSi}
                        onClick={() => {
                          onReiniciarPin(dniAg);
                          setConfirmando(null);
                        }}
                      >
                        Sí, reiniciar
                      </button>
                      <button style={styles.confirmarBtnNo} onClick={() => setConfirmando(null)}>
                        Cancelar
                      </button>
                    </div>
                  ) : (
                    <button style={styles.linkBtnMini} onClick={() => setConfirmando(dniAg)}>
                      Reiniciar PIN
                    </button>
                  )
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PanelSincronizarGuia({ agricultores, onSincronizar }) {
  const [filas, setFilas] = useState([]);
  const [error, setError] = useState("");
  const [nombreArchivo, setNombreArchivo] = useState("");
  const [guardando, setGuardando] = useState(false);
  const [resultado, setResultado] = useState(null);

  const procesarArchivo = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setResultado(null);
    setNombreArchivo(file.name);

    const lector = new FileReader();
    lector.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const tabla = localizarTablaGuia(wb);
        if (!tabla) {
          setError("No encuentro una tabla con columnas de DNI y Kilos en ninguna hoja del archivo.");
          return;
        }
        const { filasDatos, headersReales, precioGlobal } = tabla;
        const colDni = encontrarColumna(headersReales, ["dni"]);
        const colKilos = encontrarColumna(headersReales, ["kilo"]);
        const colPrecio = encontrarColumna(headersReales, ["precio"]);
        const colN = encontrarColumna(headersReales, ["ticket", "albar", "nº", "numero", "n°"]);
        const colNombre = encontrarColumna(headersReales, ["agricultor", "nombre", "cliente", "proveedor"]);
        const colPoblacion = encontrarColumna(headersReales, ["poblac", "municipio", "localidad"]);

        if (!colDni || !colKilos) {
          setError("He encontrado una tabla pero no distingo bien las columnas de DNI y Kilos.");
          return;
        }

        const procesadas = filasDatos
          .filter((f) => String(f[colDni] || "").trim() !== "" || Number(f[colKilos]) > 0)
          .map((f, i) => {
            const dni = String(f[colDni] || "").trim().toUpperCase();
            const nombreFila = colNombre ? String(f[colNombre] || "").trim() : "";
            const yaEsCliente = !!agricultores[dni];
            const kilos = Number(f[colKilos]) || 0;
            let estado;
            if (!dni) {
              estado = kilos > 0 ? "soloKilos" : "sinDatos";
            } else {
              estado = yaEsCliente ? "cliente" : nombreFila ? "nuevo" : "sinDatos";
            }
            return {
              dni,
              nombreFila,
              poblacionFila: colPoblacion ? String(f[colPoblacion] || "").trim() : "",
              n: colN && f[colN] !== "" ? Number(f[colN]) : Date.now() + i,
              kilos,
              precio: colPrecio && Number(f[colPrecio]) > 0 ? Number(f[colPrecio]) : precioGlobal || 0,
              estado,
            };
          });
        setFilas(procesadas);
      } catch {
        setError("No he podido leer el archivo. Comprueba que es un Excel (.xlsx) o CSV válido.");
      }
    };
    lector.readAsArrayBuffer(file);
  };

  const sincronizar = async () => {
    setGuardando(true);
    const res = await onSincronizar(filas);
    setGuardando(false);
    setResultado(res);
    setFilas([]);
    setNombreArchivo("");
  };

  const yaClientes = filas.filter((f) => f.estado === "cliente").length;
  const nuevosClientes = filas.filter((f) => f.estado === "nuevo").length;
  const kgSueltos = filas.filter((f) => f.estado === "soloKilos").reduce((s, f) => s + f.kilos, 0);

  return (
    <div style={{ marginTop: 20 }}>
      <p style={styles.cardHint}>
        Sube tu "GUÍA.xlsx" tal cual la generas. Los albaranes van a la carpeta "Por revisar";
        los kilos sin agricultor se suman al contador de báscula.
      </p>

      <label style={styles.subirBtn}>
        <Upload size={16} /> {nombreArchivo || "Subir Excel o CSV de hoy"}
        <input type="file" accept=".xlsx,.xls,.csv" onChange={procesarArchivo} style={{ display: "none" }} />
      </label>
      {error && (
        <p style={styles.errorText}>
          <AlertTriangle size={13} style={{ flexShrink: 0, marginTop: 1 }} /> {error}
        </p>
      )}

      {filas.length > 0 && (
        <>
          <div style={{ display: "flex", gap: 10, marginTop: 16, flexWrap: "wrap" }}>
            <div style={styles.statCard}>
              <p style={styles.statLabel}>Ya son clientes</p>
              <p style={styles.statValue}>{yaClientes}</p>
            </div>
            <div style={styles.statCard}>
              <p style={styles.statLabel}>Clientes nuevos</p>
              <p style={styles.statValue}>{nuevosClientes}</p>
            </div>
            {kgSueltos > 0 && (
              <div style={styles.statCard}>
                <p style={styles.statLabel}>Kg sin agricultor</p>
                <p style={styles.statValue}>{kgSueltos.toLocaleString("es-ES")}</p>
              </div>
            )}
          </div>

          <div style={{ ...styles.albaranList, marginTop: 14 }}>
            {filas.map((f, i) => (
              <div key={i} style={styles.albaranRow}>
                <div>
                  <p style={styles.albaranN}>
                    {f.estado === "sinDatos"
                      ? "Fila sin datos suficientes"
                      : f.estado === "soloKilos"
                      ? "Kilos de báscula (sin agricultor)"
                      : agricultores[f.dni]?.nombre || f.nombreFila}
                  </p>
                  <p style={styles.albaranFecha}>
                    {f.estado === "soloKilos"
                      ? "Se suma al contador de báscula"
                      : `${f.dni} · nº ${f.n}${f.estado === "nuevo" ? " · cliente nuevo" : ""}`}
                  </p>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <p style={styles.albaranKilos}>{f.kilos.toLocaleString("es-ES")} kg</p>
                  {f.estado === "cliente" && <CheckCircle2 size={17} color="#4E7D3A" />}
                  {f.estado === "nuevo" && <UserPlus size={17} color="#8A5A1F" />}
                  {f.estado === "soloKilos" && <Scale size={17} color="#5B6E32" />}
                  {f.estado === "sinDatos" && <AlertTriangle size={17} color="#B4552F" />}
                </div>
              </div>
            ))}
          </div>

          <button
            style={{ ...styles.primaryBtn, marginTop: 16 }}
            onClick={sincronizar}
            disabled={guardando || yaClientes + nuevosClientes + kgSueltos === 0}
          >
            {guardando ? <Loader2 size={16} className="animate-spin" /> : <ArrowRight size={16} />}
            Sincronizar {yaClientes + nuevosClientes} albaranes
          </button>
        </>
      )}

      {resultado && (
        <div style={{ ...styles.resultadoBox, marginTop: 16 }}>
          <p style={styles.resultadoTitulo}>
            <CheckCircle2 size={16} color="#4E7D3A" /> Hecho
          </p>
          <p style={styles.resultadoTexto}>
            {resultado.añadidos} albaranes puestos en "Por revisar"
            {resultado.clientesNuevos > 0 ? ` · ${resultado.clientesNuevos} clientes nuevos` : ""}
            {resultado.kgSueltosHoy > 0
              ? ` · +${resultado.kgSueltosHoy.toLocaleString("es-ES")} kg (total: ${resultado.kgTotalNuevo.toLocaleString("es-ES")} kg)`
              : ""}
            . Ve a la pestaña "Por revisar" para aprobarlos.
          </p>
        </div>
      )}
    </div>
  );
}

function PanelFacturasPuesto({ agricultores, facturas, onSubir, onVaciar }) {
  const [dniElegido, setDniElegido] = useState("");
  const [subiendo, setSubiendo] = useState(false);
  const [error, setError] = useState("");
  const [confirmandoVaciar, setConfirmandoVaciar] = useState(false);

  const subirArchivo = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    if (!dniElegido) {
      setError("Elige primero a qué agricultor pertenece la factura.");
      e.target.value = "";
      return;
    }
    if (file.type !== "application/pdf") {
      setError("Solo se admiten archivos PDF.");
      e.target.value = "";
      return;
    }
    if (file.size > 500 * 1024) {
      setError("El PDF pesa demasiado (máx. 500 KB por ahora). Comprímelo e inténtalo de nuevo.");
      e.target.value = "";
      return;
    }
    setSubiendo(true);
    const lector = new FileReader();
    lector.onload = async (ev) => {
      const resultado = await onSubir(dniElegido, file.name, ev.target.result);
      setSubiendo(false);
      e.target.value = "";
      if (resultado && resultado.ok === false) {
        setError("No se ha podido guardar el archivo. Prueba con un PDF más ligero.");
      }
    };
    lector.readAsDataURL(file);
  };

  const ordenadas = [...facturas].sort((a, b) => b.fechaSubida - a.fechaSubida);

  return (
    <div style={{ marginTop: 20 }}>
      <p style={styles.sectionLabel}>Subir factura</p>
      <select
        style={styles.input}
        value={dniElegido}
        onChange={(e) => setDniElegido(e.target.value)}
      >
        <option value="">Elige un agricultor...</option>
        {Object.entries(agricultores).map(([dniAg, datos]) => (
          <option key={dniAg} value={dniAg}>
            {datos.nombre} — {dniAg}
          </option>
        ))}
      </select>
      <label style={{ ...styles.subirBtn, opacity: subiendo ? 0.6 : 1 }}>
        {subiendo ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />}
        Subir PDF
        <input type="file" accept="application/pdf" onChange={subirArchivo} style={{ display: "none" }} disabled={subiendo} />
      </label>
      {error && (
        <p style={styles.errorText}>
          <AlertTriangle size={13} style={{ flexShrink: 0, marginTop: 1 }} /> {error}
        </p>
      )}

      <div style={styles.avisoEspacioBox}>
        <p style={styles.tinyHintLeft}>
          Por ahora las facturas (todas juntas) caben en 1 MB, así que el límite por PDF es de
          500 KB. Es un paso intermedio — el siguiente cambio (guardar los PDF en el "almacén de
          archivos" de Firebase) quita esta limitación por completo, sin tocar nada más de la app.
        </p>
        {confirmandoVaciar ? (
          <div style={styles.confirmarRow}>
            <button
              style={styles.confirmarBtnSi}
              onClick={() => {
                onVaciar();
                setConfirmandoVaciar(false);
              }}
            >
              Sí, vaciar
            </button>
            <button style={styles.confirmarBtnNo} onClick={() => setConfirmandoVaciar(false)}>
              Cancelar
            </button>
          </div>
        ) : (
          <button style={styles.linkBtnMini} onClick={() => setConfirmandoVaciar(true)}>
            Vaciar todas las facturas guardadas
          </button>
        )}
      </div>

      <p style={{ ...styles.sectionLabel, marginTop: 24 }}>Facturas subidas</p>
      {ordenadas.length === 0 ? (
        <p style={styles.emptyText}>Todavía no has subido ninguna factura.</p>
      ) : (
        <div style={styles.albaranList}>
          {ordenadas.map((f) => (
            <FacturaFilaPuesto key={f.id} factura={f} nombreAgricultor={agricultores[f.dni]?.nombre || f.dni} />
          ))}
        </div>
      )}
    </div>
  );
}

function FacturaFilaPuesto({ factura: f, nombreAgricultor }) {
  const [error, setError] = useState("");

  const descargar = (tipo) => {
    setError("");
    const url = tipo === "firmada" ? f.dataUrlFirmada : f.dataUrl;
    if (url) {
      descargarDataUrl(url, tipo === "firmada" ? `firmada-${f.nombreArchivo}` : f.nombreArchivo);
    } else {
      setError("Este archivo no está disponible.");
    }
  };

  return (
    <div style={styles.facturaCard}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ display: "flex", gap: 10 }}>
          <FileText size={20} color="#8A5A1F" style={{ flexShrink: 0, marginTop: 2 }} />
          <div>
            <p style={styles.albaranN}>{nombreAgricultor}</p>
            <p style={styles.albaranFecha}>{f.nombreArchivo}</p>
          </div>
        </div>
        <span
          style={{
            ...styles.estadoBadge,
            ...(f.firmada ? styles.estadoBadgeOk : styles.estadoBadgeAviso),
          }}
        >
          {f.firmada ? "Firmada" : "Sin firmar"}
        </span>
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <button style={{ ...styles.ghostBtn, flex: 1 }} onClick={() => descargar("original")}>
          <Download size={14} /> Original
        </button>
        {f.firmada && (
          <button style={{ ...styles.ghostBtn, flex: 1 }} onClick={() => descargar("firmada")}>
            <CheckCircle2 size={14} color="#4E7D3A" /> Firmada
          </button>
        )}
      </div>
      {error && (
        <p style={styles.errorText}>
          <AlertTriangle size={13} style={{ flexShrink: 0, marginTop: 1 }} /> {error}
        </p>
      )}
    </div>
  );
}

function PanelRevisarAlbaranes({ kgTotales, onReiniciarKgTotales, pendientes, onAlternar, onMarcarTodos, onEnviar, onEmpezarCampañaNueva }) {
  const [confirmandoReinicio, setConfirmandoReinicio] = useState(false);
  const [confirmandoCampaña, setConfirmandoCampaña] = useState(false);
  const [textoConfirmacion, setTextoConfirmacion] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [confirmandoEnvio, setConfirmandoEnvio] = useState(false);

  const revisados = pendientes.filter((p) => p.revisado).length;

  const enviar = async () => {
    setEnviando(true);
    await onEnviar();
    setEnviando(false);
    setConfirmandoEnvio(false);
  };

  const confirmarCampaña = () => {
    if (textoConfirmacion.trim().toUpperCase() !== "BORRAR") return;
    onEmpezarCampañaNueva();
    setConfirmandoCampaña(false);
    setTextoConfirmacion("");
  };

  return (
    <div style={{ marginTop: 20 }}>
      <div style={styles.kgCard}>
        <p style={styles.kgLabel}>Kilos pesados en báscula puesto</p>
        <p style={styles.kgValor}>{kgTotales.toLocaleString("es-ES")} kg</p>
        {confirmandoReinicio ? (
          <div style={styles.confirmarRow}>
            <button
              style={styles.confirmarBtnSi}
              onClick={() => {
                onReiniciarKgTotales();
                setConfirmandoReinicio(false);
              }}
            >
              Sí, poner a 0
            </button>
            <button style={styles.confirmarBtnNo} onClick={() => setConfirmandoReinicio(false)}>
              Cancelar
            </button>
          </div>
        ) : (
          <button style={styles.linkBtnMini} onClick={() => setConfirmandoReinicio(true)}>
            Reiniciar contador
          </button>
        )}
      </div>

      <div style={styles.campañaCard}>
        <p style={styles.kgLabel}>Empezar campaña nueva</p>
        <p style={styles.tinyHintLeft}>
          Borra albaranes, pedidos, avisos, kilos y la carpeta de revisión. Mantiene a los
          agricultores dados de alta (con su DNI, PIN y facturas) — no hace falta que se
          registren otra vez.
        </p>
        {confirmandoCampaña ? (
          <div style={{ marginTop: 10 }}>
            <p style={styles.tinyHintLeft}>Escribe BORRAR para confirmar:</p>
            <input
              style={{ ...styles.input, marginTop: 6 }}
              value={textoConfirmacion}
              onChange={(e) => setTextoConfirmacion(e.target.value)}
              placeholder="BORRAR"
            />
            <div style={styles.confirmarRow}>
              <button
                style={{ ...styles.confirmarBtnSi, ...(textoConfirmacion.trim().toUpperCase() !== "BORRAR" ? { opacity: 0.5 } : {}) }}
                onClick={confirmarCampaña}
                disabled={textoConfirmacion.trim().toUpperCase() !== "BORRAR"}
              >
                Sí, empezar de cero
              </button>
              <button
                style={styles.confirmarBtnNo}
                onClick={() => {
                  setConfirmandoCampaña(false);
                  setTextoConfirmacion("");
                }}
              >
                Cancelar
              </button>
            </div>
          </div>
        ) : (
          <button style={{ ...styles.ghostBtn, marginTop: 10, width: "100%" }} onClick={() => setConfirmandoCampaña(true)}>
            Empezar campaña nueva
          </button>
        )}
      </div>

      <p style={{ ...styles.sectionLabel, marginTop: 24 }}>
        Albaranes del día por revisar {pendientes.length > 0 ? `(${revisados}/${pendientes.length} revisados)` : ""}
      </p>
      <p style={styles.cardHint}>
        Estos albaranes han llegado de la guía pero todavía NO los ve el agricultor. Marca cada
        uno como revisado y pulsa "Enviar" — solo se mandan los que estén en verde; los que
        dejes en pendiente se quedan aquí esperando.
      </p>

      {pendientes.length > 0 && (
        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <button style={{ ...styles.tipoBtn, flex: 1 }} onClick={() => onMarcarTodos(true)}>
            <CheckCircle2 size={13} /> Todos revisados
          </button>
          <button style={{ ...styles.tipoBtn, flex: 1 }} onClick={() => onMarcarTodos(false)}>
            <AlertTriangle size={13} /> Todos pendientes
          </button>
        </div>
      )}

      {pendientes.length === 0 ? (
        <p style={styles.emptyText}>No hay albaranes esperando revisión.</p>
      ) : (
        <div style={styles.albaranList}>
          {pendientes.map((p) => (
            <div key={p.id} style={styles.albaranRow}>
              <div>
                <p style={styles.albaranN}>{p.nombreAgricultor}</p>
                <p style={styles.albaranFecha}>
                  {p.dni} · Albarán nº {p.n} · {fecha(p.fecha)}
                </p>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ textAlign: "right" }}>
                  <p style={styles.albaranKilos}>{p.kilos.toLocaleString("es-ES")} kg</p>
                  <p style={styles.albaranPrecio}>{p.precio.toFixed(2)} €/kg</p>
                </div>
                <button
                  style={{
                    ...styles.estadoToggleBtn,
                    ...(p.revisado ? styles.estadoToggleOk : styles.estadoTogglePendiente),
                  }}
                  onClick={() => onAlternar(p.id)}
                >
                  {p.revisado ? "REVISADO" : "PENDIENTE"}
                </button>
              </div>
            </div>
          ))}
        </div>
      )}

      {pendientes.length > 0 && (
        <>
          {confirmandoEnvio ? (
            <div style={{ ...styles.confirmarRow, justifyContent: "center", marginTop: 16 }}>
              <button style={styles.confirmarBtnSi} onClick={enviar} disabled={enviando}>
                {enviando ? <Loader2 size={14} className="animate-spin" /> : "Sí, enviar ya"}
              </button>
              <button style={styles.confirmarBtnNo} onClick={() => setConfirmandoEnvio(false)}>
                Cancelar
              </button>
            </div>
          ) : (
            <button
              style={{ ...styles.primaryBtn, marginTop: 18, ...(revisados === 0 ? { opacity: 0.5 } : {}) }}
              onClick={() => setConfirmandoEnvio(true)}
              disabled={revisados === 0}
            >
              <Send size={16} /> Enviar {revisados} albaranes a los agricultores
            </button>
          )}
          {revisados < pendientes.length && (
            <p style={styles.tinyHint}>
              {pendientes.length - revisados} albarán{pendientes.length - revisados === 1 ? "" : "es"} sigue
              {pendientes.length - revisados === 1 ? "" : "n"} en pendiente y no se enviará
              {pendientes.length - revisados === 1 ? "" : "n"} hasta que lo marques como revisado.
            </p>
          )}
        </>
      )}
    </div>
  );
}

function PanelPedidosPuesto({ precios, onActualizarPrecios, pedidos, onCambiarEstado }) {
  const [garrafa, setGarrafa] = useState(String(precios.garrafa));
  const [caja, setCaja] = useState(String(precios.caja));
  const [guardado, setGuardado] = useState(false);

  const guardarPrecios = () => {
    const g = parseFloat(garrafa.replace(",", "."));
    const c = parseFloat(caja.replace(",", "."));
    if (isNaN(g) || isNaN(c)) return;
    onActualizarPrecios({ garrafa: g, caja: c });
    setGuardado(true);
    setTimeout(() => setGuardado(false), 1800);
  };

  const ordenados = [...pedidos].sort((a, b) => (b.ts || b.id) - (a.ts || a.id));
  const etiquetaEstado = { pendiente: "Pendiente", preparado: "Preparado", entregado: "Entregado" };

  return (
    <div style={{ marginTop: 20 }}>
      <p style={styles.sectionLabel}>Precio por unidad</p>
      <div style={{ display: "flex", gap: 10, marginBottom: 8 }}>
        <div style={{ flex: 1 }}>
          <p style={styles.tinyLabel}>Garrafa (€)</p>
          <input
            style={styles.input}
            inputMode="decimal"
            value={garrafa}
            onChange={(e) => setGarrafa(e.target.value)}
          />
        </div>
        <div style={{ flex: 1 }}>
          <p style={styles.tinyLabel}>Caja (€)</p>
          <input
            style={styles.input}
            inputMode="decimal"
            value={caja}
            onChange={(e) => setCaja(e.target.value)}
          />
        </div>
      </div>
      <button style={styles.primaryBtn} onClick={guardarPrecios}>
        Guardar precios
      </button>
      {guardado && <p style={styles.tinyHint}>Precios actualizados.</p>}

      <p style={{ ...styles.sectionLabel, marginTop: 28 }}>Pedidos de los agricultores</p>
      {ordenados.length === 0 ? (
        <p style={styles.emptyText}>Todavía no hay pedidos.</p>
      ) : (
        <div style={styles.albaranList}>
          {ordenados.map((p) => {
            const precioUnidad = p.tipo === "garrafa" ? precios.garrafa : precios.caja;
            const total = precioUnidad * p.cantidad;
            return (
              <button
                key={p.id}
                style={{ ...styles.albaranRow, cursor: "pointer", border: "none", width: "100%", textAlign: "left" }}
                onClick={() => onCambiarEstado(p.id)}
              >
                <div>
                  <p style={styles.albaranN}>
                    {p.nombre} — {p.cantidad} {p.tipo}
                    {p.cantidad === 1 ? "" : "s"}
                  </p>
                  <p style={styles.albaranFecha}>
                    {new Date(p.ts || p.id).toLocaleString("es-ES", {
                      day: "2-digit",
                      month: "short",
                      hour: "2-digit",
                      minute: "2-digit",
                    })}{" "}
                    · {money(total)}
                  </p>
                </div>
                <span
                  style={{
                    ...styles.estadoBadge,
                    ...(p.estado === "entregado"
                      ? styles.estadoBadgeOk
                      : p.estado === "preparado"
                      ? styles.estadoBadgePreparado
                      : styles.estadoBadgeAviso),
                  }}
                >
                  {etiquetaEstado[p.estado]}
                </span>
              </button>
            );
          })}
        </div>
      )}
      {ordenados.length > 0 && (
        <p style={styles.tinyHint}>Toca un pedido para cambiar su estado.</p>
      )}
    </div>
  );
}

function PanelTiempoEsperaPuesto({ tiempoEspera, onActualizar }) {
  const [valor, setValor] = useState(tiempoEspera);
  const [guardado, setGuardado] = useState(false);

  useEffect(() => {
    setValor(tiempoEspera);
  }, [tiempoEspera]);

  const ajustar = (delta) => {
    const nuevo = Math.max(0, valor + delta);
    setValor(nuevo);
    onActualizar(nuevo);
    setGuardado(true);
    setTimeout(() => setGuardado(false), 1400);
  };

  return (
    <div>
      <p style={styles.sectionLabel}>Tiempo de espera del puesto</p>
      <EsperaBar minutos={valor} />
      <div style={styles.cantidadRow}>
        <button style={styles.cantidadBtn} onClick={() => ajustar(-5)}>
          −
        </button>
        <span style={styles.cantidadValor}>{valor} min</span>
        <button style={styles.cantidadBtn} onClick={() => ajustar(5)}>
          +
        </button>
      </div>
      <p style={styles.tinyHint}>
        {guardado ? "Actualizado — los agricultores ya lo ven." : "Se guarda solo, sin botón."}
      </p>
    </div>
  );
}

function PanelCodigo({ codigo, onActualizar }) {
  const [valor, setValor] = useState(codigo);
  const [guardado, setGuardado] = useState(false);

  const generarNuevo = () => {
    const random = String(Math.floor(Math.random() * 90) + 10); // 10-99
    setValor(random);
    onActualizar(random);
    setGuardado(true);
    setTimeout(() => setGuardado(false), 1800);
  };

  const guardarManual = () => {
    if (!/^\d{2}$/.test(valor)) return;
    onActualizar(valor);
    setGuardado(true);
    setTimeout(() => setGuardado(false), 1800);
  };

  return (
    <div style={{ marginTop: 20 }}>
      <p style={styles.cardHint}>
        Solo los agricultores que te pidan este número en persona podrán crear su cuenta y
        elegir su PIN. Cámbialo cuando quieras.
      </p>
      <div style={styles.codigoBox}>{codigo}</div>
      <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
        <input
          style={{ ...styles.input, marginBottom: 0, flex: 1 }}
          inputMode="numeric"
          placeholder="Nuevo número (2 cifras)"
          value={valor}
          onChange={(e) => setValor(e.target.value)}
        />
        <button
          style={{ ...styles.ghostBtn, width: "auto", padding: "0 16px" }}
          onClick={guardarManual}
        >
          Guardar
        </button>
      </div>
      <button style={{ ...styles.primaryBtn, marginTop: 12 }} onClick={generarNuevo}>
        Generar número aleatorio
      </button>
      {guardado && <p style={styles.tinyHint}>Número actualizado.</p>}
    </div>
  );
}

function PanelChatPuesto({ mensajes, perfiles, onEliminar, onEnviar, onEnviarAudio }) {
  const [texto, setTexto] = useState("");

  const enviar = () => {
    if (!texto.trim()) return;
    onEnviar(texto.trim());
    setTexto("");
  };

  return (
    <div style={{ marginTop: 20 }}>
      <p style={styles.cardHint}>
        Este es el mismo Whatsapillo que ven los agricultores. Escribe como El Jaralillo, o
        pasa el ratón/toca cualquier mensaje para borrarlo.
      </p>
      <div style={styles.chatBox}>
        <HiloWhatsapillo
          mensajes={mensajes}
          perfiles={perfiles}
          propioDni={PUESTO_DNI}
          onEliminar={onEliminar}
        />
      </div>
      <div style={styles.chatInputRow}>
        <input
          style={{ ...styles.input, marginBottom: 0, flex: 1 }}
          placeholder="Escribe un mensaje al chat..."
          value={texto}
          onChange={(e) => setTexto(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && enviar()}
        />
        <BotonGrabar onGrabado={onEnviarAudio} />
        <button style={styles.chatEnviarBtn} onClick={enviar}>
          <Send size={17} />
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Piezas comunes
// ---------------------------------------------------------------------------
function Avatar({ foto, nombre, size = 40 }) {
  if (foto) {
    return (
      <img
        src={foto}
        alt={nombre}
        style={{
          width: size,
          height: size,
          borderRadius: "50%",
          objectFit: "cover",
          flexShrink: 0,
        }}
      />
    );
  }
  const iniciales = (nombre || "?")
    .split(" ")
    .filter(Boolean)
    .slice(0, 2)
    .map((p) => p[0].toUpperCase())
    .join("");
  return (
    <div
      style={{
        width: size,
        height: size,
        borderRadius: "50%",
        background: "#E0C88C",
        color: "#4A3A14",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontWeight: 700,
        fontSize: Math.round(size * 0.4),
        fontFamily: FONT_DISPLAY,
        flexShrink: 0,
      }}
    >
      {iniciales}
    </div>
  );
}

function TopBar({ label, sub, onSalir, avatarFoto, avatarNombre }) {
  return (
    <div>
      <img src={LOGO_B64} alt="El Jaralillo" style={styles.logoHeader} />
      <div style={styles.topBar}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {avatarNombre && <Avatar foto={avatarFoto} nombre={avatarNombre} size={38} />}
          <div>
            <p style={styles.topBarLabel}>{label}</p>
            <p style={styles.topBarSub}>{sub}</p>
          </div>
        </div>
        <button style={styles.logoutBtn} onClick={onSalir}>
          <LogOut size={16} />
        </button>
      </div>
    </div>
  );
}

function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,wght@0,500;0,600;1,500&family=Inter:wght@400;500;600;700&display=swap');
      * { box-sizing: border-box; }
      body { margin: 0; }
      input::placeholder, textarea::placeholder { color: #9C9182; }
      input:focus, textarea:focus, button:focus-visible {
        outline: 2px solid #C9982B;
        outline-offset: 2px;
      }
      .animate-spin {
        animation: girar 1s linear infinite;
      }
      @keyframes girar {
        from { transform: rotate(0deg); }
        to { transform: rotate(360deg); }
      }
    `}</style>
  );
}

// ---------------------------------------------------------------------------
// Estilos (paleta olivo / oro de aceite / arcilla)
// ---------------------------------------------------------------------------
const FONT_DISPLAY = "'Fraunces', serif";
const FONT_BODY = "'Inter', sans-serif";

const styles = {
  app: {
    minHeight: "100vh",
    background: "#F6F1E4",
    fontFamily: FONT_BODY,
    color: "#2A2118",
    display: "flex",
    justifyContent: "center",
  },
  toast: {
    position: "fixed",
    top: 14,
    left: "50%",
    transform: "translateX(-50%)",
    zIndex: 50,
    display: "flex",
    gap: 8,
    alignItems: "flex-start",
    maxWidth: "90%",
    width: 360,
    background: "#2E3A1F",
    color: "#F6F1E4",
    borderRadius: 12,
    padding: "12px 14px",
    boxShadow: "0 10px 28px rgba(46,58,31,0.35)",
    cursor: "pointer",
  },
  toastTexto: {
    fontSize: 13,
    lineHeight: 1.4,
    margin: 0,
  },
  loadingScreen: {
    minHeight: "100vh",
    width: "100%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "#F6F1E4",
  },
  portalWrap: {
    width: "100%",
    maxWidth: 420,
    padding: "56px 28px 32px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    textAlign: "center",
  },
  seal: {
    width: 64,
    height: 64,
    borderRadius: "50%",
    background: "#2E3A1F",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 20,
    boxShadow: "0 6px 18px rgba(46,58,31,0.25)",
  },
  eyebrow: {
    fontSize: 12,
    letterSpacing: "0.14em",
    textTransform: "uppercase",
    color: "#8A7A55",
    fontWeight: 600,
    margin: "0 0 34px",
  },
  title: {
    fontFamily: FONT_DISPLAY,
    fontWeight: 600,
    fontSize: 34,
    lineHeight: 1.15,
    margin: "0 0 10px",
    color: "#2E3A1F",
  },
  subtitle: {
    fontSize: 15,
    color: "#5B5340",
    lineHeight: 1.5,
    margin: "0 0 32px",
    maxWidth: 320,
  },
  primaryBtn: {
    width: "100%",
    maxWidth: 340,
    background: "#2E3A1F",
    color: "#F6F1E4",
    border: "none",
    borderRadius: 12,
    padding: "15px 18px",
    fontSize: 15,
    fontWeight: 600,
    fontFamily: FONT_BODY,
    display: "flex",
    alignItems: "center",
    gap: 8,
    cursor: "pointer",
    marginBottom: 12,
    justifyContent: "center",
  },
  ghostBtn: {
    width: "100%",
    maxWidth: 340,
    background: "transparent",
    color: "#5B5340",
    border: "1px solid #D9CFB8",
    borderRadius: 12,
    padding: "13px 18px",
    fontSize: 14,
    fontWeight: 500,
    fontFamily: FONT_BODY,
    cursor: "pointer",
  },
  portalFoot: {
    marginTop: 40,
    fontSize: 12,
    color: "#A79B7E",
  },
  logoPortal: {
    width: 148,
    height: "auto",
    marginBottom: 10,
  },
  logoHeader: {
    width: 44,
    height: "auto",
    display: "block",
    marginBottom: 10,
  },
  centerCard: {
    width: "100%",
    maxWidth: 420,
    padding: "28px 28px 40px",
    display: "flex",
    flexDirection: "column",
  },
  backLink: {
    background: "none",
    border: "none",
    color: "#8A7A55",
    fontSize: 13,
    fontFamily: FONT_BODY,
    cursor: "pointer",
    padding: 0,
    marginBottom: 28,
    alignSelf: "flex-start",
  },
  sealSmall: {
    width: 44,
    height: 44,
    borderRadius: "50%",
    background: "#2E3A1F",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 16,
  },
  logoSello: {
    width: 26,
    height: 26,
    objectFit: "contain",
    filter: "brightness(0) invert(1)",
  },
  cardTitle: {
    fontFamily: FONT_DISPLAY,
    fontSize: 24,
    fontWeight: 600,
    margin: "0 0 6px",
    color: "#2E3A1F",
  },
  cardHint: {
    fontSize: 14,
    color: "#5B5340",
    margin: "0 0 20px",
  },
  input: {
    width: "100%",
    padding: "13px 14px",
    borderRadius: 10,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    fontSize: 15,
    fontFamily: FONT_BODY,
    marginBottom: 14,
  },
  errorText: {
    color: "#B4552F",
    fontSize: 13,
    margin: "-6px 0 14px",
  },
  tinyHint: {
    fontSize: 12,
    color: "#A79B7E",
    marginTop: 16,
    textAlign: "center",
  },
  linkBtn: {
    background: "none",
    border: "none",
    color: "#8A7A55",
    fontSize: 13,
    fontFamily: FONT_BODY,
    textDecoration: "underline",
    cursor: "pointer",
    marginTop: 14,
    padding: 0,
    alignSelf: "center",
  },
  linkBtnMini: {
    background: "none",
    border: "none",
    color: "#B4552F",
    fontSize: 11.5,
    fontFamily: FONT_BODY,
    textDecoration: "underline",
    cursor: "pointer",
    padding: 0,
    marginTop: 8,
  },
  perfilAvatarWrap: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: 8,
    padding: "10px 0 4px",
  },
  perfilCambiarFotoBtn: {
    display: "inline-block",
    fontSize: 12.5,
    fontWeight: 600,
    color: "#2E3A1F",
    background: "#F1E4C4",
    border: "1px solid #E0C88C",
    borderRadius: 999,
    padding: "7px 14px",
    cursor: "pointer",
  },
  tinyHintLeft: {
    fontSize: 12,
    color: "#8A7A55",
    margin: "8px 0 0",
    lineHeight: 1.4,
  },
  confirmarRow: {
    display: "flex",
    gap: 6,
    marginTop: 8,
    justifyContent: "flex-end",
  },
  confirmarBtnSi: {
    fontSize: 11.5,
    padding: "5px 9px",
    borderRadius: 7,
    border: "none",
    background: "#B4552F",
    color: "#FBF1EC",
    cursor: "pointer",
  },
  confirmarBtnNo: {
    fontSize: 11.5,
    padding: "5px 9px",
    borderRadius: 7,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    color: "#5B5340",
    cursor: "pointer",
  },
  screen: {
    width: "100%",
    maxWidth: 460,
    padding: "24px 20px 48px",
  },
  topBar: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: 22,
  },
  topBarLabel: {
    fontFamily: FONT_DISPLAY,
    fontSize: 20,
    fontWeight: 600,
    margin: 0,
    color: "#2E3A1F",
  },
  topBarSub: {
    fontSize: 13,
    color: "#8A7A55",
    margin: "2px 0 0",
  },
  logoutBtn: {
    width: 36,
    height: 36,
    borderRadius: 10,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    color: "#5B5340",
  },
  avisosBlock: {
    display: "flex",
    flexDirection: "column",
    gap: 8,
    marginBottom: 20,
  },
  avisoCard: {
    display: "flex",
    gap: 10,
    padding: "12px 14px",
    borderRadius: 10,
  },
  avisoText: {
    fontSize: 13.5,
    lineHeight: 1.45,
    margin: 0,
    color: "#2A2118",
    whiteSpace: "pre-wrap",
  },
  statsRow: {
    display: "flex",
    gap: 12,
    marginBottom: 26,
  },
  statCard: {
    flex: 1,
    background: "#FFFDF8",
    border: "1px solid #EAE1CB",
    borderRadius: 12,
    padding: "14px 16px",
  },
  statLabel: {
    fontSize: 11.5,
    color: "#8A7A55",
    margin: "0 0 6px",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  statValue: {
    fontFamily: FONT_DISPLAY,
    fontSize: 19,
    fontWeight: 600,
    margin: 0,
    color: "#2E3A1F",
  },
  sectionLabel: {
    fontSize: 12,
    fontWeight: 600,
    color: "#8A7A55",
    textTransform: "uppercase",
    letterSpacing: "0.06em",
    margin: "0 0 12px",
  },
  emptyText: {
    fontSize: 14,
    color: "#8A7A55",
  },
  albaranList: {
    display: "flex",
    flexDirection: "column",
    gap: 10,
  },
  albaranRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    background: "#FFFDF8",
    border: "1px solid #EAE1CB",
    borderRadius: 12,
    padding: "13px 16px",
  },
  albaranN: {
    fontSize: 14,
    fontWeight: 600,
    margin: 0,
    color: "#2A2118",
  },
  albaranFecha: {
    fontSize: 12.5,
    color: "#8A7A55",
    margin: "2px 0 0",
    textTransform: "capitalize",
  },
  albaranKilos: {
    fontSize: 14,
    fontWeight: 600,
    margin: 0,
    color: "#2E3A1F",
  },
  albaranPrecio: {
    fontSize: 12.5,
    color: "#8A7A55",
    margin: "2px 0 0",
  },
  tipoRow: {
    display: "flex",
    gap: 8,
    marginBottom: 14,
    flexWrap: "wrap",
  },
  tipoBtn: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    padding: "9px 13px",
    borderRadius: 999,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    color: "#5B5340",
    fontSize: 13,
    fontFamily: FONT_BODY,
    cursor: "pointer",
  },
  tipoBtnActivo: {
    background: "#2E3A1F",
    borderColor: "#2E3A1F",
    color: "#F6F1E4",
  },
  textarea: {
    width: "100%",
    minHeight: 90,
    padding: "13px 14px",
    borderRadius: 10,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    fontSize: 14,
    fontFamily: FONT_BODY,
    marginBottom: 14,
    resize: "vertical",
  },
  agricultorList: {
    display: "flex",
    flexDirection: "column",
    gap: 6,
    background: "#FFFDF8",
    border: "1px solid #EAE1CB",
    borderRadius: 12,
    padding: 8,
    maxHeight: 220,
    overflowY: "auto",
  },
  agricultorRow: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "9px 10px",
    borderRadius: 9,
    cursor: "pointer",
  },
  agricultorRowActivo: {
    background: "#EFE7CE",
  },
  checkbox: {
    width: 17,
    height: 17,
    accentColor: "#2E3A1F",
    flexShrink: 0,
  },
  agricultorNombre: {
    fontSize: 13.5,
    fontWeight: 600,
    margin: 0,
    color: "#2A2118",
  },
  agricultorDni: {
    fontSize: 12,
    color: "#8A7A55",
    margin: "1px 0 0",
  },
  destinoTag: {
    fontSize: 11.5,
    color: "#A79B7E",
    margin: "0 0 4px 2px",
  },
  tabRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
    borderBottom: "1px solid #EAE1CB",
    paddingBottom: 12,
  },
  tabBtn: {
    padding: "8px 14px",
    borderRadius: 999,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    color: "#5B5340",
    fontSize: 13,
    fontWeight: 500,
    fontFamily: FONT_BODY,
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    gap: 6,
    whiteSpace: "nowrap",
  },
  tabBtnActivo: {
    background: "#2E3A1F",
    borderColor: "#2E3A1F",
    color: "#F6F1E4",
  },
  avisoHora: {
    fontSize: 11,
    color: "#6B6350",
    margin: "4px 0 0",
  },
  esperaCard: {
    borderRadius: 14,
    padding: "14px 16px",
    marginBottom: 4,
  },
  esperaTop: {
    display: "flex",
    alignItems: "center",
    gap: 7,
    marginBottom: 4,
  },
  esperaTitulo: {
    fontSize: 12,
    fontWeight: 600,
    color: "#5B5340",
    margin: 0,
  },
  esperaMinutos: {
    fontFamily: FONT_DISPLAY,
    fontSize: 24,
    fontWeight: 600,
    margin: "0 0 10px",
  },
  esperaEtiqueta: {
    fontFamily: FONT_BODY,
    fontSize: 12.5,
    fontWeight: 500,
  },
  esperaTrack: {
    position: "relative",
    height: 10,
    borderRadius: 999,
    background: "rgba(255,255,255,0.55)",
    overflow: "hidden",
  },
  esperaFill: {
    position: "absolute",
    top: 0,
    left: 0,
    height: "100%",
    borderRadius: 999,
    transition: "width 0.3s ease",
  },
  esperaMarca: {
    position: "absolute",
    top: 0,
    width: 2,
    height: "100%",
    background: "rgba(255,255,255,0.7)",
  },
  chatBox: {
    display: "flex",
    flexDirection: "column",
    gap: 8,
    marginBottom: 14,
    maxHeight: 380,
    overflowY: "auto",
    padding: "4px 2px",
  },
  chatBurbuja: {
    maxWidth: "78%",
    padding: "9px 12px",
    borderRadius: 14,
  },
  chatBurbujaMia: {
    alignSelf: "flex-end",
    background: "#2E3A1F",
    color: "#F6F1E4",
    borderBottomRightRadius: 4,
  },
  chatBurbujaOtro: {
    alignSelf: "flex-start",
    background: "#FFFDF8",
    border: "1px solid #EAE1CB",
    color: "#2A2118",
    borderBottomLeftRadius: 4,
  },
  chatBurbujaPuesto: {
    alignSelf: "center",
    background: "#F1E4C4",
    border: "1px solid #E0C88C",
    color: "#4A3A14",
    maxWidth: "88%",
  },
  chatAutorPuesto: {
    fontSize: 11.5,
    fontWeight: 700,
    margin: "0 0 2px",
    color: "#8A5A1F",
  },
  chatAutor: {
    fontSize: 11.5,
    fontWeight: 700,
    margin: "0 0 2px",
    color: "#C9982B",
  },
  chatTexto: {
    fontSize: 14,
    margin: 0,
    lineHeight: 1.4,
  },
  chatHora: {
    fontSize: 10,
    opacity: 0.65,
    margin: "4px 0 0",
    textAlign: "right",
  },
  chatBorrarMini: {
    position: "absolute",
    top: -7,
    right: -7,
    width: 20,
    height: 20,
    borderRadius: "50%",
    border: "1px solid #E3C9BC",
    background: "#FBF1EC",
    color: "#B4552F",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    padding: 0,
  },
  chatInputRow: {
    display: "flex",
    gap: 8,
    alignItems: "center",
  },
  chatEnviarBtn: {
    width: 44,
    height: 44,
    flexShrink: 0,
    borderRadius: 10,
    border: "none",
    background: "#2E3A1F",
    color: "#F6F1E4",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
  },
  micBtn: {
    width: 44,
    height: 44,
    flexShrink: 0,
    borderRadius: 10,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    color: "#5B5340",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
  },
  grabandoBtn: {
    height: 44,
    flexShrink: 0,
    borderRadius: 10,
    border: "1px solid #E3C9BC",
    background: "#FBF1EC",
    color: "#B4552F",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
    padding: "0 12px",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
  },
  puntoGrabando: {
    width: 8,
    height: 8,
    borderRadius: "50%",
    background: "#B4552F",
  },
  audioMsgRow: {
    display: "flex",
    alignItems: "center",
    gap: 6,
  },
  audioPlayer: {
    height: 34,
    maxWidth: 190,
  },
  audioDuracion: {
    fontSize: 11,
    opacity: 0.7,
    flexShrink: 0,
  },
  audioErrorTooltip: {
    position: "absolute",
    bottom: "115%",
    right: 0,
    width: 200,
    background: "#2E3A1F",
    color: "#F6F1E4",
    fontSize: 11.5,
    lineHeight: 1.4,
    padding: "8px 10px",
    borderRadius: 8,
    boxShadow: "0 6px 16px rgba(0,0,0,0.25)",
    zIndex: 20,
  },
  codigoBox: {
    fontFamily: FONT_DISPLAY,
    fontSize: 42,
    fontWeight: 600,
    letterSpacing: "0.1em",
    color: "#2E3A1F",
    background: "#FFFDF8",
    border: "1px solid #EAE1CB",
    borderRadius: 14,
    textAlign: "center",
    padding: "18px 0",
  },
  chatTextoPuesto: {
    fontSize: 13.5,
    color: "#3A3123",
    margin: "3px 0 4px",
    lineHeight: 1.4,
  },
  borrarBtn: {
    width: 30,
    height: 30,
    flexShrink: 0,
    borderRadius: 8,
    border: "1px solid #E3C9BC",
    background: "#FBF1EC",
    color: "#B4552F",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
  },
  estadoBadge: {
    display: "inline-block",
    fontSize: 11,
    fontWeight: 600,
    padding: "3px 9px",
    borderRadius: 999,
  },
  estadoBadgeOk: {
    background: "#E7EAD9",
    color: "#3F5A25",
  },
  estadoBadgeAviso: {
    background: "#F1E4C4",
    color: "#8A5A1F",
  },
  estadoBadgePreparado: {
    background: "#DCE4EE",
    color: "#2E4A6B",
  },
  tinyLabel: {
    fontSize: 11.5,
    color: "#8A7A55",
    margin: "0 0 5px",
  },
  cantidadRow: {
    display: "flex",
    alignItems: "center",
    gap: 18,
    margin: "4px 0 16px",
  },
  cantidadBtn: {
    width: 38,
    height: 38,
    borderRadius: 10,
    border: "1px solid #D9CFB8",
    background: "#FFFDF8",
    color: "#2E3A1F",
    fontSize: 18,
    fontWeight: 600,
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  cantidadValor: {
    fontFamily: FONT_DISPLAY,
    fontSize: 22,
    fontWeight: 600,
    color: "#2E3A1F",
    minWidth: 24,
    textAlign: "center",
  },
  conexionTexto: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    gap: 6,
    fontSize: 12,
    color: "#5B5340",
    margin: "6px 0 0",
  },
  puntoEstado: {
    width: 7,
    height: 7,
    borderRadius: "50%",
    flexShrink: 0,
  },
  kgCard: {
    background: "#E7EAD9",
    border: "1px solid #CFDBB8",
    borderRadius: 14,
    padding: "16px 18px",
    textAlign: "center",
  },
  campañaCard: {
    background: "#FBF1EC",
    border: "1px solid #E3C9BC",
    borderRadius: 14,
    padding: "16px 18px",
    marginTop: 14,
  },
  avisoEspacioBox: {
    background: "#FBF1EC",
    border: "1px solid #E3C9BC",
    borderRadius: 12,
    padding: "12px 14px",
    marginTop: 14,
  },
  kgLabel: {
    fontSize: 12,
    color: "#5B5340",
    margin: "0 0 6px",
  },
  kgValor: {
    fontFamily: FONT_DISPLAY,
    fontSize: 30,
    fontWeight: 600,
    color: "#3F5A25",
    margin: "0 0 8px",
  },
  estadoToggleBtn: {
    fontSize: 11,
    fontWeight: 700,
    letterSpacing: "0.03em",
    padding: "7px 11px",
    borderRadius: 8,
    border: "none",
    cursor: "pointer",
    flexShrink: 0,
  },
  estadoToggleOk: {
    background: "#4E7D3A",
    color: "#F6F1E4",
  },
  estadoTogglePendiente: {
    background: "#B4552F",
    color: "#F6F1E4",
  },
  subirBtn: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    justifyContent: "center",
    background: "#2E3A1F",
    color: "#F6F1E4",
    border: "none",
    borderRadius: 10,
    padding: "13px 16px",
    fontSize: 13.5,
    fontWeight: 600,
    cursor: "pointer",
    textAlign: "center",
  },
  resultadoBox: {
    background: "#E7EAD9",
    border: "1px solid #CFDBB8",
    borderRadius: 12,
    padding: "14px 16px",
  },
  resultadoTitulo: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    fontSize: 13.5,
    fontWeight: 700,
    color: "#3F5A25",
    margin: "0 0 4px",
  },
  resultadoTexto: {
    fontSize: 12.5,
    color: "#3A3123",
    lineHeight: 1.5,
    margin: 0,
  },
  borrarAvisoBtn: {
    position: "absolute",
    top: 8,
    right: 8,
    width: 22,
    height: 22,
    borderRadius: "50%",
    border: "1px solid #E3C9BC",
    background: "#FBF1EC",
    color: "#B4552F",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    padding: 0,
  },
  tablonFondo: {
    marginTop: 18,
    background:
      "repeating-linear-gradient(0deg, #DCC9A3, #DCC9A3 2px, #D6C098 2px, #D6C098 4px)",
    border: "1px solid #B8A377",
    borderRadius: 14,
    padding: 14,
    minHeight: 160,
  },
  tablonGrid: {
    display: "flex",
    flexDirection: "column",
    gap: 12,
  },
  notaTablon: {
    background: "#FFFBF0",
    borderRadius: 4,
    padding: "10px 12px 12px",
    boxShadow: "0 4px 10px rgba(60,46,20,0.25)",
    transform: "rotate(-0.6deg)",
  },
  facturaCard: {
    background: "#FFFDF8",
    border: "1px solid #EAE1CB",
    borderRadius: 12,
    padding: "13px 14px",
  },
  firmaCanvas: {
    width: "100%",
    height: 140,
    background: "#FFFDF8",
    border: "1px dashed #C9BFA5",
    borderRadius: 10,
    touchAction: "none",
    cursor: "crosshair",
  },
};
