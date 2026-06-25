"use strict";
/* Revolut → EUR — app estática. Lee el informe de Revolut, obtiene los tipos del BCE
   (API SDMX, con CORS) y descarga el mismo archivo en euros. Sin servidor. */

const MARCA_VENTAS = "income from sells";
const MARCA_OTROS = "other income & fees";
const SDMX = "https://data-api.ecb.europa.eu/service/data/EXR/D.USD.EUR.SP00.A" +
  "?startPeriod={ini}&endPeriod={fin}&format=csvdata";
const REQUERIDAS = ["date acquired", "date sold", "symbol", "security name", "isin",
  "country", "quantity", "cost basis", "gross proceeds", "gross pnl", "currency"];
const COLS_EXTRA = ["FX compra (USD/EUR)", "FX venta (USD/EUR)"];
const MAX_DIAS_ATRAS = 10, MARGEN_INICIO = 15;

/* ----------------------------- utilidades ----------------------------- */
const round2 = (x) => Math.round((x + Number.EPSILON) * 100) / 100;

// Corrige texto doble-codificado ('€' guardado como 'â‚¬'). Seguro: si no aplica, devuelve igual.
function repararMojibake(s) {
  if (typeof s !== "string") return s;
  for (const ch of s) if (ch.charCodeAt(0) > 255) return s;
  try {
    const bytes = Uint8Array.from([...s].map((c) => c.charCodeAt(0) & 0xff));
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch { return s; }
}

function aDecimal(v) {
  if (typeof v === "number") return v;
  const t = String(v).replace(/[€$,\s]/g, "").trim();
  if (t === "") return 0;
  const n = Number(t);
  return Number.isNaN(n) ? 0 : n;
}

function fmtISO(d) {
  const y = d.getUTCFullYear(), m = String(d.getUTCMonth() + 1).padStart(2, "0"),
    day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}
function isoDe(v) {
  if (v instanceof Date) {                          // xlsx/ods (cellDates): usa la fecha "de pared", sin desfase horario
    return fmtISO(new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate())));
  }
  if (typeof v === "number") {                      // serie de Excel/SheetJS: días desde 1899-12-30
    return fmtISO(new Date(Math.round((v - 25569) * 86400000)));
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);          // ISO aaaa-mm-dd
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);  // dd/mm/aaaa (formato europeo)
  if (m) {
    const d = m[1].padStart(2, "0"), mo = m[2].padStart(2, "0");
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${y}-${mo}-${d}`;
  }
  const dt = new Date(s);
  if (Number.isNaN(dt.getTime())) throw new Error(`Fecha no reconocida en el archivo: "${s}".`);
  return fmtISO(dt);
}
function sumarDias(iso, n) {
  const d = new Date(iso + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return fmtISO(d);
}

/* --------------------------- lectura archivo --------------------------- */
async function leerFilas(file) {
  const ext = file.name.toLowerCase().split(".").pop();
  let wb;
  if (ext === "csv") {
    wb = XLSX.read(await file.text(), { type: "string" });          // texto UTF-8 nativo
  } else {
    wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  }
  const ws = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  return filas.map((r) => r.map((c) => (typeof c === "string" ? repararMojibake(c.trim()) : c)));
}

const norm = (s) => String(s).trim().toLowerCase();
const vacia = (fila) => fila.every((c) => String(c).trim() === "");

function parsearInforme(filas) {
  const idxVentas = filas.findIndex((f) => f.length && norm(f[0]) === MARCA_VENTAS);
  if (idxVentas < 0) throw new Error("No se encontró el bloque 'Income from Sells' en el archivo.");

  const cabeceras = filas[idxVentas + 1].map((c) => String(c).trim());
  const col = {};
  cabeceras.forEach((h, i) => (col[norm(h)] = i));
  const faltan = REQUERIDAS.filter((c) => !(c in col));
  if (faltan.length) throw new Error("Faltan columnas en 'Income from Sells': " + faltan.join(", "));

  const operaciones = [];
  for (let k = idxVentas + 2; k < filas.length; k++) {
    const fila = filas[k];
    if (vacia(fila) || norm(fila[0]) === MARCA_OTROS) break;
    operaciones.push(filaAOperacion(fila, col));
  }

  const bloques = [];
  const idxOtros = filas.findIndex((f) => f.length && norm(f[0]) === MARCA_OTROS);
  if (idxOtros >= 0) {
    const cab = filas[idxOtros + 1].map((c) => String(c).trim());
    const datos = [];
    for (let k = idxOtros + 2; k < filas.length; k++) {
      if (vacia(filas[k])) break;
      datos.push(filas[k].map((c) => (typeof c === "string" ? c.trim() : c)));
    }
    bloques.push({ titulo: String(filas[idxOtros][0]).trim(), cabeceras: cab, filas: datos });
  }
  return { operaciones, cabeceras, bloques };
}

function filaAOperacion(fila, col) {
  const v = (n) => fila[col[n]];
  const divisa = String(v("currency") || "USD").toUpperCase();
  return {
    fc: isoDe(v("date acquired")), fv: isoDe(v("date sold")),
    symbol: v("symbol"), nombre: v("security name"), isin: v("isin"), pais: v("country"),
    cantidad: aDecimal(v("quantity")),
    coste: aDecimal(v("cost basis")), ingreso: aDecimal(v("gross proceeds")),
    divisa,
  };
}

/* ----------------------------- tipos BCE ------------------------------ */
async function cargarTipos(iniISO, finISO, status) {
  const url = SDMX.replace("{ini}", sumarDias(iniISO, -MARGEN_INICIO)).replace("{fin}", finISO);
  status("Obteniendo tipos del BCE…");
  let resp;
  try {
    resp = await fetch(url);
  } catch {
    throw new Error("No se pudo conectar con el BCE. Revisa tu conexión; si usas un bloqueador (uBlock, Brave, etc.) desactívalo para esta página.");
  }
  if (!resp.ok) throw new Error("No se pudieron obtener los tipos del BCE (HTTP " + resp.status + ").");
  const text = await resp.text();
  const lineas = text.split(/\r?\n/);
  const hdr = lineas[0].split(",");
  const iD = hdr.indexOf("TIME_PERIOD"), iV = hdr.indexOf("OBS_VALUE");
  if (iD < 0 || iV < 0) throw new Error("Respuesta del BCE en formato inesperado.");
  const mapa = new Map();
  for (let k = 1; k < lineas.length; k++) {
    const f = lineas[k].split(",");
    if (f.length <= Math.max(iD, iV)) continue;
    const d = f[iD].trim(), val = f[iV].trim();
    if (d && val) mapa.set(d, Number(val));
  }
  return mapa;
}

function usdPorEur(mapa, iso) {
  for (let i = 0; i <= MAX_DIAS_ATRAS; i++) {
    const d = sumarDias(iso, -i);
    if (mapa.has(d)) return mapa.get(d);
  }
  throw new Error("No hay tipo de cambio del BCE disponible para " + iso + " ni días previos.");
}

/* --------------------------- conversión ------------------------------- */
function convertir(op, mapa) {
  if (op.divisa === "EUR") {                       // ya en euros: no se toca
    op.costeEur = round2(op.coste); op.ingresoEur = round2(op.ingreso);
    op.pnlEur = round2(op.ingreso - op.coste); op.fxC = op.fxV = "";
    return;
  }
  const fxC = usdPorEur(mapa, op.fc), fxV = usdPorEur(mapa, op.fv);
  const costeU = op.coste / fxC, ingresoU = op.ingreso / fxV;
  op.fxC = fxC; op.fxV = fxV;
  op.costeEur = round2(costeU); op.ingresoEur = round2(ingresoU);
  op.pnlEur = round2(ingresoU - costeU);           // PnL con importes SIN redondear
}

/* ------------------------- generar el xlsx ---------------------------- */
function construirAoA(informe) {
  const cab = informe.cabeceras;
  const ancho = cab.length + COLS_EXTRA.length;
  const aoa = [];

  aoa.push(["Income from Sells"]);
  aoa.push([...cab, ...COLS_EXTRA]);
  for (const op of informe.operaciones) aoa.push(filaOperacion(op, cab));

  const totalPnl = round2(informe.operaciones.reduce((s, o) => s + o.pnlEur, 0));
  aoa.push([]);
  aoa.push(filaTotal("TOTAL PnL EUR", totalPnl, ancho));

  for (const b of informe.bloques) {
    aoa.push([]);
    aoa.push([b.titulo]);
    aoa.push(b.cabeceras);
    for (const f of b.filas) aoa.push(f);
    const iNet = b.cabeceras.findIndex((h) => norm(h) === "net amount");
    const totalNet = round2(b.filas.reduce((s, f) => s + (iNet >= 0 ? aDecimal(f[iNet]) : 0), 0));
    aoa.push(filaTotal("TOTAL Net Amount EUR", totalNet, ancho));
  }
  return aoa;
}

function filaOperacion(op, cab) {
  const base = {
    "date acquired": op.fc, "date sold": op.fv, "symbol": op.symbol,
    "security name": op.nombre, "isin": op.isin, "country": op.pais,
    "quantity": op.cantidad, "cost basis": op.costeEur, "gross proceeds": op.ingresoEur,
    "gross pnl": op.pnlEur, "currency": "EUR",
  };
  const fila = cab.map((h) => (norm(h) in base ? base[norm(h)] : ""));
  fila.push(op.fxC, op.fxV);
  return fila;
}

function filaTotal(etiqueta, total, ancho) {
  const fila = new Array(ancho).fill("");
  fila[ancho - 1] = total;
  fila[ancho - 2] = etiqueta;
  return fila;
}

/* ----------------------------- flujo UI ------------------------------- */
let archivo = null;

document.addEventListener("DOMContentLoaded", () => {
  const dz = document.getElementById("dropzone");
  const input = document.getElementById("file");
  const dzText = document.getElementById("dz-text");
  const btn = document.getElementById("convert");
  const bar = document.getElementById("bar");
  const statusEl = document.getElementById("status");

  const status = (msg, clase = "") => { statusEl.textContent = msg; statusEl.className = "status " + clase; };

  function elegir(f) {
    if (!f) return;
    archivo = f;
    dz.classList.add("ok");
    dzText.innerHTML = "Archivo:<br><small>" + f.name + "</small>";
    btn.disabled = false;
    status("Listo para convertir «" + f.name + "».");
  }

  input.addEventListener("change", () => elegir(input.files[0]));
  dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("drag"); });
  dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
  dz.addEventListener("drop", (e) => {
    e.preventDefault(); dz.classList.remove("drag");
    elegir(e.dataTransfer.files[0]);
  });

  btn.addEventListener("click", async () => {
    if (!archivo) return;
    btn.disabled = true; bar.hidden = false;
    try {
      status("Leyendo archivo…");
      const informe = parsearInforme(await leerFilas(archivo));
      if (!informe.operaciones.length) throw new Error("El archivo no tiene operaciones de venta.");

      const fechas = informe.operaciones.flatMap((o) => [o.fc, o.fv]).sort();
      const mapa = await cargarTipos(fechas[0], fmtISO(new Date()), status);

      status("Convirtiendo " + informe.operaciones.length + " operaciones…");
      informe.operaciones.forEach((op) => convertir(op, mapa));

      status("Generando archivo…");
      const ws = XLSX.utils.aoa_to_sheet(construirAoA(informe));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Revolut EUR");
      const nombre = archivo.name.replace(/\.[^.]+$/, "") + "_EUR.xlsx";
      XLSX.writeFile(wb, nombre);

      status("✓ Descargado: " + nombre, "ok");
    } catch (err) {
      status("✕ " + err.message, "err");
    } finally {
      bar.hidden = true; btn.disabled = false;
    }
  });
});
