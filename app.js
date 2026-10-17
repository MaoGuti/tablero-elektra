/* Tablero de Elektra: lee en el navegador el Excel de control de cargas y dibuja el tablero.
   Nada se envía a ningún servidor. */
"use strict";

const DIA = 864e5;
const PRECIO_CORRIENTE = 16377; // Galón de gasolina corriente en Bogotá, octubre de 2026 (precio regulado, CREG).
const UMBRAL_SALIDA = 120; // km entre dos cargas seguidas a partir de los cuales el tramo cuenta como salida de Bogotá.
const MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

const $ = (sel) => document.querySelector(sel);
const sum = (arr, f) => arr.reduce((s, x) => s + (f(x) || 0), 0);
const median = (arr) => {
  if (!arr.length) return null;
  const a = [...arr].sort((x, y) => x - y);
  const m = a.length >> 1;
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
};

const fmtNum = (n, dec = 0) => n.toLocaleString("es-CO", { minimumFractionDigits: dec, maximumFractionDigits: dec });
const fmtPesos = (n) => "$" + fmtNum(Math.round(n));
const fmtPesosCorto = (n) => {
  const a = Math.abs(n);
  if (a >= 1e6) return `$${fmtNum(n / 1e6, a >= 1e7 ? 0 : 1)} millones`;
  if (a >= 1e3) return `$${fmtNum(n / 1e3)} mil`;
  return fmtPesos(n);
};
const fmtEjePesos = (v) => (v >= 1e6 ? `$${fmtNum(v / 1e6, 1)} M` : v >= 1e3 ? `$${fmtNum(v / 1e3)} mil` : `$${fmtNum(v)}`);
const fmtPct = (x) => Math.round(x * 100) + " %";
const fmtFecha = (d) => `${d.getUTCDate()} ${MESES[d.getUTCMonth()]}`;
const claveMes = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
const nombreMes = (clave) => MESES[Number(clave.slice(5)) - 1];
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- Lectura del Excel ---------- */

const norm = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();

function num(v) {
  if (typeof v === "number" && isFinite(v)) return v;
  if (typeof v === "string") {
    const t = v.replace(/[$\s]/g, "").replace(",", ".");
    if (t !== "" && isFinite(Number(t))) return Number(t);
  }
  return null;
}

// Las fechas de Excel llegan como número de serie en hora local; se guardan como UTC "de pared".
function fecha(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number" && v > 20000) return new Date(Math.round((v - 25569) * DIA));
  if (typeof v === "string" && v.trim()) {
    const d = new Date(v);
    return isNaN(d) ? null : d;
  }
  return null;
}

const fraccion = (v) => (v != null && v > 1.0001 ? v / 100 : v);

// columnas: { clave: ["Encabezado", "Alternativa"] }; las claves en `opcionales` pueden faltar.
function leerHoja(libro, hoja, columnas, opcionales = []) {
  const nombre = libro.SheetNames.find((n) => norm(n) === norm(hoja));
  if (!nombre) throw new Error(`No encontré la hoja "${hoja}".`);
  const filas = XLSX.utils.sheet_to_json(libro.Sheets[nombre], { header: 1, raw: true, defval: null });
  const opciones = Object.entries(columnas).map(([k, v]) => [k, (Array.isArray(v) ? v : [v]).map(norm)]);
  const primera = opciones[0][1];
  const iEnc = filas.findIndex((f) => f && f.some((c) => primera.includes(norm(c))));
  if (iEnc < 0) throw new Error(`En la hoja "${hoja}" no encontré la fila de encabezados.`);
  const enc = filas[iEnc].map(norm);
  const pos = {};
  for (const [k, ops] of opciones) {
    const i = enc.findIndex((e) => ops.includes(e));
    if (i >= 0) pos[k] = i;
    else if (!opcionales.includes(k)) {
      const v = columnas[k];
      throw new Error(`En la hoja "${hoja}" falta la columna "${Array.isArray(v) ? v[0] : v}".`);
    }
  }
  return filas.slice(iEnc + 1).map((f) => {
    const o = {};
    for (const k of Object.keys(columnas)) o[k] = k in pos ? f[pos[k]] ?? null : null;
    return o;
  });
}

function leerExcel(buffer) {
  const libro = XLSX.read(buffer, { type: "array" });

  const cargas = leerHoja(
    libro, "Cargas",
    {
      enchufado: "Enchufado",
      odoKm: "Odómetro km", odoMi: "Odómetro mi",
      socIni: "% al enchufar", socFin: "% al desenchufar",
      kwh: ["kWh exactos", "kWh cargados"], kwhApp: ["kWh app (redondeado)", "kWh app"],
      costoApp: "Costo app", tarifaApp: "Tarifa app $/kWh",
    },
    ["odoKm", "odoMi", "kwh", "kwhApp", "costoApp", "tarifaApp"]
  ).map((r) => {
    const costoApp = num(r.costoApp), tarifaApp = num(r.tarifaApp);
    const kwh = num(r.kwh) ?? (costoApp != null && tarifaApp ? costoApp / tarifaApp : num(r.kwhApp));
    const odo = num(r.odoKm) ?? (num(r.odoMi) != null ? num(r.odoMi) * 1.609344 : null);
    return { fecha: fecha(r.enchufado), odo, socIni: fraccion(num(r.socIni)), socFin: fraccion(num(r.socFin)), kwh: kwh ?? 0 };
  }).filter((r) => r.fecha && r.odo != null && r.socIni != null && r.socFin != null);

  const tarifas = leerHoja(
    libro, "Tarifas",
    {
      desde: "Periodo desde", hasta: "Periodo hasta",
      precio: "Precio final $/kWh", tarifa: "Tarifa $/kWh", contribucion: "Contribución",
      consumo: ["Consumo Qatar kWh", "Consumo casa kWh", "Consumo kWh"],
    },
    ["precio", "tarifa", "contribucion"]
  ).map((r) => {
    // Se guarda la tarifa sin recargo y la contribución aparte, para poder incluirla o quitarla.
    const precioFinal = num(r.precio);
    let tarifa = num(r.tarifa), contrib = fraccion(num(r.contribucion));
    if (tarifa == null && precioFinal != null) tarifa = precioFinal / (1 + (contrib ?? 0));
    if (contrib == null && tarifa && precioFinal != null) contrib = Math.max(0, precioFinal / tarifa - 1);
    return { desde: fecha(r.desde), hasta: fecha(r.hasta), tarifa, contrib: contrib ?? 0, consumo: num(r.consumo) };
  }).filter((r) => r.desde && r.hasta && r.tarifa != null && r.consumo != null);

  const gasolina = leerHoja(
    libro, "Gasolina",
    {
      fecha: "Fecha y hora", galones: "Galones", precioGal: "Precio por galón",
      valor: "Valor", neto: "Valor neto", descuento: ["Descuento TC", "Descuento"], intereses: "Intereses TC",
    },
    ["precioGal", "neto", "descuento", "intereses"]
  ).map((r) => {
    const valor = num(r.valor);
    const neto = num(r.neto) ?? (valor != null ? valor - (num(r.descuento) ?? 0) + (num(r.intereses) ?? 0) : null);
    return { fecha: fecha(r.fecha), galones: num(r.galones), precioGal: num(r.precioGal), neto };
  }).filter((r) => r.fecha && r.galones && r.neto != null);

  if (cargas.length < 2) throw new Error("La hoja Cargas necesita al menos dos cargas.");
  if (!tarifas.length) throw new Error("La hoja Tarifas no tiene periodos con tarifa.");
  return { cargas, tarifas, gasolina };
}

/* ---------- Plantilla de Excel para llenar (se arma en el navegador) ---------- */

function construirPlantilla() {
  // Número de serie de Excel para una fecha y hora (sin zona horaria).
  const serie = (a, m, d, h = 0, mi = 0) => Date.UTC(a, m - 1, d, h, mi) / DIA + 25569;
  const hoja = (filas, anchos, formatos = {}) => {
    const ws = XLSX.utils.aoa_to_sheet(filas);
    ws["!cols"] = anchos.map((wch) => ({ wch }));
    const rango = XLSX.utils.decode_range(ws["!ref"]);
    for (let r = 1; r <= rango.e.r; r++) {
      for (const [col, z] of Object.entries(formatos)) {
        const celda = ws[XLSX.utils.encode_cell({ r, c: Number(col) })];
        if (celda && typeof celda.v === "number") celda.z = z;
      }
    }
    return ws;
  };
  const instrucciones = [
    ["Plantilla del tablero del vehículo a analizar"],
    [""],
    ["Cómo usarla"],
    ["1. Llena las hojas Cargas, Tarifas y Gasolina con los datos de tu carro. Las filas que trae son de ejemplo: bórralas o escribe encima."],
    ["2. No cambies los nombres de las hojas ni los títulos de las columnas."],
    ["3. Guarda el archivo y súbelo en https://tablero-elektra.vercel.app con el botón «Escoger mi Excel». Se lee en tu navegador y no se sube a ningún servidor."],
    [""],
    ["Hoja Cargas: una fila por carga (sale del historial de carga de la app del carro)"],
    ["Enchufado: fecha y hora en que conectaste el carro."],
    ["Odómetro km: los kilómetros del carro al enchufar. Si la app da millas, cambia el título a «Odómetro mi»."],
    ["% al enchufar y % al desenchufar: la batería al conectar y al desconectar."],
    ["kWh cargados: la energía que entró en esa carga."],
    [""],
    ["Hoja Tarifas: una fila por factura de energía de la casa"],
    ["Periodo desde y Periodo hasta: el periodo de consumo que trae la factura."],
    ["Tarifa $/kWh: el precio del kWh de ese periodo, sin contribución."],
    ["Contribución: el recargo del estrato (20 % en los estratos 5 y 6 de Bogotá; 0 % en los demás)."],
    ["Consumo casa kWh: el consumo total de la factura, incluido el carro."],
    [""],
    ["Hoja Gasolina: una fila por tanqueada"],
    ["Fecha y hora, Galones, Precio por galón y Valor: salen de la factura de la estación."],
    ["Descuento: lo que te descontaron en esa tanqueada (por ejemplo, la tarjeta). Si no hubo, déjalo vacío."],
    [""],
    ["Importante: el tablero separa la ciudad de las salidas con el odómetro de cada carga. Si falta el odómetro de una carga, esa carga no se cuenta."],
  ];
  const libro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(libro, hoja(instrucciones, [110]), "Instrucciones");
  XLSX.utils.book_append_sheet(libro, hoja([
    ["Enchufado", "Odómetro km", "% al enchufar", "% al desenchufar", "kWh cargados"],
    [serie(2026, 5, 4, 20, 10), 1250, 0.18, 0.8, 9.3],
    [serie(2026, 5, 6, 19, 45), 1296, 0.2, 0.8, 9.0],
    [serie(2026, 5, 9, 21, 5), 1338, 0.25, 0.8, 8.3],
    [serie(2026, 5, 12, 20, 30), 1612, 0.03, 0.8, 11.6],
    [serie(2026, 5, 14, 19, 50), 1655, 0.21, 0.8, 8.9],
  ], [18, 13, 14, 16, 14], { 0: "yyyy-mm-dd hh:mm", 2: "0%", 3: "0%", 4: "0.0" }), "Cargas");
  XLSX.utils.book_append_sheet(libro, hoja([
    ["Periodo desde", "Periodo hasta", "Tarifa $/kWh", "Contribución", "Consumo casa kWh"],
    [serie(2026, 4, 20), serie(2026, 5, 19), 870.5, 0.2, 290],
  ], [14, 14, 13, 13, 17], { 0: "yyyy-mm-dd", 1: "yyyy-mm-dd", 2: "#,##0.00", 3: "0%", 4: "0" }), "Tarifas");
  XLSX.utils.book_append_sheet(libro, hoja([
    ["Fecha y hora", "Galones", "Precio por galón", "Valor", "Descuento"],
    [serie(2026, 5, 11, 7, 30), 4.5, 16377, 73697, 0],
  ], [18, 10, 16, 12, 12], { 0: "yyyy-mm-dd hh:mm", 1: "0.000", 2: "#,##0", 3: "#,##0", 4: "#,##0" }), "Gasolina");
  return libro;
}

/* ---------- Datos de ejemplo (inventados, para mostrar el tablero sin datos reales) ---------- */

function datosEjemplo() {
  let semilla = 20261014;
  const rnd = () => (semilla = (semilla * 48271) % 2147483647) / 2147483647;
  const precios = [905, 1010, 995, 990, 1012, 1048, 1061, 1035, 1044];
  const base = [214, 206, 221, 209, 218, 212, 226, 215, 219];
  const tarifas = precios.map((precio, m) => ({
    desde: new Date(Date.UTC(2026, m, 1)), hasta: new Date(Date.UTC(2026, m + 1, 0)),
    tarifa: Math.round((precio / 1.2) * 100) / 100, contrib: 0.2, casaBase: base[m],
  }));
  const KWH_POR_BATERIA = 15.2, CONSUMO = 0.205;
  const cargas = [], gasolina = [];
  let odo = 42, soc = 0.8, dia = Date.UTC(2026, 4, 10);
  while (dia < Date.UTC(2026, 8, 29)) {
    dia += (1 + Math.floor(rnd() * 3)) * DIA;
    let km;
    if (rnd() < 0.08) {
      km = 150 + Math.round(rnd() * 120);
      const kmBateria = ((soc - 0.03) * KWH_POR_BATERIA) / CONSUMO;
      soc = 0.03;
      const galones = (km - kmBateria) / (52 + rnd() * 8);
      const precioGal = 21900 + Math.round(rnd() * 7) * 100;
      const valor = Math.round(galones * precioGal);
      gasolina.push({ fecha: new Date(dia - DIA / 2), galones: Math.round(galones * 1000) / 1000, precioGal, neto: valor - (valor >= 100000 ? 20000 : 0) });
    } else {
      km = 18 + Math.round(rnd() * 34);
      soc = Math.max(0.02, soc - (km * CONSUMO) / KWH_POR_BATERIA);
    }
    odo += km;
    const hora = rnd() < 0.7 ? 18 + Math.floor(rnd() * 4) : 7 + Math.floor(rnd() * 10);
    const socFin = rnd() < 0.85 ? 0.8 : 1;
    const socIni = Math.round(soc * 100) / 100;
    cargas.push({ fecha: new Date(dia + hora * 36e5 + Math.floor(rnd() * 60) * 6e4), odo, socIni, socFin, kwh: (socFin - socIni) * KWH_POR_BATERIA });
    soc = socFin;
  }
  for (const t of tarifas) {
    const kwhCarro = sum(cargas.filter((c) => c.fecha >= t.desde && c.fecha < +t.hasta + DIA), (c) => c.kwh);
    t.consumo = Math.round(t.casaBase + kwhCarro);
  }
  return { cargas, tarifas, gasolina };
}

/* ---------- Cálculos ---------- */

function analizar(datos, ajustes) {
  const cargas = [...datos.cargas].sort((a, b) => a.fecha - b.fecha);
  const tarifas = [...datos.tarifas].sort((a, b) => a.desde - b.desde);
  const gasolina = [...datos.gasolina].sort((a, b) => a.fecha - b.fecha);

  const conContribucion = ajustes.contribucion === true;
  for (const t of tarifas) t.precio = t.tarifa * (1 + (conContribucion ? t.contrib : 0));
  const contribucion = median(tarifas.filter((t) => t.contrib > 0).map((t) => t.contrib));
  const periodoDe = (f) => tarifas.find((t) => f >= t.desde && f < +t.hasta + DIA);
  for (const c of cargas) {
    const t = periodoDe(c.fecha);
    c.costo = t ? c.kwh * t.precio : null;
  }
  const conEnergia = cargas.filter((c) => c.kwh > 0);
  const conPrecio = conEnergia.filter((c) => c.costo != null);
  const kwhTotal = sum(conEnergia, (c) => c.kwh);
  const costoEnergia = sum(conPrecio, (c) => c.costo);
  const precioKwh = conPrecio.length ? costoEnergia / sum(conPrecio, (c) => c.kwh) : null;
  const sinTarifa = conEnergia.length - conPrecio.length;

  // Consumo: batería gastada entre dos cargas seguidas, en tramos urbanos sin tanqueadas.
  const tanqueos = gasolina.map((g) => +g.fecha);
  const tramos = [];
  for (let i = 1; i < conEnergia.length; i++) {
    const a = conEnergia[i - 1], b = conEnergia[i];
    const km = b.odo - a.odo;
    const usada = a.socFin - b.socIni;
    const agregada = b.socFin - b.socIni;
    const tanqueo = tanqueos.some((t) => t > +a.fecha && t < +b.fecha);
    if (km >= 5 && km <= 90 && usada >= 0.1 && agregada >= 0.1 && !tanqueo) {
      const kwh = usada * (b.kwh / agregada);
      tramos.push(kwh / km);
    }
  }
  const consumo = median(tramos);
  const kwhPorBateria = median(conEnergia.filter((c) => c.socFin - c.socIni >= 0.3).map((c) => c.kwh / (c.socFin - c.socIni)));
  const kmPorCarga = consumo && kwhPorBateria ? (0.78 * kwhPorBateria) / consumo : null;

  const conPrecioGal = gasolina.filter((g) => g.precioGal);
  const precioGalDatos = conPrecioGal.length ? conPrecioGal[conPrecioGal.length - 1].precioGal : 22000;
  const tipo = ajustes.tipo === "extra" ? "extra" : "corriente";
  const precioGal = ajustes.precioGal || (tipo === "extra" ? precioGalDatos : PRECIO_CORRIENTE);
  const kmGal = ajustes.kmGalCiudad || 30;
  const kmGalCarretera = ajustes.kmGalCarretera || 40;
  const costoKmElec = precioKwh != null && consumo ? precioKwh * consumo : null;
  const costoKmGas = precioGal / kmGal; // carro de referencia en ciudad
  const costoKmGasCarretera = precioGal / kmGalCarretera;

  const ini = cargas[0], fin = cargas[cargas.length - 1];
  const kmVentana = fin.odo - ini.odo;
  const netoVentana = sum(gasolina.filter((g) => g.fecha >= ini.fecha && g.fecha <= fin.fecha), (g) => g.neto);
  const kmElec = consumo && kmVentana > 0 ? Math.min(kwhTotal / consumo, kmVentana) : null;
  const pctElec = kmElec != null ? kmElec / kmVentana : null;
  const ahorro = kmElec != null && costoKmElec != null ? kmElec * (costoKmGas - costoKmElec) : null;
  const costoKmTotal = kmVentana > 0 && !sinTarifa ? (costoEnergia + netoVentana) / kmVentana : null;

  // Bogotá frente a salidas: cada tramo entre dos cargas seguidas. Los km con batería salen de la
  // batería gastada en el tramo; el resto del tramo se hizo con gasolina.
  const segmentos = [];
  for (let i = 1; i < conEnergia.length; i++) {
    const a0 = conEnergia[i - 1], b0 = conEnergia[i];
    const km = b0.odo - a0.odo;
    if (km <= 0) continue;
    const kwh = Math.max(0, a0.socFin - b0.socIni) * (kwhPorBateria || 0);
    const kmElec = consumo ? Math.min(km, kwh / consumo) : 0;
    segmentos.push({ desde: a0.fecha, hasta: b0.fecha, km, kwh, kmElec, kmGas: km - kmElec, salida: km >= UMBRAL_SALIDA });
  }
  const resumir = (segs) => ({
    n: segs.length, km: sum(segs, (s) => s.km), kwh: sum(segs, (s) => s.kwh),
    kmElec: sum(segs, (s) => s.kmElec), kmGas: sum(segs, (s) => s.kmGas),
  });
  const bogota = resumir(segmentos.filter((s) => !s.salida));
  const salidas = resumir(segmentos.filter((s) => s.salida));
  const listaSalidas = segmentos.filter((s) => s.salida);
  const galVentana = sum(gasolina.filter((g) => g.fecha >= ini.fecha && g.fecha <= fin.fecha), (g) => g.galones);
  const kmGasTotal = bogota.kmGas + salidas.kmGas;
  const costoKmGasElektra = kmGasTotal > 20 && netoVentana > 0 ? netoVentana / kmGasTotal : null;
  const kmPorGalon = kmGasTotal > 20 && galVentana > 0 ? kmGasTotal / galVentana : null;
  for (const z of [bogota, salidas]) {
    z.costo = precioKwh != null && (costoKmGasElektra != null || z.kmGas === 0)
      ? z.kwh * precioKwh + z.kmGas * (costoKmGasElektra || 0) : null;
    z.costoKm = z.costo != null && z.km > 0 ? z.costo / z.km : null;
    z.pctElec = z.km > 0 ? z.kmElec / z.km : null;
  }

  const ultimoMesCargas = claveMes(fin.fecha);
  const claves = [...new Set([...cargas.map((c) => claveMes(c.fecha)), ...gasolina.map((g) => claveMes(g.fecha))])].sort();
  const meses = claves.map((k) => {
    const cs = conEnergia.filter((c) => claveMes(c.fecha) === k);
    const gs = gasolina.filter((g) => claveMes(g.fecha) === k);
    return {
      clave: k, nombre: nombreMes(k), cargas: cs.length, kwh: sum(cs, (c) => c.kwh),
      energia: sum(cs, (c) => c.costo), sinTarifa: cs.some((c) => c.costo == null),
      galones: sum(gs, (g) => g.galones), gasolina: sum(gs, (g) => g.neto), parcial: k > ultimoMesCargas,
    };
  });

  const casa = tarifas.map((t) => {
    const kwhCarro = sum(conEnergia.filter((c) => c.fecha >= t.desde && c.fecha < +t.hasta + DIA), (c) => c.kwh);
    const medio = new Date((+t.desde + +t.hasta) / 2);
    const carro = Math.min(kwhCarro, t.consumo);
    return { nombre: MESES[medio.getUTCMonth()], desde: t.desde, hasta: t.hasta, precio: t.precio, consumo: t.consumo, carro, resto: t.consumo - carro };
  });
  const conCarro = casa.filter((p) => p.carro > 0);
  const pctCasa = conCarro.length ? sum(conCarro, (p) => p.carro) / sum(conCarro, (p) => p.consumo) : null;
  const sinCarro = casa.filter((p) => p.carro === 0 && p.hasta < ini.fecha);
  const promedioAntes = sinCarro.length ? sum(sinCarro, (p) => p.consumo) / sinCarro.length : null;

  const bins = Array.from({ length: 10 }, (_, i) => ({ desde: i * 10, n: 0 }));
  for (const c of conEnergia) bins[Math.min(9, Math.floor((c.socIni * 100 + 1e-9) / 10))].n++;
  let ultimoBin = 9;
  while (ultimoBin > 5 && bins[ultimoBin].n === 0) ultimoBin--;
  const en80 = conEnergia.filter((c) => Math.abs(c.socFin - 0.8) <= 0.011).length;
  const deNoche = conEnergia.filter((c) => { const h = c.fecha.getUTCHours(); return h >= 18 || h < 6; }).length;
  const bajo20 = conEnergia.filter((c) => c.socIni < 0.2).length;

  return {
    nCargas: conEnergia.length, kwhTotal, precioKwh, sinTarifa, consumo, kmPorCarga, conContribucion, contribucion,
    tipo, precioGal, kmGal, kmGalCarretera, costoKmElec, costoKmGas, costoKmGasCarretera,
    ini: ini.fecha, fin: fin.fecha, kmVentana, kmElec, pctElec, ahorro, costoKmTotal,
    bogota, salidas, listaSalidas, kmPorGalon, costoKmGasElektra,
    meses, casa, pctCasa, promedioAntes, bins: bins.slice(0, ultimoBin + 1), en80, deNoche, bajo20,
  };
}

/* ---------- Gráficas (SVG a mano) ---------- */

function escala(max, enteros) {
  if (!(max > 0)) return { tope: enteros ? 4 : 1, paso: enteros ? 1 : 0.25 };
  const bruto = max / 4;
  const mag = Math.pow(10, Math.floor(Math.log10(bruto)));
  const pasos = enteros ? [1, 2, 5, 10] : [1, 2, 2.5, 5, 10];
  let paso = pasos.find((p) => p * mag >= bruto) * mag;
  if (enteros) paso = Math.max(1, Math.round(paso));
  return { tope: Math.ceil(max / paso) * paso, paso };
}

function topeRedondeado(x, y, w, h, r) {
  r = Math.min(r, w / 2, h);
  return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`;
}

const graficas = [];

// series: [{ nombre, clase, valores }]; tip(i) devuelve el HTML del tooltip.
function columnas(el, { categorias, series, fmtEje, enteros = false, etiquetas = [], tip, titulo }) {
  const dibujar = () => {
    const W = Math.max(280, el.clientWidth), H = 250;
    const m = { t: 22, r: 6, b: 28, l: 56 };
    const ih = H - m.t - m.b, iw = W - m.l - m.r;
    const totales = categorias.map((_, i) => sum(series, (s) => s.valores[i]));
    const { tope, paso } = escala(Math.max(...totales, 0), enteros);
    const y = (v) => m.t + ih - (v / tope) * ih;
    const banda = iw / categorias.length;
    const ancho = Math.min(24, banda * 0.6);
    let svg = `<svg width="${W}" height="${H}" role="img" aria-label="${esc(titulo)}">`;
    for (let v = 0; v <= tope + 1e-9; v += paso) {
      svg += `<line class="${v === 0 ? "base" : "grid"}" x1="${m.l}" x2="${W - m.r}" y1="${y(v)}" y2="${y(v)}"/>`;
      svg += `<text class="tick" x="${m.l - 8}" y="${y(v)}" text-anchor="end" dominant-baseline="middle">${esc(fmtEje(v))}</text>`;
    }
    categorias.forEach((cat, i) => {
      const x = m.l + banda * i + (banda - ancho) / 2;
      const segs = series.map((s) => ({ s, v: s.valores[i] || 0 })).filter((d) => d.v > 0);
      let acum = 0;
      segs.forEach((d, j) => {
        const abajo = y(acum) - (j > 0 ? 2 : 0);
        acum += d.v;
        const arriba = y(acum);
        const h = Math.max(0, abajo - arriba);
        if (h <= 0) return;
        svg += j === segs.length - 1
          ? `<path class="${d.s.clase}" d="${topeRedondeado(x, arriba, ancho, h, 4)}"/>`
          : `<rect class="${d.s.clase}" x="${x}" y="${arriba}" width="${ancho}" height="${h}"/>`;
      });
      if (etiquetas[i]) svg += `<text class="cap" x="${x + ancho / 2}" y="${y(totales[i]) - 7}" text-anchor="middle">${esc(etiquetas[i])}</text>`;
      svg += `<text class="tick" x="${x + ancho / 2}" y="${H - m.b + 18}" text-anchor="middle">${esc(cat)}</text>`;
      svg += `<rect class="hit" data-i="${i}" tabindex="0" x="${m.l + banda * i}" y="${m.t}" width="${banda}" height="${ih}"><title>${esc(cat)}</title></rect>`;
    });
    svg += "</svg>";
    el.innerHTML = svg;
    conectarTooltip(el, tip);
  };
  graficas.push({ el, dibujar });
  dibujar();
}

function finRedondeado(x, y, w, h, r) {
  r = Math.min(r, h / 2, w);
  return `M${x},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h - r}Q${x + w},${y + h} ${x + w - r},${y + h}H${x}Z`;
}

// Barras horizontales apiladas: una fila por grupo.
function barras(el, { filas, series, fmtEje, etiquetas = [], tip, titulo }) {
  const dibujar = () => {
    const W = Math.max(280, el.clientWidth), alto = 24, sep = 24;
    const m = { t: 10, r: W < 480 ? 92 : 120, b: 28, l: W < 480 ? 64 : 76 };
    const H = m.t + filas.length * (alto + sep) - sep + m.b + 10;
    const iw = W - m.l - m.r, yBase = H - m.b;
    const totales = filas.map((_, i) => sum(series, (s) => s.valores[i]));
    const { tope, paso } = escala(Math.max(...totales, 0), false);
    const x = (v) => m.l + (v / tope) * iw;
    let svg = `<svg width="${W}" height="${H}" role="img" aria-label="${esc(titulo)}">`;
    for (let v = 0; v <= tope + 1e-9; v += paso) {
      svg += `<line class="${v === 0 ? "base" : "grid"}" x1="${x(v)}" x2="${x(v)}" y1="${m.t - 4}" y2="${yBase}"/>`;
      svg += `<text class="tick" x="${x(v)}" y="${yBase + 18}" text-anchor="middle">${esc(fmtEje(v))}</text>`;
    }
    filas.forEach((fila, i) => {
      const y = m.t + i * (alto + sep);
      svg += `<text class="tick fila" x="${m.l - 10}" y="${y + alto / 2}" text-anchor="end" dominant-baseline="middle">${esc(fila)}</text>`;
      const segs = series.map((s) => ({ s, v: s.valores[i] || 0 })).filter((d) => d.v > 0);
      let acum = 0;
      segs.forEach((d, j) => {
        const izq = x(acum) + (j > 0 ? 2 : 0);
        acum += d.v;
        const w = Math.max(0, x(acum) - izq);
        if (w <= 0) return;
        svg += j === segs.length - 1
          ? `<path class="${d.s.clase}" d="${finRedondeado(izq, y, w, alto, 4)}"/>`
          : `<rect class="${d.s.clase}" x="${izq}" y="${y}" width="${w}" height="${alto}"/>`;
      });
      if (etiquetas[i]) svg += `<text class="cap" x="${x(totales[i]) + 8}" y="${y + alto / 2}" dominant-baseline="middle">${esc(etiquetas[i])}</text>`;
      svg += `<rect class="hit" data-i="${i}" tabindex="0" x="${m.l}" y="${y - sep / 2}" width="${iw}" height="${alto + sep}"><title>${esc(fila)}</title></rect>`;
    });
    svg += "</svg>";
    el.innerHTML = svg;
    conectarTooltip(el, tip);
  };
  graficas.push({ el, dibujar });
  dibujar();
}

function conectarTooltip(el, tip) {
  const caja = $("#tip");
  const mostrar = (hit, cx, cy) => {
    el.querySelectorAll(".hit.on").forEach((h) => h.classList.remove("on"));
    hit.classList.add("on");
    caja.innerHTML = tip(Number(hit.dataset.i));
    caja.hidden = false;
    const r = caja.getBoundingClientRect();
    let left = cx + 14, top = cy - r.height - 10;
    if (left + r.width > window.innerWidth - 8) left = cx - r.width - 14;
    if (top < 8) top = cy + 16;
    caja.style.left = Math.max(8, left) + "px";
    caja.style.top = top + "px";
  };
  const ocultar = () => {
    caja.hidden = true;
    el.querySelectorAll(".hit.on").forEach((h) => h.classList.remove("on"));
  };
  el.querySelectorAll(".hit").forEach((hit) => {
    hit.addEventListener("pointermove", (e) => mostrar(hit, e.clientX, e.clientY));
    hit.addEventListener("pointerleave", ocultar);
    hit.addEventListener("focus", () => { const b = hit.getBoundingClientRect(); mostrar(hit, b.left + b.width / 2, b.top + 40); });
    hit.addEventListener("blur", ocultar);
  });
}

const filaTip = (clase, nombre, valor) => `<div class="row"><span>${clase ? `<span class="sw ${clase}"></span>` : ""}${esc(nombre)}</span><span>${esc(valor)}</span></div>`;

/* ---------- Pintar el tablero ---------- */

// El tablero arranca sin la contribución del estrato; ?contribucion en el enlace la marca.
const estado = {
  datos: null, fuente: "",
  ajustes: { kmGalCiudad: 30, kmGalCarretera: 40, tipo: "corriente", precioGal: null, contribucion: new URLSearchParams(location.search).has("contribucion") },
};

function pintar() {
  const a = analizar(estado.datos, estado.ajustes);
  $("#fuente").textContent = estado.fuente;
  if (!estado.ajustes.precioGal) $("#precioGal").value = Math.round(a.precioGal);
  $("#contribucion").disabled = !a.contribucion;
  $("#contribucion").checked = a.conContribucion && !!a.contribucion;
  $("#texto-contribucion").textContent = a.contribucion
    ? `Incluir la contribución del ${fmtPct(a.contribucion)} que trae la factura`
    : "La factura no trae contribución";

  const veces = a.costoKmElec ? a.costoKmGas / a.costoKmElec : null;
  $("#hero").innerHTML = a.costoKmElec
    ? `<p class="hero-label">En eléctrico, cada kilómetro de Elektra cuesta</p>
       <p class="hero-num">${fmtPesos(a.costoKmElec)}</p>
       <p class="hero-sub">Un carro a gasolina ${a.tipo} que en ciudad hace ${fmtNum(a.kmGal)} km por galón gasta ${fmtPesos(a.costoKmGas)} por kilómetro${veces > 1 ? `: <strong>${fmtNum(veces, 1)} veces más</strong>` : ""}.</p>`
    : `<p class="hero-label">Todavía no se puede calcular el costo por kilómetro: faltan tarifas o tramos entre cargas.</p>`;

  const tile = (label, value, note) => `<div class="tile"><p class="tile-label">${label}</p><p class="tile-value">${value}</p><p class="tile-note">${note}</p></div>`;
  const tiles = [];
  if (a.precioKwh) tiles.push(tile(
    "Precio del kWh", fmtPesos(a.precioKwh),
    a.contribucion ? (a.conContribucion ? `Factura de la casa, con la contribución del ${fmtPct(a.contribucion)}` : "Tarifa de la factura de la casa, sin contribución") : "Tarifa de la factura de la casa"));
  if (a.consumo) tiles.push(tile("Consumo eléctrico", `${fmtNum(a.consumo * 100, 1)} <span class="tile-unit">kWh/100 km</span>`, a.kmPorCarga ? `≈ ${fmtNum(a.kmPorCarga)} km por carga, de casi vacía a 80 %` : "Medido entre cargas"));
  if (a.pctElec != null) tiles.push(tile("Km con batería (aprox.)", fmtPct(a.pctElec), `De ${fmtNum(a.kmVentana)} km, del ${fmtFecha(a.ini)} al ${fmtFecha(a.fin)}`));
  if (a.ahorro != null) tiles.push(tile("Ahorro frente a gasolina (aprox.)", fmtPesosCorto(a.ahorro), `Frente al carro de referencia (${a.tipo}, en ciudad), en los km que hizo con batería`));
  if (a.costoKmTotal != null) tiles.push(tile("Todo incluido (aprox.)", `${fmtPesos(a.costoKmTotal)} <span class="tile-unit">por km</span>`, "Energía y gasolina, en el mismo periodo"));
  $("#tiles").innerHTML = tiles.join("");

  graficas.length = 0;

  // Bogotá frente a salidas
  const bo = a.bogota, sa = a.salidas;
  const veces2 = (z, ref) => !z.costoKm ? ""
    : ref >= z.costoKm * 1.05 ? `: ${fmtNum(ref / z.costoKm, 1)} veces más`
    : ref >= z.costoKm * 0.95 ? ": casi lo mismo que Elektra"
    : ": aquí Elektra no ahorra";
  const zona = (titulo, z, lineas) => `<div class="zona"><p class="tile-label">${titulo}</p>
    <p class="zona-valor">${z.costoKm != null ? `${fmtPesos(z.costoKm)} <span class="tile-unit">por km</span>` : "—"}</p>
    ${lineas.map((l) => `<p class="tile-note">${l}</p>`).join("")}</div>`;
  $("#zonas").innerHTML =
    zona("En Bogotá", bo, [
      `${fmtNum(bo.km)} km · ${bo.pctElec != null ? fmtPct(bo.pctElec) : "—"} con batería`,
      `Un carro a gasolina ${a.tipo} de ${fmtNum(a.kmGal)} km por galón en ciudad gastaría ${fmtPesos(a.costoKmGas)} por km${veces2(bo, a.costoKmGas)}.`,
    ]) +
    (sa.n
      ? zona("En las salidas de Bogotá", sa, [
          `${sa.n} ${sa.n === 1 ? "salida" : "salidas"} · ${fmtNum(sa.km)} km · ${fmtPct(sa.pctElec)} con batería`,
          `Un carro a gasolina ${a.tipo} de ${fmtNum(a.kmGalCarretera)} km por galón en carretera gastaría ${fmtPesos(a.costoKmGasCarretera)} por km${veces2(sa, a.costoKmGasCarretera)}.`,
          a.kmPorGalon ? `Con gasolina, Elektra hace unos ${fmtNum(a.kmPorGalon)} km por galón.` : "",
        ].filter(Boolean))
      : `<div class="zona"><p class="tile-label">En las salidas de Bogotá</p><p class="tile-note">No hay tramos de ${UMBRAL_SALIDA} km o más entre dos cargas en este periodo.</p></div>`);
  $("#sub-zonas").textContent = sa.n
    ? `En Bogotá, el ${fmtPct(bo.pctElec)} de los km van con batería. En las salidas, la gasolina hace el ${fmtPct(1 - sa.pctElec)}.`
    : `En Bogotá, el ${fmtPct(bo.pctElec)} de los km van con batería.`;
  barras($("#c-zonas"), {
    titulo: "Km con batería y con gasolina, en Bogotá y en las salidas",
    filas: ["Bogotá", "Salidas"],
    series: [
      { nombre: "Con batería", clase: "s1", valores: [bo.kmElec, sa.kmElec] },
      { nombre: "Con gasolina", clase: "s2", valores: [bo.kmGas, sa.kmGas] },
    ],
    etiquetas: [bo, sa].map((z) => (z.km > 0 ? `${fmtPct(z.pctElec)} batería` : "")),
    fmtEje: (v) => `${fmtNum(v)} km`,
    tip: (i) => {
      const z = [bo, sa][i];
      return `<b>${i === 0 ? "En Bogotá" : "En las salidas"}</b>` +
        filaTip("s1", "Con batería", `${fmtNum(z.kmElec)} km`) +
        filaTip("s2", "Con gasolina", `${fmtNum(z.kmGas)} km`) +
        filaTip("", "Total", `${fmtNum(z.km)} km`) +
        (z.costoKm != null ? filaTip("", "Costo por km", fmtPesos(z.costoKm)) : "");
    },
  });
  $("#lista-salidas").innerHTML = a.listaSalidas.length
    ? `<p class="mini-title">Las salidas que encontré</p><ul class="salidas">${a.listaSalidas.map((s) =>
        `<li><span>${fmtFecha(s.desde)} – ${fmtFecha(s.hasta)}</span><span>${fmtNum(s.km)} km: ${fmtNum(s.kmElec)} con batería y ${fmtNum(s.kmGas)} con gasolina</span></li>`).join("")}</ul>`
    : "";

  // Mes a mes
  const meses = a.meses;
  $("#sub-meses").textContent = `${fmtPesosCorto(sum(meses, (m) => m.energia))} en energía y ${fmtPesosCorto(sum(meses, (m) => m.gasolina))} en gasolina, de ${meses[0].nombre} a ${meses[meses.length - 1].nombre}.`;
  const parciales = meses.filter((m) => m.parcial || m.sinTarifa).map((m) => m.nombre);
  $("#nota-meses").textContent = "La gasolina cae en el mes en que se tanquea, no en el que se gasta." +
    (parciales.length ? ` ${parciales.join(", ")}: faltan cargas o tarifas, va incompleto.` : "");
  columnas($("#c-meses"), {
    titulo: "Costo mensual de energía y gasolina",
    categorias: meses.map((m) => m.nombre + (m.parcial || m.sinTarifa ? "*" : "")),
    series: [
      { nombre: "Energía", clase: "s1", valores: meses.map((m) => m.energia) },
      { nombre: "Gasolina", clase: "s2", valores: meses.map((m) => m.gasolina) },
    ],
    fmtEje: fmtEjePesos,
    tip: (i) => {
      const m = meses[i];
      return `<b>${m.nombre}${m.parcial ? " (incompleto)" : ""}</b>` +
        filaTip("s1", `Energía (${m.cargas} cargas)`, fmtPesos(m.energia)) +
        filaTip("s2", `Gasolina (${fmtNum(m.galones, 1)} gal)`, fmtPesos(m.gasolina)) +
        filaTip("", "Total", fmtPesos(m.energia + m.gasolina));
    },
  });

  // Casa
  const casa = a.casa;
  $("#sub-casa").textContent = a.pctCasa != null
    ? `Desde que llegó, Elektra se lleva el ${fmtPct(a.pctCasa)} de la energía de la casa` +
      (a.promedioAntes ? ` (antes, la casa gastaba unos ${fmtNum(a.promedioAntes)} kWh al mes).` : ".")
    : "Todavía no hay cargas dentro de los periodos de las facturas.";
  columnas($("#c-casa"), {
    titulo: "Consumo de la casa por factura, con la parte del carro",
    categorias: casa.map((p) => p.nombre),
    series: [
      { nombre: "Resto de la casa", clase: "neutral", valores: casa.map((p) => p.resto) },
      { nombre: "Elektra", clase: "s1", valores: casa.map((p) => p.carro) },
    ],
    etiquetas: casa.map((p) => (p.carro > 0 ? fmtPct(p.carro / p.consumo) : "")),
    fmtEje: (v) => fmtNum(v),
    tip: (i) => {
      const p = casa[i];
      return `<b>Factura de ${p.nombre}</b><div class="tip-sub">${fmtFecha(p.desde)} – ${fmtFecha(p.hasta)}</div>` +
        filaTip("neutral", "Resto de la casa", `${fmtNum(p.resto)} kWh`) +
        filaTip("s1", "Elektra", `${fmtNum(p.carro)} kWh`) +
        filaTip("", "Precio del kWh", fmtPesos(p.precio));
    },
  });

  // Hábito
  const bins = a.bins;
  const maxBin = Math.max(...bins.map((b) => b.n));
  $("#sub-habito").textContent = `${a.bajo20} de ${a.nCargas} cargas empiezan con menos del 20 %; ${a.en80} paran en 80 %, y ${a.deNoche} arrancan entre las 6 p. m. y las 6 a. m.`;
  columnas($("#c-habito"), {
    titulo: "Cargas según la batería al enchufar",
    categorias: bins.map((b) => `${b.desde} %`),
    series: [{ nombre: "Cargas", clase: "s1", valores: bins.map((b) => b.n) }],
    etiquetas: bins.map((b) => (b.n === maxBin && b.n > 0 ? String(b.n) : "")),
    enteros: true,
    fmtEje: (v) => fmtNum(v),
    tip: (i) => `<b>Entre ${bins[i].desde} % y ${bins[i].desde + 10} %</b>` + filaTip("s1", "Cargas", String(bins[i].n)),
  });

  // Tablas
  const filasMes = meses.map((m) => `<tr><td>${m.nombre}${m.parcial ? " *" : ""}</td><td>${m.cargas}</td><td>${fmtNum(m.kwh, 1)}</td><td>${fmtPesos(m.energia)}</td><td>${fmtNum(m.galones, 2)}</td><td>${fmtPesos(m.gasolina)}</td></tr>`).join("");
  const filasCasa = casa.map((p) => `<tr><td>${p.nombre} (${fmtFecha(p.desde)} – ${fmtFecha(p.hasta)})</td><td>${fmtPesos(p.precio)}</td><td>${fmtNum(p.consumo)}</td><td>${fmtNum(p.carro, 1)}</td><td>${p.consumo ? fmtPct(p.carro / p.consumo) : ""}</td></tr>`).join("");
  const filaZona = (n, z) => `<tr><td>${n}</td><td>${fmtNum(z.km)}</td><td>${fmtNum(z.kmElec)}</td><td>${fmtNum(z.kmGas)}</td><td>${z.costo != null ? fmtPesos(z.costo) : "—"}</td><td>${z.costoKm != null ? fmtPesos(z.costoKm) : "—"}</td></tr>`;
  $("#tablas").innerHTML =
    `<div class="tabla-wrap"><table><caption>Bogotá y salidas</caption><thead><tr><th>Dónde</th><th>Km</th><th>Con batería</th><th>Con gasolina</th><th>Costo</th><th>Por km</th></tr></thead><tbody>${filaZona("Bogotá", bo)}${filaZona("Salidas", sa)}</tbody></table></div>` +
    `<div class="tabla-wrap"><table><caption>Mes a mes</caption><thead><tr><th>Mes</th><th>Cargas</th><th>kWh</th><th>Energía</th><th>Galones</th><th>Gasolina</th></tr></thead><tbody>${filasMes}</tbody></table></div>` +
    `<div class="tabla-wrap"><table><caption>Facturas de la casa</caption><thead><tr><th>Periodo</th><th>Precio kWh</th><th>kWh casa</th><th>kWh carro</th><th>Carro</th></tr></thead><tbody>${filasCasa}</tbody></table></div>`;
}

/* ---------- Eventos ---------- */

function mostrarTablero(datos, fuente) {
  estado.datos = datos;
  estado.fuente = fuente;
  estado.ajustes.precioGal = null;
  $("#inicio").hidden = true;
  $("#tablero").hidden = false;
  pintar();
}

function mostrarError(msg) {
  const e = $("#error");
  e.textContent = msg;
  e.hidden = false;
}

async function cargarArchivo(archivo) {
  $("#error").hidden = true;
  try {
    const datos = leerExcel(await archivo.arrayBuffer());
    mostrarTablero(datos, `Datos: ${archivo.name} (leído en este navegador).`);
  } catch (err) {
    mostrarError(`No pude leer ese archivo. ${err.message}`);
  }
}

$("#archivo").addEventListener("change", (e) => e.target.files[0] && cargarArchivo(e.target.files[0]));
$("#plantilla").addEventListener("click", () => XLSX.writeFile(construirPlantilla(), "Plantilla Tablero del vehículo a analizar.xlsx"));
$("#ejemplo").addEventListener("click", () => mostrarTablero(datosEjemplo(), "Datos de ejemplo: inventados, para ver cómo funciona."));
$("#cambiar").addEventListener("click", () => {
  $("#tablero").hidden = true;
  $("#inicio").hidden = false;
  $("#archivo").value = "";
});

const zona = $("#inicio");
["dragenter", "dragover"].forEach((t) => zona.addEventListener(t, (e) => { e.preventDefault(); zona.classList.add("over"); }));
["dragleave", "drop"].forEach((t) => zona.addEventListener(t, (e) => { e.preventDefault(); zona.classList.remove("over"); }));
zona.addEventListener("drop", (e) => e.dataTransfer.files[0] && cargarArchivo(e.dataTransfer.files[0]));

$("#contribucion").addEventListener("change", (e) => {
  estado.ajustes.contribucion = e.target.checked;
  pintar();
});
document.querySelectorAll('input[name="tipoGas"]').forEach((r) => r.addEventListener("change", (e) => {
  estado.ajustes.tipo = e.target.value;
  estado.ajustes.precioGal = null;
  pintar();
}));
[["#kmGalCiudad", "kmGalCiudad"], ["#kmGalCarretera", "kmGalCarretera"]].forEach(([sel, clave]) =>
  $(sel).addEventListener("input", (e) => {
    const v = Number(e.target.value);
    if (v >= 5) { estado.ajustes[clave] = v; pintar(); }
  }));
$("#precioGal").addEventListener("input", (e) => {
  const v = Number(e.target.value);
  if (v >= 1000) { estado.ajustes.precioGal = v; pintar(); }
});

let anchoPrevio = 0;
new ResizeObserver(() => {
  const w = $("#tablero").clientWidth;
  if (w && Math.abs(w - anchoPrevio) > 4) {
    anchoPrevio = w;
    graficas.forEach((g) => g.dibujar());
  }
}).observe($("#tablero"));

// ?ejemplo abre el tablero directamente con los datos de ejemplo (para compartir el enlace).
if (new URLSearchParams(location.search).has("ejemplo")) {
  mostrarTablero(datosEjemplo(), "Datos de ejemplo: inventados, para ver cómo funciona.");
}

// Solo para probar en el computador: ?archivo=ruta-relativa carga un Excel servido localmente.
const prueba = new URLSearchParams(location.search).get("archivo");
if (prueba && !prueba.startsWith("//") && /^\/?[\w\-./ %]+\.xlsx$/i.test(prueba)) {
  fetch(prueba).then((r) => r.blob()).then((b) => cargarArchivo(new File([b], decodeURIComponent(prueba.split("/").pop()))));
}
