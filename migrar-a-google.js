/* =============================================================================
   PASAR LOS DATOS DEL SISTEMA ANTERIOR (SQL Server) AL SISTEMA DE GITHUB
   -----------------------------------------------------------------------------
   1. Copia este archivo en la carpeta donde está server.js (la del sistema de
      bienes reposeídos, la que tiene la carpeta node_modules).
   2. Abre una ventana de comandos (cmd) en esa carpeta y escribe:
          node migrar-a-google.js
   3. Escribe tu usuario y tu contraseña del sistema de GitHub (administrador).
   4. Revisa el resumen y escribe SI para pasar los datos.

   - Solo LEE la base BIENES_REPOSEIDOS: no cambia ni borra nada del sistema anterior.
   - Usa la misma conexión del sistema anterior (bienesReposeidos.js y
     bienes-reposeidos.config). La contraseña de la base no sale de este servidor.
   - En el sistema de GitHub reemplaza cortes, fincas, ventas, gestiones, documentos,
     alquileres y contactos por los del sistema anterior. Los avalúos no se tocan.
   ============================================================================= */
'use strict';
const fs = require('fs'), path = require('path'), https = require('https'), http = require('http'), readline = require('readline');
const URL_API = process.env.MIGRAR_URL || 'https://script.google.com/macros/s/AKfycbzksqHyIr1N13az-EPB-Rx0qlhU2MBJg6Nr9ftnBu8bO3ss4uvb9bqolqLu-jAISiA5PQ/exec';

// ---------- buscar el sistema anterior y su conexión ----------
function buscar(nombre, dir, nivel) {
  const p = path.join(dir, nombre);
  if (fs.existsSync(p)) return p;
  if (!nivel) return null;
  let hijos = [];
  try { hijos = fs.readdirSync(dir, { withFileTypes: true }).filter(d => d.isDirectory() && !/^(node_modules|\.git)$/.test(d.name)); } catch (e) { return null; }
  for (const d of hijos) { const r = buscar(nombre, path.join(dir, d.name), nivel - 1); if (r) return r; }
  return null;
}
function requerir(modulo, carpetas) {
  try { return require(require.resolve(modulo, { paths: carpetas })); }
  catch (e) { console.error('\n❌ No encontré el paquete "' + modulo + '". Corre este programa en la carpeta del sistema (donde está node_modules).'); process.exit(1); }
}
function leerConfig(carpetas) {
  for (const d of carpetas) for (const n of ['bienes-reposeidos.config', 'bienes-reposeidos.config.txt']) {
    const ruta = path.join(d, n);
    if (!fs.existsSync(ruta)) continue;
    const buf = fs.readFileSync(ruta);
    let txt = buf[0] === 0xFF && buf[1] === 0xFE ? buf.slice(2).toString('utf16le')
      : buf[0] === 0xFE && buf[1] === 0xFF ? Buffer.from(buf.slice(2)).swap16().toString('utf16le') : buf.toString('utf8');
    try { return { ruta, cfg: JSON.parse(txt.replace(/^﻿/, '').trim()) }; } catch (e) { console.error('⚠ ' + ruta + ' no es JSON válido; se ignora.'); }
  }
  return null;
}

// ---------- preguntar en la consola ----------
let RL = null, MUDO = false;
const LINEAS = [], ESPERAN = [];
function preguntar(texto, oculto) {
  if (!RL) {
    RL = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: !!process.stdin.isTTY });
    RL._writeToOutput = function (s) { if (MUDO && s !== '\r\n' && s !== '\n') return; process.stdout.write(s); };
    RL.on('line', l => { MUDO = false; if (ESPERAN.length) ESPERAN.shift()(l); else LINEAS.push(l); });
  }
  process.stdout.write(texto);
  if (LINEAS.length) { process.stdout.write('\n'); return Promise.resolve(LINEAS.shift()); }
  MUDO = !!oculto;
  return new Promise(ok => ESPERAN.push(ok));
}

// ---------- llamar al sistema de GitHub (Apps Script) ----------
function pedir(url, metodo, cuerpo, saltos) {
  return new Promise((ok, mal) => {
    const u = new URL(url);
    const opc = { method: metodo, hostname: u.hostname, path: u.pathname + u.search, headers: {} };
    if (cuerpo) { opc.headers['Content-Type'] = 'text/plain;charset=utf-8'; opc.headers['Content-Length'] = Buffer.byteLength(cuerpo); }
    const req = (u.protocol === 'http:' ? http : https).request(Object.assign(opc, { port: u.port || undefined }), res => {
      if ([301, 302, 303, 307, 308].indexOf(res.statusCode) !== -1 && res.headers.location) {
        res.resume();
        if ((saltos || 0) > 5) return mal(new Error('Demasiadas redirecciones'));
        const sig = new URL(res.headers.location, url).toString();
        return ok(pedir(sig, res.statusCode === 307 || res.statusCode === 308 ? metodo : 'GET', res.statusCode === 307 || res.statusCode === 308 ? cuerpo : null, (saltos || 0) + 1));
      }
      const partes = []; res.on('data', d => partes.push(d)); res.on('end', () => ok(Buffer.concat(partes).toString('utf8')));
    });
    req.on('error', mal); req.setTimeout(330000, () => req.destroy(new Error('El servidor tardó demasiado')));
    if (cuerpo) req.write(cuerpo); req.end();
  });
}
let TOKEN = '';
async function api(metodo, ruta, cuerpo, intentos) {
  const txt = await pedir(URL_API, 'POST', JSON.stringify({ m: metodo, p: ruta, b: cuerpo || {}, t: TOKEN })).catch(e => ({ error: e }));
  let r = null;
  if (typeof txt === 'string') { try { r = JSON.parse(txt); } catch (e) { r = null; } }
  if (!r) {
    if ((intentos || 0) < 2) { await new Promise(s => setTimeout(s, 4000)); return api(metodo, ruta, cuerpo, (intentos || 0) + 1); }
    throw new Error('El sistema de GitHub no respondió bien' + (txt && txt.error ? ': ' + txt.error.message : ''));
  }
  if (r.token !== undefined && r.token) TOKEN = r.token;
  if (!r.success) throw new Error(r.error || 'Error del sistema de GitHub');
  return r;
}

// ---------- convertir filas de SQL ----------
function limpio(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return isNaN(v) ? null : v.toISOString();
  if (Buffer.isBuffer(v)) return undefined;
  if (typeof v === 'bigint') return Number(v);
  return v;
}
function fila(f) {
  const o = {};
  Object.keys(f).forEach(k => { const v = limpio(f[k]); if (v !== undefined) o[k] = v; });
  return o;
}
// tandas de hasta ~2 MB o 400 filas
function tandas(filas) {
  const out = []; let act = [], peso = 0;
  filas.forEach(f => {
    const p = JSON.stringify(f).length;
    if (act.length && (act.length >= 400 || peso + p > 2000000)) { out.push(act); act = []; peso = 0; }
    act.push(f); peso += p;
  });
  if (act.length || !out.length) out.push(act);
  return out;
}

// ======================= LECTURA DE LOS EXCEL (igual que la página) =======================
// Valor "limpio" de una celda de ExcelJS (fórmulas, texto enriquecido, etc.)
function valor(celda) {
  const v = celda ? celda.value : null;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (Object.prototype.hasOwnProperty.call(v, 'result')) return v.result === undefined ? null : v.result;
    // Fórmula sin resultado guardado: la librería lo omite cuando el resultado es 0
    if (v.formula || v.sharedFormula) return 0;
    if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
    if (v.text !== undefined) return v.text;
    if (v.error) return null;
    return null;
  }
  return v;
}

function dosDig(n) { return String(n).padStart(2, '0'); }
function fechaISO(d) { return d.getUTCFullYear() + '-' + dosDig(d.getUTCMonth() + 1) + '-' + dosDig(d.getUTCDate()); }
function fechaDMY(d) { return dosDig(d.getUTCDate()) + '/' + dosDig(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear(); }
function serialAFecha(n) { return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000); }

function texto(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return fechaDMY(v);
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(Math.round(v * 100) / 100);
  const s = String(v).replace(/\u00a0/g, ' ').trim();
  return s === '' ? null : s;
}

function numero(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return null;
  if (typeof v === 'number') return isFinite(v) ? Math.round(v * 100) / 100 : null;
  const s = String(v).replace(/B\/\.?/gi, '').replace(/,/g, '').replace(/\s/g, '');
  if (s === '' || !/^-?\d+(\.\d+)?$/.test(s)) return null;
  return Math.round(Number(s) * 100) / 100;
}

// "No. Finca" -> "nofinca", "Sup. (m²)" -> "supm2", "Código Ubic." -> "codigoubic"
function norm(v) {
  const s = texto(v);
  if (!s) return '';
  return s.normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9#]/g, '');
}

function esNumeroFila(v) {
  if (typeof v === 'number') return true;
  return typeof v === 'string' && /^\s*\d+\s*$/.test(v);
}

// Finca como texto sin ".0" ni espacios
function finca(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return String(Math.round(v));
  const s = String(v).trim();
  return s === '' ? null : s;
}

const MESES = { enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8,
  septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, niviembre: 11, diciembre: 12 };

// "17 de octubre de  2024", "16 de abril 2014", "01/09/2026", Date, serial
function interpretarFecha(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return fechaISO(v);
  if (typeof v === 'number') return v > 20000 && v < 80000 ? fechaISO(serialAFecha(v)) : null;
  const s = String(v).toLowerCase();
  let m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const d = +m[1], mes = +m[2], a = +m[3];
    if (mes >= 1 && mes <= 12 && d >= 1 && d <= 31) return a + '-' + dosDig(mes) + '-' + dosDig(d);
    return null;
  }
  m = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').match(/(\d{1,2})\s*de\s*([a-z]+)\s*(?:de\s*)?(\d{4})/);
  if (m && MESES[m[2]]) return m[3] + '-' + dosDig(MESES[m[2]]) + '-' + dosDig(+m[1]);
  return null;
}

function buscarHoja(wb, prueba) {
  let hallada = null;
  wb.eachSheet(ws => { if (!hallada && prueba(ws)) hallada = ws; });
  return hallada;
}

// Fila donde está el encabezado "#" + "No. Finca"
function filaEncabezado(ws, maxFila) {
  const tope = Math.min(ws.rowCount, maxFila || 20);
  for (let r = 1; r <= tope; r++) {
    const fila = ws.getRow(r);
    for (let c = 1; c < Math.max(fila.cellCount, 1); c++) {
      if (norm(valor(fila.getCell(c))) === '#' && norm(valor(fila.getCell(c + 1))) === 'nofinca') return r;
    }
  }
  return 0;
}

// Mapa encabezado-normalizado -> [columnas] en un rango de columnas
function mapaColumnas(ws, r, desde, hasta) {
  const mapa = {};
  const fila = ws.getRow(r);
  for (let c = desde; c <= hasta; c++) {
    const k = norm(valor(fila.getCell(c)));
    if (!k) continue;
    (mapa[k] = mapa[k] || []).push(c);
  }
  return mapa;
}

// Primer valor no vacío entre las columnas de cualquiera de los nombres
function leer(fila, mapa, ...nombres) {
  for (const n of nombres) {
    const cols = typeof n === 'function'
      ? Object.keys(mapa).filter(n).reduce((a, k) => a.concat(mapa[k]), [])
      : (mapa[n] || []);
    for (const c of cols) {
      const v = valor(fila.getCell(c));
      if (v !== null && v !== undefined && !(typeof v === 'string' && v.trim() === '')) return v;
    }
  }
  return null;
}

// ---- Hoja "Ventas en Proceso": tres bloques lado a lado ------------------
function leerVentas(wb, avisos) {
  const ws = buscarHoja(wb, h => filaEncabezado(h) && /venta/.test(norm(h.name)));
  if (!ws) return null;
  const rh = filaEncabezado(ws);
  const encab = ws.getRow(rh);
  const ultimaCol = Math.max(ws.columnCount, encab.cellCount);

  // Cada bloque empieza en una columna "#" seguida de "No. Finca"
  const inicios = [];
  for (let c = 1; c <= ultimaCol; c++) {
    if (norm(valor(encab.getCell(c))) === '#' && norm(valor(encab.getCell(c + 1))) === 'nofinca') inicios.push(c);
  }

  const bloques = inicios.map((ini, i) => {
    const fin = (inicios[i + 1] || ultimaCol + 1) - 1;
    // El título del bloque está en alguna de las filas de arriba, en su primera columna
    let titulo = '', fechaTabla = null;
    for (let r = 1; r < rh; r++) {
      const t = texto(valor(ws.getRow(r).getCell(ini)));
      if (!t) continue;
      const n = norm(t);
      if (/noconcretad/.test(n)) titulo = 'NO_CONCRETADA';
      else if (/finalizad/.test(n)) titulo = 'FINALIZADA';
      else if (/proceso/.test(n)) titulo = 'EN_PROCESO';
      if (/fechadereporte/.test(n)) {
        fechaTabla = interpretarFecha(t);
        // "Fecha de reporte:" en una celda y la fecha en la de al lado
        for (let c = ini + 1; !fechaTabla && c <= Math.min(ini + 6, (inicios[i + 1] || ultimaCol + 1) - 1); c++) {
          fechaTabla = interpretarFecha(valor(ws.getRow(r).getCell(c)));
        }
      }
    }
    return { ini, fin, tabla: titulo, fechaTabla, mapa: mapaColumnas(ws, rh, ini, fin) };
  });

  const filas = [];
  for (const b of bloques) {
    if (!b.tabla) { avisos.push('Hay un bloque de ventas sin título reconocible (columna ' + b.ini + '); no se leyó.'); continue; }
    for (let r = rh + 1; r <= ws.rowCount; r++) {
      const fila = ws.getRow(r);
      const num = valor(fila.getCell(b.ini));
      const nf = finca(valor(fila.getCell(b.ini + 1)));
      if (!nf || !esNumeroFila(num)) continue;
      filas.push({
        Tabla: b.tabla,
        FechaTabla: b.fechaTabla,
        Orden: Number(num),
        NoFinca: nf,
        CodigoUbic: texto(leer(fila, b.mapa, k => /^cod.*ubic/.test(k))),
        Ubicacion: texto(leer(fila, b.mapa, 'ubicacion')),
        Superficie: texto(leer(fila, b.mapa, k => k.startsWith('sup'))),
        ValorVentaNegociado: numero(leer(fila, b.mapa, 'valorventanegociado')),
        ValorLibros: numero(leer(fila, b.mapa, 'valorenlibros')),
        ValorVentaRapida: numero(leer(fila, b.mapa, 'valordeventarapida')),
        MontoFinanciado: numero(leer(fila, b.mapa, 'montofinanciado')),
        Cliente: texto(leer(fila, b.mapa, 'clienteinteresado', 'comprador')),
        Ejecutivo: texto(leer(fila, b.mapa, 'ejecutivo')),
        Estatus: texto(leer(fila, b.mapa, k => k.startsWith('estatus'))),
        AntiguoPropietario: texto(leer(fila, b.mapa, 'antiguopropietario')),
        ComisionPagada: numero(leer(fila, b.mapa, 'comisionpagada'))
      });
    }
  }
  const fechas = bloques.map(b => b.fechaTabla).filter(Boolean).sort();
  const proceso = bloques.filter(b => b.tabla === 'EN_PROCESO' && b.fechaTabla)[0];
  return { hoja: ws.name, filas, fechaProceso: proceso ? proceso.fechaTabla : null,
    fechaReporte: fechas.length ? fechas[fechas.length - 1] : null };
}

// ---- Hoja "Visitas y Fotografías" ---------------------------------------
function leerVisitas(wb) {
  const ws = buscarHoja(wb, h => /visita/.test(norm(h.name)) && filaEncabezado(h));
  if (!ws) return null;
  const rh = filaEncabezado(ws);
  const mapa = mapaColumnas(ws, rh, 1, Math.max(ws.columnCount, ws.getRow(rh).cellCount));
  const col = (k) => (mapa[k] || [])[0];
  const filas = [];
  for (let r = rh + 1; r <= ws.rowCount; r++) {
    const fila = ws.getRow(r);
    const nf = finca(valor(fila.getCell(col('nofinca'))));
    if (!nf) continue;
    const num = valor(fila.getCell(col('#')));
    filas.push({
      Orden: esNumeroFila(num) ? Number(num) : null,
      NoFinca: nf,
      CodigoUbic: texto(leer(fila, mapa, k => /^cod.*ubic/.test(k))),
      Ubicacion: texto(leer(fila, mapa, 'ubicacion')),
      TipoPropiedad: texto(leer(fila, mapa, 'tipodepropiedad')),
      Condicion: texto(leer(fila, mapa, k => k.startsWith('condicion'))),
      Letrero: texto(leer(fila, mapa, k => k.startsWith('letrero'))),
      Observaciones: texto(leer(fila, mapa, k => k.startsWith('observaciones'))),
      Limpieza: texto(leer(fila, mapa, 'limpieza')),
      ProximaLimpieza: texto(leer(fila, mapa, 'proximalimpieza'))
    });
  }
  return { hoja: ws.name, filas };
}

// ---- Hoja "Resumen Ejecutivo" -------------------------------------------
function leerResumen(wb) {
  const ws = buscarHoja(wb, h => /resumenejecutivo/.test(norm(h.name)));
  if (!ws) return null;
  let rh = 0;
  for (let r = 1; r <= Math.min(ws.rowCount, 20) && !rh; r++) {
    if (/clientefinca/.test(norm(valor(ws.getRow(r).getCell(1))))) rh = r;
  }
  if (!rh) return { hoja: ws.name, filas: [] };
  const filas = [];
  let seccion = null, orden = 0;
  for (let r = rh + 1; r <= ws.rowCount; r++) {
    const a = texto(valor(ws.getRow(r).getCell(1)));
    const b = texto(valor(ws.getRow(r).getCell(2)));
    if (!a) continue;
    // Las filas de título vienen combinadas A:B, ExcelJS repite el texto en B
    if (!b || b === a) { seccion = a; continue; }
    filas.push({ Orden: ++orden, Seccion: seccion, ClienteFinca: a, Estatus: b });
  }
  return { hoja: ws.name, filas };
}

// ---- Análisis de trazabilidad -------------------------------------------
function leerTrazabilidad(wb) {
  const ws = buscarHoja(wb, h => {
    const r = filaEncabezado(h, 40);
    return r && Object.keys(mapaColumnas(h, r, 1, 30)).some(k => k.startsWith('fechadeadjudicacion'));
  });
  if (!ws) return null;

  let fechaReporte = null;
  for (let r = 1; r <= 5 && !fechaReporte; r++) {
    const v = valor(ws.getRow(r).getCell(1));
    if (v instanceof Date || (typeof v === 'number' && v > 20000)) fechaReporte = interpretarFecha(v);
  }

  const filas = [];
  let grupo = 'VENDIDA', mapa = null, colNum = 1;
  for (let r = 1; r <= ws.rowCount; r++) {
    const fila = ws.getRow(r);
    const primera = norm(valor(fila.getCell(1)));
    if (/ventasenproceso/.test(primera)) { grupo = 'EN_PROCESO'; continue; }
    if (/fincasreposeidasvendidas|ventafinalizada/.test(primera)) { grupo = 'VENDIDA'; continue; }
    if (primera === '#' && norm(valor(fila.getCell(2))) === 'nofinca') {
      mapa = mapaColumnas(ws, r, 1, Math.max(ws.columnCount, fila.cellCount));
      colNum = (mapa['#'] || [1])[0];
      continue;
    }
    if (!mapa) continue;
    const num = valor(fila.getCell(colNum));
    const nf = finca(leer(fila, mapa, 'nofinca'));
    if (!nf || !esNumeroFila(num)) continue;
    const fa = leer(fila, mapa, k => k.startsWith('fechadeadjudicacion'));
    filas.push({
      Grupo: grupo,
      Orden: Number(num),
      NoFinca: nf,
      FechaAdjudicacion: texto(fa),
      FechaAdjudicacionD: interpretarFecha(fa),
      Ubicacion: texto(leer(fila, mapa, 'ubicacion')),
      TipoPropiedad: texto(leer(fila, mapa, 'tipodepropiedad')),
      ImpuestoInmueble: numero(leer(fila, mapa, 'impuestodeinmueble')),
      GastoMantenimiento: numero(leer(fila, mapa, 'gastodemantenimiento')),
      GastosLegales: numero(leer(fila, mapa, 'gastoslegales')),
      Impuesto2: numero(leer(fila, mapa, 'impuesto2')),
      Impuesto3: numero(leer(fila, mapa, 'impuesto3')),
      Comision: numero(leer(fila, mapa, 'comision')),
      TotalGastos: numero(leer(fila, mapa, 'totaldegastos')),
      ValorLibros: numero(leer(fila, mapa, 'valorenlibros')),
      ValorVentaRapida: numero(leer(fila, mapa, 'valordeventarapida')),
      ValorVentaNegociado: numero(leer(fila, mapa, 'valorventanegociado')),
      Balance: numero(leer(fila, mapa, 'balance')),
      MontoFinanciado: numero(leer(fila, mapa, 'montofinanciado')),
      Abono: numero(leer(fila, mapa, 'abono')),
      Contado: numero(leer(fila, mapa, 'contado'))
    });
  }
  return { hoja: ws.name, filas, fechaReporte };
}

// Lee un libro completo y dice qué encontró
// ---- "Activos disponibles para la venta": la lista completa de activos ----
// Hoja con encabezado "No. De Finca" … "Valor en Libros" … "Estado".
// Estados del Excel → Dashboard: DISPONIBLE, ALQUILADA (disponible y alquilada),
// VIVEN PERSONAS (disponible y ocupada), EN TRÁMITE, EN LEGAL, VENDIDA.
function leerActivos(wb, avisos) {
  const propio = cel => (cel && cel.type === 1 ? null : valor(cel));   // celda combinada: solo cuenta la principal
  let ws = null, rh = 0, mapa = null;
  wb.eachSheet(h => {
    if (ws) return;
    for (let r = 1; r <= Math.min(h.rowCount, 20); r++) {
      const m = mapaColumnas(h, r, 1, Math.max(h.columnCount, h.getRow(r).cellCount));
      if ((m.nodefinca || m.nodefinca === undefined && m.nofinca) && Object.keys(m).some(k => k.startsWith('valorenlibros')) && (m.estado || Object.keys(m).some(k => k.startsWith('valordemercado')))) { ws = h; rh = r; mapa = m; return; }
    }
  });
  if (!ws) return null;
  // Fecha del título: "AL 29 DE SEPTIEMBRE DE 2026"
  let fecha = null;
  for (let r = 1; r < rh && !fecha; r++) {
    ws.getRow(r).eachCell(c => { if (!fecha) { const t = texto(valor(c)); if (t && /\bal\b/i.test(t)) fecha = interpretarFecha(t); } });
  }
  const col = (...ks) => { for (const k of ks) { const x = Object.keys(mapa).filter(m => m === k || m.startsWith(k))[0]; if (x) return mapa[x][0]; } return null; };
  const C = {
    num: 1, fecha: col('fechadeadjudicacion'), finca: col('nodefinca'), obs: col('observaciones'), deudor: col('deudor'),
    alquilado: col('alquilado'), letrero: col('letrero'), ubic: col('ubicacion'), detalle: col('detalledelapropiedad'),
    m2: col('m2'), tipo: col('tipo'), terreno: col('terrenoauto', 'terreno'), mejoras: col('mejoras'), mercado: col('valordemercado'),
    dias: col('diastranscurridos'), fideicomiso: col('fideicomiso', 'fideidcomiso'), deterioro: col('valordeperdida'), provision: col('totaldeprovision'),
    rapida: col('valordeventarapida'), libros: col('valorenlibros'), estado: col('estado'), politica: col('politica'), pct: null
  };
  ws.getRow(rh).eachCell((c, j) => { if (!C.pct && String(texto(valor(c)) || '').trim() === '%') C.pct = j; });
  const ESTADOS = { DISPONIBLE: ['DISPONIBLE', false, false], ALQUILADA: ['DISPONIBLE', true, false], VIVENPERSONAS: ['DISPONIBLE', false, true],
    ENTRAMITE: ['EN_TRAMITE', false, false], ENLEGAL: ['LEGAL', false, false], VENDIDA: ['VENDIDA', false, false] };
  const filas = [], vistos = {};
  const marcas = marcasActivos(ws, rh);
  // Formato actual de contabilidad (sin columna Estado): todas son activos disponibles para la venta;
  // el estatus de cada finca (en trámite, legal, alquilada) se conserva el que ya tiene en el Dashboard.
  const sinEstado = !C.estado;
  if (sinEstado) avisos.push('Este Excel no trae la columna "Estado": las fincas nuevas entran como disponibles y las que ya están en el Dashboard conservan su estatus (en trámite, legal, alquilada). El estatus se cambia en el Dashboard.');
  const ultimaCol = Math.max(C.estado || 0, C.provision || 0, 18);
  for (let r = rh + 1; r <= ws.rowCount; r++) {
    const fila = ws.getRow(r);
    const n = propio(C.num ? fila.getCell(C.num) : null);
    if (!esNumeroFila(n)) continue;
    let rojas = 0, poliza = false;
    for (let c = 1; c <= ultimaCol; c++) {
      const cel = fila.getCell(c);
      if (cel.font && cel.font.color && cel.font.color.argb === 'FFFF0000') rojas++;
      if (marcas.poliza && cel.fill && cel.fill.fgColor && JSON.stringify(cel.fill.fgColor) === marcas.poliza) poliza = true;
    }
    const bruto = texto(propio(C.finca ? fila.getCell(C.finca) : null)) || '';
    const detalle = texto(propio(C.detalle ? fila.getCell(C.detalle) : null)) || '';
    let noFinca = null, codigo = null;
    const mf = bruto.match(/^\s*([A-Za-z]?\d+)\s*-\s*([0-9A-Za-z]+)\s*$/);
    if (mf) { noFinca = mf[1]; codigo = mf[2]; }
    else if (bruto.trim()) noFinca = bruto.trim().replace(/\s+/g, '');
    else {   // autos y camiones: se identifican por la placa
      const mp = detalle.match(/placa\s*(?:n[°ºo.]*)?\s*:?\s*([A-Z0-9]{4,10})/i);
      noFinca = mp ? mp[1].toUpperCase() : 'ACTIVO-' + Number(n);
    }
    noFinca = noFinca.slice(0, 30);
    if (vistos[noFinca]) { avisos.push('La finca ' + noFinca + ' aparece dos veces en el Excel de activos; se usa la primera.'); continue; }
    vistos[noFinca] = 1;
    const estTxt = sinEstado ? 'DISPONIBLE' : norm(propio(C.estado ? fila.getCell(C.estado) : null)).replace(/#/g, '').toUpperCase();
    let est = ESTADOS[estTxt];
    if (!est) { avisos.push('Finca ' + noFinca + ': estado "' + (texto(propio(C.estado ? fila.getCell(C.estado) : null)) || 'vacío') + '" no reconocido; se toma como DISPONIBLE.'); est = ESTADOS.DISPONIBLE; }
    // Una celda combinada entre dos filas se cuenta solo en la primera (como en el Excel)
    const num = c => { if (!c) return null; const cel = fila.getCell(c); return cel.type === 1 ? null : numero(valor(cel)); };
    const mej = num(C.mejoras);
    filas.push({
      Orden: Number(n), NoFinca: noFinca, CodigoUbic: codigo, EstadoExcel: sinEstado ? 'DISPONIBLE PARA LA VENTA' : (texto(propio(C.estado ? fila.getCell(C.estado) : null)) || ''), SinEstado: sinEstado,
      Estado: est[0], Alquilada: est[1], Ocupada: est[2], Mejora: mej !== null && mej > 0 ? 'CON' : 'SIN',
      TipoPropiedad: texto(propio(C.tipo ? fila.getCell(C.tipo) : null)), Ubicacion: texto(propio(C.ubic ? fila.getCell(C.ubic) : null)),
      Detalle: detalle ? detalle.slice(0, 1000) : null, Superficie: texto(propio(C.m2 ? fila.getCell(C.m2) : null)),
      Deudor: texto(propio(C.deudor ? fila.getCell(C.deudor) : null)), FechaAdjudicacion: texto(propio(C.fecha ? fila.getCell(C.fecha) : null)),
      Letrero: texto(propio(C.letrero ? fila.getCell(C.letrero) : null)), ValorMercado: num(C.mercado),
      ValorVentaRapida: num(C.rapida), ValorLibros: num(C.libros),
      ValorMejoras: mej, Observaciones: texto(propio(C.obs ? fila.getCell(C.obs) : null)),
      Dias: C.dias ? (numero(propio(C.dias ? fila.getCell(C.dias) : null)) !== null ? Math.round(numero(propio(C.dias ? fila.getCell(C.dias) : null))) : null) : null,
      Fideicomiso: C.fideicomiso ? texto(propio(C.fideicomiso ? fila.getCell(C.fideicomiso) : null)) : null,
      ValorTerreno: num(C.terreno), Deterioro: num(C.deterioro), Provision: num(C.provision),
      NoSeVende: rojas >= Math.ceil(ultimaCol * 0.6), PolizaIncendio: poliza,
      Porcentaje: num(C.pct), Politica: C.politica ? texto(propio(fila.getCell(C.politica))) : null
    });
  }
  return { hoja: ws.name, filas, fecha, pie: marcas.pie };
}

// ---- Hoja "Gestión de Desalojo" del Excel de activos ----
// Secciones ("1. DESALOJO EN CURSO", "OTRAS PROPIEDADES CON OCUPANTES"…) y filas con # y No. Finca.
function leerDesalojo(wb) {
  let ws = null, rh = 0, mapa = null;
  wb.eachSheet(h => {
    if (ws) return;
    for (let r = 1; r <= Math.min(h.rowCount, 15); r++) {
      const m = mapaColumnas(h, r, 1, Math.max(h.columnCount, h.getRow(r).cellCount));
      if (m.nofinca && Object.keys(m).some(k => k.startsWith('quienreside'))) { ws = h; rh = r; mapa = m; return; }
    }
  });
  if (!ws) return [];
  const col = (...ks) => { for (const k of ks) { const x = Object.keys(mapa).filter(m => m === k || m.startsWith(k))[0]; if (x) return mapa[x][0]; } return null; };
  const C = { num: col('#', 'numero') || mapa.nofinca[0] - 1, finca: col('nofinca'), prop: col('antiguopropietario'), ubic: col('ubicacion'),
    reside: col('quienreside'), accion: col('accionatomar'), situacion: col('situacionactual'), gestion: col('ultimagestion'),
    visita: col('fechault'), mercado: col('valorde', 'valordemercado'), rapida: col('valordeventa'), libros: col('valoren'),
    estatus: col('estatus'), notif: col('fechanotificacion', 'fechanotif'), limite: col('fechalimite', 'fechalim'),
    responsable: col('responsable'), proxima: col('proximaaccion') };
  const filas = []; let seccion = '';
  for (let r = rh + 1; r <= ws.rowCount; r++) {
    const fila = ws.getRow(r);
    const n = valor(fila.getCell(C.num));
    const primero = texto(n);
    if (!esNumeroFila(n)) {
      if (primero && /^(\d+\.\s|otras|total|control)/i.test(primero.trim()) && !/^(total|control)/i.test(primero.trim())) seccion = primero.replace(/\s*\(\d+\)\s*$/, '').trim();
      continue;
    }
    const bruto = texto(valor(fila.getCell(C.finca))) || '';
    const mf = bruto.match(/^\s*([A-Za-z]?\d+)\s*-\s*[0-9A-Za-z]+\s*$/);
    const noFinca = (mf ? mf[1] : bruto.trim().replace(/\s+/g, '')).slice(0, 30);
    if (!noFinca) continue;
    const t = c => { const v = c ? texto(valor(fila.getCell(c))) : null; return v && v !== '—' ? v : null; };
    const fv = C.visita ? valor(fila.getCell(C.visita)) : null;
    filas.push({ NoFinca: noFinca, Seccion: seccion.slice(0, 150), AntiguoPropietario: t(C.prop), Ubicacion: t(C.ubic), QuienReside: t(C.reside),
      Accion: t(C.accion), Situacion: t(C.situacion), UltimaGestion: t(C.gestion),
      FechaVisita: fv instanceof Date ? fechaISO(fv) : (interpretarFecha(texto(fv)) || null),
      ValorMercado: C.mercado ? numero(valor(fila.getCell(C.mercado))) : null, ValorVentaRapida: C.rapida ? numero(valor(fila.getCell(C.rapida))) : null,
      ValorLibros: C.libros ? numero(valor(fila.getCell(C.libros))) : null, EstatusDesalojo: t(C.estatus), FechaNotificacion: t(C.notif),
      FechaLimite: t(C.limite), Responsable: t(C.responsable), ProximaAccion: t(C.proxima), Orden: filas.length + 1 });
  }
  return filas;
}

// ---- Marcas por color y notas del pie en la hoja de activos ----
// "NO SE PUEDEN VENDER": fila completa en letra roja. "CUENTA CON PÓLIZA DE INCENDIO VIGENTE":
// fila sombreada con el mismo color de la leyenda. Pie: provisión contable, saldo de garantías y diferencia.
function marcasActivos(ws, filaEnc) {
  const claveRelleno = cel => { const f = cel.fill; return f && f.type === 'pattern' && f.pattern === 'solid' && f.fgColor ? JSON.stringify(f.fgColor) : null; };
  let poliza = null; const pie = {};
  for (let r = filaEnc + 1; r <= ws.rowCount; r++) {
    const fila = ws.getRow(r);
    let txt = '';
    fila.eachCell(c => { const t = texto(valor(c)); if (t && isNaN(Number(t))) txt += ' ' + t; });
    if (!txt) continue;
    const nums = []; fila.eachCell(c => { const v = valor(c); if (typeof v === 'number') nums.push(v); });
    const plano = txt.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (/POLIZA DE INCENDIO/i.test(plano)) {
      for (let c = 1; c <= 12 && !poliza; c++) poliza = claveRelleno(fila.getCell(c));
    }
    if (/PROVISION CONTABLE/i.test(plano) && nums.length) pie.provision = nums[nums.length - 1];
    if (/Saldo en la cuenta de Garantias recuperadas/i.test(plano)) {
      const m = plano.match(/al\s+(.+?)\s*$/i); pie.saldoAl = m ? m[1].trim() : null;
      if (nums.length) pie.saldo = nums[0];
      if (nums.length > 1) pie.saldo2 = nums[1];
    }
    if (/\bDiferencia\b/i.test(plano) && nums.length) pie.diferencia = nums[nums.length - 1];
  }
  return { poliza, pie };
}

// ---- Hoja "datos de ventas": garantías vendidas por año (2026, 2025, 2024…) ----
function leerVentasAnteriores(wb) {
  const out = [];
  wb.eachSheet(ws => {
    let anio = null, mapa = null, orden = 0, vacias = 0;
    for (let r = 1; r <= ws.rowCount; r++) {
      const fila = ws.getRow(r);
      let t = ''; fila.eachCell(c => { const v = texto(valor(c)); if (v) t += ' ' + v; });
      const mt = t.normalize('NFD').replace(/[\u0300-\u036f]/g, '').match(/GARANTIAS VENDIDAS EN EL ANO\s*(\d{4})/i);
      if (mt) { anio = Number(mt[1]); mapa = null; orden = 0; vacias = 0; continue; }
      if (!anio) continue;
      if (!mapa) {
        const m = mapaColumnas(ws, r, 1, Math.max(ws.columnCount, fila.cellCount));
        if (m.nodefinca || m.nofinca || m.nombredelcliente) mapa = m;
        continue;
      }
      if (!t.trim()) { if (++vacias >= 2) { mapa = null; } continue; }
      vacias = 0;
      if (/^\s*total/i.test(t) || /TOTAL DE COMISI/i.test(t)) { mapa = null; continue; }
      // primero el nombre exacto; si no, el que empieza igual (sin confundir "valor de venta" con "valor de venta rápida")
      const col = (...ks) => {
        for (const k of ks) if (mapa[k]) return mapa[k][0];
        for (const k of ks) { const x = Object.keys(mapa).filter(m => m.startsWith(k) && !(k === 'valordeventa' && /rapida/.test(m)))[0]; if (x) return mapa[x][0]; }
        return null;
      };
      const tx = (...ks) => { const c = col(...ks); const v = c ? texto(valor(fila.getCell(c))) : null; return v || null; };
      const nu = (...ks) => { const c = col(...ks); if (!c) return null; const cel = fila.getCell(c); return cel.type === 1 ? null : numero(valor(cel)); };
      let finca = tx('nodefinca', 'nofinca');
      let codigo = tx('codigoubic');
      if (finca) { const mf = finca.match(/^\s*([A-Za-z]?\d+)\s*-\s*([0-9A-Za-z]+)\s*$/); if (mf) { finca = mf[1]; codigo = codigo || mf[2]; } }
      const deudor = tx('deudor', 'nombredelcliente', 'antiguopropietario');
      if (!finca && !deudor) continue;
      out.push({ Anio: anio, Orden: ++orden, NoFinca: finca ? finca.slice(0, 30) : null, CodigoUbic: codigo ? String(codigo).slice(0, 20) : null,
        FechaAdjudicacion: tx('fechadeadjudicacion'), Deudor: deudor, Comprador: tx('comprador'), Ubicacion: tx('ubicacion'),
        Detalle: tx('detalledelapropiedad', 'descripcionesdelasgarantias', 'descripciones'), Tipo: tx('tipo'),
        ValorLibros: nu('valorenlibros', 'saldosinicialesengarantia'), ValorMercado: nu('valordemercado', 'valordeavaluo'),
        ValorVentaRapida: nu('valordeventarapida'), ValorVenta: nu('valorventanegociado', 'valordeventa'),
        MontoFinanciado: nu('montofinanciado', 'financiamientoatraves'), Comision: nu('comisionpagada'),
        Ejecutivo: tx('ejecutivo'), Observaciones: tx('estatusobservaciones', 'observaciones') });
    }
  });
  return out;
}

// ---- Cuadro de bienes reposeídos – lotes con mejoras (casas) ----
// Columnas: Act. · No. · No. de Finca · Antiguo propietario · Ubicación · Detalle · Fecha de adjudicación · Valor de mejoras ·
// Valor de mercado · Venta rápida · Valor negociado de venta (=MAX(venta rápida, libros)×1.05) · ¿Residen? · Quién reside · ¿Visitada? ·
// Condición · Acción a tomar · Comentarios · Responsable/Fecha visita · (valor en libros, columna sin título) · Acción (nota)
const MESES_ES = { enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12 };
function fechaEnTexto(t) {
  if (t instanceof Date && !isNaN(t)) return fechaISO(t);
  const s = String(t || '');
  let m = s.match(/(\d{1,2})\s*[\/-]\s*(\d{1,2})\s*[\/-]\s*(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = norm(s) && s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').match(/(\d{1,2})\s+de\s+([a-z]+)\s+(?:de\s+|del\s+)?(\d{4})/);
  if (m && MESES_ES[m[2]]) return m[3] + '-' + ('0' + MESES_ES[m[2]]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return null;
}
function leerMejoras(wb) {
  let ws = null, rh = 0, mapa = null;
  wb.eachSheet(h => {
    if (ws) return;
    for (let r = 1; r <= Math.min(h.rowCount, 15); r++) {
      const m = mapaColumnas(h, r, 1, Math.max(h.columnCount, h.getRow(r).cellCount));
      const ks = Object.keys(m);
      if (ks.some(k => k.startsWith('nodefinca')) && ks.some(k => k.startsWith('valornegociado')) && ks.some(k => k.startsWith('accionatomar'))) { ws = h; rh = r; mapa = m; return; }
    }
  });
  if (!ws) return null;
  const col = (...ks) => { for (const k of ks) { const x = Object.keys(mapa).filter(m => m === k || m.startsWith(k))[0]; if (x) return mapa[x][0]; } return null; };
  const C = { act: col('act'), num: col('no'), finca: col('nodefinca'), prop: col('antiguopropietario'), ubic: col('ubicacion'), detalle: col('detalle'),
    adj: col('fechadeadjudicacion'), mejoras: col('valordemejoras'), mercado: col('valordemercado'), rapida: col('valordeventarapida'),
    negociado: col('valornegociado'), residen: col('residen'), quien: col('nombredequienreside', 'quienreside'), visitada: col('visitada'),
    condicion: col('condicion'), accion: col('accionatomar'), coment: col('comentarios'), resp: col('responsable') };
  // El valor en libros está en la columna sin título que sigue a "Responsable / Fecha visita"; la nota, en la columna "Acción"
  C.libros = C.resp ? C.resp + 1 : null;
  C.nota = Object.keys(mapa).filter(k => k === 'accion').map(k => mapa[k][0])[0] || null;
  // Fecha del corte: título "AL 31 DE MARZO DE 2026 | Visitas y Fotografías al 18/03/2026"
  let fecha = null, fechaVisitas = null, titulo = '';
  for (let r = 1; r < rh; r++) {
    const vistos = {};
    const t = ws.getRow(r).values.filter(v => v != null).map(v => texto(valor({ value: v })))
      .filter(x => x && !vistos[x] && (vistos[x] = 1)).join(' ');   // las celdas combinadas repiten el mismo texto
    if (/cuadro/i.test(t) && !titulo) titulo = t.trim();
    const mAl = t.match(/al\s+(\d{1,2}\s+de\s+[a-záéíóú]+\s+(?:de\s+)?\d{4})/i);
    if (mAl && !fecha) fecha = fechaEnTexto(mAl[1]);
    const mVis = t.match(/visitas[^|]*?al\s+(\d{1,2}\/\d{1,2}\/\d{4})/i);
    if (mVis) fechaVisitas = fechaEnTexto(mVis[1]);
  }
  const filas = [], avisos = [];
  const txt = c => c ? texto(valor(ws.getRow(rr).getCell(c))) : null;
  // una fórmula sin resultado guardado (archivo no recalculado) se toma como vacía, no como 0
  const sinResultado = cel => cel && cel.value && typeof cel.value === 'object' && cel.value.formula != null && cel.value.result == null;
  const nu = c => { if (!c) return null; const cel = ws.getRow(rr).getCell(c); return sinResultado(cel) ? null : numero(valor(cel)); };
  let rr, filaTotal = null;
  for (rr = rh + 1; rr <= ws.rowCount; rr++) {
    const fila = ws.getRow(rr);
    const a = norm(valor(fila.getCell(1)));
    if (/^totales?/.test(a)) { filaTotal = rr; break; }
    if (/^leyenda/.test(a)) break;
    const bruto = txt(C.finca) || '';
    // solo números de finca ("958-4411", "30385875", "F2000773"); las notas de fila completa se saltan
    if (!/^\s*[A-Za-z]{0,2}\d[\d\s]*(?:-\s*[0-9A-Za-z]+)?\s*$/.test(bruto)) continue;
    const mf = bruto.match(/^\s*([A-Za-z]?\d+)\s*-\s*([0-9A-Za-z]+)\s*$/);
    const noFinca = (mf ? mf[1] : bruto.trim().replace(/\s+/g, '')).slice(0, 30), codigo = mf ? mf[2] : null;
    const adjCel = valor(fila.getCell(C.adj));
    const respCel = C.resp ? valor(fila.getCell(C.resp)) : null;
    const respTxt = respCel instanceof Date ? null : texto(respCel);
    const fv = respCel instanceof Date ? fechaISO(respCel) : fechaEnTexto(respTxt);
    const si = v => { const k = norm(v); return k === 'si' ? true : k === 'no' ? false : null; };
    filas.push({
      _fila: rr, Orden: numero(valor(fila.getCell(C.num))) || filas.length + 1, NoFinca: noFinca, CodigoUbic: codigo,
      Activo: C.act ? si(valor(fila.getCell(C.act))) : null,
      AntiguoPropietario: txt(C.prop), Ubicacion: txt(C.ubic), Detalle: (txt(C.detalle) || '').slice(0, 1000) || null,
      FechaAdjudicacion: adjCel instanceof Date ? fechaISO(adjCel) : (fechaEnTexto(texto(adjCel)) || texto(adjCel)),
      ValorMejoras: nu(C.mejoras), ValorMercado: nu(C.mercado), ValorVentaRapida: nu(C.rapida), ValorNegociado: nu(C.negociado),
      ValorLibros: nu(C.libros), _negCalc: false,
      Residen: (txt(C.residen) || '').replace(/^—$/, '') || null, QuienReside: (txt(C.quien) || '').replace(/^—$/, '') || null,
      Visitada: (txt(C.visitada) || '').replace(/^—$/, '') || null, Condicion: (txt(C.condicion) || '').replace(/^—$/, '') || null,
      Accion: (txt(C.accion) || '').replace(/^—$/, '') || null, Comentarios: (txt(C.coment) || '').slice(0, 4000) || null,
      Responsable: respTxt && respTxt !== '—' ? respTxt.slice(0, 200) : null, FechaVisita: fv, Nota: (txt(C.nota) || '').slice(0, 500) || null
    });
  }
  // Valor negociado que vino como fórmula sin resultado: se calcula con la regla del cuadro
  filas.forEach(f => {
    if (f.ValorNegociado == null && (f.ValorVentaRapida != null || f.ValorLibros != null)) {
      f.ValorNegociado = Math.round(Math.max(f.ValorVentaRapida || 0, f.ValorLibros || 0) * 1.05 * 100) / 100; f._negCalc = true;
    }
  });
  if (filas.some(f => f._negCalc)) avisos.push('El archivo no traía calculada la columna "Valor negociado" (no se guardó desde Excel después de cambiarlo): ' +
    'el sistema la calculó con su fórmula, MAX(venta rápida, libros) × 1.05.');
  // La fila TOTALES del Excel debe sumar todas las fincas
  if (filaTotal) {
    const ft = ws.getRow(filaTotal), malas = [];
    let formula = null;
    [['mejoras', 'ValorMejoras', 'mejoras'], ['mercado', 'ValorMercado', 'mercado'], ['rapida', 'ValorVentaRapida', 'venta rápida'],
      ['negociado', 'ValorNegociado', 'negociado']].forEach(([c, k, nom]) => {
      if (!C[c]) return;
      const cel = ft.getCell(C[c]), tot = sinResultado(cel) ? null : numero(valor(cel));
      const suma = Math.round(filas.reduce((a, f) => a + (f[k] || 0), 0) * 100) / 100;
      if (tot != null && Math.abs(tot - suma) > 0.05) { malas.push(nom + ' ' + suma.toLocaleString('en-US', { minimumFractionDigits: 2 })); formula = formula || (cel.value && cel.value.formula); }
    });
    if (malas.length) {
      const m = formula && String(formula).match(/[A-Z]+(\d+):[A-Z]+(\d+)/);
      const fuera = m ? filas.filter(f => f._fila < Number(m[1]) || f._fila > Number(m[2])).map(f => f.NoFinca + (f.CodigoUbic ? '-' + f.CodigoUbic : '') + ' (fila ' + f._fila + ')') : [];
      avisos.push('La fila TOTALES (fila ' + filaTotal + ') del Excel no suma todas las fincas' + (formula ? ' (=' + formula + ')' : '') +
        (fuera.length ? ': deja fuera ' + fuera.join(', ') : '') + '. Totales con las ' + filas.length + ' fincas: ' + malas.join(' · ') + '.');
    }
  }
  // El valor negociado debe ser MAX(venta rápida, libros) × 1.05: se avisa si en alguna fila no da eso
  filas.forEach(f => {
    const esperado = Math.round(Math.max(f.ValorVentaRapida || 0, f.ValorLibros || 0) * 1.05 * 100) / 100;
    if (f.ValorNegociado != null && Math.abs(esperado - f.ValorNegociado) > 0.05)
      avisos.push('Finca ' + f.NoFinca + (f.CodigoUbic ? '-' + f.CodigoUbic : '') + ' (fila ' + f._fila + '): el valor negociado (' + f.ValorNegociado.toFixed(2) +
        ') no es MAX(venta rápida, libros) × 1.05 (' + esperado.toFixed(2) + '); revisar si los valores de esa fila están en sus columnas.');
  });
  filas.forEach(f => { delete f._fila; delete f._negCalc; });
  return { hoja: ws.name, titulo, fecha, fechaVisitas, filas, avisos };
}

// ======================= CONTROL DE ALQUILERES =======================
// Meses del contrato, como en el Excel: DATEDIF(inicio, fin + 1, "m")
function mesesContrato(ini, fin) {
  const a = new Date(ini + 'T00:00:00Z'), b = new Date(fin + 'T00:00:00Z'); b.setUTCDate(b.getUTCDate() + 1);
  let m = (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + (b.getUTCMonth() - a.getUTCMonth());
  if (b.getUTCDate() < a.getUTCDate()) m--;
  return Math.max(m, 1);
}
// Fecha con ese día del mes; si el mes es más corto (30 de febrero), el último día del mes
function diaDelMes(anio, mes0, dia) {
  const ult = new Date(Date.UTC(anio, mes0 + 1, 0)).getUTCDate();
  return fechaISO(new Date(Date.UTC(anio, mes0, Math.min(dia, ult))));
}
// Calendario de pagos: por ADELANTADO el primer pago es en el día de pago desde el inicio (incluido);
// AL VENCIMIENTO, en el primer día de pago después del inicio. Una cuota por mes del contrato.
function calendarioAlquiler(c) {
  const n = Number(c.DuracionMeses) || mesesContrato(c.FechaInicio, c.FechaFin), ini = new Date(c.FechaInicio + 'T00:00:00Z'), dia = Number(c.DiaPago) || ini.getUTCDate();
  let y = ini.getUTCFullYear(), m = ini.getUTCMonth();
  let primera = diaDelMes(y, m, dia);
  if (c.Modalidad === 'VENCIMIENTO' ? primera <= c.FechaInicio : primera < c.FechaInicio) { m++; primera = diaDelMes(y, m, dia); }
  if (c.Frecuencia === 'UNICO') return [{ Numero: 1, FechaVence: primera, Monto: Number(c.Canon) }];   // un solo pago por todo el contrato
  const cuotas = [];
  for (let k = 0; k < n; k++) cuotas.push({ Numero: k + 1, FechaVence: diaDelMes(y, m + k, dia), Monto: Number(c.Canon) });
  return cuotas;
}
function modalidadDeTexto(t) {
  const k = norm(t);
  return /adelant/.test(k) ? 'ADELANTADO' : /vencim/.test(k) ? 'VENCIMIENTO' : null;
}
// Excel "Control de Alquileres": hoja de control (un contrato por fila) y una hoja por arrendatario con sus pagos
function leerAlquileres(wb) {
  let hc = null, rh = 0, mapa = null;
  wb.eachSheet(h => {
    if (hc) return;
    for (let r = 1; r <= Math.min(h.rowCount, 10); r++) {
      const m = mapaColumnas(h, r, 1, Math.max(h.columnCount, 12)), ks = Object.keys(m);
      if (ks.some(k => k.startsWith('nombredelarrendatario')) && ks.some(k => k.startsWith('canon'))) { hc = h; rh = r; mapa = m; return; }
    }
  });
  if (!hc) return null;
  const col = k => { const x = Object.keys(mapa).filter(m => m.startsWith(k))[0]; return x ? mapa[x][0] : null; };
  const C = { num: col('no'), arr: col('nombredelarrendatario'), exp: col('nombredelexpediente'), ini: col('fechainicial'), fin: col('fechafinal'),
    canon: col('canon'), dia: col('diadepago'), obs: col('observacion'), dur: col('duracion') };
  const aFecha = v => v instanceof Date ? fechaISO(v) : fechaEnTexto(texto(v));
  const contratos = [], avisos = [], notas = [];
  let fechaControl = null;
  for (let r = 1; r < rh; r++) { const v = valor(hc.getRow(r).getCell(1)); if (v instanceof Date) fechaControl = fechaISO(v); }
  for (let r = rh + 1; r <= hc.rowCount; r++) {
    const f = hc.getRow(r), arr = texto(valor(f.getCell(C.arr)));
    const a1 = texto(valor(f.getCell(1))) || '', b1 = texto(valor(f.getCell(2))) || '';
    if (/^nota/i.test(b1.trim()) || /^nota/i.test(a1.trim())) { notas.push((b1 || a1).trim()); continue; }
    if (!arr || !arr.trim()) continue;
    const ini = aFecha(valor(f.getCell(C.ini))), fin = aFecha(valor(f.getCell(C.fin)));
    const diaCel = valor(f.getCell(C.dia));
    const dia = diaCel instanceof Date ? diaCel.getUTCDate() : numero(diaCel);
    const obs = C.obs ? texto(valor(f.getCell(C.obs))) : null;
    const c = { Numero: numero(valor(f.getCell(C.num))), Arrendatario: arr.trim(), Expediente: (texto(valor(f.getCell(C.exp))) || '').trim() || null,
      FechaInicio: ini, FechaFin: fin, Canon: numero(valor(f.getCell(C.canon))), DiaPago: dia ? Math.round(dia) : null,
      Observacion: obs, Modalidad: modalidadDeTexto(obs), DuracionMeses: C.dur ? numero(valor(f.getCell(C.dur))) : null,
      Frecuencia: /\bun\s+(solo\s+|único\s+|unico\s+)?pago\b/i.test(obs || '') ? 'UNICO' : 'MENSUAL', _fila: r };
    if (c.FechaInicio && c.FechaFin) {
      const calc = mesesContrato(c.FechaInicio, c.FechaFin);
      if (!c.DuracionMeses) c.DuracionMeses = calc;
      else if (c.DuracionMeses !== calc) avisos.push(c.Arrendatario + ': el Excel dice ' + c.DuracionMeses + ' meses, pero del ' + c.FechaInicio.split('-').reverse().join('/') +
        ' al ' + c.FechaFin.split('-').reverse().join('/') + ' son ' + calc + '. Se usan ' + c.DuracionMeses + ' cuotas, como dice el Excel (se puede cambiar).');
    }
    if (!c.FechaInicio || !c.FechaFin) { avisos.push(c.Arrendatario + ' (fila ' + r + '): falta la fecha inicial o final; no se puede armar el calendario.'); }
    if (!c.Modalidad) { c.Modalidad = 'ADELANTADO'; avisos.push(c.Arrendatario + ': la observación no dice si paga por adelantado o al vencimiento; se tomó "por adelantado" (se puede cambiar).'); }
    if (!c.DiaPago && c.FechaInicio) c.DiaPago = Number(c.FechaInicio.slice(8, 10));
    contratos.push(c);
  }
  // Hojas de pagos: se unen al contrato por el nombre de la hoja (y, si no alcanza, por el título)
  const palabras = t => String(t || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(w => w.length > 2);
  const puntaje = (a, b) => { const B = palabras(b); return palabras(a).filter(w => B.indexOf(w) !== -1).length; };
  wb.eachSheet(h => {
    if (h === hc) return;
    let rp = 0;
    for (let r = 1; r <= Math.min(h.rowCount, 10); r++) {
      const t = norm(h.getRow(r).values.filter(v => v != null).map(v => texto(valor({ value: v }))).join(' '));
      if (t.indexOf('depagos') !== -1 && t.indexOf('fechadepago') !== -1) { rp = r; break; }
    }
    if (!rp) return;
    const mp = mapaColumnas(h, rp, 1, Math.max(h.columnCount, 10)), cp = k => { const x = Object.keys(mp).filter(m => m.indexOf(k) !== -1)[0]; return x ? mp[x][0] : null; };
    const P = { n: cp('depagos'), rec: cp('pagorecibido'), monto: cp('monto'), fecha: cp('fechadepago') };
    if (!P.n || !P.fecha) return;
    let titulo = '', fincaTxt = '';
    for (let r = 1; r < rp; r++) {
      const vistos = {};   // las celdas combinadas repiten el mismo texto
      const t = h.getRow(r).values.filter(v => v != null).map(v => texto(valor({ value: v }))).filter(x => x && !vistos[x] && (vistos[x] = 1)).join(' ').trim();
      if (/finca/i.test(t)) fincaTxt = t; else if (t && !titulo) titulo = t;
    }
    let mejor = null, pMejor = 0;
    contratos.forEach(c => { const p = puntaje(h.name, c.Arrendatario) * 10 + puntaje(titulo, c.Arrendatario); if (p > pMejor) { pMejor = p; mejor = c; } });
    if (!mejor || pMejor < 10) { avisos.push('La hoja "' + h.name + '" no se pudo unir a ningún contrato de la hoja de control.'); return; }
    if (titulo && puntaje(titulo, mejor.Arrendatario) === 0)
      avisos.push('Hoja "' + h.name + '": el título dice "' + titulo + '", que no es el arrendatario ' + mejor.Arrendatario + ' (se unió por el nombre de la hoja).');
    const mf = fincaTxt.match(/([A-Za-z]{0,2}\d{3,})\s*(?:-\s*([0-9A-Za-z]+))?/);
    if (mf) { mejor.NoFinca = mf[1]; mejor.CodigoUbic = mf[2] || null; }
    mejor.hoja = h.name; mejor.pagos = [];
    for (let r = rp + 1; r <= h.rowCount; r++) {
      const f = h.getRow(r), nro = numero(valor(f.getCell(P.n)));
      if (/^total/i.test(texto(valor(f.getCell(P.n))) || '')) break;
      if (!nro) continue;
      const fc = valor(f.getCell(P.fecha)), fecha = fc instanceof Date ? fechaISO(fc) : null, fechaTxt = fc instanceof Date ? null : texto(fc);
      const nota = [];
      for (let c2 = P.fecha + 1; c2 <= Math.max(h.columnCount, P.fecha + 4); c2++) { const t = texto(valor(f.getCell(c2))); if (t) nota.push(t.trim()); }
      mejor.pagos.push({ Numero: nro, Recibido: P.rec ? numero(valor(f.getCell(P.rec))) : null, Monto: P.monto ? numero(valor(f.getCell(P.monto))) : null,
        Fecha: fecha, FechaTexto: fechaTxt, Nota: nota.join(' · ') || null, _fila: r });
    }
  });
  return { hoja: hc.name, fechaControl, contratos, notas, avisos };
}
// Une lo que trae el Excel con el calendario que arma el sistema: los pagos recibidos se registran con su fecha;
// las fechas del Excel que no cuadran con el calendario se avisan (no se toman).
function cuotasDesdeExcel(c, avisos) {
  const cal = calendarioAlquiler(c), por = {};
  (c.pagos || []).forEach(p => { por[p.Numero] = p; });
  cal.forEach(q => {
    const p = por[q.Numero]; if (!p) return;
    if (p.Monto != null && p.Monto > 0) q.Monto = p.Monto;
    if (p.Recibido != null && p.Recibido > 0) { q.MontoPagado = p.Recibido; q.FechaPago = p.Fecha || null; }
    else if (p.Fecha && p.Fecha !== q.FechaVence)
      avisos.push(c.Arrendatario + ', cuota ' + q.Numero + ': el Excel dice ' + p.Fecha.split('-').reverse().join('/') + ' y el calendario da ' +
        q.FechaVence.split('-').reverse().join('/') + ' (se usa la del calendario; se puede cambiar a mano).');
    else if (!p.Fecha && p.FechaTexto) avisos.push(c.Arrendatario + ', cuota ' + q.Numero + ': la fecha "' + p.FechaTexto + '" del Excel no es una fecha válida; se usa ' +
      q.FechaVence.split('-').reverse().join('/') + '.');
    if (p.Nota) q.Nota = p.Nota.slice(0, 500);
  });
  (c.pagos || []).filter(p => p.Numero > cal.length && (p.Recibido || p.Monto)).forEach(p =>
    avisos.push(c.Arrendatario + ': la hoja trae la cuota ' + p.Numero + ', pero el contrato es de ' + cal.length + (cal.length === 1 ? ' mes' : ' meses') + ' (no se tomó).'));
  return cal;
}
function leerLibroWb(wb) {
  const avisos = [];
  // El Excel de activos es un tipo aparte: no se mezcla con ventas ni visitas
  const activos = leerActivos(wb, avisos);
  if (activos && activos.filas.length) {
    return { tipo: 'ACTIVOS', activos: activos.filas, desalojo: leerDesalojo(wb), ventasAnteriores: leerVentasAnteriores(wb), pie: activos.pie,
      fechaActivos: activos.fecha, fechaReporte: activos.fecha,
      ventas: [], visitas: [], resumen: [], trazabilidad: [], avisos };
  }
  // El cuadro de lotes con mejoras (casas) también es un tipo aparte
  const mejoras = leerMejoras(wb);
  if (mejoras && mejoras.filas.length) {
    return { tipo: 'MEJORAS', mejoras: mejoras.filas, tituloMejoras: mejoras.titulo, fechaVisitas: mejoras.fechaVisitas,
      fechaActivos: mejoras.fecha, fechaReporte: mejoras.fecha, ventas: [], visitas: [], resumen: [], trazabilidad: [], avisos: avisos.concat(mejoras.avisos) };
  }
  const ventas = leerVentas(wb, avisos);
  const visitas = leerVisitas(wb);
  const resumen = leerResumen(wb);
  const traza = leerTrazabilidad(wb);

  const tieneReporte = !!(ventas && ventas.filas.length);
  const tieneTraza = !!(traza && traza.filas.length);
  const tipo = tieneReporte && tieneTraza ? 'MIXTO' : tieneReporte ? 'REPORTE' : tieneTraza ? 'TRAZABILIDAD' : null;

  const hojas = [];
  wb.eachSheet(ws => hojas.push(ws.name));
  const leidas = [ventas, visitas, resumen, traza].filter(Boolean).map(x => x.hoja);
  const ignoradas = hojas.filter(h => !leidas.includes(h));
  if (ignoradas.length) avisos.push('Hojas que no se guardan (no tienen tabla reconocible): ' + ignoradas.join(', ') + '.');

  return {
    tipo,
    fechaReporte: (ventas && ventas.fechaReporte) || (traza && traza.fechaReporte) || null,
    fechaProceso: ventas ? ventas.fechaProceso : null,
    fechaTraza: traza ? traza.fechaReporte : null,
    ventas: ventas ? ventas.filas : [],
    visitas: visitas ? visitas.filas : [],
    resumen: resumen ? resumen.filas : [],
    trazabilidad: traza ? traza.filas : [],
    avisos
  };
}

// =============================================================================
const ORDEN = ['BR_Usuario', 'BR_Plantilla', 'BR_Carga', 'BR_Venta', 'BR_Visita', 'BR_ResumenEjecutivo', 'BR_Trazabilidad', 'BR_VentaAnterior',
  'BR_Desalojo', 'BR_Mejora', 'BR_Finca', 'BR_FincaHistorial', 'BR_VentaManual', 'BR_Gestion', 'BR_Contacto', 'BR_Documento', 'BR_Alquiler',
  'BR_AlquilerCuota', 'BR_Config', 'BR_Bitacora'];
const ARCHIVO = { BR_Carga: ['Archivo', 'NombreArchivo'], BR_Plantilla: ['Archivo', 'ArchivoNombre'], BR_Documento: ['Adjunto', 'AdjuntoNombre'] };

(async () => {
  console.log('\n=== Pasar los datos del sistema anterior al sistema de GitHub ===\n');
  const carpetas = [process.cwd(), __dirname].filter((d, i, a) => a.indexOf(d) === i);
  const modRuta = carpetas.map(d => buscar('bienesReposeidos.js', d, 3)).filter(Boolean)[0];
  if (!modRuta) { console.error('❌ No encontré bienesReposeidos.js. Pon este archivo en la carpeta del sistema anterior.'); process.exit(1); }
  const dirs = [path.dirname(modRuta)].concat(carpetas);
  const sql = requerir('mssql', dirs), ExcelJS = requerir('exceljs', dirs);
  const conf = Object.assign({}, require(modRuta)._interno.config);
  conf.options = Object.assign({}, conf.options); conf.pool = { max: 2, min: 0, idleTimeoutMillis: 30000 };
  const c = leerConfig(dirs);
  if (c) {
    if (c.cfg.password && !process.env.BIENES_DB_PASSWORD) conf.password = String(c.cfg.password);
    if (c.cfg.usuario && !process.env.BIENES_DB_USER) conf.user = String(c.cfg.usuario);
    if (c.cfg.servidor && !process.env.BIENES_DB_SERVER) conf.server = String(c.cfg.servidor);
    if (c.cfg.base && !process.env.BIENES_DB_NAME) conf.database = String(c.cfg.base);
  }
  console.log('Base del sistema anterior: ' + conf.database + ' en ' + conf.server + (c ? ' (conexión de ' + c.ruta + ')' : ''));
  const pool = await sql.connect(conf).catch(e => { console.error('❌ No pude entrar a la base: ' + e.message); process.exit(1); });

  // contar lo que hay
  const cuenta = {};
  for (const t of ORDEN) {
    const r = await pool.request().query(`SELECT COUNT(*) AS n FROM dbo.${t}`).catch(() => null);
    cuenta[t] = r ? r.recordset[0].n : null;
  }

  console.log('\nAhora entra al sistema de GitHub como administrador.');
  const usuario = (await preguntar('Usuario: ')).trim();
  const clave = await preguntar('Contraseña: ', true);
  await api('POST', '/api/bienes/entrar', { usuario, contrasena: clave }).catch(e => { console.error('❌ ' + e.message); process.exit(1); });
  const ses = await api('GET', '/api/bienes/sesion');
  if (!ses.usuario || ses.usuario.rol !== 'ADMIN') { console.error('❌ Ese usuario no es administrador.'); process.exit(1); }
  const est = await api('POST', '/api/bienes/migrar', { accion: 'estado' });
  const aqui = {}; est.tablas.forEach(x => { aqui[x.tabla] = x.filas; });

  console.log('\nTabla                   Sistema anterior   GitHub ahora');
  ORDEN.forEach(t => console.log((t + '                        ').slice(0, 24) + String(cuenta[t] == null ? '(no existe)' : cuenta[t]).padStart(16) + String(aqui[t] || 0).padStart(15)));
  console.log('\nEn GitHub se REEMPLAZAN cortes, fincas, ventas, visitas, gestiones, documentos, plantillas, alquileres y contactos');
  console.log('por los del sistema anterior. Los usuarios de GitHub se conservan y se agregan los que falten (sin contraseña).');
  console.log('Los avalúos de GitHub NO se tocan. El sistema anterior NO se modifica.');
  const si = (await preguntar('\n¿Pasar los datos? Escribe SI para continuar: ')).trim().toUpperCase();
  if (si !== 'SI' && si !== 'SÍ') { console.log('Cancelado. No se cambió nada.'); process.exit(0); }

  const resumen = [], usuariosSinClave = [];
  for (const t of ORDEN) {
    if (cuenta[t] == null) continue;
    process.stdout.write('· ' + t + ' … ');
    const arch = ARCHIVO[t];
    let filas;
    if (t === 'BR_Carga') {
      const ids = (await pool.request().query('SELECT Id FROM dbo.BR_Carga ORDER BY Id')).recordset.map(x => x.Id);
      let primera = true, n = 0;
      for (const id of ids) {
        const f = (await pool.request().input('id', sql.Int, id).query('SELECT * FROM dbo.BR_Carga WHERE Id = @id')).recordset[0];
        const o = fila(f);
        if (f.Archivo) {
          o.__archivoB64 = f.Archivo.toString('base64'); o.__archivoNombre = f.NombreArchivo || ('carga_' + id + '.xlsx');
          if (f.TipoArchivo === 'ACTIVOS' || f.TipoArchivo === 'ALQUILERES') {
            try {
              const wb = new ExcelJS.Workbook(); await wb.xlsx.load(f.Archivo);
              if (f.TipoArchivo === 'ACTIVOS') {
                const l = leerLibroWb(wb);
                if (l.activos) o.__lectura = JSON.parse(JSON.stringify({ activos: l.activos, ventasAnteriores: l.ventasAnteriores || [], desalojo: l.desalojo || [],
                  pie: l.pie || null, fechaReporte: o.FechaReporte || l.fechaReporte }));
              } else {
                const lx = leerAlquileres(wb); if (lx) o.__lectura = JSON.parse(JSON.stringify(lx));
              }
            } catch (e) { console.log('\n  ⚠ no pude leer el Excel de la carga #' + id + ': ' + e.message); }
          }
        }
        await api('POST', '/api/bienes/migrar', { tabla: t, primera, filas: [o] });
        primera = false; n++;
        process.stdout.write(n % 5 ? '' : n + ' ');
      }
      if (!ids.length) await api('POST', '/api/bienes/migrar', { tabla: t, primera: true, filas: [] });
      console.log(n + ' cortes con sus archivos'); resumen.push(t + ' ' + n);
      continue;
    }
    filas = (await pool.request().query(`SELECT * FROM dbo.${t}`)).recordset.map(f => {
      const o = fila(f);
      if (arch && f[arch[0]]) { o.__archivoB64 = f[arch[0]].toString('base64'); o.__archivoNombre = f[arch[1]] || 'archivo'; }
      return o;
    });
    if (t === 'BR_Usuario') filas.forEach(f => delete f.Hash);
    let primera = true, n = 0, salt = [];
    for (const tanda of tandas(filas)) {
      const r = await api('POST', '/api/bienes/migrar', { tabla: t, primera, filas: tanda });
      primera = false; n += r.insertadas; salt = salt.concat(r.saltados || []);
    }
    if (t === 'BR_Usuario') filas.forEach(f => { if (!salt.some(s => s.indexOf(f.Usuario + ' ') === 0)) usuariosSinClave.push(f.Usuario + ' (' + f.Nombre + ')'); });
    console.log(n + ' filas' + (salt.length ? ' · ya estaban: ' + salt.join(', ') : ''));
    resumen.push(t + ' ' + n);
  }
  await api('POST', '/api/bienes/migrar', { accion: 'fin', resumen: resumen.join(', ') });
  await pool.close();
  console.log('\n✅ Listo. Los datos del sistema anterior ya están en el sistema de GitHub.');
  if (usuariosSinClave.length) {
    console.log('\nEstos usuarios se agregaron SIN contraseña. En la página, Usuarios → Editar, ponles una temporal:');
    usuariosSinClave.forEach(u => console.log('   - ' + u));
  }
  process.exit(0);
})().catch(e => { console.error('\n❌ ' + (e && e.message ? e.message : e)); process.exit(1); });
