/* =============================================================================
   BIENES REPOSEÍDOS — Servidor en Google Apps Script sobre Google Sheets
   -----------------------------------------------------------------------------
   Réplica de bienesReposeidos.js (Node + SQL Server) para correr sin servidor:
     · la base son las pestañas de esta hoja de Google (una por tabla, BR_...);
     · los archivos (Excel originales, modelos de plantilla, adjuntos y PDF de
       avalúos) se guardan en una carpeta de Google Drive;
     · la página (index.html en GitHub Pages) habla con este script.

   Cómo se instala (una sola vez), en esta hoja:
     1. Extensiones → Apps Script. Pega este archivo como Code.gs y, aparte,
        el archivo Plantillas.gs (las plantillas de fábrica de documentos).
     2. Recarga la hoja: aparece el menú "Bienes reposeídos".
        Menú → "Instalar" crea las pestañas, la carpeta de archivos y te muestra
        la CLAVE DE INSTALACIÓN (la pide la página para crear el administrador).
     3. Implementar → Nueva implementación → Aplicación web
        (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
        La URL que termina en /exec va en index.html (APPS_SCRIPT_URL).

   Roles: ADMIN (todo, usuarios y bitácora) · GESTOR (sube y edita) · CONSULTA (ve y descarga).
   ============================================================================= */

// ID de la hoja (entre /d/ y /edit en su dirección). Vacío = la hoja donde está pegado el script.
var HOJA_ID = '';
var ZONA = 'America/Panama';

var HORAS_SESION = 8;          // la sesión vence sola a las 8 horas
var MAX_INTENTOS = 5;          // contraseñas malas seguidas antes de bloquear
var MINUTOS_BLOQUEO = 15;      // tiempo de bloqueo del usuario
var MIN_LARGO_CLAVE = 8;
var ITERACIONES_CLAVE = 1000;  // repeticiones del SHA-256 al cifrar contraseñas
var ROLES = ['ADMIN', 'GESTOR', 'CONSULTA'];
var MAX_BYTES = 20 * 1024 * 1024;
var VERSION_SERVIDOR = '2026-11-14';
// ===== Utilidades copiadas tal cual de bienesReposeidos.js =====
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
// ===========================================================================
// CÁLCULOS DEL CUADRO
// ===========================================================================
function entero(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return Number.isInteger(n) && n >= 0 ? n : NaN;
}

// ventas: filas de BR_Venta; conteos: { totalActivos, enLegal, alquiladas, conMejora }
function calcularCuadro(ventas, conteos) {
  const setDe = t => new Set(ventas.filter(v => v.Tabla === t).map(v => String(v.NoFinca)));
  const proceso = setDe('EN_PROCESO');
  const finalizadas = setDe('FINALIZADA');
  // El Excel llama "en proceso" a la lista de TODAS las ventas registradas;
  // las que ya aparecen en "venta finalizada" son vendidas, el resto sigue en trámite.
  const registradas = new Set([...proceso, ...finalizadas]);
  const vendidas = finalizadas.size;
  const enProceso = registradas.size - vendidas;

  const c = conteos || {};
  const totalActivos = c.totalActivos === undefined ? null : c.totalActivos;
  const enLegal = c.enLegal === undefined ? null : c.enLegal;
  const alquiladas = c.alquiladas === undefined ? null : c.alquiladas;
  const conMejora = c.conMejora === undefined ? null : c.conMejora;

  const disponibles = totalActivos !== null && enLegal !== null ? totalActivos - registradas.size - enLegal : null;
  const sinMejora = disponibles !== null && conMejora !== null ? disponibles - conMejora : null;

  const avisos = [];
  if (disponibles !== null && disponibles < 0) avisos.push('Total de activos menor que ventas + legal: los disponibles quedan en negativo.');
  if (sinMejora !== null && sinMejora < 0) avisos.push('Las propiedades con mejora son más que las disponibles.');
  if (alquiladas !== null && disponibles !== null && alquiladas > disponibles) avisos.push('Hay más alquiladas que disponibles.');
  const enProcesoLista = [...proceso].filter(f => !finalizadas.has(f));
  const soloFinal = [...finalizadas].filter(f => !proceso.has(f));
  if (soloFinal.length) avisos.push('Fincas vendidas que no están en la lista de ventas registradas: ' + soloFinal.join(', ') + '.');

  return {
    totalActivos, enTramiteOVendidos: registradas.size, enLegal, disponibles, alquiladas,
    conMejora, sinMejora,
    ventasRegistradas: registradas.size, vendidas, enProceso,
    fincasEnProceso: enProcesoLista,
    avisos
  };
}

function validarConteos(body) {
  const c = {
    totalActivos: entero(body && body.totalActivos),
    enLegal: entero(body && body.enLegal),
    alquiladas: entero(body && body.alquiladas),
    conMejora: entero(body && body.conMejora)
  };
  const malos = Object.keys(c).filter(k => Number.isNaN(c[k]));
  return { conteos: c, error: malos.length ? 'Los conteos deben ser números enteros de 0 en adelante: ' + malos.join(', ') + '.' : null };
}

// Diferencias entre el reporte de ventas y la trazabilidad, finca por finca
function cruzar(ventas, traza) {
  if (!ventas.length || !traza.length) return [];
  const porFinca = {};
  // En el reporte la misma finca puede venir en "en proceso" y en "finalizada";
  // se toma la de finalizada si existe.
  for (const v of ventas) {
    if (v.Tabla === 'NO_CONCRETADA') continue;
    if (!porFinca[v.NoFinca] || v.Tabla === 'FINALIZADA') porFinca[v.NoFinca] = v;
  }
  const dif = [];
  const campos = [
    ['ValorVentaNegociado', 'Valor venta negociado'],
    ['ValorLibros', 'Valor en libros'],
    ['MontoFinanciado', 'Monto financiado']
  ];
  for (const t of traza) {
    const v = porFinca[t.NoFinca];
    if (!v) { dif.push({ NoFinca: t.NoFinca, campo: 'Finca', reporte: null, trazabilidad: 'está', nota: 'No aparece en el reporte de ventas' }); continue; }
    for (const [k, nombre] of campos) {
      const a = v[k] === null || v[k] === undefined ? null : Number(v[k]);
      const b = t[k] === null || t[k] === undefined ? null : Number(t[k]);
      if (a === null || b === null) continue;
      if (Math.abs(a - b) > 0.5) dif.push({ NoFinca: t.NoFinca, campo: nombre, reporte: a, trazabilidad: b });
    }
    // Forma de pago: financiado + abono + contado debería dar el valor negociado
    const pago = (Number(t.MontoFinanciado) || 0) + (Number(t.Abono) || 0) + (Number(t.Contado) || 0);
    const vn = Number(t.ValorVentaNegociado) || 0;
    if (vn && Math.abs(pago - vn) > 0.5) {
      dif.push({ NoFinca: t.NoFinca, campo: 'Forma de pago', reporte: null, trazabilidad: pago, venta: vn,
        nota: 'Financiado + abono + contado = ' + pago.toFixed(2) + ' y la venta negociada es ' + vn.toFixed(2) });
    }
  }
  return dif;
}

function totalesTraza(filas) {
  const campos = ['ImpuestoInmueble', 'GastoMantenimiento', 'GastosLegales', 'Impuesto2', 'Impuesto3', 'Comision',
    'TotalGastos', 'ValorLibros', 'ValorVentaRapida', 'ValorVentaNegociado', 'Balance', 'MontoFinanciado', 'Abono', 'Contado'];
  const suma = (lista) => {
    const o = { fincas: lista.length };
    for (const c of campos) o[c] = Math.round(lista.reduce((a, f) => a + (Number(f[c]) || 0), 0) * 100) / 100;
    return o;
  };
  return {
    VENDIDA: suma(filas.filter(f => f.Grupo === 'VENDIDA')),
    EN_PROCESO: suma(filas.filter(f => f.Grupo === 'EN_PROCESO')),
    TOTAL: suma(filas)
  };
}
// Limpieza del HTML que se guarda (quita scripts, eventos y enlaces peligrosos)
function limpiarHtml(html) {
  let s = String(html || '');
  s = s.replace(/<!--[\s\S]*?-->/g, '');
  s = s.replace(/<\s*(script|style|iframe|object|textarea|select|button|noscript|template|applet|frameset)\b[\s\S]*?<\s*\/\s*\1\s*>/gi, '');
  s = s.replace(/<\/?\s*(script|style|iframe|object|embed|link|meta|base|form|input|textarea|select|button|noscript|template|frame|frameset|applet)\b[^>]*>/gi, '');
  s = s.replace(/\s+on[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '');
  s = s.replace(/(href|src|xlink:href|action|formaction|background)\s*=\s*(["']?)\s*(javascript|vbscript|data(?!:image\/))\s*:[^"'\s>]*\2/gi, '$1="#"');
  s = s.replace(/expression\s*\(/gi, '(');
  return s;
}
function validarClaveNueva(clave, usuario) {
  const c = String(clave || '');
  if (c.length < MIN_LARGO_CLAVE) return 'La contraseña debe tener al menos ' + MIN_LARGO_CLAVE + ' caracteres.';
  if (!/[a-zA-Z]/.test(c) || !/\d/.test(c)) return 'La contraseña debe tener letras y números.';
  if (usuario && c.toLowerCase().indexOf(String(usuario).toLowerCase()) !== -1) return 'La contraseña no puede contener el nombre de usuario.';
  return null;
}

// =============================================================================
// BASE DE DATOS SOBRE LA HOJA
// Cada tabla es una pestaña. Tipos: I entero · N decimal · B sí/no · F fecha ·
// T fecha y hora · S texto · J objeto (JSON). Todo se guarda como texto para que
// Sheets no cambie números de finca ni fechas.
// Las filas se entregan igual que las devolvía SQL Server (mssql): fechas como
// Date (la hora de Panamá marcada como UTC), sí/no como true/false, números.
// =============================================================================
var ESQ = {
  BR_Carga: 'Id:I FechaCarga:T Usuario:S UsuarioId:I NombreArchivo:S TipoArchivo:S FechaReporte:F HashArchivo:S TamanoBytes:I ArchivoId:S ' +
    'TotalActivos:I EnLegal:I Alquiladas:I ConMejora:I Nota:S Anulada:B FechaAnulada:T UsuarioAnula:S Lectura:J',
  BR_Venta: 'Id:I CargaId:I Tabla:S FechaTabla:F Orden:I NoFinca:S CodigoUbic:S Ubicacion:S Superficie:S ValorVentaNegociado:N ValorLibros:N ' +
    'ValorVentaRapida:N MontoFinanciado:N Cliente:S Ejecutivo:S Estatus:S AntiguoPropietario:S ComisionPagada:N',
  BR_Visita: 'Id:I CargaId:I Orden:I NoFinca:S CodigoUbic:S Ubicacion:S TipoPropiedad:S Condicion:S Letrero:S Observaciones:S Limpieza:S ProximaLimpieza:S',
  BR_ResumenEjecutivo: 'Id:I CargaId:I Orden:I Seccion:S ClienteFinca:S Estatus:S',
  BR_Trazabilidad: 'Id:I CargaId:I Grupo:S Orden:I NoFinca:S FechaAdjudicacion:S FechaAdjudicacionD:F Ubicacion:S TipoPropiedad:S ImpuestoInmueble:N ' +
    'GastoMantenimiento:N GastosLegales:N Impuesto2:N Impuesto3:N Comision:N TotalGastos:N ValorLibros:N ValorVentaRapida:N ValorVentaNegociado:N ' +
    'Balance:N MontoFinanciado:N Abono:N Contado:N',
  BR_Usuario: 'Id:I Usuario:S Nombre:S Rol:S Hash:S Activo:B DebeCambiar:B Intentos:I BloqueadoHasta:T SesionVersion:I FechaCreacion:T CreadoPor:S UltimoIngreso:T',
  BR_Config: 'Clave:S Valor:S',
  BR_Bitacora: 'Id:I Fecha:T UsuarioId:I Usuario:S Accion:S Detalle:S IP:S',
  BR_Gestion: 'Id:I NoFinca:S Fecha:F Tipo:S Lugar:S Descripcion:S HayInteresado:B Cliente:S Contacto:S MontoOferta:N FormaPago:S EstadoOferta:S ' +
    'ProximoPaso:S FechaProximo:F UsuarioId:I Usuario:S FechaRegistro:T UsuarioModifica:S FechaModifica:T Anulada:B Canal:S Enlace:S Consultas:I',
  BR_Plantilla: 'Id:I Codigo:S Nombre:S Tipo:S Contenido:S Activa:B ArchivoNombre:S ArchivoId:S UsuarioCrea:S FechaCrea:T UsuarioModifica:S FechaModifica:T',
  BR_Documento: 'Id:I PlantillaId:I Tipo:S Titulo:S NoFinca:S Cliente:S Contenido:S Campos:S Estado:S Version:I AdjuntoNombre:S AdjuntoId:S ' +
    'UsuarioCrea:S FechaCrea:T UsuarioModifica:S FechaModifica:T Anulado:B',
  BR_Finca: 'NoFinca:S Estado:S Alquilada:B Mejora:S TipoPropiedad:S Ubicacion:S CodigoUbic:S ValorLibros:N Nota:S Activa:B UsuarioCrea:S FechaCrea:T ' +
    'UsuarioModifica:S FechaModifica:T Ocupada:B ValorMercado:N ValorVentaRapida:N ValorMejoras:N Deudor:S FechaAdjudicacion:S Superficie:S Detalle:S ' +
    'Letrero:S Observaciones:S Dias:I Fideicomiso:S ValorTerreno:N Deterioro:N Provision:N EstadoExcel:S OrdenExcel:I NoSeVende:B PolizaIncendio:B',
  BR_VentaAnterior: 'Id:I CargaId:I Anio:I Orden:I NoFinca:S CodigoUbic:S FechaAdjudicacion:S Deudor:S Comprador:S Ubicacion:S Detalle:S Tipo:S ' +
    'ValorLibros:N ValorMercado:N ValorVentaRapida:N ValorVenta:N MontoFinanciado:N Comision:N Ejecutivo:S Observaciones:S',
  BR_Desalojo: 'Id:I CargaId:I Orden:I NoFinca:S Seccion:S AntiguoPropietario:S Ubicacion:S QuienReside:S Accion:S Situacion:S UltimaGestion:S ' +
    'FechaVisita:F ValorMercado:N ValorVentaRapida:N ValorLibros:N EstatusDesalojo:S FechaNotificacion:S FechaLimite:S Responsable:S ProximaAccion:S',
  BR_Alquiler: 'Id:I Numero:I Arrendatario:S Expediente:S NoFinca:S CodigoUbic:S FechaInicio:F FechaFin:F Canon:N DiaPago:I Modalidad:S ' +
    'DuracionMeses:I Observacion:S Anulado:B Origen:S CreadoPor:S FechaCrea:T ModificadoPor:S FechaModifica:T Frecuencia:S',
  BR_AlquilerCuota: 'Id:I AlquilerId:I Numero:I FechaVence:F Monto:N FechaPago:F MontoPagado:N Nota:S FechaManual:B RegistradoPor:S FechaRegistro:T',
  BR_Mejora: 'Id:I CargaId:I Orden:I NoFinca:S CodigoUbic:S Activo:B AntiguoPropietario:S Ubicacion:S Detalle:S FechaAdjudicacion:S ValorMejoras:N ' +
    'ValorMercado:N ValorVentaRapida:N ValorNegociado:N ValorLibros:N Residen:S QuienReside:S Visitada:S Condicion:S Accion:S Comentarios:S ' +
    'Responsable:S FechaVisita:F Nota:S',
  BR_FincaHistorial: 'Id:I NoFinca:S Campo:S ValorAnterior:S ValorNuevo:S Usuario:S Fecha:T',
  BR_VentaManual: 'Id:I NoFinca:S Tabla:S Fecha:F CodigoUbic:S Ubicacion:S Superficie:S ValorVentaNegociado:N ValorLibros:N ValorVentaRapida:N ' +
    'MontoFinanciado:N Cliente:S Ejecutivo:S Estatus:S AntiguoPropietario:S ComisionPagada:N UsuarioCrea:S FechaCrea:T UsuarioModifica:S ' +
    'FechaModifica:T Anulada:B',
  BR_Contacto: 'Id:I Fecha:F Canal:S Nombre:S Contacto:S NoFinca:S Mensaje:S Estado:S Seguimiento:S Usuario:S FechaRegistro:T UsuarioModifica:S ' +
    'FechaModifica:T Anulado:B',
  // Nuevo: avalúos (PDF en SharePoint/OneDrive o subidos aquí), unidos a su finca
  BR_Avaluo: 'Id:I Archivo:S Enlace:S Titular:S FincasManual:S FechaAvaluo:F Nota:S Origen:S ArchivoId:S UsuarioCrea:S FechaCrea:T ' +
    'UsuarioModifica:S FechaModifica:T Anulado:B',
  _Textos: 'Clave:S Parte:I Texto:S'
};
// Valores por omisión (los DEFAULT de las tablas SQL)
var POR_OMISION = {
  BR_Carga: { FechaCarga: 'AHORA', Anulada: false },
  BR_Usuario: { Activo: true, DebeCambiar: true, Intentos: 0, SesionVersion: 1, FechaCreacion: 'AHORA' },
  BR_Bitacora: { Fecha: 'AHORA' },
  BR_Gestion: { HayInteresado: false, FechaRegistro: 'AHORA', Anulada: false },
  BR_Plantilla: { Activa: true, FechaCrea: 'AHORA' },
  BR_Documento: { Estado: 'BORRADOR', Version: 1, FechaCrea: 'AHORA', Anulado: false },
  BR_Finca: { Alquilada: false, Activa: true, FechaCrea: 'AHORA', Ocupada: false, NoSeVende: false, PolizaIncendio: false },
  BR_Alquiler: { Modalidad: 'ADELANTADO', Anulado: false, Origen: 'MANUAL', FechaCrea: 'AHORA', Frecuencia: 'MENSUAL' },
  BR_AlquilerCuota: { FechaManual: false },
  BR_FincaHistorial: { Fecha: 'AHORA' },
  BR_VentaManual: { FechaCrea: 'AHORA', Anulada: false },
  BR_Contacto: { FechaRegistro: 'AHORA', Anulado: false },
  BR_Avaluo: { FechaCrea: 'AHORA', Anulado: false }
};
var MAX_CELDA = 45000;   // una celda de Sheets aguanta 50.000 caracteres; lo que pase va a la pestaña _Textos

var DB = {
  cache: {},        // tabla -> { cols, tipos, filas, nuevas[], sucias{}, rehacer }
  textos: null,     // clave -> texto (de _Textos)
  textosSucio: false,
  claveDe: {},      // texto largo -> su clave en _Textos
  libro: null
};

function libro_() {
  if (!DB.libro) DB.libro = HOJA_ID ? SpreadsheetApp.openById(HOJA_ID) : SpreadsheetApp.getActiveSpreadsheet();
  return DB.libro;
}
function esquema_(tabla) {
  var cols = [], tipos = {};
  ESQ[tabla].split(/\s+/).forEach(function (p) { var x = p.split(':'); cols.push(x[0]); tipos[x[0]] = x[1]; });
  return { cols: cols, tipos: tipos };
}
function hojaDe_(tabla) {
  var sh = libro_().getSheetByName(tabla);
  var e = esquema_(tabla);
  if (!sh) {
    sh = libro_().insertSheet(tabla);
    sh.getRange(1, 1, 1, e.cols.length).setValues([e.cols]).setFontWeight('bold').setBackground('#12365b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), e.cols.length).setNumberFormat('@');
  }
  return sh;
}

// ---- conversión celda <-> valor ----
function ahora_() {
  var s = Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd'T'HH:mm:ss");
  return new Date(s + 'Z');   // hora de Panamá marcada como UTC (como la daba SQL Server)
}
function hoyISO_() { return fechaISO(ahora_()); }
function aFecha_(v, conHora) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], conHora && m[4] ? +m[4] : 0, conHora && m[5] ? +m[5] : 0, conHora && m[6] ? +m[6] : 0));
}
function deCelda_(tipo, v) {
  if (v === '' || v === null || v === undefined) return tipo === 'B' ? false : null;
  if (typeof v === 'string' && v.indexOf('@@txt:') === 0) v = leerTexto_(v.slice(6));
  switch (tipo) {
    case 'I': case 'N': { var n = Number(v); return isFinite(n) ? n : null; }
    case 'B': return v === true || String(v).toUpperCase() === 'TRUE' || v === 1 || v === '1';
    case 'F': return aFecha_(v, false);
    case 'T': return aFecha_(v, true);
    case 'J': try { return typeof v === 'string' ? JSON.parse(v) : v; } catch (e) { return null; }
    default: return v instanceof Date ? Utilities.formatDate(v, ZONA, 'yyyy-MM-dd') : String(v);
  }
}
function aCelda_(tipo, v) {
  if (v === null || v === undefined || v === '') return tipo === 'B' ? 'FALSE' : '';
  var s;
  switch (tipo) {
    case 'I': case 'N': { var n = Number(v); return isFinite(n) ? String(n) : ''; }
    case 'B': return v ? 'TRUE' : 'FALSE';
    case 'F': { var d = aFecha_(v, false); return d ? fechaISO(d) : ''; }
    case 'T': { var t = aFecha_(v, true); return t ? fechaISO(t) + ' ' + dosDig(t.getUTCHours()) + ':' + dosDig(t.getUTCMinutes()) + ':' + dosDig(t.getUTCSeconds()) : ''; }
    case 'J': s = JSON.stringify(v); break;
    default: s = String(v);
  }
  if (s.length > MAX_CELDA) return '@@txt:' + guardarTexto_(s);
  return s;
}

// ---- textos largos (_Textos) ----
function cargarTextos_() {
  if (DB.textos) return;
  DB.textos = {};
  var sh = hojaDe_('_Textos'), n = sh.getLastRow();
  if (n < 2) return;
  var V = sh.getRange(2, 1, n - 1, 3).getValues();
  var partes = {};
  V.forEach(function (r) { var k = String(r[0]); (partes[k] = partes[k] || []).push([Number(r[1]), String(r[2])]); });
  Object.keys(partes).forEach(function (k) { DB.textos[k] = partes[k].sort(function (a, b) { return a[0] - b[0]; }).map(function (x) { return x[1]; }).join(''); });
}
function leerTexto_(k) { cargarTextos_(); var s = DB.textos[k] || ''; DB.claveDe[s] = k; return s; }
function guardarTexto_(s) {
  cargarTextos_();
  if (DB.claveDe[s] && DB.textos[DB.claveDe[s]] === s) return DB.claveDe[s];   // el mismo texto que ya estaba guardado
  var k = Utilities.getUuid(); DB.textos[k] = s; DB.claveDe[s] = k; DB.textosSucio = true; return k;
}

// ---- lectura de tablas ----
function T_(tabla) {
  if (DB.cache[tabla]) return DB.cache[tabla].filas;
  var e = esquema_(tabla), sh = hojaDe_(tabla);
  var nCol = Math.max(sh.getLastColumn(), 1);
  var enc = sh.getRange(1, 1, 1, nCol).getValues()[0].map(String);
  // columnas nuevas del esquema que la pestaña todavía no tiene
  var faltan = e.cols.filter(function (c) { return enc.indexOf(c) === -1; });
  if (faltan.length) {
    sh.getRange(1, enc.length + 1, 1, faltan.length).setValues([faltan]).setFontWeight('bold').setBackground('#12365b').setFontColor('#ffffff');
    sh.getRange(1, enc.length + 1, sh.getMaxRows(), faltan.length).setNumberFormat('@');
    enc = enc.concat(faltan);
  }
  var n = sh.getLastRow(), filas = [];
  if (n > 1) {
    var V = sh.getRange(2, 1, n - 1, enc.length).getValues();
    V.forEach(function (r, i) {
      if (r.every(function (x) { return x === '' || x === null; })) return;
      var o = {};
      enc.forEach(function (c, j) { if (e.tipos[c]) o[c] = deCelda_(e.tipos[c], r[j]); });
      Object.defineProperty(o, '_r', { value: i + 2, writable: true, enumerable: false });
      filas.push(o);
    });
  }
  DB.cache[tabla] = { enc: enc, tipos: e.tipos, filas: filas, nuevas: [], sucias: [], rehacer: false };
  return filas;
}
function sigId_(tabla) {
  var p = PropertiesService.getScriptProperties(), k = 'SEQ_' + tabla;
  var max = T_(tabla).reduce(function (a, f) { return Math.max(a, Number(f.Id) || 0); }, 0);
  var n = Math.max(Number(p.getProperty(k)) || 0, max) + 1;
  p.setProperty(k, String(n));
  return n;
}
// INSERT: aplica los valores por omisión y el Id; devuelve la fila
function ins_(tabla, obj) {
  var c = DB.cache[tabla] || (T_(tabla), DB.cache[tabla]);
  var o = {}, def = POR_OMISION[tabla] || {};
  Object.keys(c.tipos).forEach(function (k) {
    var v = obj[k];
    if (v === undefined || v === null) v = def[k] === 'AHORA' ? ahora_() : (def[k] !== undefined ? def[k] : null);
    o[k] = normalizarTipo_(c.tipos[k], v);
  });
  if (c.tipos.Id === 'I' && !obj.Id) o.Id = sigId_(tabla);
  Object.defineProperty(o, '_r', { value: null, writable: true, enumerable: false });
  c.filas.push(o); c.nuevas.push(o);
  return o;
}
function normalizarTipo_(t, v) {
  if (v === null || v === undefined || v === '') return t === 'B' ? false : null;
  if (t === 'F') return aFecha_(v, false);
  if (t === 'T') return aFecha_(v, true);
  if (t === 'B') return !!v;
  if (t === 'I' || t === 'N') { var n = Number(v); return isFinite(n) ? n : null; }
  if (t === 'J') return v;
  return String(v);
}
// UPDATE de una fila ya leída
function upd_(tabla, fila, cambios) {
  var c = DB.cache[tabla];
  Object.keys(cambios || {}).forEach(function (k) { if (c.tipos[k]) fila[k] = normalizarTipo_(c.tipos[k], cambios[k]); });
  if (fila._r && c.sucias.indexOf(fila) === -1 && c.nuevas.indexOf(fila) === -1) c.sucias.push(fila);
  return fila;
}
// DELETE de las filas que cumplan la condición; devuelve cuántas
function del_(tabla, pred) {
  var filas = T_(tabla), c = DB.cache[tabla];
  var quedan = filas.filter(function (f) { return !pred(f); });
  var n = filas.length - quedan.length;
  if (n) { c.filas.length = 0; Array.prototype.push.apply(c.filas, quedan); c.rehacer = true; }
  return n;
}
function aFila_(c, o) { return c.enc.map(function (k) { return c.tipos[k] ? aCelda_(c.tipos[k], o[k]) : ''; }); }

// Escribe en la hoja todo lo que cambió en esta petición
function guardarTodo_() {
  Object.keys(DB.cache).forEach(function (tabla) {
    var c = DB.cache[tabla];
    if (!c.rehacer && !c.nuevas.length && !c.sucias.length) return;
    var sh = hojaDe_(tabla), nc = c.enc.length;
    if (c.rehacer || c.sucias.length > 40) {
      var n = sh.getLastRow();
      if (n > 1) sh.getRange(2, 1, n - 1, nc).clearContent();
      if (c.filas.length) {
        var rg = sh.getRange(2, 1, c.filas.length, nc); rg.setNumberFormat('@');
        rg.setValues(c.filas.map(function (o) { return aFila_(c, o); }));
      }
      c.filas.forEach(function (o, i) { o._r = i + 2; });
    } else {
      c.sucias.forEach(function (o) { var rg = sh.getRange(o._r, 1, 1, nc); rg.setNumberFormat('@'); rg.setValues([aFila_(c, o)]); });
      if (c.nuevas.length) {
        var ini = sh.getLastRow() + 1, rg2 = sh.getRange(ini, 1, c.nuevas.length, nc);
        rg2.setNumberFormat('@');
        rg2.setValues(c.nuevas.map(function (o) { return aFila_(c, o); }));
        c.nuevas.forEach(function (o, i) { o._r = ini + i; });
      }
    }
    c.nuevas = []; c.sucias = []; c.rehacer = false;
  });
  if (DB.textosSucio) guardarTextos_();
}
// Reescribe _Textos solo con los textos que todavía usa alguna celda
function guardarTextos_() {
  var usados = {};
  libro_().getSheets().forEach(function (sh) {
    if (!/^BR_/.test(sh.getName()) || sh.getLastRow() < 2) return;
    var t = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).createTextFinder('@@txt:').findAll();
    t.forEach(function (r) { usados[String(r.getValue()).slice(6)] = 1; });
  });
  var sh = hojaDe_('_Textos'), n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, 3).clearContent();
  var filas = [];
  Object.keys(DB.textos).forEach(function (k) {
    if (!usados[k]) return;
    var s = DB.textos[k];
    for (var i = 0, p = 0; i < s.length; i += MAX_CELDA, p++) filas.push([k, String(p), s.slice(i, i + MAX_CELDA)]);
  });
  if (filas.length) { var rg = sh.getRange(2, 1, filas.length, 3); rg.setNumberFormat('@'); rg.setValues(filas); }
  DB.textosSucio = false;
}

// ---- configuración (BR_Config) ----
function cfg_(clave) { var f = T_('BR_Config').filter(function (x) { return x.Clave === clave; })[0]; return f ? f.Valor : null; }
function ponerCfg_(clave, valor) {
  var f = T_('BR_Config').filter(function (x) { return x.Clave === clave; })[0];
  if (f) upd_('BR_Config', f, { Valor: valor }); else ins_('BR_Config', { Clave: clave, Valor: valor });
}
function quitarCfg_(claves) { del_('BR_Config', function (x) { return claves.indexOf(x.Clave) !== -1; }); }

// ---- archivos en Google Drive ----
function carpeta_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('CARPETA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* se borró: se crea otra */ } }
  var c = DriveApp.createFolder('Bienes reposeídos – archivos');
  p.setProperty('CARPETA_ID', c.getId());
  return c;
}
var TIPOS_MIME = { pdf: 'application/pdf', doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', csv: 'text/csv' };
function extension_(n) { var m = String(n || '').match(/\.([a-z0-9]+)$/i); return m ? m[1].toLowerCase() : ''; }
function guardarArchivo_(nombre, b64) {
  var bytes = Utilities.base64Decode(String(b64).replace(/^data:.*;base64,/, ''));
  var f = carpeta_().createFile(Utilities.newBlob(bytes, TIPOS_MIME[extension_(nombre)] || 'application/octet-stream', nombre));
  return { id: f.getId(), bytes: bytes.length };
}
function leerArchivo_(id) {
  var b = DriveApp.getFileById(id).getBlob();
  return Utilities.base64Encode(b.getBytes());
}
function borrarArchivo_(id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* ya no estaba */ } }
// Respuesta de descarga: la página la convierte en archivo
function archivoRespuesta_(nombre, b64) {
  return { success: true, __archivo: { nombre: nombre, tipo: TIPOS_MIME[extension_(nombre)] || 'application/octet-stream', b64: b64 } };
}

// =============================================================================
// ENTRADA WEB, RUTAS Y SEGURIDAD
// La página manda: { m: método, p: ruta (/api/bienes/...), q: query, b: cuerpo, t: token }
// y recibe el mismo JSON que daba el servidor Node.
// =============================================================================
ESQ._Registro = 'Fecha:T Etapa:S Detalle:S';

function doGet(e) {
  return salida_({ success: true, servicio: 'Bienes reposeídos', version: VERSION_SERVIDOR });
}
function doPost(e) {
  var req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return salida_({ success: false, status: 400, error: 'Solicitud inválida.' }); }
  return salida_(atender_(req));
}
function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Error "para el usuario": se responde con su mensaje y no se guarda nada de esa petición
function err_(msg, status, extra) { var e = new Error(msg); e.status = status || 400; e.extra = extra || null; return e; }
// Respuesta de error que SÍ guarda lo anotado antes (intentos fallidos, bitácora)
function falla_(msg, status, extra) { return Object.assign({ success: false, status: status || 400, error: msg }, extra || {}); }

var LOCK_ = null;
function conLock_() {
  if (LOCK_) return;
  LOCK_ = LockService.getScriptLock();
  if (!LOCK_.tryLock(30000)) { LOCK_ = null; throw err_('El sistema está ocupado guardando otro cambio. Prueba otra vez en unos segundos.', 503); }
}

function atender_(req) {
  var metodo = String(req.m || 'GET').toUpperCase();
  var partes = String(req.p || '').split('?');
  var ruta = partes[0], query = Object.assign({}, req.q || {});
  if (partes[1]) partes[1].split('&').forEach(function (kv) {
    if (!kv) return; var i = kv.indexOf('='); var k = decodeURIComponent(i < 0 ? kv : kv.slice(0, i));
    query[k] = i < 0 ? '' : decodeURIComponent(kv.slice(i + 1).replace(/\+/g, ' '));
  });
  try {
    if (metodo !== 'GET') conLock_();     // las escrituras van de una en una
    var r = despachar_(metodo, ruta, query, req.b || {}, req.t);
    if (hayCambios_()) { conLock_(); guardarTodo_(); }
    return r;
  } catch (e) {
    if (e && e.status) return Object.assign({ success: false, status: e.status, error: e.message }, e.extra || {});
    anotar_('error ' + metodo + ' ' + ruta, e);
    return { success: false, status: 500, error: 'Error del servidor: ' + (e && e.message || e) };
  } finally {
    if (LOCK_) { try { LOCK_.releaseLock(); } catch (x) { /* nada */ } LOCK_ = null; }
  }
}
function hayCambios_() {
  return DB.textosSucio || Object.keys(DB.cache).some(function (t) { var c = DB.cache[t]; return c.rehacer || c.nuevas.length || c.sucias.length; });
}

// ---- tabla de rutas ----
var RUTAS = [];
// acceso: null (público) · 'cualquiera' · 'gestor' (ADMIN o GESTOR) · 'admin'
function ruta_(metodo, patron, acceso, fn) {
  var nombres = [];
  var re = new RegExp('^' + patron.replace(/:([a-zA-Z]+)/g, function (_, n) { nombres.push(n); return '([^/]+)'; }) + '$');
  RUTAS.push({ metodo: metodo, re: re, nombres: nombres, acceso: acceso, fn: fn });
}
function despachar_(metodo, ruta, query, body, token) {
  for (var i = 0; i < RUTAS.length; i++) {
    var r = RUTAS[i];
    if (r.metodo !== metodo) continue;
    var m = ruta.match(r.re); if (!m) continue;
    var params = {}; r.nombres.forEach(function (n, k) { params[n] = decodeURIComponent(m[k + 1]); });
    var ctx = { params: params, query: query, body: body, token: token, br: null };
    if (r.acceso) {
      var u = usuarioDeSesion_(token);
      if (!u) return { success: false, status: 401, sinSesion: true, error: 'Tu sesión terminó. Entra otra vez con tu usuario y contraseña.' };
      if (u.DebeCambiar) return { success: false, status: 403, debeCambiar: true, error: 'Antes de seguir, cambia tu contraseña.' };
      var roles = r.acceso === 'admin' ? ['ADMIN'] : r.acceso === 'gestor' ? ['ADMIN', 'GESTOR'] : null;
      if (roles && roles.indexOf(u.Rol) === -1) return { success: false, status: 403, error: 'Tu usuario no tiene permiso para esto.' };
      ctx.br = u;
    }
    return r.fn(ctx);
  }
  return { success: false, status: 404, error: 'Endpoint no encontrado: ' + metodo + ' ' + ruta };
}

// ---- registro de fallas (pestaña _Registro) ----
function anotar_(etapa, err) {
  try {
    var det = err ? String(err.stack || err.message || err).slice(0, 4000) : '';
    var sh = hojaDe_('_Registro');
    if (sh.getLastRow() > 400) sh.deleteRows(2, 200);
    sh.appendRow([aCelda_('T', ahora_()), etapa, det]);
  } catch (e) { console.error(e); }
}

// ---- contraseñas: SHA-256 repetido con sal (sha$iteraciones$sal$hash) ----
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
function sha_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }
function cifrar_(sal, clave, iter) { var h = sha_(sal + '|' + clave); for (var i = 1; i < iter; i++) h = sha_(sal + '|' + h); return h; }
function hashClave_(clave) { var sal = Utilities.getUuid().replace(/-/g, ''); return 'sha$' + ITERACIONES_CLAVE + '$' + sal + '$' + cifrar_(sal, String(clave), ITERACIONES_CLAVE); }
function igualSeguro_(a, b) { a = String(a); b = String(b); if (a.length !== b.length) return false; var d = 0; for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i); return d === 0; }
function verificarClave_(clave, guardado) {
  var p = String(guardado || '').split('$');
  if (p.length !== 4 || p[0] !== 'sha') { cifrar_('relleno', String(clave), ITERACIONES_CLAVE); return false; }
  return igualSeguro_(cifrar_(p[2], String(clave), Number(p[1]) || ITERACIONES_CLAVE), p[3]);
}

// ---- sesión: token firmado con HMAC (como la cookie del servidor Node) ----
function secreto_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRETO_SESION');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRETO_SESION', s); }
  return s;
}
function b64u_(s) { return Utilities.base64EncodeWebSafe(s, Utilities.Charset.UTF_8).replace(/=+$/, ''); }
function firma_(cuerpo) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(cuerpo, secreto_())).replace(/=+$/, ''); }
function firmarSesion_(u) {
  var cuerpo = b64u_(JSON.stringify({ u: u.Id, v: u.SesionVersion, exp: Date.now() + HORAS_SESION * 3600000 }));
  return cuerpo + '.' + firma_(cuerpo);
}
function leerSesion_(token) {
  if (!token || typeof token !== 'string') return null;
  var p = token.split('.'); if (p.length !== 2) return null;
  if (!igualSeguro_(firma_(p[0]), p[1])) return null;
  try {
    var pad = p[0] + '==='.slice((p[0].length + 3) % 4);
    var o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(pad)).getDataAsString('UTF-8'));
    return o && o.exp && o.exp > Date.now() ? o : null;
  } catch (e) { return null; }
}
function usuarioDeSesion_(token) {
  var t = leerSesion_(token); if (!t) return null;
  var u = T_('BR_Usuario').filter(function (x) { return x.Id === t.u; })[0];
  if (!u || !u.Activo || u.SesionVersion !== t.v) return null;
  return u;
}
function hayAdmin_() { return T_('BR_Usuario').some(function (u) { return u.Rol === 'ADMIN' && u.Activo; }); }
function bitacora_(ctx, accion, detalle, u) {
  try {
    var q = u || (ctx && ctx.br) || {};
    ins_('BR_Bitacora', { UsuarioId: q.Id || null, Usuario: q.Usuario ? String(q.Usuario).slice(0, 50) : null, Accion: accion,
      Detalle: detalle ? String(detalle).slice(0, 500) : null, IP: null });
  } catch (e) { console.error(e); }
}
function publico_(u) { return u ? { id: u.Id, usuario: u.Usuario, nombre: u.Nombre, rol: u.Rol, debeCambiar: !!u.DebeCambiar } : null; }
function usuarioValido_(t) { return /^[a-zA-Z0-9._-]{3,50}$/.test(String(t || '')); }
function claveInstalacion_() { return PropertiesService.getScriptProperties().getProperty('CLAVE_INSTALACION') || ''; }

// ---- quita los campos internos antes de mandar filas a la página ----
function limpio_(o) { return o; }

// =============================================================================
// SESIÓN, INSTALACIÓN Y USUARIOS
// =============================================================================
ruta_('GET', '/api/bienes/diagnostico', null, function () {
  var sh = libro_().getSheetByName('_Registro'), txt = '(todavía no hay registro de fallas)';
  if (sh && sh.getLastRow() > 1) {
    var n = sh.getLastRow(), desde = Math.max(2, n - 59);
    txt = sh.getRange(desde, 1, n - desde + 1, 3).getDisplayValues().map(function (r) { return r[0] + '  [' + r[1] + ']' + (r[2] ? '\n' + r[2] : ''); }).join('\n');
  }
  return { success: true, node: 'Google Apps Script', pid: '', servidor: 'Google Sheets', base: libro_().getName(),
    usuarioSql: Session.getEffectiveUser().getEmail() || 'propietario de la hoja', contrasenaPuesta: true,
    claveInstalacionPuesta: !!claveInstalacion_(), registro: txt };
});

ruta_('GET', '/api/bienes/sesion', null, function (ctx) {
  var instalado = hayAdmin_();
  var u = instalado ? usuarioDeSesion_(ctx.token) : null;
  return { success: true, instalado: instalado, usuario: publico_(u) };
});

// Crear el primer administrador (solo si no hay ningún ADMIN activo)
ruta_('POST', '/api/bienes/instalar', null, function (ctx) {
  var b = ctx.body || {};
  var clave = claveInstalacion_();
  if (!clave) throw err_('Falta la clave de instalación: en la hoja de Google abre el menú "Bienes reposeídos" → "Instalar".', 400);
  if (!igualSeguro_(String(b.clave || ''), clave)) {
    bitacora_(ctx, 'INSTALAR_FALLIDO', 'Clave de instalación incorrecta', { Usuario: b.usuario });
    return falla_('La clave de instalación no es correcta.', 403);
  }
  var usuario = String(b.usuario || '').trim(), nombre = String(b.nombre || '').trim().slice(0, 100);
  if (!usuarioValido_(usuario)) throw err_('El usuario debe tener de 3 a 50 letras, números, punto, guion o guion bajo.', 400);
  if (!nombre) throw err_('Escribe el nombre de la persona.', 400);
  var malo = validarClaveNueva(b.contrasena, usuario);
  if (malo) throw err_(malo, 400);
  if (hayAdmin_()) throw err_('Ya existe un administrador. Entra con tu usuario.', 409);
  var hash = hashClave_(b.contrasena);
  var u = T_('BR_Usuario').filter(function (x) { return String(x.Usuario).toLowerCase() === usuario.toLowerCase(); })[0];
  if (u) upd_('BR_Usuario', u, { Rol: 'ADMIN', Nombre: nombre, Hash: hash, Activo: true, DebeCambiar: false, Intentos: 0, BloqueadoHasta: null, SesionVersion: u.SesionVersion + 1 });
  else u = ins_('BR_Usuario', { Usuario: usuario, Nombre: nombre, Rol: 'ADMIN', Hash: hash, DebeCambiar: false, CreadoPor: 'INSTALACION' });
  bitacora_(ctx, 'INSTALAR', 'Administrador creado: ' + usuario, u);
  return { success: true, usuario: publico_(u), token: firmarSesion_(u) };
});

ruta_('POST', '/api/bienes/entrar', null, function (ctx) {
  var b = ctx.body || {};
  var usuario = String(b.usuario || '').trim(), clave = String(b.contrasena || '');
  var generico = 'Usuario o contraseña incorrectos.';
  if (!usuario || !clave) throw err_('Escribe tu usuario y tu contraseña.', 400);
  conLock_();
  var u = T_('BR_Usuario').filter(function (x) { return String(x.Usuario).toLowerCase() === usuario.slice(0, 50).toLowerCase(); })[0];
  if (!u) {
    verificarClave_(clave, '');   // mismo tiempo de respuesta
    bitacora_(ctx, 'ENTRAR_FALLIDO', 'Usuario inexistente: ' + usuario.slice(0, 50), {});
    return falla_(generico, 401);
  }
  var ahora = ahora_();
  if (u.BloqueadoHasta && u.BloqueadoHasta > ahora) {
    var min = Math.ceil((u.BloqueadoHasta - ahora) / 60000);
    bitacora_(ctx, 'ENTRAR_BLOQUEADO', null, u);
    return falla_('Usuario bloqueado por intentos fallidos. Prueba en ' + min + ' minutos o pide al administrador que lo desbloquee.', 429);
  }
  var ok = verificarClave_(clave, u.Hash);
  if (!ok || !u.Activo) {
    if (!ok) {
      var n = (u.Intentos || 0) + 1;
      if (n >= MAX_INTENTOS) upd_('BR_Usuario', u, { Intentos: 0, BloqueadoHasta: new Date(ahora.getTime() + MINUTOS_BLOQUEO * 60000) });
      else upd_('BR_Usuario', u, { Intentos: n });
    }
    bitacora_(ctx, 'ENTRAR_FALLIDO', ok ? 'Usuario desactivado' : 'Contraseña incorrecta', u);
    return falla_(generico, 401);
  }
  upd_('BR_Usuario', u, { Intentos: 0, BloqueadoHasta: null, UltimoIngreso: ahora });
  bitacora_(ctx, 'ENTRAR', null, u);
  return { success: true, usuario: publico_(u), token: firmarSesion_(u) };
});

ruta_('POST', '/api/bienes/salir', null, function () { return { success: true, token: '' }; });

// Cambiar la propia contraseña (también sirve cuando es obligatorio cambiarla)
ruta_('POST', '/api/bienes/mi-contrasena', null, function (ctx) {
  var b = ctx.body || {};
  var u = usuarioDeSesion_(ctx.token);
  if (!u) return { success: false, status: 401, sinSesion: true, error: 'Tu sesión terminó. Entra otra vez.' };
  if (!verificarClave_(b.actual, u.Hash)) {
    bitacora_(ctx, 'CLAVE_FALLIDO', 'Contraseña actual incorrecta', u);
    return falla_('La contraseña actual no es correcta.', 400);
  }
  var malo = validarClaveNueva(b.nueva, u.Usuario);
  if (malo) throw err_(malo, 400);
  if (String(b.nueva) === String(b.actual)) throw err_('La contraseña nueva debe ser distinta de la actual.', 400);
  upd_('BR_Usuario', u, { Hash: hashClave_(b.nueva), DebeCambiar: false, SesionVersion: u.SesionVersion + 1 });
  bitacora_(ctx, 'CAMBIAR_CLAVE', null, u);
  return { success: true, usuario: publico_(u), token: firmarSesion_(u) };   // las otras sesiones abiertas quedan cerradas
});

// ---------------- USUARIOS (solo ADMIN) ----------------
ruta_('GET', '/api/bienes/usuarios', 'admin', function (ctx) {
  var ahora = ahora_();
  var us = T_('BR_Usuario').slice().sort(function (a, b) { return (b.Activo - a.Activo) || String(a.Nombre).localeCompare(String(b.Nombre)); })
    .map(function (u) { return { Id: u.Id, Usuario: u.Usuario, Nombre: u.Nombre, Rol: u.Rol, Activo: u.Activo, DebeCambiar: u.DebeCambiar,
      FechaCreacion: u.FechaCreacion, CreadoPor: u.CreadoPor, UltimoIngreso: u.UltimoIngreso, Bloqueado: u.BloqueadoHasta && u.BloqueadoHasta > ahora ? 1 : 0 }; });
  return { success: true, usuarios: us, yo: ctx.br.Id };
});

// Crear usuario. La contraseña que pone el admin es temporal: al entrar la primera vez se obliga a cambiarla.
ruta_('POST', '/api/bienes/usuarios', 'admin', function (ctx) {
  var b = ctx.body || {};
  var usuario = String(b.usuario || '').trim(), nombre = String(b.nombre || '').trim().slice(0, 100);
  var rol = String(b.rol || '').toUpperCase();
  if (!usuarioValido_(usuario)) throw err_('El usuario debe tener de 3 a 50 letras, números, punto, guion o guion bajo (sin espacios).', 400);
  if (!nombre) throw err_('Escribe el nombre de la persona.', 400);
  if (ROLES.indexOf(rol) === -1) throw err_('Rol inválido.', 400);
  var malo = validarClaveNueva(b.contrasena, usuario);
  if (malo) throw err_(malo, 400);
  if (T_('BR_Usuario').some(function (x) { return String(x.Usuario).toLowerCase() === usuario.toLowerCase(); })) throw err_('Ya existe el usuario "' + usuario + '".', 409);
  var u = ins_('BR_Usuario', { Usuario: usuario, Nombre: nombre, Rol: rol, Hash: hashClave_(b.contrasena), DebeCambiar: true, CreadoPor: ctx.br.Usuario });
  bitacora_(ctx, 'USUARIO_CREAR', usuario + ' (' + rol + ')');
  return { success: true, id: u.Id };
});

// Editar: nombre, rol, activo, restablecer contraseña o desbloquear
ruta_('PUT', '/api/bienes/usuarios/:id', 'admin', function (ctx) {
  var id = parseInt(ctx.params.id, 10);
  if (!id) throw err_('Usuario inválido.', 400);
  var b = ctx.body || {};
  var u = T_('BR_Usuario').filter(function (x) { return x.Id === id; })[0];
  if (!u) throw err_('No existe ese usuario.', 404);
  var nombre = b.nombre !== undefined ? String(b.nombre).trim().slice(0, 100) : u.Nombre;
  var rol = b.rol !== undefined ? String(b.rol).toUpperCase() : u.Rol;
  var activo = b.activo !== undefined ? !!b.activo : !!u.Activo;
  if (!nombre) throw err_('El nombre no puede quedar vacío.', 400);
  if (ROLES.indexOf(rol) === -1) throw err_('Rol inválido.', 400);
  if (id === ctx.br.Id && (rol !== u.Rol || !activo)) throw err_('No puedes cambiar tu propio rol ni desactivarte. Pídeselo a otro administrador.', 400);
  var hash = null;
  if (b.contrasena) {
    var malo = validarClaveNueva(b.contrasena, u.Usuario);
    if (malo) throw err_(malo, 400);
    hash = hashClave_(b.contrasena);
  }
  // Cambiar rol, desactivar o poner clave nueva cierra las sesiones abiertas de ese usuario
  var corta = !!hash || rol !== u.Rol || (!activo && u.Activo);
  var desb = !!b.desbloquear || !!hash;
  var cambios = [];
  if (nombre !== u.Nombre) cambios.push('nombre');
  if (rol !== u.Rol) cambios.push('rol ' + u.Rol + '→' + rol);
  if (activo !== !!u.Activo) cambios.push(activo ? 'activado' : 'desactivado');
  if (hash) cambios.push('contraseña restablecida');
  if (b.desbloquear) cambios.push('desbloqueado');
  var c = { Nombre: nombre, Rol: rol, Activo: activo, SesionVersion: u.SesionVersion + (corta ? 1 : 0) };
  if (hash) { c.Hash = hash; c.DebeCambiar = true; }
  if (desb) { c.Intentos = 0; c.BloqueadoHasta = null; }
  upd_('BR_Usuario', u, c);
  bitacora_(ctx, 'USUARIO_EDITAR', u.Usuario + ': ' + (cambios.join(', ') || 'sin cambios'));
  return { success: true };
});

ruta_('GET', '/api/bienes/bitacora', 'admin', function () {
  var ev = T_('BR_Bitacora').slice().sort(function (a, b) { return b.Id - a.Id; }).slice(0, 300)
    .map(function (x) { return { Id: x.Id, Fecha: x.Fecha, Usuario: x.Usuario, Accion: x.Accion, Detalle: x.Detalle, IP: x.IP }; });
  return { success: true, eventos: ev };
});

// =============================================================================
// AYUDAS PARA CONSULTAR (lo que en SQL era ORDER BY, TOP, etc.)
// =============================================================================
// ord_(lista, 'Tabla, Orden DESC, Id') — los vacíos van primero, como en SQL Server
function ord_(lista, espec) {
  var campos = espec.split(',').map(function (s) { var p = s.trim().split(/\s+/); return { k: p[0], desc: /desc/i.test(p[1] || '') }; });
  var val = function (v) { return v instanceof Date ? v.getTime() : typeof v === 'boolean' ? (v ? 1 : 0) : v; };
  return lista.slice().sort(function (a, b) {
    for (var i = 0; i < campos.length; i++) {
      var c = campos[i], x = val(a[c.k]), y = val(b[c.k]), r;
      if (x === y) continue;
      if (x === null || x === undefined) r = -1;
      else if (y === null || y === undefined) r = 1;
      else if (typeof x === 'number' && typeof y === 'number') r = x - y;
      else r = String(x).localeCompare(String(y));
      if (r) return c.desc ? -r : r;
    }
    return 0;
  });
}
function copia_(o, campos) { var r = {}; (campos || Object.keys(o)).forEach(function (k) { r[k] = o[k] === undefined ? null : o[k]; }); return r; }
function iso_(v) { if (v === null || v === undefined || v === '') return null; if (v instanceof Date) return fechaISO(v); var d = aFecha_(v, false); return d ? fechaISO(d) : null; }
function porId_(tabla, id) { return T_(tabla).filter(function (x) { return x.Id === id; })[0] || null; }
// "ORDER BY ISNULL(FechaReporte, '19000101') DESC, Id DESC"
function ordenCortes_(l) {
  return l.slice().sort(function (a, b) {
    var x = iso_(a.FechaReporte) || '1900-01-01', y = iso_(b.FechaReporte) || '1900-01-01';
    return x < y ? 1 : x > y ? -1 : b.Id - a.Id;
  });
}
var COLS_CARGA = ['Id', 'FechaCarga', 'Usuario', 'NombreArchivo', 'TipoArchivo', 'FechaReporte', 'TamanoBytes',
  'TotalActivos', 'EnLegal', 'Alquiladas', 'ConMejora', 'Nota', 'Anulada', 'UsuarioAnula', 'FechaAnulada'];

// Última carga vigente de esos tipos; con fecha, la última cuyo corte sea hasta ese día
function ultimaCarga_(tipos, fecha) {
  var l = T_('BR_Carga').filter(function (c) {
    return !c.Anulada && tipos.indexOf(c.TipoArchivo) !== -1 && (!fecha || (c.FechaReporte && iso_(c.FechaReporte) <= fecha));
  });
  var u = ordenCortes_(l)[0];
  return u ? u.Id : null;
}
function cargaInfo_(id) {
  var c = id ? porId_('BR_Carga', id) : null;
  return c ? copia_(c, ['Id', 'NombreArchivo', 'FechaReporte', 'FechaCarga', 'Usuario']) : null;
}
function filasDeCarga_(tabla, id) { return id ? T_(tabla).filter(function (x) { return x.CargaId === id; }) : []; }
function ultimaActivos_() { return ordenCortes_(T_('BR_Carga').filter(function (c) { return !c.Anulada && c.TipoArchivo === 'ACTIVOS'; }))[0] || null; }
function hayExcelDeActivos_() { return T_('BR_Carga').some(function (c) { return !c.Anulada && c.TipoArchivo === 'ACTIVOS'; }); }

var NOMBRE_ESTADO = { DISPONIBLE: 'Disponible', EN_TRAMITE: 'En trámite de venta', VENDIDA: 'Vendida', LEGAL: 'En legal' };
var ESTADOS_FINCA = ['DISPONIBLE', 'EN_TRAMITE', 'VENDIDA', 'LEGAL'];
function mon_(v) { return v === null || v === undefined || v === '' ? '—' : 'B/. ' + Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function distintoNum_(a, b) { return (a === null || a === undefined ? null : Math.round(Number(a) * 100)) !== (b === null || b === undefined ? null : Math.round(Number(b) * 100)); }
function distintoTxt_(a, b) { return String(a || '').trim().toLowerCase() !== String(b || '').trim().toLowerCase(); }
function grupo_(titulo, tono, filas) { return { titulo: titulo, tono: tono, filas: filas }; }

function registrarCambio_(noFinca, campo, antes, despues, usuario) {
  ins_('BR_FincaHistorial', { NoFinca: noFinca, Campo: campo,
    ValorAnterior: antes === null || antes === undefined ? null : String(antes).slice(0, 500),
    ValorNuevo: despues === null || despues === undefined ? null : String(despues).slice(0, 500), Usuario: usuario });
}
function fincaDe_(nf) { return T_('BR_Finca').filter(function (f) { return f.NoFinca === nf; })[0] || null; }

// =============================================================================
// EXCEL DE ACTIVOS → DASHBOARD
// Compara la lista del Excel con el Dashboard y (si no es vista previa) la aplica.
// Si el Excel no trae la columna "Estado" (formato actual de contabilidad), las
// fincas que ya están conservan su estatus, alquilada y ocupada.
// =============================================================================
function TXT_ALQ(v) { return v ? 'Sí' : 'No'; }
function TXT_MEJ(v) { return v === 'CON' ? 'Con mejora' : v === 'SIN' ? 'Sin mejora' : 'Sin definir'; }
function nombreCampo_(c) {
  return ({ Estado: 'Estatus', Alquilada: 'Alquilada', Ocupada: 'Viven personas', Mejora: 'Mejora', Activa: 'Dashboard',
    NoSeVende: 'No se puede vender', PolizaIncendio: 'Póliza de incendio' })[c] || c;
}
function compararActivos_(activos) {
  var actuales = {};
  T_('BR_Finca').forEach(function (f) { actuales[f.NoFinca] = f; });
  var enExcel = {}; activos.forEach(function (a) { enExcel[a.NoFinca] = 1; });
  var nuevas = activos.filter(function (a) { return !actuales[a.NoFinca]; });
  var cambios = [];
  activos.forEach(function (a) {
    var f = actuales[a.NoFinca]; if (!f) return;
    var c = [];
    if (!f.Activa) c.push({ campo: 'Activa', antes: 'Retirada del inventario', despues: 'Activa' });
    if (!a.SinEstado) {
      if (f.Estado !== a.Estado) c.push({ campo: 'Estado', antes: NOMBRE_ESTADO[f.Estado] || f.Estado, despues: NOMBRE_ESTADO[a.Estado] });
      if (!!f.Alquilada !== a.Alquilada) c.push({ campo: 'Alquilada', antes: TXT_ALQ(f.Alquilada), despues: TXT_ALQ(a.Alquilada) });
    }
    if ((f.Mejora || null) !== a.Mejora) c.push({ campo: 'Mejora', antes: TXT_MEJ(f.Mejora), despues: TXT_MEJ(a.Mejora) });
    if (!a.SinEstado && !!f.Ocupada !== a.Ocupada) c.push({ campo: 'Ocupada', antes: TXT_ALQ(f.Ocupada), despues: TXT_ALQ(a.Ocupada) });
    if (distintoNum_(f.ValorLibros, a.ValorLibros)) c.push({ campo: 'ValorLibros', antes: mon_(f.ValorLibros), despues: mon_(a.ValorLibros) });
    if (!!f.NoSeVende !== !!a.NoSeVende) c.push({ campo: 'NoSeVende', antes: TXT_ALQ(f.NoSeVende), despues: TXT_ALQ(a.NoSeVende) });
    if (!!f.PolizaIncendio !== !!a.PolizaIncendio) c.push({ campo: 'PolizaIncendio', antes: TXT_ALQ(f.PolizaIncendio), despues: TXT_ALQ(a.PolizaIncendio) });
    if (c.length) cambios.push({ NoFinca: a.NoFinca, cambios: c });
  });
  var retiradas = Object.keys(actuales).filter(function (k) { return actuales[k].Activa && !enExcel[k]; });
  return { actuales: actuales, nuevas: nuevas, cambios: cambios, retiradas: retiradas };
}
function aplicarActivos_(activos, usuario, fechaTxt) {
  var cmp = compararActivos_(activos);
  var nota = ' (Excel de activos' + (fechaTxt ? ' al ' + fechaTxt : '') + ')';
  var corta = function (v, n) { return v ? String(v).trim().slice(0, n) : null; };
  activos.forEach(function (a) {
    var f = cmp.actuales[a.NoFinca];
    var datos = { Mejora: a.Mejora, ValorLibros: a.ValorLibros, ValorMercado: a.ValorMercado, ValorVentaRapida: a.ValorVentaRapida,
      ValorMejoras: a.ValorMejoras, Deudor: corta(a.Deudor, 250), FechaAdjudicacion: corta(a.FechaAdjudicacion, 200), Superficie: corta(a.Superficie, 80),
      Detalle: a.Detalle, Letrero: corta(a.Letrero, 30), Observaciones: corta(a.Observaciones, 1000), Dias: a.Dias, Fideicomiso: corta(a.Fideicomiso, 150),
      ValorTerreno: a.ValorTerreno, Deterioro: a.Deterioro, Provision: a.Provision, EstadoExcel: String(a.EstadoExcel || '').slice(0, 40) || null,
      OrdenExcel: a.Orden || null, NoSeVende: !!a.NoSeVende, PolizaIncendio: !!a.PolizaIncendio, Activa: true };
    if (!a.SinEstado || !f) { datos.Estado = a.Estado; datos.Alquilada = a.Alquilada; datos.Ocupada = a.Ocupada; }
    var t = corta(a.TipoPropiedad, 100), ub = corta(a.Ubicacion, 500), cu = corta(a.CodigoUbic, 20);
    if (f) {
      if (t) datos.TipoPropiedad = t; if (ub) datos.Ubicacion = ub; if (cu) datos.CodigoUbic = cu;
      datos.UsuarioModifica = usuario; datos.FechaModifica = ahora_();
      upd_('BR_Finca', f, datos);
    } else {
      datos.NoFinca = a.NoFinca; datos.TipoPropiedad = t; datos.Ubicacion = ub; datos.CodigoUbic = cu; datos.UsuarioCrea = usuario;
      ins_('BR_Finca', datos);
      registrarCambio_(a.NoFinca, 'Alta', null, NOMBRE_ESTADO[a.Estado] + nota, usuario);
    }
  });
  cmp.cambios.forEach(function (c) { c.cambios.forEach(function (x) { registrarCambio_(c.NoFinca, x.campo, x.antes, x.despues + nota, usuario); }); });
  cmp.retiradas.forEach(function (k) {
    upd_('BR_Finca', cmp.actuales[k], { Activa: false, UsuarioModifica: usuario, FechaModifica: ahora_() });
    registrarCambio_(k, 'Activa', 'Activa', 'Retirada: no está en el Excel de activos' + (fechaTxt ? ' al ' + fechaTxt : ''), usuario);
  });
  return cmp;
}
function gruposActivos_(cmp, activos) {
  return [
    grupo_('Fincas nuevas (se agregan al Dashboard)', 'nuevo', cmp.nuevas.map(function (a) { return { finca: a.NoFinca,
      texto: a.EstadoExcel + (a.Ubicacion ? ' · ' + a.Ubicacion : '') + (a.ValorLibros !== null ? ' · libros ' + mon_(a.ValorLibros) : '') }; })),
    grupo_('Fincas que cambian', 'cambia', cmp.cambios.map(function (c) { return { finca: c.NoFinca,
      texto: c.cambios.map(function (x) { return x.campo === 'ValorLibros' ? 'Valor en libros: ' + x.antes + ' → ' + x.despues : nombreCampo_(x.campo) + ': ' + x.antes + ' → ' + x.despues; }).join(' · ') }; })),
    grupo_('Fincas que ya no están en el Excel (salen del Dashboard)', 'sale', cmp.retiradas.map(function (k) {
      var f = cmp.actuales[k]; return { finca: k, texto: (NOMBRE_ESTADO[f.Estado] || f.Estado) + (f.Ubicacion ? ' · ' + f.Ubicacion : '') };
    }))
  ];
}

// ---- reporte de ventas y trazabilidad: qué cambia ----
function estadosVenta_(filas) {
  var m = {};
  filas.forEach(function (v) {
    var x = m[v.NoFinca] = m[v.NoFinca] || {};
    if (v.Tabla === 'FINALIZADA') x.fin = v; else if (v.Tabla === 'EN_PROCESO') x.proc = v; else x.noc = v;
  });
  Object.keys(m).forEach(function (k) { var x = m[k]; x.estado = x.fin ? 'VENDIDA' : x.proc ? 'EN_PROCESO' : 'NO_CONCRETADA'; x.fila = x.fin || x.proc || x.noc; });
  return m;
}
function compararReporte_(libro) {
  var idAnt = ultimaCarga_(['REPORTE', 'MIXTO'], libro.fechaReporte);
  var anterior = cargaInfo_(idAnt);
  var antes = estadosVenta_(filasDeCarga_('BR_Venta', idAnt));
  var ahora = estadosVenta_(libro.ventas);
  var desc = function (x) { return (x.fila.Cliente ? x.fila.Cliente + ' · ' : '') + mon_(x.fila.ValorVentaNegociado); };
  var g = { nuevas: [], vendidas: [], salen: [], noConc: [], datos: [] };
  Object.keys(ahora).forEach(function (k) {
    var n = ahora[k], a = antes[k];
    var aReg = a && a.estado !== 'NO_CONCRETADA';
    if (n.estado !== 'NO_CONCRETADA' && !aReg) g.nuevas.push({ finca: k, texto: (n.estado === 'VENDIDA' ? 'vendida · ' : 'en proceso · ') + desc(n) });
    else if (n.estado === 'VENDIDA' && a && a.estado === 'EN_PROCESO') g.vendidas.push({ finca: k, texto: desc(n) });
    if (n.estado === 'NO_CONCRETADA' && (!a || a.estado !== 'NO_CONCRETADA')) g.noConc.push({ finca: k, texto: desc(n) + (a ? ' · antes ' + (a.estado === 'VENDIDA' ? 'vendida' : 'en proceso') : '') });
    if (n.estado !== 'NO_CONCRETADA' && aReg) {
      var c = [], x = n.fila, y = a.fila;
      if (distintoNum_(y.ValorVentaNegociado, x.ValorVentaNegociado)) c.push('Venta negociada: ' + mon_(y.ValorVentaNegociado) + ' → ' + mon_(x.ValorVentaNegociado));
      if (distintoTxt_(y.Cliente, x.Cliente)) c.push('Cliente: ' + (y.Cliente || '—') + ' → ' + (x.Cliente || '—'));
      if (distintoNum_(y.MontoFinanciado, x.MontoFinanciado)) c.push('Financiado: ' + mon_(y.MontoFinanciado) + ' → ' + mon_(x.MontoFinanciado));
      if (distintoTxt_(y.Ejecutivo, x.Ejecutivo)) c.push('Ejecutivo: ' + (y.Ejecutivo || '—') + ' → ' + (x.Ejecutivo || '—'));
      if (c.length) g.datos.push({ finca: k, texto: c.join(' · ') });
    }
  });
  Object.keys(antes).forEach(function (k) {
    var a = antes[k];
    if (a.estado !== 'NO_CONCRETADA' && !ahora[k]) g.salen.push({ finca: k, texto: (a.estado === 'VENDIDA' ? 'estaba vendida · ' : 'estaba en proceso · ') + desc(a) });
  });
  var inv = {}; T_('BR_Finca').forEach(function (f) { inv[f.NoFinca] = f; });
  var manual = {}; T_('BR_VentaManual').forEach(function (v) { if (!v.Anulada && v.Tabla !== 'NO_CONCRETADA') manual[v.NoFinca] = 1; });
  var listaManda = hayExcelDeActivos_();
  var efectos = [], regresan = [];
  Object.keys(ahora).forEach(function (k) {
    var n = ahora[k], f = inv[k];
    if (!f) { if (n.estado !== 'NO_CONCRETADA') efectos.push({ finca: k, texto: listaManda ? 'no está en el Excel de activos: no se agrega' : 'se agrega al Dashboard' }); return; }
    if (!f.Activa) return;
    if (n.estado === 'VENDIDA' && (f.Estado === 'DISPONIBLE' || f.Estado === 'EN_TRAMITE')) efectos.push({ finca: k, texto: NOMBRE_ESTADO[f.Estado] + ' → Vendida' });
    else if (n.estado === 'EN_PROCESO' && f.Estado === 'DISPONIBLE') efectos.push({ finca: k, texto: 'Disponible → En trámite de venta' });
  });
  g.salen.concat(g.noConc).forEach(function (r) {
    var f = inv[r.finca], n = ahora[r.finca];
    if (f && f.Activa && f.Estado === 'EN_TRAMITE' && !manual[r.finca] && (!n || n.estado === 'NO_CONCRETADA') && regresan.indexOf(r.finca) === -1) {
      regresan.push(r.finca); efectos.push({ finca: r.finca, texto: 'En trámite de venta → Disponible (ya no está en proceso)' });
    }
  });
  var grupos = anterior ? [
    grupo_('Ventas nuevas', 'nuevo', g.nuevas), grupo_('Pasan a vendida', 'nuevo', g.vendidas),
    grupo_('Cambian datos de la venta', 'cambia', g.datos), grupo_('Nuevas no concretadas', 'sale', g.noConc),
    grupo_('Ya no aparecen en el reporte', 'sale', g.salen)
  ] : [grupo_('Primer reporte de ventas: se cargan todas', 'info', [])];
  grupos.push(grupo_('Cambios en el Dashboard', 'info', efectos));
  return { anterior: anterior, grupos: grupos, regresan: regresan };
}
function compararTraza_(libro) {
  var idAnt = ultimaCarga_(['TRAZABILIDAD', 'MIXTO'], libro.fechaReporte);
  var anterior = cargaInfo_(idAnt);
  if (!anterior) return { anterior: null, grupos: [grupo_('Primera trazabilidad: se carga completa', 'info', [])] };
  var antes = {}; filasDeCarga_('BR_Trazabilidad', idAnt).forEach(function (t) { antes[t.NoFinca] = t; });
  var ahora = {}; libro.trazabilidad.forEach(function (t) { ahora[t.NoFinca] = t; });
  var nuevas = [], salen = [], cambian = [];
  Object.keys(ahora).forEach(function (k) {
    var n = ahora[k], a = antes[k];
    if (!a) { nuevas.push({ finca: k, texto: (n.Grupo === 'VENDIDA' ? 'vendida' : 'en proceso') + ' · balance ' + mon_(n.Balance) }); return; }
    var c = [];
    if (a.Grupo !== n.Grupo) c.push('Grupo: ' + (a.Grupo === 'VENDIDA' ? 'vendida' : 'en proceso') + ' → ' + (n.Grupo === 'VENDIDA' ? 'vendida' : 'en proceso'));
    if (distintoNum_(a.TotalGastos, n.TotalGastos)) c.push('Gastos: ' + mon_(a.TotalGastos) + ' → ' + mon_(n.TotalGastos));
    if (distintoNum_(a.ValorVentaNegociado, n.ValorVentaNegociado)) c.push('Venta: ' + mon_(a.ValorVentaNegociado) + ' → ' + mon_(n.ValorVentaNegociado));
    if (distintoNum_(a.Balance, n.Balance)) c.push('Balance: ' + mon_(a.Balance) + ' → ' + mon_(n.Balance));
    if (c.length) cambian.push({ finca: k, texto: c.join(' · ') });
  });
  Object.keys(antes).forEach(function (k) { if (!ahora[k]) salen.push({ finca: k, texto: 'balance ' + mon_(antes[k].Balance) }); });
  return { anterior: anterior, grupos: [grupo_('Fincas nuevas en la trazabilidad', 'nuevo', nuevas), grupo_('Cambian', 'cambia', cambian),
    grupo_('Ya no aparecen', 'sale', salen)] };
}
function compararLibro_(libro) {
  var r = { anterior: null, grupos: [], regresan: [] };
  if (libro.tipo === 'REPORTE' || libro.tipo === 'MIXTO') { var x = compararReporte_(libro); r.anterior = x.anterior; r.grupos = r.grupos.concat(x.grupos); r.regresan = x.regresan; }
  if (libro.tipo === 'TRAZABILIDAD' || libro.tipo === 'MIXTO') {
    var y = compararTraza_(libro); r.anterior = r.anterior || y.anterior;
    r.grupos = r.grupos.concat(y.grupos.map(function (g) { return Object.assign({}, g, { titulo: libro.tipo === 'MIXTO' ? 'Trazabilidad: ' + g.titulo : g.titulo }); }));
  }
  return r;
}
// Fincas que estaban en trámite y el reporte nuevo ya no las tiene en proceso: vuelven a disponible
function regresarADisponible_(fincas, usuario, fechaTxt) {
  fincas.forEach(function (k) {
    var f = fincaDe_(k);
    if (f && f.Estado === 'EN_TRAMITE' && f.Activa) {
      upd_('BR_Finca', f, { Estado: 'DISPONIBLE', UsuarioModifica: usuario, FechaModifica: ahora_() });
      registrarCambio_(k, 'Estado', NOMBRE_ESTADO.EN_TRAMITE, NOMBRE_ESTADO.DISPONIBLE + ' (ya no está en proceso en el reporte de ventas' + (fechaTxt ? ' del ' + fechaTxt : '') + ')', usuario);
    }
  });
}
// Aplica todo lo que trae un Excel de activos: fincas, marcas, ventas por año, notas del pie y desalojo
function guardarDatosActivos_(libro, cargaId, usuario) {
  aplicarActivos_(libro.activos, usuario, libro.fechaReporte ? fechaDMY(aFecha_(libro.fechaReporte)) : '');
  if (libro.ventasAnteriores && libro.ventasAnteriores.length) {
    del_('BR_VentaAnterior', function () { return true; });
    libro.ventasAnteriores.forEach(function (v) { ins_('BR_VentaAnterior', Object.assign({}, v, { Id: null, CargaId: cargaId })); });
  }
  if (libro.pie && Object.keys(libro.pie).length) {
    ponerCfg_('activos_pie', JSON.stringify({ p: libro.pie.provision, s: libro.pie.saldo, s2: libro.pie.saldo2, a: libro.pie.saldoAl, d: libro.pie.diferencia }));
  }
  if (libro.desalojo && libro.desalojo.length) {
    del_('BR_Desalojo', function () { return true; });
    libro.desalojo.forEach(function (d) { ins_('BR_Desalojo', Object.assign({}, d, { Id: null, CargaId: cargaId })); });
  }
}

// ---- cuadro de lotes con mejoras (casas): seguimiento por cortes ----
function cortesMejoras_(hasta) {
  var h = /^\d{4}-\d{2}-\d{2}$/.test(String(hasta || '')) ? hasta : null;
  return T_('BR_Carga').filter(function (c) { return !c.Anulada && c.TipoArchivo === 'MEJORAS' && (!h || !c.FechaReporte || iso_(c.FechaReporte) <= h); });
}
function mejorasActuales_(hasta) {
  var c = ordenCortes_(cortesMejoras_(hasta))[0];
  return c ? ord_(T_('BR_Mejora').filter(function (m) { return m.CargaId === c.Id; }), 'Orden') : [];
}
var CAMPOS_SEGUIMIENTO = [['Accion', 'Acción a tomar'], ['Condicion', 'Condición'], ['Residen', '¿Residen?'], ['QuienReside', 'Quién reside'],
  ['Visitada', '¿Visitada?'], ['FechaVisita', 'Fecha de visita'], ['ValorNegociado', 'Valor negociado'], ['Nota', 'Nota'], ['Comentarios', 'Comentarios']];
function historialMejoras_(hasta) {
  var cs = cortesMejoras_(hasta).slice().sort(function (a, b) {
    var x = iso_(a.FechaReporte) || '1900-01-01', y = iso_(b.FechaReporte) || '1900-01-01'; return x < y ? -1 : x > y ? 1 : a.Id - b.Id; });
  var cortes = [], porCorte = {};
  cs.forEach(function (c) {
    var filas = ord_(T_('BR_Mejora').filter(function (m) { return m.CargaId === c.Id; }), 'Orden');
    if (!filas.length) return;
    porCorte[c.Id] = filas; cortes.push({ CargaId: c.Id, Fecha: c.FechaReporte, Archivo: c.NombreArchivo });
  });
  var val = function (m, k) { var v = m[k]; if (v == null || v === '') return ''; if (k === 'FechaVisita') return fechaISO(new Date(v)); if (k === 'ValorNegociado') return (Math.round(Number(v) * 100) / 100).toFixed(2); return String(v).trim(); };
  var hist = {}, previo = {};
  cortes.forEach(function (c, i) {
    var aqui = {};
    porCorte[c.CargaId].forEach(function (m) {
      aqui[m.NoFinca] = 1;
      var a = previo[m.NoFinca], h = hist[m.NoFinca] = hist[m.NoFinca] || [];
      if (!a) h.push({ fecha: c.Fecha, archivo: c.Archivo, tipo: i === 0 ? 'inicio' : 'entra', cambios: [], accion: m.Accion, condicion: m.Condicion });
      else {
        var d = CAMPOS_SEGUIMIENTO.filter(function (x) { return val(a, x[0]) !== val(m, x[0]); }).map(function (x) { return { campo: x[1], antes: val(a, x[0]), ahora: val(m, x[0]) }; });
        if (d.length) h.push({ fecha: c.Fecha, archivo: c.Archivo, tipo: 'cambio', cambios: d, accion: m.Accion, condicion: m.Condicion });
      }
      previo[m.NoFinca] = m;
    });
    if (i > 0) Object.keys(previo).forEach(function (nf) {
      if (!aqui[nf] && previo[nf] && !previo[nf]._fuera) { (hist[nf] = hist[nf] || []).push({ fecha: c.Fecha, archivo: c.Archivo, tipo: 'sale', cambios: [] }); previo[nf] = Object.assign({}, previo[nf], { _fuera: 1 }); }
    });
    Object.keys(aqui).forEach(function (nf) { if (previo[nf] && previo[nf]._fuera) { previo[nf] = Object.assign({}, previo[nf]); delete previo[nf]._fuera; } });
  });
  return { cortes: cortes, hist: hist };
}
function guardarMejoras_(filas, cargaId) {
  del_('BR_Mejora', function (m) { return m.CargaId === cargaId; });
  filas.forEach(function (m) { ins_('BR_Mejora', Object.assign({}, m, { Id: null, CargaId: cargaId })); });
}
function compararMejoras_(filas) {
  var ant = {}, inv = {};
  mejorasActuales_().forEach(function (m) { ant[m.NoFinca] = m; });
  T_('BR_Finca').forEach(function (f) { inv[f.NoFinca] = f; });
  var id = function (m) { return m.NoFinca + (m.CodigoUbic ? '-' + m.CodigoUbic : ''); };
  var n2 = function (v) { return v == null ? null : Math.round(Number(v) * 100) / 100; };
  var hayAnterior = Object.keys(ant).length > 0, nuevas = [], cambios = [], diferencias = [], noEstan = [];
  var campos = [['Accion', 'Acción a tomar'], ['Condicion', 'Condición'], ['Residen', '¿Residen?'], ['QuienReside', 'Quién reside'],
    ['Visitada', '¿Visitada?'], ['ValorNegociado', 'Valor negociado'], ['Nota', 'Nota']];
  filas.forEach(function (m) {
    var a = ant[m.NoFinca];
    if (hayAnterior && !a) nuevas.push(id(m));
    if (a) {
      var d = campos.filter(function (x) { var k = x[0]; return String(k === 'ValorNegociado' ? n2(a[k]) : (a[k] || '')) !== String(k === 'ValorNegociado' ? n2(m[k]) : (m[k] || '')); })
        .map(function (x) { var k = x[0]; return x[1] + ': ' + (a[k] == null || a[k] === '' ? '—' : a[k]) + ' → ' + (m[k] == null || m[k] === '' ? '—' : m[k]); });
      if ((a.Comentarios || '').trim() !== (m.Comentarios || '').trim() && m.Comentarios)
        d.push('Comentario nuevo: ' + String(m.Comentarios).slice(0, 160) + (String(m.Comentarios).length > 160 ? '…' : ''));
      if (d.length) cambios.push({ finca: id(m), cambios: d });
    }
    var f = inv[m.NoFinca];
    if (!f) { if (Object.keys(inv).length) noEstan.push(id(m)); return; }
    var dif = [['ValorMejoras', 'mejoras'], ['ValorMercado', 'mercado'], ['ValorVentaRapida', 'venta rápida'], ['ValorLibros', 'libros']]
      .filter(function (x) { var k = x[0]; return f[k] != null && (m[k] == null || Math.abs(n2(m[k]) - n2(f[k])) > 0.05); })
      .map(function (x) { var k = x[0]; return x[1] + ' ' + (m[k] == null ? 'vacío' : n2(m[k]).toFixed(2)) + ' (Excel de activos: ' + n2(f[k]).toFixed(2) + ')'; });
    if (dif.length) diferencias.push({ finca: id(m), detalle: dif });
  });
  var salen = hayAnterior ? Object.keys(ant).filter(function (k) { return !filas.some(function (m) { return m.NoFinca === k; }); }).map(function (k) { return id(ant[k]); }) : [];
  return { hayAnterior: hayAnterior, nuevas: nuevas, salen: salen, cambios: cambios, diferencias: diferencias, noEstan: noEstan };
}

// =============================================================================
// SUBIR UN EXCEL
// La página lee el Excel en el navegador (con las mismas funciones del servidor
// Node) y manda lo leído: { libro, alquileres, hash, tamano, nombreArchivo, ... }.
// El archivo original (archivoB64) solo llega al guardar, para dejarlo en Drive.
// =============================================================================
function insertarCarga_(ctx, b, tipo, fechaReporte, extra) {
  var arch = guardarArchivo_(b.nombreArchivo, b.archivoB64);
  return ins_('BR_Carga', Object.assign({ Usuario: ctx.br.Nombre, UsuarioId: ctx.br.Id, NombreArchivo: String(b.nombreArchivo).slice(0, 260),
    TipoArchivo: tipo, FechaReporte: fechaReporte, HashArchivo: b.hash, TamanoBytes: arch.bytes || b.tamano, ArchivoId: arch.id,
    Nota: b.nota ? String(b.nota).slice(0, 500) : null }, extra || {}));
}
function duplicado_(hash) {
  var d = ord_(T_('BR_Carga').filter(function (c) { return c.HashArchivo === hash && !c.Anulada; }), 'Id DESC')[0];
  return d ? copia_(d, ['Id', 'FechaCarga', 'Usuario']) : null;
}

ruta_('POST', '/api/bienes/cargar', 'gestor', function (ctx) {
  var b = ctx.body || {};
  var nombre = String(b.nombreArchivo || '').trim();
  if (!nombre || !b.libro) throw err_('Falta el archivo.', 400);
  if (!/\.xlsx$/i.test(nombre)) throw err_('El archivo debe ser .xlsx (Excel). Si es .xls, ábrelo en Excel y guárdalo como .xlsx.', 400);
  if (!b.soloVista && !b.archivoB64) throw err_('No llegó el archivo.', 400);
  var libro = b.libro;
  if (!libro.tipo) {
    // ¿Es el Excel de control de alquileres? Se guarda en Alquileres, igual que desde esa pestaña
    var alq = importarAlquileres_(ctx, b, !!b.soloVista);
    if (alq) return Object.assign({ success: true, vista: !!b.soloVista, tipo: 'ALQUILERES' }, alq);
    throw err_('No encontré ninguna tabla conocida. El archivo debe ser el reporte de ventas, el Excel de activos, el análisis de trazabilidad, ' +
      'el cuadro de lotes con mejoras o el control de alquileres.', 400);
  }
  // Fecha del corte: la tabla "ventas en proceso" manda; si no la trae, el nombre del archivo; si tampoco, la fecha más alta.
  var deNombre = interpretarFecha(nombre.replace(/[_\-.]+/g, ' '));
  var fuenteFecha = '';
  if (libro.tipo === 'ACTIVOS' || libro.tipo === 'MEJORAS') { libro.fechaReporte = libro.fechaActivos || deNombre; fuenteFecha = libro.fechaActivos ? 'el título del Excel' : 'el nombre del archivo'; }
  else if (libro.tipo === 'TRAZABILIDAD') { libro.fechaReporte = libro.fechaTraza || deNombre; fuenteFecha = libro.fechaTraza ? 'la hoja de trazabilidad' : 'el nombre del archivo'; }
  else if (libro.fechaProceso) { libro.fechaReporte = libro.fechaProceso; fuenteFecha = 'la tabla de ventas en proceso'; }
  else if (deNombre) { libro.fechaReporte = deNombre; fuenteFecha = 'el nombre del archivo'; }
  else fuenteFecha = libro.fechaReporte ? 'la fecha más reciente dentro del archivo' : '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(b.fechaCorte || '')) && aFecha_(b.fechaCorte)) { libro.fechaReporte = b.fechaCorte; fuenteFecha = 'corregida a mano'; }
  ['ventas', 'visitas', 'resumen', 'trazabilidad', 'avisos'].forEach(function (k) { libro[k] = libro[k] || []; });

  // ---- Excel de activos: actualiza el Dashboard ----
  if (libro.tipo === 'ACTIVOS') {
    var dup = duplicado_(b.hash);
    var cmp = compararActivos_(libro.activos);
    var porEstado = {}; libro.activos.forEach(function (a) { porEstado[a.EstadoExcel] = (porEstado[a.EstadoExcel] || 0) + 1; });
    var anteriorAct = ultimaActivos_();
    var resultado = { tipo: 'ACTIVOS', fechaReporte: libro.fechaReporte, fuenteFecha: fuenteFecha, avisos: libro.avisos,
      cambios: { anterior: anteriorAct ? copia_(anteriorAct, ['Id', 'NombreArchivo', 'FechaReporte', 'FechaCarga', 'Usuario']) : null, grupos: gruposActivos_(cmp, libro.activos) },
      activos: { total: libro.activos.length, porEstado: porEstado, desalojo: (libro.desalojo || []).length,
        anos: Array.from(new Set((libro.ventasAnteriores || []).map(function (v) { return v.Anio; }))).sort().reverse(), ventasAnteriores: (libro.ventasAnteriores || []).length,
        noSeVende: libro.activos.filter(function (a) { return a.NoSeVende; }).length, poliza: libro.activos.filter(function (a) { return a.PolizaIncendio; }).length,
        nuevas: cmp.nuevas.length, cambian: cmp.cambios.length, retiradas: cmp.retiradas,
        cuadro: cuadroInventario_(libro.activos.map(function (a) {
          var f = cmp.actuales[a.NoFinca];   // sin columna Estado, cuenta el estatus que ya tiene la finca
          return Object.assign({ Activa: true }, a, a.SinEstado && f ? { Estado: f.Estado, Alquilada: f.Alquilada, Ocupada: f.Ocupada } : {}); })) } };
    if (b.soloVista) return Object.assign({ success: true, vista: true, duplicado: dup }, resultado);
    if (dup && !b.forzar) return { success: false, status: 409, duplicado: dup, error: 'Este mismo archivo ya se subió (carga #' + dup.Id + ').' };
    var c = insertarCarga_(ctx, b, 'ACTIVOS', libro.fechaReporte, { Lectura: { activos: libro.activos, ventasAnteriores: libro.ventasAnteriores || [],
      desalojo: libro.desalojo || [], pie: libro.pie || null, fechaReporte: libro.fechaReporte } });
    guardarDatosActivos_(libro, c.Id, ctx.br.Nombre);
    bitacora_(ctx, 'CARGAR', 'Carga #' + c.Id + ' ACTIVOS: ' + libro.activos.length + ' activos, ' + cmp.nuevas.length + ' nuevos, ' +
      cmp.cambios.length + ' con cambios, ' + cmp.retiradas.length + ' retirados · ' + nombre);
    return Object.assign({ success: true, cargaId: c.Id }, resultado);
  }

  // ---- Cuadro de lotes con mejoras (casas) ----
  if (libro.tipo === 'MEJORAS') {
    var dupM = duplicado_(b.hash);
    var cmpM = compararMejoras_(libro.mejoras);
    var sum = function (k) { return Math.round(libro.mejoras.reduce(function (a, m) { return a + (Number(m[k]) || 0); }, 0) * 100) / 100; };
    var porAccion = {}; libro.mejoras.forEach(function (m) { var k = m.Accion || 'Sin acción'; porAccion[k] = (porAccion[k] || 0) + 1; });
    var esSi = function (v) { return /^s[ií]/i.test(String(v || '').trim()); };
    var cortesExistentes = T_('BR_Carga').filter(function (x) { return !x.Anulada && x.TipoArchivo === 'MEJORAS'; }).map(function (x) { return copia_(x, ['Id', 'NombreArchivo', 'FechaReporte']); });
    var resM = { tipo: 'MEJORAS', fechaReporte: libro.fechaReporte, fuenteFecha: fuenteFecha, fechaVisitas: libro.fechaVisitas, titulo: libro.tituloMejoras, cortesExistentes: cortesExistentes,
      avisos: libro.avisos, comparacion: cmpM,
      mejoras: { total: libro.mejoras.length, residen: libro.mejoras.filter(function (m) { return esSi(m.Residen); }).length,
        visitadas: libro.mejoras.filter(function (m) { return esSi(m.Visitada); }).length, porAccion: porAccion,
        totales: { mejoras: sum('ValorMejoras'), mercado: sum('ValorMercado'), rapida: sum('ValorVentaRapida'), negociado: sum('ValorNegociado'), libros: sum('ValorLibros') } } };
    if (b.soloVista) return Object.assign({ success: true, vista: true, duplicado: dupM }, resM);
    if (dupM && !b.forzar) return { success: false, status: 409, duplicado: dupM, error: 'Este mismo archivo ya se subió (carga #' + dupM.Id + ').' };
    var cM = insertarCarga_(ctx, b, 'MEJORAS', libro.fechaReporte);
    guardarMejoras_(libro.mejoras, cM.Id);
    bitacora_(ctx, 'CARGAR', 'Carga #' + cM.Id + ' LOTES CON MEJORAS: ' + libro.mejoras.length + ' fincas · ' + nombre);
    return Object.assign({ success: true, cargaId: cM.Id }, resM);
  }

  var vc = validarConteos(b.conteos);
  if (vc.error) throw err_(vc.error, 400);
  var conteos = vc.conteos;
  var esReporte = libro.tipo !== 'TRAZABILIDAD';
  var cuadro = esReporte ? calcularCuadro(libro.ventas, conteos) : null;
  var resumenLectura = {
    tipo: libro.tipo, fechaReporte: libro.fechaReporte, fuenteFecha: fuenteFecha,
    filas: {
      ventasEnProceso: libro.ventas.filter(function (v) { return v.Tabla === 'EN_PROCESO'; }).length,
      ventasFinalizadas: libro.ventas.filter(function (v) { return v.Tabla === 'FINALIZADA'; }).length,
      ventasNoConcretadas: libro.ventas.filter(function (v) { return v.Tabla === 'NO_CONCRETADA'; }).length,
      visitas: libro.visitas.length, resumenEjecutivo: libro.resumen.length, trazabilidad: libro.trazabilidad.length
    },
    cuadro: cuadro, avisos: libro.avisos
  };
  var duplicado = duplicado_(b.hash);
  var cambios = compararLibro_(libro);
  resumenLectura.cambios = { anterior: cambios.anterior, grupos: cambios.grupos };
  if (b.soloVista) return Object.assign({ success: true, vista: true, duplicado: duplicado }, resumenLectura);
  if (duplicado && !b.forzar) return { success: false, status: 409, duplicado: duplicado, error: 'Este mismo archivo ya se subió (carga #' + duplicado.Id + '). No se guardó otra vez.' };
  if (cuadro && cuadro.disponibles !== null && cuadro.disponibles < 0) throw err_('Revisa los conteos: ' + cuadro.avisos[0], 400);

  var carga = insertarCarga_(ctx, b, libro.tipo, libro.fechaReporte, {
    TotalActivos: esReporte ? conteos.totalActivos : null, EnLegal: esReporte ? conteos.enLegal : null,
    Alquiladas: esReporte ? conteos.alquiladas : null, ConMejora: esReporte ? conteos.conMejora : null });
  var cargaId = carga.Id;
  libro.ventas.forEach(function (f) { ins_('BR_Venta', Object.assign({}, f, { Id: null, CargaId: cargaId })); });
  libro.visitas.forEach(function (f) { ins_('BR_Visita', Object.assign({}, f, { Id: null, CargaId: cargaId })); });
  libro.resumen.forEach(function (f) { ins_('BR_ResumenEjecutivo', Object.assign({}, f, { Id: null, CargaId: cargaId })); });
  libro.trazabilidad.forEach(function (f) { ins_('BR_Trazabilidad', Object.assign({}, f, { Id: null, CargaId: cargaId })); });
  bitacora_(ctx, 'CARGAR', 'Carga #' + cargaId + ' ' + libro.tipo + ': ' + nombre);
  try {
    sincronizarInventario_(ctx.br.Nombre);
    regresarADisponible_(cambios.regresan, ctx.br.Nombre, libro.fechaReporte ? fechaDMY(aFecha_(libro.fechaReporte)) : '');
  } catch (e) { anotar_('no se pudo poner al día el Dashboard', e); }
  return Object.assign({ success: true, cargaId: cargaId }, resumenLectura);
});

// =============================================================================
// DATOS DE UN CORTE (los usan la pantalla y la descarga)
// =============================================================================
function obtenerCorte_(q, soloReporte) {
  var fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(q.fecha || '')) ? q.fecha : null;
  var idR = parseInt(q.reporte, 10) || ultimaCarga_(['REPORTE', 'MIXTO'], fecha);
  var idT = parseInt(q.traza, 10) || ultimaCarga_(['TRAZABILIDAD', 'MIXTO'], fecha);
  var cargaDe = function (id) { var c = id ? porId_('BR_Carga', id) : null; return c ? copia_(c, COLS_CARGA) : null; };
  var cargaR = cargaDe(idR), cargaT = cargaDe(idT);
  var ventas = idR ? ord_(filasDeCarga_('BR_Venta', idR), 'Tabla, Orden, Id').map(function (v) { return copia_(v); }) : [];
  // Ventas anotadas a mano: entran si su fecha es hasta el corte que se mira (en el más reciente, todas)
  var hasta = q.fecha || (q.reporte && cargaR && cargaR.FechaReporte ? fechaISO(cargaR.FechaReporte) : null);
  var manuales = ord_(T_('BR_VentaManual').filter(function (m) { return !m.Anulada && (!hasta || iso_(m.Fecha) <= hasta); }), 'Fecha, Id');
  var ya = {}; ventas.forEach(function (v) { ya[v.Tabla + '|' + v.NoFinca] = 1; });
  manuales.forEach(function (m) {
    if (ya[m.Tabla + '|' + m.NoFinca]) return;
    ventas.push(Object.assign(copia_(m), { Manual: true, IdManual: m.Id, FechaTabla: m.Fecha }));
  });
  var visitas = idR ? ord_(filasDeCarga_('BR_Visita', idR), 'Orden, Id') : [];
  var resumen = idR ? ord_(filasDeCarga_('BR_ResumenEjecutivo', idR), 'Orden, Id') : [];
  var traza = idT ? ord_(filasDeCarga_('BR_Trazabilidad', idT), 'Grupo DESC, Orden, Id') : [];
  var cuadro = cargaR ? calcularCuadro(ventas, { totalActivos: cargaR.TotalActivos, enLegal: cargaR.EnLegal, alquiladas: cargaR.Alquiladas, conMejora: cargaR.ConMejora }) : null;

  // Corte más reciente: si hay Excel de activos, el cuadro completo sale del Dashboard
  if (!soloReporte && !q.reporte && !q.traza && !q.fecha) {
    var inv = T_('BR_Finca');
    if (!inv.length && cargaR) {
      try { sincronizarInventario_('Automático'); inv = T_('BR_Finca'); } catch (e) { anotar_('no se pudo llenar el Dashboard solo', e); }
    }
    var act = ultimaActivos_();
    act = act ? { FechaReporte: act.FechaReporte, NombreArchivo: act.NombreArchivo } : null;
    if (inv.some(function (f) { return f.Activa; }) && (act || cuadro)) {
      var c = cuadroInventario_(inv);
      var vivas = inv.filter(function (f) { return f.Activa; });
      var n = function (e) { return vivas.filter(function (f) { return f.Estado === e; }).length; };
      var base = cuadro || { avisos: [], fincasEnProceso: [] };
      var avisos = (base.avisos || []).filter(function (a) { return /vendidas que no están/.test(a) && !act; });
      if (c.mejoraSinDefinir) avisos.push(c.mejoraSinDefinir + ' fincas disponibles no tienen marcado si son con o sin mejora (se marca en el Dashboard).');
      cuadro = Object.assign({}, base, {
        totalActivos: c.totalActivos, enTramiteOVendidos: c.enTramiteOVendidos, enLegal: c.enLegal, disponibles: c.disponibles,
        alquiladas: c.alquiladas, conMejora: c.conMejora, sinMejora: c.sinMejora, alquiladasConMejora: c.alquiladasConMejora || [],
        avisos: avisos, fuente: 'DASHBOARD', excelActivos: act
      });
      if (act) Object.assign(cuadro, { ventasRegistradas: c.enTramiteOVendidos, vendidas: n('VENDIDA'), enProceso: n('EN_TRAMITE'),
        fincasEnProceso: vivas.filter(function (f) { return f.Estado === 'EN_TRAMITE'; }).map(function (f) { return f.NoFinca; }) });
    }
  }
  return { cargaReporte: cargaR, cargaTrazabilidad: cargaT, cuadro: cuadro, ventas: ventas, visitas: visitas, resumenEjecutivo: resumen,
    trazabilidad: traza, totalesTrazabilidad: totalesTraza(traza), cruces: cruzar(ventas, traza) };
}

ruta_('GET', '/api/bienes/resumen', 'cualquiera', function (ctx) {
  var datos = obtenerCorte_(ctx.query);
  var ca = ultimaActivos_();
  datos.cargaActivos = ca ? copia_(ca, ['Id', 'FechaReporte', 'NombreArchivo', 'FechaCarga', 'Usuario']) : null;
  datos.codigos = T_('BR_Finca').filter(function (f) { return f.CodigoUbic; }).map(function (f) { return { NoFinca: f.NoFinca, CodigoUbic: f.CodigoUbic }; });
  var va = T_('BR_VentaAnterior'), maxA = va.reduce(function (a, v) { return Math.max(a, v.Anio || 0); }, 0);
  datos.ventasAnio = datos.cargaActivos && !ctx.query.reporte && !ctx.query.fecha ? ord_(va.filter(function (v) { return v.Anio === maxA; }), 'Orden, Id') : [];
  return Object.assign({ success: true }, datos);
});

// Descarga en Excel: el servidor junta los datos y la página arma el libro (mismo formato que antes)
ruta_('GET', '/api/bienes/exportar', 'cualquiera', function (ctx) {
  var datos = obtenerCorte_(ctx.query);
  if (!datos.cargaReporte && !datos.cargaTrazabilidad && !hayExcelDeActivos_()) throw err_('Todavía no hay datos para descargar.', 404);
  var hist = ord_(T_('BR_Carga').filter(function (c) { return c.TipoArchivo !== 'ALQUILERES'; }), 'Id DESC').map(function (c) { return copia_(c, COLS_CARGA); });
  var extra = null;
  if (hayExcelDeActivos_()) {
    var fq = /^\d{4}-\d{2}-\d{2}$/.test(String(ctx.query.fecha || '')) ? ctx.query.fecha : null;
    var inv = leerInventario_(fq);
    var ca = ultimaActivos_();
    var pieA = null; try { pieA = cfg_('activos_pie') ? JSON.parse(cfg_('activos_pie')) : null; } catch (e) { pieA = null; }
    extra = { activos: inv.filas.filter(function (f) { return f.Activa; }), fechaActivos: ca && ca.FechaReporte, pie: pieA,
      ventasAnteriores: ord_(T_('BR_VentaAnterior'), 'Anio DESC, Orden'), desalojo: ord_(T_('BR_Desalojo'), 'Orden'),
      mejoras: mejorasActuales_(fq), seguimientoMejoras: historialMejoras_(fq) };
  }
  var corte = (datos.cargaReporte || datos.cargaTrazabilidad || { FechaReporte: extra && extra.fechaActivos }).FechaReporte || ahora_();
  var nombre = 'Bienes_Reposeidos_corte_' + (corte ? fechaISO(new Date(corte)) : 'sin_fecha') + '.xlsx';
  bitacora_(ctx, 'DESCARGAR', nombre);
  return { success: true, __excel: 'corte', nombre: nombre, datos: datos, cargas: hist, extra: extra };
});

// =============================================================================
// FINCAS CONOCIDAS (para autocompletar y llenar documentos)
// =============================================================================
function fincasConocidas_(fecha) {
  var c = obtenerCorte_(fecha ? { fecha: fecha } : {}, true);
  var m = {};
  var dar = function (nf) { return (m[nf] = m[nf] || { NoFinca: nf }); };
  var poner = function (o, k, v) { if (v !== null && v !== undefined && v !== '' && (o[k] === undefined || o[k] === null || o[k] === '')) o[k] = v; };
  var ordenVentas = (c.ventas || []).slice().sort(function (a, b) { return (a.Tabla === 'FINALIZADA' ? -1 : 0) - (b.Tabla === 'FINALIZADA' ? -1 : 0); });
  ordenVentas.forEach(function (v) {
    var o = dar(v.NoFinca);
    poner(o, 'CodigoUbic', v.CodigoUbic); poner(o, 'Ubicacion', v.Ubicacion); poner(o, 'Superficie', v.Superficie);
    poner(o, 'ValorVentaNegociado', v.ValorVentaNegociado); poner(o, 'ValorLibros', v.ValorLibros);
    poner(o, 'ValorVentaRapida', v.ValorVentaRapida); poner(o, 'Cliente', v.Cliente); poner(o, 'Ejecutivo', v.Ejecutivo);
    poner(o, 'AntiguoPropietario', v.AntiguoPropietario);
    poner(o, 'EstadoVenta', v.Tabla === 'FINALIZADA' ? 'Vendida' : v.Tabla === 'NO_CONCRETADA' ? 'No concretada' : 'En proceso');
  });
  (c.trazabilidad || []).forEach(function (t) {
    var o = dar(t.NoFinca);
    poner(o, 'FechaAdjudicacion', t.FechaAdjudicacion); poner(o, 'TipoPropiedad', t.TipoPropiedad);
    poner(o, 'Ubicacion', t.Ubicacion); poner(o, 'ValorLibros', t.ValorLibros); poner(o, 'ValorVentaRapida', t.ValorVentaRapida);
  });
  (c.visitas || []).forEach(function (v) {
    var o = dar(v.NoFinca);
    poner(o, 'CodigoUbic', v.CodigoUbic); poner(o, 'Ubicacion', v.Ubicacion); poner(o, 'Mejoras', v.TipoPropiedad);
    poner(o, 'Condicion', v.Condicion);
  });
  // Última gestión y última oferta de cada finca
  var gs = ord_(T_('BR_Gestion').filter(function (g) { return !g.Anulada && (!fecha || iso_(g.Fecha) <= fecha); }), 'Fecha DESC, Id DESC');
  var vistoG = {}, vistoO = {};
  gs.forEach(function (r) {
    var o = dar(r.NoFinca);
    if (!vistoG[r.NoFinca]) { vistoG[r.NoFinca] = 1; o.UltimaGestion = r.Fecha; }
    if (r.MontoOferta !== null && !vistoO[r.NoFinca]) {
      vistoO[r.NoFinca] = 1;
      o.UltimaOferta = Number(r.MontoOferta); o.UltimaOfertaCliente = r.Cliente; o.UltimaOfertaEstado = r.EstadoOferta; o.UltimaOfertaContacto = r.Contacto;
    }
  });
  return Object.keys(m).sort().map(function (k) { return m[k]; });
}
ruta_('GET', '/api/bienes/fincas', 'cualquiera', function () {
  // Lo de los reportes, más las fincas del Dashboard (Excel de activos) que no salen en ningún reporte
  var l = fincasConocidas_(), por = {};
  l.forEach(function (c) { por[c.NoFinca] = c; });
  T_('BR_Finca').forEach(function (f) {
    var o = por[f.NoFinca];
    if (!o) { o = por[f.NoFinca] = { NoFinca: f.NoFinca }; l.push(o); }
    [['CodigoUbic', f.CodigoUbic], ['Ubicacion', f.Ubicacion], ['ValorLibros', f.ValorLibros], ['ValorVentaRapida', f.ValorVentaRapida], ['TipoPropiedad', f.TipoPropiedad],
      ['AntiguoPropietario', f.Deudor], ['FechaAdjudicacion', f.FechaAdjudicacion], ['Superficie', f.Superficie], ['ValorMercado', f.ValorMercado]]
      .forEach(function (x) { if (x[1] != null && x[1] !== '' && (o[x[0]] == null || o[x[0]] === '')) o[x[0]] = x[1]; });
  });
  l.sort(function (a, b) { return String(a.NoFinca).localeCompare(String(b.NoFinca)); });
  return { success: true, fincas: l };
});

// =============================================================================
// INVENTARIO DE FINCAS (DASHBOARD)
// =============================================================================
function sugerirEstado_(c) { return c.EstadoVenta === 'Vendida' ? 'VENDIDA' : c.EstadoVenta === 'En proceso' ? 'EN_TRAMITE' : 'DISPONIBLE'; }
function sugerirMejora_(t) { return /sin mejora/i.test(t || '') ? 'SIN' : /con mejora|residencia|casa|apartamento/i.test(t || '') ? 'CON' : null; }

function cuadroInventario_(filas) {
  var act = filas.filter(function (f) { return f.Activa; });
  var n = function (e) { return act.filter(function (f) { return f.Estado === e; }).length; };
  var disp = act.filter(function (f) { return f.Estado === 'DISPONIBLE'; });
  return {
    totalActivos: act.length, enTramite: n('EN_TRAMITE'), vendidas: n('VENDIDA'),
    enTramiteOVendidos: n('EN_TRAMITE') + n('VENDIDA'), enLegal: n('LEGAL'), disponibles: disp.length,
    // Mejoras con la regla del cuadro de Farzin: las alquiladas cuentan como "sin mejora" aunque tengan mejoras
    alquiladas: disp.filter(function (f) { return f.Alquilada; }).length,
    conMejora: disp.filter(function (f) { return f.Mejora === 'CON' && !f.Alquilada; }).length,
    mejoraSinDefinir: disp.filter(function (f) { return !f.Mejora && !f.Alquilada; }).length,
    sinMejora: disp.filter(function (f) { return f.Alquilada || f.Mejora === 'SIN'; }).length,
    alquiladasConMejora: disp.filter(function (f) { return f.Mejora === 'CON' && f.Alquilada; }).map(function (f) { return f.NoFinca; }),
    alquiladasFueraDeDisponibles: act.filter(function (f) { return f.Alquilada && f.Estado !== 'DISPONIBLE'; }).length,
    ocupadas: disp.filter(function (f) { return f.Ocupada; }).length
  };
}

// Mantiene el Dashboard al día con los reportes subidos (agrega fincas nuevas y avanza el estatus)
function sincronizarInventario_(usuario) {
  var conocidas = fincasConocidas_();
  var listaManda = hayExcelDeActivos_();
  var hay = {}; T_('BR_Finca').forEach(function (r) { hay[r.NoFinca] = r; });
  var orden = { DISPONIBLE: 0, EN_TRAMITE: 1, VENDIDA: 2 };
  var agregadas = 0, actualizadas = 0;
  conocidas.forEach(function (c) {
    var sug = sugerirEstado_(c), f = hay[c.NoFinca];
    if (!f && listaManda) return;
    if (!f) {
      ins_('BR_Finca', { NoFinca: c.NoFinca, Estado: sug, Mejora: sugerirMejora_(c.Mejoras || c.TipoPropiedad),
        TipoPropiedad: String(c.TipoPropiedad || c.Mejoras || '').slice(0, 100) || null, Ubicacion: String(c.Ubicacion || '').slice(0, 500) || null,
        CodigoUbic: c.CodigoUbic ? String(c.CodigoUbic).slice(0, 20) : null, ValorLibros: c.ValorLibros != null ? c.ValorLibros : null, UsuarioCrea: usuario });
      registrarCambio_(c.NoFinca, 'Alta', null, NOMBRE_ESTADO[sug] + ' (traída de los reportes)', usuario);
      agregadas++;
    } else if (f.Activa && c.EstadoVenta && orden[f.Estado] !== undefined && orden[sug] > orden[f.Estado]) {
      var antes = f.Estado;
      upd_('BR_Finca', f, { Estado: sug, UsuarioModifica: usuario, FechaModifica: ahora_() });
      registrarCambio_(c.NoFinca, 'Estado', NOMBRE_ESTADO[antes], NOMBRE_ESTADO[sug] + ' (según el reporte de ventas)', usuario);
      actualizadas++;
    }
  });
  if (agregadas || actualizadas) anotar_('Dashboard al día: ' + agregadas + ' fincas agregadas, ' + actualizadas + ' con estatus avanzado según el reporte');
  return { agregadas: agregadas, actualizadas: actualizadas };
}

// Sin fecha: el Dashboard de hoy. Con fecha (AAAA-MM-DD): cómo estaba ese día.
function leerInventario_(fecha) {
  var filas = ord_(T_('BR_Finca'), 'Activa DESC, Estado, NoFinca').map(function (f) { return copia_(f); });
  var conocidas = fincasConocidas_(fecha);
  var porFinca = {}; conocidas.forEach(function (c) { porFinca[c.NoFinca] = c; });
  var corte = obtenerCorte_(fecha ? { fecha: fecha } : {}, true);
  if (fecha) {
    var hist = {};
    ord_(T_('BR_FincaHistorial').filter(function (h) { return iso_(h.Fecha) <= fecha && ['Alta', 'Estado', 'Alquilada', 'Mejora', 'Activa'].indexOf(h.Campo) !== -1; }), 'Id')
      .forEach(function (h) { var x = hist[h.NoFinca] = hist[h.NoFinca] || {}; x[h.Campo === 'Alta' ? 'Estado' : h.Campo] = h.ValorNuevo; });
    var estadoDeTexto = function (t) { return ESTADOS_FINCA.filter(function (k) { return String(t || '').indexOf(NOMBRE_ESTADO[k]) === 0; })[0] || null; };
    filas = filas.filter(function (f) {
      var h = hist[f.NoFinca], c = porFinca[f.NoFinca];
      if (h && h.Estado && estadoDeTexto(h.Estado)) {
        f.Estado = estadoDeTexto(h.Estado);
        f.Alquilada = h.Alquilada !== undefined ? h.Alquilada === 'Sí' : false;
        if (h.Mejora !== undefined) f.Mejora = h.Mejora === 'Con mejora' ? 'CON' : h.Mejora === 'Sin mejora' ? 'SIN' : null;
        f.Activa = h.Activa !== undefined ? h.Activa === 'Activa' : true;
        return true;
      }
      if (c && c.EstadoVenta !== undefined) { f.Estado = sugerirEstado_(c); f.Alquilada = false; f.Activa = true; return true; }
      if (c) { f.Estado = 'DISPONIBLE'; f.Alquilada = false; f.Activa = true; return true; }
      return false;
    });
  }
  var visita = {};
  (corte.visitas || []).forEach(function (v) { visita[v.NoFinca] = { condicion: v.Condicion, letrero: v.Letrero, observaciones: v.Observaciones,
    limpieza: v.Limpieza, proximaLimpieza: v.ProximaLimpieza, tipo: v.TipoPropiedad }; });
  var gs = ord_(T_('BR_Gestion').filter(function (g) { return !g.Anulada && (!fecha || iso_(g.Fecha) <= fecha); }), 'Fecha DESC, Id DESC');
  var gestion = {}, conteo = {}, canales = {};
  gs.forEach(function (g) {
    var k = conteo[g.NoFinca] = conteo[g.NoFinca] || { Visitas: 0, UltimaVisita: null, Publicaciones: 0, Consultas: 0, Interesados: 0, Total: 0 };
    k.Total++;
    if (!gestion[g.NoFinca]) gestion[g.NoFinca] = g;
    if (g.Tipo === 'Visita a la propiedad') { k.Visitas++; if (!k.UltimaVisita || g.Fecha > k.UltimaVisita) k.UltimaVisita = g.Fecha; }
    if (g.Canal) { k.Publicaciones++; var cn = canales[g.NoFinca] = canales[g.NoFinca] || {}; cn[g.Canal] = (cn[g.Canal] || 0) + 1; }
    k.Consultas += g.Consultas || 0;
    if (g.HayInteresado) k.Interesados++;
  });
  var utilidad = {};
  (corte.trazabilidad || []).forEach(function (t) { utilidad[t.NoFinca] = { balance: t.Balance, gastos: t.TotalGastos, venta: t.ValorVentaNegociado, grupo: t.Grupo, fila: t }; });
  var desalojo = {}; ord_(T_('BR_Desalojo'), 'Orden').forEach(function (d) { desalojo[d.NoFinca] = d; });
  var mejoras = mejorasActuales_(fecha).map(function (m) { return copia_(m); }), mejoraDe = {};
  var seguimiento = historialMejoras_(fecha);
  mejoras.forEach(function (m) { m.Historial = seguimiento.hist[m.NoFinca] || []; mejoraDe[m.NoFinca] = m; });
  var avaluos = avaluosPorFinca_();
  var enInventario = {};
  var corteISO = corte.cargaReporte && corte.cargaReporte.FechaReporte ? fechaISO(corte.cargaReporte.FechaReporte) : null;
  filas.forEach(function (f) {
    enInventario[f.NoFinca] = 1;
    var c = porFinca[f.NoFinca] || {};
    f.Ubicacion = f.Ubicacion || c.Ubicacion || null;
    f.TipoPropiedad = f.TipoPropiedad || c.TipoPropiedad || c.Mejoras || null;
    if ((f.ValorLibros === null || f.ValorLibros === undefined) && !f.EstadoExcel) f.ValorLibros = c.ValorLibros != null ? c.ValorLibros : null;
    f.SegunReporte = c.EstadoVenta || null;
    f.EstadoSugerido = c.EstadoVenta ? sugerirEstado_(c) : null;
    f.UltimaOferta = c.UltimaOferta != null ? c.UltimaOferta : null;
    f.UltimaOfertaCliente = c.UltimaOfertaCliente || null;
    f.UltimaOfertaEstado = c.UltimaOfertaEstado || null;
    f.UltimaGestion = c.UltimaGestion || null;
    f.Utilidad = utilidad[f.NoFinca] ? utilidad[f.NoFinca].balance : null;
    f.Traza = utilidad[f.NoFinca] ? utilidad[f.NoFinca].fila : null;
    f.Desalojo = desalojo[f.NoFinca] || null;
    // Cuadro de lotes con mejoras (casas). En el servidor Node esto pisaba f.Mejora (CON/SIN);
    // aquí va aparte para que el Dashboard no pierda la marca con/sin mejora.
    f.LoteMejora = mejoraDe[f.NoFinca] || null;
    f.Gastos = utilidad[f.NoFinca] ? utilidad[f.NoFinca].gastos : null;
    f.ValorVenta = utilidad[f.NoFinca] ? utilidad[f.NoFinca].venta : (c.ValorVentaNegociado != null ? c.ValorVentaNegociado : null);
    f.Visita = visita[f.NoFinca] || null;
    f.Avaluos = avaluos[f.NoFinca] || [];
    var g = gestion[f.NoFinca];
    f.UltGestion = g ? { fecha: g.Fecha, tipo: g.Tipo, lugar: g.Lugar, descripcion: g.Descripcion, usuario: g.Usuario } : null;
    var k = conteo[f.NoFinca] || {};
    f.NumGestiones = k.Total || 0;
    f.NumVisitas = k.Visitas || 0; f.UltimaVisita = k.UltimaVisita || null;
    var fechasExcel = {};
    if (f.Visita) {
      var mejor = null, rx = /(\d{1,2})\/(\d{1,2})\/(\d{4})/g, mm;
      var txt = [f.Visita.observaciones, f.Visita.limpieza].join(' ');
      while ((mm = rx.exec(txt))) {
        var iso = mm[3] + '-' + ('0' + mm[2]).slice(-2) + '-' + ('0' + mm[1]).slice(-2);
        if ((!corteISO || iso <= corteISO) && (!mejor || iso > mejor)) mejor = iso;
      }
      if (mejor || corteISO) fechasExcel[mejor || corteISO] = 'reporte de visitas';
    }
    if (f.Desalojo && f.Desalojo.FechaVisita) { var i1 = fechaISO(f.Desalojo.FechaVisita); if (!fecha || i1 <= fecha) fechasExcel[i1] = fechasExcel[i1] || 'gestión de desalojo'; }
    if (mejoraDe[f.NoFinca] && mejoraDe[f.NoFinca].FechaVisita) { var i2 = fechaISO(mejoraDe[f.NoFinca].FechaVisita); if (!fecha || i2 <= fecha) fechasExcel[i2] = fechasExcel[i2] || 'lotes con mejoras'; }
    f.VisitasExcel = Object.keys(fechasExcel).length;
    f.FechasVisitaExcel = Object.keys(fechasExcel).sort();
    f.NumVisitas += f.VisitasExcel;
    var ultEx = Object.keys(fechasExcel).sort().pop() || null;
    if (ultEx && (!f.UltimaVisita || ultEx > fechaISO(f.UltimaVisita))) f.UltimaVisita = ultEx;
    f.NumPublicaciones = k.Publicaciones || 0; f.NumConsultas = k.Consultas || 0; f.NumInteresados = k.Interesados || 0;
    f.Canales = canales[f.NoFinca] || {};
  });
  var sinInventario = conocidas.filter(function (c) { return !enInventario[c.NoFinca]; }).length;
  return { filas: filas, sinInventario: sinInventario, mejoras: mejoras, cortesMejoras: seguimiento.cortes, cuadro: cuadroInventario_(filas),
    cargaReporte: corte.cargaReporte, ventasReporte: corte.cuadro, cargaTrazabilidad: corte.cargaTrazabilidad, fechaVista: fecha || null };
}

var DE_ARCHIVO_ = /\(Excel de activos|según el reporte de ventas|^Retirada: no está|ya no está en proceso|traída de los reportes/;
ruta_('GET', '/api/bienes/inventario', 'cualquiera', function (ctx) {
  var fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(ctx.query.fecha || '')) && aFecha_(ctx.query.fecha) ? ctx.query.fecha : null;
  var inv = leerInventario_(fecha);
  inv.ventasAnteriores = ord_(T_('BR_VentaAnterior'), 'Anio DESC, Orden');
  try { inv.pieActivos = cfg_('activos_pie') ? JSON.parse(cfg_('activos_pie')) : null; } catch (e) { inv.pieActivos = null; }
  var H = ord_(T_('BR_FincaHistorial').filter(function (h) { return !fecha || iso_(h.Fecha) <= fecha; }), 'Id DESC');
  inv.recientes = H.filter(function (h) { return h.Campo !== 'Alta' && !DE_ARCHIVO_.test(h.ValorNuevo || ''); }).slice(0, 15)
    .map(function (h) { return copia_(h, ['NoFinca', 'Campo', 'ValorAnterior', 'ValorNuevo', 'Usuario', 'Fecha']); });
  // Importaciones agrupadas por persona y minuto
  var imp = {}, ordenImp = [];
  H.filter(function (h) { return DE_ARCHIVO_.test(h.ValorNuevo || ''); }).forEach(function (h) {
    var k = h.Usuario + '|' + (h.Fecha ? h.Fecha.toISOString().slice(0, 16) : '');
    var x = imp[k];
    if (!x) { x = imp[k] = { Usuario: h.Usuario, Ultima: h.Fecha, f: {}, Altas: 0, Retiros: 0, Estatus: 0, act: false }; ordenImp.push(x); }
    x.f[h.NoFinca] = 1;
    if (h.Campo === 'Alta') x.Altas++; if (h.Campo === 'Activa') x.Retiros++; if (h.Campo === 'Estado') x.Estatus++;
    if (/Excel de activos|^Retirada/.test(h.ValorNuevo || '')) x.act = true;
  });
  inv.importaciones = ordenImp.slice(0, 8).map(function (x) { return { Usuario: x.Usuario, Ultima: x.Ultima, Fincas: Object.keys(x.f).length, Altas: x.Altas,
    Retiros: x.Retiros, Estatus: x.Estatus, Origen: x.act ? 'el Excel de activos' : 'el reporte de ventas' }; });
  var altas = {}, ordenAl = [];
  H.filter(function (h) { return h.Campo === 'Alta'; }).forEach(function (h) {
    var k = (h.Fecha ? fechaISO(h.Fecha) : '') + '|' + h.Usuario;
    if (!altas[k]) { altas[k] = { Dia: h.Fecha ? aFecha_(fechaISO(h.Fecha)) : null, Usuario: h.Usuario, Fincas: 0, Ultima: h.Fecha }; ordenAl.push(altas[k]); }
    altas[k].Fincas++;
  });
  inv.altas = ordenAl.slice(0, 5);
  return Object.assign({ success: true, estados: NOMBRE_ESTADO }, inv);
});

ruta_('POST', '/api/bienes/inventario/importar', 'gestor', function (ctx) {
  var r = sincronizarInventario_(ctx.br.Nombre);
  bitacora_(ctx, 'INVENTARIO_IMPORTAR', r.agregadas + ' fincas agregadas, ' + r.actualizadas + ' actualizadas');
  return { success: true, agregadas: r.agregadas, actualizadas: r.actualizadas };
});

ruta_('POST', '/api/bienes/inventario', 'gestor', function (ctx) {
  var b = ctx.body || {};
  var nf = String(b.noFinca || '').trim().slice(0, 30);
  if (!nf) throw err_('Escribe el número de finca.', 400);
  var est = ESTADOS_FINCA.indexOf(b.estado) !== -1 ? b.estado : 'DISPONIBLE';
  if (fincaDe_(nf)) throw err_('La finca ' + nf + ' ya está en el inventario.', 409);
  ins_('BR_Finca', { NoFinca: nf, Estado: est, Alquilada: !!b.alquilada, Mejora: b.mejora === 'CON' || b.mejora === 'SIN' ? b.mejora : null,
    TipoPropiedad: String(b.tipo || '').slice(0, 100) || null, Ubicacion: String(b.ubicacion || '').slice(0, 500) || null,
    ValorLibros: numero(b.valorLibros), Nota: String(b.nota || '').slice(0, 500) || null, UsuarioCrea: ctx.br.Nombre });
  registrarCambio_(nf, 'Alta', null, NOMBRE_ESTADO[est] + ' (agregada a mano)', ctx.br.Nombre);
  bitacora_(ctx, 'INVENTARIO_AGREGAR', 'Finca ' + nf);
  return { success: true };
});

// Cambiar estatus, alquilada, mejora, datos o nota. Cada cambio queda en el historial.
ruta_('PUT', '/api/bienes/inventario/:finca', 'gestor', function (ctx) {
  var nf = String(ctx.params.finca || '').trim(), b = ctx.body || {};
  var f = fincaDe_(nf);
  if (!f) throw err_('La finca ' + nf + ' no está en el inventario.', 404);
  var nuevo = {
    Estado: b.estado !== undefined ? b.estado : f.Estado,
    Alquilada: b.alquilada !== undefined ? !!b.alquilada : !!f.Alquilada,
    Mejora: b.mejora !== undefined ? (b.mejora === 'CON' || b.mejora === 'SIN' ? b.mejora : null) : f.Mejora,
    TipoPropiedad: b.tipo !== undefined ? (String(b.tipo).slice(0, 100) || null) : f.TipoPropiedad,
    Ubicacion: b.ubicacion !== undefined ? (String(b.ubicacion).slice(0, 500) || null) : f.Ubicacion,
    ValorLibros: b.valorLibros !== undefined ? numero(b.valorLibros) : f.ValorLibros,
    Nota: b.nota !== undefined ? (String(b.nota).slice(0, 500) || null) : f.Nota,
    Activa: b.activa !== undefined ? !!b.activa : !!f.Activa
  };
  if (ESTADOS_FINCA.indexOf(nuevo.Estado) === -1) throw err_('Estatus inválido.', 400);
  var txt = { Estado: function (v) { return NOMBRE_ESTADO[v] || v; }, Alquilada: function (v) { return v ? 'Sí' : 'No'; },
    Mejora: function (v) { return v === 'CON' ? 'Con mejora' : v === 'SIN' ? 'Sin mejora' : 'Sin definir'; },
    Activa: function (v) { return v ? 'Activa' : 'Retirada del inventario'; } };
  var cambios = Object.keys(nuevo).filter(function (k) {
    var a = f[k] === undefined ? null : f[k], n = nuevo[k];
    if (k === 'ValorLibros') return (a === null ? null : Number(a)) !== (n === null ? null : Number(n));
    if (k === 'Alquilada' || k === 'Activa') return !!a !== !!n;
    return (a || null) !== (n || null);
  });
  if (!cambios.length) return { success: true, sinCambios: true };
  var antes = copia_(f);
  upd_('BR_Finca', f, Object.assign({}, nuevo, { UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() }));
  cambios.forEach(function (k) {
    var fmt = txt[k] || function (v) { return v; };
    registrarCambio_(nf, k, antes[k] === null ? null : fmt(antes[k]), nuevo[k] === null ? null : fmt(nuevo[k]), ctx.br.Nombre);
  });
  bitacora_(ctx, 'INVENTARIO_CAMBIO', 'Finca ' + nf + ': ' + cambios.map(function (k) { return k + ' → ' + (txt[k] ? txt[k](nuevo[k]) : nuevo[k]); }).join(', '));
  return { success: true };
});

ruta_('GET', '/api/bienes/inventario/:finca/historial', 'cualquiera', function (ctx) {
  var nf = String(ctx.params.finca || '').trim();
  return { success: true, historial: ord_(T_('BR_FincaHistorial').filter(function (h) { return h.NoFinca === nf; }), 'Id DESC').slice(0, 300)
    .map(function (h) { return copia_(h, ['Id', 'Campo', 'ValorAnterior', 'ValorNuevo', 'Usuario', 'Fecha']); }) };
});

// Pasa los números del inventario a los conteos del corte vigente
ruta_('POST', '/api/bienes/inventario/aplicar-cuadro', 'gestor', function (ctx) {
  var inv = leerInventario_();
  if (!inv.cargaReporte) throw err_('Todavía no hay un reporte de ventas subido para ponerle los conteos.', 400);
  var c = inv.cuadro, carga = porId_('BR_Carga', inv.cargaReporte.Id);
  if (!carga || carga.Anulada) throw err_('No se pudo actualizar el corte.', 400);
  upd_('BR_Carga', carga, { TotalActivos: c.totalActivos, EnLegal: c.enLegal, Alquiladas: c.alquiladas, ConMejora: c.conMejora });
  bitacora_(ctx, 'CONTEOS', 'Desde el inventario, carga #' + carga.Id + ': ' + JSON.stringify(c));
  return { success: true, cargaId: carga.Id, cuadro: c };
});

ruta_('GET', '/api/bienes/inventario-exportar', 'cualquiera', function (ctx) {
  var inv = leerInventario_();
  bitacora_(ctx, 'DESCARGAR', 'Inventario (' + inv.filas.length + ' fincas)');
  return { success: true, __excel: 'inventario', nombre: 'Inventario_bienes_reposeidos_' + hoyISO_() + '.xlsx', filas: inv.filas };
});

// =============================================================================
// VENTAS ANOTADAS A MANO
// =============================================================================
var TABLAS_VENTA = ['EN_PROCESO', 'FINALIZADA', 'NO_CONCRETADA'];
var NOMBRE_TABLA = { EN_PROCESO: 'en proceso', FINALIZADA: 'vendida', NO_CONCRETADA: 'no concretada' };
function validarVentaManual_(b) {
  var v = {
    NoFinca: String(b.noFinca || '').trim().slice(0, 30), Tabla: String(b.tabla || ''), Fecha: String(b.fecha || '').slice(0, 10),
    CodigoUbic: String(b.codigoUbic || '').trim().slice(0, 20) || null, Ubicacion: String(b.ubicacion || '').trim().slice(0, 500) || null,
    Superficie: String(b.superficie || '').trim().slice(0, 80) || null, Cliente: String(b.cliente || '').trim().slice(0, 250) || null,
    Ejecutivo: String(b.ejecutivo || '').trim().slice(0, 100) || null, Estatus: String(b.estatus || '').trim() || null,
    AntiguoPropietario: String(b.antiguoPropietario || '').trim().slice(0, 250) || null
  };
  if (!v.NoFinca) return { error: 'Escribe el número de finca.' };
  if (TABLAS_VENTA.indexOf(v.Tabla) === -1) return { error: 'Elige el estado de la venta.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v.Fecha) || !aFecha_(v.Fecha)) return { error: 'La fecha no es válida.' };
  var nums = [['valorVentaNegociado', 'ValorVentaNegociado'], ['valorLibros', 'ValorLibros'], ['valorVentaRapida', 'ValorVentaRapida'],
    ['montoFinanciado', 'MontoFinanciado'], ['comisionPagada', 'ComisionPagada']];
  for (var i = 0; i < nums.length; i++) {
    var k = nums[i][0], K = nums[i][1];
    if (b[k] === '' || b[k] === null || b[k] === undefined) { v[K] = null; continue; }
    v[K] = numero(b[k]);
    if (v[K] === null) return { error: 'Revisa el valor: ' + k.replace(/([A-Z])/g, ' $1').toLowerCase() + '.' };
  }
  if (v.Tabla !== 'NO_CONCRETADA' && !v.Cliente) return { error: 'Escribe el cliente o comprador.' };
  if (v.Tabla !== 'NO_CONCRETADA' && v.ValorVentaNegociado === null) return { error: 'Escribe el valor de venta negociado.' };
  return { v: v };
}
ruta_('POST', '/api/bienes/ventas-manuales', 'gestor', function (ctx) {
  var r = validarVentaManual_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var v = r.v, fila = ins_('BR_VentaManual', Object.assign({}, v, { UsuarioCrea: ctx.br.Nombre }));
  bitacora_(ctx, 'VENTA_MANUAL', 'Finca ' + v.NoFinca + ' ' + NOMBRE_TABLA[v.Tabla] + (v.ValorVentaNegociado != null ? ' · ' + v.ValorVentaNegociado : ''));
  try { sincronizarInventario_(ctx.br.Nombre); } catch (e) { anotar_('no se pudo poner al día el Dashboard', e); }
  return { success: true, id: fila.Id };
});
ruta_('PUT', '/api/bienes/ventas-manuales/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), r = validarVentaManual_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = porId_('BR_VentaManual', id);
  if (!f || f.Anulada) throw err_('Esa venta no existe o ya se quitó.', 404);
  upd_('BR_VentaManual', f, Object.assign({}, r.v, { UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() }));
  bitacora_(ctx, 'VENTA_MANUAL', 'Editó venta #' + id + ' finca ' + r.v.NoFinca + ' → ' + NOMBRE_TABLA[r.v.Tabla]);
  try { sincronizarInventario_(ctx.br.Nombre); } catch (e) { anotar_('no se pudo poner al día el Dashboard', e); }
  return { success: true };
});
ruta_('POST', '/api/bienes/ventas-manuales/:id/quitar', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), f = porId_('BR_VentaManual', id);
  if (!f || f.Anulada) throw err_('Esa venta no existe o ya se quitó.', 404);
  upd_('BR_VentaManual', f, { Anulada: true, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  bitacora_(ctx, 'VENTA_MANUAL', 'Quitó venta anotada #' + id);
  return { success: true };
});

// =============================================================================
// CONTACTOS DE REDES
// =============================================================================
var CANALES_CONTACTO = ['Instagram', 'Facebook', 'Marketplace', 'Encuentra24', 'WhatsApp', 'Página web', 'Otro'];
var ESTADOS_CONTACTO = ['NUEVO', 'RESPONDIDO', 'INTERESADO', 'VISITA', 'OFERTA', 'NO_INTERESADO'];
function validarContacto_(b) {
  var c = {
    Fecha: String(b.fecha || '').slice(0, 10), Canal: String(b.canal || ''), Nombre: String(b.nombre || '').trim().slice(0, 150),
    Contacto: String(b.contacto || '').trim().slice(0, 150) || null, NoFinca: String(b.noFinca || '').trim().slice(0, 30) || null,
    Mensaje: String(b.mensaje || '').trim() || null, Estado: String(b.estado || 'NUEVO'),
    Seguimiento: String(b.seguimiento || '').trim().slice(0, 1000) || null
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(c.Fecha) || !aFecha_(c.Fecha)) return { error: 'La fecha no es válida.' };
  if (CANALES_CONTACTO.indexOf(c.Canal) === -1) return { error: 'Elige por dónde escribió.' };
  if (!c.Nombre) return { error: 'Escribe el nombre del cliente.' };
  if (ESTADOS_CONTACTO.indexOf(c.Estado) === -1) return { error: 'Estado inválido.' };
  return { c: c };
}
ruta_('GET', '/api/bienes/contactos', 'cualquiera', function (ctx) {
  var f = ctx.query.finca ? String(ctx.query.finca).trim() : null;
  var l = ord_(T_('BR_Contacto').filter(function (x) { return !x.Anulado && (!f || x.NoFinca === f); }), 'Fecha DESC, Id DESC').slice(0, 3000);
  return { success: true, contactos: l, canales: CANALES_CONTACTO };
});
ruta_('POST', '/api/bienes/contactos', 'gestor', function (ctx) {
  var r = validarContacto_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = ins_('BR_Contacto', Object.assign({}, r.c, { Usuario: ctx.br.Nombre }));
  bitacora_(ctx, 'CONTACTO', r.c.Canal + ': ' + r.c.Nombre + (r.c.NoFinca ? ' (finca ' + r.c.NoFinca + ')' : ''));
  return { success: true, id: f.Id };
});
ruta_('PUT', '/api/bienes/contactos/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), r = validarContacto_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = porId_('BR_Contacto', id); if (!f || f.Anulado) throw err_('Ese contacto no existe o se anuló.', 404);
  upd_('BR_Contacto', f, Object.assign({}, r.c, { UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() }));
  bitacora_(ctx, 'CONTACTO', 'Editó contacto #' + id + ' (' + r.c.Nombre + ')');
  return { success: true };
});
ruta_('POST', '/api/bienes/contactos/:id/anular', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), f = porId_('BR_Contacto', id);
  if (!f || f.Anulado) throw err_('Ese contacto no existe o ya se anuló.', 404);
  upd_('BR_Contacto', f, { Anulado: true, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  bitacora_(ctx, 'CONTACTO', 'Anuló contacto #' + id);
  return { success: true };
});

// =============================================================================
// HISTORIAL DE CARGAS
// =============================================================================
ruta_('GET', '/api/bienes/cargas', 'cualquiera', function () {
  var nV = {}, nT = {};
  T_('BR_Venta').forEach(function (v) { nV[v.CargaId] = (nV[v.CargaId] || 0) + 1; });
  T_('BR_Trazabilidad').forEach(function (t) { nT[t.CargaId] = (nT[t.CargaId] || 0) + 1; });
  var l = ord_(T_('BR_Carga').filter(function (c) { return c.TipoArchivo !== 'ALQUILERES'; }), 'Id DESC')
    .map(function (c) { return Object.assign(copia_(c, COLS_CARGA), { FilasVentas: nV[c.Id] || 0, FilasTrazabilidad: nT[c.Id] || 0 }); });
  return { success: true, cargas: l };
});
// Descargar el Excel original tal como se subió
ruta_('GET', '/api/bienes/cargas/:id/archivo', 'cualquiera', function (ctx) {
  var id = parseInt(ctx.params.id, 10); if (!id) throw err_('Carga inválida.', 400);
  var c = porId_('BR_Carga', id); if (!c) throw err_('No existe la carga #' + id + '.', 404);
  if (!c.ArchivoId) throw err_('Esa carga no tiene el archivo guardado.', 404);
  bitacora_(ctx, 'DESCARGAR_ORIGINAL', 'Carga #' + id + ': ' + c.NombreArchivo);
  return archivoRespuesta_(c.NombreArchivo, leerArchivo_(c.ArchivoId));
});
// Corregir los conteos a mano de una carga
ruta_('PUT', '/api/bienes/cargas/:id/conteos', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10); if (!id) throw err_('Carga inválida.', 400);
  var vc = validarConteos(ctx.body); if (vc.error) throw err_(vc.error, 400);
  var cuadro = calcularCuadro(filasDeCarga_('BR_Venta', id), vc.conteos);
  if (cuadro.disponibles !== null && cuadro.disponibles < 0) throw err_('Revisa los conteos: ' + cuadro.avisos[0], 400);
  var c = porId_('BR_Carga', id);
  if (!c || c.Anulada || ['REPORTE', 'MIXTO'].indexOf(c.TipoArchivo) === -1) throw err_('La carga no existe, está anulada o no es un reporte de ventas.', 404);
  upd_('BR_Carga', c, { TotalActivos: vc.conteos.totalActivos, EnLegal: vc.conteos.enLegal, Alquiladas: vc.conteos.alquiladas, ConMejora: vc.conteos.conMejora });
  bitacora_(ctx, 'CONTEOS', 'Carga #' + id + ': ' + JSON.stringify(vc.conteos));
  return { success: true, cuadro: cuadro };
});

// Después de eliminar o anular un Excel de activos: se vuelve a aplicar el último que quede
// (con lo que se guardó de él); si no queda ninguno, se quita lo que venía de ese Excel.
function rehacerActivos_() {
  quitarCfg_(['activos_pie']);
  var queda = ultimaActivos_();
  if (queda && queda.Lectura && queda.Lectura.activos) { guardarDatosActivos_(queda.Lectura, queda.Id, 'Sistema (vuelve al Excel de activos anterior)'); return; }
  if (!queda) { del_('BR_Desalojo', function () { return true; }); del_('BR_VentaAnterior', function () { return true; }); }
}
// Si ya no queda ninguna carga, las fincas del Dashboard también se borran
function limpiarSiNoHayCargas_() {
  if (T_('BR_Carga').some(function (c) { return !c.Anulada && c.TipoArchivo !== 'ALQUILERES'; })) return false;
  ['BR_FincaHistorial', 'BR_Finca', 'BR_Desalojo', 'BR_VentaAnterior', 'BR_Mejora'].forEach(function (t) { del_(t, function () { return true; }); });
  quitarCfg_(['activos_version', 'activos_pie']);
  return true;
}
ruta_('POST', '/api/bienes/borrar-todo', 'admin', function (ctx) {
  if (String((ctx.body || {}).confirmar || '').trim().toUpperCase() !== 'BORRAR TODO') throw err_('Falta confirmar escribiendo BORRAR TODO.', 400);
  var antes = { cargas: T_('BR_Carga').filter(function (c) { return c.TipoArchivo !== 'ALQUILERES'; }).length, fincas: T_('BR_Finca').length };
  ['BR_Venta', 'BR_Visita', 'BR_ResumenEjecutivo', 'BR_Trazabilidad', 'BR_Desalojo', 'BR_VentaAnterior', 'BR_Mejora', 'BR_FincaHistorial', 'BR_Finca']
    .forEach(function (t) { del_(t, function () { return true; }); });
  del_('BR_Carga', function (c) { if (c.TipoArchivo !== 'ALQUILERES') { borrarArchivo_(c.ArchivoId); return true; } return false; });
  quitarCfg_(['activos_version', 'activos_pie']);
  bitacora_(ctx, 'BORRAR_TODO', 'Se borraron ' + antes.cargas + ' cargas y ' + antes.fincas + ' fincas del Dashboard');
  return { success: true, cargas: antes.cargas, fincas: antes.fincas };
});
ruta_('POST', '/api/bienes/cargas/:id/eliminar', 'admin', function (ctx) {
  var id = parseInt(ctx.params.id, 10); if (!id) throw err_('Carga inválida.', 400);
  var c = porId_('BR_Carga', id); if (!c) throw err_('No existe la carga #' + id + '.', 404);
  ['BR_Venta', 'BR_Visita', 'BR_ResumenEjecutivo', 'BR_Trazabilidad', 'BR_Desalojo', 'BR_VentaAnterior', 'BR_Mejora']
    .forEach(function (t) { del_(t, function (x) { return x.CargaId === id; }); });
  del_('BR_Carga', function (x) { return x.Id === id; });
  borrarArchivo_(c.ArchivoId);
  if (c.TipoArchivo === 'ACTIVOS') rehacerActivos_();
  if (limpiarSiNoHayCargas_()) bitacora_(ctx, 'BORRAR_TODO', 'Se eliminó la última carga: se vaciaron las fincas del Dashboard');
  bitacora_(ctx, 'ELIMINAR_CARGA', 'Carga #' + id + ' (' + c.TipoArchivo + ', corte ' + (c.FechaReporte ? fechaISO(c.FechaReporte) : 's/f') + '): ' + c.NombreArchivo);
  return { success: true };
});
ruta_('POST', '/api/bienes/cargas/:id/anular', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10); if (!id) throw err_('Carga inválida.', 400);
  var c = porId_('BR_Carga', id); if (!c || c.Anulada) throw err_('La carga no existe o ya estaba anulada.', 404);
  upd_('BR_Carga', c, { Anulada: true, FechaAnulada: ahora_(), UsuarioAnula: ctx.br.Nombre });
  bitacora_(ctx, 'ANULAR', 'Carga #' + id);
  if (c.TipoArchivo === 'ACTIVOS') rehacerActivos_();
  if (limpiarSiNoHayCargas_()) bitacora_(ctx, 'BORRAR_TODO', 'Se anuló la última carga: se vaciaron las fincas del Dashboard');
  return { success: true };
});
ruta_('GET', '/api/bienes/estado', 'cualquiera', function () {
  return { success: true, base: libro_().getName(), servidor: 'Google Sheets', cargas: T_('BR_Carga').filter(function (c) { return !c.Anulada; }).length };
});

// =============================================================================
// CONTROL DE ALQUILERES
// =============================================================================
function fISO_(v) { return v == null || v === '' ? null : v instanceof Date ? fechaISO(v) : /^\d{4}-\d{2}-\d{2}/.test(String(v)) ? String(v).slice(0, 10) : null; }
function contratoValido_(b) {
  var c = {
    Numero: b.Numero != null && b.Numero !== '' ? parseInt(b.Numero, 10) : null, Arrendatario: String(b.Arrendatario || '').trim().slice(0, 200),
    Expediente: String(b.Expediente || '').trim().slice(0, 250) || null, FechaInicio: fISO_(b.FechaInicio), FechaFin: fISO_(b.FechaFin),
    Canon: Number(b.Canon), DiaPago: parseInt(b.DiaPago, 10), Modalidad: b.Modalidad === 'VENCIMIENTO' ? 'VENCIMIENTO' : 'ADELANTADO',
    DuracionMeses: b.DuracionMeses != null && b.DuracionMeses !== '' ? parseInt(b.DuracionMeses, 10) : null,
    Observacion: String(b.Observacion || '').trim().slice(0, 1000) || null, Frecuencia: b.Frecuencia === 'UNICO' ? 'UNICO' : 'MENSUAL'
  };
  var fin = String(b.Finca || '').trim(), mf = fin.match(/^([A-Za-z]{0,2}\d+)\s*(?:-\s*([0-9A-Za-z]+))?$/);
  c.NoFinca = mf ? mf[1] : (fin.slice(0, 30) || null); c.CodigoUbic = mf ? (mf[2] || null) : null;
  if (!c.Arrendatario) throw err_('Falta el nombre del arrendatario.', 400);
  if (!c.FechaInicio || !c.FechaFin) throw err_('Faltan la fecha inicial o la final.', 400);
  if (c.FechaFin < c.FechaInicio) throw err_('La fecha final es anterior a la inicial.', 400);
  if (!(c.Canon > 0)) throw err_('El canon de arrendamiento debe ser mayor que 0.', 400);
  if (!(c.DiaPago >= 1 && c.DiaPago <= 31)) throw err_('El día de pago debe estar entre 1 y 31.', 400);
  if (c.DuracionMeses != null && !(c.DuracionMeses >= 1 && c.DuracionMeses <= 120)) throw err_('La duración debe estar entre 1 y 120 meses.', 400);
  return c;
}
var CAMPOS_ALQ = ['Numero', 'Arrendatario', 'Expediente', 'NoFinca', 'CodigoUbic', 'FechaInicio', 'FechaFin', 'Canon', 'DiaPago', 'Modalidad',
  'DuracionMeses', 'Observacion', 'Frecuencia'];
function leerAlquileresBD_(id) {
  var cs = ord_(T_('BR_Alquiler').filter(function (c) { return !c.Anulado && (!id || c.Id === id); }), 'Numero, FechaInicio, Id');
  // ISNULL(Numero, 9999): los que no tienen número van al final
  cs = cs.filter(function (c) { return c.Numero != null; }).concat(cs.filter(function (c) { return c.Numero == null; }));
  var por = {};
  ord_(T_('BR_AlquilerCuota'), 'AlquilerId, Numero, Id').forEach(function (q) { (por[q.AlquilerId] = por[q.AlquilerId] || []).push(copia_(q)); });
  return cs.map(function (c) { var o = copia_(c); o.Cuotas = por[c.Id] || []; o.Meses = c.DuracionMeses || mesesContrato(fISO_(c.FechaInicio), fISO_(c.FechaFin)); return o; });
}
// Rehace el calendario conservando lo ya registrado: pagos, notas y fechas cambiadas a mano
function rehacerCuotas_(id, contrato, usuario) {
  var viejas = T_('BR_AlquilerCuota').filter(function (q) { return q.AlquilerId === id; });
  var vPor = {}; viejas.forEach(function (v) { vPor[v.Numero] = v; });
  calendarioAlquiler(contrato).forEach(function (q) {
    var v = vPor[q.Numero];
    if (!v) { ins_('BR_AlquilerCuota', Object.assign({ AlquilerId: id, FechaManual: false, RegistradoPor: usuario }, q)); return; }
    upd_('BR_AlquilerCuota', v, { FechaVence: v.FechaManual ? v.FechaVence : q.FechaVence, Monto: v.MontoPagado != null ? v.Monto : q.Monto });
    delete vPor[q.Numero];
  });
  var sobran = Object.keys(vPor).map(function (k) { return vPor[k]; }).filter(function (v) { return v.MontoPagado == null; });
  if (sobran.length) del_('BR_AlquilerCuota', function (q) { return sobran.indexOf(q) !== -1; });
}
function datosContrato_(c) { var o = {}; CAMPOS_ALQ.forEach(function (k) { o[k] = c[k] === undefined ? null : c[k]; }); return o; }

ruta_('GET', '/api/bienes/alquileres', 'cualquiera', function () { return { success: true, contratos: leerAlquileresBD_() }; });

ruta_('POST', '/api/bienes/alquileres', 'gestor', function (ctx) {
  var c = contratoValido_(ctx.body || {});
  var f = ins_('BR_Alquiler', Object.assign(datosContrato_(c), { Origen: 'MANUAL', CreadoPor: ctx.br.Nombre }));
  rehacerCuotas_(f.Id, c, ctx.br.Nombre);
  bitacora_(ctx, 'ALQUILER', 'Nuevo contrato #' + f.Id + ': ' + c.Arrendatario + ' · B/. ' + c.Canon.toFixed(2) + ' mensual');
  return { success: true, id: f.Id, contrato: leerAlquileresBD_(f.Id)[0] };
});
ruta_('PUT', '/api/bienes/alquileres/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), c = contratoValido_(ctx.body || {});
  var f = porId_('BR_Alquiler', id); if (!f || f.Anulado) throw err_('El contrato no existe.', 404);
  upd_('BR_Alquiler', f, Object.assign(datosContrato_(c), { ModificadoPor: ctx.br.Nombre, FechaModifica: ahora_() }));
  if (ctx.body.rehacer) rehacerCuotas_(id, c, ctx.br.Nombre);
  bitacora_(ctx, 'ALQUILER', 'Editó el contrato #' + id + ': ' + c.Arrendatario + (ctx.body.rehacer ? ' (rehízo el calendario de pagos)' : ''));
  return { success: true, contrato: leerAlquileresBD_(id)[0] };
});
ruta_('POST', '/api/bienes/alquileres/:id/anular', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), f = porId_('BR_Alquiler', id);
  if (!f || f.Anulado) throw err_('El contrato no existe o ya estaba anulado.', 404);
  upd_('BR_Alquiler', f, { Anulado: true, ModificadoPor: ctx.br.Nombre, FechaModifica: ahora_() });
  bitacora_(ctx, 'ALQUILER', 'Anuló el contrato #' + id);
  return { success: true };
});
// Cambiar una cuota: fecha que vence, monto, el pago recibido y su fecha, nota
ruta_('PUT', '/api/bienes/alquileres/cuotas/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), b = ctx.body || {};
  var v = porId_('BR_AlquilerCuota', id); if (!v) throw err_('La cuota no existe.', 404);
  var fv = fISO_(b.FechaVence), monto = Number(b.Monto), pagado = b.MontoPagado === '' || b.MontoPagado == null ? null : Number(b.MontoPagado);
  var fp = fISO_(b.FechaPago);
  if (!fv) throw err_('Falta la fecha en que vence.', 400);
  if (!(monto > 0)) throw err_('El monto de la cuota debe ser mayor que 0.', 400);
  if (pagado != null && !(pagado > 0)) throw err_('El monto pagado debe ser mayor que 0 (o dejarlo vacío si no ha pagado).', 400);
  if (pagado != null && !fp) throw err_('Falta la fecha en que pagó.', 400);
  var antesVence = fISO_(v.FechaVence);
  var manual = v.FechaManual || antesVence !== fv;
  upd_('BR_AlquilerCuota', v, { FechaVence: fv, Monto: monto, FechaPago: pagado != null ? fp : null, MontoPagado: pagado,
    Nota: String(b.Nota || '').trim().slice(0, 500) || null, FechaManual: !!manual, RegistradoPor: ctx.br.Nombre, FechaRegistro: ahora_() });
  bitacora_(ctx, 'ALQUILER', 'Cuota ' + v.Numero + ' del contrato #' + v.AlquilerId + ': ' +
    (pagado != null ? 'pago de B/. ' + pagado.toFixed(2) + ' el ' + fp : 'sin pago') + (manual && antesVence !== fv ? ' · vence ' + fv : ''));
  return { success: true, contrato: leerAlquileresBD_(v.AlquilerId)[0] };
});
// Agregar una cuota al final (por ejemplo, un mes extra o una prórroga)
ruta_('POST', '/api/bienes/alquileres/:id/cuotas', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), b = ctx.body || {};
  var fv = fISO_(b.FechaVence), monto = Number(b.Monto);
  if (!fv || !(monto > 0)) throw err_('Faltan la fecha o el monto de la cuota.', 400);
  var n = T_('BR_AlquilerCuota').filter(function (q) { return q.AlquilerId === id; }).reduce(function (a, q) { return Math.max(a, q.Numero || 0); }, 0);
  ins_('BR_AlquilerCuota', { AlquilerId: id, Numero: n + 1, FechaVence: fv, Monto: monto, FechaManual: true, RegistradoPor: ctx.br.Nombre,
    Nota: String(b.Nota || '').trim().slice(0, 500) || null });
  bitacora_(ctx, 'ALQUILER', 'Agregó la cuota ' + (n + 1) + ' al contrato #' + id);
  return { success: true, contrato: leerAlquileresBD_(id)[0] };
});
ruta_('DELETE', '/api/bienes/alquileres/cuotas/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), v = porId_('BR_AlquilerCuota', id);
  if (!v) throw err_('La cuota no existe.', 404);
  if (v.MontoPagado != null) throw err_('Esa cuota tiene un pago registrado; primero quita el pago.', 400);
  del_('BR_AlquilerCuota', function (q) { return q.Id === id; });
  bitacora_(ctx, 'ALQUILER', 'Quitó la cuota ' + v.Numero + ' del contrato #' + v.AlquilerId);
  return { success: true, contrato: leerAlquileresBD_(v.AlquilerId)[0] };
});

// Importa el Excel "Control de Alquileres" (la página ya lo leyó: b.alquileres).
// Devuelve null si el archivo no es un control de alquileres; con soloVista, solo dice qué entra y qué cambia.
function importarAlquileres_(ctx, b, soloVista) {
  var lx = b.alquileres;
  if (!lx || !lx.contratos || !lx.contratos.length) return null;
  var existentes = leerAlquileresBD_();
  var clave = function (a, f) { return norm(a) + '|' + f; };
  var ePor = {}; existentes.forEach(function (e) { ePor[clave(e.Arrendatario, fISO_(e.FechaInicio))] = e; });
  var avisos = (lx.avisos || []).slice(), plan = [];
  lx.contratos.filter(function (c) { return c.FechaInicio && c.FechaFin; }).forEach(function (c) {
    var cuotas = cuotasDesdeExcel(c, avisos), e = ePor[clave(c.Arrendatario, c.FechaInicio)];
    var pagos = cuotas.filter(function (q) { return q.MontoPagado != null; });
    plan.push({ c: c, cuotas: cuotas, existente: e ? e.Id : null, resumen: { Numero: c.Numero, Arrendatario: c.Arrendatario,
      Finca: c.NoFinca ? c.NoFinca + (c.CodigoUbic ? '-' + c.CodigoUbic : '') : null, FechaInicio: c.FechaInicio, FechaFin: c.FechaFin, Canon: c.Canon,
      DiaPago: c.DiaPago, Modalidad: c.Modalidad, Frecuencia: c.Frecuencia, Cuotas: cuotas.length, Pagos: pagos.length,
      Cobrado: Math.round(pagos.reduce(function (a, q) { return a + q.MontoPagado; }, 0) * 100) / 100, Accion: e ? 'Se actualiza' : 'Nuevo', Hoja: c.hoja || null } });
  });
  var vista = { contratos: plan.map(function (p) { return p.resumen; }), avisos: avisos, notas: lx.notas || [], fechaControl: lx.fechaControl || null };
  if (soloVista) return vista;
  if (!b.archivoB64) throw err_('No llegó el archivo.', 400);
  plan.forEach(function (p) {
    var c = Object.assign({}, p.c, { Observacion: (p.c.Observacion || '').slice(0, 1000) || null });
    var id = p.existente;
    if (!id) {
      id = ins_('BR_Alquiler', Object.assign(datosContrato_(c), { Origen: 'EXCEL', CreadoPor: ctx.br.Nombre })).Id;
      p.cuotas.forEach(function (q) { ins_('BR_AlquilerCuota', Object.assign({ AlquilerId: id, FechaManual: false, RegistradoPor: ctx.br.Nombre }, q)); });
    } else {
      upd_('BR_Alquiler', porId_('BR_Alquiler', id), Object.assign(datosContrato_(c), { ModificadoPor: ctx.br.Nombre, FechaModifica: ahora_() }));
      rehacerCuotas_(id, c, ctx.br.Nombre);
      var act = T_('BR_AlquilerCuota').filter(function (x) { return x.AlquilerId === id; });
      p.cuotas.forEach(function (q) {
        var a = act.filter(function (x) { return x.Numero === q.Numero; })[0]; if (!a) return;
        if (q.MontoPagado != null || (q.Nota && q.Nota !== a.Nota))
          upd_('BR_AlquilerCuota', a, { MontoPagado: q.MontoPagado != null ? q.MontoPagado : a.MontoPagado,
            FechaPago: q.MontoPagado != null ? q.FechaPago : a.FechaPago, Nota: q.Nota || a.Nota });
      });
    }
  });
  bitacora_(ctx, 'ALQUILER', 'Subió el control de alquileres (' + String(b.nombreArchivo || '').slice(0, 120) + '): ' + plan.length + ' contratos');
  vista.contratosGuardados = leerAlquileresBD_().length;
  var c2 = insertarCarga_(ctx, Object.assign({}, b, { nombreArchivo: b.nombreArchivo || 'Control_de_Alquileres.xlsx', nota: plan.length + ' contratos' }),
    'ALQUILERES', lx.fechaControl || hoyISO_(), { Lectura: lx });
  vista.cargaId = c2.Id;
  return vista;
}
ruta_('GET', '/api/bienes/alquileres/cargas', 'cualquiera', function () {
  return { success: true, cargas: ordenCortes_(T_('BR_Carga').filter(function (c) { return c.TipoArchivo === 'ALQUILERES'; }))
    .map(function (c) { return copia_(c, ['Id', 'FechaCarga', 'Usuario', 'NombreArchivo', 'FechaReporte', 'TamanoBytes', 'Nota']); }) };
});
// Lo que traía un Excel subido, exactamente (lo leído al subirlo)
ruta_('GET', '/api/bienes/alquileres/carga/:id', 'cualquiera', function (ctx) {
  var id = parseInt(ctx.params.id, 10), c = porId_('BR_Carga', id);
  if (!c || c.TipoArchivo !== 'ALQUILERES') throw err_('No existe ese Excel de alquileres.', 404);
  var lx = c.Lectura || { contratos: [], avisos: [] }, avisos = (lx.avisos || []).slice(), k = 0;
  var contratos = (lx.contratos || []).filter(function (x) { return x.FechaInicio && x.FechaFin; }).map(function (x, i) {
    var cuotas = cuotasDesdeExcel(x, avisos).map(function (q) { return Object.assign({ Id: -(++k), FechaManual: false }, q); });
    return { Id: -(i + 1), Numero: x.Numero, Arrendatario: x.Arrendatario, Expediente: x.Expediente, NoFinca: x.NoFinca || null, CodigoUbic: x.CodigoUbic || null,
      FechaInicio: x.FechaInicio, FechaFin: x.FechaFin, Canon: x.Canon, DiaPago: x.DiaPago, Modalidad: x.Modalidad, Frecuencia: x.Frecuencia,
      DuracionMeses: x.DuracionMeses, Observacion: x.Observacion, Meses: x.DuracionMeses || mesesContrato(x.FechaInicio, x.FechaFin), Cuotas: cuotas };
  });
  return { success: true, carga: copia_(c, ['Id', 'FechaCarga', 'Usuario', 'NombreArchivo', 'FechaReporte']), contratos: contratos, avisos: avisos };
});
ruta_('DELETE', '/api/bienes/alquileres/cargas/:id', 'admin', function (ctx) {
  var id = parseInt(ctx.params.id, 10), c = porId_('BR_Carga', id);
  if (!c || c.TipoArchivo !== 'ALQUILERES') throw err_('No existe ese Excel de alquileres.', 404);
  del_('BR_Carga', function (x) { return x.Id === id; });
  borrarArchivo_(c.ArchivoId);
  bitacora_(ctx, 'ALQUILER', 'Eliminó del historial el Excel de alquileres #' + id + ': ' + c.NombreArchivo);
  return { success: true };
});
ruta_('POST', '/api/bienes/alquileres/importar', 'gestor', function (ctx) {
  var b = ctx.body || {};
  if (!b.alquileres) throw err_('No llegó el archivo.', 400);
  var vista = importarAlquileres_(ctx, b, !!b.soloVista);
  if (!vista) throw err_('El archivo no tiene la hoja de control de alquileres (con "Nombre del arrendatario" y "Canon de arrendamiento").', 400);
  return Object.assign({ success: true, vista: !!b.soloVista }, vista);
});
ruta_('GET', '/api/bienes/alquileres/excel', 'cualquiera', function () {
  return { success: true, __excel: 'alquileres', nombre: 'Control_de_Alquileres_' + hoyISO_() + '.xlsx', contratos: leerAlquileresBD_(), hoy: hoyISO_() };
});

// =============================================================================
// GESTIÓN DE CAMPO
// =============================================================================
var TIPOS_GESTION = ['Visita a la propiedad', 'Publicación (redes o portales)', 'Reunión con cliente', 'Llamada o mensaje',
  'Trámite (Registro, notaría, ANATI)', 'Limpieza o mantenimiento', 'Publicidad / letrero', 'Avalúo', 'Otro'];
var CANALES = ['Instagram', 'Facebook', 'Marketplace', 'Encuentra24', 'WhatsApp', 'Página web', 'Letrero', 'Otro'];
var ESTADOS_OFERTA = ['EN_EVALUACION', 'CONTRAOFERTA', 'ACEPTADA', 'RECHAZADA'];
function esPublicacion_(t) { return t === 'Publicación (redes o portales)' || t === 'Publicidad / letrero'; }
function fechaValida_(f) { return /^\d{4}-\d{2}-\d{2}$/.test(String(f || '')) && !!aFecha_(f); }
function validarGestion_(b) {
  var g = {
    NoFinca: String(b.noFinca || '').trim().slice(0, 30), Fecha: String(b.fecha || '').slice(0, 10), Tipo: String(b.tipo || ''),
    Lugar: String(b.lugar || '').trim().slice(0, 300) || null, Descripcion: String(b.descripcion || '').trim(), HayInteresado: !!b.hayInteresado,
    Cliente: null, Contacto: null, MontoOferta: null, FormaPago: null, EstadoOferta: null,
    ProximoPaso: String(b.proximoPaso || '').trim().slice(0, 500) || null, FechaProximo: b.fechaProximo ? String(b.fechaProximo).slice(0, 10) : null,
    Canal: null, Enlace: null, Consultas: null
  };
  if (esPublicacion_(g.Tipo)) {
    g.Canal = CANALES.indexOf(b.canal) !== -1 ? b.canal : null;
    if (g.Tipo === 'Publicación (redes o portales)' && !g.Canal) return { error: 'Elige dónde se publicó (Instagram, Facebook, …).' };
    var enl = String(b.enlace || '').trim().slice(0, 500);
    if (enl && !/^https?:\/\//i.test(enl)) return { error: 'El enlace debe empezar con http:// o https://' };
    g.Enlace = enl || null;
    if (b.consultas !== '' && b.consultas !== null && b.consultas !== undefined) {
      var c = Number(b.consultas);
      if (!Number.isInteger(c) || c < 0) return { error: 'Las consultas deben ser un número entero de 0 en adelante.' };
      g.Consultas = c;
    }
  }
  if (!g.NoFinca) return { error: 'Escribe el número de finca.' };
  if (!fechaValida_(g.Fecha)) return { error: 'La fecha no es válida.' };
  if (TIPOS_GESTION.indexOf(g.Tipo) === -1) return { error: 'Elige el tipo de gestión.' };
  if (!g.Descripcion) return { error: 'Escribe qué se hizo.' };
  if (g.FechaProximo && !fechaValida_(g.FechaProximo)) return { error: 'La fecha del próximo paso no es válida.' };
  if (g.HayInteresado) {
    g.Cliente = String(b.cliente || '').trim().slice(0, 200) || null;
    g.Contacto = String(b.contacto || '').trim().slice(0, 150) || null;
    g.FormaPago = String(b.formaPago || '').trim().slice(0, 200) || null;
    g.EstadoOferta = ESTADOS_OFERTA.indexOf(b.estadoOferta) !== -1 ? b.estadoOferta : 'EN_EVALUACION';
    if (!g.Cliente) return { error: 'Escribe el nombre del cliente interesado.' };
    if (b.montoOferta !== '' && b.montoOferta !== null && b.montoOferta !== undefined) {
      g.MontoOferta = numero(b.montoOferta);
      if (g.MontoOferta === null || g.MontoOferta < 0) return { error: 'El monto ofertado no es un número válido.' };
    }
  }
  return { g: g };
}
function filtrarGestiones_(q, conOfertas) {
  var f = q.finca ? String(q.finca).trim() : null;
  return ord_(T_('BR_Gestion').filter(function (g) {
    return !g.Anulada && (!f || g.NoFinca === f) && (!fechaValida_(q.desde) || iso_(g.Fecha) >= q.desde) &&
      (!fechaValida_(q.hasta) || iso_(g.Fecha) <= q.hasta) && (!conOfertas || q.soloOfertas !== '1' || g.HayInteresado);
  }), 'Fecha DESC, Id DESC');
}
ruta_('GET', '/api/bienes/gestion', 'cualquiera', function (ctx) {
  return { success: true, gestiones: filtrarGestiones_(ctx.query, true).slice(0, 2000), tipos: TIPOS_GESTION, canales: CANALES };
});
ruta_('POST', '/api/bienes/gestion', 'gestor', function (ctx) {
  var r = validarGestion_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = ins_('BR_Gestion', Object.assign({}, r.g, { UsuarioId: ctx.br.Id, Usuario: ctx.br.Nombre }));
  bitacora_(ctx, 'GESTION', 'Finca ' + r.g.NoFinca + ': ' + r.g.Tipo + (r.g.MontoOferta !== null ? ' · oferta ' + r.g.MontoOferta : ''));
  return { success: true, id: f.Id };
});
ruta_('PUT', '/api/bienes/gestion/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10); if (!id) throw err_('Gestión inválida.', 400);
  var r = validarGestion_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = porId_('BR_Gestion', id); if (!f || f.Anulada) throw err_('La gestión no existe o está anulada.', 404);
  upd_('BR_Gestion', f, Object.assign({}, r.g, { UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() }));
  bitacora_(ctx, 'GESTION_EDITAR', 'Gestión #' + id + ' finca ' + r.g.NoFinca);
  return { success: true };
});
ruta_('POST', '/api/bienes/gestion/:id/anular', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), f = porId_('BR_Gestion', id);
  if (!f || f.Anulada) throw err_('La gestión no existe o ya estaba anulada.', 404);
  upd_('BR_Gestion', f, { Anulada: true, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  bitacora_(ctx, 'GESTION_ANULAR', 'Gestión #' + id);
  return { success: true };
});
ruta_('GET', '/api/bienes/gestion/exportar', 'cualquiera', function (ctx) {
  var filas = filtrarGestiones_(ctx.query, false);
  bitacora_(ctx, 'DESCARGAR', 'Gestión de campo (' + filas.length + ' filas)');
  return { success: true, __excel: 'gestion', nombre: 'Gestion_de_campo_' + hoyISO_() + '.xlsx', filas: filas };
});

// =============================================================================
// PLANTILLAS Y DOCUMENTOS
// =============================================================================
var TIPOS_DOC = ['MEMO', 'PROFORMA', 'VENTA', 'ALQUILER', 'MINUTA', 'CAJA', 'OTRO'];
var MAX_HTML = 3 * 1024 * 1024, MAX_ADJUNTO = 15 * 1024 * 1024;
function leerArchivoSubido_(b) {
  var nombre = String(b.nombreArchivo || '').trim().slice(0, 260), ext = extension_(nombre);
  if (!b.archivoB64 || !nombre) return { error: 'Falta el archivo.' };
  if (!TIPOS_MIME[ext] || ext === 'csv') return { error: 'Tipo de archivo no permitido. Usa PDF, Word, Excel, JPG o PNG.' };
  var largo = Math.floor(String(b.archivoB64).replace(/^data:.*;base64,/, '').length * 3 / 4);
  if (!largo) return { error: 'El archivo llegó vacío.' };
  if (largo > MAX_ADJUNTO) return { error: 'El archivo pasa de 15 MB.' };
  return { nombre: nombre };
}
// Plantillas de fábrica (Plantillas.gs): solo se agregan si su código no existe todavía
function sembrarPlantillas_() {
  if (typeof PLANTILLAS_BASE === 'undefined') return 0;
  var n = 0, ya = {};
  T_('BR_Plantilla').forEach(function (p) { if (p.Codigo) ya[p.Codigo] = 1; });
  PLANTILLAS_BASE.forEach(function (p) {
    if (ya[p.Codigo]) return;
    ins_('BR_Plantilla', { Codigo: p.Codigo, Nombre: p.Nombre, Tipo: p.Tipo, Contenido: p.Contenido, UsuarioCrea: 'SISTEMA' }); n++;
  });
  return n;
}
ruta_('GET', '/api/bienes/plantillas', 'cualquiera', function () {
  if (!T_('BR_Plantilla').length) sembrarPlantillas_();
  var l = ord_(T_('BR_Plantilla'), 'Activa DESC, Tipo, Nombre').map(function (p) {
    return copia_(p, ['Id', 'Codigo', 'Nombre', 'Tipo', 'Contenido', 'Activa', 'ArchivoNombre', 'UsuarioCrea', 'FechaCrea', 'UsuarioModifica', 'FechaModifica']); });
  return { success: true, plantillas: l };
});
function validarPlantilla_(b) {
  var p = { Nombre: String(b.nombre || '').trim().slice(0, 150), Tipo: String(b.tipo || '').toUpperCase(), Contenido: limpiarHtml(b.contenido) };
  if (!p.Nombre) return { error: 'Escribe el nombre de la plantilla.' };
  if (TIPOS_DOC.indexOf(p.Tipo) === -1) return { error: 'Elige el tipo de plantilla.' };
  if (!p.Contenido.replace(/<[^>]+>/g, '').trim()) return { error: 'La plantilla está vacía.' };
  if (p.Contenido.length > MAX_HTML) return { error: 'La plantilla es demasiado grande (más de 3 MB). Quita imágenes pegadas.' };
  return { p: p };
}
ruta_('POST', '/api/bienes/plantillas', 'gestor', function (ctx) {
  var r = validarPlantilla_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = ins_('BR_Plantilla', Object.assign({}, r.p, { UsuarioCrea: ctx.br.Nombre }));
  bitacora_(ctx, 'PLANTILLA_CREAR', r.p.Nombre);
  return { success: true, id: f.Id };
});
ruta_('PUT', '/api/bienes/plantillas/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), b = ctx.body || {}, f = porId_('BR_Plantilla', id);
  if (b.activa !== undefined && b.nombre === undefined) {
    if (!f) throw err_('No existe esa plantilla.', 404);
    upd_('BR_Plantilla', f, { Activa: !!b.activa, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
    bitacora_(ctx, 'PLANTILLA_EDITAR', 'Plantilla #' + id + (b.activa ? ' activada' : ' desactivada'));
    return { success: true };
  }
  var r = validarPlantilla_(b); if (r.error) throw err_(r.error, 400);
  if (!f) throw err_('No existe esa plantilla.', 404);
  upd_('BR_Plantilla', f, Object.assign({}, r.p, { UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() }));
  bitacora_(ctx, 'PLANTILLA_EDITAR', r.p.Nombre);
  return { success: true };
});
ruta_('POST', '/api/bienes/plantillas/:id/archivo', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), r = leerArchivoSubido_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = porId_('BR_Plantilla', id); if (!f) throw err_('No existe esa plantilla.', 404);
  var a = guardarArchivo_(r.nombre, ctx.body.archivoB64), viejo = f.ArchivoId;
  upd_('BR_Plantilla', f, { ArchivoNombre: r.nombre, ArchivoId: a.id, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  borrarArchivo_(viejo);
  bitacora_(ctx, 'PLANTILLA_ARCHIVO', 'Plantilla #' + id + ': ' + r.nombre);
  return { success: true };
});
ruta_('GET', '/api/bienes/plantillas/:id/archivo', 'cualquiera', function (ctx) {
  var f = porId_('BR_Plantilla', parseInt(ctx.params.id, 10));
  if (!f || !f.ArchivoId) throw err_('Esa plantilla no tiene modelo original guardado.', 404);
  return archivoRespuesta_(f.ArchivoNombre, leerArchivo_(f.ArchivoId));
});

ruta_('GET', '/api/bienes/documentos', 'cualquiera', function (ctx) {
  var fq = ctx.query.finca ? String(ctx.query.finca).trim() : null, pl = {};
  T_('BR_Plantilla').forEach(function (p) { pl[p.Id] = p.Nombre; });
  var l = T_('BR_Documento').filter(function (d) { return !d.Anulado && (!fq || d.NoFinca === fq); })
    .sort(function (a, b) { return (b.FechaModifica || b.FechaCrea) - (a.FechaModifica || a.FechaCrea); }).slice(0, 2000)
    .map(function (d) { return Object.assign(copia_(d, ['Id', 'PlantillaId', 'Tipo', 'Titulo', 'NoFinca', 'Cliente', 'Estado', 'Version', 'AdjuntoNombre',
      'UsuarioCrea', 'FechaCrea', 'UsuarioModifica', 'FechaModifica']), { Plantilla: pl[d.PlantillaId] || null }); });
  return { success: true, documentos: l };
});
ruta_('GET', '/api/bienes/documentos/:id', 'cualquiera', function (ctx) {
  var d = porId_('BR_Documento', parseInt(ctx.params.id, 10)); if (!d) throw err_('No existe ese documento.', 404);
  return { success: true, documento: copia_(d, ['Id', 'PlantillaId', 'Tipo', 'Titulo', 'NoFinca', 'Cliente', 'Contenido', 'Campos', 'Estado', 'Version',
    'AdjuntoNombre', 'UsuarioCrea', 'FechaCrea', 'UsuarioModifica', 'FechaModifica', 'Anulado']) };
});
function validarDocumento_(b) {
  var d = { Titulo: String(b.titulo || '').trim().slice(0, 200), Tipo: String(b.tipo || '').toUpperCase(),
    NoFinca: String(b.noFinca || '').trim().slice(0, 30) || null, Cliente: String(b.cliente || '').trim().slice(0, 200) || null,
    Contenido: limpiarHtml(b.contenido), Estado: b.estado === 'FINAL' ? 'FINAL' : 'BORRADOR',
    Campos: b.campos ? JSON.stringify(b.campos).slice(0, 200000) : null, PlantillaId: parseInt(b.plantillaId, 10) || null };
  if (!d.Titulo) return { error: 'Escribe un título para el documento.' };
  if (TIPOS_DOC.indexOf(d.Tipo) === -1) return { error: 'Tipo de documento inválido.' };
  if (!d.Contenido.replace(/<[^>]+>/g, '').trim()) return { error: 'El documento está vacío.' };
  if (d.Contenido.length > MAX_HTML) return { error: 'El documento es demasiado grande (más de 3 MB). Quita imágenes pegadas.' };
  return { d: d };
}
ruta_('POST', '/api/bienes/documentos', 'gestor', function (ctx) {
  var r = validarDocumento_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = ins_('BR_Documento', Object.assign({}, r.d, { UsuarioCrea: ctx.br.Nombre }));
  bitacora_(ctx, 'DOC_CREAR', r.d.Titulo);
  return { success: true, id: f.Id, version: f.Version };
});
// Guardar cambios. Si otra persona lo guardó mientras tanto, avisa en vez de pisarlo.
ruta_('PUT', '/api/bienes/documentos/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), b = ctx.body || {};
  var r = validarDocumento_(b); if (r.error) throw err_(r.error, 400);
  var f = porId_('BR_Documento', id);
  if (!f || f.Anulado) throw err_('El documento no existe o fue anulado.', 404);
  if (f.Version !== (parseInt(b.version, 10) || 0))
    throw err_('Otra persona (' + (f.UsuarioModifica || '¿?') + ') guardó este documento mientras lo editabas. Copia tus cambios, vuelve a abrirlo y pégalos.', 409);
  var c = Object.assign({}, r.d, { Version: f.Version + 1, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  delete c.PlantillaId; if (c.Campos === null) delete c.Campos;
  upd_('BR_Documento', f, c);
  bitacora_(ctx, 'DOC_EDITAR', r.d.Titulo);
  return { success: true, version: f.Version };
});
ruta_('POST', '/api/bienes/documentos/:id/anular', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), f = porId_('BR_Documento', id);
  if (!f || f.Anulado) throw err_('El documento no existe o ya estaba anulado.', 404);
  upd_('BR_Documento', f, { Anulado: true, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  bitacora_(ctx, 'DOC_ANULAR', 'Documento #' + id);
  return { success: true };
});
ruta_('POST', '/api/bienes/documentos/:id/adjunto', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), r = leerArchivoSubido_(ctx.body || {}); if (r.error) throw err_(r.error, 400);
  var f = porId_('BR_Documento', id); if (!f || f.Anulado) throw err_('El documento no existe o está anulado.', 404);
  var a = guardarArchivo_(r.nombre, ctx.body.archivoB64), viejo = f.AdjuntoId;
  upd_('BR_Documento', f, { AdjuntoNombre: r.nombre, AdjuntoId: a.id, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  borrarArchivo_(viejo);
  bitacora_(ctx, 'DOC_ADJUNTO', 'Documento #' + id + ': ' + r.nombre);
  return { success: true };
});
ruta_('GET', '/api/bienes/documentos/:id/adjunto', 'cualquiera', function (ctx) {
  var f = porId_('BR_Documento', parseInt(ctx.params.id, 10));
  if (!f || !f.AdjuntoId) throw err_('Ese documento no tiene archivo adjunto.', 404);
  return archivoRespuesta_(f.AdjuntoNombre, leerArchivo_(f.AdjuntoId));
});

// =============================================================================
// AVALÚOS (nuevo)
// Lista de PDF de avalúos (enlaces de SharePoint/OneDrive o PDF subidos aquí).
// Cada avalúo se une solo a su finca: por el número en el nombre del archivo
// ("FINCA 56575", "F339402", "30357148-4B12", "placa 936979") y, si no trae
// número, por el nombre del deudor. Se puede asignar la finca a mano.
// =============================================================================
function sinAcentos_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function fincasDeNombre_(nombre) {
  var s = sinAcentos_(nombre).replace(/\.pdf$/i, ''), out = [], m;
  var add = function (x) { x = String(x).toUpperCase(); if (out.indexOf(x) === -1) out.push(x); };
  var re1 = /\b(?:fincas?|placa|f)\s*(?:n[o°º.]*\s*)?#?\s*(\d{3,})/gi;
  while ((m = re1.exec(s))) add(m[1]);
  var re2 = /(?:^|[^\dA-Za-z])(\d{5,})(?![\d])/g;
  while ((m = re2.exec(s))) add(m[1]);
  return out;
}
var PALABRAS_FUERA = ['avaluo', 'avaluos', 'finca', 'fincas', 'veh', 'vehiculo', 'auto', 'placa', 'camion', 'pdf', 'del', 'las', 'los', 'por', 'y'];
function palabras_(t) {
  return sinAcentos_(t).toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/)
    .filter(function (w) { return w.length > 2 && !/^\d+$/.test(w) && PALABRAS_FUERA.indexOf(w) === -1; });
}
function titularDeNombre_(nombre) {
  var t = String(nombre || '').replace(/\.pdf$/i, '')
    .replace(/\baval[uú]os?\b/gi, ' ').replace(/\bfincas?\s*\d+(?:\s*-\s*\d[0-9a-z]*)?(?:\s*y\s*\d+)?/gi, ' ')
    .replace(/\b(?:placa|f)\s*\d+/gi, ' ').replace(/\b\d{3,}(?:\s*-\s*\d[0-9a-z]*)?\b/gi, ' ')
    .replace(/\b(?:auto|veh|veh[ií]culo|cami[oó]n)\b/gi, ' ').replace(/\(\d+\)/g, ' ').replace(/\s\d{1,2}\s*$/, ' ')
    .replace(/\s+-\s+|^\s*-\s*|\s*-\s*$/g, ' ').replace(/\s{2,}/g, ' ').trim();
  return t.replace(/^[,.\s]+|[,\s]+$/g, '') || null;
}
// Fincas que conoce el sistema, con el deudor para unir por nombre
function universoFincas_() {
  var u = {};
  var poner = function (nf, d) {
    if (!nf) return; var x = u[nf] = u[nf] || { NoFinca: nf, nombres: [], Estado: null, Activa: null, Ubicacion: null, Vendida: null };
    if (d.Deudor || d.AntiguoPropietario) x.nombres.push(d.Deudor || d.AntiguoPropietario);
    Object.keys(d).forEach(function (k) { if (x[k] == null && d[k] != null && k !== 'Deudor') x[k] = d[k]; });
  };
  T_('BR_Finca').forEach(function (f) { poner(f.NoFinca, { Deudor: f.Deudor, Estado: f.Estado, Activa: f.Activa, Ubicacion: f.Ubicacion, CodigoUbic: f.CodigoUbic }); });
  T_('BR_VentaAnterior').forEach(function (v) { poner(v.NoFinca, { Deudor: v.Deudor, Ubicacion: v.Ubicacion, Vendida: true, Anio: v.Anio }); });
  T_('BR_Mejora').forEach(function (m) { poner(m.NoFinca, { AntiguoPropietario: m.AntiguoPropietario }); });
  T_('BR_Venta').forEach(function (v) { poner(v.NoFinca, { AntiguoPropietario: v.AntiguoPropietario, Ubicacion: v.Ubicacion }); });
  Object.keys(u).forEach(function (k) { var w = {}; u[k].nombres.forEach(function (n) { palabras_(n).forEach(function (p) { w[p] = 1; }); }); u[k].pal = w; });
  return u;
}
function unirAvaluo_(a, U) {
  var r = { fincas: [], forma: null, numeros: [] };
  if (a.FincasManual) {
    r.forma = 'manual';
    r.numeros = String(a.FincasManual).split(/[,;\s]+/).filter(Boolean);
  } else {
    r.numeros = fincasDeNombre_(a.Archivo);
    if (r.numeros.length) r.forma = 'número';
  }
  if (r.numeros.length) {
    r.fincas = r.numeros.map(function (n) { var x = U[n]; return x ? x : { NoFinca: n, noEsta: true }; });
    return r;
  }
  // Por nombre del deudor: las fincas con más palabras en común (mínimo 2)
  var pa = palabras_(a.Titular || titularDeNombre_(a.Archivo)), mejor = 0, cand = [];
  if (pa.length < 2) return r;
  Object.keys(U).forEach(function (k) {
    var p = pa.filter(function (w) { return U[k].pal[w]; }).length;
    if (p > mejor) { mejor = p; cand = [U[k]]; } else if (p === mejor && p) cand.push(U[k]);
  });
  if (mejor >= 2) { r.fincas = cand.slice(0, 6); r.forma = 'nombre'; }
  return r;
}
function avaluoPublico_(a, U) {
  var r = unirAvaluo_(a, U);
  return { Id: a.Id, Archivo: a.Archivo, Enlace: a.Enlace, Titular: a.Titular, FincasManual: a.FincasManual, FechaAvaluo: a.FechaAvaluo, Nota: a.Nota,
    Origen: a.Origen, TieneArchivo: !!a.ArchivoId, UsuarioCrea: a.UsuarioCrea, FechaCrea: a.FechaCrea, UsuarioModifica: a.UsuarioModifica, FechaModifica: a.FechaModifica,
    Forma: r.forma, Fincas: r.fincas.map(function (f) { return { NoFinca: f.NoFinca, CodigoUbic: f.CodigoUbic || null, Estado: f.Estado || null,
      Activa: f.Activa, Ubicacion: f.Ubicacion || null, Deudor: (f.nombres || [])[0] || null, Vendida: !!f.Vendida, NoEsta: !!f.noEsta }; }) };
}
function avaluosPorFinca_() {
  var U = universoFincas_(), por = {};
  T_('BR_Avaluo').forEach(function (a) {
    if (a.Anulado) return;
    var r = unirAvaluo_(a, U);
    r.fincas.forEach(function (f) { (por[f.NoFinca] = por[f.NoFinca] || []).push({ Id: a.Id, Archivo: a.Archivo, Enlace: a.Enlace, TieneArchivo: !!a.ArchivoId, Forma: r.forma }); });
  });
  return por;
}
ruta_('GET', '/api/bienes/avaluos', 'cualquiera', function () {
  var U = universoFincas_();
  var l = ord_(T_('BR_Avaluo').filter(function (a) { return !a.Anulado; }), 'Archivo').map(function (a) { return avaluoPublico_(a, U); });
  var con = {}; l.forEach(function (a) { a.Fincas.forEach(function (f) { if (!f.NoEsta) con[f.NoFinca] = 1; }); });
  var activas = T_('BR_Finca').filter(function (f) { return f.Activa; });
  var sinAv = activas.filter(function (f) { return !con[f.NoFinca]; }).map(function (f) {
    return { NoFinca: f.NoFinca, CodigoUbic: f.CodigoUbic, Deudor: f.Deudor, Ubicacion: f.Ubicacion, Estado: f.Estado, ValorMercado: f.ValorMercado, Observaciones: f.Observaciones }; });
  return { success: true, avaluos: l, fincasSinAvaluo: sinAv, totalActivas: activas.length,
    resumen: { total: l.length, porNumero: l.filter(function (a) { return a.Forma === 'número'; }).length, porNombre: l.filter(function (a) { return a.Forma === 'nombre'; }).length,
      manual: l.filter(function (a) { return a.Forma === 'manual'; }).length, sinFinca: l.filter(function (a) { return !a.Fincas.length; }).length,
      activasConAvaluo: activas.length - sinAv.length } };
});
// Listado (CSV o Excel con columnas Archivo y Link): agrega lo nuevo y conserva lo que ya estaba
function importarAvaluos_(filas, origen, usuario, soloVista) {
  var ya = {};
  T_('BR_Avaluo').forEach(function (a) { if (!a.Anulado) { if (a.Enlace) ya['e|' + a.Enlace.trim().toLowerCase()] = 1; ya['n|' + String(a.Archivo).trim().toLowerCase()] = 1; } });
  var nuevos = [], repetidos = 0;
  (filas || []).forEach(function (f) {
    var nombre = String(f.Archivo || f.archivo || '').trim(), enlace = String(f.Enlace || f.Link || f.enlace || f.link || '').trim();
    if (!nombre && !enlace) return;
    if (!nombre) nombre = decodeURIComponent(enlace.split('/').pop() || 'Avalúo');
    if (enlace && !/^https?:\/\//i.test(enlace)) enlace = '';
    if ((enlace && ya['e|' + enlace.toLowerCase()]) || ya['n|' + nombre.toLowerCase()]) { repetidos++; return; }
    ya['e|' + enlace.toLowerCase()] = 1; ya['n|' + nombre.toLowerCase()] = 1;
    nuevos.push({ Archivo: nombre.slice(0, 260), Enlace: enlace.slice(0, 1000) || null, Titular: titularDeNombre_(nombre), Origen: String(origen || 'Listado').slice(0, 200), UsuarioCrea: usuario });
  });
  if (!soloVista) nuevos.forEach(function (a) { ins_('BR_Avaluo', a); });
  var U = universoFincas_();
  return { nuevos: nuevos.length, repetidos: repetidos, filas: nuevos.slice(0, 300).map(function (a) { return avaluoPublico_(a, U); }) };
}
ruta_('POST', '/api/bienes/avaluos/importar', 'gestor', function (ctx) {
  var b = ctx.body || {};
  if (!Array.isArray(b.filas) || !b.filas.length) throw err_('El listado no trae filas. Debe tener las columnas "Archivo" y "Link".', 400);
  var r = importarAvaluos_(b.filas, 'Listado: ' + String(b.nombreArchivo || '').slice(0, 150), ctx.br.Nombre, !!b.soloVista);
  if (!b.soloVista) bitacora_(ctx, 'AVALUOS', 'Subió el listado ' + (b.nombreArchivo || '') + ': ' + r.nuevos + ' nuevos, ' + r.repetidos + ' ya estaban');
  return Object.assign({ success: true, vista: !!b.soloVista }, r);
});
function validarAvaluo_(b, parcial) {
  var a = { Archivo: String(b.archivo || '').trim().slice(0, 260), Enlace: String(b.enlace || '').trim().slice(0, 1000) || null,
    FincasManual: String(b.fincas || '').trim().slice(0, 200) || null, FechaAvaluo: b.fechaAvaluo && fechaValida_(b.fechaAvaluo) ? b.fechaAvaluo : null,
    Nota: String(b.nota || '').trim().slice(0, 1000) || null, Titular: String(b.titular || '').trim().slice(0, 200) || null };
  if (a.Enlace && !/^https?:\/\//i.test(a.Enlace)) throw err_('El enlace debe empezar con http:// o https://', 400);
  if (!parcial && !a.Archivo && !b.nombreArchivo) throw err_('Escribe el nombre del avalúo o sube el PDF.', 400);
  return a;
}
ruta_('POST', '/api/bienes/avaluos', 'gestor', function (ctx) {
  var b = ctx.body || {}, a = validarAvaluo_(b);
  if (b.archivoB64) {
    var r = leerArchivoSubido_(b); if (r.error) throw err_(r.error, 400);
    a.ArchivoId = guardarArchivo_(r.nombre, b.archivoB64).id;
    if (!a.Archivo) a.Archivo = r.nombre;
  }
  if (!a.Enlace && !a.ArchivoId) throw err_('Pon el enlace del PDF o sube el archivo.', 400);
  a.Titular = a.Titular || titularDeNombre_(a.Archivo);
  var f = ins_('BR_Avaluo', Object.assign(a, { Origen: b.archivoB64 ? 'PDF subido' : 'Agregado a mano', UsuarioCrea: ctx.br.Nombre }));
  bitacora_(ctx, 'AVALUOS', 'Agregó el avalúo ' + f.Archivo);
  return { success: true, id: f.Id };
});
ruta_('PUT', '/api/bienes/avaluos/:id', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), f = porId_('BR_Avaluo', id), b = ctx.body || {};
  if (!f || f.Anulado) throw err_('Ese avalúo no existe o se quitó.', 404);
  var a = validarAvaluo_(b, true), c = { UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() };
  ['fincas', 'nota', 'fechaAvaluo', 'enlace', 'titular', 'archivo'].forEach(function (k) {
    var K = { fincas: 'FincasManual', nota: 'Nota', fechaAvaluo: 'FechaAvaluo', enlace: 'Enlace', titular: 'Titular', archivo: 'Archivo' }[k];
    if (b[k] !== undefined) c[K] = a[K];
  });
  if (c.Archivo === '') delete c.Archivo;
  upd_('BR_Avaluo', f, c);
  bitacora_(ctx, 'AVALUOS', 'Editó el avalúo #' + id + (c.FincasManual !== undefined ? ' · fincas: ' + (c.FincasManual || 'automático') : ''));
  return { success: true };
});
ruta_('POST', '/api/bienes/avaluos/:id/anular', 'gestor', function (ctx) {
  var id = parseInt(ctx.params.id, 10), f = porId_('BR_Avaluo', id);
  if (!f || f.Anulado) throw err_('Ese avalúo no existe o ya se quitó.', 404);
  upd_('BR_Avaluo', f, { Anulado: true, UsuarioModifica: ctx.br.Nombre, FechaModifica: ahora_() });
  bitacora_(ctx, 'AVALUOS', 'Quitó el avalúo #' + id + ': ' + f.Archivo);
  return { success: true };
});
ruta_('GET', '/api/bienes/avaluos/:id/archivo', 'cualquiera', function (ctx) {
  var f = porId_('BR_Avaluo', parseInt(ctx.params.id, 10));
  if (!f || !f.ArchivoId) throw err_('Ese avalúo no tiene PDF guardado aquí.', 404);
  return archivoRespuesta_(/\.pdf$/i.test(f.Archivo) ? f.Archivo : f.Archivo + '.pdf', leerArchivo_(f.ArchivoId));
});

// =============================================================================
// INSTALACIÓN (menú de la hoja)
// =============================================================================
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Bienes reposeídos')
      .addItem('Instalar', 'instalar')
      .addItem('Ver la clave de instalación', 'verClaveInstalacion')
      .addItem('Nueva clave de instalación', 'nuevaClaveInstalacion')
      .addToUi();
  } catch (e) { /* sin interfaz */ }
}
function mostrar_(titulo, msg) { try { SpreadsheetApp.getUi().alert(titulo, msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* desde el editor */ } Logger.log(titulo + '\n' + msg); }
function nuevaClave_() {
  var L = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789', hx = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''), s = '';
  for (var i = 0; i < 16; i++) { s += L[parseInt(hx.substr(i * 2, 2), 16) % L.length]; if (i === 3 || i === 7 || i === 11) s += '-'; }
  PropertiesService.getScriptProperties().setProperty('CLAVE_INSTALACION', s);
  return s;
}
function instalar() {
  Object.keys(ESQ).forEach(function (t) { T_(t); });
  var carpeta = carpeta_();
  var nP = sembrarPlantillas_();
  var nA = 0;
  if (typeof AVALUOS_INICIALES !== 'undefined' && !T_('BR_Avaluo').length) {
    nA = importarAvaluos_(AVALUOS_INICIALES.map(function (x) { return { Archivo: x[0], Enlace: x[1] }; }), 'Listado inicial (AVALUOS de OneDrive)', 'Instalación', false).nuevos;
  }
  guardarTodo_();
  var clave = claveInstalacion_() || nuevaClave_();
  secreto_();
  var hoja = libro_().getSheetByName('Hoja 1') || libro_().getSheetByName('Sheet1');
  if (hoja && hoja.getLastRow() === 0 && libro_().getSheets().length > 1) { try { libro_().deleteSheet(hoja); } catch (e) { /* nada */ } }
  mostrar_('Bienes reposeídos instalado',
    'Pestañas listas, carpeta de archivos en Drive: "' + carpeta.getName() + '".\n' +
    (nP ? nP + ' plantillas de documentos cargadas.\n' : '') + (nA ? nA + ' avalúos cargados.\n' : '') +
    (hayAdmin_() ? '\nYa hay administrador: entra con tu usuario.' : '\nCLAVE DE INSTALACIÓN: ' + clave +
      '\n\nLa pide la página la primera vez, para crear el administrador.') +
    '\n\nFalta publicar: Implementar → Nueva implementación → Aplicación web (Ejecutar como: Yo · Acceso: Cualquier usuario).');
}
function verClaveInstalacion() {
  var c = claveInstalacion_() || nuevaClave_();
  mostrar_('Clave de instalación', c + '\n\nSolo sirve para crear el primer administrador (o uno nuevo si se desactivaron todos).');
}
function nuevaClaveInstalacion() { mostrar_('Nueva clave de instalación', nuevaClave_()); }
