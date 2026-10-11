/**
 * TPT — Indumentaria · Apps Script (backend de la planilla)
 * ------------------------------------------------------------
 * URL deployment activo:
 * https://script.google.com/macros/s/AKfycbwREGjI3V2_vjnpSNbe5SbZIw5m2mkn7d1fmp0ZKsu4cGCT4NKpG5IirzkH4TA-A0_2/exec
 *
 * Acciones disponibles:
 *   GET  ?action=getAll&clave=XXX          → pedidos + estado de gestión (JSON)
 *   GET  ?action=getPrendas&clave=XXX      → catálogo de la hoja Prendas
 *   GET  ?action=getAtletas&clave=XXX      → lista de atletas de la hoja Atletas
 *   POST {action:'setGestion', clave, key, data}  → guarda estado de un pedido
 *   POST {action:'delGestion', clave, key}         → borra estado de un pedido
 *   POST {action:'setPrenda',  clave, data}         → crea o actualiza una prenda
 *   POST {action:'delPrenda',  clave, key}          → borra una prenda
 *   POST {action:'setAtleta',  clave, data}         → crea o actualiza un atleta
 *   POST {action:'delAtleta',  clave, key}          → borra un atleta (por ID)
 *
 * Instalación:
 *   1. Abrí el proyecto de Apps Script de la planilla de pedidos.
 *   2. Pegá este archivo completo en Code.gs.
 *   3. Completá SHEET_ID y CLAVE abajo.
 *   4. Implementar → Administrar implementaciones → lápiz → Nueva versión → Implementar.
 */

var SHEET_ID     = '1ofI_n6kdGZDrMsNcEJbwNL9Mx0Rp-uMDimcoIGtkV6s';
var CLAVE        = 'n3w3r4.2026';
var HOJA_PEDIDOS = '';          // vacío = primera hoja
var HOJA_GESTION = 'Gestion';  // se crea sola si no existe
var HOJA_PRENDAS = 'Prendas';  // debe existir
var HOJA_ATLETAS = 'Atletas';  // se crea sola si no existe

var COLS_PEDIDO  = ['fecha','nombre','producto','talle','cantidad','precio','subtotal','total'];
var COLS_GESTION = ['key','alumno_id','alumno','telefono','pagado','medio','fecha_pago','entrega','notas','actualizado'];
var COLS_PRENDAS = ['proveedor','prenda','costo','precio de venta','talles'];
var COLS_ATLETAS = ['id','apellido','nombre','telefono','profe','activo'];

// ─────────────────────────────── ENTRADAS ───────────────────────────────
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action) return responder(ejecutar(p));
  return responder(agregarPedido(p));
}

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents || '{}'); } catch (err) { p = (e && e.parameter) || {}; }
  if (p.action) return responder(ejecutar(p));
  return responder(agregarPedido(p));
}

function ejecutar(p) {
  if (String(p.clave || '') !== CLAVE) return { ok: false, error: 'No autorizado' };
  try {
    switch (p.action) {
      case 'getAll':      return { ok: true, data: { pedidos: leerPedidos(), gestion: leerGestion() } };
      case 'getPedidos':  return { ok: true, data: leerPedidos() };
      case 'getGestion':  return { ok: true, data: leerGestion() };
      case 'setGestion':  return { ok: true, data: guardarGestion(p.key, p.data || {}) };
      case 'delGestion':  return { ok: true, data: borrarGestion(p.key) };
      case 'getPrendas':  return { ok: true, data: leerPrendas() };
      case 'setPrenda':   return { ok: true, data: guardarPrenda(p.data || {}) };
      case 'delPrenda':   return { ok: true, data: borrarPrenda(p.key) };
      case 'getAtletas':  return { ok: true, data: leerAtletas() };
      case 'setAtleta':   return { ok: true, data: guardarAtleta(p.data || {}) };
      case 'delAtleta':   return { ok: true, data: borrarAtleta(p.key) };
      default:            return { ok: false, error: 'Acción desconocida: ' + p.action };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────── PLANILLA ───────────────────────────────
function ss() { return SpreadsheetApp.openById(SHEET_ID); }

function hojaPedidos() {
  var s = ss();
  var h = HOJA_PEDIDOS ? s.getSheetByName(HOJA_PEDIDOS) : s.getSheets()[0];
  if (!h) throw new Error('No existe la hoja de pedidos "' + HOJA_PEDIDOS + '"');
  if (h.getLastRow() === 0) h.appendRow(COLS_PEDIDO);
  return h;
}

function hojaGestion() {
  var s = ss();
  var h = s.getSheetByName(HOJA_GESTION);
  if (!h) { h = s.insertSheet(HOJA_GESTION); h.appendRow(COLS_GESTION); h.setFrozenRows(1); }
  if (h.getLastRow() === 0) h.appendRow(COLS_GESTION);
  return h;
}

function hojaPrendas() {
  var h = ss().getSheetByName(HOJA_PRENDAS);
  if (!h) throw new Error('No existe la hoja "' + HOJA_PRENDAS + '" en la planilla');
  return h;
}

function hojaAtletas() {
  var s = ss();
  var h = s.getSheetByName(HOJA_ATLETAS);
  if (!h) {
    h = s.insertSheet(HOJA_ATLETAS);
    h.appendRow(['ID', 'Apellido', 'Nombre', 'Telefono', 'Profe', 'Activo']);
    h.setFrozenRows(1);
  }
  return h;
}

function agregarPedido(p) {
  if (!p.producto && !p.nombre) return { ok: false, error: 'Sin datos' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hojaPedidos();
    var hdr = encabezados(h);
    var fila = hdr.map(function (k) {
      var v = p[k];
      if (v == null) return '';
      if (['cantidad','precio','subtotal','total'].indexOf(k) >= 0) { var n = Number(v); return isNaN(n) ? v : n; }
      return v;
    });
    h.appendRow(fila);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function encabezados(h) {
  var lastCol = Math.max(h.getLastColumn(), COLS_PEDIDO.length);
  var raw = h.getRange(1, 1, 1, lastCol).getValues()[0].map(norm);
  var reconocidos = raw.filter(function (c) { return COLS_PEDIDO.indexOf(c) >= 0; }).length;
  if (reconocidos < 3) return COLS_PEDIDO.concat(raw.slice(COLS_PEDIDO.length));
  return raw;
}

function norm(s) {
  return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

// ─────────────────────────────── PEDIDOS ───────────────────────────────
function leerPedidos() {
  var h = hojaPedidos();
  var n = h.getLastRow();
  if (n < 2) return [];
  var hdr = encabezados(h);
  var vals = h.getRange(2, 1, n - 1, hdr.length).getValues();
  var out = [];
  vals.forEach(function (r, i) {
    var o = { fila: i + 2 };
    var vacia = true;
    hdr.forEach(function (k, j) {
      if (!k) return;
      var v = r[j];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      if (v !== '' && v != null) vacia = false;
      o[k] = v;
    });
    if (!vacia) out.push(o);
  });
  return out;
}

// ─────────────────────────────── GESTION ───────────────────────────────
function leerGestion() {
  var h = hojaGestion();
  var n = h.getLastRow();
  if (n < 2) return {};
  var vals = h.getRange(2, 1, n - 1, COLS_GESTION.length).getValues();
  var out = {};
  vals.forEach(function (r) {
    if (!r[0]) return;
    var o = {};
    COLS_GESTION.forEach(function (k, j) {
      var v = r[j];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      o[k] = v;
    });
    out[String(r[0])] = o;
  });
  return out;
}

function guardarGestion(key, data) {
  if (!key) throw new Error('Falta key');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hojaGestion();
    var n = h.getLastRow();
    var keys = n >= 2 ? h.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    var idx = keys.indexOf(String(key));
    var actual = {};
    if (idx >= 0) {
      var r = h.getRange(idx + 2, 1, 1, COLS_GESTION.length).getValues()[0];
      COLS_GESTION.forEach(function (k, j) { actual[k] = r[j]; });
    }
    COLS_GESTION.forEach(function (k) { if (data[k] !== undefined) actual[k] = data[k]; });
    actual.key = key;
    actual.actualizado = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var fila = COLS_GESTION.map(function (k) { return actual[k] == null ? '' : actual[k]; });
    if (idx >= 0) h.getRange(idx + 2, 1, 1, COLS_GESTION.length).setValues([fila]);
    else h.appendRow(fila);
    return actual;
  } finally { lock.releaseLock(); }
}

function borrarGestion(key) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hojaGestion();
    var n = h.getLastRow();
    if (n < 2) return false;
    var keys = h.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); });
    var idx = keys.indexOf(String(key));
    if (idx < 0) return false;
    h.deleteRow(idx + 2);
    return true;
  } finally { lock.releaseLock(); }
}

// ─────────────────────────────── PRENDAS ───────────────────────────────
function clavePrenda(proveedor, prenda) {
  return norm(proveedor) + '|' + norm(prenda);
}

function numeroOVacio(v) {
  if (v === '' || v == null) return '';
  var n = Number(String(v).replace(/[$,]/g, '').replace(',', '.'));
  return isNaN(n) ? '' : n;
}

function leerPrendas() {
  var h = hojaPrendas();
  var n = h.getLastRow();
  if (n < 2) return [];
  var vals = h.getRange(2, 1, n - 1, COLS_PRENDAS.length).getValues();
  var out = [];
  vals.forEach(function (r, i) {
    if (String(r[0]).trim() === '' && String(r[1]).trim() === '') return;
    out.push({
      fila: i + 2,
      key: clavePrenda(r[0], r[1]),
      proveedor: String(r[0]),
      prenda: String(r[1]),
      costo: numeroOVacio(r[2]),
      precio: numeroOVacio(r[3]),
      talles: String(r[4] || '')
    });
  });
  return out;
}

function guardarPrenda(data) {
  if (!data.proveedor || !data.prenda) throw new Error('Falta proveedor o prenda');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hojaPrendas();
    var filas = h.getLastRow() >= 2 ? leerPrendas() : [];
    var buscada = data.key_anterior || clavePrenda(data.proveedor, data.prenda);
    var existente = filas.filter(function (f) { return f.key === buscada; })[0];
    var actual = existente
      ? { proveedor: existente.proveedor, prenda: existente.prenda, costo: existente.costo, precio: existente.precio, talles: existente.talles }
      : { proveedor: '', prenda: '', costo: '', precio: '', talles: 'XS a 3XL' };
    if (data.proveedor !== undefined) actual.proveedor = String(data.proveedor).trim();
    if (data.prenda    !== undefined) actual.prenda    = String(data.prenda).trim();
    if (data.costo     !== undefined) actual.costo     = numeroOVacio(data.costo);
    if (data.precio    !== undefined) actual.precio    = numeroOVacio(data.precio);
    if (data.talles    !== undefined) actual.talles    = String(data.talles).trim();
    var nueva = clavePrenda(actual.proveedor, actual.prenda);
    var duplicada = filas.filter(function (f) { return f.key === nueva && f.key !== buscada; })[0];
    if (duplicada) throw new Error('Ya existe "' + actual.prenda + '" para ' + actual.proveedor);
    var fila = [actual.proveedor, actual.prenda, actual.costo, actual.precio, actual.talles];
    if (existente) h.getRange(existente.fila, 1, 1, COLS_PRENDAS.length).setValues([fila]);
    else h.appendRow(fila);
    return { key: nueva, proveedor: actual.proveedor, prenda: actual.prenda, costo: actual.costo, precio: actual.precio, talles: actual.talles };
  } finally { lock.releaseLock(); }
}

function borrarPrenda(key) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hojaPrendas();
    var filas = leerPrendas();
    var f = filas.filter(function (x) { return x.key === String(key); })[0];
    if (!f) return false;
    h.deleteRow(f.fila);
    return true;
  } finally { lock.releaseLock(); }
}

// ─────────────────────────────── ATLETAS ───────────────────────────────
function leerAtletas() {
  var h = hojaAtletas();
  var n = h.getLastRow();
  if (n < 2) return [];
  var vals = h.getRange(2, 1, n - 1, COLS_ATLETAS.length).getValues();
  var out = [];
  vals.forEach(function (r, i) {
    if (String(r[0]).trim() === '' && String(r[2]).trim() === '') return;
    out.push({
      fila: i + 2,
      id:        String(r[0]),
      apellido:  String(r[1]),
      nombre:    String(r[2]),
      telefono:  String(r[3]),
      profe:     String(r[4]),
      activo:    r[5] !== false && r[5] !== 'false' && r[5] !== 0
    });
  });
  return out;
}

function guardarAtleta(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hojaAtletas();
    var n = h.getLastRow();
    var filas = n >= 2 ? leerAtletas() : [];
    // Busca por ID si viene, sino por apellido+nombre
    var existente = data.id
      ? filas.filter(function (f) { return f.id === String(data.id); })[0]
      : null;
    if (!existente && data.apellido && data.nombre) {
      var k = norm(data.apellido) + '|' + norm(data.nombre);
      existente = filas.filter(function (f) { return norm(f.apellido) + '|' + norm(f.nombre) === k; })[0];
    }
    var actual = existente
      ? { id: existente.id, apellido: existente.apellido, nombre: existente.nombre, telefono: existente.telefono, profe: existente.profe, activo: existente.activo }
      : { id: String(Date.now()), apellido: '', nombre: '', telefono: '', profe: '', activo: true };
    if (data.apellido  !== undefined) actual.apellido  = String(data.apellido).trim();
    if (data.nombre    !== undefined) actual.nombre    = String(data.nombre).trim();
    if (data.telefono  !== undefined) actual.telefono  = String(data.telefono).trim();
    if (data.profe     !== undefined) actual.profe     = String(data.profe).trim();
    if (data.activo    !== undefined) actual.activo    = data.activo !== false && data.activo !== 'false';
    var fila = [actual.id, actual.apellido, actual.nombre, actual.telefono, actual.profe, actual.activo];
    if (existente) h.getRange(existente.fila, 1, 1, COLS_ATLETAS.length).setValues([fila]);
    else h.appendRow(fila);
    return actual;
  } finally { lock.releaseLock(); }
}

function borrarAtleta(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hojaAtletas();
    var filas = leerAtletas();
    var f = filas.filter(function (x) { return x.id === String(id); })[0];
    if (!f) return false;
    h.deleteRow(f.fila);
    return true;
  } finally { lock.releaseLock(); }
}

// ─────────────────────────────── TEST ───────────────────────────────────
function testLeer() {
  Logger.log('Pedidos: ' + JSON.stringify(leerPedidos().slice(0, 2)));
  Logger.log('Gestion: ' + JSON.stringify(leerGestion()));
  Logger.log('Prendas: ' + JSON.stringify(leerPrendas().slice(0, 2)));
  Logger.log('Atletas: ' + JSON.stringify(leerAtletas().slice(0, 5)));
}
