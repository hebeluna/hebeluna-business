/**
 * Hebeluna Business Management · API
 * Vive dentro del Google Sheet "Inventario Hebeluna y Rashi Studio".
 * El Sheet ES la base de datos: productos en las hojas Hebeluna / Rashi Studio,
 * movimientos en Movimientos, y todo lo demás del CRM en la hoja CRM_Datos.
 */

var TABS = { heb: 'Hebeluna', rs: 'Rashi Studio', mov: 'Movimientos', cfg: 'Config',
  datos: 'CRM_Datos', users: 'CRM_Usuarios' };
var HEADER_ROW = 4, FIRST_ROW = 5;
var LOC_SHEET = { jinotega: 'Local Rashi Studio', bodega: 'Bodega' };
var OWNER_ONLY = /^(data\/owner|config|requests|roles)(\/|$)/;
var TOKEN_DAYS = 30;

/* ================= entrada web ================= */
function doGet(e) {
  return out_({ ok: true, app: 'Hebeluna Business Management', ver: ver_() });
}
function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Solicitud inválida' }); }
  try {
    var a = req.action;
    if (a === 'status') return out_({ ok: true, needsOwner: users_().length === 0, ver: ver_() });
    if (a === 'firstRun') return out_(firstRun_(req));
    if (a === 'login') return out_(login_(req));
    var me = auth_(req.token);
    if (a === 'ver') return out_({ ok: true, ver: ver_() });
    if (a === 'snapshot') return out_(snapshot_(me));
    if (a === 'write') return out_(write_(me, req.ops || []));
    if (a === 'upload') return out_(upload_(me, req));
    if (a === 'changePassword') return out_(changePassword_(me, req));
    if (a === 'users.list') return out_(usersList_(me));
    if (a === 'users.save') return out_(usersSave_(me, req));
    return out_({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ================= utilidades ================= */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { return ss_().getSheetByName(name); }
function props_() { return PropertiesService.getScriptProperties(); }
function ver_() { return Number(props_().getProperty('ver') || 0); }
function bump_() { var v = ver_() + 1; props_().setProperty('ver', String(v)); return v; }
function norm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9$]+/g, ' ').trim(); }
function num_(v) { var n = parseFloat(String(v).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : 0; }
function nowIso_() { return new Date().toISOString(); }
function b64_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function secret_() {
  var s = props_().getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props_().setProperty('SECRET', s); }
  return s;
}
function hash_(salt, pw) {
  var h = salt + '|' + pw;
  for (var i = 0; i < 300; i++) h = b64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function sign_(payload) { return b64_(Utilities.computeHmacSha256Signature(payload, secret_())); }

/* se ejecuta sola cuando alguien edita el Sheet a mano: avisa al CRM que hay cambios */
function onEdit(e) { try { bump_(); } catch (err) {} }

/* ================= instalación (correr una vez) ================= */
function instalar() {
  ensureTabs_();
  secret_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'avisarCambios') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('avisarCambios').timeBased().everyMinutes(30).create();
  bump_();
  return 'Listo';
}
function ensureTabs_() {
  var s = ss_();
  if (!s.getSheetByName(TABS.datos)) {
    var d = s.insertSheet(TABS.datos);
    d.getRange(1, 1, 1, 5).setValues([['coleccion', 'id', 'json', 'actualizado', 'por']]).setFontWeight('bold');
    d.setFrozenRows(1); d.setColumnWidth(3, 500);
  }
  if (!s.getSheetByName(TABS.users)) {
    var u = s.insertSheet(TABS.users);
    u.getRange(1, 1, 1, 9).setValues([['usuario', 'nombre', 'email', 'rol', 'permisos', 'salt', 'hash', 'activo', 'creado']]).setFontWeight('bold');
    u.setFrozenRows(1); u.hideSheet();
  }
  var mv = s.getSheetByName(TABS.mov);
  var oc = originCol_(mv);
  if (!oc) {
    var c = mv.getLastColumn() + 1;
    mv.getRange(HEADER_ROW, c).setValue('Origen (CRM)');
    mv.getRange(HEADER_ROW, c - 1).copyFormatToRange(mv, c, c, HEADER_ROW, HEADER_ROW);
  }
}
function originCol_(mv) {
  var h = mv.getRange(HEADER_ROW, 1, 1, mv.getLastColumn()).getValues()[0];
  for (var i = 0; i < h.length; i++) if (norm_(h[i]).indexOf('origen') === 0) return i + 1;
  return 0;
}

/* ================= usuarios y sesiones ================= */
function users_() {
  var u = sheet_(TABS.users); if (!u) return [];
  var v = u.getDataRange().getValues(); var out = [];
  for (var i = 1; i < v.length; i++) if (v[i][0]) out.push({ row: i + 1, usuario: String(v[i][0]), nombre: v[i][1], email: v[i][2], rol: v[i][3],
    perms: v[i][4] ? JSON.parse(v[i][4]) : {}, salt: v[i][5], hash: v[i][6], activo: v[i][7] !== false && v[i][7] !== 'FALSE' });
  return out;
}
function pub_(u) { return { id: u.usuario, usuario: u.usuario, nombre: u.nombre, email: u.email, rol: u.rol, perms: u.perms, activo: u.activo, isOwner: u.rol === 'duena' }; }
function cleanUser_(s) { return String(s || '').trim().toLowerCase().replace(/[^a-z0-9._-]/g, ''); }
function checkPw_(pw) { if (String(pw || '').length < 8) throw new Error('La contraseña debe tener al menos 8 caracteres'); }
function firstRun_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ensureTabs_();
    if (users_().length) throw new Error('La cuenta de la dueña ya existe. Inicia sesión.');
    var usuario = cleanUser_(req.usuario); if (!usuario) throw new Error('Escribe un usuario');
    checkPw_(req.password);
    var salt = Utilities.getUuid();
    sheet_(TABS.users).appendRow([usuario, req.nombre || '', req.email || '', 'duena', '{}', salt, hash_(salt, req.password), true, nowIso_()]);
    return login_({ usuario: usuario, password: req.password });
  } finally { lock.releaseLock(); }
}
function login_(req) {
  var usuario = cleanUser_(req.usuario);
  var cache = CacheService.getScriptCache(); var k = 'fail_' + usuario;
  var fails = Number(cache.get(k) || 0);
  if (fails >= 8) throw new Error('Demasiados intentos. Espera 15 minutos.');
  var u = users_().filter(function (x) { return x.usuario === usuario; })[0];
  if (!u || !u.activo || hash_(u.salt, String(req.password || '')) !== u.hash) {
    cache.put(k, String(fails + 1), 900);
    throw new Error('Usuario o contraseña incorrectos');
  }
  cache.remove(k);
  var exp = Date.now() + TOKEN_DAYS * 864e5;
  var payload = u.usuario + '|' + exp + '|' + u.hash.slice(0, 8);
  return { ok: true, token: payload + '|' + sign_(payload), me: pub_(u) };
}
function auth_(token) {
  var p = String(token || '').split('|');
  if (p.length !== 4) throw new Error('SESION');
  var payload = p[0] + '|' + p[1] + '|' + p[2];
  if (sign_(payload) !== p[3] || Number(p[1]) < Date.now()) throw new Error('SESION');
  var u = users_().filter(function (x) { return x.usuario === p[0]; })[0];
  if (!u || !u.activo || u.hash.slice(0, 8) !== p[2]) throw new Error('SESION');
  return pub_(u);
}
function changePassword_(me, req) {
  var u = users_().filter(function (x) { return x.usuario === me.usuario; })[0];
  if (hash_(u.salt, String(req.actual || '')) !== u.hash) throw new Error('La contraseña actual no coincide');
  checkPw_(req.nueva);
  var salt = Utilities.getUuid();
  sheet_(TABS.users).getRange(u.row, 6, 1, 2).setValues([[salt, hash_(salt, req.nueva)]]);
  return login_({ usuario: me.usuario, password: req.nueva });
}
function usersList_(me) {
  if (!me.isOwner) throw new Error('Solo la dueña');
  return { ok: true, users: users_().map(pub_) };
}
function usersSave_(me, req) {
  if (!me.isOwner) throw new Error('Solo la dueña');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var d = req.user || {}; var usuario = cleanUser_(d.usuario); if (!usuario) throw new Error('Escribe un usuario');
    var sh = sheet_(TABS.users);
    var u = users_().filter(function (x) { return x.usuario === usuario; })[0];
    if (!u) {
      checkPw_(d.password);
      var salt = Utilities.getUuid();
      sh.appendRow([usuario, d.nombre || '', d.email || '', 'vendedora', JSON.stringify(d.perms || {}), salt, hash_(salt, d.password), d.activo !== false, nowIso_()]);
    } else {
      if (u.rol === 'duena' && d.activo === false) throw new Error('No puedes desactivar a la dueña');
      sh.getRange(u.row, 2, 1, 4).setValues([[d.nombre != null ? d.nombre : u.nombre, d.email != null ? d.email : u.email, u.rol, JSON.stringify(d.perms || u.perms)]]);
      sh.getRange(u.row, 8).setValue(d.activo !== false);
      if (d.password) { checkPw_(d.password); var s2 = Utilities.getUuid(); sh.getRange(u.row, 6, 1, 2).setValues([[s2, hash_(s2, d.password)]]); }
    }
    bump_();
    return { ok: true, users: users_().map(pub_) };
  } finally { lock.releaseLock(); }
}

/* ================= lectura: productos y movimientos desde el Sheet ================= */
function cols_(sh) {
  var h = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  function f(k) { for (var i = 0; i < h.length; i++) if (h[i] === k || h[i].indexOf(k) === 0) return i; return -1; }
  return { id: f('id'), cat: f('categoria'), marca: f('marca'), nombre: f('producto'), tono: f('tono'), talla: f('talla'), necesita: f('necesita'),
    sku: f('codigo'), precio: f('precio us$'), precioN: f('precio c$'), iniL: f('inicial local'), iniB: f('inicial bodega'), stL: f('stock local'),
    stB: f('stock bodega'), nota: f('notas'), etiqueta: f('etiqueta'), costo: f('costo'), envio: f('envio'), n: h.length };
}
function cfgSheet_() {
  var c = sheet_(TABS.cfg); var v = c.getRange('B4:B8').getValues();
  return { tasa: num_(v[0][0]) || 37.5, redondeo: num_(v[1][0]) || 10, stockCritico: num_(v[2][0]) || 1, impuesto: num_(v[3][0]), ganancia: num_(v[4][0]) };
}
function readProducts_(isOwner, cfg) {
  var prods = [], costs = [];
  [[TABS.heb, 'hebeluna'], [TABS.rs, 'rashi']].forEach(function (t) {
    var sh = sheet_(t[0]); var C = cols_(sh); var last = sh.getLastRow(); if (last < FIRST_ROW) return;
    var v = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, C.n).getValues();
    v.forEach(function (r, i) {
      var id = String(r[C.id] || '').trim(); if (!/^(HB|RS)-\d+$/.test(id)) return;
      var precio = num_(r[C.precio]); var pn = C.precioN >= 0 ? num_(r[C.precioN]) : 0;
      var calc = Math.round(precio * cfg.tasa / cfg.redondeo) * cfg.redondeo;
      prods.push({ id: id, negocio: t[1], row: FIRST_ROW + i, cat: r[C.cat] || '', marca: r[C.marca] || '', nombre: r[C.nombre] || '', tono: r[C.tono] || '',
        talla: C.talla >= 0 ? String(r[C.talla] || '') : '', necesita: C.necesita >= 0 ? (r[C.necesita] || 'No') : 'No', sku: C.sku >= 0 ? String(r[C.sku] || '') : '',
        precio: precio, precioNio: pn && precio && Math.abs(pn - calc) > cfg.redondeo ? pn : null,
        base: { jinotega: num_(r[C.iniL]), managua: 0, bodega: num_(r[C.iniB]) }, nota: C.nota >= 0 ? String(r[C.nota] || '') : '',
        etiqueta: C.etiqueta >= 0 ? String(r[C.etiqueta] || '') : '' });
      if (isOwner && C.costo >= 0 && r[C.costo] !== '' && r[C.costo] !== null) costs.push({ id: id, costo: num_(r[C.costo]), envio: C.envio >= 0 ? num_(r[C.envio]) : 0 });
    });
  });
  return { prods: prods, costs: costs };
}
function movCols_(mv) {
  var h = mv.getRange(HEADER_ROW, 1, 1, mv.getLastColumn()).getValues()[0].map(norm_);
  function f(k) { for (var i = 0; i < h.length; i++) if (h[i] === k || h[i].indexOf(k) === 0) return i; return -1; }
  return { fecha: f('fecha'), prod: f('producto'), id: f('id'), tipo: f('tipo'), ubic: f('ubicacion'), cant: f('cantidad'), cobrado: f('precio cobrado'),
    dL: f('loc') >= 0 ? f('loc') : f('local'), dB: f('bod'), nota: f('cliente'), origen: f('origen'), n: h.length };
}
function readMovs_() {
  var mv = sheet_(TABS.mov); var C = movCols_(mv); var last = mv.getLastRow(); var out = [];
  if (last < FIRST_ROW) return out;
  var v = mv.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, C.n).getValues();
  var tz = ss_().getSpreadsheetTimeZone();
  v.forEach(function (r, i) {
    var pid = String(r[C.id] || '').trim(); if (!pid) { var m = String(r[C.prod] || '').match(/^(HB|RS)-\d+/); pid = m ? m[0] : ''; }
    if (!pid) return;
    var origen = C.origen >= 0 ? String(r[C.origen] || '') : '';
    if (origen.indexOf('CRM-VENTA:') === 0) return; // las ventas del CRM ya vienen de la colección sales
    var f = r[C.fecha]; var fecha = f instanceof Date ? Utilities.formatDate(f, tz, 'yyyy-MM-dd') : String(f || '');
    out.push({ id: origen.indexOf('CRM:') === 0 ? origen.slice(4) : 'fila-' + (FIRST_ROW + i), fecha: fecha, pid: pid, tipo: 'sheet',
      tipoSheet: String(r[C.tipo] || ''), ubicSheet: String(r[C.ubic] || ''), qty: num_(r[C.cant]), dj: C.dL >= 0 ? num_(r[C.dL]) : 0, db: C.dB >= 0 ? num_(r[C.dB]) : 0,
      nota: C.nota >= 0 ? String(r[C.nota] || '') : '', origen: origen ? 'CRM' : 'Google Sheet', ts: fecha });
  });
  return out;
}

/* ================= CRM_Datos (todo lo demás) ================= */
function datosAll_() {
  var d = sheet_(TABS.datos); var v = d.getDataRange().getValues(); var out = [];
  for (var i = 1; i < v.length; i++) if (v[i][0]) out.push({ row: i + 1, col: String(v[i][0]), id: String(v[i][1]), json: v[i][2] });
  return out;
}
function snapshot_(me) {
  var cfg = cfgSheet_();
  var P = readProducts_(me.isOwner, cfg);
  var datos = datosAll_(); var cols = {};
  datos.forEach(function (x) {
    if (!me.isOwner && OWNER_ONLY.test(x.col) && x.col !== 'config' && x.col !== 'roles') return;
    if (x.col === 'requests') return;
    (cols[x.col] = cols[x.col] || []).push({ id: x.id, data: JSON.parse(x.json) });
  });
  // extras de producto (foto, proveedor, fechas, activo…) se guardan en prodx
  var extras = {}; (cols.prodx || []).forEach(function (x) { extras[x.id] = x.data; }); delete cols.prodx;
  cols.products = P.prods.map(function (p) { var e = extras[p.id] || {}; var o = {}; for (var k in e) o[k] = e[k]; for (var k2 in p) if (k2 !== 'row') o[k2] = p[k2]; return { id: p.id, data: o }; });
  if (me.isOwner) cols['data/owner/costs'] = P.costs.map(function (c) { return { id: c.id, data: { costo: c.costo, envio: c.envio } }; });
  cols.movs = readMovs_().map(function (m) { return { id: m.id, data: m }; }).concat(cols.movs || []);
  // config: lo del Sheet (tasa, redondeo, impuesto, ganancia) manda
  var main = ((cols.config || []).filter(function (x) { return x.id === 'main'; })[0] || { id: 'main', data: {} });
  main.data.tasa = cfg.tasa; main.data.redondeo = cfg.redondeo; main.data.impuesto = cfg.impuesto; main.data.ganancia = cfg.ganancia;
  cols.config = [main];
  // roles salen de los usuarios
  cols.roles = users_().map(function (u) { return { id: u.usuario, data: { nombre: u.nombre, email: u.email, rol: u.rol, perms: u.perms, activo: u.activo } }; });
  return { ok: true, ver: ver_(), me: me, cols: cols };
}

/* ================= escritura ================= */
function canWrite_(me, col) {
  if (me.isOwner) return true;
  if (OWNER_ONLY.test(col)) return false;
  var p = me.perms || {};
  if (col === 'products' || col === 'movs') return !!p.inventario;
  if (col === 'sales') return !!(p.vender || p.pedidos);
  if (col === 'customers') return !!(p.clientes || p.vender);
  if (col === 'posts') return !!p.redes;
  return col === 'audit';
}
function write_(me, ops) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var datos = null; var cfg = cfgSheet_();
    ops.forEach(function (op) {
      var parts = String(op.path).split('/'); var id = parts.pop(); var col = parts.join('/');
      if (!id || !col) throw new Error('Ruta inválida');
      if (!canWrite_(me, col)) throw new Error('No tienes permiso para cambiar ' + col);
      if (col === 'audit' && op.data) { op.data.owner = !!me.isOwner; op.data.uid = me.usuario; }
      if (col === 'products') return writeProduct_(me, id, op, cfg);
      if (col === 'data/owner/costs') return writeCost_(id, op.data || {});
      if (col === 'movs' && !isManagua_(op.data)) return writeMov_(id, op);
      if (col === 'sales') mirrorSale_(id, op);
      if (col === 'config' && id === 'main' && op.data) writeCfgSheet_(op.data);
      if (col === 'roles') return; // se manejan con users.save
      datos = datos || datosIndex_();
      upsertDatos_(datos, col, id, op.op === 'delete' ? null : op.data, me);
    });
    return { ok: true, ver: bump_() };
  } finally { lock.releaseLock(); }
}
function datosIndex_() { var m = {}; datosAll_().forEach(function (x) { m[x.col + '\u0001' + x.id] = x.row; }); return m; }
function upsertDatos_(idx, col, id, data, me) {
  var d = sheet_(TABS.datos); var key = col + '\u0001' + id; var row = idx[key];
  if (data === null) { if (row) { d.deleteRow(row); for (var k in idx) if (idx[k] > row) idx[k]--; delete idx[key]; } return; }
  var vals = [[col, id, JSON.stringify(data), nowIso_(), me.usuario]];
  if (row) d.getRange(row, 1, 1, 5).setValues(vals);
  else { d.appendRow(vals[0]); idx[key] = d.getLastRow(); }
}
function isManagua_(m) { return m && (m.loc === 'managua' || m.from === 'managua' || m.to === 'managua'); }

function findProductRow_(id) {
  var sh = sheet_(id.indexOf('RS-') === 0 ? TABS.rs : TABS.heb); var C = cols_(sh);
  var last = sh.getLastRow(); var ids = last >= FIRST_ROW ? sh.getRange(FIRST_ROW, C.id + 1, last - FIRST_ROW + 1, 1).getValues() : [];
  var lastId = FIRST_ROW - 1;
  for (var i = 0; i < ids.length; i++) { var v = String(ids[i][0] || '').trim(); if (v === id) return { sh: sh, C: C, row: FIRST_ROW + i }; if (/^(HB|RS)-\d+$/.test(v)) lastId = FIRST_ROW + i; }
  return { sh: sh, C: C, row: 0, lastId: lastId };
}
function writeProduct_(me, id, op, cfg) {
  var d = op.data || {};
  var f = findProductRow_(id); var sh = f.sh, C = f.C, row = f.row;
  if (op.op === 'delete') { d = { activo: false }; }
  if (!row && op.op !== 'delete') {
    row = f.lastId + 1;
    var nextHasId = String(sh.getRange(row, C.id + 1).getValue() || '').trim();
    if (nextHasId) { sh.insertRowAfter(f.lastId); }
    if (f.lastId >= FIRST_ROW) sh.getRange(f.lastId, 1, 1, C.n).copyTo(sh.getRange(row, 1, 1, C.n), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    sh.getRange(row, 1, 1, C.n).getFormulas()[0].forEach(function (fx, i) { if (!fx) sh.getRange(row, i + 1).clearContent(); });
    sh.getRange(row, C.id + 1).setValue(id);
    if (d.base) { sh.getRange(row, C.iniL + 1).setValue(num_(d.base.jinotega)); sh.getRange(row, C.iniB + 1).setValue(num_(d.base.bodega)); }
  }
  if (row && op.op !== 'delete') {
    var set = function (k, v) { if (C[k] >= 0 && v !== undefined) sh.getRange(row, C[k] + 1).setValue(v === null ? '' : v); };
    set('cat', d.cat); set('marca', d.marca); set('nombre', d.nombre); set('tono', d.tono); set('talla', d.talla);
    set('necesita', d.necesita || 'No'); set('sku', d.sku); set('precio', d.precio != null ? num_(d.precio) : undefined); set('nota', d.nota);
    if (C.precioN >= 0 && d.precioNio !== undefined) {
      var cell = sh.getRange(row, C.precioN + 1);
      if (d.precioNio) cell.setValue(num_(d.precioNio));
      else if (!cell.getFormula()) { var src = findFormulaRow_(sh, C.precioN + 1, row); if (src) sh.getRange(src, C.precioN + 1).copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false); }
    }
  }
  // lo que el Sheet no tiene va a prodx
  var extra = {}; ['proveedor', 'fechaIngreso', 'vence', 'estadoManual', 'foto', 'activo', 'transito'].forEach(function (k) { if (d[k] !== undefined) extra[k] = d[k]; });
  if (Object.keys(extra).length) {
    var idx = datosIndex_(); var prev = {};
    var r = idx['prodx\u0001' + id]; if (r) prev = JSON.parse(sheet_(TABS.datos).getRange(r, 3).getValue() || '{}');
    for (var k in extra) prev[k] = extra[k];
    upsertDatos_(idx, 'prodx', id, prev, me);
  }
}
function findFormulaRow_(sh, col, around) {
  for (var d = 1; d < 400; d++) { var a = [around - d, around + d]; for (var i = 0; i < 2; i++) if (a[i] >= FIRST_ROW && a[i] <= sh.getLastRow() && sh.getRange(a[i], col).getFormula()) return a[i]; }
  return 0;
}
function writeCost_(id, d) {
  var f = findProductRow_(id); if (!f.row) return;
  if (f.C.costo >= 0) f.sh.getRange(f.row, f.C.costo + 1).setValue(d.costo == null ? '' : num_(d.costo));
  if (f.C.envio >= 0 && d.envio !== undefined) f.sh.getRange(f.row, f.C.envio + 1).setValue(num_(d.envio) || '');
}

/* --- Movimientos --- */
function listValues_(rng) {
  var dv = rng.getDataValidation(); if (!dv) return [];
  var t = dv.getCriteriaType(); var c = dv.getCriteriaValues();
  if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return c[0];
  if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return c[0].getValues().map(function (r) { return r[0]; }).filter(String);
  return [];
}
function pickValue_(list, words, fallback) {
  for (var i = 0; i < list.length; i++) { var n = norm_(list[i]); if (words.every(function (w) { return n.indexOf(w) >= 0; })) return list[i]; }
  return fallback;
}
function etiquetaOf_(pid) {
  var f = findProductRow_(pid); if (!f.row) return pid;
  if (f.C.etiqueta >= 0) { var e = String(f.sh.getRange(f.row, f.C.etiqueta + 1).getValue() || ''); if (e) return e; }
  return pid;
}
function freeMovRow_(mv, C) {
  var last = mv.getLastRow(); var n = Math.max(0, last - FIRST_ROW + 1);
  var v = n ? mv.getRange(FIRST_ROW, 1, n, 2).getValues() : [];
  for (var i = 0; i < v.length; i++) if (!v[i][0] && !v[i][1]) return FIRST_ROW + i;
  var row = last + 1;
  mv.getRange(last, 1, 1, mv.getLastColumn()).copyTo(mv.getRange(row, 1, 1, mv.getLastColumn()), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  mv.getRange(last, 1, 1, mv.getLastColumn()).copyTo(mv.getRange(row, 1, 1, mv.getLastColumn()), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  mv.getRange(row, 1, 1, mv.getLastColumn()).getFormulas()[0].forEach(function (fx, i) { if (!fx) mv.getRange(row, i + 1).clearContent(); });
  return row;
}
function writeMovRow_(mv, C, row, m) {
  var put = function (k, v) { if (C[k] >= 0) mv.getRange(row, C[k] + 1).setValue(v); };
  put('fecha', m.fecha ? new Date(m.fecha + 'T12:00:00') : new Date());
  put('prod', etiquetaOf_(m.pid)); put('tipo', m.tipo); put('ubic', m.ubic); put('cant', m.cant);
  if (m.cobrado != null) put('cobrado', m.cobrado);
  put('nota', m.nota || ''); put('origen', m.origen);
}
function deleteMovRows_(mv, C, prefix) {
  var last = mv.getLastRow(); if (last < FIRST_ROW || C.origen < 0) return;
  var v = mv.getRange(FIRST_ROW, C.origen + 1, last - FIRST_ROW + 1, 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]).indexOf(prefix) === 0) {
    var r = FIRST_ROW + i; var f = mv.getRange(r, 1, 1, mv.getLastColumn()).getFormulas()[0];
    f.forEach(function (fx, j) { if (!fx) mv.getRange(r, j + 1).clearContent(); });
  }
}
function tipoSheet_(mv, C, tipo, from) {
  var list = listValues_(mv.getRange(FIRST_ROW, C.tipo + 1));
  if (tipo === 'venta') return pickValue_(list, ['venta'], 'Venta');
  if (tipo === 'ingreso') return pickValue_(list, ['ingreso'], 'Ingreso');
  if (tipo === 'egreso') return pickValue_(list, ['egreso'], 'Egreso');
  if (tipo === 'ajuste') return pickValue_(list, ['ajuste'], 'Ajuste');
  if (tipo === 'traslado') return from === 'bodega' ? pickValue_(list, ['traslado', 'bodega local'], 'Traslado Bodega → Local') : pickValue_(list, ['traslado', 'local bodega'], 'Traslado Local → Bodega');
  return tipo;
}
function ubicSheet_(mv, C, loc) {
  var list = listValues_(mv.getRange(FIRST_ROW, C.ubic + 1));
  return loc === 'bodega' ? pickValue_(list, ['bodega'], 'Bodega') : pickValue_(list, ['local'], LOC_SHEET.jinotega);
}
function writeMov_(id, op) {
  var mv = sheet_(TABS.mov); var C = movCols_(mv);
  deleteMovRows_(mv, C, 'CRM:' + id);
  if (op.op === 'delete' || !op.data) return;
  var m = op.data; var q = num_(m.qty);
  var row = freeMovRow_(mv, C);
  writeMovRow_(mv, C, row, { fecha: m.fecha, pid: m.pid, tipo: tipoSheet_(mv, C, m.tipo, m.from), ubic: ubicSheet_(mv, C, m.tipo === 'traslado' ? m.from : m.loc),
    cant: m.tipo === 'ajuste' ? q : Math.abs(q), nota: m.nota || '', origen: 'CRM:' + id });
}
var VENDIDO = ['Pagado', 'Preparando', 'Listo', 'Enviado', 'Entregado'];
function mirrorSale_(id, op) {
  var mv = sheet_(TABS.mov); var C = movCols_(mv);
  deleteMovRows_(mv, C, 'CRM-VENTA:' + id);
  var s = op.data; if (op.op === 'delete' || !s || VENDIDO.indexOf(s.estado) < 0 || s.loc === 'managua') return;
  (s.items || []).forEach(function (it, n) {
    var q = num_(it.qty); if (!q || !/^(HB|RS)-\d+$/.test(it.pid || '')) return; // productos por encargo no están en el inventario
    var row = freeMovRow_(mv, C);
    writeMovRow_(mv, C, row, { fecha: s.fecha, pid: it.pid, tipo: tipoSheet_(mv, C, 'venta'), ubic: ubicSheet_(mv, C, s.loc || 'jinotega'), cant: q,
      cobrado: Math.round((q * num_(it.precio) - num_(it.desc)) / q * 100) / 100, nota: (s.num || '') + (s.clienteNombre ? ' · ' + s.clienteNombre : ''),
      origen: 'CRM-VENTA:' + id + ':' + n });
  });
}
function writeCfgSheet_(d) {
  var c = sheet_(TABS.cfg);
  if (d.tasa) c.getRange('B4').setValue(num_(d.tasa));
  if (d.redondeo) c.getRange('B5').setValue(num_(d.redondeo));
  if (d.impuesto != null) c.getRange('B7').setValue(num_(d.impuesto));
  if (d.ganancia != null) c.getRange('B8').setValue(num_(d.ganancia));
}

/* ================= fotos ================= */
function upload_(me, req) {
  var folders = DriveApp.getFoldersByName('Hebeluna CRM · Fotos');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Hebeluna CRM · Fotos');
  var bytes = Utilities.base64Decode(req.base64);
  if (bytes.length > 8 * 1024 * 1024) throw new Error('La foto pesa más de 8 MB');
  var f = folder.createFile(Utilities.newBlob(bytes, req.mime || 'image/jpeg', req.name || 'foto.jpg'));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: f.getId() };
}

/* ================= avisos por correo (cada 30 min) ================= */
function avisarCambios() {
  var d = sheet_(TABS.datos); if (!d) return;
  var v = d.getDataRange().getValues(); var pend = [];
  for (var i = 1; i < v.length; i++) if (v[i][0] === 'audit') { var a = JSON.parse(v[i][2]); if (!a.owner && !a.notificado) pend.push({ row: i + 1, a: a }); }
  if (!pend.length) return;
  var cfg = (datosAll_().filter(function (x) { return x.col === 'config' && x.id === 'main'; })[0] || {});
  var to = (cfg.json ? JSON.parse(cfg.json).ownerEmail : '') || Session.getEffectiveUser().getEmail();
  var tz = ss_().getSpreadsheetTimeZone();
  var names = {}; users_().forEach(function (u) { names[u.usuario] = u.nombre || u.usuario; });
  var lines = pend.sort(function (x, y) { return String(x.a.ts).localeCompare(String(y.a.ts)); }).map(function (p) {
    return '• ' + Utilities.formatDate(new Date(p.a.ts), tz, 'dd/MM HH:mm') + ' · ' + (names[p.a.uid] || p.a.uid) + ' · ' + p.a.accion + (p.a.detalle ? ' — ' + p.a.detalle : '');
  });
  MailApp.sendEmail(to, 'Hebeluna CRM · ' + pend.length + ' cambios de tu equipo', 'Esto cambió tu equipo en el CRM:\n\n' + lines.join('\n') + '\n\nHebeluna Business Management');
  pend.forEach(function (p) { p.a.notificado = true; d.getRange(p.row, 3).setValue(JSON.stringify(p.a)); });
}

/* ================= prueba rápida (opcional) ================= */
function revisar() {
  ensureTabs_();
  var cfg = cfgSheet_(); var P = readProducts_(true, cfg); var M = readMovs_();
  var mv = sheet_(TABS.mov); var C = movCols_(mv);
  return JSON.stringify({ productos: P.prods.length, costos: P.costs.length, movimientos: M.length, tasa: cfg.tasa,
    tipos: listValues_(mv.getRange(FIRST_ROW, C.tipo + 1)), ubicaciones: listValues_(mv.getRange(FIRST_ROW, C.ubic + 1)), columnasMov: C });
}

/* Prueba completa sin tocar datos reales: lee, compara stock con el Sheet y hace escrituras de prueba que luego borra */
function probar() {
  ensureTabs_();
  var me = { usuario: 'prueba', isOwner: true, perms: {} };
  var snap = snapshot_(me); var c = snap.cols;
  var st = {};
  (c.products || []).forEach(function (p) { st[p.id] = { j: p.data.base.jinotega, b: p.data.base.bodega }; });
  (c.movs || []).forEach(function (m) { var d = m.data; if (d.tipo === 'sheet' && st[d.pid]) { st[d.pid].j += d.dj; st[d.pid].b += d.db; } });
  var bad = [];
  [TABS.heb, TABS.rs].forEach(function (t) {
    var sh = sheet_(t); var C = cols_(sh); var v = sh.getRange(FIRST_ROW, 1, sh.getLastRow() - FIRST_ROW + 1, C.n).getValues();
    v.forEach(function (r) { var id = String(r[C.id] || ''); if (st[id] && (num_(r[C.stL]) !== st[id].j || num_(r[C.stB]) !== st[id].b)) bad.push(id + ' sheet ' + r[C.stL] + '/' + r[C.stB] + ' crm ' + st[id].j + '/' + st[id].b); });
  });
  var mv = sheet_(TABS.mov); var C = movCols_(mv); var res = {};
  function deltaFor(prefix) {
    var last = mv.getLastRow(); var v = mv.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, C.n).getValues();
    for (var i = 0; i < v.length; i++) if (String(v[i][C.origen]).indexOf(prefix) === 0) return [v[i][C.tipo], v[i][C.ubic], v[i][C.cant], v[i][C.dL], v[i][C.dB], v[i][C.id]];
    return null;
  }
  writeMov_('prueba-ajuste', { op: 'set', data: { pid: 'HB-001', tipo: 'ajuste', qty: 1, loc: 'jinotega', fecha: '2026-09-25', nota: 'PRUEBA (se borra sola)' } });
  SpreadsheetApp.flush(); res.ajuste = deltaFor('CRM:prueba-ajuste');
  writeMov_('prueba-tras', { op: 'set', data: { pid: 'HB-064', tipo: 'traslado', qty: 1, from: 'bodega', to: 'jinotega', fecha: '2026-09-25', nota: 'PRUEBA' } });
  SpreadsheetApp.flush(); res.traslado = deltaFor('CRM:prueba-tras');
  mirrorSale_('prueba-venta', { op: 'set', data: { estado: 'Pagado', loc: 'jinotega', fecha: '2026-09-25', num: 'PRUEBA', items: [{ pid: 'HB-001', qty: 1, precio: 26.3, desc: 0 }] } });
  SpreadsheetApp.flush(); res.venta = deltaFor('CRM-VENTA:prueba-venta');
  deleteMovRows_(mv, C, 'CRM:prueba-'); deleteMovRows_(mv, C, 'CRM-VENTA:prueba-'); SpreadsheetApp.flush();
  res.limpio = !deltaFor('CRM:prueba-') && !deltaFor('CRM-VENTA:prueba-');
  var out = { productos: (c.products || []).length, costos: (c['data/owner/costs'] || []).length, movimientos: (c.movs || []).length, stockDistinto: bad.slice(0, 10), nStockDistinto: bad.length,
    tipos: listValues_(mv.getRange(FIRST_ROW, C.tipo + 1)), ubicaciones: listValues_(mv.getRange(FIRST_ROW, C.ubic + 1)), pruebas: res, config: c.config[0].data };
  console.log(JSON.stringify(out));
  return out;
}
