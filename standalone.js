/* Hebeluna Business Management · versión independiente
   Reemplaza el runtime de Claude (db, user, assets, downloads) por la API del Google Sheet (Apps Script). */
(function () {
  "use strict";
  var API = window.HBM_API_URL;
  var TK = "hbm.token";
  var store = new Map();          // "col/id" -> data
  var subs = [];                  // {col, q, cb}
  var session = null;             // {token, me}
  var ver = 0, pending = 0, queue = [], flushT = null;
  var readyResolve; var ready = new Promise(function (r) { readyResolve = r; });

  function getTok() { try { return localStorage.getItem(TK); } catch (e) { return null; } }
  function setTok(t) { try { t ? localStorage.setItem(TK, t) : localStorage.removeItem(TK); } catch (e) {} }

  async function api(action, payload) {
    var body = Object.assign({ action: action, token: session && session.token }, payload || {});
    var r, j;
    try { r = await fetch(API, { method: "POST", body: JSON.stringify(body), redirect: "follow" }); j = await r.json(); }
    catch (e) { throw new Error("Sin conexión con el servidor. Revisa tu internet e intenta de nuevo."); }
    if (!j.ok) { if (j.error === "SESION") { logout(true); throw new Error("Tu sesión venció. Vuelve a entrar."); } throw new Error(j.error || "Error"); }
    return j;
  }

  /* ---------- base de datos en memoria + sincronía ---------- */
  function colOf(p) { return p.split("/").slice(0, -1).join("/"); }
  function clone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }
  function snapOf(col, q) {
    var docs = [];
    store.forEach(function (d, p) { if (colOf(p) === col) docs.push({ id: p.split("/").pop(), exists: true, data: (function (x) { return function () { return clone(x); }; })(d) }); });
    if (q && q.order) { var f = q.order[0], dir = q.order[1] === "desc" ? -1 : 1; docs.sort(function (a, b) { var x = a.data()[f], y = b.data()[f]; return (x > y ? 1 : x < y ? -1 : 0) * dir; }); }
    if (q && q.lim) docs = docs.slice(0, q.lim);
    return { docs: docs, size: docs.length, empty: !docs.length };
  }
  function notify(col) { subs.forEach(function (s) { if (col == null || s.col === col) try { s.cb(snapOf(s.col, s.q)); } catch (e) { console.error(e); } }); }
  function load(cols) { store.clear(); Object.keys(cols).forEach(function (c) { cols[c].forEach(function (d) { store.set(c + "/" + d.id, d.data); }); }); }
  async function refresh() { var j = await api("snapshot"); ver = j.ver; load(j.cols); notify(null); }

  function enqueue(op) {
    if (/^requests\//.test(op.path)) return Promise.resolve();
    return new Promise(function (res, rej) { queue.push({ op: op, res: res, rej: rej }); clearTimeout(flushT); flushT = setTimeout(flush, 120); });
  }
  async function flush() {
    var batch = queue.splice(0, 40); if (!batch.length) return;
    pending++;
    try { var j = await api("write", { ops: batch.map(function (b) { return b.op; }) }); ver = j.ver; batch.forEach(function (b) { b.res(); }); }
    catch (e) { batch.forEach(function (b) { b.rej(e); }); try { await refresh(); } catch (e2) {} }
    finally { pending--; if (queue.length) flush(); }
  }
  function query(col, q) {
    return {
      orderBy: function (f, d) { return query(col, Object.assign({}, q, { order: [f, d || "asc"] })); },
      limit: function (n) { return query(col, Object.assign({}, q, { lim: n })); },
      where: function () { return query(col, q); },
      onSnapshot: function (cb) { var s = { col: col, q: q, cb: cb }; subs.push(s); setTimeout(function () { cb(snapOf(col, q)); }, 0); return function () { subs.splice(subs.indexOf(s), 1); }; },
      get: async function () { return snapOf(col, q); },
      doc: function (id) { return docRef(col + "/" + id); }
    };
  }
  function docRef(path) {
    return {
      get: async function () { var d = store.get(path); return { id: path.split("/").pop(), exists: d !== undefined, data: function () { return clone(d); } }; },
      set: async function (d) { var prev = store.get(path); store.set(path, clone(d)); notify(colOf(path)); try { await enqueue({ op: "set", path: path, data: d }); } catch (e) { if (prev === undefined) store.delete(path); else store.set(path, prev); notify(colOf(path)); throw e; } },
      update: async function (d) { var n = Object.assign({}, store.get(path) || {}, d); return this.set(n); },
      delete: async function () { store.delete(path); notify(colOf(path)); await enqueue({ op: "delete", path: path }); },
      onSnapshot: function (cb) { return query(colOf(path), {}).onSnapshot(function () { var d = store.get(path); cb({ id: path.split("/").pop(), exists: d !== undefined, data: function () { return clone(d); } }); }); }
    };
  }
  var db = { collection: function (c) { return query(c, {}); }, doc: docRef };

  /* ---------- usuario, fotos, descargas ---------- */
  var user = {
    me: async function () { var m = session.me; return { id: m.usuario, name: m.nombre || m.usuario, email: m.email || null, avatarUrl: "", color: "#9A9CEA", isOwner: !!m.isOwner, canEdit: !!m.isOwner }; },
    isOwner: async function () { return !!session.me.isOwner; },
    canEdit: async function () { return !!session.me.isOwner; },
    can: async function () { return true; },
    id: async function () { return session.me.usuario; },
    profiles: async function (ids) { var o = {}; [].concat(ids).forEach(function (i) { var r = store.get("roles/" + i) || {}; o[i] = { id: i, name: r.nombre || i, avatarUrl: "", color: "#9A9CEA", email: r.email || null, isMe: i === session.me.usuario, guest: false }; }); return o; }
  };
  var assets = {
    upload: async function (blob) {
      var small = await shrink(blob);
      var b64 = await new Promise(function (res) { var fr = new FileReader(); fr.onload = function () { res(String(fr.result).split(",")[1]); }; fr.readAsDataURL(small); });
      var j = await api("upload", { base64: b64, mime: small.type || "image/jpeg", name: blob.name || "foto.jpg" });
      return { id: j.id, url: "https://drive.google.com/thumbnail?sz=w600&id=" + j.id, sizeBytes: small.size, contentType: small.type };
    }
  };
  function shrink(file) {
    return new Promise(function (res) {
      if (!/^image\//.test(file.type) || file.size < 400000) return res(file);
      var img = new Image(); img.onload = function () { var k = Math.min(1, 1400 / Math.max(img.width, img.height)); var c = document.createElement("canvas"); c.width = img.width * k; c.height = img.height * k; c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); c.toBlob(function (b) { res(b || file); }, "image/jpeg", 0.85); };
      img.onerror = function () { res(file); }; img.src = URL.createObjectURL(file);
    });
  }
  var downloads = {
    save: async function (r) {
      var blob = r.data instanceof Blob ? r.data : new Blob([r.data]);
      var a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = r.filename; document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
      return { status: "saved" };
    }
  };

  /* ---------- sesión ---------- */
  function logout(silent) { setTok(null); session = null; if (!silent) location.reload(); else setTimeout(function () { location.reload(); }, 1200); }
  async function changePassword(actual, nueva) { var j = await api("changePassword", { actual: actual, nueva: nueva }); session.token = j.token; setTok(j.token); }
  window.HBM = { api: api, logout: function () { logout(false); }, changePassword: changePassword, refresh: refresh };
  window.claude = { use: async function (n) { await ready; return { db: db, user: user, assets: assets, downloads: downloads }[n] || null; } };

  /* poll: si alguien (o el Sheet) cambió algo, recargar */
  setInterval(async function () {
    if (!session || pending || queue.length || document.hidden) return;
    try { var j = await api("ver"); if (j.ver !== ver) await refresh(); } catch (e) {}
  }, 15000);
  document.addEventListener("visibilitychange", function () { if (!document.hidden && session && !pending) api("ver").then(function (j) { if (j.ver !== ver) refresh(); }).catch(function () {}); });

  /* ---------- pantalla de entrada ---------- */
  function gate(html) {
    var root = document.getElementById("root");
    root.innerHTML = '<div class="gate"><div class="brand" style="align-items:center"><div class="logo" role="img" aria-label="Hebeluna"></div><div class="sub">Business Management</div><div class="muted" style="font-style:italic;font-size:13px;margin-top:4px">Self-care isn\'t selfish</div></div>' + html + "</div>";
  }
  function esc(s) { return String(s || "").replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function form(title, fields, btn, note, onSubmit) {
    gate('<form id="hbm_f" class="card" style="width:min(380px,100%);display:grid;gap:10px;text-align:left" novalidate><h2 style="margin:0">' + title + "</h2>" +
      fields.map(function (f) { return '<label class="f"><span>' + f[1] + '</span><input id="hbm_' + f[0] + '" type="' + (f[2] || "text") + '" autocomplete="' + (f[3] || "off") + '" ' + (f[4] ? "required" : "") + "></label>"; }).join("") +
      '<button class="btn p" style="justify-content:center" type="submit">' + btn + '</button><p id="hbm_msg" class="muted" style="margin:0;font-size:13px">' + (note || "") + "</p></form>");
    var f = document.getElementById("hbm_f"); var first = f.querySelector("input"); if (first) first.focus();
    f.addEventListener("submit", async function (e) {
      e.preventDefault(); var v = {}; fields.forEach(function (x) { v[x[0]] = document.getElementById("hbm_" + x[0]).value; });
      var b = f.querySelector("button"); b.disabled = true; var msg = document.getElementById("hbm_msg"); msg.textContent = "Un momento…";
      try { await onSubmit(v); } catch (err) { msg.textContent = err.message; b.disabled = false; }
    });
  }
  async function start(j) {
    session = { token: j.token, me: j.me }; setTok(j.token);
    gate('<p class="muted">Cargando tu negocio…</p>');
    await refresh(); readyResolve();
  }
  async function boot() {
    if (!API) { gate("<p>Falta configurar la dirección del servidor (config.js).</p>"); return; }
    var tok = getTok();
    if (tok) { session = { token: tok, me: null }; try { var s = await api("snapshot"); session.me = s.me; ver = s.ver; load(s.cols); readyResolve(); return; } catch (e) { session = null; setTok(null); } }
    var st; try { st = await api("status"); } catch (e) { gate('<p>' + esc(e.message) + '</p><button class="btn" onclick="location.reload()">Reintentar</button>'); return; }
    if (st.needsOwner) {
      form("Crea tu cuenta de dueña", [["nombre", "Tu nombre", "text", "name"], ["email", "Correo para avisos", "email", "email"], ["usuario", "Usuario", "text", "username", 1], ["password", "Contraseña (mínimo 8)", "password", "new-password", 1]],
        "Crear cuenta y entrar", "Solo se hace una vez. Después, tú creas los usuarios de tu equipo desde Equipo.",
        async function (v) { await start(await api("firstRun", v)); });
    } else {
      form("Entrar", [["usuario", "Usuario", "text", "username", 1], ["password", "Contraseña", "password", "current-password", 1]], "Entrar", "",
        async function (v) { await start(await api("login", v)); });
    }
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot); else boot();
})();
