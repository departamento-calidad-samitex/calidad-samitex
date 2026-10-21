/**
 * AUDITORÍA FINAL · Samitex S.A. · Formato SA-AC-PR-05.01
 * Servidor (API) de la app web instalable. La app (PWA) vive en GitHub Pages y funciona sin señal;
 * cuando hay conexión envía los reportes aquí. Este servidor guarda en Google Sheets/Drive y
 * envía cada reporte con su PDF al CORREO, más un Excel diario con todas las auditorías.
 * Instalación: ejecutar instalar() una vez y publicar como Aplicación web (Ejecutar como: Yo · Acceso: Cualquier persona).
 */

/* ===== CONFIGURACIÓN ===== */
var HOJA_ID = '1F5_tC2rSSNT1PpU1Ox4bMbGYgv22XRBo81a_hHCJ4a8'; // hoja de cálculo «Auditoría Final · Datos»
var CORREO = '';            // correo que recibe los reportes. Vacío = la cuenta dueña de la app.
var RESUMEN_HORA = 19;      // hora (0-23, Lima) del Excel diario con todas las auditorías
var ENVIAR_CADA_REPORTE = true;


/* Código de administradora: solo se guarda su huella cifrada (SHA-256 con sal), nunca el número. */
var ADMIN_HASH = '31bbff557f1b249bf9633fb7305135d19be64f41ba0ec236900b493ffd543cd0';
var SAL = 'samitex-af:';

/* Tabla de muestreo de prendas Samitex: [lote máximo, muestra, acepta mayor (AQL 2.5), acepta menor (AQL 4.0)] */
var TABLA = [[8,2,0,0],[15,3,0,0],[25,5,0,0],[50,8,0,1],[90,13,1,1],[150,20,1,2],[280,32,2,3],[500,50,3,5],[1200,80,5,7],[3200,125,7,10],[10000,200,10,14],[35000,315,14,21],[150000,500,21,21]];
var MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
var NDEF = 15; // grupos COD/NOMBRE/PUNTAJE en la hoja DATA
var H_DATA = (function(){var h=['MES','FECHA AUDITORIA','LINEA','PROCEDENCIA','TALLER','CLIENTE','OPERACION FINAL','OF','ARTICULO','COLOR','LOTE','MUESTRA','STATUS','AUDITOR','CANTIDAD DEFECTOS'];
  for(var i=1;i<=NDEF;i++)h.push('COD'+i,'NOMBRE '+i,'PUNTAJE '+i);
  return h.concat(['OBSERVACIONES','ID','TIPO','DESCRIPCION','MAYORES','MENORES','% FALLAS','AC MAYOR 2.5','RE MAYOR 2.5','AC MENOR 4.0','RE MENOR 4.0','PRUEBAS RECHAZADAS','CHECKLIST NO CONFORME','DNI AUDITOR','RESPONSABLE TALLER','V°B°','PDF','FOTOS','REGISTRADO','CORREGIDO'])})();
var H_AUD = ['DNI','NOMBRE','CARGO','ESTADO','SOLICITADO','APROBADO','HISTORIAL HASTA','ULTIMO INGRESO','REPORTES','LLAVES'];
var H_REP = ['ID','FECHA','DNI','AUDITOR','OF','STATUS','JSON','PDF ID','REGISTRADO','CORREGIDO','VB FECHA'];

/* ======================= API ======================= */
var API = ['solicitarAcceso','ingresar','estado','adminIngresar','adminAuditores','adminAccion','adminRegistrar','guardarReporte','misReportes','adminReportes','adminVB','descargarPdf','adminTablero','adminConfig'];
function doPost(e) {
  var out;
  try {
    var q = JSON.parse(e.postData.contents || '{}');
    if (API.indexOf(q.fn) < 0) throw new Error('Función no permitida');
    out = { ok: true, r: globalThis[q.fn].apply(null, q.args || []) };
  } catch (err) { out = { ok: false, e: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput(JSON.stringify({ app: 'Auditoría Final', ok: true })).setMimeType(ContentService.MimeType.JSON); }
var REPORTE_URL = 'https://karinauscubilcaaguilar-droid.github.io/auditoria-final/Reporte.html';
function plantilla_() { // plantilla del PDF SA-AC-PR-05.01, publicada junto a la app
  var c = cache_(), k = 'tpl_reporte', v = c.get(k); if (v) return v;
  v = UrlFetchApp.fetch(REPORTE_URL + '?v=' + Date.now()).getContentText('UTF-8');
  if (v.length < 90000) c.put(k, v, 21600); return v;
}
var SS_; function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(HOJA_ID)); }
function correo_() { return CORREO || Session.getEffectiveUser().getEmail(); }

/* ======================= instalación ======================= */
function instalar() {
  var ss = ss_();
  hoja_('DATA', H_DATA); hoja_('AUDITORES', H_AUD); hoja_('REPORTES', H_REP);
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('CARPETA')) {
    var raiz = DriveApp.createFolder('Auditoría Final · Samitex');
    p.setProperty('CARPETA', raiz.getId());
    p.setProperty('CARPETA_FOTOS', raiz.createFolder('Fotos').getId());
    p.setProperty('CARPETA_FIRMAS', raiz.createFolder('Firmas').getId());
    p.setProperty('CARPETA_PDF', raiz.createFolder('PDF').getId());
  }
  ss.rename('Auditoría Final · Samitex · Datos');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'enviarResumen') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('enviarResumen').timeBased().everyDays(1).atHour(RESUMEN_HORA).inTimezone('America/Lima').create();
  return 'Listo: hojas, carpeta en Drive y Excel diario a ' + correo_() + ' a las ' + RESUMEN_HORA + ':00.';
}
function hoja_(nombre, h) {
  var ss = ss_(), sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#0a1f74').setFontColor('#ffffff'); sh.setFrozenRows(1); }
  else { // agrega columnas que falten, sin tocar las existentes (por ejemplo si pegas tu DATA histórica)
    var cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
    var falt = h.filter(function (x) { return cur.indexOf(norm_(x)) < 0; });
    if (falt.length) sh.getRange(1, cur.length + 1, 1, falt.length).setValues([falt]).setFontWeight('bold');
  }
  return sh;
}
function sh_(n) { var s = ss_().getSheetByName(n); if (!s) throw new Error('Falta la hoja ' + n + '. Ejecuta instalar().'); return s; }

/* ======================= utilidades ======================= */
function norm_(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toUpperCase(); }
function up_(s) { return String(s == null ? '' : s).trim().toUpperCase(); }
function hex_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join(''); }
function dniOk_(d) { return /^\d{8}$/.test(String(d || '')); }
function cache_() { return CacheService.getScriptCache(); }
function headers_(sh) { return sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(norm_); }
function rows_(sh) { var n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues(); }
function plan_(lote) {
  lote = Math.floor(Number(lote) || 0); if (lote < 1) return null;
  var k = TABLA.length - 1; for (var i = 0; i < TABLA.length; i++) if (lote <= TABLA[i][0]) { k = i; break; }
  var t = TABLA[k]; return { n: Math.min(t[1], lote), acM: t[2], reM: t[2] + 1, acm: t[3], rem: t[3] + 1 };
}
function lock_(fn) { var l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
function fecha_(f) { var m = String(f || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + '/' + m[2] + '/' + m[1] : String(f || ''); }

/* ======================= sesiones ======================= */
function nuevaSesion_(datos) { var t = Utilities.getUuid().replace(/-/g, ''); cache_().put('s_' + t, JSON.stringify(datos), 21600); return t; }
function sesion_(t, rol) {
  t = String(t || '');
  if (t.indexOf(':') > 0) { // auditor: DNI + llave del dispositivo (no vence; se revoca al bloquear)
    if (rol && rol !== 'aud') throw new Error('SESION');
    var dni = t.split(':')[0], k = t.split(':')[1], a = buscarAud_(dni);
    if (!a || a.ESTADO !== 'ACTIVO' || String(a.LLAVES || '').split(',').indexOf(k) < 0) throw new Error('SESION');
    return { rol: 'aud', dni: dni, nombre: a.NOMBRE, aud: a };
  }
  var v = t && cache_().get('s_' + t); if (!v) throw new Error('SESION');
  var s = JSON.parse(v); if (rol && s.rol !== rol) throw new Error('SESION'); return s;
}

/* ======================= auditores ======================= */
function buscarAud_(dni) {
  var sh = sh_('AUDITORES'), H = headers_(sh), r = rows_(sh), c = H.indexOf('DNI');
  for (var i = 0; i < r.length; i++) if (String(r[i][c]).replace(/\D/g, '') === dni) {
    var o = { fila: i + 2 }; H.forEach(function (h, j) { o[h] = r[i][j]; }); return o;
  }
  return null;
}
function setAud_(fila, campos) {
  var sh = sh_('AUDITORES'), H = headers_(sh);
  Object.keys(campos).forEach(function (k) { var c = H.indexOf(norm_(k)); if (c >= 0) sh.getRange(fila, c + 1).setValue(campos[k]); });
}
function solicitarAcceso(nombre, dni, cargo) {
  dni = String(dni || '').replace(/\D/g, ''); nombre = up_(nombre);
  if (!nombre || !dniOk_(dni)) return { ok: false, msg: 'Escribe tu nombre y un DNI de 8 dígitos.' };
  return lock_(function () {
    var a = buscarAud_(dni);
    if (a) {
      if (a.ESTADO === 'BLOQUEADO') return { ok: false, msg: 'Tu acceso está bloqueado. Habla con la jefatura de calidad.' };
      if (a.ESTADO === 'ACTIVO') return { ok: true, estado: 'ACTIVO' };
      setAud_(a.fila, { NOMBRE: nombre, CARGO: cargo || a.CARGO, ESTADO: 'PENDIENTE', SOLICITADO: new Date() });
      return { ok: true, estado: 'PENDIENTE' };
    }
    sh_('AUDITORES').appendRow(H_AUD.map(function (h) { return ({ 'DNI': "'" + dni, 'NOMBRE': nombre, 'CARGO': cargo || 'AUDITOR DE CALIDAD', 'ESTADO': 'PENDIENTE', 'SOLICITADO': new Date(), 'REPORTES': 0 })[h] || ''; }));
    return { ok: true, estado: 'PENDIENTE' };
  });
}
function ingresar(dni) {
  dni = String(dni || '').replace(/\D/g, '');
  if (!dniOk_(dni)) return { ok: false, msg: 'El DNI tiene 8 dígitos.' };
  var a = buscarAud_(dni);
  if (!a) return { ok: false, estado: 'NUEVO', msg: 'Este DNI no tiene acceso. Envía tu solicitud.' };
  if (a.ESTADO === 'PENDIENTE') return { ok: false, estado: 'PENDIENTE', nombre: a.NOMBRE };
  if (a.ESTADO !== 'ACTIVO') return { ok: false, estado: a.ESTADO, msg: a.ESTADO === 'BLOQUEADO' ? 'Tu acceso fue bloqueado por la jefatura de calidad.' : 'Tu solicitud no fue aprobada. Habla con la jefatura de calidad.' };
  var llave = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
  var ll = String(a.LLAVES || '').split(',').filter(String).concat(llave).slice(-5); // hasta 5 equipos por auditor
  setAud_(a.fila, { 'ULTIMO INGRESO': new Date(), LLAVES: ll.join(',') });
  var hasta = a['HISTORIAL HASTA'] instanceof Date ? a['HISTORIAL HASTA'].getTime() : 0;
  return { ok: true, token: dni + ':' + llave, nombre: a.NOMBRE, historialHasta: hasta > Date.now() ? hasta : 0 };
}
/* Comprueba si el equipo sigue autorizado (al abrir la app con señal) */
function estado(t) {
  try { var s = sesion_(t, 'aud'); var h = s.aud['HISTORIAL HASTA'] instanceof Date ? s.aud['HISTORIAL HASTA'].getTime() : 0;
    return { ok: true, nombre: s.nombre, historialHasta: h > Date.now() ? h : 0 }; }
  catch (e) { var a = buscarAud_(String(t || '').split(':')[0]); return { ok: false, estado: a ? a.ESTADO : 'NUEVO' }; }
}

/* ======================= administradora ======================= */
function adminIngresar(codigo) {
  var c = cache_(), f = Number(c.get('adm_fallos') || 0);
  if (f >= 5) return { ok: false, msg: 'Demasiados intentos. Espera 15 minutos.' };
  if (hex_(SAL + String(codigo || '').trim()) !== ADMIN_HASH) { c.put('adm_fallos', String(f + 1), 900); return { ok: false, msg: 'Código incorrecto.' }; }
  c.remove('adm_fallos');
  return { ok: true, token: nuevaSesion_({ rol: 'adm', nombre: 'ADMINISTRADORA' }) };
}
function adminAuditores(t) {
  sesion_(t, 'adm'); var sh = sh_('AUDITORES'), H = headers_(sh);
  return rows_(sh).map(function (r) { var o = {}; H.forEach(function (h, j) { o[h] = r[j] instanceof Date ? r[j].getTime() : r[j]; }); o.DNI = String(o.DNI).replace(/\D/g, ''); return o; })
    .filter(function (o) { return o.DNI; });
}
function adminAccion(t, dni, accion, horas) {
  sesion_(t, 'adm'); dni = String(dni || '').replace(/\D/g, '');
  return lock_(function () {
    var a = buscarAud_(dni); if (!a) return { ok: false, msg: 'No encontré ese DNI.' };
    if (accion === 'aprobar') setAud_(a.fila, { ESTADO: 'ACTIVO', APROBADO: new Date() });
    else if (accion === 'rechazar') setAud_(a.fila, { ESTADO: 'RECHAZADO' });
    else if (accion === 'bloquear') setAud_(a.fila, { ESTADO: 'BLOQUEADO', 'HISTORIAL HASTA': '', LLAVES: '' });
    else if (accion === 'reactivar') setAud_(a.fila, { ESTADO: 'ACTIVO' });
    else if (accion === 'historial') setAud_(a.fila, { 'HISTORIAL HASTA': new Date(Date.now() + Number(horas || 24) * 3600e3) });
    else if (accion === 'quitarHistorial') setAud_(a.fila, { 'HISTORIAL HASTA': '' });
    else return { ok: false, msg: 'Acción desconocida.' };
    return { ok: true };
  });
}
function adminRegistrar(t, nombre, dni, cargo) {
  sesion_(t, 'adm'); dni = String(dni || '').replace(/\D/g, '');
  if (!up_(nombre) || !dniOk_(dni)) return { ok: false, msg: 'Escribe el nombre y un DNI de 8 dígitos.' };
  return lock_(function () {
    if (buscarAud_(dni)) return { ok: false, msg: 'Ese DNI ya está registrado.' };
    sh_('AUDITORES').appendRow(H_AUD.map(function (h) { return ({ 'DNI': "'" + dni, 'NOMBRE': up_(nombre), 'CARGO': cargo || 'AUDITOR DE CALIDAD', 'ESTADO': 'ACTIVO', 'SOLICITADO': new Date(), 'APROBADO': new Date(), 'REPORTES': 0 })[h] || ''; }));
    return { ok: true };
  });
}

/* ======================= reportes ======================= */
function guardarReporte(t, rec, fotos, firmas) {
  var s;
  try { s = sesion_(t); } catch (e) { return { ok: false, sesion: true, msg: 'Tu sesión venció. Vuelve a ingresar.' }; }
  var esAdm = s.rol === 'adm';
  return lock_(function () {
    var p = PropertiesService.getScriptProperties(), corr = false, prev = null;
    if (!rec || !rec.id) return { ok: false, msg: 'Reporte vacío.' };
    var repSh = sh_('REPORTES'), R = rows_(repSh), fila = -1;
    for (var i = 0; i < R.length; i++) if (R[i][0] === rec.id) { fila = i + 2; prev = JSON.parse(R[i][6] || '{}'); break; }
    if (fila > 0) { // es una corrección
      if (!esAdm) {
        var a = buscarAud_(s.dni), h = a && a['HISTORIAL HASTA'] instanceof Date ? a['HISTORIAL HASTA'].getTime() : 0;
        if (String(prev.dni) === s.dni && !rec.esCorreccion) return { ok: true, id: rec.id, rej: prev.rej, repetido: true }; // reenvío desde la cola
        if (String(prev.dni) !== s.dni || h < Date.now()) return { ok: false, msg: 'No tienes autorización para corregir este reporte.' };
      }
      corr = true;
    }

    /* datos de identidad: los pone el servidor, no el celular */
    rec.dni = corr ? prev.dni : (esAdm ? (rec.dni || '') : s.dni);
    rec.aud = corr ? prev.aud : (esAdm ? up_(rec.aud || 'ADMINISTRADORA') : s.nombre);
    /* dictamen recalculado en el servidor con la tabla Samitex */
    var pl = plan_(rec.lote), may = 0, men = 0;
    Object.keys(rec.d || {}).forEach(function (c) { may += Number(rec.d[c].M) || 0; men += Number(rec.d[c].m) || 0; });
    rec.may = may; rec.men = men; rec.def = may + men; rec.mue = Number(rec.mue) || (pl ? pl.n : 0);
    rec.acM = pl ? pl.acM : ''; rec.reM = pl ? pl.reM : ''; rec.acm = pl ? pl.acm : ''; rec.rem = pl ? pl.rem : '';
    var prR = Object.keys(rec.pr || {}).filter(function (k) { return rec.pr[k] === 'R'; });
    rec.rej = (pl && (may >= pl.reM || men >= pl.rem)) || prR.length ? 1 : 0;
    rec.fallas = rec.mue ? Math.round(rec.def / rec.mue * 10000) / 100 : 0;
    var fe = String(rec.fecha || ''); rec.mes = MESES[Number(fe.slice(5, 7)) - 1] || '';

    /* fotos y firmas a Drive */
    var nombreBase = 'OF' + String(rec.of || 'SIN-OF').replace(/[^A-Za-z0-9-]/g, '') + '_' + fe + '_' + rec.id;
    var keep = rec.fotosKeep || [];
    rec.fotos = corr ? (prev.fotos || []).filter(function (id) { return keep.indexOf(id) >= 0; }) : [];
    (fotos || []).slice(0, 6).forEach(function (b64, i) {
      var f = DriveApp.getFolderById(p.getProperty('CARPETA_FOTOS')).createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', nombreBase + '_foto' + (rec.fotos.length + 1) + '.jpg'));
      rec.fotos.push(f.getId());
    });
    ['A', 'T'].forEach(function (k) {
      var b = firmas && firmas[k];
      if (b) rec['firma' + k] = DriveApp.getFolderById(p.getProperty('CARPETA_FIRMAS')).createFile(Utilities.newBlob(Utilities.base64Decode(b), 'image/png', nombreBase + '_firma' + k + '.png')).getId();
      else if (corr) rec['firma' + k] = prev['firma' + k] || '';
    });
    delete rec.fotosKeep; delete rec.esCorreccion;
    if (corr) { rec.ts = prev.ts; rec.vb = prev.vb || null; rec.editado = Date.now(); } else rec.ts = Date.now();

    /* PDF con el formato */
    var pdfId = pdf_(rec, corr ? (R[fila - 2][7] || '') : '');

    /* hojas */
    var json = JSON.stringify(rec);
    var fila_rep = [rec.id, rec.fecha, "'" + rec.dni, rec.aud, rec.of, rec.rej ? 'RECHAZADO' : 'APROBADO', json, pdfId, new Date(rec.ts), corr ? new Date() : '', rec.vb ? new Date(rec.vb) : ''];
    if (corr) repSh.getRange(fila, 1, 1, fila_rep.length).setValues([fila_rep]); else repSh.appendRow(fila_rep);
    escribirData_(rec, pdfId, corr);
    if (!corr && !esAdm) { var au = buscarAud_(s.dni); if (au) setAud_(au.fila, { REPORTES: (Number(au.REPORTES) || 0) + 1 }); }
    if (ENVIAR_CADA_REPORTE) { try { correoReporte_(rec, pdfId, corr); } catch (e) { console.log('correo: ' + e); } }
    return { ok: true, id: rec.id, rej: rec.rej, corregido: corr };
  });
}
function escribirData_(r, pdfId, corr) {
  var sh = sh_('DATA'), H = headers_(sh), row = new Array(H.length).fill('');
  var set = function (h, v) { var c = H.indexOf(norm_(h)); if (c >= 0) row[c] = v; };
  set('MES', r.mes); set('FECHA AUDITORIA', fecha_(r.fecha)); set('LINEA', r.lin); set('PROCEDENCIA', r.proc); set('TALLER', r.tal); set('CLIENTE', r.cli);
  set('OPERACION FINAL', r.op); set('OF', r.of); set('ARTICULO', r.art); set('COLOR', r.color); set('LOTE', r.lote); set('MUESTRA', r.mue);
  set('STATUS', r.rej ? 'RECHAZADO' : 'APROBADO'); set('AUDITOR', r.aud); set('CANTIDAD DEFECTOS', r.def);
  var cods = H.map(function (h, i) { return /^COD\d*$/.test(h) ? i : -1; }).filter(function (i) { return i >= 0; }),
      noms = H.map(function (h, i) { return h.indexOf('NOMB') === 0 ? i : -1; }).filter(function (i) { return i >= 0; }),
      pts = H.map(function (h, i) { return h.indexOf('PUNTAJE') === 0 ? i : -1; }).filter(function (i) { return i >= 0; });
  var defs = Object.keys(r.d || {}).map(function (c) { return [c, r.nombres && r.nombres[c] || c, (Number(r.d[c].M) || 0) + (Number(r.d[c].m) || 0)]; }).sort(function (a, b) { return b[2] - a[2]; });
  defs.slice(0, Math.min(cods.length, noms.length, pts.length)).forEach(function (d, i) { row[cods[i]] = d[0]; row[noms[i]] = d[1]; row[pts[i]] = d[2]; });
  set('OBSERVACIONES', (r.obs || []).map(function (o) { return [o.t, o.ft && 'FT: ' + o.ft, o.pf && 'PF: ' + o.pf].filter(String).join(' '); }).join(' | '));
  set('ID', r.id); set('TIPO', r.tipo); set('DESCRIPCION', r.desc); set('MAYORES', r.may); set('MENORES', r.men); set('% FALLAS', r.fallas);
  set('AC MAYOR 2.5', r.acM); set('RE MAYOR 2.5', r.reM); set('AC MENOR 4.0', r.acm); set('RE MENOR 4.0', r.rem);
  set('PRUEBAS RECHAZADAS', Object.keys(r.pr || {}).filter(function (k) { return r.pr[k] === 'R'; }).join(' '));
  set('CHECKLIST NO CONFORME', Object.keys(r.ck || {}).filter(function (k) { return r.ck[k] === 'X'; }).map(function (k) { return CHECK[k]; }).join(' | '));
  set('DNI AUDITOR', "'" + r.dni); set('RESPONSABLE TALLER', r.talResp); set('V°B°', r.vb ? new Date(r.vb) : '');
  set('PDF', pdfId ? 'https://drive.google.com/file/d/' + pdfId + '/view' : ''); set('FOTOS', (r.fotos || []).length);
  set('REGISTRADO', new Date(r.ts)); set('CORREGIDO', r.editado ? new Date(r.editado) : '');
  var idc = H.indexOf('ID');
  if (corr && idc >= 0) {
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, idc + 1, sh.getLastRow() - 1, 1).getValues() : [];
    for (var i = 0; i < ids.length; i++) if (ids[i][0] === r.id) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}
var CHECK = ['CALIDAD DE CONFECCIÓN', 'MEDIDAS/FIT', 'CUMPL. DE FICHA TÉCNICA', 'MATERIAL/COLOR/TEST', 'EMPAQUE Y EMBALAJE'];

/* ======================= consultas ======================= */
function reportes_(filtro) {
  return rows_(sh_('REPORTES')).filter(function (r) { return r[0]; }).map(function (r) {
    var o = JSON.parse(r[6] || '{}'); o.pdfId = r[7]; return o;
  }).filter(filtro || function () { return true; }).sort(function (a, b) { return (b.ts || 0) - (a.ts || 0); });
}
function misReportes(t) {
  var s = sesion_(t, 'aud'), a = buscarAud_(s.dni), h = a && a['HISTORIAL HASTA'] instanceof Date ? a['HISTORIAL HASTA'].getTime() : 0;
  if (h < Date.now()) return { ok: false, msg: 'No tienes autorización para ver tu historial.' };
  var desde = Date.now() - 60 * 864e5;
  return { ok: true, hasta: h, rows: reportes_(function (o) { return String(o.dni) === s.dni && (o.ts || 0) >= desde; }).slice(0, 150) };
}
function adminReportes(t) { sesion_(t, 'adm'); return reportes_().slice(0, 400); }
function adminVB(t, id) {
  sesion_(t, 'adm');
  return lock_(function () {
    var sh = sh_('REPORTES'), R = rows_(sh);
    for (var i = 0; i < R.length; i++) if (R[i][0] === id) {
      var o = JSON.parse(R[i][6]); o.vb = Date.now(); var pdfId = pdf_(o, R[i][7]);
      sh.getRange(i + 2, 7, 1, 5).setValues([[JSON.stringify(o), pdfId, R[i][8], R[i][9], new Date(o.vb)]]);
      escribirData_(o, pdfId, true); return { ok: true };
    }
    return { ok: false, msg: 'No encontré el reporte.' };
  });
}
/* Descarga del PDF: devuelve el archivo en base64 para que el celular lo guarde */
function descargarPdf(t, id) {
  var s = sesion_(t), R = rows_(sh_('REPORTES'));
  for (var i = 0; i < R.length; i++) if (R[i][0] === id) {
    var o = JSON.parse(R[i][6]);
    if (s.rol !== 'adm' && String(o.dni) !== s.dni) return { ok: false, msg: 'No puedes descargar este reporte.' };
    var pid = R[i][7] || pdf_(o, '');
    return { ok: true, nombre: DriveApp.getFileById(pid).getName(), b64: Utilities.base64Encode(DriveApp.getFileById(pid).getBlob().getBytes()) };
  }
  return { ok: false, msg: 'No encontré el reporte.' };
}
/* Tablero: lee la hoja DATA (incluye tu histórico si lo pegas ahí) */
function adminTablero(t) {
  sesion_(t, 'adm'); var sh = sh_('DATA'), H = headers_(sh), R = rows_(sh);
  var c = function (k) { return H.indexOf(k); }, cods = [], noms = [], pts = [];
  H.forEach(function (h, i) { if (/^COD\d*$/.test(h)) cods.push(i); if (h.indexOf('NOMB') === 0) noms.push(i); if (h.indexOf('PUNTAJE') === 0) pts.push(i); });
  var C = { mes: c('MES'), fecha: c('FECHA AUDITORIA'), lin: c('LINEA'), proc: c('PROCEDENCIA'), tal: c('TALLER'), cli: c('CLIENTE'), op: c('OPERACION FINAL'), of: c('OF'), st: c('STATUS'), aud: c('AUDITOR'), def: c('CANTIDAD DEFECTOS'), lote: c('LOTE'), mue: c('MUESTRA'), id: c('ID') };
  var g = function (r, i) { return i < 0 ? '' : r[i]; };
  return R.filter(function (r) { return up_(g(r, C.st)); }).map(function (r) {
    var defs = []; for (var j = 0; j < Math.max(cods.length, noms.length); j++) { var n = up_(g(r, noms[j])) || up_(g(r, cods[j])); if (n) defs.push([n, Number(g(r, pts[j])) || 1]); }
    var f = g(r, C.fecha); f = f instanceof Date ? Utilities.formatDate(f, 'America/Lima', 'yyyy-MM-dd') : String(f).replace(/^(\d{1,2})\/(\d{1,2})\/(\d{4}).*$/, function (_, d, m, y) { return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2); });
    var mes = up_(g(r, C.mes)); if (mes === 'SEPTIEMBRE') mes = 'SETIEMBRE';
    return [mes, f, up_(g(r, C.lin)), up_(g(r, C.proc)), up_(g(r, C.tal)), up_(g(r, C.cli)), up_(g(r, C.op)), String(g(r, C.of)), up_(g(r, C.st)) === 'RECHAZADO' ? 1 : 0, up_(g(r, C.aud)), Number(g(r, C.def)) || 0, Number(g(r, C.lote)) || 0, Number(g(r, C.mue)) || 0, defs, g(r, C.id) ? 1 : 0];
  });
}

/* ======================= PDF SA-AC-PR-05.01 ======================= */
function img64_(id, mime) { try { return id ? 'data:' + mime + ';base64,' + Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes()) : ''; } catch (e) { return ''; } }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function pdf_(r, viejoId) {
  var t = HtmlService.createTemplate(plantilla_()); t.r = r; t.CHECK = CHECK; t.fecha = fecha_; t.esc = esc_;
  t.firmaA = img64_(r.firmaA, 'image/png'); t.firmaT = img64_(r.firmaT, 'image/png');
  t.fotos = (r.fotos || []).slice(0, 6).map(function (id) { return img64_(id, 'image/jpeg'); }).filter(String);
  t.vbFecha = r.vb ? Utilities.formatDate(new Date(r.vb), 'America/Lima', 'dd/MM/yyyy HH:mm') : '';
  var nombre = 'AF_OF' + String(r.of || 'SIN-OF').replace(/[^A-Za-z0-9-]/g, '') + '_' + (r.fecha || '') + '_' + String(r.id).slice(-6) + '.pdf';
  var blob = t.evaluate().getBlob().getAs(MimeType.PDF).setName(nombre);
  var carpeta = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('CARPETA_PDF'));
  if (viejoId) { try { DriveApp.getFileById(viejoId).setTrashed(true); } catch (e) {} }
  return carpeta.createFile(blob).getId();
}

/* ======================= correos ======================= */
function correoReporte_(r, pdfId, corr) {
  var est = r.rej ? 'RECHAZADO' : 'APROBADO';
  var asunto = 'Auditoría Final · OF ' + (r.of || '—') + ' · ' + est + ' · ' + r.aud + (corr ? ' (corregido)' : '');
  var defs = Object.keys(r.d || {}).map(function (c) { var x = r.d[c]; return c + ' ' + ((r.nombres || {})[c] || '') + ': ' + (x.M ? x.M + ' may. ' : '') + (x.m ? x.m + ' men.' : ''); });
  var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0b1030">' +
    '<h2 style="color:#0a1f74;margin:0 0 6px">Reporte de auditoría final · <span style="color:' + (r.rej ? '#d32027' : '#0b7a2a') + '">' + est + '</span></h2>' +
    '<p style="margin:0 0 10px;color:#4a5070">SA-AC-PR-05.01 · ' + fecha_(r.fecha) + ' · ' + esc_(r.proc) + ' · ' + esc_(r.op) + '</p>' +
    '<table cellpadding="4" style="border-collapse:collapse;font-size:13px">' +
    [['Cliente', r.cli], ['Proveedor / taller', r.tal], ['OF / OC', r.of], ['Artículo', r.art], ['Color', r.color], ['Línea', r.lin], ['Auditor', r.aud],
     ['Lote / muestra', r.lote + ' / ' + r.mue], ['Mayores (Ac ' + r.acM + ')', r.may], ['Menores (Ac ' + r.acm + ')', r.men], ['% fallas', r.fallas + '%']]
      .map(function (x) { return '<tr><td style="color:#767c96;border-bottom:1px solid #e2e4ec">' + x[0] + '</td><td style="border-bottom:1px solid #e2e4ec"><b>' + esc_(x[1]) + '</b></td></tr>'; }).join('') +
    '</table>' + (defs.length ? '<p><b>Defectos:</b><br>' + defs.map(esc_).join('<br>') + '</p>' : '<p>Sin defectos encontrados.</p>') +
    '<p style="color:#767c96;font-size:12px">El PDF del formato va adjunto. Fotos y firmas en la carpeta «Auditoría Final · Samitex» de Drive.</p></div>';
  var att = []; try { att.push(DriveApp.getFileById(pdfId).getBlob()); } catch (e) {}
  MailApp.sendEmail({ to: correo_(), subject: asunto, htmlBody: html, attachments: att, name: 'Auditoría Final' });
}
/* Excel con todas las auditorías (hoja DATA). Se envía cada día por el activador; también puedes ejecutarlo a mano. */
function enviarResumen() {
  var ss = ss_(), hoy = Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd');
  var R = reportes_(function (o) { return Utilities.formatDate(new Date(o.ts || 0), 'America/Lima', 'yyyy-MM-dd') === hoy; });
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName('Auditoria_Final_' + hoy + '.xlsx');
  var rej = R.filter(function (o) { return o.rej; }).length;
  MailApp.sendEmail({ to: correo_(), name: 'Auditoría Final', subject: 'Auditoría Final · resumen ' + hoy + ' · ' + R.length + ' auditorías hoy',
    htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px"><h2 style="color:#0a1f74">Resumen del ' + fecha_(hoy) + '</h2><p>Auditorías registradas hoy: <b>' + R.length + '</b> · aprobadas <b>' + (R.length - rej) + '</b> · rechazadas <b style="color:#d32027">' + rej + '</b>.</p><p>Adjunto el Excel con todas las hojas (DATA tiene las mismas columnas que tu tablero).</p></div>',
    attachments: [blob] });
}
function adminConfig(t) { sesion_(t, 'adm'); return { correo: correo_(), hora: RESUMEN_HORA, cadaReporte: ENVIAR_CADA_REPORTE, hoja: ss_().getUrl() }; }
