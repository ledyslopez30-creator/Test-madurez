/**
 * Test de madurez digital · Colombia Travel Expo 2026
 * Pegar en una hoja de Google NUEVA: Extensiones > Apps Script.
 * Implementar > Nueva implementación > Aplicación web
 *   Ejecutar como: Yo · Quién tiene acceso: Cualquier persona.
 *
 * doPost: guarda cada respuesta del test en la pestaña "Respuestas".
 * doGet : devuelve SOLO cifras agregadas (sin nombres) para el tablero en vivo.
 */
const HOJA = 'Respuestas';
const DIMS = ['presenciaDigital', 'gestionCliente', 'usoDatos', 'iaAutomatizacion', 'sostenibilidad'];
const PERFILES = ['Inicio digital', 'Adoptante temprana', 'Empresa conectada', 'Empresa inteligente'];

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sh = hoja_();
    let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    Object.keys(data).forEach(k => {
      if (headers.indexOf(k) < 0) { headers.push(k); sh.getRange(1, headers.length).setValue(k); }
    });
    sh.appendRow(headers.map(h => (data[h] !== undefined ? data[h] : '')));
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const sh = hoja_();
  if (sh.getLastRow() < 2) return json_({ n: 0, perfiles: {}, dims: {}, subsectores: {}, roles: {}, actualizado: new Date() });
  const rows = sh.getDataRange().getValues();
  const h = rows.shift();
  const col = k => h.indexOf(k);
  const out = { n: 0, total: 0, perfiles: {}, dims: {}, subsectores: {}, roles: {}, actualizado: new Date() };
  PERFILES.forEach(p => out.perfiles[p] = 0);
  DIMS.forEach(d => out.dims[d] = 0);
  rows.forEach(r => {
    const tot = Number(r[col('puntajeTotal')]);
    if (isNaN(tot)) return;
    out.n++; out.total += tot;
    const p = r[col('perfil')]; out.perfiles[p] = (out.perfiles[p] || 0) + 1;
    const sub = (col('subsector') >= 0 && r[col('subsector')]) || 'Sin dato';
    const rol = (col('rol') >= 0 && r[col('rol')]) || 'Sin dato';
    out.roles[rol] = (out.roles[rol] || 0) + 1;
    const S = out.subsectores[sub] = out.subsectores[sub] || { n: 0, total: 0, dims: {}, perfiles: {} };
    S.n++; S.total += tot; S.perfiles[p] = (S.perfiles[p] || 0) + 1;
    DIMS.forEach(d => {
      const v = Number(r[col(d)]) || 0;
      out.dims[d] += v; S.dims[d] = (S.dims[d] || 0) + v;
    });
  });
  // promedios
  const avg = (x, n) => (n ? Math.round((x / n) * 10) / 10 : 0);
  out.total = avg(out.total, out.n);
  DIMS.forEach(d => out.dims[d] = avg(out.dims[d], out.n));
  Object.values(out.subsectores).forEach(S => {
    S.total = avg(S.total, S.n);
    DIMS.forEach(d => S.dims[d] = avg(S.dims[d] || 0, S.n));
  });
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
