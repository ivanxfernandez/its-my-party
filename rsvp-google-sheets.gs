// Confirmaciones de la fiesta: guarda una fila por invitado en la pestaña "RSVP".
// Pega esto en Extensiones > Apps Script de tu hoja y publícalo como Aplicación web.

var SHEET_NAME = 'RSVP';
var VALID = { si: 'Sí voy', talvez: 'Tal vez', no: 'No puede' };

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Nombre', 'Respuesta', 'Código', 'Fecha']);
  }
  return sh;
}

function readAll_() {
  var rows = getSheet_().getDataRange().getValues();
  var out = {};
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] && VALID[rows[i][2]]) out[rows[i][0]] = { status: rows[i][2] };
  }
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  var d = {};
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  var name = String(d.name || '').trim().slice(0, 40);
  var st = String(d.status || '');
  if (!name || !VALID[st]) return json_({ error: 'Datos inválidos' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_();
    var rows = sh.getDataRange().getValues();
    var row = -1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === name) { row = i + 1; break; }
    }
    var values = [name, VALID[st], st, new Date()];
    if (row > 0) sh.getRange(row, 1, 1, 4).setValues([values]);
    else sh.appendRow(values);
  } finally {
    lock.releaseLock();
  }
  return json_(readAll_());
}
