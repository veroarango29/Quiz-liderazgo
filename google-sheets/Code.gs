/**
 * The Authority Map · receptor de leads en Google Sheets
 * Pega este código en Extensiones → Apps Script de tu hoja y publícalo como "Aplicación web".
 */
const SHEET_NAME = 'Leads';
const COLS = ['createdAt','name','email','role','typeName','merit','auth','answers','source','consentText'];
const HEAD = ['Fecha','Nombre','Correo','Etapa','Tipo','Mérito /8','Autoridad /8','Respuestas','Origen','Consentimiento'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) { sh.appendRow(HEAD); sh.setFrozenRows(1); }
    sh.appendRow(COLS.map(c => clean(d[c])));
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error');
  } finally {
    lock.releaseLock();
  }
}

// Evita que un texto que empiece por = + - @ se interprete como fórmula.
function clean(v) {
  const s = String(v === undefined || v === null ? '' : v).slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
