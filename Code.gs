// Gastos de a dos - receptor para Google Sheets
// 1) Cambiá KEY por una clave larga y propia (la misma que vas a poner en la app).
const KEY = 'CAMBIAR_POR_UNA_CLAVE_LARGA';
const SHEET = 'Gastos';
const HEAD = ['id','fecha','tipo','descripcion','grupo','monto','pago','pct_persona1','parte_persona1','parte_persona2','recurrente','usd','tipo_cambio','nota','subgrupo'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.key !== KEY) return out_({ ok: false, error: 'clave incorrecta' });
    const s = sheet_();
    if (p.action === 'upsert') upsert_(s, p.row);
    else if (p.action === 'delete') remove_(s, p.id);
    else if (p.action === 'replaceAll') {
      s.clearContents();
      s.appendRow(HEAD);
      if (p.rows.length) s.getRange(2, 1, p.rows.length, HEAD.length).setValues(p.rows);
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.release();
  }
}

function doGet() { return out_({ ok: true, msg: 'activo' }); }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) s = ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) { s.appendRow(HEAD); s.setFrozenRows(1); }
  return s;
}
function find_(s, id) {
  const n = s.getLastRow();
  if (n < 2) return -1;
  const ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function upsert_(s, row) {
  const r = find_(s, row[0]);
  if (r < 0) s.appendRow(row); else s.getRange(r, 1, 1, row.length).setValues([row]);
}
function remove_(s, id) { const r = find_(s, id); if (r > 0) s.deleteRow(r); }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
