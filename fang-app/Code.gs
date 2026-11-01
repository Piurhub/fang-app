// Google Apps Script: วางทั้งหมดใน Extensions > Apps Script ของ Google Sheets
// ใช้ 3 ชีต: รายได้ / รายจ่าย / สต็อก (สร้างให้เองอัตโนมัติ)
const TOKEN = 'เปลี่ยนเป็นรหัสลับของคุณ'; // ต้องตรงกับที่ใส่ในหน้าเว็บ
const SALES = ['รายได้', ['id', 'วันที่', 'จำนวนเงิน', 'ประเภท', 'โน้ต', 'จำนวน']];
const EXP = ['รายจ่าย', ['id', 'วันที่', 'ประเภท', 'จำนวนเงิน', 'โน้ต']];
const STOCK = ['สต็อก', ['ประเภท', 'คงเหลือ']];

function sheet_(def) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(def[0]);
  if (!s) {
    s = ss.insertSheet(def[0]);
    s.appendRow(def[1]);
    if (def !== STOCK) s.getRange('B:B').setNumberFormat('@');
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function day_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
}
function body_(s) {
  return s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues() : [];
}
function del_(s, id) {
  const v = s.getRange(1, 1, s.getLastRow(), 1).getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (Number(v[i][0]) === Number(id)) { s.deleteRow(i + 1); return true; }
  }
  return false;
}
function replace_(s, rows, w) {
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, w).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, w).setValues(rows);
}
const sale_ = r => [r.id, r.d, r.a, r.c, r.n || '', r.q || 1];
const exp_ = r => [r.id, r.d, r.c, r.a, r.n || ''];

function doGet(e) {
  if (e.parameter.token !== TOKEN) return out_({ ok: false, error: 'token' });
  const stock = {};
  body_(sheet_(STOCK)).forEach(r => { stock[String(r[0])] = Number(r[1]); });
  return out_({
    ok: true,
    rows: body_(sheet_(SALES)).filter(r => r[0] !== '').map(r => ({ id: Number(r[0]), d: day_(r[1]), a: Number(r[2]), c: String(r[3]), n: String(r[4] || ''), q: Number(r[5] || 1) })),
    exp: body_(sheet_(EXP)).filter(r => r[0] !== '').map(r => ({ id: Number(r[0]), d: day_(r[1]), c: String(r[2]), a: Number(r[3]), n: String(r[4] || '') })),
    stock: stock
  });
}
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.token !== TOKEN) return out_({ ok: false, error: 'token' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (b.action === 'add') sheet_(SALES).appendRow(sale_(b.row));
    else if (b.action === 'add_exp') sheet_(EXP).appendRow(exp_(b.row));
    else if (b.action === 'set_stock') replace_(sheet_(STOCK), Object.keys(b.stock).map(k => [k, b.stock[k]]), 2);
    else if (b.action === 'delete') { if (!del_(sheet_(SALES), b.id)) del_(sheet_(EXP), b.id); }
    else if (b.action === 'bulk') { // เขียนทับทั้งสามชีตด้วยข้อมูลจากหน้าเว็บ
      replace_(sheet_(SALES), (b.rows || []).map(sale_), 6);
      replace_(sheet_(EXP), (b.exp || []).map(exp_), 5);
      replace_(sheet_(STOCK), Object.keys(b.stock || {}).map(k => [k, b.stock[k]]), 2);
    }
  } finally { lock.releaseLock(); }
  return out_({ ok: true });
}
