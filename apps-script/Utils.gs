/**
 * Utils.gs — دوال مساعدة لقراءة وكتابة الصفوف بأسماء الأعمدة.
 */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('الورقة غير موجودة: ' + name + ' — شغّل "تجهيز النظام" أولاً');
  return sh;
}

/** يرجّع { sheet, headers, rows: [{_row, عمود: قيمة, ...}] } */
function readTable_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  const rows = values
    .map((r, i) => {
      const o = { _row: i + 2 };
      headers.forEach((h, j) => (o[h] = r[j]));
      return o;
    })
    .filter((o) => headers.some((h) => o[h] !== '' && o[h] !== null));
  return { sheet: sh, headers, rows };
}

function colIndex_(headers, name) {
  const i = headers.indexOf(name);
  if (i < 0) throw new Error('العمود غير موجود: ' + name);
  return i + 1;
}

function setCell_(table, rowNum, colName, value) {
  table.sheet.getRange(rowNum, colIndex_(table.headers, colName)).setValue(value);
}

/** يضيف صفاً من كائن {عمود: قيمة}. */
function appendObject_(name, obj) {
  const sh = sheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  sh.appendRow(headers.map((h) => (h in obj ? obj[h] : '')));
}

/** رقم تسلسلي بالبادئة، مثلاً TR-0007 */
function nextId_(name, colName, prefix) {
  const t = readTable_(name);
  let max = 0;
  t.rows.forEach((r) => {
    const m = String(r[colName] || '').match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return prefix + '-' + String(max + 1).padStart(4, '0');
}

function toDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function daysBetween_(from, to) {
  return Math.round((startOfDay_(to) - startOfDay_(from)) / 86400000);
}

function fmt_(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}

function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, 'المساعد الإداري', 6);
}
