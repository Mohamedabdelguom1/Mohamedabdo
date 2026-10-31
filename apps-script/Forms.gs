/**
 * Forms.gs — نماذج إدخال سريعة (بدل الكتابة المباشرة في الجدول) + فحص وتنظيف البيانات.
 */

function showTraineeForm() {
  const courses = readTable_(SHEETS.COURSES).rows
    .filter((c) => !['منتهي', 'ملغي'].includes(c['الحالة']))
    .map((c) => ({ code: String(c['كود الكورس']), name: String(c['اسم الكورس']) }));
  const t = HtmlService.createTemplateFromFile('TraineeForm');
  t.courses = courses;
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('متدرب جديد'));
}

function saveTrainee(form) {
  if (!form.name || !form.course) throw new Error('الاسم والكورس مطلوبين');
  const phone = normalizePhone_(form.phone);
  const dup = readTable_(SHEETS.TRAINEES).rows.find(
    (r) => phone && normalizePhone_(r['الهاتف']) === phone && String(r['كود الكورس']) === form.course
  );
  if (dup) throw new Error('المتدرب مسجّل مسبقاً في هذا الكورس: ' + dup['رقم المتدرب']);
  const id = nextId_(SHEETS.TRAINEES, 'رقم المتدرب', 'TR');
  appendObject_(SHEETS.TRAINEES, {
    'رقم المتدرب': id,
    'الاسم': form.name.trim(),
    'الهاتف': phone,
    'البريد': (form.email || '').trim().toLowerCase(),
    'كود الكورس': form.course,
    'تاريخ التسجيل': new Date(),
    'الحالة': 'مسجّل',
    'ملاحظات': form.notes || '',
  });
  return id;
}

function showTaskForm() {
  const staff = readTable_(SHEETS.STAFF).rows.filter((s) => s['نشط'] !== 'لا').map((s) => String(s['الاسم']));
  const t = HtmlService.createTemplateFromFile('TaskForm');
  t.staff = staff;
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('مهمة / موعد جديد'));
}

function saveTask(form) {
  if (!form.title || !form.date) throw new Error('العنوان والتاريخ مطلوبين');
  const id = nextId_(SHEETS.TASKS, 'رقم', 'T');
  const [y, m, d] = form.date.split('-').map(Number);
  appendObject_(SHEETS.TASKS, {
    'رقم': id,
    'العنوان': form.title.trim(),
    'النوع': form.type || 'مهمة',
    'التاريخ': new Date(y, m - 1, d),
    'الوقت': form.time || '',
    'المسؤول': form.owner || '',
    'الأولوية': form.priority || 'متوسطة',
    'الحالة': 'مفتوحة',
    'ملاحظات': form.notes || '',
  });
  return id;
}

/** توحيد أرقام الهواتف السودانية: 09xxxxxxxx / +2499xxxxxxxx → 2499xxxxxxxx */
function normalizePhone_(p) {
  // تحويل الأرقام العربية/الهندية ثم حذف أي رمز غير رقمي
  let s = String(p || '')
    .replace(/[٠-٩]/g, (c) => String(c.charCodeAt(0) - 0x0660))
    .replace(/[^\d]/g, '');
  if (s.startsWith('00')) s = s.slice(2);
  if (s.startsWith('0') && s.length === 10) s = '249' + s.slice(1);
  if (s.length === 9 && /^[19]/.test(s)) s = '249' + s;
  return s;
}

/** فحص وتنظيف: مسافات زائدة، توحيد الهواتف والبريد، وتمييز التكرار بالأصفر. */
function cleanData() {
  let fixed = 0, dups = 0;
  [SHEETS.TRAINEES, SHEETS.STAFF].forEach((name) => {
    const t = readTable_(name);
    const seen = {};
    t.rows.forEach((r) => {
      ['الاسم', 'الهاتف', 'البريد'].forEach((col) => {
        if (!(col in r)) return;
        let v = r[col];
        let nv = col === 'الهاتف' ? normalizePhone_(v) : col === 'البريد' ? String(v).trim().toLowerCase() : String(v).replace(/\s+/g, ' ').trim();
        if (String(v) !== nv) {
          setCell_(t, r._row, col, nv);
          fixed++;
        }
      });
      const key = normalizePhone_(r['الهاتف']) + '|' + (r['كود الكورس'] || '');
      const range = t.sheet.getRange(r._row, 1, 1, t.headers.length);
      if (normalizePhone_(r['الهاتف']) && seen[key]) {
        range.setBackground('#fff3cd');
        dups++;
      } else {
        range.setBackground(null);
        seen[key] = true;
      }
    });
  });
  toast_(`تم تعديل ${fixed} خلية، ووجدنا ${dups} صف مكرر (ملوّن بالأصفر)`);
}
