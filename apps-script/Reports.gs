/**
 * Reports.gs — لوحة المتابعة داخل الملف + التقرير الشهري كملف Word.
 */

function collectStats_() {
  const trainees = readTable_(SHEETS.TRAINEES).rows;
  const courses = readTable_(SHEETS.COURSES).rows;
  const certs = readTable_(SHEETS.CERTS).rows;
  const staff = readTable_(SHEETS.STAFF).rows;
  const tasks = readTable_(SHEETS.TASKS).rows;
  const mail = readTable_(SHEETS.MAIL).rows;
  const today = new Date();
  const count = (arr, col) => arr.reduce((m, r) => ((m[r[col] || 'غير محدد'] = (m[r[col] || 'غير محدد'] || 0) + 1), m), {});
  return {
    trainees: trainees.length,
    traineesByStatus: count(trainees, 'الحالة'),
    courses: courses.length,
    coursesByStatus: count(courses, 'الحالة'),
    certs: certs.length,
    staffActive: staff.filter((s) => s['نشط'] !== 'لا').length,
    staffByRole: count(staff.filter((s) => s['نشط'] !== 'لا'), 'الدور'),
    tasksOpen: tasks.filter((t) => t['الحالة'] !== 'منجزة').length,
    tasksOverdue: tasks.filter((t) => t['الحالة'] !== 'منجزة' && toDate_(t['التاريخ']) && daysBetween_(today, toDate_(t['التاريخ'])) < 0).length,
    mailOpen: mail.filter((m) => !['تم الرد', 'مؤرشفة'].includes(m['الحالة'])).length,
    courseRows: courses.map((c) => [c['كود الكورس'], c['اسم الكورس'], c['المدرب'], fmt_(toDate_(c['تاريخ البداية'])), c['عدد المسجلين'] || 0, c['الحالة']]),
  };
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.DASH) || ss.insertSheet(SHEETS.DASH, 0);
  sh.clear();
  sh.setRightToLeft(true);
  const s = collectStats_();
  const rows = [
    ['لوحة المتابعة — آخر تحديث', fmt_(new Date())],
    ['', ''],
    ['إجمالي المتدربين', s.trainees],
    ['الكورسات', s.courses],
    ['الشهادات الصادرة', s.certs],
    ['فريق العمل النشط', s.staffActive],
    ['مهام مفتوحة', s.tasksOpen],
    ['مهام متأخرة', s.tasksOverdue],
    ['مراسلات تحتاج متابعة', s.mailOpen],
    ['', ''],
    ['المتدربين حسب الحالة', ''],
    ...Object.entries(s.traineesByStatus),
    ['', ''],
    ['الفريق حسب الدور', ''],
    ...Object.entries(s.staffByRole),
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A1:B1').setFontWeight('bold').setBackground('#1f4e78').setFontColor('#fff');
  sh.getRange(3, 1, 7, 1).setFontWeight('bold');
  sh.getRange(8, 2).setFontColor(s.tasksOverdue ? '#c0392b' : '#117a65').setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}

/** يولّد تقرير Word (.docx) في نفس مجلد الملف، ويرسله للمدير لو sendMail = true */
function generateMonthlyReport(sendMail) {
  const ss = SpreadsheetApp.getActive();
  const s = collectStats_();
  const center = getSetting('اسم المركز') || ss.getName();
  const title = `التقرير الإداري — ${center} — ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM')}`;

  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  const rtl = (p) => p.setLeftToRight(false).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  rtl(body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1));
  rtl(body.appendParagraph('ملخص عام').setHeading(DocumentApp.ParagraphHeading.HEADING2));
  [
    `إجمالي المتدربين: ${s.trainees}`,
    `عدد الكورسات: ${s.courses}`,
    `الشهادات الصادرة: ${s.certs}`,
    `فريق العمل النشط: ${s.staffActive}`,
    `مهام مفتوحة: ${s.tasksOpen} (منها متأخرة: ${s.tasksOverdue})`,
    `مراسلات تحتاج متابعة: ${s.mailOpen}`,
  ].forEach((l) => rtl(body.appendListItem(l)));

  rtl(body.appendParagraph('الكورسات').setHeading(DocumentApp.ParagraphHeading.HEADING2));
  if (s.courseRows.length) {
    const table = body.appendTable([['الكود', 'الكورس', 'المدرب', 'البداية', 'المسجلين', 'الحالة']].concat(
      s.courseRows.map((r) => r.map(String))));
    table.getRow(0).editAsText().setBold(true);
  }
  rtl(body.appendParagraph('المتدربين حسب الحالة').setHeading(DocumentApp.ParagraphHeading.HEADING2));
  Object.entries(s.traineesByStatus).forEach(([k, v]) => rtl(body.appendListItem(`${k}: ${v}`)));
  doc.saveAndClose();

  // تحويل إلى Word عبر رابط التصدير
  const url = `https://docs.google.com/document/d/${doc.getId()}/export?format=docx`;
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName(title + '.docx');
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const docx = folder.createFile(blob);

  if (sendMail === true) {
    const to = getSetting('بريد المدير');
    if (to) MailApp.sendEmail({ to, subject: title, body: 'التقرير الشهري مرفق.', attachments: [blob] });
  } else {
    toast_('تم إنشاء التقرير ✅ ' + docx.getUrl());
  }
  return docx.getUrl();
}
