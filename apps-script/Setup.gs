/**
 * Setup.gs — تجهيز الأوراق، القائمة العربية، والمشغّلات (Triggers = تشغيل تلقائي بجدول زمني).
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗂️ المساعد الإداري')
    .addItem('1) تجهيز النظام (أول مرة)', 'setupSystem')
    .addItem('2) تفعيل التذكير والتقارير التلقائية', 'installTriggers')
    .addSeparator()
    .addItem('➕ متدرب جديد', 'showTraineeForm')
    .addItem('➕ مهمة / موعد جديد', 'showTaskForm')
    .addSeparator()
    .addItem('🎓 إصدار شهادات المكتملين', 'issuePendingCertificates')
    .addItem('🔔 إرسال تذكير الآن', 'sendDailyReminder')
    .addItem('📊 تحديث لوحة المتابعة', 'refreshDashboard')
    .addItem('📄 تقرير شهري (Word)', 'generateMonthlyReport')
    .addItem('📥 سحب المراسلات من Gmail', 'importCorrespondence')
    .addItem('🧹 فحص وتنظيف البيانات', 'cleanData')
    .addToUi();
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach((name) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sh.setRightToLeft(true);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center');

    // القوائم المنسدلة
    const dd = DROPDOWNS[name] || {};
    Object.keys(dd).forEach((col) => {
      const c = headers.indexOf(col) + 1;
      if (c > 0) {
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(dd[col], true).setAllowInvalid(false).build();
        sh.getRange(2, c, sh.getMaxRows() - 1, 1).setDataValidation(rule);
      }
    });

    // أعمدة التواريخ
    headers.forEach((h, i) => {
      if (DATE_COLUMNS.includes(h)) {
        const rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
        sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd').setDataValidation(rule);
      }
    });
    sh.autoResizeColumns(1, headers.length);
  });

  // عدد المسجلين في كل كورس — معادلة تلقائية
  const courses = ss.getSheetByName(SHEETS.COURSES);
  const cH = HEADERS[SHEETS.COURSES];
  const tH = HEADERS[SHEETS.TRAINEES];
  const countCol = cH.indexOf('عدد المسجلين') + 1;
  const codeColL = columnLetter_(cH.indexOf('كود الكورس') + 1);
  const tCodeColL = columnLetter_(tH.indexOf('كود الكورس') + 1);
  courses.getRange(2, countCol).setFormula(
    `=ARRAYFORMULA(IF(${codeColL}2:${codeColL}="",,COUNTIF('${SHEETS.TRAINEES}'!${tCodeColL}2:${tCodeColL},${codeColL}2:${codeColL})))`
  );

  // إعدادات افتراضية
  const st = ss.getSheetByName(SHEETS.SETTINGS);
  if (st.getLastRow() < 2) {
    st.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    st.getRange(2, 2).setValue(ss.getName());
    st.getRange(3, 2).setValue(Session.getEffectiveUser().getEmail());
  }

  // تلوين المهام المتأخرة بالأحمر
  addOverdueFormatting_();

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  refreshDashboard();
  toast_('تم تجهيز النظام ✅ — راجع ورقة "الإعدادات" ثم فعّل التذكير من القائمة.');
}

function addOverdueFormatting_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TASKS);
  const h = HEADERS[SHEETS.TASKS];
  const dL = columnLetter_(h.indexOf('التاريخ') + 1);
  const sL = columnLetter_(h.indexOf('الحالة') + 1);
  const range = sh.getRange(2, 1, sh.getMaxRows() - 1, h.length);
  const overdue = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${dL}2<>"",$${dL}2<TODAY(),$${sL}2<>"منجزة")`)
    .setBackground('#f8d7da').setRanges([range]).build();
  const today = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${dL}2=TODAY(),$${sL}2<>"منجزة")`)
    .setBackground('#fff3cd').setRanges([range]).build();
  sh.setConditionalFormatRules([overdue, today]);
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** يمسح المشغّلات القديمة ويضيف: تذكير يومي، سحب مراسلات كل ساعة، تقرير شهري. */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));
  const hour = Number(getSetting('ساعة التذكير')) || 8;
  ScriptApp.newTrigger('sendDailyReminder').timeBased().everyDays(1).atHour(hour).create();
  ScriptApp.newTrigger('importCorrespondence').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('monthlyJob').timeBased().onMonthDay(1).atHour(hour).create();
  toast_('تم التفعيل ✅ — تذكير يومي الساعة ' + hour + '، سحب مراسلات كل ساعة، وتقرير أول كل شهر.');
}

function monthlyJob() {
  generateMonthlyReport(true);
}
