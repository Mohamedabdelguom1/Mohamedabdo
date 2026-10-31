/**
 * Reminders.gs — التذكير اليومي للمدير بالبريد.
 * يجمع: المهام المتأخرة، مواعيد اليوم، القادمة خلال N أيام، كورسات تبدأ/تنتهي قريباً،
 * مراسلات محتاجة متابعة، ومتدربين مكتملين بدون شهادة.
 */

function sendDailyReminder() {
  const to = getSetting('بريد المدير');
  if (!to) throw new Error('اكتب بريد المدير في ورقة الإعدادات');
  const ahead = Number(getSetting('أيام التنبيه المسبق')) || 3;
  const today = new Date();
  const center = getSetting('اسم المركز') || 'المركز';

  const tasks = readTable_(SHEETS.TASKS);
  const overdue = [], todayList = [], upcoming = [];
  tasks.rows.forEach((r) => {
    if (r['الحالة'] === 'منجزة') return;
    const d = toDate_(r['التاريخ']);
    if (!d) return;
    const diff = daysBetween_(today, d);
    if (diff < 0) overdue.push(r);
    else if (diff === 0) todayList.push(r);
    else if (diff <= ahead) upcoming.push(r);
  });

  const courses = readTable_(SHEETS.COURSES).rows;
  const courseAlerts = [];
  courses.forEach((c) => {
    if (c['الحالة'] === 'ملغي') return;
    const s = toDate_(c['تاريخ البداية']);
    const e = toDate_(c['تاريخ النهاية']);
    if (s && daysBetween_(today, s) >= 0 && daysBetween_(today, s) <= ahead)
      courseAlerts.push(`يبدأ "${c['اسم الكورس']}" يوم ${fmt_(s)} (المسجلين: ${c['عدد المسجلين'] || 0}/${c['السعة'] || '-'})`);
    if (e && daysBetween_(today, e) >= 0 && daysBetween_(today, e) <= ahead)
      courseAlerts.push(`ينتهي "${c['اسم الكورس']}" يوم ${fmt_(e)} — جهّز الشهادات`);
  });

  const mailFollow = readTable_(SHEETS.MAIL).rows.filter((m) => {
    if (['تم الرد', 'مؤرشفة'].includes(m['الحالة'])) return false;
    const f = toDate_(m['متابعة بتاريخ']);
    return m['الحالة'] === 'جديدة' || (f && daysBetween_(today, f) <= 0);
  });

  const certifiedIds = new Set(readTable_(SHEETS.CERTS).rows.map((c) => String(c['رقم المتدرب'])));
  const noCert = readTable_(SHEETS.TRAINEES).rows.filter(
    (t) => t['الحالة'] === 'مكتمل' && !certifiedIds.has(String(t['رقم المتدرب']))
  );

  const total = overdue.length + todayList.length + upcoming.length + courseAlerts.length + mailFollow.length + noCert.length;
  if (total === 0) return; // ما في شي — ما نزعج المدير

  const taskLine = (r) =>
    `<li><b>${escapeHtml_(r['العنوان'])}</b> — ${fmt_(toDate_(r['التاريخ']))} ${escapeHtml_(r['الوقت'] || '')}` +
    ` | ${escapeHtml_(r['المسؤول'] || '')} | أولوية: ${escapeHtml_(r['الأولوية'] || '-')}</li>`;
  const section = (title, color, items) =>
    items.length ? `<h3 style="color:${color};margin:16px 0 4px">${title} (${items.length})</h3><ul>${items.join('')}</ul>` : '';

  const html =
    `<div dir="rtl" style="font-family:Tahoma,Arial;font-size:14px">` +
    `<h2>صباح الخير 👋 — ملخص ${escapeHtml_(center)} ليوم ${fmt_(today)}</h2>` +
    section('🔴 متأخرة', '#c0392b', overdue.map(taskLine)) +
    section('🟡 اليوم', '#b9770e', todayList.map(taskLine)) +
    section(`🔵 خلال ${ahead} أيام`, '#1f4e78', upcoming.map(taskLine)) +
    section('📚 الكورسات', '#1f4e78', courseAlerts.map((s) => `<li>${escapeHtml_(s)}</li>`)) +
    section('✉️ مراسلات تحتاج متابعة', '#6c3483', mailFollow.map((m) =>
      `<li>${escapeHtml_(m['الموضوع'])} — ${escapeHtml_(m['الجهة'])} (${escapeHtml_(m['الحالة'])})</li>`)) +
    section('🎓 مكتملين بدون شهادة', '#117a65', noCert.map((t) =>
      `<li>${escapeHtml_(t['الاسم'])} — ${escapeHtml_(t['كود الكورس'])}</li>`)) +
    `<p style="color:#888">رابط الملف: ${SpreadsheetApp.getActive().getUrl()}</p></div>`;

  MailApp.sendEmail({
    to,
    subject: `🔔 ${center}: ${overdue.length} متأخرة، ${todayList.length} اليوم، ${upcoming.length} قادمة`,
    htmlBody: html,
  });

  const stamp = new Date();
  overdue.concat(todayList, upcoming).forEach((r) => setCell_(tasks, r._row, 'آخر تذكير', stamp));
}
