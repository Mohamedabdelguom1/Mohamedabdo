/**
 * Correspondence.gs — تنظيم المراسلات.
 * أي رسالة في Gmail عليها التسمية (Label) المحددة في الإعدادات تُسجّل تلقائياً في ورقة "المراسلات".
 * التصنيف يتم بكلمات مفتاحية بسيطة (بدون ذكاء اصطناعي = مجاني بالكامل).
 */

const CATEGORY_KEYWORDS = {
  'متدربين': ['تسجيل', 'متدرب', 'شهادة', 'كورس', 'دورة', 'حجز'],
  'مدربين': ['مدرب', 'محاضر', 'تدريس'],
  'مالية': ['فاتورة', 'دفع', 'رسوم', 'تحويل', 'إيصال', 'invoice'],
  'شراكات': ['شراكة', 'تعاون', 'رعاية', 'اتفاقية'],
};

function classify_(text) {
  const t = String(text).toLowerCase();
  for (const cat in CATEGORY_KEYWORDS) {
    if (CATEGORY_KEYWORDS[cat].some((k) => t.includes(k))) return cat;
  }
  return 'إدارية';
}

function importCorrespondence() {
  const labelName = getSetting('تسمية Gmail للمراسلات');
  if (!labelName) return;
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) return;

  const table = readTable_(SHEETS.MAIL);
  const seen = new Set(table.rows.map((r) => String(r['معرف Gmail'])));
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  let added = 0;

  label.getThreads(0, 50).forEach((thread) => {
    thread.getMessages().forEach((msg) => {
      const id = msg.getId();
      if (seen.has(id)) return;
      const from = msg.getFrom();
      const outgoing = from.toLowerCase().includes(me);
      appendObject_(SHEETS.MAIL, {
        'رقم': nextId_(SHEETS.MAIL, 'رقم', 'MSG'),
        'التاريخ': msg.getDate(),
        'الجهة': outgoing ? msg.getTo() : from,
        'الاتجاه': outgoing ? 'صادر' : 'وارد',
        'الموضوع': msg.getSubject(),
        'التصنيف': classify_(msg.getSubject() + ' ' + msg.getPlainBody().slice(0, 500)),
        'الحالة': outgoing ? 'تم الرد' : 'جديدة',
        'الرابط': 'https://mail.google.com/mail/u/0/#all/' + thread.getId(),
        'معرف Gmail': id,
      });
      seen.add(id);
      added++;
    });
  });
  if (added) Logger.log('مراسلات جديدة: ' + added);
}
