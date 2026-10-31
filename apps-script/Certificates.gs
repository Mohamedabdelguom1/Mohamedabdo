/**
 * Certificates.gs — إصدار الشهادات تلقائياً من قالب Google Docs إلى PDF.
 *
 * القالب: ملف Docs عادي فيه نصوص بين أقواس معقوفة تُستبدل تلقائياً:
 *   {{الاسم}} {{الكورس}} {{المدرب}} {{التاريخ}} {{رقم_الشهادة}} {{المركز}}
 */

function issuePendingCertificates() {
  const templateId = getSetting('معرف قالب الشهادة');
  const folderId = getSetting('معرف مجلد الشهادات');
  if (!templateId || !folderId) throw new Error('اكتب "معرف قالب الشهادة" و"معرف مجلد الشهادات" في الإعدادات');

  const template = DriveApp.getFileById(templateId);
  const folder = DriveApp.getFolderById(folderId);
  const center = getSetting('اسم المركز');

  const courses = {};
  readTable_(SHEETS.COURSES).rows.forEach((c) => (courses[String(c['كود الكورس'])] = c));
  const issued = new Set(readTable_(SHEETS.CERTS).rows.map((c) => String(c['رقم المتدرب']) + '|' + String(c['كود الكورس'])));
  const pending = readTable_(SHEETS.TRAINEES).rows.filter(
    (t) => t['الحالة'] === 'مكتمل' && !issued.has(String(t['رقم المتدرب']) + '|' + String(t['كود الكورس']))
  );

  let count = 0;
  pending.forEach((t) => {
    const course = courses[String(t['كود الكورس'])] || {};
    const certId = nextId_(SHEETS.CERTS, 'رقم الشهادة', 'CERT');
    const today = new Date();
    const copy = template.makeCopy(`شهادة - ${t['الاسم']} - ${certId}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    const map = {
      '{{الاسم}}': t['الاسم'],
      '{{الكورس}}': course['اسم الكورس'] || t['كود الكورس'],
      '{{المدرب}}': course['المدرب'] || '',
      '{{التاريخ}}': fmt_(today),
      '{{رقم_الشهادة}}': certId,
      '{{المركز}}': center,
    };
    [doc.getBody(), doc.getHeader(), doc.getFooter()].forEach((part) => {
      if (!part) return;
      Object.keys(map).forEach((k) => part.replaceText(k.replace(/[{}]/g, '\\$&'), String(map[k] || '')));
    });
    doc.saveAndClose();
    const pdf = folder.createFile(copy.getAs(MimeType.PDF)).setName(`${certId} - ${t['الاسم']}.pdf`);
    copy.setTrashed(true);

    appendObject_(SHEETS.CERTS, {
      'رقم الشهادة': certId,
      'رقم المتدرب': t['رقم المتدرب'],
      'الاسم': t['الاسم'],
      'كود الكورس': t['كود الكورس'],
      'تاريخ الإصدار': today,
      'رابط PDF': pdf.getUrl(),
      'الحالة': 'جاهزة',
    });
    count++;
  });
  toast_(count ? `تم إصدار ${count} شهادة ✅` : 'ما في متدربين مكتملين بدون شهادة');
}
