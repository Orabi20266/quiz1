/**
 * خدمة استقبال نتائج اختبار التجويد (اختياري)
 * تسجّل كل نتيجة في جدول Google وتُرسل نسخة إلى بريد المعلم تلقائياً.
 *
 * خطوات التشغيل مرة واحدة فقط:
 * 1) افتح https://sheets.new لإنشاء جدول Google جديد.
 * 2) من القائمة: Extensions ← Apps Script.
 * 3) امسح الكود الموجود والصق هذا الملف كاملاً.
 * 4) عدّل TEACHER_EMAIL إلى بريد المعلم.
 * 5) Deploy ← New deployment ← Web app
 *      - Description: tajweed-results
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    ثم Authorize واختر حسابك واسمح بالصلاحيات.
 * 6) انسخ رابط النشر الذي ينتهي بـ /exec.
 * 7) ضع الرابط في index.html داخل:
 *    RESULT_CONFIG.appsScriptUrl = "الرابط هنا";
 */

const TEACHER_EMAIL = 'teacher@example.com'; // ← ضع بريد المعلم هنا
const SHEET_NAME = 'النتائج';

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'التاريخ',
        'اسم الطالب',
        'بريد الطالب',
        'الدرجة',
        'الوقت المستغرق',
        'حالة التسليم',
        'التقرير الكامل'
      ]);
      sheet.setRightToLeft(true);
      sheet.getRange(1, 1, 1, 7)
        .setFontWeight('bold')
        .setBackground('#1b6a45')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data['تاريخ الاختبار'] || new Date(),
      data['اسم الطالب'] || '',
      data['بريد الطالب'] || '',
      data['الدرجة'] || '',
      data['الوقت المستغرق'] || '',
      data['حالة التسليم'] || '',
      data['تقرير الأسئلة والتوثيق'] || ''
    ]);

    MailApp.sendEmail({
      to: TEACHER_EMAIL,
      subject: data.subject || 'نتيجة اختبار التجويد',
      body: data['تقرير الأسئلة والتوثيق'] || JSON.stringify(data, null, 2)
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// فحص سريع من المتصفح بعد النشر
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'tajweed-results' }))
    .setMimeType(ContentService.MimeType.JSON);
}
