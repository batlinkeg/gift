// 1. اذهب إلى Google Sheets وأنشئ ملفاً جديداً.
// 2. من القائمة العلوية، اختر Extensions (الإضافات) ثم Apps Script (برمجة التطبيقات).
// 3. امسح أي كود موجود والصق هذا الكود بالكامل:

var sheetName = 'Sheet1'; // تأكد أن اسم الشيت في الأسفل هو Sheet1 (أو الورقة 1)
var scriptProp = PropertiesService.getScriptProperties();

function initialSetup () {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', activeSpreadsheet.getId());
}

function doPost (e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    var sheet = doc.getSheetByName(sheetName);

    // هذه هي البيانات التي يتم إرسالها من موقعك بالإضافة إلى وقت التسجيل
    var headers = ['Name', 'Phone', 'Address', 'Age', 'Status', 'Location', 'Timestamp'];
    
    // إذا كان الشيت فارغاً، سيتم إضافة أسماء الأعمدة في الصف الأول
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    var nextRow = sheet.getLastRow() + 1;

    // تجهيز الصف الجديد بناءً على البيانات القادمة
    var newRow = headers.map(function(header) {
      if (header === 'Timestamp') {
        return new Date(); // يضيف تاريخ ووقت التسجيل
      }
      return e.parameter[header] || ''; // يجلب القيمة المرسلة من HTML
    });

    // إضافة الصف في الشيت
    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': error }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  finally {
    lock.releaseLock();
  }
}

/*
--- خطوات التشغيل خطوة بخطوة ---

1. بعد لصق الكود، اضغط على زر الحفظ (أيقونة القرص المرن) أو Ctrl+S.
2. من القائمة المنسدلة بالأعلى التي تحتوي على كلمة (doPost)، قم بتغييرها إلى (initialSetup) ثم اضغط على زر "Run" (تشغيل).
3. سيطلب منك إعطاء صلاحيات (Review Permissions)، وافق عليها كلها (اختر إيميلك -> Advanced -> Go to Untitled project).
4. الآن من أعلى اليمين، اضغط على الزر الأزرق "Deploy" (نشر) ثم اختر "New deployment" (نشر جديد).
5. اضغط على أيقونة الترس ⚙️ بجوار "Select type" واختر "Web app" (تطبيق ويب).
6. في خانة Description اكتب مثلاً "Version 1".
7. في خانة "Execute as" اختر "Me" (أنا).
8. في خانة "Who has access" (من يمكنه الوصول) **مهم جداً**: اختر "Anyone" (أي شخص).
9. اضغط "Deploy".
10. سيظهر لك رابط طويل تحت عنوان "Web app URL". قم بنسخ هذا الرابط.
11. اذهب إلى ملف index.html في موقعك، وفي سطر 249 قم باستبدال:
    'ضع_هنا_رابط_الـ_Web_App_الخاص_بـ_Google_Script'
    بالرابط الذي نسخته.
*/
