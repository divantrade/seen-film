/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف الدوال المساعدة
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على الدوال المساعدة المستخدمة في جميع أجزاء النظام
 */

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التاريخ والوقت
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على التاريخ والوقت الحالي بتنسيق معين
 * @param {string} format تنسيق التاريخ (اختياري)
 * @returns {string} التاريخ والوقت بالتنسيق المطلوب
 */
function getCurrentDateTime(format) {
  const now = new Date();
  const timezone = CONFIG.TIMEZONE;
  format = format || CONFIG.DATETIME_FORMAT;
  return Utilities.formatDate(now, timezone, format);
}

/**
 * الحصول على التاريخ الحالي فقط
 * @returns {string} التاريخ الحالي
 */
function getCurrentDate() {
  return getCurrentDateTime(CONFIG.DATE_FORMAT);
}

/**
 * الحصول على الوقت الحالي فقط
 * @returns {string} الوقت الحالي
 */
function getCurrentTime() {
  return getCurrentDateTime(CONFIG.TIME_FORMAT);
}

/**
 * تنسيق تاريخ معين
 * @param {Date} date التاريخ
 * @param {string} format التنسيق المطلوب
 * @returns {string} التاريخ بالتنسيق المطلوب
 */
function formatDate(date, format) {
  if (!date) return '';
  format = format || CONFIG.DATE_FORMAT;
  return Utilities.formatDate(new Date(date), CONFIG.TIMEZONE, format);
}

/**
 * حساب الفرق بين تاريخين بالأيام
 * @param {Date} startDate تاريخ البداية
 * @param {Date} endDate تاريخ النهاية
 * @returns {number} عدد الأيام
 */
function daysBetween(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const diffTime = Math.abs(end - start);
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * حساب الأيام المتبقية من تاريخ معين
 * @param {Date} targetDate التاريخ المستهدف
 * @returns {number} عدد الأيام المتبقية (سالب إذا كان في الماضي)
 */
function daysRemaining(targetDate) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(targetDate);
  target.setHours(0, 0, 0, 0);
  const diffTime = target - today;
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * إضافة أيام لتاريخ معين
 * @param {Date} date التاريخ
 * @param {number} days عدد الأيام المراد إضافتها
 * @returns {Date} التاريخ الجديد
 */
function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إنشاء المعرفات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء معرف فريد جديد
 * @param {string} prefix بادئة المعرف (اختياري)
 * @returns {string} المعرف الفريد
 */
function generateId(prefix) {
  prefix = prefix || '';
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000);
  return `${prefix}${timestamp}${random}`;
}

/**
 * إنشاء معرف تسلسلي للشيت
 * @param {string} sheetName اسم الشيت
 * @param {string} prefix بادئة المعرف
 * @returns {string} المعرف التسلسلي
 */
function generateSequentialId(sheetName, prefix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return prefix + '001';

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + '001';

  // الحصول على آخر معرف
  const lastId = sheet.getRange(lastRow, 1).getValue();
  if (!lastId) return prefix + '001';

  // استخراج الرقم وزيادته
  const numPart = parseInt(lastId.toString().replace(prefix, ''), 10);
  const newNum = isNaN(numPart) ? 1 : numPart + 1;

  return prefix + newNum.toString().padStart(3, '0');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الشيتات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على شيت بالاسم
 * @param {string} sheetName اسم الشيت
 * @returns {Sheet} الشيت
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * الحصول على جميع البيانات من شيت (بدون الهيدر)
 * @param {string} sheetName اسم الشيت
 * @returns {Array} مصفوفة البيانات
 */
function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1 || lastCol <= 0) return [];

  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

/**
 * الحصول على البيانات مع الهيدر ككائنات
 * @param {string} sheetName اسم الشيت
 * @returns {Array} مصفوفة من الكائنات
 */
function getSheetDataAsObjects(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1 || lastCol <= 0) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return data.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

/**
 * إضافة صف جديد لشيت
 * @param {string} sheetName اسم الشيت
 * @param {Array} rowData بيانات الصف
 */
function appendRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  if (sheet) {
    sheet.appendRow(rowData);
  }
}

/**
 * تحديث صف في شيت
 * @param {string} sheetName اسم الشيت
 * @param {number} rowIndex رقم الصف (يبدأ من 1)
 * @param {Array} rowData بيانات الصف
 */
function updateRow(sheetName, rowIndex, rowData) {
  const sheet = getSheet(sheetName);
  if (sheet && rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  }
}

/**
 * حذف صف من شيت
 * @param {string} sheetName اسم الشيت
 * @param {number} rowIndex رقم الصف
 */
function deleteRow(sheetName, rowIndex) {
  const sheet = getSheet(sheetName);
  if (sheet && rowIndex > 1) {
    sheet.deleteRow(rowIndex);
  }
}

/**
 * البحث عن صف بقيمة معينة في عمود محدد
 * @param {string} sheetName اسم الشيت
 * @param {number} columnIndex رقم العمود (يبدأ من 1)
 * @param {*} value القيمة المراد البحث عنها
 * @returns {number} رقم الصف أو -1 إذا لم يوجد
 */
function findRowByValue(sheetName, columnIndex, value) {
  const sheet = getSheet(sheetName);
  if (!sheet) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const columnData = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();

  for (let i = 0; i < columnData.length; i++) {
    if (columnData[i][0] == value) {
      return i + 2; // +2 لأننا بدأنا من الصف 2 والمصفوفة تبدأ من 0
    }
  }

  return -1;
}

/**
 * الحصول على قيم فريدة من عمود معين
 * @param {string} sheetName اسم الشيت
 * @param {number} columnIndex رقم العمود
 * @returns {Array} قائمة القيم الفريدة
 */
function getUniqueValues(sheetName, columnIndex) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const columnData = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();
  const uniqueValues = [...new Set(columnData.map(row => row[0]).filter(val => val !== ''))];

  return uniqueValues;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التنسيق والعرض
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تطبيق تلوين حسب الحالة على نطاق
 * @param {Range} range النطاق
 * @param {string} status الحالة
 */
function applyStatusColor(range, status) {
  const color = getStatusColor(status);
  range.setBackground(color);
}

/**
 * تنسيق رقم بفواصل الآلاف
 * @param {number} num الرقم
 * @returns {string} الرقم المنسق
 */
function formatNumber(num) {
  if (num === null || num === undefined) return '0';
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * تحويل الدقائق إلى صيغة ساعات:دقائق
 * @param {number} minutes عدد الدقائق
 * @returns {string} الوقت بصيغة HH:MM
 */
function minutesToHoursMinutes(minutes) {
  if (!minutes || minutes <= 0) return '00:00';
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}`;
}

/**
 * تحويل الثواني إلى صيغة دقائق:ثواني
 * @param {number} seconds عدد الثواني
 * @returns {string} الوقت بصيغة MM:SS
 */
function secondsToMinutesSeconds(seconds) {
  if (!seconds || seconds <= 0) return '00:00';
  const mins = Math.floor(seconds / 60);
  const secs = seconds % 60;
  return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التحقق والتنظيف
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * التحقق من صحة البريد الإلكتروني
 * @param {string} email البريد الإلكتروني
 * @returns {boolean} صحيح إذا كان البريد صالحاً
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * التحقق من صحة رقم الهاتف
 * @param {string} phone رقم الهاتف
 * @returns {boolean} صحيح إذا كان الرقم صالحاً
 */
function isValidPhone(phone) {
  if (!phone) return false;
  const phoneRegex = /^[\d\s\-\+\(\)]+$/;
  return phoneRegex.test(phone) && phone.replace(/\D/g, '').length >= 7;
}

/**
 * تنظيف النص من المسافات الزائدة
 * @param {string} text النص
 * @returns {string} النص النظيف
 */
function cleanText(text) {
  if (!text) return '';
  return text.toString().trim().replace(/\s+/g, ' ');
}

/**
 * التحقق من أن القيمة ليست فارغة
 * @param {*} value القيمة
 * @returns {boolean} صحيح إذا كانت القيمة ليست فارغة
 */
function isNotEmpty(value) {
  if (value === null || value === undefined) return false;
  if (typeof value === 'string') return value.trim() !== '';
  return true;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإشعارات والرسائل
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * عرض رسالة نجاح
 * @param {string} message الرسالة
 * @param {string} title العنوان (اختياري)
 */
function showSuccess(message, title) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title || 'نجاح ✅', message, ui.ButtonSet.OK);
}

/**
 * عرض رسالة خطأ
 * @param {string} message الرسالة
 * @param {string} title العنوان (اختياري)
 */
function showError(message, title) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title || 'خطأ ❌', message, ui.ButtonSet.OK);
}

/**
 * عرض رسالة تحذير
 * @param {string} message الرسالة
 * @param {string} title العنوان (اختياري)
 */
function showWarning(message, title) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title || 'تحذير ⚠️', message, ui.ButtonSet.OK);
}

/**
 * عرض رسالة معلومات
 * @param {string} message الرسالة
 * @param {string} title العنوان (اختياري)
 */
function showInfo(message, title) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title || 'معلومات ℹ️', message, ui.ButtonSet.OK);
}

/**
 * عرض سؤال تأكيد
 * @param {string} message الرسالة
 * @param {string} title العنوان (اختياري)
 * @returns {boolean} صحيح إذا اختار المستخدم نعم
 */
function confirmAction(message, title) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    title || 'تأكيد',
    message,
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}

/**
 * عرض إشعار عائم (Toast)
 * @param {string} message الرسالة
 * @param {string} title العنوان (اختياري)
 * @param {number} timeout مدة العرض بالثواني (اختياري)
 */
function showToast(message, title, timeout) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast(message, title || '', timeout || 3);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على بريد المستخدم الحالي
 * @returns {string} البريد الإلكتروني
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * الحصول على اسم المستخدم الحالي (من البريد)
 * @returns {string} اسم المستخدم
 */
function getCurrentUserName() {
  const email = getCurrentUserEmail();
  if (!email) return 'مجهول';
  return email.split('@')[0];
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال القوائم المنسدلة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء قائمة منسدلة من قائمة قيم
 * @param {Range} range النطاق
 * @param {Array} values قائمة القيم
 */
function createDropdown(range, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

/**
 * إنشاء قائمة منسدلة من نطاق مسمى
 * @param {Range} range النطاق
 * @param {string} namedRange اسم النطاق المسمى
 */
function createDropdownFromRange(range, namedRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceRange = ss.getRangeByName(namedRange);

  if (sourceRange) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    range.setDataValidation(rule);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التصدير
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تصدير شيت كـ PDF
 * @param {string} sheetName اسم الشيت
 * @param {string} fileName اسم الملف
 * @returns {string} رابط الملف
 */
function exportSheetAsPDF(sheetName, fileName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('الشيت غير موجود: ' + sheetName);
  }

  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();

  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
    `format=pdf&gid=${sheetId}&portrait=false&fitw=true&gridlines=false`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  const blob = response.getBlob().setName(fileName + '.pdf');
  const file = DriveApp.createFile(blob);

  return file.getUrl();
}

/**
 * تسجيل عملية تصدير
 * @param {string} reportType نوع التقرير
 * @param {string} period الفترة
 * @param {string} fileUrl رابط الملف
 * @param {string} notes ملاحظات
 */
function logExport(reportType, period, fileUrl, notes) {
  const rowData = [
    generateSequentialId(SHEETS.EXPORT_LOG, 'EXP'),
    getCurrentDateTime(),
    reportType,
    period,
    getCurrentUserEmail(),
    fileUrl,
    notes || ''
  ];

  appendRow(SHEETS.EXPORT_LOG, rowData);
}
