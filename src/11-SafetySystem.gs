/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * نظام الأمان والحماية (Safety System)
 * ═══════════════════════════════════════════════════════════════════════════════
 * الإصدار: 1.0.0
 * يشمل:
 *   1. معالجة الأخطاء الشاملة
 *   2. النسخ الاحتياطي التلقائي
 *   3. سجل التدقيق للتغييرات
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// الإعدادات
// ═══════════════════════════════════════════════════════════════════════════════

const SAFETY_CONFIG = {
  // شيت سجل التدقيق
  AUDIT_SHEET_NAME: 'سجل التغييرات',

  // الحد الأقصى لسجلات التدقيق (لتجنب امتلاء الشيت)
  MAX_AUDIT_ROWS: 5000,

  // اسم فولدر النسخ الاحتياطية
  BACKUP_FOLDER_NAME: 'Backups',

  // الاحتفاظ بالنسخ الاحتياطية (بالأيام)
  BACKUP_RETENTION_DAYS: 30
};

// هيدرات شيت سجل التدقيق
const AUDIT_HEADERS = [
  'التاريخ والوقت',
  'المستخدم',
  'الشيت',
  'الصف',
  'العمود',
  'اسم العمود',
  'القيمة السابقة',
  'القيمة الجديدة',
  'نوع العملية'
];

// ═══════════════════════════════════════════════════════════════════════════════
// 1. نظام معالجة الأخطاء الشاملة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تنفيذ دالة مع معالجة أخطاء شاملة
 * @param {Function} fn - الدالة المراد تنفيذها
 * @param {string} operationName - اسم العملية للعرض
 * @param {boolean} showSuccessMessage - عرض رسالة نجاح
 * @returns {*} نتيجة الدالة أو null في حالة الخطأ
 */
function safeExecute(fn, operationName, showSuccessMessage = false) {
  try {
    const result = fn();
    if (showSuccessMessage) {
      showSuccess('✅ تمت العملية: ' + operationName);
    }
    return result;
  } catch (error) {
    handleError(error, operationName);
    return null;
  }
}

/**
 * معالجة الأخطاء وعرض رسالة مناسبة للمستخدم
 * @param {Error} error - كائن الخطأ
 * @param {string} operationName - اسم العملية
 */
function handleError(error, operationName) {
  // تسجيل الخطأ في الـ console
  console.error('═══ خطأ في: ' + operationName + ' ═══');
  console.error('الرسالة:', error.message);
  console.error('Stack:', error.stack);

  // تحديد نوع الخطأ ورسالة مناسبة
  let userMessage = '';
  let errorType = 'عام';

  if (error.message.includes('permission') || error.message.includes('access')) {
    userMessage = '🔒 لا تملك صلاحية للوصول. تحقق من صلاحيات Google Drive.';
    errorType = 'صلاحيات';
  } else if (error.message.includes('not found') || error.message.includes('غير موجود')) {
    userMessage = '🔍 العنصر المطلوب غير موجود. تحقق من البيانات.';
    errorType = 'غير موجود';
  } else if (error.message.includes('network') || error.message.includes('timeout')) {
    userMessage = '🌐 مشكلة في الاتصال. تحقق من الإنترنت وحاول مرة أخرى.';
    errorType = 'اتصال';
  } else if (error.message.includes('quota') || error.message.includes('limit')) {
    userMessage = '⏱️ تم تجاوز الحد المسموح. انتظر قليلاً ثم حاول مرة أخرى.';
    errorType = 'حد الاستخدام';
  } else if (error.message.includes('invalid') || error.message.includes('غير صالح')) {
    userMessage = '⚠️ البيانات المدخلة غير صالحة. تحقق من المدخلات.';
    errorType = 'بيانات غير صالحة';
  } else {
    userMessage = '❌ حدث خطأ غير متوقع. حاول مرة أخرى.';
  }

  // عرض الرسالة للمستخدم
  const fullMessage = userMessage + '\n\n' +
    'العملية: ' + operationName + '\n' +
    'التفاصيل: ' + error.message;

  showError(fullMessage);

  // تسجيل الخطأ في سجل التدقيق
  try {
    logAuditEntry({
      action: 'خطأ',
      sheetName: operationName,
      details: errorType + ': ' + error.message
    });
  } catch (e) {
    // تجاهل أخطاء التسجيل لتجنب حلقة لا نهائية
    console.error('فشل تسجيل الخطأ:', e);
  }
}

/**
 * التحقق من صحة البيانات المطلوبة
 * @param {Object} data - البيانات للتحقق
 * @param {Array} requiredFields - الحقول المطلوبة
 * @returns {Object} {isValid: boolean, missingFields: Array}
 */
function validateRequiredFields(data, requiredFields) {
  const missingFields = [];

  for (const field of requiredFields) {
    if (!data[field] || data[field].toString().trim() === '') {
      missingFields.push(field);
    }
  }

  return {
    isValid: missingFields.length === 0,
    missingFields: missingFields
  };
}

/**
 * التحقق من وجود شيت وإرجاع رسالة مناسبة
 * @param {string} sheetName - اسم الشيت
 * @returns {Sheet|null}
 */
function getSheetSafe(sheetName) {
  const sheet = getSheet(sheetName);

  if (!sheet) {
    showError('⚠️ الشيت "' + sheetName + '" غير موجود.\n\n' +
      'الحل: اذهب إلى القائمة ← الإعدادات ← تهيئة النظام');
    return null;
  }

  return sheet;
}

/**
 * التحقق من إعداد الفولدر الرئيسي
 * @returns {Folder|null}
 */
function getMainFolderSafe() {
  const folder = getMainProductionFolder();

  if (!folder) {
    showError('⚠️ لم يتم تحديد فولدر الإنتاج الرئيسي.\n\n' +
      'الحل:\n' +
      '1. اذهب إلى شيت الإعدادات\n' +
      '2. ضع رابط الفولدر في الخلية B3\n' +
      '3. تأكد من صلاحيات المشاركة');
    return null;
  }

  return folder;
}

// ═══════════════════════════════════════════════════════════════════════════════
// 2. نظام النسخ الاحتياطي التلقائي
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء نسخة احتياطية كاملة
 * @param {string} reason - سبب النسخ الاحتياطي
 * @returns {string|null} رابط النسخة الاحتياطية
 */
function createBackup(reason) {
  return safeExecute(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm');
    const backupName = 'نسخة احتياطية - ' + ss.getName() + ' - ' + dateStr;

    // إنشاء نسخة
    const backup = ss.copy(backupName);

    // محاولة نقلها لفولدر النسخ الاحتياطية
    const backupFolder = getOrCreateBackupFolder();
    if (backupFolder) {
      const file = DriveApp.getFileById(backup.getId());
      file.moveTo(backupFolder);
    }

    // تسجيل في سجل التدقيق
    logAuditEntry({
      action: 'نسخ احتياطي',
      sheetName: 'النظام',
      newValue: backupName,
      details: reason || 'نسخ احتياطي يدوي'
    });

    return backup.getUrl();
  }, 'إنشاء نسخة احتياطية');
}

/**
 * إنشاء نسخة احتياطية يدوية من القائمة
 */
function createManualBackup() {
  const url = createBackup('نسخ احتياطي يدوي من القائمة');

  if (url) {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial; direction: rtl; text-align: center; padding: 20px;">
        <h3>✅ تم إنشاء النسخة الاحتياطية بنجاح!</h3>
        <p>يمكنك الوصول إليها من الرابط التالي:</p>
        <a href="${url}" target="_blank" style="background: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; margin: 10px 0;">
          فتح النسخة الاحتياطية
        </a>
        <p style="color: #666; font-size: 12px;">النسخة محفوظة في فولدر Backups</p>
      </div>
    `).setWidth(400).setHeight(200);

    ui.showModalDialog(html, 'نسخة احتياطية');
  }
}

/**
 * الحصول على أو إنشاء فولدر النسخ الاحتياطية
 * @returns {Folder|null}
 */
function getOrCreateBackupFolder() {
  try {
    const mainFolder = getMainProductionFolder();

    if (mainFolder) {
      // البحث عن فولدر Backups
      let backupFolder = findFolderByName(mainFolder, SAFETY_CONFIG.BACKUP_FOLDER_NAME);

      if (!backupFolder) {
        // إنشاء فولدر جديد
        backupFolder = mainFolder.createFolder(SAFETY_CONFIG.BACKUP_FOLDER_NAME);
      }

      return backupFolder;
    }

    // إذا لم يكن هناك فولدر رئيسي، نحفظ في الجذر
    return DriveApp.getRootFolder();

  } catch (e) {
    console.error('خطأ في إنشاء فولدر النسخ الاحتياطية:', e);
    return null;
  }
}

/**
 * نسخ احتياطي تلقائي يومي (يُنفذ عبر Trigger)
 */
function dailyAutoBackup() {
  const url = createBackup('نسخ احتياطي تلقائي يومي');

  if (url) {
    console.log('تم إنشاء النسخة الاحتياطية اليومية:', url);

    // تنظيف النسخ القديمة
    cleanOldBackups();
  }
}

/**
 * حذف النسخ الاحتياطية القديمة
 */
function cleanOldBackups() {
  try {
    const backupFolder = getOrCreateBackupFolder();
    if (!backupFolder) return;

    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - SAFETY_CONFIG.BACKUP_RETENTION_DAYS);

    const files = backupFolder.getFiles();
    let deletedCount = 0;

    while (files.hasNext()) {
      const file = files.next();

      // التحقق من أن الملف نسخة احتياطية وقديم
      if (file.getName().includes('نسخة احتياطية') &&
          file.getDateCreated() < cutoffDate) {
        file.setTrashed(true);
        deletedCount++;
      }
    }

    if (deletedCount > 0) {
      console.log('تم حذف ' + deletedCount + ' نسخة احتياطية قديمة');
    }

  } catch (e) {
    console.error('خطأ في تنظيف النسخ الاحتياطية:', e);
  }
}

/**
 * نسخ احتياطي قبل العمليات الخطيرة
 * @param {string} operationName - اسم العملية
 * @returns {boolean} نجاح إنشاء النسخة
 */
function backupBeforeDangerousOperation(operationName) {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    '⚠️ تحذير - عملية خطيرة',
    'هذه العملية قد تؤدي لفقدان بيانات.\n\n' +
    'العملية: ' + operationName + '\n\n' +
    'هل تريد إنشاء نسخة احتياطية أولاً؟',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (result === ui.Button.CANCEL) {
    return false; // إلغاء العملية
  }

  if (result === ui.Button.YES) {
    const backupUrl = createBackup('قبل: ' + operationName);
    if (!backupUrl) {
      showError('فشل إنشاء النسخة الاحتياطية. تم إلغاء العملية.');
      return false;
    }
    showSuccess('تم إنشاء نسخة احتياطية. جاري المتابعة...');
  }

  return true;
}

// ═══════════════════════════════════════════════════════════════════════════════
// 3. سجل التدقيق للتغييرات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء شيت سجل التدقيق
 */
function createAuditLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SAFETY_CONFIG.AUDIT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SAFETY_CONFIG.AUDIT_SHEET_NAME);

    // تعيين الهيدرات
    const headerRange = sheet.getRange(1, 1, 1, AUDIT_HEADERS.length);
    headerRange.setValues([AUDIT_HEADERS]);

    // تنسيق الهيدر
    headerRange.setBackground('#1565C0')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // تعيين عرض الأعمدة
    sheet.setColumnWidth(1, 150); // التاريخ
    sheet.setColumnWidth(2, 150); // المستخدم
    sheet.setColumnWidth(3, 100); // الشيت
    sheet.setColumnWidth(4, 60);  // الصف
    sheet.setColumnWidth(5, 60);  // العمود
    sheet.setColumnWidth(6, 120); // اسم العمود
    sheet.setColumnWidth(7, 150); // القيمة السابقة
    sheet.setColumnWidth(8, 150); // القيمة الجديدة
    sheet.setColumnWidth(9, 100); // نوع العملية

    // تجميد الصف الأول
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * تسجيل إدخال في سجل التدقيق
 * @param {Object} entry - بيانات الإدخال
 */
function logAuditEntry(entry) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SAFETY_CONFIG.AUDIT_SHEET_NAME);

    // إنشاء الشيت إذا لم يكن موجوداً
    if (!sheet) {
      sheet = createAuditLogSheet();
    }

    // الحصول على معلومات المستخدم
    let userEmail = 'غير معروف';
    try {
      userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'غير معروف';
    } catch (e) {
      userEmail = 'نظام';
    }

    // التاريخ والوقت
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

    // إعداد الصف
    const rowData = [
      timestamp,
      userEmail,
      entry.sheetName || '',
      entry.row || '',
      entry.column || '',
      entry.columnName || '',
      entry.oldValue !== undefined ? String(entry.oldValue) : '',
      entry.newValue !== undefined ? String(entry.newValue) : '',
      entry.action || 'تعديل'
    ];

    // إضافة الصف في الأعلى (بعد الهيدر)
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);

    // تنظيف السجلات القديمة إذا تجاوز الحد
    const lastRow = sheet.getLastRow();
    if (lastRow > SAFETY_CONFIG.MAX_AUDIT_ROWS) {
      sheet.deleteRows(SAFETY_CONFIG.MAX_AUDIT_ROWS + 1, lastRow - SAFETY_CONFIG.MAX_AUDIT_ROWS);
    }

  } catch (e) {
    console.error('خطأ في تسجيل التدقيق:', e);
  }
}

/**
 * تسجيل تغييرات onEdit في سجل التدقيق
 * يُستدعى من onEdit الرئيسية
 * @param {Event} e - حدث التعديل
 */
function logEditChange(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();

    // تجاهل شيت سجل التغييرات نفسه
    if (sheetName === SAFETY_CONFIG.AUDIT_SHEET_NAME) return;

    // تجاهل الشيتات المخفية والداشبورد
    if (sheetName === SHEETS.FOLDER_LINKS) return;

    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    // تجاهل الهيدر
    if (row === 1) return;

    // الحصول على اسم العمود
    let columnName = '';
    try {
      columnName = sheet.getRange(1, col).getValue() || '';
    } catch (err) {
      columnName = 'عمود ' + col;
    }

    // تسجيل التغيير
    logAuditEntry({
      sheetName: sheetName,
      row: row,
      column: col,
      columnName: columnName,
      oldValue: e.oldValue,
      newValue: e.value,
      action: 'تعديل'
    });

  } catch (e) {
    console.error('خطأ في تسجيل التعديل:', e);
  }
}

/**
 * تسجيل عملية إضافة
 * @param {string} sheetName - اسم الشيت
 * @param {string} itemName - اسم العنصر المضاف
 */
function logAddition(sheetName, itemName) {
  logAuditEntry({
    sheetName: sheetName,
    newValue: itemName,
    action: 'إضافة'
  });
}

/**
 * تسجيل عملية حذف
 * @param {string} sheetName - اسم الشيت
 * @param {string} itemName - اسم العنصر المحذوف
 */
function logDeletion(sheetName, itemName) {
  logAuditEntry({
    sheetName: sheetName,
    oldValue: itemName,
    action: 'حذف'
  });
}

/**
 * عرض سجل التدقيق
 */
function showAuditLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SAFETY_CONFIG.AUDIT_SHEET_NAME);

  if (!sheet) {
    sheet = createAuditLogSheet();
    showInfo('تم إنشاء سجل التغييرات. لا توجد سجلات بعد.');
  }

  // الانتقال للشيت
  ss.setActiveSheet(sheet);
  showSuccess('تم فتح سجل التغييرات');
}

/**
 * تصدير سجل التدقيق كملف
 */
function exportAuditLog() {
  return safeExecute(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SAFETY_CONFIG.AUDIT_SHEET_NAME);

    if (!sheet || sheet.getLastRow() <= 1) {
      showInfo('سجل التغييرات فارغ.');
      return null;
    }

    const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const exportName = 'سجل التدقيق - ' + dateStr;

    // إنشاء ملف جديد
    const newSS = SpreadsheetApp.create(exportName);
    const newSheet = newSS.getActiveSheet();

    // نسخ البيانات
    const data = sheet.getDataRange().getValues();
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // تنسيق
    newSheet.setFrozenRows(1);
    newSheet.getRange(1, 1, 1, data[0].length)
      .setBackground('#1565C0')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    return newSS.getUrl();
  }, 'تصدير سجل التدقيق');
}

// ═══════════════════════════════════════════════════════════════════════════════
// 4. تثبيت Triggers للنظام الآمن
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تثبيت جميع Triggers للنظام الآمن
 */
function installSafetyTriggers() {
  const ss = SpreadsheetApp.getActive();

  // حذف Triggers القديمة للنسخ الاحتياطي
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'dailyAutoBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // إضافة trigger للنسخ الاحتياطي اليومي (الساعة 3 صباحاً)
  ScriptApp.newTrigger('dailyAutoBackup')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  showSuccess('تم تثبيت نظام الأمان:\n• نسخ احتياطي يومي الساعة 3 صباحاً\n• سجل التغييرات مفعّل');
}

// ═══════════════════════════════════════════════════════════════════════════════
// 5. دوال مساعدة للتكامل
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * نسخة آمنة من resetSystem مع نسخ احتياطي
 */
function resetSystemSafe() {
  // التحقق والنسخ الاحتياطي
  if (!backupBeforeDangerousOperation('إعادة تهيئة النظام')) {
    return;
  }

  // تسجيل العملية
  logAuditEntry({
    sheetName: 'النظام',
    action: 'إعادة تهيئة',
    details: 'بدء إعادة تهيئة النظام'
  });

  // تنفيذ العملية الأصلية
  resetSystem();
}

/**
 * ملخص حالة النظام الآمن
 */
function showSafetyStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let status = '═══════════════════════════════════\n';
  status += '📊 حالة نظام الأمان\n';
  status += '═══════════════════════════════════\n\n';

  // سجل التدقيق
  const auditSheet = ss.getSheetByName(SAFETY_CONFIG.AUDIT_SHEET_NAME);
  if (auditSheet) {
    const auditRows = auditSheet.getLastRow() - 1;
    status += '✅ سجل التغييرات: مفعّل (' + auditRows + ' سجل)\n';
  } else {
    status += '⚠️ سجل التغييرات: غير مفعّل\n';
  }

  // النسخ الاحتياطي
  const backupFolder = getOrCreateBackupFolder();
  if (backupFolder) {
    const files = backupFolder.getFiles();
    let backupCount = 0;
    while (files.hasNext()) {
      if (files.next().getName().includes('نسخة احتياطية')) {
        backupCount++;
      }
    }
    status += '✅ النسخ الاحتياطية: ' + backupCount + ' نسخة\n';
  } else {
    status += '⚠️ فولدر النسخ الاحتياطية: غير محدد\n';
  }

  // Triggers
  const triggers = ScriptApp.getProjectTriggers();
  const hasBackupTrigger = triggers.some(t => t.getHandlerFunction() === 'dailyAutoBackup');
  status += hasBackupTrigger ?
    '✅ النسخ التلقائي: مفعّل\n' :
    '⚠️ النسخ التلقائي: غير مفعّل\n';

  status += '\n═══════════════════════════════════';

  ui.alert('حالة نظام الأمان', status, ui.ButtonSet.OK);
}
