/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * ملف التهيئة وإنشاء الشيتات
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * تهيئة النظام بالكامل
 * يُنشئ جميع الشيتات المطلوبة مع الهيدرات والتنسيقات
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'تهيئة النظام',
    'سيتم إنشاء الشيتات التالية:\n\n' +
    '• المشاريع\n' +
    '• الفريق\n' +
    '• الحركة\n' +
    '• داشبورد\n' +
    '• الإعدادات\n\n' +
    'هل تريد المتابعة؟',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  try {
    // إنشاء الشيتات
    createProjectsSheet(ss);
    createTeamSheet(ss);
    createMovementSheet(ss);
    createDashboardSheet(ss);
    createSettingsSheet(ss);

    // حذف الشيت الافتراضي إن وجد
    deleteDefaultSheet(ss);

    ui.alert('تم بنجاح', 'تم تهيئة النظام بنجاح!', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('خطأ', 'حدث خطأ أثناء التهيئة:\n' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * إنشاء شيت المشاريع
 */
function createProjectsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.PROJECTS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.PROJECTS);
  }

  // تعيين الهيدرات
  const headerRange = sheet.getRange(1, 1, 1, PROJECT_HEADERS.length);
  headerRange.setValues([PROJECT_HEADERS]);

  // تنسيق الهيدر
  formatHeader(headerRange);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(PROJECT_COLS.CODE, 100);
  sheet.setColumnWidth(PROJECT_COLS.NAME, 200);
  sheet.setColumnWidth(PROJECT_COLS.TYPE, 120);
  sheet.setColumnWidth(PROJECT_COLS.PRODUCER, 120);
  sheet.setColumnWidth(PROJECT_COLS.EDITOR, 120);
  sheet.setColumnWidth(PROJECT_COLS.START_DATE, 110);
  sheet.setColumnWidth(PROJECT_COLS.END_DATE, 110);
  sheet.setColumnWidth(PROJECT_COLS.STATUS, 100);
  sheet.setColumnWidth(PROJECT_COLS.FOLDER_LINK, 250);
  sheet.setColumnWidth(PROJECT_COLS.NOTES, 200);

  // تجميد الصف الأول
  sheet.setFrozenRows(1);

  // إضافة القوائم المنسدلة
  setDropdown(sheet, 2, PROJECT_COLS.TYPE, 100, PROJECT_TYPES);
  setDropdown(sheet, 2, PROJECT_COLS.STATUS, 100, PROJECT_STATUS);

  return sheet;
}

/**
 * إنشاء شيت الفريق
 */
function createTeamSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.TEAM);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.TEAM);
  }

  // تعيين الهيدرات
  const headerRange = sheet.getRange(1, 1, 1, TEAM_HEADERS.length);
  headerRange.setValues([TEAM_HEADERS]);

  // تنسيق الهيدر
  formatHeader(headerRange);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(TEAM_COLS.CODE, 100);
  sheet.setColumnWidth(TEAM_COLS.NAME, 150);
  sheet.setColumnWidth(TEAM_COLS.ROLE, 120);
  sheet.setColumnWidth(TEAM_COLS.EMAIL, 200);
  sheet.setColumnWidth(TEAM_COLS.PHONE, 130);
  sheet.setColumnWidth(TEAM_COLS.STATUS, 100);
  sheet.setColumnWidth(TEAM_COLS.JOIN_DATE, 110);
  sheet.setColumnWidth(TEAM_COLS.NOTES, 200);

  // تجميد الصف الأول
  sheet.setFrozenRows(1);

  // إضافة القوائم المنسدلة
  setDropdown(sheet, 2, TEAM_COLS.ROLE, 100, TEAM_ROLES);
  setDropdown(sheet, 2, TEAM_COLS.STATUS, 100, TEAM_STATUS);

  return sheet;
}

/**
 * إنشاء شيت الحركة
 */
function createMovementSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.MOVEMENT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.MOVEMENT);
  }

  // تعيين الهيدرات
  const headerRange = sheet.getRange(1, 1, 1, MOVEMENT_HEADERS.length);
  headerRange.setValues([MOVEMENT_HEADERS]);

  // تنسيق الهيدر
  formatHeader(headerRange);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(MOVEMENT_COLS.NUMBER, 50);
  sheet.setColumnWidth(MOVEMENT_COLS.DATE, 100);
  sheet.setColumnWidth(MOVEMENT_COLS.PROJECT, 150);
  sheet.setColumnWidth(MOVEMENT_COLS.STAGE, 100);
  sheet.setColumnWidth(MOVEMENT_COLS.SUBTYPE, 120);
  sheet.setColumnWidth(MOVEMENT_COLS.ELEMENT, 150);
  sheet.setColumnWidth(MOVEMENT_COLS.DETAILS, 200);
  sheet.setColumnWidth(MOVEMENT_COLS.ASSIGNED_TO, 120);
  sheet.setColumnWidth(MOVEMENT_COLS.STATUS, 100);
  sheet.setColumnWidth(MOVEMENT_COLS.DUE_DATE, 110);
  sheet.setColumnWidth(MOVEMENT_COLS.LINK, 250);
  sheet.setColumnWidth(MOVEMENT_COLS.NOTES, 200);

  // تجميد الصف الأول
  sheet.setFrozenRows(1);

  // إضافة القوائم المنسدلة للمراحل
  setDropdown(sheet, 2, MOVEMENT_COLS.STAGE, 500, STAGE_NAMES);

  // إضافة القوائم المنسدلة للحالات
  setDropdown(sheet, 2, MOVEMENT_COLS.STATUS, 500, STATUS_WITH_ICONS);

  return sheet;
}

/**
 * إنشاء شيت الداشبورد
 */
function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.DASHBOARD);
  }

  // تنظيف الشيت
  sheet.clear();

  // العنوان الرئيسي
  sheet.getRange('A1').setValue('لوحة التحكم - نظام إدارة الإنتاج');
  sheet.getRange('A1:F1').merge()
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // اختيار الفيلم
  sheet.getRange('A3').setValue('اختر الفيلم:');
  sheet.getRange('B3').setValue('الكل');

  // ملخص
  sheet.getRange('A5').setValue('ملخص المهام');
  sheet.getRange('A5:D5').merge()
    .setBackground(COLORS.INFO)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // جدول الملخص
  sheet.getRange('A6:D6').setValues([['إجمالي', 'تم', 'جاري', 'متأخر']]);
  sheet.getRange('A6:D6').setBackground(COLORS.BACKGROUND).setFontWeight('bold');

  // صيغ الحساب
  sheet.getRange('A7').setFormula('=COUNTA(الحركة!A:A)-1');
  sheet.getRange('B7').setFormula('=COUNTIF(الحركة!I:I,"*تم*")');
  sheet.getRange('C7').setFormula('=COUNTIF(الحركة!I:I,"*جاري*")');
  sheet.getRange('D7').setFormula('=COUNTIF(الحركة!I:I,"*متأخر*")');

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);

  return sheet;
}

/**
 * إنشاء شيت الإعدادات
 */
function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.SETTINGS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.SETTINGS);
  }

  // تنظيف الشيت
  sheet.clear();

  // العنوان
  sheet.getRange('A1').setValue('إعدادات النظام');
  sheet.getRange('A1:C1').merge()
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight('bold');

  // رابط فولدر الإنتاج الرئيسي
  sheet.getRange('A3').setValue('فولدر الإنتاج الرئيسي:');
  sheet.getRange('B3').setValue('(أدخل رابط الفولدر هنا)');
  sheet.getRange('A3').setFontWeight('bold');

  // قسم أنواع المشاريع
  sheet.getRange('A5').setValue('أنواع المشاريع');
  sheet.getRange('A5').setBackground(COLORS.INFO).setFontWeight('bold');

  for (let i = 0; i < PROJECT_TYPES.length; i++) {
    sheet.getRange(6 + i, 1).setValue(PROJECT_TYPES[i]);
  }

  // قسم الأدوار
  sheet.getRange('B5').setValue('أدوار الفريق');
  sheet.getRange('B5').setBackground(COLORS.INFO).setFontWeight('bold');

  for (let i = 0; i < TEAM_ROLES.length; i++) {
    sheet.getRange(6 + i, 2).setValue(TEAM_ROLES[i]);
  }

  // قسم الحالات
  sheet.getRange('C5').setValue('الحالات');
  sheet.getRange('C5').setBackground(COLORS.INFO).setFontWeight('bold');

  const statusList = STATUS_WITH_ICONS;
  for (let i = 0; i < statusList.length; i++) {
    sheet.getRange(6 + i, 3).setValue(statusList[i]);
  }

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);

  return sheet;
}

/**
 * تنسيق صف الهيدر
 */
function formatHeader(range) {
  range.setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  range.getSheet().setRowHeight(1, 35);
}

/**
 * إضافة قائمة منسدلة
 */
function setDropdown(sheet, startRow, column, numRows, values) {
  const range = sheet.getRange(startRow, column, numRows, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

/**
 * حذف الشيت الافتراضي
 */
function deleteDefaultSheet(ss) {
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

/**
 * إعادة تعيين النظام (للتطوير فقط)
 */
function resetSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'تحذير!',
    'سيتم حذف جميع البيانات وإعادة إنشاء الشيتات!\n\nهل أنت متأكد؟',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  // حذف جميع الشيتات
  const sheets = ss.getSheets();

  // إنشاء شيت مؤقت
  const tempSheet = ss.insertSheet('_temp_');

  for (const sheet of sheets) {
    ss.deleteSheet(sheet);
  }

  // إعادة تهيئة النظام
  initializeSystem();

  // حذف الشيت المؤقت
  const temp = ss.getSheetByName('_temp_');
  if (temp) {
    ss.deleteSheet(temp);
  }
}

/**
 * إصلاح شيت الحركة - حذف العمود الزائد وتحديث الهيدرات
 */
function fixMovementSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MOVEMENT);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('شيت الحركة غير موجود!');
    return;
  }

  const currentCols = sheet.getLastColumn();

  // إذا كان عدد الأعمدة أكثر من المطلوب (13 بدلاً من 12)
  if (currentCols > MOVEMENT_HEADERS.length) {
    // حذف العمود 11 (عمود الـ checkbox القديم)
    sheet.deleteColumn(11);
  }

  // تحديث الهيدرات
  const headerRange = sheet.getRange(1, 1, 1, MOVEMENT_HEADERS.length);
  headerRange.setValues([MOVEMENT_HEADERS]);

  // تنسيق الهيدر
  headerRange.setBackground('#1565C0')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(MOVEMENT_COLS.NUMBER, 50);
  sheet.setColumnWidth(MOVEMENT_COLS.DATE, 100);
  sheet.setColumnWidth(MOVEMENT_COLS.PROJECT, 150);
  sheet.setColumnWidth(MOVEMENT_COLS.STAGE, 100);
  sheet.setColumnWidth(MOVEMENT_COLS.SUBTYPE, 120);
  sheet.setColumnWidth(MOVEMENT_COLS.ELEMENT, 150);
  sheet.setColumnWidth(MOVEMENT_COLS.DETAILS, 200);
  sheet.setColumnWidth(MOVEMENT_COLS.ASSIGNED_TO, 120);
  sheet.setColumnWidth(MOVEMENT_COLS.STATUS, 100);
  sheet.setColumnWidth(MOVEMENT_COLS.DUE_DATE, 110);
  sheet.setColumnWidth(MOVEMENT_COLS.LINK, 250);
  sheet.setColumnWidth(MOVEMENT_COLS.NOTES, 200);

  SpreadsheetApp.getActiveSpreadsheet().toast('تم إصلاح شيت الحركة بنجاح!', 'نجاح', 3);
}
