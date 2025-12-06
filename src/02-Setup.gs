/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف الإعداد وإنشاء الشيتات
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على دوال إنشاء وإعداد جميع الشيتات في النظام
 */

// ═══════════════════════════════════════════════════════════════════════════════
// الدالة الرئيسية لإعداد النظام
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إعداد النظام بالكامل - إنشاء جميع الشيتات
 * هذه الدالة تنشئ جميع الشيتات المطلوبة مع التنسيق
 */
function setupSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'إعداد النظام',
    'سيتم إنشاء جميع الشيتات المطلوبة.\nهل تريد المتابعة؟',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('تم إلغاء العملية');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // إنشاء الشيتات بالترتيب
    createSettingsSheet(ss);
    createProjectsSheet(ss);
    createTeamSheet(ss);
    createPhotographersSheet(ss);
    createGuestsSheet(ss);
    createMovementSheet(ss);
    createVoiceOverSheet(ss);
    createAnimationSheet(ss);
    createArchiveSheet(ss);
    createDashboardSheet(ss);
    createPersonReportSheet(ss);
    createExportLogSheet(ss);

    // حذف الشيت الافتراضي إذا كان موجوداً
    deleteDefaultSheet(ss);

    ui.alert('تم إعداد النظام بنجاح!', 'تم إنشاء جميع الشيتات المطلوبة.', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('خطأ!', 'حدث خطأ أثناء إعداد النظام:\n' + error.message, ui.ButtonSet.OK);
    console.error('Setup Error:', error);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إنشاء الشيتات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء شيت الإعدادات
 * يحتوي على المراحل والحالات
 * @param {Spreadsheet} ss الجدول
 */
function createSettingsSheet(ss) {
  const sheetName = SHEETS.SETTINGS;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // مسح المحتوى الحالي
  sheet.clear();

  // ═══════════════════════════════════════════════════════════════════════════
  // قسم المراحل
  // ═══════════════════════════════════════════════════════════════════════════

  // عنوان المراحل
  sheet.getRange('A1').setValue('المراحل').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A1:D1').merge().setBackground(COLORS.HEADER).setFontColor(COLORS.HEADER_TEXT);

  // رؤوس أعمدة المراحل
  const stageHeaders = ['المعرف', 'الاسم', 'الأيقونة', 'الترتيب'];
  sheet.getRange('A2:D2').setValues([stageHeaders])
    .setFontWeight('bold')
    .setBackground(COLORS.BACKGROUND_LIGHT);

  // بيانات المراحل
  const stageData = Object.values(STAGES).map(stage => [
    stage.id,
    stage.name,
    stage.icon,
    stage.order
  ]);
  sheet.getRange(3, 1, stageData.length, 4).setValues(stageData);

  // ═══════════════════════════════════════════════════════════════════════════
  // قسم الحالات
  // ═══════════════════════════════════════════════════════════════════════════

  const statusStartRow = stageData.length + 5;

  // عنوان الحالات
  sheet.getRange(statusStartRow, 1).setValue('الحالات').setFontWeight('bold').setFontSize(14);
  sheet.getRange(statusStartRow, 1, 1, 4).merge()
    .setBackground(COLORS.HEADER).setFontColor(COLORS.HEADER_TEXT);

  // رؤوس أعمدة الحالات
  const statusHeaders = ['المعرف', 'الاسم', 'الأيقونة', 'اللون'];
  sheet.getRange(statusStartRow + 1, 1, 1, 4).setValues([statusHeaders])
    .setFontWeight('bold')
    .setBackground(COLORS.BACKGROUND_LIGHT);

  // بيانات الحالات مع تلوين الخلايا
  const statusData = Object.values(STATUS);
  statusData.forEach((status, index) => {
    const row = statusStartRow + 2 + index;
    sheet.getRange(row, 1, 1, 4).setValues([[
      status.id,
      status.name,
      status.icon,
      status.color
    ]]);
    // تلوين خلية اللون
    sheet.getRange(row, 4).setBackground(status.color);
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // قسم أنواع المشاريع
  // ═══════════════════════════════════════════════════════════════════════════

  const typesStartRow = statusStartRow + statusData.length + 4;

  sheet.getRange(typesStartRow, 1).setValue('أنواع المشاريع').setFontWeight('bold').setFontSize(14);
  sheet.getRange(typesStartRow, 1, 1, 2).merge()
    .setBackground(COLORS.HEADER).setFontColor(COLORS.HEADER_TEXT);

  PROJECT_TYPES.forEach((type, index) => {
    sheet.getRange(typesStartRow + 1 + index, 1).setValue(type);
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // قسم أدوار الفريق
  // ═══════════════════════════════════════════════════════════════════════════

  const rolesStartRow = typesStartRow;

  sheet.getRange(rolesStartRow, 3).setValue('أدوار الفريق').setFontWeight('bold').setFontSize(14);
  sheet.getRange(rolesStartRow, 3, 1, 2).merge()
    .setBackground(COLORS.HEADER).setFontColor(COLORS.HEADER_TEXT);

  TEAM_ROLES.forEach((role, index) => {
    sheet.getRange(rolesStartRow + 1 + index, 3).setValue(role);
  });

  // تنسيق الشيت
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);

  // تسمية النطاقات للقوائم المنسدلة
  ss.setNamedRange('StagesList', sheet.getRange(3, 2, stageData.length, 1));
  ss.setNamedRange('StatusList', sheet.getRange(statusStartRow + 2, 2, statusData.length, 1));
  ss.setNamedRange('ProjectTypesList', sheet.getRange(typesStartRow + 1, 1, PROJECT_TYPES.length, 1));
  ss.setNamedRange('TeamRolesList', sheet.getRange(rolesStartRow + 1, 3, TEAM_ROLES.length, 1));
}

/**
 * إنشاء شيت المشاريع
 * @param {Spreadsheet} ss الجدول
 */
function createProjectsSheet(ss) {
  const sheetName = SHEETS.PROJECTS;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  // رؤوس الأعمدة
  const headers = [
    'المعرف',
    'اسم المشروع',
    'الكود',
    'النوع',
    'الحالة',
    'تاريخ البداية',
    'تاريخ النهاية',
    'المنتج',
    'المخرج',
    'ملاحظات',
    'تاريخ الإنشاء',
    'آخر تحديث'
  ];

  // تعيين الرؤوس
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // تجميد الصف الأول
  sheet.setFrozenRows(1);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(1, 80);   // المعرف
  sheet.setColumnWidth(2, 200);  // اسم المشروع
  sheet.setColumnWidth(3, 100);  // الكود
  sheet.setColumnWidth(4, 120);  // النوع
  sheet.setColumnWidth(5, 100);  // الحالة
  sheet.setColumnWidth(6, 120);  // تاريخ البداية
  sheet.setColumnWidth(7, 120);  // تاريخ النهاية
  sheet.setColumnWidth(8, 120);  // المنتج
  sheet.setColumnWidth(9, 120);  // المخرج
  sheet.setColumnWidth(10, 200); // ملاحظات
  sheet.setColumnWidth(11, 150); // تاريخ الإنشاء
  sheet.setColumnWidth(12, 150); // آخر تحديث

  // تعيين تنسيق الاتجاه من اليمين لليسار
  sheet.setRightToLeft(true);
}

/**
 * إنشاء شيت الفريق
 * @param {Spreadsheet} ss الجدول
 */
function createTeamSheet(ss) {
  const sheetName = SHEETS.TEAM;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'الاسم',
    'الدور',
    'البريد الإلكتروني',
    'رقم الهاتف',
    'الحالة',
    'تاريخ الانضمام',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 200);
}

/**
 * إنشاء شيت المصورين
 * @param {Spreadsheet} ss الجدول
 */
function createPhotographersSheet(ss) {
  const sheetName = SHEETS.PHOTOGRAPHERS;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'الاسم',
    'التخصص',
    'رقم الهاتف',
    'البريد الإلكتروني',
    'السعر اليومي',
    'الحالة',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 200);
}

/**
 * إنشاء شيت الضيوف
 * @param {Spreadsheet} ss الجدول
 */
function createGuestsSheet(ss) {
  const sheetName = SHEETS.GUESTS;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'الاسم',
    'الصفة',
    'الجهة',
    'رقم الهاتف',
    'البريد الإلكتروني',
    'المشروع',
    'تاريخ المقابلة',
    'الحالة',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 200);
}

/**
 * إنشاء شيت الحركة (الإدخال الرئيسي)
 * @param {Spreadsheet} ss الجدول
 */
function createMovementSheet(ss) {
  const sheetName = SHEETS.MOVEMENT;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'التاريخ',
    'المشروع',
    'المرحلة',
    'المهمة',
    'المسؤول',
    'الحالة',
    'تاريخ البداية',
    'تاريخ النهاية',
    'المدة (دقيقة)',
    'ملاحظات',
    'أنشئ بواسطة',
    'تاريخ الإنشاء'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 200);
  sheet.setColumnWidth(12, 120);
  sheet.setColumnWidth(13, 150);

  // تلوين تبادلي للصفوف
  sheet.getRange('A:M').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * إنشاء شيت التعليق الصوتي
 * @param {Spreadsheet} ss الجدول
 */
function createVoiceOverSheet(ss) {
  const sheetName = SHEETS.VOICEOVER;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'المشروع',
    'الحلقة',
    'النوع',
    'حالة النص',
    'حالة التسجيل',
    'المعلق',
    'المدة (دقيقة)',
    'تاريخ التسجيل',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 200);
}

/**
 * إنشاء شيت الرسوم المتحركة
 * @param {Spreadsheet} ss الجدول
 */
function createAnimationSheet(ss) {
  const sheetName = SHEETS.ANIMATION;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'المشروع',
    'الحلقة',
    'المشهد',
    'النوع',
    'المحرك',
    'الحالة',
    'تاريخ البداية',
    'تاريخ النهاية',
    'المدة (ثانية)',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 200);
}

/**
 * إنشاء شيت الأرشيف
 * @param {Spreadsheet} ss الجدول
 */
function createArchiveSheet(ss) {
  const sheetName = SHEETS.ARCHIVE;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'المشروع',
    'النوع',
    'الوصف',
    'المصدر',
    'المدة (ثانية)',
    'الحالة',
    'الباحث',
    'تاريخ الطلب',
    'تاريخ الاستلام',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 120);
  sheet.setColumnWidth(11, 200);
}

/**
 * إنشاء شيت الداشبورد
 * @param {Spreadsheet} ss الجدول
 */
function createDashboardSheet(ss) {
  const sheetName = SHEETS.DASHBOARD;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  // عنوان الداشبورد
  sheet.getRange('A1').setValue('لوحة المتابعة الرئيسية')
    .setFontSize(18)
    .setFontWeight('bold');
  sheet.getRange('A1:H1').merge()
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // أقسام الداشبورد
  sheet.getRange('A3').setValue('ملخص المشاريع').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3:D3').merge().setBackground(COLORS.BACKGROUND_LIGHT);

  // جدول ملخص المشاريع
  const summaryHeaders = ['الحالة', 'العدد', 'النسبة', 'شريط'];
  sheet.getRange('A4:D4').setValues([summaryHeaders])
    .setFontWeight('bold')
    .setBackground(COLORS.BACKGROUND_ALT);

  // قسم المهام الحالية
  sheet.getRange('A12').setValue('المهام الجارية').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A12:H12').merge().setBackground(COLORS.BACKGROUND_LIGHT);

  const tasksHeaders = ['المشروع', 'المرحلة', 'المهمة', 'المسؤول', 'الحالة', 'تاريخ البداية', 'تاريخ النهاية', 'أيام متبقية'];
  sheet.getRange('A13:H13').setValues([tasksHeaders])
    .setFontWeight('bold')
    .setBackground(COLORS.BACKGROUND_ALT);

  sheet.setRightToLeft(true);

  // تعيين عرض الأعمدة
  for (let i = 1; i <= 8; i++) {
    sheet.setColumnWidth(i, 120);
  }
}

/**
 * إنشاء شيت تقرير الشخص
 * @param {Spreadsheet} ss الجدول
 */
function createPersonReportSheet(ss) {
  const sheetName = SHEETS.PERSON_REPORT;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  // عنوان التقرير
  sheet.getRange('A1').setValue('تقرير أداء الشخص')
    .setFontSize(16)
    .setFontWeight('bold');
  sheet.getRange('A1:E1').merge()
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // حقل اختيار الشخص
  sheet.getRange('A3').setValue('اختر الشخص:').setFontWeight('bold');
  sheet.getRange('B3').setValue('').setBackground(COLORS.WARNING);

  // حقل الفترة
  sheet.getRange('A4').setValue('من تاريخ:').setFontWeight('bold');
  sheet.getRange('C4').setValue('إلى تاريخ:').setFontWeight('bold');

  // ملخص الأداء
  sheet.getRange('A6').setValue('ملخص الأداء').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A6:E6').merge().setBackground(COLORS.BACKGROUND_LIGHT);

  const summaryLabels = [
    ['إجمالي المهام', ''],
    ['المهام المكتملة', ''],
    ['المهام الجارية', ''],
    ['المهام المتأخرة', ''],
    ['نسبة الإنجاز', '']
  ];
  sheet.getRange('A7:B11').setValues(summaryLabels);

  // جدول المهام التفصيلي
  sheet.getRange('A13').setValue('تفاصيل المهام').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A13:E13').merge().setBackground(COLORS.BACKGROUND_LIGHT);

  const detailHeaders = ['التاريخ', 'المشروع', 'المهمة', 'الحالة', 'المدة'];
  sheet.getRange('A14:E14').setValues([detailHeaders])
    .setFontWeight('bold')
    .setBackground(COLORS.BACKGROUND_ALT);

  sheet.setRightToLeft(true);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
}

/**
 * إنشاء شيت سجل التصديرات
 * @param {Spreadsheet} ss الجدول
 */
function createExportLogSheet(ss) {
  const sheetName = SHEETS.EXPORT_LOG;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'المعرف',
    'تاريخ التصدير',
    'نوع التقرير',
    'الفترة',
    'المستخدم',
    'رابط الملف',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 200);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال مساعدة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * حذف الشيت الافتراضي (Sheet1) إذا كان موجوداً
 * @param {Spreadsheet} ss الجدول
 */
function deleteDefaultSheet(ss) {
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

/**
 * إعادة ترتيب الشيتات حسب الترتيب المنطقي
 */
function reorderSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrder = [
    SHEETS.DASHBOARD,
    SHEETS.MOVEMENT,
    SHEETS.PROJECTS,
    SHEETS.TEAM,
    SHEETS.PHOTOGRAPHERS,
    SHEETS.GUESTS,
    SHEETS.VOICEOVER,
    SHEETS.ANIMATION,
    SHEETS.ARCHIVE,
    SHEETS.PERSON_REPORT,
    SHEETS.EXPORT_LOG,
    SHEETS.SETTINGS
  ];

  sheetOrder.forEach((sheetName, index) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });
}

/**
 * إعادة تعيين شيت معين (مسح البيانات مع الحفاظ على الهيكل)
 * @param {string} sheetName اسم الشيت
 */
function resetSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
    }
  }
}

/**
 * إعادة تعيين جميع الشيتات (مسح البيانات مع الحفاظ على الهيكل)
 */
function resetAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'تحذير!',
    'سيتم مسح جميع البيانات من الشيتات.\nهذه العملية لا يمكن التراجع عنها.\nهل أنت متأكد؟',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    Object.values(SHEETS).forEach(sheetName => {
      if (sheetName !== SHEETS.SETTINGS) {
        resetSheet(sheetName);
      }
    });
    ui.alert('تم مسح جميع البيانات بنجاح');
  }
}
