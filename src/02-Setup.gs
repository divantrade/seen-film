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
 * يحتوي على أعمدة checkboxes لكل مرحلة
 * @param {Spreadsheet} ss الجدول
 */
function createProjectsSheet(ss) {
  const sheetName = SHEETS.PROJECTS;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  // رؤوس الأعمدة الأساسية
  const basicHeaders = [
    'الكود',
    'اسم الفيلم',
    'النوع',
    'تاريخ البداية',
    'التسليم المتوقع',
    'الحالة',
    'ملاحظات'
  ];

  // أعمدة المراحل (checkboxes)
  const phaseHeaders = Object.values(STAGES).map(s => s.icon + ' ' + s.name);

  // أعمدة النظام
  const systemHeaders = ['تاريخ الإنشاء', 'آخر تحديث'];

  // دمج جميع الرؤوس
  const headers = [...basicHeaders, ...phaseHeaders, ...systemHeaders];

  // تعيين الرؤوس
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // تجميد الصف الأول
  sheet.setFrozenRows(1);

  // تعيين عرض الأعمدة الأساسية
  sheet.setColumnWidth(1, 80);   // الكود
  sheet.setColumnWidth(2, 180);  // اسم الفيلم
  sheet.setColumnWidth(3, 100);  // النوع
  sheet.setColumnWidth(4, 110);  // تاريخ البداية
  sheet.setColumnWidth(5, 110);  // التسليم المتوقع
  sheet.setColumnWidth(6, 80);   // الحالة
  sheet.setColumnWidth(7, 150);  // ملاحظات

  // تعيين عرض أعمدة المراحل (أضيق)
  for (let i = 8; i < 8 + phaseHeaders.length; i++) {
    sheet.setColumnWidth(i, 50);
  }

  // تعيين عرض أعمدة النظام
  sheet.setColumnWidth(8 + phaseHeaders.length, 130);
  sheet.setColumnWidth(9 + phaseHeaders.length, 130);

  // إضافة checkboxes لأعمدة المراحل (الصفوف 2-100)
  const checkboxRange = sheet.getRange(2, 8, 99, phaseHeaders.length);
  checkboxRange.insertCheckboxes();

  // تلوين أعمدة المراحل
  sheet.getRange(1, 8, 1, phaseHeaders.length).setBackground('#E3F2FD');

  // تعيين تنسيق الاتجاه من اليمين لليسار
  sheet.setRightToLeft(true);
}

/**
 * إنشاء شيت الفريق
 * يحتوي على عمود المراحل المسؤول عنها
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
    'الكود',
    'الاسم',
    'الدور',
    'البريد الإلكتروني',
    'رقم الهاتف',
    'المراحل المسؤول عنها',
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
  sheet.setColumnWidth(1, 70);   // الكود
  sheet.setColumnWidth(2, 140);  // الاسم
  sheet.setColumnWidth(3, 100);  // الدور
  sheet.setColumnWidth(4, 180);  // البريد
  sheet.setColumnWidth(5, 120);  // الهاتف
  sheet.setColumnWidth(6, 200);  // المراحل
  sheet.setColumnWidth(7, 80);   // الحالة
  sheet.setColumnWidth(8, 110);  // تاريخ الانضمام
  sheet.setColumnWidth(9, 180);  // ملاحظات
}

/**
 * إنشاء شيت المصورين
 * يحتوي على عمود المعدات
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
    'الكود',
    'الاسم',
    'التخصص',
    'البريد الإلكتروني',
    'رقم الهاتف',
    'المعدات',
    'الحالة',
    'السعر اليومي',
    'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 70);   // الكود
  sheet.setColumnWidth(2, 140);  // الاسم
  sheet.setColumnWidth(3, 100);  // التخصص
  sheet.setColumnWidth(4, 180);  // البريد
  sheet.setColumnWidth(5, 120);  // الهاتف
  sheet.setColumnWidth(6, 180);  // المعدات
  sheet.setColumnWidth(7, 80);   // الحالة
  sheet.setColumnWidth(8, 90);   // السعر
  sheet.setColumnWidth(9, 180);  // ملاحظات
}

/**
 * إنشاء شيت الضيوف
 * يحتوي على أعمدة متابعة التواصل والتصوير
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
    'الكود',
    'الاسم',
    'المشروع',
    'نوع المشاركة',
    'حالة التواصل',
    'حالة الأسئلة',
    'البلد',
    'مكان التصوير',
    'تاريخ التصوير',
    'حالة التصوير',
    'يحتاج دوبلاج',
    'البريد الإلكتروني',
    'رقم الهاتف',
    'ملاحظات',
    'تاريخ الإنشاء',
    'آخر تحديث'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  sheet.setColumnWidth(1, 70);   // الكود
  sheet.setColumnWidth(2, 140);  // الاسم
  sheet.setColumnWidth(3, 140);  // المشروع
  sheet.setColumnWidth(4, 100);  // نوع المشاركة
  sheet.setColumnWidth(5, 100);  // حالة التواصل
  sheet.setColumnWidth(6, 90);   // حالة الأسئلة
  sheet.setColumnWidth(7, 80);   // البلد
  sheet.setColumnWidth(8, 120);  // مكان التصوير
  sheet.setColumnWidth(9, 100);  // تاريخ التصوير
  sheet.setColumnWidth(10, 90);  // حالة التصوير
  sheet.setColumnWidth(11, 80);  // يحتاج دوبلاج
  sheet.setColumnWidth(12, 180); // البريد
  sheet.setColumnWidth(13, 120); // الهاتف
  sheet.setColumnWidth(14, 150); // ملاحظات
  sheet.setColumnWidth(15, 110); // تاريخ الإنشاء
  sheet.setColumnWidth(16, 110); // آخر تحديث
}

/**
 * إنشاء شيت الحركة (الإدخال الرئيسي)
 * يحتوي على أعمدة ديناميكية للمشروع والمرحلة
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
    'النوع الفرعي',
    'العنصر',
    'الإجراء',
    'المسؤول',
    'الحالة',
    'تاريخ الاستحقاق',
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

  sheet.setColumnWidth(1, 70);   // المعرف
  sheet.setColumnWidth(2, 90);   // التاريخ
  sheet.setColumnWidth(3, 140);  // المشروع
  sheet.setColumnWidth(4, 140);  // المرحلة
  sheet.setColumnWidth(5, 90);   // النوع الفرعي
  sheet.setColumnWidth(6, 150);  // العنصر
  sheet.setColumnWidth(7, 120);  // الإجراء
  sheet.setColumnWidth(8, 110);  // المسؤول
  sheet.setColumnWidth(9, 90);   // الحالة
  sheet.setColumnWidth(10, 100); // تاريخ الاستحقاق
  sheet.setColumnWidth(11, 180); // ملاحظات
  sheet.setColumnWidth(12, 110); // أنشئ بواسطة
  sheet.setColumnWidth(13, 130); // تاريخ الإنشاء

  // تلوين تبادلي للصفوف
  sheet.getRange('A:M').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * إنشاء شيت التعليق الصوتي
 * يتضمن أعمدة: الكود، المشروع، النوع، رقم المقطع، النص، المؤدي، اللغة، الاستديو، الحالة، المدة، رابط الملف، ملاحظات
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
    'الكود',
    'المشروع',
    'النوع',
    'رقم المقطع',
    'النص',
    'المؤدي',
    'اللغة',
    'الاستديو',
    'الحالة',
    'المدة (دقيقة)',
    'رابط الملف',
    'ملاحظات',
    'تاريخ الإنشاء',
    'آخر تحديث'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(1, 80);   // الكود
  sheet.setColumnWidth(2, 140);  // المشروع
  sheet.setColumnWidth(3, 100);  // النوع
  sheet.setColumnWidth(4, 80);   // رقم المقطع
  sheet.setColumnWidth(5, 250);  // النص
  sheet.setColumnWidth(6, 120);  // المؤدي
  sheet.setColumnWidth(7, 80);   // اللغة
  sheet.setColumnWidth(8, 120);  // الاستديو
  sheet.setColumnWidth(9, 100);  // الحالة
  sheet.setColumnWidth(10, 90);  // المدة
  sheet.setColumnWidth(11, 200); // رابط الملف
  sheet.setColumnWidth(12, 180); // ملاحظات
  sheet.setColumnWidth(13, 120); // تاريخ الإنشاء
  sheet.setColumnWidth(14, 120); // آخر تحديث

  // تلوين تبادلي للصفوف
  sheet.getRange('A:N').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * إنشاء شيت الرسوم المتحركة
 * يتضمن أعمدة: الكود، المشروع، رقم المشهد، الوصف، النوع، المدة، رابط السكريبت، الاستديو، المحرك، الحالة، رابط الملف
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
    'الكود',
    'المشروع',
    'رقم المشهد',
    'الوصف',
    'النوع',
    'المدة (ثانية)',
    'رابط السكريبت',
    'الاستديو',
    'المحرك',
    'الحالة',
    'رابط الملف',
    'ملاحظات',
    'تاريخ الإنشاء',
    'آخر تحديث'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(1, 80);   // الكود
  sheet.setColumnWidth(2, 140);  // المشروع
  sheet.setColumnWidth(3, 80);   // رقم المشهد
  sheet.setColumnWidth(4, 200);  // الوصف
  sheet.setColumnWidth(5, 130);  // النوع
  sheet.setColumnWidth(6, 90);   // المدة
  sheet.setColumnWidth(7, 200);  // رابط السكريبت
  sheet.setColumnWidth(8, 120);  // الاستديو
  sheet.setColumnWidth(9, 120);  // المحرك
  sheet.setColumnWidth(10, 100); // الحالة
  sheet.setColumnWidth(11, 200); // رابط الملف
  sheet.setColumnWidth(12, 180); // ملاحظات
  sheet.setColumnWidth(13, 120); // تاريخ الإنشاء
  sheet.setColumnWidth(14, 120); // آخر تحديث

  // تلوين تبادلي للصفوف
  sheet.getRange('A:N').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * إنشاء شيت الأرشيف
 * يتضمن أعمدة: الكود، المشروع، نوع الأرشيف، العنوان، الوصف، المصدر، تاريخ الحصول، حالة الترخيص، انتهاء الترخيص، تكلفة الترخيص، رابط الملف، مستخدم في
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
    'الكود',
    'المشروع',
    'نوع الأرشيف',
    'العنوان',
    'الوصف',
    'المصدر',
    'تاريخ المادة الأصلي',
    'تاريخ الحصول',
    'حالة الترخيص',
    'انتهاء الترخيص',
    'تكلفة الترخيص',
    'رابط الملف',
    'صورة مصغرة',
    'مستخدم في',
    'المدة',
    'الدقة',
    'ملاحظات',
    'أضافه',
    'تاريخ الإضافة',
    'آخر تحديث'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  // تعيين عرض الأعمدة
  sheet.setColumnWidth(1, 80);   // الكود
  sheet.setColumnWidth(2, 140);  // المشروع
  sheet.setColumnWidth(3, 100);  // نوع الأرشيف
  sheet.setColumnWidth(4, 160);  // العنوان
  sheet.setColumnWidth(5, 200);  // الوصف
  sheet.setColumnWidth(6, 120);  // المصدر
  sheet.setColumnWidth(7, 110);  // تاريخ المادة الأصلي
  sheet.setColumnWidth(8, 100);  // تاريخ الحصول
  sheet.setColumnWidth(9, 100);  // حالة الترخيص
  sheet.setColumnWidth(10, 110); // انتهاء الترخيص
  sheet.setColumnWidth(11, 100); // تكلفة الترخيص
  sheet.setColumnWidth(12, 200); // رابط الملف
  sheet.setColumnWidth(13, 150); // صورة مصغرة
  sheet.setColumnWidth(14, 150); // مستخدم في
  sheet.setColumnWidth(15, 80);  // المدة
  sheet.setColumnWidth(16, 80);  // الدقة
  sheet.setColumnWidth(17, 180); // ملاحظات
  sheet.setColumnWidth(18, 100); // أضافه
  sheet.setColumnWidth(19, 110); // تاريخ الإضافة
  sheet.setColumnWidth(20, 110); // آخر تحديث

  // تلوين تبادلي للصفوف
  sheet.getRange('A:T').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
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
