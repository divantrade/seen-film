/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف القائمة المخصصة
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على القائمة المخصصة وجميع عناصرها
 * يتم تحميل القائمة تلقائياً عند فتح الجدول
 */

// ═══════════════════════════════════════════════════════════════════════════════
// دالة تشغيل القائمة عند فتح الجدول
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * دالة تُنفذ تلقائياً عند فتح الجدول
 * تقوم بإنشاء القائمة المخصصة
 */
function onOpen() {
  createCustomMenu();
}

/**
 * دالة تُنفذ عند تثبيت الإضافة
 */
function onInstall() {
  onOpen();
}

// ═══════════════════════════════════════════════════════════════════════════════
// إنشاء القائمة المخصصة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء القائمة المخصصة الرئيسية
 * تحتوي على جميع وظائف النظام منظمة في قوائم فرعية
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🎬 نظام الإنتاج')

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة الإعداد
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('⚙️ الإعداد')
      .addItem('🔧 إعداد النظام الكامل', 'setupSystem')
      .addSeparator()
      .addItem('📋 إنشاء شيت الإعدادات', 'createSettingsSheetMenu')
      .addItem('📁 إنشاء شيت المشاريع', 'createProjectsSheetMenu')
      .addItem('👥 إنشاء شيت الفريق', 'createTeamSheetMenu')
      .addItem('📸 إنشاء شيت المصورين', 'createPhotographersSheetMenu')
      .addItem('🎤 إنشاء شيت الضيوف', 'createGuestsSheetMenu')
      .addItem('📝 إنشاء شيت الحركة', 'createMovementSheetMenu')
      .addItem('🎙️ إنشاء شيت التعليق الصوتي', 'createVoiceOverSheetMenu')
      .addItem('🎨 إنشاء شيت الرسوم المتحركة', 'createAnimationSheetMenu')
      .addItem('🗄️ إنشاء شيت الأرشيف', 'createArchiveSheetMenu')
      .addSeparator()
      .addItem('🔄 إعادة ترتيب الشيتات', 'reorderSheets')
      .addItem('⚠️ إعادة تعيين جميع البيانات', 'resetAllSheets')
    )

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة المشاريع
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('📁 المشاريع')
      .addItem('➕ إضافة مشروع جديد', 'addNewProject')
      .addItem('📋 عرض قائمة المشاريع', 'goToProjectsSheet')
      .addSeparator()
      .addItem('📊 ملخص المشاريع', 'showProjectsSummary')
      .addItem('🔍 البحث في المشاريع', 'searchProjects')
    )

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة الفريق
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('👥 الفريق')
      .addItem('➕ إضافة عضو جديد', 'addNewTeamMember')
      .addItem('📋 عرض قائمة الفريق', 'goToTeamSheet')
      .addItem('📸 عرض المصورين', 'goToPhotographersSheet')
      .addSeparator()
      .addItem('🔢 توليد أكواد الفريق المفقودة', 'generateMissingTeamCodes')
      .addItem('📊 تقرير أداء الفريق', 'showTeamPerformance')
    )

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة الإنتاج
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('🎬 الإنتاج')
      .addItem('📝 إضافة حركة جديدة', 'addNewMovement')
      .addItem('✨ إضافة حركة ذكية', 'showSmartMovementForm')
      .addItem('📋 عرض شيت الحركة', 'goToMovementSheet')
      .addSeparator()
      .addItem('🎤 إضافة ضيف جديد', 'addNewGuest')
      .addItem('🎙️ إضافة تعليق صوتي', 'addNewVoiceOver')
      .addItem('🎨 إضافة رسوم متحركة', 'addNewAnimation')
      .addItem('🗄️ إضافة طلب أرشيف', 'addNewArchiveRequest')
    )

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة التقارير
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('📊 التقارير')
      .addItem('📈 لوحة المتابعة', 'goToDashboard')
      .addItem('👤 تقرير الشخص', 'goToPersonReport')
      .addSeparator()
      .addItem('📋 تقرير المشاريع', 'generateProjectsReport')
      .addItem('📋 تقرير الفريق', 'generateTeamReport')
      .addItem('📋 تقرير الحركة اليومي', 'generateDailyMovementReport')
      .addItem('📋 تقرير الحركة الأسبوعي', 'generateWeeklyMovementReport')
    )

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة التصدير
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('📤 التصدير')
      .addItem('📸 تصدير للمصور...', 'showExportForPhotographerDialog')
      .addItem('🎙️ تصدير للاستوديو الصوتي...', 'showExportForVoiceStudioDialog')
      .addItem('🎨 تصدير لاستوديو الرسوم...', 'showExportForAnimationStudioDialog')
      .addItem('🎞️ تصدير للمونتير...', 'showExportForEditorDialog')
      .addItem('📋 تصدير تقرير مشروع...', 'showExportProjectReportDialog')
      .addSeparator()
      .addItem('📄 تصدير الشيت الحالي كـ PDF', 'exportCurrentSheetAsPDF')
      .addItem('📊 تصدير الداشبورد كـ PDF', 'exportDashboardAsPDF')
      .addSeparator()
      .addItem('📜 عرض سجل التصديرات', 'goToExportLog')
      .addItem('🔄 آخر 10 تصديرات', 'showRecentExports')
    )

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة الإشعارات
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('📧 الإشعارات')
      .addItem('📋 إرسال تقرير يومي الآن', 'sendDailyReportNow')
      .addItem('📊 إرسال تقرير أسبوعي الآن', 'sendWeeklyReportNow')
      .addSeparator()
      .addItem('⚙️ إعدادات الإشعارات...', 'showNotificationSettings')
      .addSeparator()
      .addItem('✅ تفعيل الإشعارات التلقائية', 'enableAutoNotifications')
      .addItem('❌ إيقاف الإشعارات التلقائية', 'disableAutoNotifications')
    )

    // ═══════════════════════════════════════════════════════════════════════════
    // قائمة الأدوات
    // ═══════════════════════════════════════════════════════════════════════════
    .addSubMenu(ui.createMenu('🔧 الأدوات')
      .addItem('🔄 تحديث القوائم المنسدلة', 'refreshAllDropdowns')
      .addItem('🎨 تحديث ألوان الحالات', 'refreshStatusColors')
      .addItem('✅ تطبيق تنسيق الـ Checkboxes الخضراء', 'applyCheckboxFormatting')
      .addSeparator()
      .addItem('➕ مزامنة أعمدة المراحل الجديدة', 'syncPhaseColumns')
      .addItem('🔄 إعادة بناء أعمدة المراحل', 'rebuildPhaseColumns')
      .addItem('🔧 إصلاح شيت المشاريع', 'fixProjectsSheet')
      .addSeparator()
      .addItem('🔍 التحقق من البيانات', 'validateAllData')
      .addItem('🧹 تنظيف البيانات المكررة', 'cleanDuplicateData')
    )

    .addSeparator()

    // ═══════════════════════════════════════════════════════════════════════════
    // معلومات النظام
    // ═══════════════════════════════════════════════════════════════════════════
    .addItem('ℹ️ حول النظام', 'showAbout')
    .addItem('❓ المساعدة', 'showHelp')

    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التنقل بين الشيتات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الانتقال إلى شيت معين
 * @param {string} sheetName اسم الشيت
 */
function goToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.setActiveSheet(sheet);
    sheet.getRange('A1').activate();
  } else {
    showError('الشيت غير موجود: ' + sheetName);
  }
}

/** الانتقال إلى شيت المشاريع */
function goToProjectsSheet() {
  goToSheet(SHEETS.PROJECTS);
}

/** الانتقال إلى شيت الفريق */
function goToTeamSheet() {
  goToSheet(SHEETS.TEAM);
}

/** الانتقال إلى شيت المصورين */
function goToPhotographersSheet() {
  goToSheet(SHEETS.PHOTOGRAPHERS);
}

/** الانتقال إلى شيت الحركة */
function goToMovementSheet() {
  goToSheet(SHEETS.MOVEMENT);
}

/** الانتقال إلى الداشبورد */
function goToDashboard() {
  goToSheet(SHEETS.DASHBOARD);
}

/** الانتقال إلى تقرير الشخص */
function goToPersonReport() {
  goToSheet(SHEETS.PERSON_REPORT);
}

/** الانتقال إلى سجل التصديرات */
function goToExportLog() {
  goToSheet(SHEETS.EXPORT_LOG);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إنشاء الشيتات من القائمة
// ═══════════════════════════════════════════════════════════════════════════════

/** إنشاء شيت الإعدادات */
function createSettingsSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSettingsSheet(ss);
  showToast('تم إنشاء شيت الإعدادات', 'نجاح');
}

/** إنشاء شيت المشاريع */
function createProjectsSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createProjectsSheet(ss);
  showToast('تم إنشاء شيت المشاريع', 'نجاح');
}

/** إنشاء شيت الفريق */
function createTeamSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createTeamSheet(ss);
  showToast('تم إنشاء شيت الفريق', 'نجاح');
}

/** إنشاء شيت المصورين */
function createPhotographersSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createPhotographersSheet(ss);
  showToast('تم إنشاء شيت المصورين', 'نجاح');
}

/** إنشاء شيت الضيوف */
function createGuestsSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createGuestsSheet(ss);
  showToast('تم إنشاء شيت الضيوف', 'نجاح');
}

/** إنشاء شيت الحركة */
function createMovementSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createMovementSheet(ss);
  showToast('تم إنشاء شيت الحركة', 'نجاح');
}

/** إنشاء شيت التعليق الصوتي */
function createVoiceOverSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createVoiceOverSheet(ss);
  showToast('تم إنشاء شيت التعليق الصوتي', 'نجاح');
}

/** إنشاء شيت الرسوم المتحركة */
function createAnimationSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createAnimationSheet(ss);
  showToast('تم إنشاء شيت الرسوم المتحركة', 'نجاح');
}

/** إنشاء شيت الأرشيف */
function createArchiveSheetMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createArchiveSheet(ss);
  showToast('تم إنشاء شيت الأرشيف', 'نجاح');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإضافة (Placeholders - سيتم تنفيذها في ملفات أخرى)
// ═══════════════════════════════════════════════════════════════════════════════

/** إضافة مشروع جديد */
function addNewProject() {
  showAddProjectDialog();
}

/** إضافة عضو فريق جديد */
function addNewTeamMember() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 06-Team.gs');
  goToTeamSheet();
}

/** إضافة حركة جديدة */
function addNewMovement() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 08-Movement.gs');
  goToMovementSheet();
}

/** إضافة ضيف جديد */
function addNewGuest() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 07-Guests.gs');
  goToSheet(SHEETS.GUESTS);
}

/** إضافة تعليق صوتي */
function addNewVoiceOver() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 09-VoiceOver.gs');
  goToSheet(SHEETS.VOICEOVER);
}

/** إضافة رسوم متحركة */
function addNewAnimation() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 10-Animation.gs');
  goToSheet(SHEETS.ANIMATION);
}

/** إضافة طلب أرشيف */
function addNewArchiveRequest() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 11-Archive.gs');
  goToSheet(SHEETS.ARCHIVE);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التقارير (Placeholders - سيتم تنفيذها في ملفات أخرى)
// ═══════════════════════════════════════════════════════════════════════════════

// ملاحظة: showProjectsSummary موجودة في 05-Projects.gs
// ملاحظة: searchProjects موجودة في 05-Projects.gs

/** تقرير أداء الفريق */
function showTeamPerformance() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 13-Reports.gs');
}

/** تقرير المشاريع */
function generateProjectsReport() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 13-Reports.gs');
}

/** تقرير الفريق */
function generateTeamReport() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 13-Reports.gs');
}

/** تقرير الحركة اليومي */
function generateDailyMovementReport() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 13-Reports.gs');
}

/** تقرير الحركة الأسبوعي */
function generateWeeklyMovementReport() {
  showInfo('سيتم تنفيذ هذه الوظيفة في ملف 13-Reports.gs');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التصدير
// ═══════════════════════════════════════════════════════════════════════════════

/** تصدير الشيت الحالي كـ PDF */
function exportCurrentSheetAsPDF() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const fileName = `${sheet.getName()}_${getCurrentDate()}`;
    const fileUrl = exportSheetAsPDF(sheet.getName(), fileName);

    logExport(sheet.getName(), getCurrentDate(), fileUrl, 'تصدير من القائمة');
    showSuccess(`تم التصدير بنجاح!\n\nرابط الملف:\n${fileUrl}`);
  } catch (error) {
    showError('حدث خطأ أثناء التصدير:\n' + error.message);
  }
}

/** تصدير الداشبورد كـ PDF */
function exportDashboardAsPDF() {
  try {
    const fileName = `لوحة_المتابعة_${getCurrentDate()}`;
    const fileUrl = exportSheetAsPDF(SHEETS.DASHBOARD, fileName);

    logExport('داشبورد', getCurrentDate(), fileUrl, 'تصدير من القائمة');
    showSuccess(`تم التصدير بنجاح!\n\nرابط الملف:\n${fileUrl}`);
  } catch (error) {
    showError('حدث خطأ أثناء التصدير:\n' + error.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الأدوات (Placeholders - سيتم تنفيذها في ملفات أخرى)
// ═══════════════════════════════════════════════════════════════════════════════

// ملاحظة: refreshAllDropdowns موجودة في 04-Dropdowns.gs
// ملاحظة: refreshStatusColors موجودة في 04-Dropdowns.gs

/** التحقق من البيانات */
function validateAllData() {
  showInfo('سيتم تنفيذ هذه الوظيفة لاحقاً');
}

/** تنظيف البيانات المكررة */
function cleanDuplicateData() {
  showInfo('سيتم تنفيذ هذه الوظيفة لاحقاً');
}

// ═══════════════════════════════════════════════════════════════════════════════
// معلومات النظام والمساعدة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * عرض معلومات حول النظام
 */
function showAbout() {
  const message = `
${CONFIG.SYSTEM_NAME}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

الإصدار: ${CONFIG.VERSION}
المنطقة الزمنية: ${CONFIG.TIMEZONE}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

نظام متكامل لمتابعة الإنتاج الفني
في شركات الأفلام الوثائقية

يشمل:
• إدارة المشاريع والفريق
• متابعة مراحل الإنتاج
• إدارة الضيوف والمقابلات
• التعليق الصوتي والرسوم المتحركة
• إدارة الأرشيف والتراخيص
• التقارير والداشبورد التفاعلي
• تصدير شيتات منفصلة للفريق
• إشعارات تلقائية بالإيميل

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  `.trim();

  showInfo(message, 'حول النظام');
}

/**
 * عرض المساعدة
 */
function showHelp() {
  const message = `
دليل الاستخدام السريع
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1️⃣ البداية:
   • استخدم "إعداد النظام الكامل" لإنشاء جميع الشيتات

2️⃣ إضافة البيانات:
   • أضف المشاريع من قائمة "المشاريع"
   • أضف الفريق من قائمة "الفريق"
   • سجل الحركة اليومية من قائمة "الإنتاج"

3️⃣ المتابعة:
   • استخدم "الداشبورد" لرؤية ملخص العمل
   • استخدم "تقرير الشخص" لمتابعة أداء فرد معين

4️⃣ التصدير:
   • تصدير شيتات مخصصة للمصورين والاستوديوهات
   • تصدير تقارير PDF للمشاريع
   • جميع عمليات التصدير تُسجل تلقائياً

5️⃣ الإشعارات:
   • تقارير يومية وأسبوعية بالإيميل
   • تنبيهات المهام المتأخرة
   • إعداد الإشعارات من "📧 الإشعارات"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

المراحل المتاحة:
📄 الورق/البحث
🔧 الفيكسر
🎬 التصوير الميداني
🎤 تصوير المقابلات
🎭 تصوير الدراما
🎙️ التعليق الصوتي
🎨 الرسوم المتحركة
📊 الانفوجراف
🎞️ المونتاج
🗄️ الأرشيف
✔️ المراجعة
📦 التسليم

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

الحالات:
⬜ لم يبدأ
⏳ في الانتظار
🔄 جاري
✅ تم
🔴 متأخر
❌ ملغي
  `.trim();

  showInfo(message, 'المساعدة');
}
