/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * لوحة التحكم (Dashboard)
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * تحديث لوحة التحكم
 */
function refreshDashboard() {
  const sheet = getSheet(SHEETS.DASHBOARD);
  if (!sheet) {
    showError('شيت الداشبورد غير موجود');
    return;
  }

  // الحصول على الفيلم المحدد
  const selectedProject = sheet.getRange('B3').getValue();
  const isAll = !selectedProject || selectedProject === 'الكل';

  // الحصول على الإحصائيات
  const stats = isAll ? getMovementStats() : getMovementStats(selectedProject);
  const projects = getActiveProjects();

  // تنظيف الداشبورد (باستثناء العنوان والاختيار)
  sheet.getRange('A5:F100').clear();

  let currentRow = 5;

  // ═══════════════════════════════════════════════════════════════════════════════
  // ملخص المهام
  // ═══════════════════════════════════════════════════════════════════════════════

  sheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('ملخص المهام')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  currentRow++;

  // جدول الملخص
  sheet.getRange(currentRow, 1, 1, 4).setValues([['إجمالي', 'تم ✅', 'جاري 🔄', 'متأخر 🔴']]);
  sheet.getRange(currentRow, 1, 1, 4)
    .setBackground(COLORS.BACKGROUND)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  currentRow++;

  sheet.getRange(currentRow, 1, 1, 4).setValues([[
    stats.total,
    stats.completed,
    stats.inProgress,
    stats.delayed
  ]]);
  sheet.getRange(currentRow, 1, 1, 4)
    .setHorizontalAlignment('center')
    .setFontSize(14);

  // تلوين الخلايا
  sheet.getRange(currentRow, 2).setBackground(COLORS.SUCCESS);
  sheet.getRange(currentRow, 3).setBackground(COLORS.INFO);
  sheet.getRange(currentRow, 4).setBackground(COLORS.DANGER);

  currentRow += 2;

  // ═══════════════════════════════════════════════════════════════════════════════
  // تقدم المراحل
  // ═══════════════════════════════════════════════════════════════════════════════

  sheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('تقدم المراحل')
    .setBackground(COLORS.HEADER)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  currentRow++;

  sheet.getRange(currentRow, 1, 1, 4).setValues([['المرحلة', 'الإجمالي', 'المكتمل', 'النسبة']]);
  sheet.getRange(currentRow, 1, 1, 4)
    .setBackground(COLORS.BACKGROUND)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  currentRow++;

  // المراحل
  for (const stageName of STAGE_NAMES) {
    const stageStats = stats.byStage[stageName] || { total: 0, completed: 0 };
    const percentage = stageStats.total > 0
      ? Math.round((stageStats.completed / stageStats.total) * 100)
      : 0;

    sheet.getRange(currentRow, 1, 1, 4).setValues([[
      stageName,
      stageStats.total,
      stageStats.completed,
      percentage + '%'
    ]]);

    // تلوين حسب النسبة
    let bgColor = COLORS.BACKGROUND;
    if (percentage >= 100) bgColor = COLORS.SUCCESS;
    else if (percentage >= 50) bgColor = COLORS.WARNING;
    else if (percentage > 0) bgColor = COLORS.DANGER;

    sheet.getRange(currentRow, 4).setBackground(bgColor);

    currentRow++;
  }

  currentRow++;

  // ═══════════════════════════════════════════════════════════════════════════════
  // المهام المتأخرة
  // ═══════════════════════════════════════════════════════════════════════════════

  const delayedTasks = getDelayedTasks();

  if (delayedTasks.length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue('⚠️ المهام المتأخرة (' + delayedTasks.length + ')')
      .setBackground(COLORS.DANGER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    currentRow++;

    sheet.getRange(currentRow, 1, 1, 4).setValues([['الفيلم', 'المرحلة', 'العنصر', 'المسؤول']]);
    sheet.getRange(currentRow, 1, 1, 4)
      .setBackground(COLORS.BACKGROUND)
      .setFontWeight('bold');

    currentRow++;

    // عرض أول 10 مهام متأخرة
    const displayTasks = delayedTasks.slice(0, 10);
    for (const task of displayTasks) {
      sheet.getRange(currentRow, 1, 1, 4).setValues([[
        task.project,
        task.stage,
        task.element,
        task.assignedTo || '-'
      ]]);
      currentRow++;
    }
  }

  currentRow++;

  // ═══════════════════════════════════════════════════════════════════════════════
  // ملخص المشاريع (إذا كان "الكل")
  // ═══════════════════════════════════════════════════════════════════════════════

  if (isAll && projects.length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue('ملخص المشاريع')
      .setBackground(COLORS.HEADER)
      .setFontColor(COLORS.HEADER_TEXT)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    currentRow++;

    sheet.getRange(currentRow, 1, 1, 4).setValues([['المشروع', 'المهام', 'المكتمل', 'النسبة']]);
    sheet.getRange(currentRow, 1, 1, 4)
      .setBackground(COLORS.BACKGROUND)
      .setFontWeight('bold');

    currentRow++;

    for (const project of projects) {
      const projectStats = getMovementStats(project.name);
      const percentage = projectStats.total > 0
        ? Math.round((projectStats.completed / projectStats.total) * 100)
        : 0;

      sheet.getRange(currentRow, 1, 1, 4).setValues([[
        project.name,
        projectStats.total,
        projectStats.completed,
        percentage + '%'
      ]]);

      // تلوين حسب النسبة
      let bgColor = COLORS.BACKGROUND;
      if (percentage >= 100) bgColor = COLORS.SUCCESS;
      else if (percentage >= 50) bgColor = COLORS.WARNING;
      else if (percentage > 0) bgColor = COLORS.DANGER;

      sheet.getRange(currentRow, 4).setBackground(bgColor);

      currentRow++;
    }
  }

  // تحديث قائمة المشاريع المنسدلة
  const projectNames = ['الكل', ...getActiveProjectNames()];
  const projectDropdown = SpreadsheetApp.newDataValidation()
    .requireValueInList(projectNames, true)
    .build();
  sheet.getRange('B3').setDataValidation(projectDropdown);

  // ضبط عرض الأعمدة
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);

  // إضافة تاريخ آخر تحديث
  currentRow += 2;
  sheet.getRange(currentRow, 1).setValue('آخر تحديث: ' + getCurrentDateTime());
  sheet.getRange(currentRow, 1).setFontColor('#757575').setFontSize(10);

  showSuccess('تم تحديث لوحة التحكم');
}

/**
 * Trigger عند تغيير اختيار الفيلم في الداشبورد
 */
function onDashboardEdit(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== SHEETS.DASHBOARD) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // إذا تم تغيير اختيار الفيلم
  if (row === 3 && col === 2) {
    refreshDashboard();
  }
}

/**
 * إنشاء تقرير مشروع
 */
function generateProjectReport(projectName) {
  const stats = getMovementStats(projectName);
  const movements = getProjectMovements(projectName);

  let report = '═══════════════════════════════════════\n';
  report += '📊 تقرير مشروع: ' + projectName + '\n';
  report += '═══════════════════════════════════════\n\n';

  report += '📈 ملخص الإحصائيات:\n';
  report += '  • إجمالي المهام: ' + stats.total + '\n';
  report += '  • المكتملة: ' + stats.completed + '\n';
  report += '  • جاري العمل: ' + stats.inProgress + '\n';
  report += '  • المتأخرة: ' + stats.delayed + '\n\n';

  report += '📋 المراحل:\n';
  for (const stage in stats.byStage) {
    const s = stats.byStage[stage];
    const pct = s.total > 0 ? Math.round((s.completed / s.total) * 100) : 0;
    report += '  • ' + stage + ': ' + s.completed + '/' + s.total + ' (' + pct + '%)\n';
  }

  return report;
}

/**
 * عرض تقرير المشروع المحدد
 */
function showProjectReport() {
  const sheet = getSheet(SHEETS.DASHBOARD);
  const selectedProject = sheet.getRange('B3').getValue();

  if (!selectedProject || selectedProject === 'الكل') {
    showError('اختر مشروعاً محدداً لعرض تقريره');
    return;
  }

  const report = generateProjectReport(selectedProject);

  const html = HtmlService.createHtmlOutput(`
    <pre style="font-family: monospace; direction: rtl; text-align: right; white-space: pre-wrap;">
${report}
    </pre>
  `)
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'تقرير المشروع');
}

/**
 * حذف شيت لوحة التحكم القديم
 */
function deleteControlPanelSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // البحث عن شيت لوحة التحكم
  const sheetNames = ['لوحة التحكم', '🎛️ لوحة التحكم'];
  let sheetToDelete = null;

  for (const name of sheetNames) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      sheetToDelete = sheet;
      break;
    }
  }

  if (!sheetToDelete) {
    ui.alert('ℹ️ معلومة', 'لم يتم العثور على شيت "لوحة التحكم"', ui.ButtonSet.OK);
    return;
  }

  // تأكيد الحذف
  const response = ui.alert(
    '🗑️ حذف شيت لوحة التحكم',
    'هل أنت متأكد من حذف شيت "لوحة التحكم"؟\n\nسيتم الاعتماد على شيت "داشبورد" فقط.',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    ss.deleteSheet(sheetToDelete);
    ui.alert('✅ تم', 'تم حذف شيت "لوحة التحكم" بنجاح.\n\nيمكنك الآن استخدام شيت "داشبورد" فقط.', ui.ButtonSet.OK);
  }
}

/**
 * إضافة أيقونة لاسم شيت الداشبورد وتحسين المظهر
 */
function upgradeDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // البحث عن شيت الداشبورد
  let sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sheet) {
    sheet = ss.getSheetByName('داشبورد');
  }

  if (!sheet) {
    ui.alert('❌ خطأ', 'لم يتم العثور على شيت الداشبورد', ui.ButtonSet.OK);
    return;
  }

  // تغيير اسم التاب مع أيقونة
  const newName = '📊 داشبورد';
  try {
    sheet.setName(newName);
  } catch (e) {
    // إذا كان الاسم موجود مسبقاً، نتجاهل الخطأ
  }

  // تحسين لون التاب
  sheet.setTabColor('#1a73e8'); // أزرق جوجل

  // تحديث العنوان الرئيسي
  sheet.getRange('A1').setValue('📊 لوحة المتابعة - نظام إدارة الإنتاج');
  sheet.getRange('A1:D1').merge()
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // تحديث الداشبورد بالبيانات
  refreshDashboard();

  ui.alert('✅ تم التحسين',
    'تم تحسين شيت الداشبورد:\n\n' +
    '• تم إضافة أيقونة 📊 للتاب\n' +
    '• تم تغيير لون التاب للأزرق\n' +
    '• تم تحسين العنوان الرئيسي\n' +
    '• تم تحديث البيانات',
    ui.ButtonSet.OK);
}

/**
 * تنظيف وترتيب الشيتات مع أيقونات
 */
function organizeAllSheetTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // قائمة الشيتات مع الأيقونات والألوان
  const sheetConfig = [
    { oldName: 'المشاريع', newName: '🎬 المشاريع', color: '#34a853' },
    { oldName: 'الفريق', newName: '👥 الفريق', color: '#fbbc04' },
    { oldName: 'الحركة', newName: '📋 الحركة', color: '#ea4335' },
    { oldName: 'داشبورد', newName: '📊 داشبورد', color: '#1a73e8' },
    { oldName: 'الإعدادات', newName: '⚙️ الإعدادات', color: '#9e9e9e' },
    { oldName: 'المستخدمين', newName: '👤 المستخدمين', color: '#9c27b0' },
    { oldName: 'روابط الفولدرات', newName: '📁 روابط الفولدرات', color: '#ff9800' }
  ];

  let updatedCount = 0;

  for (const config of sheetConfig) {
    // جرب الاسم القديم أو الاسم الجديد (إذا كان محدث مسبقاً)
    let sheet = ss.getSheetByName(config.oldName) || ss.getSheetByName(config.newName);
    // خاص بشيت المستخدمين - جرب الاسم القديم بأيقونة القفل
    if (!sheet && config.oldName === 'المستخدمين') {
      sheet = ss.getSheetByName('🔐 المستخدمين');
    }

    if (sheet) {
      try {
        const currentName = sheet.getName();
        // إذا كان الاسم القديم بأيقونة القفل، غيره للأيقونة الجديدة
        if (currentName === '🔐 المستخدمين') {
          sheet.setName('👤 المستخدمين');
        }
        // لا نغير الاسم إذا كان يبدأ بأيقونة بالفعل (غير القفل)
        else if (!currentName.match(/^[📊🎬👥📋⚙️👤📁]/)) {
          sheet.setName(config.newName);
        }
        sheet.setTabColor(config.color);
        updatedCount++;
      } catch (e) {
        console.log('Could not update sheet: ' + config.oldName);
      }
    }
  }

  SpreadsheetApp.getUi().alert('✅ تم الترتيب',
    'تم تحديث ' + updatedCount + ' شيت بأيقونات وألوان مميزة.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * إزالة جميع الحمايات الوهمية من الشيتات
 * هذه الحمايات تظهر كقفل صغير بجانب اسم التاب لكنها لا تمنع أحداً من التعديل
 */
function removeAllSheetProtections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    '⚠️ إزالة الحمايات',
    'سيتم إزالة جميع الحمايات (الأقفال) من كل الشيتات.\n\nهل تريد المتابعة؟',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  let removedCount = 0;
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    try {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (const protection of protections) {
        protection.remove();
        removedCount++;
      }

      // إزالة حمايات النطاقات أيضاً
      const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (const protection of rangeProtections) {
        protection.remove();
        removedCount++;
      }
    } catch (e) {
      console.log('Could not remove protection from: ' + sheet.getName(), e);
    }
  }

  ui.alert('✅ تم',
    'تم إزالة ' + removedCount + ' حماية من الشيتات.\n\nالآن لن تظهر أيقونات القفل الوهمية.',
    ui.ButtonSet.OK);
}
