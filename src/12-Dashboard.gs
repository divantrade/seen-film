/**
 * ===================================================
 * 12-Dashboard.gs - الداشبورد التفاعلي
 * ===================================================
 *
 * هذا الملف يحتوي على:
 * - عرض معلومات المشروع
 * - نسب إكمال المراحل
 * - إحصائيات الضيوف والتعليق الصوتي والرسوم المتحركة
 * - تنبيهات المهام المتأخرة
 */

// ====================================================
// ثوابت الداشبورد
// ====================================================

/**
 * موقع خلايا الداشبورد
 */
const DASHBOARD_CELLS = {
  PROJECT_CODE_INPUT: 'B2',     // خلية إدخال كود المشروع
  PROJECT_NAME: 'D2',           // اسم المشروع
  PROJECT_STATUS: 'F2',         // حالة المشروع
  PROJECT_CLIENT: 'B4',         // العميل (اسم القناة)
  PROJECT_START: 'D4',          // تاريخ البداية
  PROJECT_DEADLINE: 'F4',       // الموعد النهائي

  // نسب إكمال المراحل
  PHASES_START_ROW: 7,
  PHASES_START_COL: 2,

  // الإحصائيات
  STATS_START_ROW: 7,
  STATS_START_COL: 5,

  // التنبيهات
  ALERTS_START_ROW: 20,
  ALERTS_START_COL: 2
};

// ====================================================
// دوال تحديث الداشبورد
// ====================================================

/**
 * تحديث الداشبورد بناءً على كود المشروع المدخل
 * يتم استدعاؤها من onEdit عند تغيير خلية كود المشروع
 */
function refreshDashboard() {
  const sheet = getSheet(SHEETS.DASHBOARD);
  if (!sheet) return;

  // الحصول على كود المشروع من خلية الإدخال
  const projectCode = sheet.getRange(DASHBOARD_CELLS.PROJECT_CODE_INPUT).getValue();

  if (!projectCode) {
    clearDashboard();
    return;
  }

  // الحصول على بيانات المشروع
  const project = getProjectByCode(projectCode);
  if (!project) {
    showDashboardError('المشروع غير موجود');
    return;
  }

  // تحديث معلومات المشروع الأساسية
  updateProjectInfo(sheet, project);

  // تحديث نسب إكمال المراحل
  updatePhaseProgress(sheet, project, projectCode);

  // تحديث الإحصائيات
  updateDashboardStats(sheet, projectCode);

  // تحديث التنبيهات
  updateDashboardAlerts(sheet, projectCode);
}

/**
 * تحديث معلومات المشروع الأساسية
 * @param {Sheet} sheet - ورقة الداشبورد
 * @param {Object} project - بيانات المشروع (كائن من getAllProjects)
 */
function updateProjectInfo(sheet, project) {
  sheet.getRange(DASHBOARD_CELLS.PROJECT_NAME).setValue(project.name || '');
  sheet.getRange(DASHBOARD_CELLS.PROJECT_STATUS).setValue(project.status || '');
  sheet.getRange(DASHBOARD_CELLS.PROJECT_CLIENT).setValue(project.channel || '');
  sheet.getRange(DASHBOARD_CELLS.PROJECT_START).setValue(project.startDate || '');
  sheet.getRange(DASHBOARD_CELLS.PROJECT_DEADLINE).setValue(project.endDate || '');

  // تلوين حالة المشروع
  const statusColor = getStatusColor(project.status);
  if (statusColor) {
    sheet.getRange(DASHBOARD_CELLS.PROJECT_STATUS).setBackground(statusColor);
  }
}

/**
 * حساب نسبة إكمال مرحلة معينة
 * @param {string} projectCode - كود المشروع
 * @param {string} stageName - اسم المرحلة
 * @returns {Object} نسبة الإكمال وعدد المهام
 */
function calculatePhaseProgress(projectCode, stageName) {
  const movements = getMovementByProject(projectCode);
  // تصفية الحركات حسب المرحلة (مع مراعاة وجود الأيقونة)
  const stageMovements = movements.filter(m =>
    m.stage === stageName || m.stage.includes(stageName)
  );

  if (stageMovements.length === 0) {
    return { percentage: 0, total: 0, completed: 0 };
  }

  // التحقق من الحالة المكتملة (تم أو ✅)
  const completed = stageMovements.filter(m =>
    m.status.includes('تم') || m.status.includes('✅')
  ).length;

  return {
    percentage: Math.round((completed / stageMovements.length) * 100),
    total: stageMovements.length,
    completed: completed
  };
}

/**
 * تحديث نسب إكمال المراحل في الداشبورد
 * @param {Sheet} sheet - ورقة الداشبورد
 * @param {Object} project - بيانات المشروع
 * @param {string} projectCode - كود المشروع
 */
function updatePhaseProgress(sheet, project, projectCode) {
  const startRow = DASHBOARD_CELLS.PHASES_START_ROW;
  const startCol = DASHBOARD_CELLS.PHASES_START_COL;

  // الحصول على المراحل المفعلة للمشروع (مصفوفة من كائنات STAGES)
  const activePhases = getProjectPhases(projectCode);
  // استخراج أسماء المراحل المفعلة
  const activePhaseNames = activePhases.map(p => p.name);

  // مسح البيانات القديمة
  sheet.getRange(startRow, startCol, 12, 3).clearContent();

  let currentRow = startRow;

  // عرض كل مرحلة مفعلة
  Object.entries(STAGES).forEach(([key, stage]) => {
    if (activePhaseNames.includes(stage.name)) {
      const progress = calculatePhaseProgress(projectCode, stage.name);

      // اسم المرحلة مع الأيقونة
      sheet.getRange(currentRow, startCol).setValue(`${stage.icon} ${stage.name}`);

      // نسبة الإكمال
      sheet.getRange(currentRow, startCol + 1).setValue(`${progress.percentage}%`);

      // شريط التقدم (باستخدام التنسيق الشرطي)
      const progressCell = sheet.getRange(currentRow, startCol + 2);
      progressCell.setValue(progress.percentage / 100);
      progressCell.setNumberFormat('0%');

      // تلوين حسب النسبة
      let color = '#FFCDD2'; // أحمر فاتح للنسب المنخفضة
      if (progress.percentage >= 75) {
        color = '#C8E6C9'; // أخضر فاتح
      } else if (progress.percentage >= 50) {
        color = '#FFF9C4'; // أصفر فاتح
      } else if (progress.percentage >= 25) {
        color = '#FFE0B2'; // برتقالي فاتح
      }
      progressCell.setBackground(color);

      currentRow++;
    }
  });
}

/**
 * تحديث إحصائيات الداشبورد
 * @param {Sheet} sheet - ورقة الداشبورد
 * @param {string} projectCode - كود المشروع
 */
function updateDashboardStats(sheet, projectCode) {
  const startRow = DASHBOARD_CELLS.STATS_START_ROW;
  const startCol = DASHBOARD_CELLS.STATS_START_COL;

  // إحصائيات الضيوف (استخدام الخصائص المباشرة من كائنات الضيوف)
  const guests = getGuestsByProject(projectCode);
  const guestStats = {
    total: guests.length,
    confirmed: guests.filter(g =>
      g.shootStatus === SHOOT_STATUS.SCHEDULED ||
      g.shootStatus === SHOOT_STATUS.DONE
    ).length,
    pending: guests.filter(g =>
      g.contactStatus === CONTACT_STATUS.NOT_STARTED ||
      g.contactStatus === CONTACT_STATUS.IN_PROGRESS
    ).length
  };

  // إحصائيات التعليق الصوتي (stub - غير مفعل حالياً)
  const voiceOvers = getVoiceOverByProject(projectCode);
  const voStats = {
    total: voiceOvers.length,
    completed: voiceOvers.filter(v => v.status === 'مكتمل' || v.status === 'تم').length,
    totalDuration: getTotalVoiceOverDuration(projectCode)
  };

  // إحصائيات الرسوم المتحركة (stub - غير مفعل حالياً)
  const animations = getAnimationByProject(projectCode);
  const animStats = {
    total: animations.length,
    completed: animations.filter(a => a.status === 'مكتمل' || a.status === 'تم').length,
    totalDuration: getTotalAnimationDuration(projectCode)
  };

  // إحصائيات الأرشيف (stub - غير مفعل حالياً)
  const archive = getArchiveByProject(projectCode);
  const archiveStats = {
    total: archive.length,
    licensed: archive.filter(a => a.licenseStatus === LICENSE_STATUS.LICENSED).length,
    pending: archive.filter(a => a.licenseStatus === LICENSE_STATUS.PENDING).length
  };

  // مسح البيانات القديمة
  sheet.getRange(startRow, startCol, 22, 2).clearContent();

  // عرض الإحصائيات
  const stats = [
    ['📊 إحصائيات المشروع', ''],
    ['', ''],
    ['👥 الضيوف', ''],
    ['   إجمالي', guestStats.total],
    ['   مؤكد/مجدول', guestStats.confirmed],
    ['   في الانتظار', guestStats.pending],
    ['', ''],
    ['🎙️ التعليق الصوتي', ''],
    ['   إجمالي', voStats.total],
    ['   مكتمل', voStats.completed],
    ['   المدة الكلية', formatDuration(voStats.totalDuration)],
    ['', ''],
    ['🎬 الرسوم المتحركة', ''],
    ['   إجمالي', animStats.total],
    ['   مكتمل', animStats.completed],
    ['   المدة الكلية', formatDuration(animStats.totalDuration)],
    ['', ''],
    ['📁 الأرشيف', ''],
    ['   إجمالي', archiveStats.total],
    ['   مرخص', archiveStats.licensed],
    ['   في الانتظار', archiveStats.pending]
  ];

  sheet.getRange(startRow, startCol, stats.length, 2).setValues(stats);
}

/**
 * تنسيق المدة بالدقائق إلى صيغة مقروءة
 * @param {number} minutes - المدة بالدقائق
 * @returns {string} المدة بصيغة مقروءة
 */
function formatDuration(minutes) {
  if (!minutes || minutes === 0) return '0 د';

  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;

  if (hours > 0) {
    return `${hours} س ${mins} د`;
  }
  return `${mins} د`;
}

/**
 * الحصول على تنبيهات المشروع
 * @param {string} projectCode - كود المشروع
 * @returns {Array} مصفوفة بالتنبيهات
 */
function getProjectAlerts(projectCode) {
  const alerts = [];
  const today = new Date();

  // 1. المهام المتأخرة (تصفية بكود أو اسم المشروع)
  const delayedMovements = getDelayedMovements().filter(m =>
    m.projectCode === projectCode || m.projectName === projectCode || m.projectName.includes(projectCode)
  );
  if (delayedMovements.length > 0) {
    alerts.push({
      type: 'error',
      icon: '🔴',
      message: `${delayedMovements.length} مهمة متأخرة`,
      details: delayedMovements.slice(0, 3).map(m => m.element || m.action).join(', ')
    });
  }

  // 2. ضيوف بحاجة للمتابعة
  const allFollowupGuests = getGuestsNeedingFollowup();
  const followupGuests = allFollowupGuests.filter(g =>
    g.project === projectCode || g.project.includes(projectCode)
  );
  if (followupGuests.length > 0) {
    alerts.push({
      type: 'warning',
      icon: '🟡',
      message: `${followupGuests.length} ضيف بحاجة للمتابعة`,
      details: followupGuests.slice(0, 3).map(g => g.name).join(', ')
    });
  }

  // 3. تصوير قادم خلال أسبوع
  const upcomingShoots = getUpcomingShoots(7).filter(g =>
    g.project === projectCode || g.project.includes(projectCode)
  );
  if (upcomingShoots.length > 0) {
    alerts.push({
      type: 'info',
      icon: '🔵',
      message: `${upcomingShoots.length} تصوير قادم خلال أسبوع`,
      details: upcomingShoots.slice(0, 3).map(g => `${g.name} - ${formatDate(g.shootDate)}`).join(', ')
    });
  }

  // 4. تعليق صوتي معلق (stub - غير مفعل حالياً)
  const pendingVO = getPendingVoiceOver().filter(v => v.project === projectCode);
  if (pendingVO.length > 0) {
    alerts.push({
      type: 'warning',
      icon: '🟡',
      message: `${pendingVO.length} تعليق صوتي معلق`,
      details: ''
    });
  }

  // 5. رسوم متحركة معلقة (stub - غير مفعل حالياً)
  const pendingAnim = getPendingAnimation().filter(a => a.project === projectCode);
  if (pendingAnim.length > 0) {
    alerts.push({
      type: 'warning',
      icon: '🟡',
      message: `${pendingAnim.length} رسوم متحركة معلقة`,
      details: ''
    });
  }

  // 6. تراخيص في الانتظار (stub - غير مفعل حالياً)
  const pendingLicenses = getPendingLicenses().filter(a => a.project === projectCode);
  if (pendingLicenses.length > 0) {
    alerts.push({
      type: 'warning',
      icon: '🟡',
      message: `${pendingLicenses.length} ترخيص في الانتظار`,
      details: ''
    });
  }

  // 7. الموعد النهائي قريب
  const project = getProjectByCode(projectCode);
  if (project && project.endDate) {
    const deadline = new Date(project.endDate);
    const daysLeft = Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));

    if (daysLeft < 0) {
      alerts.push({
        type: 'error',
        icon: '🔴',
        message: `تجاوز الموعد النهائي بـ ${Math.abs(daysLeft)} يوم`,
        details: ''
      });
    } else if (daysLeft <= 7) {
      alerts.push({
        type: 'warning',
        icon: '🟠',
        message: `متبقي ${daysLeft} يوم على الموعد النهائي`,
        details: ''
      });
    }
  }

  return alerts;
}

/**
 * تحديث قسم التنبيهات في الداشبورد
 * @param {Sheet} sheet - ورقة الداشبورد
 * @param {string} projectCode - كود المشروع
 */
function updateDashboardAlerts(sheet, projectCode) {
  const startRow = DASHBOARD_CELLS.ALERTS_START_ROW;
  const startCol = DASHBOARD_CELLS.ALERTS_START_COL;

  // مسح التنبيهات القديمة
  sheet.getRange(startRow, startCol, 15, 5).clearContent().setBackground(null);

  const alerts = getProjectAlerts(projectCode);

  // عنوان قسم التنبيهات
  sheet.getRange(startRow, startCol).setValue('⚠️ التنبيهات').setFontWeight('bold');

  if (alerts.length === 0) {
    sheet.getRange(startRow + 1, startCol).setValue('✅ لا توجد تنبيهات').setFontColor('#4CAF50');
    return;
  }

  // عرض التنبيهات
  alerts.forEach((alert, index) => {
    const row = startRow + 1 + index;
    const alertRange = sheet.getRange(row, startCol, 1, 4);

    alertRange.getCell(1, 1).setValue(alert.icon);
    alertRange.getCell(1, 2).setValue(alert.message);

    if (alert.details) {
      alertRange.getCell(1, 3).setValue(alert.details).setFontSize(9).setFontColor('#666');
    }

    // تلوين حسب نوع التنبيه
    let bgColor = '#FFF9C4'; // أصفر فاتح افتراضي
    if (alert.type === 'error') {
      bgColor = '#FFCDD2'; // أحمر فاتح
    } else if (alert.type === 'info') {
      bgColor = '#BBDEFB'; // أزرق فاتح
    }
    alertRange.setBackground(bgColor);
  });
}

/**
 * مسح الداشبورد
 */
function clearDashboard() {
  const sheet = getSheet(SHEETS.DASHBOARD);
  if (!sheet) return;

  // مسح معلومات المشروع
  sheet.getRange(DASHBOARD_CELLS.PROJECT_NAME).clearContent();
  sheet.getRange(DASHBOARD_CELLS.PROJECT_STATUS).clearContent().setBackground(null);
  sheet.getRange(DASHBOARD_CELLS.PROJECT_CLIENT).clearContent();
  sheet.getRange(DASHBOARD_CELLS.PROJECT_START).clearContent();
  sheet.getRange(DASHBOARD_CELLS.PROJECT_DEADLINE).clearContent();

  // مسح نسب الإكمال
  const phasesRange = sheet.getRange(
    DASHBOARD_CELLS.PHASES_START_ROW,
    DASHBOARD_CELLS.PHASES_START_COL,
    12, 3
  );
  phasesRange.clearContent().setBackground(null);

  // مسح الإحصائيات
  const statsRange = sheet.getRange(
    DASHBOARD_CELLS.STATS_START_ROW,
    DASHBOARD_CELLS.STATS_START_COL,
    25, 2
  );
  statsRange.clearContent();

  // مسح التنبيهات
  const alertsRange = sheet.getRange(
    DASHBOARD_CELLS.ALERTS_START_ROW,
    DASHBOARD_CELLS.ALERTS_START_COL,
    15, 5
  );
  alertsRange.clearContent().setBackground(null);
}

/**
 * عرض رسالة خطأ في الداشبورد
 * @param {string} message - رسالة الخطأ
 */
function showDashboardError(message) {
  const sheet = getSheet(SHEETS.DASHBOARD);
  if (!sheet) return;

  clearDashboard();
  sheet.getRange(DASHBOARD_CELLS.PROJECT_NAME)
    .setValue(message)
    .setFontColor('#D32F2F');
}

// ====================================================
// دوال الملخص العام
// ====================================================

/**
 * الحصول على ملخص جميع المشاريع النشطة
 * @returns {Array} مصفوفة بملخصات المشاريع
 */
function getAllProjectsSummary() {
  const projects = getActiveProjects();

  return projects.map(project => {
    const code = project.code;
    const movements = getMovementByProject(code);
    const totalTasks = movements.length;
    // التحقق من الحالة المكتملة (تم أو ✅)
    const completedTasks = movements.filter(m =>
      m.status.includes('تم') || m.status.includes('✅')
    ).length;

    return {
      code: code,
      name: project.name,
      status: project.status,
      client: project.channel,
      deadline: project.endDate,
      progress: totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0,
      totalTasks: totalTasks,
      completedTasks: completedTasks,
      alertsCount: getProjectAlerts(code).length
    };
  });
}

/**
 * تحديث ورقة الملخص العام لجميع المشاريع
 */
function updateAllProjectsDashboard() {
  const sheet = getSheet(SHEETS.DASHBOARD);
  if (!sheet) return;

  const summary = getAllProjectsSummary();

  // إنشاء قسم منفصل لملخص جميع المشاريع (في أسفل الداشبورد)
  const startRow = 40;
  const startCol = 2;

  // مسح البيانات القديمة
  sheet.getRange(startRow, startCol, 30, 8).clearContent().setBackground(null);

  // العنوان
  sheet.getRange(startRow, startCol).setValue('📋 ملخص جميع المشاريع النشطة').setFontWeight('bold').setFontSize(12);

  // العناوين
  const headers = ['الكود', 'الاسم', 'الحالة', 'العميل', 'الموعد النهائي', 'التقدم', 'المهام', 'التنبيهات'];
  sheet.getRange(startRow + 2, startCol, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#E3F2FD');

  // البيانات
  summary.forEach((project, index) => {
    const row = startRow + 3 + index;
    const values = [
      project.code,
      project.name,
      project.status,
      project.client,
      project.deadline ? formatDate(project.deadline) : '',
      `${project.progress}%`,
      `${project.completedTasks}/${project.totalTasks}`,
      project.alertsCount > 0 ? `⚠️ ${project.alertsCount}` : '✅'
    ];

    sheet.getRange(row, startCol, 1, values.length).setValues([values]);

    // تلوين نسبة التقدم
    let progressColor = '#FFCDD2';
    if (project.progress >= 75) progressColor = '#C8E6C9';
    else if (project.progress >= 50) progressColor = '#FFF9C4';
    else if (project.progress >= 25) progressColor = '#FFE0B2';

    sheet.getRange(row, startCol + 5).setBackground(progressColor);

    // تلوين حالة المشروع
    const statusColor = getStatusColor(project.status);
    if (statusColor) {
      sheet.getRange(row, startCol + 2).setBackground(statusColor);
    }
  });
}

// ====================================================
// دوال مساعدة
// ====================================================

// ملاحظة: دالة formatDate موجودة في 03-Helpers.gs

/**
 * تحديث الداشبورد عند تغيير كود المشروع
 * يُستدعى من onEdit
 * @param {Event} e - حدث التعديل
 */
function onDashboardEdit(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== SHEETS.DASHBOARD) return;

  const range = e.range;
  const cell = range.getA1Notation();

  // التحقق من أن الخلية المعدلة هي خلية كود المشروع
  if (cell === DASHBOARD_CELLS.PROJECT_CODE_INPUT) {
    refreshDashboard();
  }
}

/**
 * إعداد الداشبورد
 * يُنشئ الهيكل الأساسي للداشبورد
 */
function setupDashboard() {
  const sheet = getSheet(SHEETS.DASHBOARD);
  if (!sheet) return;

  // مسح الورقة
  sheet.clear();

  // إعداد العناوين
  sheet.getRange('A2').setValue('كود المشروع:').setFontWeight('bold');
  sheet.getRange('C2').setValue('الاسم:').setFontWeight('bold');
  sheet.getRange('E2').setValue('الحالة:').setFontWeight('bold');

  sheet.getRange('A4').setValue('العميل:').setFontWeight('bold');
  sheet.getRange('C4').setValue('تاريخ البداية:').setFontWeight('bold');
  sheet.getRange('E4').setValue('الموعد النهائي:').setFontWeight('bold');

  // عنوان نسب الإكمال
  sheet.getRange('B6').setValue('📊 نسب إكمال المراحل').setFontWeight('bold');

  // عنوان الإحصائيات
  sheet.getRange('E6').setValue('📈 الإحصائيات').setFontWeight('bold');

  // إعداد القائمة المنسدلة لكود المشروع
  const projects = getActiveProjects();
  const projectCodes = projects.map(p => p.code);

  if (projectCodes.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(projectCodes, true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(DASHBOARD_CELLS.PROJECT_CODE_INPUT).setDataValidation(rule);
  }

  // تنسيق الورقة
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);

  // تعيين الاتجاه من اليمين لليسار
  sheet.setRightToLeft(true);
}
