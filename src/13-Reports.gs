/**
 * ===================================================
 * 13-Reports.gs - التقارير
 * ===================================================
 *
 * هذا الملف يحتوي على:
 * - تقرير الشخص (المهام حسب الشخص)
 * - تقرير المهام المتأخرة
 * - تقرير المواعيد القادمة
 * - ملخص المشروع
 * - تصدير التقارير
 */

// ====================================================
// ثوابت التقارير
// ====================================================

/**
 * أنواع التقارير
 */
const REPORT_TYPES = {
  PERSON: 'تقرير الشخص',
  OVERDUE: 'المهام المتأخرة',
  UPCOMING: 'المواعيد القادمة',
  PROJECT_SUMMARY: 'ملخص المشروع',
  WEEKLY: 'تقرير أسبوعي',
  MONTHLY: 'تقرير شهري'
};

/**
 * ألوان حالات المهام في التقارير
 */
const TASK_COLORS = {
  OVERDUE: '#FFCDD2',      // أحمر فاتح - متأخر
  IN_PROGRESS: '#FFF9C4',  // أصفر فاتح - قيد العمل
  COMPLETED: '#C8E6C9',    // أخضر فاتح - مكتمل
  PENDING: '#E3F2FD'       // أزرق فاتح - في الانتظار
};

// ====================================================
// تقرير الشخص
// ====================================================

/**
 * توليد تقرير شخص معين
 * @param {string} personName - اسم الشخص
 * @returns {Object} بيانات التقرير
 */
function generatePersonReport(personName) {
  const report = {
    person: personName,
    generatedAt: new Date(),
    movements: [],
    guests: [],
    voiceOvers: [],
    animations: [],
    summary: {
      totalTasks: 0,
      completed: 0,
      inProgress: 0,
      overdue: 0,
      pending: 0
    }
  };

  const today = new Date();

  // الحصول على مهام الشخص من ورقة الحركة
  const movements = getMovementByPerson(personName);

  movements.forEach(movement => {
    const deadline = movement.dueDate;
    const status = movement.status || '';

    let taskStatus = 'pending';
    let color = TASK_COLORS.PENDING;

    if (status.includes('تم') || status.includes('✅')) {
      taskStatus = 'completed';
      color = TASK_COLORS.COMPLETED;
      report.summary.completed++;
    } else if (status.includes('جاري') || status.includes('🔄')) {
      taskStatus = 'inProgress';
      color = TASK_COLORS.IN_PROGRESS;
      report.summary.inProgress++;

      // التحقق من التأخير
      if (deadline && new Date(deadline) < today) {
        taskStatus = 'overdue';
        color = TASK_COLORS.OVERDUE;
        report.summary.overdue++;
        report.summary.inProgress--;
      }
    } else if (deadline && new Date(deadline) < today && !status.includes('تم') && !status.includes('✅')) {
      taskStatus = 'overdue';
      color = TASK_COLORS.OVERDUE;
      report.summary.overdue++;
    } else {
      report.summary.pending++;
    }

    report.movements.push({
      ...movement,
      taskStatus: taskStatus,
      color: color
    });

    report.summary.totalTasks++;
  });

  // ترتيب المهام: المتأخرة أولاً، ثم قيد العمل، ثم في الانتظار، ثم المكتملة
  const statusOrder = { overdue: 0, inProgress: 1, pending: 2, completed: 3 };
  report.movements.sort((a, b) => {
    return statusOrder[a.taskStatus] - statusOrder[b.taskStatus];
  });

  return report;
}

/**
 * عرض تقرير الشخص في ورقة التقارير
 * @param {string} personName - اسم الشخص (اختياري، إذا لم يحدد يطلب من المستخدم)
 */
function showPersonReport(personName = null) {
  const sheet = getSheet(SHEETS.PERSON_REPORT);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('ورقة تقرير الشخص غير موجودة');
    return;
  }

  // إذا لم يحدد الشخص، نحصل عليه من خلية الإدخال
  if (!personName) {
    personName = sheet.getRange('B2').getValue();
    if (!personName) {
      SpreadsheetApp.getUi().alert('يرجى إدخال اسم الشخص في الخلية B2');
      return;
    }
  }

  const report = generatePersonReport(personName);

  // مسح البيانات القديمة
  sheet.getRange('A5:H100').clearContent().setBackground(null);

  // عنوان التقرير
  sheet.getRange('A1').setValue(`📋 تقرير: ${personName}`).setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setValue(`تاريخ التقرير: ${formatDate(report.generatedAt)}`);

  // ملخص
  sheet.getRange('D1').setValue('الملخص:').setFontWeight('bold');
  sheet.getRange('D2').setValue(`إجمالي المهام: ${report.summary.totalTasks}`);
  sheet.getRange('E2').setValue(`مكتمل: ${report.summary.completed}`).setBackground(TASK_COLORS.COMPLETED);
  sheet.getRange('F2').setValue(`قيد العمل: ${report.summary.inProgress}`).setBackground(TASK_COLORS.IN_PROGRESS);
  sheet.getRange('G2').setValue(`متأخر: ${report.summary.overdue}`).setBackground(TASK_COLORS.OVERDUE);
  sheet.getRange('H2').setValue(`في الانتظار: ${report.summary.pending}`).setBackground(TASK_COLORS.PENDING);

  // عناوين الجدول
  const headers = ['المشروع', 'المرحلة', 'المهمة', 'الحالة', 'الموعد النهائي', 'الأولوية', 'ملاحظات'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#E3F2FD');

  // بيانات المهام
  let currentRow = 6;
  report.movements.forEach(movement => {
    const values = [
      movement.projectCode || movement.projectName || '',
      movement.stage || '',
      movement.element || movement.action || '',
      movement.status || '',
      movement.dueDate ? formatDate(movement.dueDate) : '',
      '', // الأولوية - غير متاحة
      movement.notes || ''
    ];

    const rowRange = sheet.getRange(currentRow, 1, 1, values.length);
    rowRange.setValues([values]);
    rowRange.setBackground(movement.color);

    currentRow++;
  });

  // تنسيق الأعمدة
  sheet.autoResizeColumns(1, headers.length);
}

// ====================================================
// تقرير المهام المتأخرة
// ====================================================

/**
 * الحصول على جميع المهام المتأخرة
 * @returns {Array} مصفوفة بالمهام المتأخرة
 */
function getOverdueTasks() {
  const allMovements = getAllMovements();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return allMovements.filter(movement => {
    const deadline = movement.dueDate;
    const status = movement.status || '';

    if (!deadline) return false;
    // التحقق من الحالة - مكتمل أو ملغي
    if (status.includes('تم') || status.includes('✅') || status.includes('ملغي') || status.includes('❌')) {
      return false;
    }

    const deadlineDate = new Date(deadline);
    deadlineDate.setHours(0, 0, 0, 0);

    return deadlineDate < today;
  }).sort((a, b) => {
    return new Date(a.dueDate) - new Date(b.dueDate);
  });
}

/**
 * عرض تقرير المهام المتأخرة
 */
function showOverdueReport() {
  const overdueTasks = getOverdueTasks();

  if (overdueTasks.length === 0) {
    SpreadsheetApp.getUi().alert('✅ لا توجد مهام متأخرة!');
    return;
  }

  const sheet = getSheet(SHEETS.PERSON_REPORT);
  if (!sheet) return;

  // مسح البيانات القديمة
  sheet.clear();

  // عنوان التقرير
  sheet.getRange('A1').setValue(`🔴 تقرير المهام المتأخرة`).setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue(`تاريخ التقرير: ${formatDate(new Date())}`);
  sheet.getRange('A3').setValue(`عدد المهام المتأخرة: ${overdueTasks.length}`).setFontColor('#D32F2F');

  // عناوين الجدول
  const headers = ['المشروع', 'المرحلة', 'المهمة', 'المسؤول', 'الموعد النهائي', 'أيام التأخير', 'الأولوية'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#FFCDD2');

  // بيانات المهام
  const today = new Date();
  let currentRow = 6;

  overdueTasks.forEach(task => {
    const deadline = new Date(task.dueDate);
    const daysOverdue = Math.ceil((today - deadline) / (1000 * 60 * 60 * 24));

    const values = [
      task.projectCode || task.projectName || '',
      task.stage || '',
      task.element || task.action || '',
      task.assignedTo || '',
      formatDate(task.dueDate),
      `${daysOverdue} يوم`,
      '' // الأولوية - غير متاحة
    ];

    const rowRange = sheet.getRange(currentRow, 1, 1, values.length);
    rowRange.setValues([values]);
    rowRange.setBackground(TASK_COLORS.OVERDUE);

    currentRow++;
  });

  sheet.autoResizeColumns(1, headers.length);
}

// ====================================================
// تقرير المواعيد القادمة
// ====================================================

/**
 * الحصول على المواعيد القادمة
 * @param {number} daysAhead - عدد الأيام القادمة (افتراضي 7)
 * @returns {Array} مصفوفة بالمواعيد القادمة
 */
function getUpcomingDeadlines(daysAhead = 7) {
  const allMovements = getAllMovements();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const futureDate = new Date(today.getTime() + (daysAhead * 24 * 60 * 60 * 1000));

  return allMovements.filter(movement => {
    const deadline = movement.dueDate;
    const status = movement.status || '';

    if (!deadline) return false;
    // التحقق من الحالة - مكتمل أو ملغي
    if (status.includes('تم') || status.includes('✅') || status.includes('ملغي') || status.includes('❌')) {
      return false;
    }

    const deadlineDate = new Date(deadline);
    deadlineDate.setHours(0, 0, 0, 0);

    return deadlineDate >= today && deadlineDate <= futureDate;
  }).sort((a, b) => {
    return new Date(a.dueDate) - new Date(b.dueDate);
  });
}

/**
 * عرض تقرير المواعيد القادمة
 * @param {number} days - عدد الأيام (افتراضي 7)
 */
function showUpcomingDeadlinesReport(days = 7) {
  const upcoming = getUpcomingDeadlines(days);

  if (upcoming.length === 0) {
    SpreadsheetApp.getUi().alert(`لا توجد مواعيد نهائية خلال ${days} أيام القادمة`);
    return;
  }

  const sheet = getSheet(SHEETS.PERSON_REPORT);
  if (!sheet) return;

  sheet.clear();

  // عنوان التقرير
  sheet.getRange('A1').setValue(`📅 المواعيد القادمة (${days} أيام)`).setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue(`تاريخ التقرير: ${formatDate(new Date())}`);
  sheet.getRange('A3').setValue(`عدد المواعيد: ${upcoming.length}`);

  // عناوين الجدول
  const headers = ['المشروع', 'المرحلة', 'المهمة', 'المسؤول', 'الموعد النهائي', 'أيام متبقية', 'الحالة'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#E3F2FD');

  const today = new Date();
  let currentRow = 6;

  upcoming.forEach(task => {
    const deadline = new Date(task.dueDate);
    const daysLeft = Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));

    const values = [
      task.projectCode || task.projectName || '',
      task.stage || '',
      task.element || task.action || '',
      task.assignedTo || '',
      formatDate(task.dueDate),
      daysLeft === 0 ? 'اليوم!' : `${daysLeft} يوم`,
      task.status || ''
    ];

    const rowRange = sheet.getRange(currentRow, 1, 1, values.length);
    rowRange.setValues([values]);

    // تلوين حسب الأيام المتبقية
    let bgColor = '#C8E6C9'; // أخضر - أكثر من 3 أيام
    if (daysLeft === 0) {
      bgColor = '#FFCDD2'; // أحمر - اليوم
    } else if (daysLeft <= 2) {
      bgColor = '#FFE0B2'; // برتقالي - يومين أو أقل
    } else if (daysLeft <= 5) {
      bgColor = '#FFF9C4'; // أصفر - 3-5 أيام
    }
    rowRange.setBackground(bgColor);

    currentRow++;
  });

  sheet.autoResizeColumns(1, headers.length);
}

// ====================================================
// ملخص المشروع
// ====================================================

/**
 * توليد ملخص شامل للمشروع
 * @param {string} projectCode - كود المشروع
 * @returns {Object} ملخص المشروع
 */
function generateProjectSummary(projectCode) {
  const project = getProjectByCode(projectCode);
  if (!project) return null;

  const movements = getMovementByProject(projectCode);
  const guests = getGuestsByProject(projectCode);
  const voiceOvers = getVoiceOverByProject(projectCode);
  const animations = getAnimationByProject(projectCode);
  const archive = getArchiveByProject(projectCode);

  const today = new Date();

  // حساب إحصائيات المهام
  const taskStats = {
    total: movements.length,
    completed: movements.filter(m => {
      const status = m.status || '';
      return status.includes('تم') || status.includes('✅');
    }).length,
    inProgress: movements.filter(m => {
      const status = m.status || '';
      return status.includes('جاري') || status.includes('🔄');
    }).length,
    overdue: movements.filter(m => {
      const deadline = m.dueDate;
      const status = m.status || '';
      return deadline && new Date(deadline) < today && !status.includes('تم') && !status.includes('✅');
    }).length,
    delayed: movements.filter(m => {
      const status = m.status || '';
      return status.includes('متأخر') || status.includes('⏰');
    }).length
  };

  // حساب نسبة الإكمال الكلية
  taskStats.completionRate = taskStats.total > 0
    ? Math.round((taskStats.completed / taskStats.total) * 100)
    : 0;

  // إحصائيات الضيوف
  const guestStats = {
    total: guests.length,
    contacted: guests.filter(g => {
      const status = g.contactStatus || '';
      return status.includes('تم التواصل') || status.includes('✅');
    }).length,
    confirmed: guests.filter(g => {
      const status = g.shootStatus || '';
      return status.includes('مجدول') || status.includes('مكتمل') || status.includes('✅');
    }).length,
    shotCompleted: guests.filter(g => {
      const status = g.shootStatus || '';
      return status.includes('مكتمل') || status.includes('✅');
    }).length
  };

  // إحصائيات التعليق الصوتي
  const voStats = {
    total: voiceOvers.length,
    completed: voiceOvers.filter(v => {
      const status = v.status || '';
      return status.includes('مكتمل') || status.includes('✅');
    }).length,
    totalDuration: getTotalVoiceOverDuration(projectCode)
  };

  // إحصائيات الرسوم المتحركة
  const animStats = {
    total: animations.length,
    completed: animations.filter(a => {
      const status = a.status || '';
      return status.includes('مكتمل') || status.includes('✅');
    }).length,
    totalDuration: getTotalAnimationDuration(projectCode)
  };

  // إحصائيات الأرشيف
  const archiveStats = {
    total: archive.length,
    licensed: archive.filter(a => {
      const status = a.licenseStatus || '';
      return status.includes('مرخص') || status.includes('✅');
    }).length,
    totalCost: getProjectLicenseCost(projectCode)
  };

  // حساب الأيام المتبقية
  let daysRemaining = null;
  if (project.deadline) {
    const deadline = new Date(project.deadline);
    daysRemaining = Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));
  }

  return {
    project: {
      code: projectCode,
      name: project.name || '',
      client: project.client || '',
      status: project.status || '',
      startDate: project.startDate || '',
      deadline: project.deadline || '',
      daysRemaining: daysRemaining
    },
    tasks: taskStats,
    guests: guestStats,
    voiceOver: voStats,
    animation: animStats,
    archive: archiveStats,
    alerts: getProjectAlerts(projectCode),
    generatedAt: new Date()
  };
}

/**
 * عرض ملخص المشروع
 * @param {string} projectCode - كود المشروع (اختياري)
 */
function showProjectSummary(projectCode = null) {
  // إذا لم يحدد الكود، نسأل المستخدم
  if (!projectCode) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('ملخص المشروع', 'أدخل كود المشروع:', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK) return;
    projectCode = response.getResponseText().trim();
  }

  const summary = generateProjectSummary(projectCode);

  if (!summary) {
    SpreadsheetApp.getUi().alert('المشروع غير موجود');
    return;
  }

  const sheet = getSheet(SHEETS.PERSON_REPORT);
  if (!sheet) return;

  sheet.clear();

  // معلومات المشروع الأساسية
  sheet.getRange('A1').setValue(`📋 ملخص المشروع: ${summary.project.code}`).setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue(`الاسم: ${summary.project.name}`);
  sheet.getRange('A3').setValue(`العميل: ${summary.project.client}`);
  sheet.getRange('A4').setValue(`الحالة: ${summary.project.status}`);

  sheet.getRange('D2').setValue(`تاريخ البداية: ${formatDate(summary.project.startDate)}`);
  sheet.getRange('D3').setValue(`الموعد النهائي: ${formatDate(summary.project.deadline)}`);

  if (summary.project.daysRemaining !== null) {
    const daysText = summary.project.daysRemaining < 0
      ? `متأخر بـ ${Math.abs(summary.project.daysRemaining)} يوم`
      : `متبقي ${summary.project.daysRemaining} يوم`;
    const daysColor = summary.project.daysRemaining < 0 ? '#D32F2F' :
                     summary.project.daysRemaining <= 7 ? '#FF6D00' : '#4CAF50';
    sheet.getRange('D4').setValue(daysText).setFontColor(daysColor).setFontWeight('bold');
  }

  // إحصائيات المهام
  let row = 6;
  sheet.getRange(row, 1).setValue('📊 إحصائيات المهام').setFontWeight('bold');
  row++;

  const taskData = [
    ['إجمالي المهام', summary.tasks.total],
    ['مكتمل', summary.tasks.completed],
    ['قيد العمل', summary.tasks.inProgress],
    ['متأخر', summary.tasks.overdue],
    ['معلق', summary.tasks.delayed],
    ['نسبة الإكمال', `${summary.tasks.completionRate}%`]
  ];

  sheet.getRange(row, 1, taskData.length, 2).setValues(taskData);
  sheet.getRange(row + 5, 2).setBackground(
    summary.tasks.completionRate >= 75 ? '#C8E6C9' :
    summary.tasks.completionRate >= 50 ? '#FFF9C4' :
    summary.tasks.completionRate >= 25 ? '#FFE0B2' : '#FFCDD2'
  );

  row += taskData.length + 1;

  // إحصائيات الضيوف
  sheet.getRange(row, 1).setValue('👥 الضيوف').setFontWeight('bold');
  row++;

  const guestData = [
    ['إجمالي الضيوف', summary.guests.total],
    ['تم التواصل', summary.guests.contacted],
    ['مؤكد للتصوير', summary.guests.confirmed],
    ['تم التصوير', summary.guests.shotCompleted]
  ];

  sheet.getRange(row, 1, guestData.length, 2).setValues(guestData);
  row += guestData.length + 1;

  // إحصائيات الإنتاج (في العمود D)
  let prodRow = 6;
  sheet.getRange(prodRow, 4).setValue('🎙️ التعليق الصوتي').setFontWeight('bold');
  prodRow++;

  const voData = [
    ['إجمالي', summary.voiceOver.total],
    ['مكتمل', summary.voiceOver.completed],
    ['المدة الكلية', formatDuration(summary.voiceOver.totalDuration)]
  ];
  sheet.getRange(prodRow, 4, voData.length, 2).setValues(voData);
  prodRow += voData.length + 1;

  sheet.getRange(prodRow, 4).setValue('🎬 الرسوم المتحركة').setFontWeight('bold');
  prodRow++;

  const animData = [
    ['إجمالي', summary.animation.total],
    ['مكتمل', summary.animation.completed],
    ['المدة الكلية', formatDuration(summary.animation.totalDuration)]
  ];
  sheet.getRange(prodRow, 4, animData.length, 2).setValues(animData);
  prodRow += animData.length + 1;

  sheet.getRange(prodRow, 4).setValue('📁 الأرشيف').setFontWeight('bold');
  prodRow++;

  const archiveData = [
    ['إجمالي المواد', summary.archive.total],
    ['مرخص', summary.archive.licensed],
    ['تكلفة التراخيص', `$${summary.archive.totalCost}`]
  ];
  sheet.getRange(prodRow, 4, archiveData.length, 2).setValues(archiveData);

  // التنبيهات
  row = Math.max(row, prodRow) + 2;
  sheet.getRange(row, 1).setValue('⚠️ التنبيهات').setFontWeight('bold');
  row++;

  if (summary.alerts.length === 0) {
    sheet.getRange(row, 1).setValue('✅ لا توجد تنبيهات').setFontColor('#4CAF50');
  } else {
    summary.alerts.forEach(alert => {
      sheet.getRange(row, 1).setValue(`${alert.icon} ${alert.message}`);
      row++;
    });
  }

  // تنسيق الورقة
  sheet.setRightToLeft(true);
  sheet.autoResizeColumns(1, 6);
}

// ====================================================
// تقرير أسبوعي
// ====================================================

/**
 * توليد تقرير أسبوعي
 * @param {string} projectCode - كود المشروع (اختياري لجميع المشاريع)
 * @returns {Object} بيانات التقرير الأسبوعي
 */
function generateWeeklyReport(projectCode = null) {
  const today = new Date();
  const weekStart = new Date(today);
  weekStart.setDate(today.getDate() - today.getDay()); // بداية الأسبوع (الأحد)
  weekStart.setHours(0, 0, 0, 0);

  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  weekEnd.setHours(23, 59, 59, 999);

  // الحصول على جميع الحركات
  let movements = getAllMovements();
  if (projectCode) {
    movements = movements.filter(m => m.projectCode === projectCode || m.projectName === projectCode);
  }

  // المهام المكتملة هذا الأسبوع
  // ملاحظة: لا يوجد عمود تاريخ التحديث، لذا نعرض جميع المهام المكتملة
  const completedThisWeek = movements.filter(m => {
    const status = m.status || '';
    return status.includes('تم') || status.includes('✅');
  });

  // المهام الجديدة هذا الأسبوع
  // ملاحظة: لا يوجد عمود تاريخ الإنشاء، لذا نعرض قائمة فارغة
  const newThisWeek = [];

  // المهام المتأخرة
  const overdue = movements.filter(m => {
    const deadline = m.dueDate;
    const status = m.status || '';
    return deadline && new Date(deadline) < today && !status.includes('تم') && !status.includes('✅');
  });

  // المواعيد القادمة الأسبوع القادم
  const nextWeekStart = new Date(weekEnd);
  nextWeekStart.setDate(nextWeekStart.getDate() + 1);
  const nextWeekEnd = new Date(nextWeekStart);
  nextWeekEnd.setDate(nextWeekStart.getDate() + 6);

  const upcomingNextWeek = movements.filter(m => {
    const deadline = m.dueDate;
    const status = m.status || '';
    return deadline && new Date(deadline) >= nextWeekStart && new Date(deadline) <= nextWeekEnd &&
           !status.includes('تم') && !status.includes('✅');
  });

  return {
    period: {
      start: weekStart,
      end: weekEnd
    },
    projectCode: projectCode || 'جميع المشاريع',
    completedThisWeek: completedThisWeek,
    newThisWeek: newThisWeek,
    overdue: overdue,
    upcomingNextWeek: upcomingNextWeek,
    summary: {
      completed: completedThisWeek.length,
      new: newThisWeek.length,
      overdue: overdue.length,
      upcoming: upcomingNextWeek.length
    },
    generatedAt: new Date()
  };
}

// ====================================================
// دوال التصدير
// ====================================================

/**
 * تصدير التقرير الحالي إلى PDF
 */
function exportReportToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PERSON_REPORT);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ورقة التقرير غير موجودة');
    return;
  }

  const url = ss.getUrl()
    .replace(/edit$/, '')
    .replace(/edit\?.*$/, '')
    + 'export?format=pdf'
    + '&gid=' + sheet.getSheetId()
    + '&size=A4'
    + '&portrait=false'
    + '&fitw=true'
    + '&gridlines=false';

  SpreadsheetApp.getUi().alert(
    'لتصدير التقرير كـ PDF:\n\n' +
    '1. انقر على ملف > تنزيل > PDF\n' +
    '2. أو افتح الرابط التالي:\n' + url
  );
}

/**
 * سجل عملية التصدير
 * @param {string} reportType - نوع التقرير
 * @param {string} details - تفاصيل إضافية
 */
function logExport(reportType, details = '') {
  const sheet = getSheet(SHEETS.EXPORT_LOG);
  if (!sheet) return;

  const user = Session.getActiveUser().getEmail() || 'غير معروف';
  const timestamp = new Date();

  sheet.appendRow([timestamp, reportType, details, user]);
}

// ====================================================
// إعداد ورقة تقرير الشخص
// ====================================================

/**
 * إعداد ورقة تقرير الشخص
 */
function setupPersonReportSheet() {
  const sheet = getSheet(SHEETS.PERSON_REPORT);
  if (!sheet) return;

  sheet.clear();

  // إعداد خلية إدخال اسم الشخص
  sheet.getRange('A2').setValue('اسم الشخص:').setFontWeight('bold');

  // إنشاء قائمة منسدلة بأسماء الفريق
  const team = getTeamMembers();
  const photographers = getPhotographers();
  const names = [...team.map(t => t.name || ''), ...photographers.map(p => p.name || '')].filter(n => n);

  if (names.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true)
      .setAllowInvalid(true)
      .build();

    sheet.getRange('B2').setDataValidation(rule);
  }

  // زر لتوليد التقرير
  sheet.getRange('C2').setValue('← أدخل الاسم ثم شغّل: نظام الإنتاج > التقارير > تقرير الشخص');

  sheet.setRightToLeft(true);
}
