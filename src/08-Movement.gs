/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف شيت الحركة (الإدخال الرئيسي)
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على جميع الدوال المتعلقة بشيت الحركة
 * وهو الشيت الرئيسي لتسجيل جميع الأنشطة والمهام اليومية
 */

// ═══════════════════════════════════════════════════════════════════════════════
// ثوابت أعمدة شيت الحركة
// ═══════════════════════════════════════════════════════════════════════════════

const MOVEMENT_COLS = {
  NUMBER: 1,          // #
  DATE: 2,            // التاريخ (auto-fill)
  PROJECT_CODE: 3,    // كود المشروع (dropdown)
  PROJECT_NAME: 4,    // اسم المشروع (dropdown)
  STAGE: 5,           // المرحلة (dropdown ديناميكي)
  SUBTYPE: 6,         // النوع الفرعي (dropdown ديناميكي)
  ELEMENT: 7,         // العنصر
  ACTION: 8,          // الإجراء
  ASSIGNED_TO: 9,     // المسؤول (dropdown من الفريق)
  STATUS: 10,         // الحالة (dropdown)
  DUE_DATE: 11,       // تاريخ الاستحقاق
  NOTES: 12,          // ملاحظات
  CREATED_BY: 13,     // أنشئ بواسطة
  CREATED_AT: 14      // تاريخ الإنشاء
};

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الاستعلام عن الحركات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على جميع الحركات
 * @returns {Array} مصفوفة من كائنات الحركات
 */
function getAllMovements() {
  const sheet = getSheet(SHEETS.MOVEMENT);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  return data.map(row => ({
    number: row[MOVEMENT_COLS.NUMBER - 1],
    date: row[MOVEMENT_COLS.DATE - 1],
    projectCode: row[MOVEMENT_COLS.PROJECT_CODE - 1],
    projectName: row[MOVEMENT_COLS.PROJECT_NAME - 1],
    stage: row[MOVEMENT_COLS.STAGE - 1],
    subtype: row[MOVEMENT_COLS.SUBTYPE - 1],
    element: row[MOVEMENT_COLS.ELEMENT - 1],
    action: row[MOVEMENT_COLS.ACTION - 1],
    assignedTo: row[MOVEMENT_COLS.ASSIGNED_TO - 1],
    status: row[MOVEMENT_COLS.STATUS - 1],
    dueDate: row[MOVEMENT_COLS.DUE_DATE - 1],
    notes: row[MOVEMENT_COLS.NOTES - 1],
    createdBy: row[MOVEMENT_COLS.CREATED_BY - 1],
    createdAt: row[MOVEMENT_COLS.CREATED_AT - 1]
  })).filter(movement => movement.number);
}

/**
 * الحصول على حركات مشروع معين
 * @param {string} projectCode كود المشروع أو اسمه
 * @returns {Array} مصفوفة الحركات
 */
function getMovementByProject(projectCode) {
  if (!projectCode) return [];

  const allMovements = getAllMovements();
  return allMovements.filter(movement =>
    movement.project === projectCode ||
    movement.project.includes(projectCode)
  );
}

/**
 * الحصول على حركات شخص معين
 * @param {string} personName اسم الشخص
 * @returns {Array} مصفوفة الحركات
 */
function getMovementByPerson(personName) {
  if (!personName) return [];

  const allMovements = getAllMovements();
  return allMovements.filter(movement =>
    movement.assignedTo === personName ||
    movement.assignedTo.includes(personName)
  );
}

/**
 * الحصول على حركات مرحلة معينة
 * @param {string} stageName اسم المرحلة
 * @returns {Array} مصفوفة الحركات
 */
function getMovementByStage(stageName) {
  if (!stageName) return [];

  const allMovements = getAllMovements();
  return allMovements.filter(movement =>
    movement.stage === stageName ||
    movement.stage.includes(stageName)
  );
}

/**
 * الحصول على حركات حسب الحالة
 * @param {string} status الحالة
 * @returns {Array} مصفوفة الحركات
 */
function getMovementByStatus(status) {
  if (!status) return [];

  const allMovements = getAllMovements();
  return allMovements.filter(movement =>
    movement.status === status ||
    movement.status.includes(status)
  );
}

/**
 * الحصول على حركات تاريخ معين
 * @param {Date} date التاريخ
 * @returns {Array} مصفوفة الحركات
 */
function getMovementByDate(date) {
  if (!date) return [];

  const allMovements = getAllMovements();
  const targetDate = new Date(date);
  targetDate.setHours(0, 0, 0, 0);

  return allMovements.filter(movement => {
    if (!movement.date) return false;
    const movementDate = new Date(movement.date);
    movementDate.setHours(0, 0, 0, 0);
    return movementDate.getTime() === targetDate.getTime();
  });
}

/**
 * الحصول على حركات فترة معينة
 * @param {Date} startDate تاريخ البداية
 * @param {Date} endDate تاريخ النهاية
 * @returns {Array} مصفوفة الحركات
 */
function getMovementByDateRange(startDate, endDate) {
  const allMovements = getAllMovements();
  const start = new Date(startDate);
  start.setHours(0, 0, 0, 0);
  const end = new Date(endDate);
  end.setHours(23, 59, 59, 999);

  return allMovements.filter(movement => {
    if (!movement.date) return false;
    const movementDate = new Date(movement.date);
    return movementDate >= start && movementDate <= end;
  });
}

/**
 * الحصول على الحركات الجارية
 * @returns {Array} مصفوفة الحركات الجارية
 */
function getInProgressMovements() {
  const allMovements = getAllMovements();
  return allMovements.filter(movement =>
    movement.status.includes('جاري') || movement.status.includes('🔄')
  );
}

/**
 * الحصول على الحركات المتأخرة
 * @returns {Array} مصفوفة الحركات المتأخرة
 */
function getDelayedMovements() {
  const allMovements = getAllMovements();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return allMovements.filter(movement => {
    // متأخرة إذا كان تاريخ الاستحقاق في الماضي والحالة ليست "تم"
    if (!movement.dueDate) return false;
    if (movement.status.includes('تم') || movement.status.includes('✅')) return false;
    if (movement.status.includes('ملغي') || movement.status.includes('❌')) return false;

    const dueDate = new Date(movement.dueDate);
    dueDate.setHours(0, 0, 0, 0);
    return dueDate < today;
  });
}

/**
 * الحصول على حركات اليوم
 * @returns {Array} مصفوفة حركات اليوم
 */
function getTodayMovements() {
  return getMovementByDate(new Date());
}

/**
 * الحصول على حركات الأسبوع الحالي
 * @returns {Array} مصفوفة حركات الأسبوع
 */
function getThisWeekMovements() {
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay());

  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);

  return getMovementByDateRange(startOfWeek, endOfWeek);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إضافة وتعديل الحركات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة حركة جديدة
 * @param {Object} movementData بيانات الحركة
 * @returns {boolean} نجاح العملية
 */
function addMovementEntry(movementData) {
  try {
    const sheet = getSheet(SHEETS.MOVEMENT);
    if (!sheet) {
      showError('شيت الحركة غير موجود');
      return false;
    }

    // التحقق من البيانات المطلوبة
    if (!movementData.project) {
      showError('المشروع مطلوب');
      return false;
    }

    if (!movementData.stage) {
      showError('المرحلة مطلوبة');
      return false;
    }

    // التحقق من صحة المشروع
    const project = getProjectByName(movementData.project);
    if (!project) {
      showWarning('المشروع غير موجود في قائمة المشاريع');
    }

    // التحقق من صحة المرحلة للمشروع
    if (project) {
      const projectPhases = getProjectPhases(project.code);
      const stageName = movementData.stage.replace(/^[^\s]+\s/, ''); // إزالة الأيقونة
      const isValidStage = projectPhases.some(p => p.name === stageName);
      if (!isValidStage && projectPhases.length > 0) {
        showWarning('المرحلة المختارة غير مفعلة لهذا المشروع');
      }
    }

    // إنشاء الرقم التسلسلي تلقائياً
    const num = generateMovementNumber();

    // تجهيز صف البيانات
    const rowData = [
      num,
      movementData.date || new Date(),
      movementData.projectCode || '',
      movementData.projectName || movementData.project || '',
      movementData.stage,
      movementData.subtype || '-',
      movementData.element || '',
      movementData.action || '',
      movementData.assignedTo || '',
      movementData.status || '⬜ لم يبدأ',
      movementData.dueDate || '',
      movementData.notes || '',
      getCurrentUserEmail() || 'النظام',
      new Date()
    ];

    sheet.appendRow(rowData);

    // تطبيق لون الحالة
    const lastRow = sheet.getLastRow();
    const statusColor = getStatusColor(movementData.status || 'لم يبدأ');
    sheet.getRange(lastRow, MOVEMENT_COLS.STATUS).setBackground(statusColor);

    return true;
  } catch (error) {
    console.error('Error adding movement:', error);
    showError('حدث خطأ أثناء إضافة الحركة');
    return false;
  }
}

/**
 * تحديث حركة
 * @param {string} num رقم الحركة
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updateMovement(num, updates) {
  try {
    const sheet = getSheet(SHEETS.MOVEMENT);
    if (!sheet) return false;

    const rowIndex = findRowByValue(SHEETS.MOVEMENT, MOVEMENT_COLS.NUMBER, num);
    if (rowIndex === -1) {
      showError('الحركة غير موجودة');
      return false;
    }

    // تحديث الحقول
    const fieldsMap = {
      date: MOVEMENT_COLS.DATE,
      projectCode: MOVEMENT_COLS.PROJECT_CODE,
      projectName: MOVEMENT_COLS.PROJECT_NAME,
      stage: MOVEMENT_COLS.STAGE,
      subtype: MOVEMENT_COLS.SUBTYPE,
      element: MOVEMENT_COLS.ELEMENT,
      action: MOVEMENT_COLS.ACTION,
      assignedTo: MOVEMENT_COLS.ASSIGNED_TO,
      status: MOVEMENT_COLS.STATUS,
      dueDate: MOVEMENT_COLS.DUE_DATE,
      notes: MOVEMENT_COLS.NOTES
    };

    Object.keys(updates).forEach(field => {
      if (fieldsMap[field] && updates[field] !== undefined) {
        sheet.getRange(rowIndex, fieldsMap[field]).setValue(updates[field]);

        // تلوين الحالة إذا تم تحديثها
        if (field === 'status') {
          const statusColor = getStatusColor(updates[field]);
          sheet.getRange(rowIndex, MOVEMENT_COLS.STATUS).setBackground(statusColor);
        }
      }
    });

    return true;
  } catch (error) {
    console.error('Error updating movement:', error);
    return false;
  }
}

/**
 * تحديث حالة حركة
 * @param {string} id معرف الحركة
 * @param {string} newStatus الحالة الجديدة
 * @returns {boolean} نجاح العملية
 */
function updateMovementStatus(num, newStatus) {
  return updateMovement(num, { status: newStatus });
}

/**
 * إنشاء رقم تسلسلي جديد للحركة
 * @returns {number} الرقم التسلسلي
 */
function generateMovementNumber() {
  const sheet = getSheet(SHEETS.MOVEMENT);
  if (!sheet) return 1;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;

  const nums = sheet.getRange(2, MOVEMENT_COLS.NUMBER, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(num => num && !isNaN(num));

  if (nums.length === 0) return 1;

  const maxNum = Math.max(...nums);
  return maxNum + 1;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإحصائيات والتقارير
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على إحصائيات الحركات
 * @param {Object} filters معايير التصفية (اختياري)
 * @returns {Object} كائن الإحصائيات
 */
function getMovementStats(filters) {
  filters = filters || {};

  let movements = getAllMovements();

  // تطبيق الفلاتر
  if (filters.project) {
    movements = movements.filter(m => m.project.includes(filters.project));
  }
  if (filters.person) {
    movements = movements.filter(m => m.assignedTo.includes(filters.person));
  }
  if (filters.startDate && filters.endDate) {
    movements = getMovementByDateRange(filters.startDate, filters.endDate);
  }

  const stats = {
    total: movements.length,
    byStatus: {},
    byStage: {},
    byProject: {},
    byPerson: {},
    delayed: 0,
    completedThisWeek: 0
  };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const weekAgo = new Date(today);
  weekAgo.setDate(weekAgo.getDate() - 7);

  movements.forEach(movement => {
    // حسب الحالة
    const status = movement.status || 'غير محدد';
    stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;

    // حسب المرحلة
    const stage = movement.stage || 'غير محدد';
    stats.byStage[stage] = (stats.byStage[stage] || 0) + 1;

    // حسب المشروع
    const project = movement.project || 'غير محدد';
    stats.byProject[project] = (stats.byProject[project] || 0) + 1;

    // حسب الشخص
    const person = movement.assignedTo || 'غير محدد';
    stats.byPerson[person] = (stats.byPerson[person] || 0) + 1;

    // المتأخرة
    if (movement.dueDate) {
      const dueDate = new Date(movement.dueDate);
      if (dueDate < today &&
        !movement.status.includes('تم') &&
        !movement.status.includes('✅') &&
        !movement.status.includes('ملغي')) {
        stats.delayed++;
      }
    }

    // المكتملة هذا الأسبوع
    if ((movement.status.includes('تم') || movement.status.includes('✅')) &&
      movement.date) {
      const completedDate = new Date(movement.date);
      if (completedDate >= weekAgo && completedDate <= today) {
        stats.completedThisWeek++;
      }
    }
  });

  return stats;
}

/**
 * الحصول على ملخص يومي
 * @param {Date} date التاريخ (اختياري - افتراضي اليوم)
 * @returns {Object} الملخص اليومي
 */
function getDailySummary(date) {
  date = date || new Date();

  const movements = getMovementByDate(date);
  const dateStr = Utilities.formatDate(new Date(date), CONFIG.TIMEZONE, 'yyyy-MM-dd');

  const summary = {
    date: dateStr,
    total: movements.length,
    byStatus: {},
    byProject: {},
    items: movements
  };

  movements.forEach(m => {
    const status = m.status || 'غير محدد';
    summary.byStatus[status] = (summary.byStatus[status] || 0) + 1;

    const project = m.project || 'غير محدد';
    summary.byProject[project] = (summary.byProject[project] || 0) + 1;
  });

  return summary;
}

/**
 * الحصول على تقرير أسبوعي
 * @returns {Object} التقرير الأسبوعي
 */
function getWeeklySummary() {
  const movements = getThisWeekMovements();
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay());

  const summary = {
    weekStart: Utilities.formatDate(startOfWeek, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    weekEnd: Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    total: movements.length,
    byDay: {},
    byStatus: {},
    byProject: {},
    byPerson: {}
  };

  // تجميع حسب اليوم
  for (let i = 0; i <= today.getDay(); i++) {
    const day = new Date(startOfWeek);
    day.setDate(startOfWeek.getDate() + i);
    const dayStr = Utilities.formatDate(day, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    summary.byDay[dayStr] = 0;
  }

  movements.forEach(m => {
    if (m.date) {
      const dayStr = Utilities.formatDate(new Date(m.date), CONFIG.TIMEZONE, 'yyyy-MM-dd');
      summary.byDay[dayStr] = (summary.byDay[dayStr] || 0) + 1;
    }

    const status = m.status || 'غير محدد';
    summary.byStatus[status] = (summary.byStatus[status] || 0) + 1;

    const project = m.project || 'غير محدد';
    summary.byProject[project] = (summary.byProject[project] || 0) + 1;

    const person = m.assignedTo || 'غير محدد';
    summary.byPerson[person] = (summary.byPerson[person] || 0) + 1;
  });

  return summary;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال واجهة المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * فتح نموذج إضافة حركة جديدة
 */
function showAddMovementDialog() {
  // الحصول على البيانات للقوائم المنسدلة
  const activeProjects = getActiveProjects();
  const projectOptions = activeProjects
    .map(p => `<option value="${p.name}">${p.name}</option>`)
    .join('');

  const stageOptions = Object.values(STAGES)
    .map(s => `<option value="${s.icon} ${s.name}">${s.icon} ${s.name}</option>`)
    .join('');

  const teamMembers = getTeamMembers();
  const teamOptions = teamMembers
    .map(t => `<option value="${t.name}">${t.name}</option>`)
    .join('');

  const statusOptions = Object.values(STATUS)
    .map(s => `<option value="${s.icon} ${s.name}">${s.icon} ${s.name}</option>`)
    .join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; direction: rtl; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
      .row { display: flex; gap: 15px; }
      .row .form-group { flex: 1; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px; }
      .btn-primary { background: #1565c0; color: white; }
      .btn-secondary { background: #757575; color: white; }
    </style>

    <h3>إضافة حركة جديدة</h3>

    <div class="row">
      <div class="form-group">
        <label>المشروع *</label>
        <select id="project" required onchange="updateStages()">
          <option value="">اختر المشروع</option>
          ${projectOptions}
        </select>
      </div>
      <div class="form-group">
        <label>التاريخ</label>
        <input type="date" id="date" value="${new Date().toISOString().split('T')[0]}">
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>المرحلة *</label>
        <select id="stage" required onchange="updateSubtypes()">
          <option value="">اختر المرحلة</option>
          ${stageOptions}
        </select>
      </div>
      <div class="form-group">
        <label>النوع الفرعي</label>
        <select id="subtype">
          <option value="-">-</option>
        </select>
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>العنصر</label>
        <input type="text" id="element" placeholder="مثال: ضيف أحمد، مشهد 5...">
      </div>
      <div class="form-group">
        <label>الإجراء</label>
        <input type="text" id="action" placeholder="مثال: تصوير، مراجعة...">
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>المسؤول</label>
        <select id="assignedTo">
          <option value="">اختر المسؤول</option>
          ${teamOptions}
        </select>
      </div>
      <div class="form-group">
        <label>الحالة</label>
        <select id="status">${statusOptions}</select>
      </div>
    </div>

    <div class="form-group">
      <label>تاريخ الاستحقاق</label>
      <input type="date" id="dueDate">
    </div>

    <div class="form-group">
      <label>ملاحظات</label>
      <textarea id="notes" rows="3"></textarea>
    </div>

    <button class="btn btn-primary" onclick="submitForm()">حفظ</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>

    <script>
      // تحديث المراحل عند اختيار المشروع
      function updateStages() {
        const project = document.getElementById('project').value;
        if (!project) return;

        google.script.run
          .withSuccessHandler(function(phases) {
            const stageSelect = document.getElementById('stage');
            stageSelect.innerHTML = '<option value="">اختر المرحلة</option>';

            if (phases && phases.length > 0) {
              phases.forEach(function(phase) {
                const option = document.createElement('option');
                option.value = phase.icon + ' ' + phase.name;
                option.textContent = phase.icon + ' ' + phase.name;
                stageSelect.appendChild(option);
              });
            }
          })
          .getProjectPhases(project);
      }

      // تحديث الأنواع الفرعية عند اختيار المرحلة
      function updateSubtypes() {
        const stage = document.getElementById('stage').value;
        const subtypeSelect = document.getElementById('subtype');
        subtypeSelect.innerHTML = '<option value="-">-</option>';

        // البحث عن المرحلة وأنواعها الفرعية
        const subtypesMap = {
          '🎙️ التعليق الصوتي': ['راوي', 'اقتباس', 'دوبلاج']
        };

        if (subtypesMap[stage]) {
          subtypesMap[stage].forEach(function(subtype) {
            const option = document.createElement('option');
            option.value = subtype;
            option.textContent = subtype;
            subtypeSelect.appendChild(option);
          });
        }
      }

      function submitForm() {
        const data = {
          date: document.getElementById('date').value,
          project: document.getElementById('project').value,
          stage: document.getElementById('stage').value,
          subtype: document.getElementById('subtype').value,
          element: document.getElementById('element').value,
          action: document.getElementById('action').value,
          assignedTo: document.getElementById('assignedTo').value,
          status: document.getElementById('status').value,
          dueDate: document.getElementById('dueDate').value,
          notes: document.getElementById('notes').value
        };

        if (!data.project || !data.stage) {
          alert('المشروع والمرحلة مطلوبان');
          return;
        }

        google.script.run
          .withSuccessHandler(function() {
            alert('تم إضافة الحركة بنجاح');
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            alert('حدث خطأ: ' + err.message);
          })
          .addMovementEntry(data);
      }
    </script>
  `).setWidth(550).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة حركة جديدة');
}

/**
 * عرض الحركات المتأخرة
 */
function showDelayedMovements() {
  const delayed = getDelayedMovements();

  if (delayed.length === 0) {
    showInfo('لا توجد حركات متأخرة 🎉', 'الحركات المتأخرة');
    return;
  }

  const list = delayed.map(m => {
    const daysLate = daysRemaining(m.dueDate) * -1;
    return `• ${m.project} - ${m.stage}\n  ${m.element || m.action}\n  المسؤول: ${m.assignedTo || 'غير محدد'} | متأخر ${daysLate} يوم`;
  }).join('\n\n');

  const message = `
الحركات المتأخرة (${delayed.length})
━━━━━━━━━━━━━━━━━━━━━━━━━

${list}
  `.trim();

  showWarning(message, 'الحركات المتأخرة');
}

/**
 * عرض ملخص اليوم
 */
function showTodaySummary() {
  const summary = getDailySummary();

  if (summary.total === 0) {
    showInfo('لا توجد حركات مسجلة اليوم', 'ملخص اليوم');
    return;
  }

  let statusList = Object.entries(summary.byStatus)
    .map(([status, count]) => `  • ${status}: ${count}`)
    .join('\n');

  let projectList = Object.entries(summary.byProject)
    .map(([project, count]) => `  • ${project}: ${count}`)
    .join('\n');

  const message = `
ملخص اليوم - ${summary.date}
━━━━━━━━━━━━━━━━━━━━━━━━━

📊 إجمالي الحركات: ${summary.total}

حسب الحالة:
${statusList}

حسب المشروع:
${projectList}
  `.trim();

  showInfo(message, 'ملخص اليوم');
}
