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
 * @param {string} projectCodeOrName كود المشروع أو اسمه
 * @returns {Array} مصفوفة الحركات
 */
function getMovementByProject(projectCodeOrName) {
  if (!projectCodeOrName) return [];

  const allMovements = getAllMovements();
  return allMovements.filter(movement =>
    movement.projectCode === projectCodeOrName ||
    movement.projectName === projectCodeOrName ||
    (movement.projectCode && movement.projectCode.includes(projectCodeOrName)) ||
    (movement.projectName && movement.projectName.includes(projectCodeOrName))
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
    movements = movements.filter(m =>
      (m.projectCode && m.projectCode.includes(filters.project)) ||
      (m.projectName && m.projectName.includes(filters.project))
    );
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

    // حسب المشروع (استخدام اسم المشروع للعرض)
    const project = movement.projectName || movement.projectCode || 'غير محدد';
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

    const project = m.projectName || m.projectCode || 'غير محدد';
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

    const project = m.projectName || m.projectCode || 'غير محدد';
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
    const projectDisplay = m.projectName || m.projectCode || 'غير محدد';
    return `• ${projectDisplay} - ${m.stage}\n  ${m.element || m.action}\n  المسؤول: ${m.assignedTo || 'غير محدد'} | متأخر ${daysLate} يوم`;
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

// ═══════════════════════════════════════════════════════════════════════════════
// النموذج الذكي لإضافة الحركات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الأنواع الفرعية الذكية حسب المرحلة
 */
const SMART_SUBTYPES = {
  'تصوير المقابلات': ['مقابلة شخصية', 'مقابلة جماعية', 'مقابلة هاتفية'],
  'التصوير الميداني': ['تصوير خارجي', 'تصوير داخلي'],
  'التعليق الصوتي': ['تعليق رئيسي', 'دوبلاج', 'مؤثرات صوتية'],
  'الرسوم المتحركة': ['2D', '3D', 'موشن جرافيك']
};

/**
 * الإجراءات المتاحة
 */
const MOVEMENT_ACTIONS = ['تصوير', 'تسجيل', 'تسليم', 'مراجعة', 'تعديل', 'إلغاء'];

/**
 * الحصول على الأنواع الفرعية حسب المرحلة
 * @param {string} stageName اسم المرحلة
 * @returns {Array} قائمة الأنواع الفرعية
 */
function getSmartSubtypes(stageName) {
  // إزالة الأيقونة من اسم المرحلة
  const cleanName = stageName.replace(/^[^\s]+\s/, '').trim();
  return SMART_SUBTYPES[cleanName] || ['عام'];
}

/**
 * الحصول على ضيوف مشروع للنموذج الذكي
 * @param {string} projectName اسم المشروع
 * @returns {Array} قائمة أسماء الضيوف
 */
function getGuestsForSmartForm(projectName) {
  if (!projectName) return [];
  const guests = getGuestsByProject(projectName);
  return guests.map(g => g.name).filter(n => n);
}

/**
 * التحقق من كون المرحلة مرحلة تصوير
 * @param {string} stageName اسم المرحلة
 * @returns {boolean}
 */
function isShootingStage(stageName) {
  const cleanName = stageName.replace(/^[^\s]+\s/, '').trim();
  return cleanName.includes('تصوير') || cleanName.includes('الميداني');
}

/**
 * التحقق من كون المرحلة مرحلة تصوير مقابلات
 * @param {string} stageName اسم المرحلة
 * @returns {boolean}
 */
function isInterviewStage(stageName) {
  const cleanName = stageName.replace(/^[^\s]+\s/, '').trim();
  return cleanName.includes('المقابلات');
}

/**
 * عرض النموذج الذكي لإضافة حركة جديدة
 */
function showSmartMovementForm() {
  // تجهيز بيانات المشاريع
  const activeProjects = getActiveProjects();
  const projectsData = activeProjects.map(p => ({
    code: p.code || '',
    name: p.name || ''
  }));

  // تجهيز بيانات المراحل
  const stagesData = Object.values(STAGES).map(s => ({
    id: s.id,
    name: s.name,
    icon: s.icon,
    displayName: `${s.icon} ${s.name}`
  }));

  // تجهيز بيانات الحالات
  const statusData = Object.values(STATUS).map(s => ({
    name: s.name,
    icon: s.icon,
    displayName: `${s.icon} ${s.name}`
  }));

  // تجهيز بيانات الفريق
  const teamMembers = getTeamMembers();
  const teamData = teamMembers.map(t => t.name).filter(n => n);

  // تجهيز بيانات المصورين
  const photographers = getPhotographers();
  const photographersData = photographers.map(p => p.name).filter(n => n);

  // حساب تاريخ الاستحقاق الافتراضي (اليوم + 7 أيام)
  const defaultDueDate = new Date();
  defaultDueDate.setDate(defaultDueDate.getDate() + 7);
  const dueDateStr = Utilities.formatDate(defaultDueDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
      <meta charset="UTF-8">
      <style>
        * {
          box-sizing: border-box;
          font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
        }
        body {
          direction: rtl;
          padding: 20px;
          background: #f5f5f5;
          margin: 0;
        }
        .form-container {
          background: white;
          border-radius: 12px;
          padding: 25px;
          box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h2 {
          color: #1565c0;
          margin-top: 0;
          margin-bottom: 20px;
          padding-bottom: 10px;
          border-bottom: 2px solid #e3f2fd;
        }
        .form-row {
          display: flex;
          gap: 15px;
          margin-bottom: 15px;
        }
        .form-group {
          flex: 1;
          margin-bottom: 15px;
        }
        .form-group.full-width {
          flex: 100%;
        }
        label {
          display: block;
          margin-bottom: 6px;
          font-weight: 600;
          color: #333;
          font-size: 14px;
        }
        label .required {
          color: #d32f2f;
        }
        input, select, textarea {
          width: 100%;
          padding: 10px 12px;
          border: 1px solid #ddd;
          border-radius: 8px;
          font-size: 14px;
          transition: border-color 0.2s, box-shadow 0.2s;
        }
        input:focus, select:focus, textarea:focus {
          outline: none;
          border-color: #1565c0;
          box-shadow: 0 0 0 3px rgba(21, 101, 192, 0.1);
        }
        select {
          background: white;
          cursor: pointer;
        }
        textarea {
          resize: vertical;
          min-height: 80px;
        }
        .info-box {
          background: #e3f2fd;
          border: 1px solid #90caf9;
          border-radius: 8px;
          padding: 12px;
          margin-bottom: 20px;
          display: none;
        }
        .info-box.show {
          display: block;
        }
        .info-box .label {
          font-size: 12px;
          color: #1565c0;
          margin-bottom: 4px;
        }
        .info-box .value {
          font-weight: 600;
          color: #0d47a1;
        }
        .btn-container {
          display: flex;
          gap: 10px;
          justify-content: flex-start;
          margin-top: 20px;
          padding-top: 20px;
          border-top: 1px solid #eee;
        }
        .btn {
          padding: 12px 28px;
          border: none;
          border-radius: 8px;
          cursor: pointer;
          font-size: 14px;
          font-weight: 600;
          transition: all 0.2s;
        }
        .btn-primary {
          background: #1565c0;
          color: white;
        }
        .btn-primary:hover {
          background: #0d47a1;
        }
        .btn-primary:disabled {
          background: #90caf9;
          cursor: not-allowed;
        }
        .btn-secondary {
          background: #757575;
          color: white;
        }
        .btn-secondary:hover {
          background: #616161;
        }
        .loading {
          display: none;
          color: #1565c0;
          font-size: 14px;
          margin-right: 10px;
        }
        .loading.show {
          display: inline;
        }
        #elementContainer {
          transition: all 0.3s;
        }
      </style>
    </head>
    <body>
      <div class="form-container">
        <h2>إضافة حركة ذكية</h2>

        <!-- معلومات المشروع المختار -->
        <div class="info-box" id="projectInfo">
          <div class="form-row" style="margin-bottom: 0;">
            <div class="form-group" style="margin-bottom: 0;">
              <div class="label">كود المشروع</div>
              <div class="value" id="projectCodeDisplay">-</div>
            </div>
            <div class="form-group" style="margin-bottom: 0;">
              <div class="label">اسم المشروع</div>
              <div class="value" id="projectNameDisplay">-</div>
            </div>
          </div>
        </div>

        <!-- المشروع والمرحلة -->
        <div class="form-row">
          <div class="form-group">
            <label>المشروع <span class="required">*</span></label>
            <select id="project" onchange="onProjectChange()">
              <option value="">-- اختر المشروع --</option>
              ${projectsData.map(p => '<option value="' + p.name + '" data-code="' + p.code + '">' + p.name + '</option>').join('')}
            </select>
          </div>
          <div class="form-group">
            <label>المرحلة <span class="required">*</span></label>
            <select id="stage" onchange="onStageChange()">
              <option value="">-- اختر المرحلة --</option>
              ${stagesData.map(s => '<option value="' + s.displayName + '" data-name="' + s.name + '">' + s.displayName + '</option>').join('')}
            </select>
          </div>
        </div>

        <!-- النوع الفرعي والإجراء -->
        <div class="form-row">
          <div class="form-group">
            <label>النوع الفرعي</label>
            <select id="subtype">
              <option value="عام">عام</option>
            </select>
          </div>
          <div class="form-group">
            <label>الإجراء</label>
            <select id="action">
              <option value="">-- اختر الإجراء --</option>
              ${MOVEMENT_ACTIONS.map(a => '<option value="' + a + '">' + a + '</option>').join('')}
            </select>
          </div>
        </div>

        <!-- العنصر -->
        <div class="form-group" id="elementContainer">
          <label>العنصر</label>
          <input type="text" id="element" placeholder="مثال: مشهد 1، ضيف أحمد...">
        </div>

        <!-- المسؤول والحالة -->
        <div class="form-row">
          <div class="form-group">
            <label>المسؤول</label>
            <select id="assignedTo">
              <option value="">-- اختر المسؤول --</option>
              ${teamData.map(t => '<option value="' + t + '">' + t + '</option>').join('')}
            </select>
          </div>
          <div class="form-group">
            <label>الحالة</label>
            <select id="status">
              ${statusData.map(s => '<option value="' + s.displayName + '">' + s.displayName + '</option>').join('')}
            </select>
          </div>
        </div>

        <!-- تاريخ الاستحقاق -->
        <div class="form-group">
          <label>تاريخ الاستحقاق</label>
          <input type="date" id="dueDate" value="${dueDateStr}">
        </div>

        <!-- ملاحظات -->
        <div class="form-group">
          <label>ملاحظات</label>
          <textarea id="notes" placeholder="أضف أي ملاحظات إضافية..."></textarea>
        </div>

        <!-- أزرار -->
        <div class="btn-container">
          <button class="btn btn-primary" id="submitBtn" onclick="submitForm()">
            حفظ الحركة
          </button>
          <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>
          <span class="loading" id="loading">جاري الحفظ...</span>
        </div>
      </div>

      <script>
        // بيانات مخزنة محلياً
        const projectsData = ${JSON.stringify(projectsData)};
        const stagesData = ${JSON.stringify(stagesData)};
        const teamData = ${JSON.stringify(teamData)};
        const photographersData = ${JSON.stringify(photographersData)};
        const smartSubtypes = ${JSON.stringify(SMART_SUBTYPES)};

        let currentProjectCode = '';
        let currentProjectName = '';
        let guestsCache = {};

        // عند تغيير المشروع
        function onProjectChange() {
          const select = document.getElementById('project');
          const selectedOption = select.options[select.selectedIndex];
          const infoBox = document.getElementById('projectInfo');

          if (select.value) {
            currentProjectCode = selectedOption.getAttribute('data-code') || '';
            currentProjectName = select.value;

            document.getElementById('projectCodeDisplay').textContent = currentProjectCode || '-';
            document.getElementById('projectNameDisplay').textContent = currentProjectName;
            infoBox.classList.add('show');

            // تحميل ضيوف المشروع للكاش
            loadProjectGuests(currentProjectName);
          } else {
            currentProjectCode = '';
            currentProjectName = '';
            infoBox.classList.remove('show');
          }

          // تحديث حقل العنصر إذا كانت المرحلة تصوير مقابلات
          updateElementField();
        }

        // عند تغيير المرحلة
        function onStageChange() {
          const stage = document.getElementById('stage').value;
          const stageName = stage.replace(/^[^\\s]+\\s/, '').trim();

          // تحديث الأنواع الفرعية
          updateSubtypes(stageName);

          // تحديث حقل العنصر
          updateElementField();

          // تحديث قائمة المسؤولين
          updateAssignedTo(stageName);
        }

        // تحديث الأنواع الفرعية
        function updateSubtypes(stageName) {
          const subtypeSelect = document.getElementById('subtype');
          subtypeSelect.innerHTML = '';

          const subtypes = smartSubtypes[stageName] || ['عام'];
          subtypes.forEach(function(st) {
            const option = document.createElement('option');
            option.value = st;
            option.textContent = st;
            subtypeSelect.appendChild(option);
          });
        }

        // تحديث حقل العنصر (dropdown للضيوف أو نص حر)
        function updateElementField() {
          const stage = document.getElementById('stage').value;
          const stageName = stage.replace(/^[^\\s]+\\s/, '').trim();
          const container = document.getElementById('elementContainer');

          if (stageName.includes('المقابلات') && currentProjectName) {
            // مرحلة تصوير مقابلات - عرض dropdown للضيوف
            const guests = guestsCache[currentProjectName] || [];

            let html = '<label>الضيف</label><select id="element">';
            html += '<option value="">-- اختر الضيف --</option>';
            guests.forEach(function(g) {
              html += '<option value="' + g + '">' + g + '</option>';
            });
            html += '</select>';
            container.innerHTML = html;
          } else {
            // حقل نص حر
            container.innerHTML = '<label>العنصر</label><input type="text" id="element" placeholder="مثال: مشهد 1، ملف صوتي...">';
          }
        }

        // تحديث قائمة المسؤولين
        function updateAssignedTo(stageName) {
          const select = document.getElementById('assignedTo');
          select.innerHTML = '<option value="">-- اختر المسؤول --</option>';

          // إذا كانت مرحلة تصوير، استخدم المصورين
          const isShootingStage = stageName.includes('تصوير') || stageName.includes('الميداني');
          const people = isShootingStage ? photographersData : teamData;

          people.forEach(function(p) {
            const option = document.createElement('option');
            option.value = p;
            option.textContent = p;
            select.appendChild(option);
          });
        }

        // تحميل ضيوف المشروع
        function loadProjectGuests(projectName) {
          if (guestsCache[projectName]) return;

          google.script.run
            .withSuccessHandler(function(guests) {
              guestsCache[projectName] = guests || [];
              // تحديث حقل العنصر إذا كانت المرحلة مقابلات
              const stage = document.getElementById('stage').value;
              if (stage && stage.includes('المقابلات')) {
                updateElementField();
              }
            })
            .getGuestsForSmartForm(projectName);
        }

        // إرسال النموذج
        function submitForm() {
          const project = document.getElementById('project').value;
          const stage = document.getElementById('stage').value;

          if (!project) {
            alert('الرجاء اختيار المشروع');
            return;
          }
          if (!stage) {
            alert('الرجاء اختيار المرحلة');
            return;
          }

          // تعطيل الزر وإظهار التحميل
          document.getElementById('submitBtn').disabled = true;
          document.getElementById('loading').classList.add('show');

          const data = {
            projectCode: currentProjectCode,
            projectName: currentProjectName,
            project: currentProjectName,
            stage: stage,
            subtype: document.getElementById('subtype').value,
            element: document.getElementById('element').value,
            action: document.getElementById('action').value,
            assignedTo: document.getElementById('assignedTo').value,
            status: document.getElementById('status').value,
            dueDate: document.getElementById('dueDate').value,
            notes: document.getElementById('notes').value
          };

          google.script.run
            .withSuccessHandler(function(success) {
              if (success) {
                alert('تم إضافة الحركة بنجاح!');
                google.script.host.close();
              } else {
                document.getElementById('submitBtn').disabled = false;
                document.getElementById('loading').classList.remove('show');
              }
            })
            .withFailureHandler(function(err) {
              alert('حدث خطأ: ' + err.message);
              document.getElementById('submitBtn').disabled = false;
              document.getElementById('loading').classList.remove('show');
            })
            .addMovementEntry(data);
        }
      </script>
    </body>
    </html>
  `).setWidth(600).setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة حركة ذكية');
}
