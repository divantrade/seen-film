/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * شيت الحركة - إدارة المهام والمتابعة
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * إضافة حركة جديدة عبر نموذج
 */
function showAddMovementForm() {
  const projects = getActiveProjectNames();
  const team = getActiveTeamNames();

  if (projects.length === 0) {
    showError('لا توجد مشاريع نشطة. أضف مشروعاً أولاً.');
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; direction: rtl; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 8px; box-sizing: border-box; }
      button { background: #1565C0; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-left: 10px; }
      button:hover { background: #0D47A1; }
      .cancel { background: #757575; }
      .row { display: flex; gap: 10px; }
      .row .form-group { flex: 1; }
    </style>
    <form id="movementForm">
      <div class="form-group">
        <label>الفيلم *</label>
        <select id="project" required>
          <option value="">اختر الفيلم</option>
          ${projects.map(p => `<option value="${p}">${p}</option>`).join('')}
        </select>
      </div>
      <div class="row">
        <div class="form-group">
          <label>المرحلة *</label>
          <select id="stage" required onchange="updateSubtypes()">
            <option value="">اختر المرحلة</option>
            ${STAGE_NAMES.map(s => `<option value="${s}">${s}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label>النوع الفرعي</label>
          <select id="subtype">
            <option value="">اختر النوع</option>
          </select>
        </div>
      </div>
      <div class="form-group">
        <label>العنصر *</label>
        <input type="text" id="element" required placeholder="مثال: مقابلة أحمد، بحث المصادر...">
      </div>
      <div class="form-group">
        <label>التفاصيل</label>
        <textarea id="details" rows="2"></textarea>
      </div>
      <div class="row">
        <div class="form-group">
          <label>المسؤول</label>
          <select id="assignedTo">
            <option value="">اختر المسؤول</option>
            ${team.map(t => `<option value="${t}">${t}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label>تاريخ الاستحقاق</label>
          <input type="date" id="dueDate">
        </div>
      </div>
      <div class="form-group">
        <label>ملاحظات</label>
        <input type="text" id="notes">
      </div>
      <button type="submit">إضافة</button>
      <button type="button" class="cancel" onclick="google.script.host.close()">إلغاء</button>
    </form>
    <script>
      const stages = ${JSON.stringify(STAGES)};

      function updateSubtypes() {
        const stage = document.getElementById('stage').value;
        const subtypeSelect = document.getElementById('subtype');
        subtypeSelect.innerHTML = '<option value="">اختر النوع</option>';

        for (const key in stages) {
          if (stages[key].name === stage && stages[key].subtypes) {
            stages[key].subtypes.forEach(s => {
              subtypeSelect.innerHTML += '<option value="' + s + '">' + s + '</option>';
            });
          }
        }
      }

      document.getElementById('movementForm').onsubmit = function(e) {
        e.preventDefault();
        const data = {
          project: document.getElementById('project').value,
          stage: document.getElementById('stage').value,
          subtype: document.getElementById('subtype').value,
          element: document.getElementById('element').value,
          details: document.getElementById('details').value,
          assignedTo: document.getElementById('assignedTo').value,
          dueDate: document.getElementById('dueDate').value,
          notes: document.getElementById('notes').value
        };
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .addMovement(data);
      };
    </script>
  `)
    .setWidth(500)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة حركة جديدة');
}

/**
 * إضافة حركة جديدة
 */
function addMovement(data) {
  const sheet = getSheet(SHEETS.MOVEMENT);
  const lastRow = getLastRowInColumn(sheet, MOVEMENT_COLS.PROJECT);
  const newRow = Math.max(lastRow + 1, 2);

  // الرقم التسلسلي
  const number = lastRow <= 1 ? 1 : lastRow;

  // إضافة البيانات
  sheet.getRange(newRow, MOVEMENT_COLS.NUMBER).setValue(number);
  sheet.getRange(newRow, MOVEMENT_COLS.DATE).setValue(getCurrentDate());
  sheet.getRange(newRow, MOVEMENT_COLS.PROJECT).setValue(data.project);
  sheet.getRange(newRow, MOVEMENT_COLS.STAGE).setValue(data.stage);
  sheet.getRange(newRow, MOVEMENT_COLS.SUBTYPE).setValue(data.subtype || '');
  sheet.getRange(newRow, MOVEMENT_COLS.ELEMENT).setValue(cleanText(data.element));
  sheet.getRange(newRow, MOVEMENT_COLS.DETAILS).setValue(data.details || '');
  sheet.getRange(newRow, MOVEMENT_COLS.ASSIGNED_TO).setValue(data.assignedTo || '');
  sheet.getRange(newRow, MOVEMENT_COLS.STATUS).setValue('⬜ لم يبدأ');
  sheet.getRange(newRow, MOVEMENT_COLS.DUE_DATE).setValue(data.dueDate || '');
  sheet.getRange(newRow, MOVEMENT_COLS.NOTES).setValue(data.notes || '');

  // إنشاء فولدر تلقائي إذا كانت المرحلة "التصوير"
  const isShootingStage = data.stage === 'التصوير' || data.stage === 'تصوير' || (data.stage && data.stage.toLowerCase() === 'shooting');
  if (isShootingStage && data.element) {
    createShootingFolder(data.project, data.subtype, newRow, data.element);
  }

  showSuccess('تم إضافة الحركة بنجاح');
}

/**
 * تحديث حالة الحركة
 */
function updateMovementStatus(newStatus) {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== SHEETS.MOVEMENT) {
    showError('يجب أن تكون في شيت الحركة');
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    showError('اختر حركة من القائمة');
    return;
  }

  sheet.getRange(row, MOVEMENT_COLS.STATUS).setValue(newStatus);
  colorRowByStatus(sheet, row, newStatus);

  showSuccess('تم تحديث الحالة');
}

/**
 * وضع علامة "تم" على الحركة المحددة
 */
function markAsCompleted() {
  updateMovementStatus('✅ تم');
}

/**
 * وضع علامة "جاري" على الحركة المحددة
 */
function markAsInProgress() {
  updateMovementStatus('🔄 جاري');
}

/**
 * وضع علامة "متأخر" على الحركة المحددة
 */
function markAsDelayed() {
  updateMovementStatus('🔴 متأخر');
}

/**
 * Trigger عند تعديل شيت الحركة
 */
function onMovementEdit(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== SHEETS.MOVEMENT) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // ملء التاريخ والرقم تلقائياً عند إدخال بيانات جديدة
  if (row > 1 && col > MOVEMENT_COLS.DATE) {
    const numberCell = sheet.getRange(row, MOVEMENT_COLS.NUMBER);
    const dateCell = sheet.getRange(row, MOVEMENT_COLS.DATE);

    // ملء الرقم إذا كان فارغاً
    if (!numberCell.getValue()) {
      const lastRow = getLastRowInColumn(sheet, MOVEMENT_COLS.PROJECT);
      const newNumber = Math.max(row - 1, lastRow > 1 ? lastRow : 1);
      numberCell.setValue(newNumber);
    }

    // ملء التاريخ إذا كان فارغاً
    if (!dateCell.getValue()) {
      dateCell.setValue(getCurrentDate());
    }
  }

  // تلوين الصف عند تغيير الحالة
  if (col === MOVEMENT_COLS.STATUS && row > 1) {
    colorRowByStatus(sheet, row, e.value);
  }

  // تحديث الأنواع الفرعية عند تغيير المرحلة
  if (col === MOVEMENT_COLS.STAGE && row > 1) {
    const stage = e.value;
    // قراءة الأنواع الفرعية من شيت الإعدادات
    const subtypes = getSubtypesFromSettings(stage);

    if (subtypes && subtypes.length > 0) {
      const subtypeCell = sheet.getRange(row, MOVEMENT_COLS.SUBTYPE);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(subtypes, true)
        .setAllowInvalid(true)
        .build();
      subtypeCell.setDataValidation(rule);
    } else {
    // إزالة التحقق للتصوير (الإنتاج) إذا كان يتطلب إدخال يدوي
    const stageKey = Object.keys(STAGES).find(key => STAGES[key].name === stage);
    if(stageKey === 'PRODUCTION' || stageKey === 'SHOOTING') {
       sheet.getRange(row, MOVEMENT_COLS.SUBTYPE).clearDataValidations();
    }
  }
}

  // إنشاء فولدر تلقائي للتصوير (الإنتاج) عند إدخال العنصر
  if (col === MOVEMENT_COLS.ELEMENT && row > 1) {
    const stage = sheet.getRange(row, MOVEMENT_COLS.STAGE).getValue();
    // التحقق من مرحلة الإنتاج أو التصوير
    const isShootingStage = stage === 'الإنتاج' || stage === 'التصوير';

    if (isShootingStage && e.value) {
      const project = sheet.getRange(row, MOVEMENT_COLS.PROJECT).getValue();
      const subtype = sheet.getRange(row, MOVEMENT_COLS.SUBTYPE).getValue();
      const existingLink = sheet.getRange(row, MOVEMENT_COLS.LINK).getValue();

      if (!existingLink && project) {
        createShootingFolder(project, subtype, row, e.value);
      }
    }
  }
}

/**
 * الحصول على حركات مشروع معين
 */
function getProjectMovements(projectName) {
  const sheet = getSheet(SHEETS.MOVEMENT);
  const lastRow = getLastRowInColumn(sheet, MOVEMENT_COLS.PROJECT);

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, MOVEMENT_COLS.NOTES).getValues();
  const movements = [];

  for (const row of data) {
    if (row[MOVEMENT_COLS.PROJECT - 1] === projectName) {
      movements.push({
        number: row[MOVEMENT_COLS.NUMBER - 1],
        date: row[MOVEMENT_COLS.DATE - 1],
        stage: row[MOVEMENT_COLS.STAGE - 1],
        subtype: row[MOVEMENT_COLS.SUBTYPE - 1],
        element: row[MOVEMENT_COLS.ELEMENT - 1],
        assignedTo: row[MOVEMENT_COLS.ASSIGNED_TO - 1],
        status: row[MOVEMENT_COLS.STATUS - 1],
        dueDate: row[MOVEMENT_COLS.DUE_DATE - 1]
      });
    }
  }

  return movements;
}

/**
 * الحصول على إحصائيات الحركات
 */
function getMovementStats(projectName) {
  const movements = projectName ? getProjectMovements(projectName) : getAllMovements();

  const stats = {
    total: movements.length,
    completed: 0,
    inProgress: 0,
    waiting: 0,
    delayed: 0,
    notStarted: 0,
    byStage: {}
  };

  for (const m of movements) {
    // إحصائيات الحالة
    if (m.status.includes('تم')) stats.completed++;
    else if (m.status.includes('جاري')) stats.inProgress++;
    else if (m.status.includes('انتظار')) stats.waiting++;
    else if (m.status.includes('متأخر')) stats.delayed++;
    else stats.notStarted++;

    // إحصائيات المراحل
    if (!stats.byStage[m.stage]) {
      stats.byStage[m.stage] = { total: 0, completed: 0 };
    }
    stats.byStage[m.stage].total++;
    if (m.status.includes('تم')) {
      stats.byStage[m.stage].completed++;
    }
  }

  return stats;
}

/**
 * الحصول على جميع الحركات
 */
function getAllMovements() {
  const sheet = getSheet(SHEETS.MOVEMENT);
  const lastRow = getLastRowInColumn(sheet, MOVEMENT_COLS.PROJECT);

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, MOVEMENT_COLS.NOTES).getValues();
  const movements = [];

  for (const row of data) {
    movements.push({
      number: row[MOVEMENT_COLS.NUMBER - 1],
      date: row[MOVEMENT_COLS.DATE - 1],
      project: row[MOVEMENT_COLS.PROJECT - 1],
      stage: row[MOVEMENT_COLS.STAGE - 1],
      subtype: row[MOVEMENT_COLS.SUBTYPE - 1],
      element: row[MOVEMENT_COLS.ELEMENT - 1],
      assignedTo: row[MOVEMENT_COLS.ASSIGNED_TO - 1],
      status: row[MOVEMENT_COLS.STATUS - 1],
      dueDate: row[MOVEMENT_COLS.DUE_DATE - 1]
    });
  }

  return movements;
}

/**
 * الحصول على المهام المتأخرة
 */
function getDelayedTasks() {
  const movements = getAllMovements();
  const today = new Date();

  return movements.filter(m => {
    if (m.status.includes('تم') || m.status.includes('ملغي')) return false;
    if (!m.dueDate) return false;

    const dueDate = new Date(m.dueDate);
    return dueDate < today;
  });
}

/**
 * تحديث حالة المهام المتأخرة تلقائياً
 */
function updateDelayedTasks() {
  const sheet = getSheet(SHEETS.MOVEMENT);
  const lastRow = getLastRowInColumn(sheet, MOVEMENT_COLS.PROJECT);

  if (lastRow <= 1) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const data = sheet.getRange(2, 1, lastRow - 1, MOVEMENT_COLS.NOTES).getValues();
  let updatedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const status = data[i][MOVEMENT_COLS.STATUS - 1];
    const dueDate = data[i][MOVEMENT_COLS.DUE_DATE - 1];

    // تخطي المكتملة والملغية
    if (status.includes('تم') || status.includes('ملغي')) continue;
    // تخطي المتأخرة بالفعل
    if (status.includes('متأخر')) continue;
    // تخطي بدون تاريخ استحقاق
    if (!dueDate) continue;

    const dueDateObj = new Date(dueDate);
    dueDateObj.setHours(0, 0, 0, 0);

    if (dueDateObj < today) {
      const row = i + 2;
      sheet.getRange(row, MOVEMENT_COLS.STATUS).setValue('🔴 متأخر');
      colorRowByStatus(sheet, row, '🔴 متأخر');
      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    showInfo('تم تحديث ' + updatedCount + ' مهمة متأخرة');
  }
}
