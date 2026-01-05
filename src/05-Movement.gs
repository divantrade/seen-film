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
  const cities = getCitiesFromSettings();

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
            ${getStagesFromSettings().map(s => `<option value="${s}">${s}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label>النوع الفرعي</label>
          <select id="subtype">
            <option value="">اختر النوع</option>
          </select>
        </div>
      </div>
      <div class="row">
        <div class="form-group">
          <label>المدينة</label>
          <select id="city">
            <option value="">اختر المدينة</option>
            ${cities.map(c => `<option value="${c}">${c}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label>العنصر *</label>
          <input type="text" id="element" required placeholder="مثال: مقابلة أحمد، بحث المصادر...">
        </div>
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
          city: document.getElementById('city').value,
          element: document.getElementById('element').value,
          details: document.getElementById('details').value,
          assignedTo: document.getElementById('assignedTo').value,
          dueDate: document.getElementById('dueDate').value,
          notes: document.getElementById('notes').value
        };
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler((err) => alert('خطأ: ' + err.message))
          .addMovement(data);
      };
    </script>
  `)
    .setWidth(500)
    .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة حركة جديدة');
}

/**
 * إضافة حركة جديدة
 */
function addMovement(data) {
  try {
    const sheet = getSheet(SHEETS.MOVEMENT);
    if (!sheet) {
      throw new Error('شيت الحركة غير موجود');
    }

    const lastRow = getLastRowInColumn(sheet, MOVEMENT_COLS.PROJECT);
    const newRow = Math.max(lastRow + 1, 2);

    // الرقم التسلسلي
    const number = lastRow <= 1 ? 1 : lastRow;

    // تجهيز كل البيانات في صف واحد (13 عمود)
    const rowData = [
      number,                           // 1: الرقم
      getCurrentDate(),                 // 2: التاريخ
      data.project || '',               // 3: الفيلم
      data.stage || '',                 // 4: المرحلة
      data.subtype || '',               // 5: المرحلة الفرعية
      data.city || '',                  // 6: المدينة
      cleanText(data.element) || '',    // 7: العنصر
      data.details || '',               // 8: التفاصيل
      data.assignedTo || '',            // 9: المسؤول
      '⬜ لم يبدأ',                      // 10: الحالة
      data.dueDate || '',               // 11: تاريخ التسليم
      data.notes || '',                 // 12: ملاحظات
      ''                                // 13: الرابط (يُملأ لاحقاً)
    ];

    // حفظ كل البيانات دفعة واحدة
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    // إنشاء فولدر تلقائي إذا كانت المرحلة "التصوير"
    const isShootingStage = data.stage === 'التصوير' || data.stage === 'تصوير' ||
                           (data.stage && data.stage.toLowerCase() === 'shooting');
    if (isShootingStage && data.element) {
      try {
        createShootingFolder(data.project, data.city || data.subtype, newRow, data.element);
      } catch (folderErr) {
        console.error('خطأ في إنشاء الفولدر:', folderErr);
        // لا نوقف العملية بسبب خطأ الفولدر
      }
    }

    showSuccess('تم إضافة الحركة بنجاح');
    return { success: true, row: newRow };

  } catch (error) {
    console.error('خطأ في إضافة الحركة:', error);
    showError('فشل في إضافة الحركة: ' + error.message);
    throw error; // إعادة الخطأ للـ client
  }
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

  const range = e.range;
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const col = range.getColumn();

  // تحديد الأعمدة ديناميكياً
  const projectCol = getColumnByHeader(sheet, 'الفيلم');
  const stageCol = getColumnByHeader(sheet, 'المرحلة');
  const subtypeCol = getColumnByHeader(sheet, 'المرحلة الفرعية');
  const elementCol = getColumnByHeader(sheet, 'العنصر');
  const statusCol = getColumnByHeader(sheet, 'الحالة');
  const numberCol = getColumnByHeader(sheet, '#');
  const dateCol = getColumnByHeader(sheet, 'التاريخ');
  const dueDateCol = getColumnByHeader(sheet, 'تاريخ التسليم');
  const linkCol = getColumnByHeader(sheet, 'الرابط');

  if (projectCol === -1 || stageCol === -1) return;

  // معالجة كل صف في النطاق المعدل لدعم النسخ واللصق
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    if (currentRow <= 1) continue;

    // 0. تطبيع التواريخ تلقائياً إلى dd/mm/yyyy
    if ((col === dateCol || col === dueDateCol) && col !== -1) {
      const cell = sheet.getRange(currentRow, col);
      const value = cell.getValue();
      if (value) {
        normalizeDateCell_(cell, value);
      }
    }

    // 1. ملء التاريخ والرقم تلقائياً عند إدخال بيانات جديدة (إذا كان المشروع موجوداً)
    if (col === projectCol) {
      const project = sheet.getRange(currentRow, projectCol).getValue();
      if (project) {
        if (numberCol !== -1) {
          const numberCell = sheet.getRange(currentRow, numberCol);
          if (!numberCell.getValue()) {
            numberCell.setValue(currentRow - 1);
          }
        }
        if (dateCol !== -1) {
          const dateCell = sheet.getRange(currentRow, dateCol);
          if (!dateCell.getValue()) {
            dateCell.setValue(getCurrentDate());
          }
        }
      }
    }

    // 2. تلوين الصف عند تغيير الحالة
    if (col === statusCol && statusCol !== -1) {
      const status = sheet.getRange(currentRow, statusCol).getValue();
      colorRowByStatus(sheet, currentRow, status);
    }

    // 3. تحديث الأنواع الفرعية عند تغيير المرحلة
    if (col === stageCol) {
      const stage = sheet.getRange(currentRow, stageCol).getValue();
      const subtypes = getSubtypesFromSettings(stage);

      if (subtypes && subtypes.length > 0 && subtypeCol !== -1) {
        const subtypeCell = sheet.getRange(currentRow, subtypeCol);
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(subtypes, true)
          .setAllowInvalid(true)
          .build();
        subtypeCell.setDataValidation(rule);
      } else if (subtypeCol !== -1) {
        // إزالة التحقق لبعض المراحل الخاصة
        const stageKey = Object.keys(STAGES).find(key => STAGES[key].name === stage);
        if (stageKey === 'PRODUCTION' || stageKey === 'SHOOTING') {
          sheet.getRange(currentRow, subtypeCol).clearDataValidations();
        }
      }
      
      // [NEW] Trigger Smart Automation (Shortcuts & Folder Logic)
      const project = (projectCol !== -1) ? sheet.getRange(currentRow, projectCol).getValue() : '';
      if (project && stage) {
        try {
          onProjectStageChange(project, stage);
        } catch (err) {
          console.error('خطأ في الأتمتة:', err);
        }
      }
    }

    // 4. إنشاء فولدر تلقائي للتصوير (الإنتاج) عند إدخال العنصر
    if (col === elementCol && elementCol !== -1) {
      const stage = (stageCol !== -1) ? sheet.getRange(currentRow, stageCol).getValue() : '';
      const isShootingStage = stage === 'الإنتاج' || stage === 'التصوير';
      const elementValue = sheet.getRange(currentRow, elementCol).getValue();

      if (isShootingStage && elementValue) {
        const project = (projectCol !== -1) ? sheet.getRange(currentRow, projectCol).getValue() : '';
        const subtype = (subtypeCol !== -1) ? sheet.getRange(currentRow, subtypeCol).getValue() : '';
        const existingLink = (linkCol !== -1) ? sheet.getRange(currentRow, linkCol).getValue() : '';

        if (!existingLink && project) {
          createShootingFolder(project, subtype, currentRow, elementValue);
        }
      }
    }
  }
}

/**
 * الحصول على حركات مشروع معين
 */
function getProjectMovements(projectName) {
  const movements = getAllMovements();
  return movements.filter(m => m.project === projectName);
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
    const status = String(m.status || '');
    // إحصائيات الحالة
    if (status.includes('تم')) stats.completed++;
    else if (status.includes('جاري')) stats.inProgress++;
    else if (status.includes('انتظار')) stats.waiting++;
    else if (status.includes('متأخر')) stats.delayed++;
    else stats.notStarted++;

    // إحصائيات المراحل
    if (!stats.byStage[m.stage]) {
      stats.byStage[m.stage] = { total: 0, completed: 0 };
    }
    stats.byStage[m.stage].total++;
    if (status.includes('تم')) {
      stats.byStage[m.stage].completed++;
    }
  }

  return stats;
}

/**
 * الحصول على جميع الحركات - نسخة محسنة للأداء
 */
function getAllMovements() {
  try {
    const sheet = getSheet(SHEETS.MOVEMENT);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1) return [];

    // جلب البيانات والهيدرات في نطاق محدد بدقة للسرعة
    const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
    const allValues = dataRange.getValues();
    const headers = allValues[0];
    const data = allValues.slice(1);

    // تحديد الفهارس (Indexes) مع معالجة المرادفات برمجياً
    const getIdx = (columnName) => {
      const search = normalizeString(columnName);
      return headers.findIndex(h => normalizeString(h) === search);
    };

    const idx = {
      number: getIdx('#'),
      date: getIdx('التاريخ'),
      project: getIdx('الفيلم'),
      stage: getIdx('المرحلة'),
      subtype: getIdx('المرحلة الفرعية'),
      city: getIdx('المدينة'),
      element: getIdx('العنصر'),
      details: getIdx('التفاصيل'),
      assignedTo: getIdx('المسؤول'),
      status: getIdx('الحالة'),
      dueDate: getIdx('تاريخ التسليم'),
      notes: getIdx('ملاحظات')
    };

    // التحقق من الهيدرات الأساسية
    if (idx.project === -1) {
      idx.project = getIdx('المشروع');
      if (idx.project === -1) return [];
    }

    const movements = [];
    for (const row of data) {
      const projectName = row[idx.project];
      if (!projectName || String(projectName).trim() === "") continue;

      const dateVal = idx.date !== -1 ? row[idx.date] : '';
      const dueDateVal = idx.dueDate !== -1 ? row[idx.dueDate] : '';

      movements.push({
        number: idx.number !== -1 ? row[idx.number] : '',
        date: (dateVal instanceof Date) ? dateVal.toISOString() : dateVal,
        project: String(projectName).trim(),
        stage: idx.stage !== -1 ? row[idx.stage] : '',
        subtype: idx.subtype !== -1 ? row[idx.subtype] : '',
        city: idx.city !== -1 ? row[idx.city] : '',
        element: idx.element !== -1 ? row[idx.element] : '',
        details: idx.details !== -1 ? row[idx.details] : '',
        assignedTo: idx.assignedTo !== -1 ? row[idx.assignedTo] : '',
        status: idx.status !== -1 ? row[idx.status] : '',
        dueDate: (dueDateVal instanceof Date) ? dueDateVal.toISOString() : dueDateVal,
        notes: idx.notes !== -1 ? row[idx.notes] : ''
      });
    }
    return movements;
  } catch (e) {
    console.error('Error in getAllMovements:', e);
    return [];
  }
}

/**
 * الحصول على المهام المتأخرة
 */
function getDelayedTasks() {
  const movements = getAllMovements();
  const today = new Date();

  return movements.filter(m => {
    const status = String(m.status || '');
    if (status.includes('تم') || status.includes('ملغي')) return false;
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
