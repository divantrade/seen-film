/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * القائمة والـ Triggers
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * إنشاء القائمة عند فتح الملف
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🎬 نظام الإنتاج')
    // المشاريع
    .addSubMenu(ui.createMenu('📁 المشاريع')
      .addItem('إضافة مشروع جديد', 'showAddProjectForm')
      .addItem('فتح فولدر المشروع', 'openProjectFolder')
      .addItem('إنشاء فولدرات المشروع', 'createProjectFoldersManual'))

    // الفريق
    .addSubMenu(ui.createMenu('👥 الفريق')
      .addItem('إضافة عضو جديد', 'showAddTeamMemberForm')
      .addItem('تغيير حالة العضو', 'toggleTeamMemberStatus'))

    // الحركة
    .addSubMenu(ui.createMenu('📋 الحركة')
      .addItem('إضافة حركة جديدة', 'showAddMovementForm')
      .addSeparator()
      .addItem('✅ وضع علامة "تم"', 'markAsCompleted')
      .addItem('🔄 وضع علامة "جاري"', 'markAsInProgress')
      .addItem('🔴 وضع علامة "متأخر"', 'markAsDelayed')
      .addSeparator()
      .addItem('📁 إنشاء فولدر', 'createFolderForMovement')
      .addItem('🔄 تحديث المهام المتأخرة', 'updateDelayedTasks'))

    // التقارير
    .addSubMenu(ui.createMenu('📊 التقارير')
      .addItem('تقرير حالة فيلم (Timeline)', 'showFilmTimelineReport')
      .addItem('📥 إنشاء تقرير تفصيلي (Sheet)', 'createDetailedFilmReport')
      .addItem('تقارير الشركة المجمعة', 'showCompanyReport')
      .addSeparator()
      .addItem('تحديث لوحة التحكم', 'refreshDashboard')
      .addItem('تقرير نصي سريع', 'showProjectReport'))

    .addSeparator()

    // الإعدادات
    .addSubMenu(ui.createMenu('⚙️ الإعدادات')
      .addItem('التحقق من فولدر الإنتاج', 'checkMainFolderSettings')
      .addItem('تحديث القوائم المنسدلة', 'updateAllDropdowns')
      .addItem('🔧 إصلاح شيت الحركة', 'fixMovementSheet')
      .addSeparator()
      .addItem('👁️ إظهار شيت الروابط', 'showFolderLinksSheet')
      .addItem('🙈 إخفاء شيت الروابط', 'hideFolderLinksSheet')
      .addSeparator()
      .addItem('🔄 تهيئة النظام', 'initializeSystem')
      .addItem('🔧 تشخيص النظام', 'debugSettings'))

    .addToUi();

  // تحديث القوائم المنسدلة
  try {
    updateMovementDropdowns();
  } catch (e) {
    console.log('تخطي تحديث القوائم:', e);
  }
}

/**
 * Trigger عند التعديل
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  try {
    switch (sheetName) {
      case SHEETS.TEAM:
        onTeamEdit(e);
        break;
      case SHEETS.MOVEMENT:
        onMovementEdit(e);
        break;
      case SHEETS.PROJECTS:
        onProjectEdit(e);
        break;
      case SHEETS.DASHBOARD:
        onDashboardEdit(e);
        break;
    }
  } catch (error) {
    console.error('خطأ في onEdit:', error);
  }
}

/**
 * Trigger عند تعديل شيت المشاريع
 */
function onProjectEdit(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row <= 1) return;

  // توليد الكود تلقائياً عند إدخال اسم المشروع
  if (col === PROJECT_COLS.NAME) {
    const currentCode = sheet.getRange(row, PROJECT_COLS.CODE).getValue();
    if (!currentCode && e.value) {
      const newCode = generateProjectCode();
      sheet.getRange(row, PROJECT_COLS.CODE).setValue(newCode);

      // ملاحظة: لا يمكن إنشاء الفولدرات من onEdit بسبب صلاحيات محدودة
      // استخدم القائمة: المشاريع > إنشاء فولدرات المشروع
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'تم إنشاء الكود. لإنشاء فولدرات المشروع، استخدم القائمة: المشاريع > إنشاء فولدرات المشروع',
        'تنبيه',
        5
      );
    }
  }
}

/**
 * تحديث جميع القوائم المنسدلة
 */
function updateAllDropdowns() {
  updateMovementDropdowns();
  showSuccess('تم تحديث القوائم المنسدلة');
}

/**
 * نموذج إضافة مشروع جديد
 */
function showAddProjectForm() {
  const team = getActiveTeamNames();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; direction: rtl; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; box-sizing: border-box; }
      button { background: #1565C0; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-left: 10px; }
      button:hover { background: #0D47A1; }
      .cancel { background: #757575; }
      .row { display: flex; gap: 10px; }
      .row .form-group { flex: 1; }
    </style>
    <form id="projectForm">
      <div class="form-group">
        <label>اسم الفيلم *</label>
        <input type="text" id="name" required>
      </div>
      <div class="form-group">
        <label>نوع الفيلم *</label>
        <select id="type" required>
          <option value="">اختر النوع</option>
          ${PROJECT_TYPES.map(t => `<option value="${t}">${t}</option>`).join('')}
        </select>
      </div>
      <div class="row">
        <div class="form-group">
          <label>المنتج المسؤول</label>
          <select id="producer">
            <option value="">اختر المنتج</option>
            ${team.map(t => `<option value="${t}">${t}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label>المونتير</label>
          <select id="editor">
            <option value="">اختر المونتير</option>
            ${team.map(t => `<option value="${t}">${t}</option>`).join('')}
          </select>
        </div>
      </div>
      <div class="row">
        <div class="form-group">
          <label>تاريخ البداية</label>
          <input type="date" id="startDate">
        </div>
        <div class="form-group">
          <label>تاريخ التسليم</label>
          <input type="date" id="endDate">
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
      document.getElementById('projectForm').onsubmit = function(e) {
        e.preventDefault();
        const data = {
          name: document.getElementById('name').value,
          type: document.getElementById('type').value,
          producer: document.getElementById('producer').value,
          editor: document.getElementById('editor').value,
          startDate: document.getElementById('startDate').value,
          endDate: document.getElementById('endDate').value,
          notes: document.getElementById('notes').value
        };
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .addProject(data);
      };
    </script>
  `)
    .setWidth(500)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة مشروع جديد');
}

/**
 * إضافة مشروع جديد
 */
function addProject(data) {
  const sheet = getSheet(SHEETS.PROJECTS);
  const lastRow = getLastRowInColumn(sheet, PROJECT_COLS.NAME);
  const newRow = Math.max(lastRow + 1, 2);

  // توليد الكود
  const code = generateProjectCode();

  // إضافة البيانات
  sheet.getRange(newRow, PROJECT_COLS.CODE).setValue(code);
  sheet.getRange(newRow, PROJECT_COLS.NAME).setValue(cleanText(data.name));
  sheet.getRange(newRow, PROJECT_COLS.TYPE).setValue(data.type);
  sheet.getRange(newRow, PROJECT_COLS.PRODUCER).setValue(data.producer || '');
  sheet.getRange(newRow, PROJECT_COLS.EDITOR).setValue(data.editor || '');
  sheet.getRange(newRow, PROJECT_COLS.START_DATE).setValue(data.startDate || '');
  sheet.getRange(newRow, PROJECT_COLS.END_DATE).setValue(data.endDate || '');
  sheet.getRange(newRow, PROJECT_COLS.STATUS).setValue('نشط');
  sheet.getRange(newRow, PROJECT_COLS.NOTES).setValue(data.notes || '');

  // إنشاء فولدرات المشروع
  const folderUrl = createProjectFolderStructure(data.name, code);
  if (folderUrl) {
    sheet.getRange(newRow, PROJECT_COLS.FOLDER_LINK).setValue(folderUrl);
  }

  // تحديث القوائم المنسدلة
  updateMovementDropdowns();

  showSuccess('تم إضافة المشروع: ' + data.name);
}

/**
 * إنشاء فولدرات المشروع يدوياً
 */
function createProjectFoldersManual() {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== SHEETS.PROJECTS) {
    showError('يجب أن تكون في شيت المشاريع');
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    showError('اختر مشروعاً من القائمة');
    return;
  }

  const projectName = sheet.getRange(row, PROJECT_COLS.NAME).getValue();
  const projectCode = sheet.getRange(row, PROJECT_COLS.CODE).getValue();

  if (!projectName) {
    showError('المشروع غير صالح');
    return;
  }

  const folderUrl = createProjectFolderStructure(projectName, projectCode);

  if (folderUrl) {
    sheet.getRange(row, PROJECT_COLS.FOLDER_LINK).setValue(folderUrl);
    showSuccess('تم إنشاء فولدرات المشروع');
  }
}

/**
 * تثبيت الـ Triggers
 */
function installTriggers() {
  // حذف الـ Triggers القديمة
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }

  // إضافة trigger للتحديث اليومي
  ScriptApp.newTrigger('updateDelayedTasks')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  showSuccess('تم تثبيت الـ Triggers');
}
