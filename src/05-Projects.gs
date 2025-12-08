/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف إدارة المشاريع
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على جميع الدوال المتعلقة بإدارة المشاريع
 * بما في ذلك الإضافة، التعديل، الحذف، والاستعلام
 */

// ═══════════════════════════════════════════════════════════════════════════════
// ثوابت أعمدة شيت المشاريع (محدثة)
// ═══════════════════════════════════════════════════════════════════════════════

const PROJECT_COLS = {
  CODE: 1,           // كود المشروع
  NAME: 2,           // اسم الفيلم
  TYPE: 3,           // النوع
  START_DATE: 4,     // تاريخ البداية
  END_DATE: 5,       // التسليم المتوقع
  STATUS: 6,         // الحالة (نشط/متوقف/منتهي)
  NOTES: 7,          // ملاحظات
  // أعمدة المراحل (checkboxes)
  PHASE_PAPER: 8,
  PHASE_FIXER: 9,
  PHASE_SHOOT_FIELD: 10,
  PHASE_SHOOT_INT: 11,
  PHASE_SHOOT_DRAMA: 12,
  PHASE_VO: 13,
  PHASE_ANIMATION: 14,
  PHASE_INFOGRAPH: 15,
  PHASE_MONTAGE: 16,
  PHASE_ARCHIVE: 17,
  PHASE_REVIEW: 18,
  PHASE_DELIVERY: 19,
  // أعمدة النظام
  CREATED_AT: 20,
  UPDATED_AT: 21
};

// حالات المشروع
const PROJECT_STATUS = {
  ACTIVE: 'نشط',
  PAUSED: 'متوقف',
  COMPLETED: 'منتهي',
  CANCELLED: 'ملغي'
};

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الاستعلام عن المشاريع
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على جميع المشاريع
 * @returns {Array} مصفوفة من كائنات المشاريع
 */
function getAllProjects() {
  const sheet = getSheet(SHEETS.PROJECTS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  return data.map(row => ({
    code: row[PROJECT_COLS.CODE - 1],
    name: row[PROJECT_COLS.NAME - 1],
    type: row[PROJECT_COLS.TYPE - 1],
    startDate: row[PROJECT_COLS.START_DATE - 1],
    endDate: row[PROJECT_COLS.END_DATE - 1],
    status: row[PROJECT_COLS.STATUS - 1],
    notes: row[PROJECT_COLS.NOTES - 1],
    phases: {
      paper: row[PROJECT_COLS.PHASE_PAPER - 1],
      fixer: row[PROJECT_COLS.PHASE_FIXER - 1],
      shootField: row[PROJECT_COLS.PHASE_SHOOT_FIELD - 1],
      shootInt: row[PROJECT_COLS.PHASE_SHOOT_INT - 1],
      shootDrama: row[PROJECT_COLS.PHASE_SHOOT_DRAMA - 1],
      vo: row[PROJECT_COLS.PHASE_VO - 1],
      animation: row[PROJECT_COLS.PHASE_ANIMATION - 1],
      infograph: row[PROJECT_COLS.PHASE_INFOGRAPH - 1],
      montage: row[PROJECT_COLS.PHASE_MONTAGE - 1],
      archive: row[PROJECT_COLS.PHASE_ARCHIVE - 1],
      review: row[PROJECT_COLS.PHASE_REVIEW - 1],
      delivery: row[PROJECT_COLS.PHASE_DELIVERY - 1]
    },
    createdAt: row[PROJECT_COLS.CREATED_AT - 1],
    updatedAt: row[PROJECT_COLS.UPDATED_AT - 1]
  })).filter(project => project.code); // تصفية الصفوف الفارغة
}

/**
 * الحصول على المشاريع النشطة فقط
 * @returns {Array} مصفوفة المشاريع النشطة
 */
function getActiveProjects() {
  const allProjects = getAllProjects();
  return allProjects.filter(project => project.status === PROJECT_STATUS.ACTIVE);
}

/**
 * الحصول على مشروع بالكود
 * @param {string} code كود المشروع
 * @returns {Object|null} كائن المشروع أو null
 */
function getProjectByCode(code) {
  if (!code) return null;

  const allProjects = getAllProjects();
  return allProjects.find(project => project.code === code) || null;
}

/**
 * الحصول على مشروع بالاسم
 * @param {string} name اسم المشروع
 * @returns {Object|null} كائن المشروع أو null
 */
function getProjectByName(name) {
  if (!name) return null;

  const allProjects = getAllProjects();
  return allProjects.find(project => project.name === name) || null;
}

/**
 * الحصول على المراحل المفعلة لمشروع معين
 * @param {string} projectCode كود المشروع أو اسمه
 * @returns {Array} مصفوفة المراحل المفعلة
 */
function getProjectPhases(projectCode) {
  // البحث بالكود أو الاسم
  let project = getProjectByCode(projectCode);
  if (!project) {
    project = getProjectByName(projectCode);
  }

  if (!project) return [];

  const enabledPhases = [];

  // مطابقة المراحل المفعلة مع كائن STAGES
  const phaseMapping = {
    paper: 'PAPER',
    fixer: 'FIXER',
    shootField: 'SHOOT_FIELD',
    shootInt: 'SHOOT_INT',
    shootDrama: 'SHOOT_DRAMA',
    vo: 'VO',
    animation: 'ANIMATION',
    infograph: 'INFOGRAPH',
    montage: 'MONTAGE',
    archive: 'ARCHIVE',
    review: 'REVIEW',
    delivery: 'DELIVERY'
  };

  Object.keys(phaseMapping).forEach(key => {
    if (project.phases[key] === true) {
      const stageId = phaseMapping[key];
      if (STAGES[stageId]) {
        enabledPhases.push(STAGES[stageId]);
      }
    }
  });

  // ترتيب المراحل حسب order
  enabledPhases.sort((a, b) => a.order - b.order);

  return enabledPhases;
}

/**
 * الحصول على أسماء المشاريع النشطة (للقوائم المنسدلة)
 * @returns {Array} مصفوفة أسماء المشاريع
 */
function getActiveProjectNames() {
  const activeProjects = getActiveProjects();
  return activeProjects.map(project => project.name);
}

/**
 * الحصول على أكواد المشاريع النشطة
 * @returns {Array} مصفوفة أكواد المشاريع
 */
function getActiveProjectCodes() {
  const activeProjects = getActiveProjects();
  return activeProjects.map(project => project.code);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إضافة وتعديل المشاريع
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة مشروع جديد
 * @param {Object} projectData بيانات المشروع
 * @returns {boolean} نجاح العملية
 */
function addProject(projectData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PROJECTS);

    if (!sheet) {
      console.error('شيت المشاريع غير موجود: ' + SHEETS.PROJECTS);
      return false;
    }

    // التحقق من البيانات المطلوبة
    if (!projectData || !projectData.name) {
      console.error('اسم المشروع مطلوب');
      return false;
    }

    // إنشاء كود المشروع تلقائياً
    const code = projectData.code || generateProjectCode();

    // التحقق من عدم تكرار الكود
    if (getProjectByCode(code)) {
      console.error('كود المشروع موجود مسبقاً: ' + code);
      return false;
    }

    // تحويل التواريخ من strings إلى Date objects
    let startDate = new Date();
    if (projectData.startDate && projectData.startDate !== '') {
      startDate = new Date(projectData.startDate);
    }

    let endDate = '';
    if (projectData.endDate && projectData.endDate !== '') {
      endDate = new Date(projectData.endDate);
    }

    // إيجاد أول صف فاضي بعد الـ header
    const lastRow = sheet.getLastRow();
    let targetRow = 2; // البداية من الصف 2 (بعد الهيدر)

    if (lastRow >= 2) {
      // البحث عن أول صف فاضي
      const firstColData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < firstColData.length; i++) {
        if (!firstColData[i][0] || firstColData[i][0] === '') {
          targetRow = i + 2;
          break;
        }
        targetRow = i + 3; // الصف التالي بعد آخر صف ممتلئ
      }
    }

    // تجهيز صف البيانات (بدون المراحل - سنضيفها كـ checkboxes)
    const basicData = [
      code,
      projectData.name,
      projectData.type || PROJECT_TYPES[0],
      startDate,
      endDate,
      projectData.status || PROJECT_STATUS.ACTIVE,
      projectData.notes || ''
    ];

    // قيم المراحل - تحويل صريح للقيم البوليانية
    const phases = projectData.phases || {};
    const phaseValues = [
      Boolean(phases.paper),
      Boolean(phases.fixer),
      Boolean(phases.shootField),
      Boolean(phases.shootInt),
      Boolean(phases.shootDrama),
      Boolean(phases.vo),
      Boolean(phases.animation),
      Boolean(phases.infograph),
      Boolean(phases.montage),
      Boolean(phases.archive),
      Boolean(phases.review),
      Boolean(phases.delivery)
    ];

    // تواريخ النظام
    const systemDates = [new Date(), new Date()];

    // دمج كل البيانات
    const rowData = [...basicData, ...phaseValues, ...systemDates];

    // إضافة البيانات في الصف المحدد
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    // تطبيق checkbox validation على أعمدة المراحل (H إلى S = أعمدة 8 إلى 19)
    const phaseRange = sheet.getRange(targetRow, 8, 1, 12); // 12 مرحلة
    const checkboxRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    phaseRange.setDataValidation(checkboxRule);

    // التأكد من حفظ البيانات فوراً
    SpreadsheetApp.flush();

    // إظهار رسالة نجاح
    SpreadsheetApp.getActiveSpreadsheet().toast('تم إضافة المشروع: ' + projectData.name, 'تمت الإضافة ✓', 3);

    console.log('تم إضافة المشروع بنجاح: ' + code + ' في الصف: ' + targetRow);
    return true;

  } catch (error) {
    console.error('Error adding project:', error.toString());
    return false;
  }
}

/**
 * تحديث مشروع موجود
 * @param {string} code كود المشروع
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updateProject(code, updates) {
  try {
    const sheet = getSheet(SHEETS.PROJECTS);
    if (!sheet) return false;

    const rowIndex = findRowByValue(SHEETS.PROJECTS, PROJECT_COLS.CODE, code);
    if (rowIndex === -1) {
      showError('المشروع غير موجود');
      return false;
    }

    // تحديث الحقول المطلوبة
    if (updates.name !== undefined) {
      sheet.getRange(rowIndex, PROJECT_COLS.NAME).setValue(updates.name);
    }
    if (updates.type !== undefined) {
      sheet.getRange(rowIndex, PROJECT_COLS.TYPE).setValue(updates.type);
    }
    if (updates.startDate !== undefined) {
      sheet.getRange(rowIndex, PROJECT_COLS.START_DATE).setValue(updates.startDate);
    }
    if (updates.endDate !== undefined) {
      sheet.getRange(rowIndex, PROJECT_COLS.END_DATE).setValue(updates.endDate);
    }
    if (updates.status !== undefined) {
      sheet.getRange(rowIndex, PROJECT_COLS.STATUS).setValue(updates.status);
    }
    if (updates.notes !== undefined) {
      sheet.getRange(rowIndex, PROJECT_COLS.NOTES).setValue(updates.notes);
    }

    // تحديث المراحل
    if (updates.phases) {
      const phaseColumns = {
        paper: PROJECT_COLS.PHASE_PAPER,
        fixer: PROJECT_COLS.PHASE_FIXER,
        shootField: PROJECT_COLS.PHASE_SHOOT_FIELD,
        shootInt: PROJECT_COLS.PHASE_SHOOT_INT,
        shootDrama: PROJECT_COLS.PHASE_SHOOT_DRAMA,
        vo: PROJECT_COLS.PHASE_VO,
        animation: PROJECT_COLS.PHASE_ANIMATION,
        infograph: PROJECT_COLS.PHASE_INFOGRAPH,
        montage: PROJECT_COLS.PHASE_MONTAGE,
        archive: PROJECT_COLS.PHASE_ARCHIVE,
        review: PROJECT_COLS.PHASE_REVIEW,
        delivery: PROJECT_COLS.PHASE_DELIVERY
      };

      Object.keys(updates.phases).forEach(phase => {
        if (phaseColumns[phase]) {
          sheet.getRange(rowIndex, phaseColumns[phase]).setValue(updates.phases[phase]);
        }
      });
    }

    // تحديث تاريخ آخر تعديل
    sheet.getRange(rowIndex, PROJECT_COLS.UPDATED_AT).setValue(new Date());

    return true;
  } catch (error) {
    console.error('Error updating project:', error);
    return false;
  }
}

/**
 * تغيير حالة المشروع
 * @param {string} code كود المشروع
 * @param {string} newStatus الحالة الجديدة
 * @returns {boolean} نجاح العملية
 */
function changeProjectStatus(code, newStatus) {
  return updateProject(code, { status: newStatus });
}

/**
 * إنشاء كود مشروع جديد تلقائياً
 * @returns {string} كود المشروع
 */
function generateProjectCode() {
  const year = new Date().getFullYear().toString().substr(-2);
  const sheet = getSheet(SHEETS.PROJECTS);

  if (!sheet) return `P${year}001`;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return `P${year}001`;

  // البحث عن آخر كود في نفس السنة
  const codes = sheet.getRange(2, PROJECT_COLS.CODE, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(code => code && code.startsWith(`P${year}`));

  if (codes.length === 0) return `P${year}001`;

  // استخراج أعلى رقم
  const numbers = codes.map(code => parseInt(code.replace(`P${year}`, ''), 10));
  const maxNum = Math.max(...numbers);

  return `P${year}${(maxNum + 1).toString().padStart(3, '0')}`;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال البحث والتصفية
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * البحث في المشاريع
 * @param {string} query نص البحث
 * @returns {Array} نتائج البحث
 */
function searchProjects(query) {
  if (!query) return [];

  const allProjects = getAllProjects();
  const searchTerm = query.toLowerCase();

  return allProjects.filter(project =>
    project.name.toLowerCase().includes(searchTerm) ||
    project.code.toLowerCase().includes(searchTerm) ||
    (project.notes && project.notes.toLowerCase().includes(searchTerm))
  );
}

/**
 * الحصول على المشاريع حسب الحالة
 * @param {string} status الحالة
 * @returns {Array} المشاريع المطابقة
 */
function getProjectsByStatus(status) {
  const allProjects = getAllProjects();
  return allProjects.filter(project => project.status === status);
}

/**
 * الحصول على المشاريع حسب النوع
 * @param {string} type النوع
 * @returns {Array} المشاريع المطابقة
 */
function getProjectsByType(type) {
  const allProjects = getAllProjects();
  return allProjects.filter(project => project.type === type);
}

/**
 * الحصول على المشاريع المتأخرة
 * @returns {Array} المشاريع المتأخرة
 */
function getDelayedProjects() {
  const activeProjects = getActiveProjects();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return activeProjects.filter(project => {
    if (!project.endDate) return false;
    const endDate = new Date(project.endDate);
    endDate.setHours(0, 0, 0, 0);
    return endDate < today;
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإحصائيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على إحصائيات المشاريع
 * @returns {Object} كائن الإحصائيات
 */
function getProjectsStats() {
  const allProjects = getAllProjects();

  const stats = {
    total: allProjects.length,
    active: 0,
    paused: 0,
    completed: 0,
    cancelled: 0,
    delayed: 0,
    byType: {}
  };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  allProjects.forEach(project => {
    // إحصاء حسب الحالة
    switch (project.status) {
      case PROJECT_STATUS.ACTIVE:
        stats.active++;
        // التحقق من التأخير
        if (project.endDate) {
          const endDate = new Date(project.endDate);
          if (endDate < today) stats.delayed++;
        }
        break;
      case PROJECT_STATUS.PAUSED:
        stats.paused++;
        break;
      case PROJECT_STATUS.COMPLETED:
        stats.completed++;
        break;
      case PROJECT_STATUS.CANCELLED:
        stats.cancelled++;
        break;
    }

    // إحصاء حسب النوع
    if (project.type) {
      stats.byType[project.type] = (stats.byType[project.type] || 0) + 1;
    }
  });

  return stats;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال واجهة المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * فتح نموذج إضافة مشروع جديد
 */
function showAddProjectDialog() {
  // تحضير خيارات الأنواع
  const typeOptions = PROJECT_TYPES.map(t => '<option value="' + t + '">' + t + '</option>').join('');

  // تحضير checkboxes المراحل
  const phaseCheckboxes = Object.values(STAGES).map(s =>
    '<div class="phase-item">' +
    '<input type="checkbox" id="phase_' + s.id + '" checked>' +
    '<label for="phase_' + s.id + '">' + s.icon + ' ' + s.name + '</label>' +
    '</div>'
  ).join('');

  // تاريخ اليوم
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; direction: rtl; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px; }
      .btn-primary { background: #1565c0; color: white; }
      .btn-secondary { background: #757575; color: white; }
      .phases-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
      .phase-item { display: flex; align-items: center; }
      .phase-item input { width: auto; margin-left: 5px; }
      .buttons { margin-top: 20px; text-align: left; }
    </style>

    <h3>إضافة مشروع جديد</h3>

    <div class="form-group">
      <label>اسم المشروع *</label>
      <input type="text" id="projectName" required>
    </div>

    <div class="form-group">
      <label>النوع</label>
      <select id="projectType">
        ${typeOptions}
      </select>
    </div>

    <div class="form-group">
      <label>تاريخ البداية</label>
      <input type="date" id="startDate" value="${today}">
    </div>

    <div class="form-group">
      <label>تاريخ التسليم المتوقع</label>
      <input type="date" id="endDate">
    </div>

    <div class="form-group">
      <label>المراحل المطلوبة</label>
      <div class="phases-grid">
        ${phaseCheckboxes}
      </div>
    </div>

    <div class="form-group">
      <label>ملاحظات</label>
      <textarea id="notes" rows="3"></textarea>
    </div>

    <div class="buttons">
      <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>
      <button class="btn btn-primary" onclick="submitForm()">حفظ</button>
    </div>

    <script>
      function submitForm() {
        var projectName = document.getElementById('projectName').value;
        if (!projectName || projectName.trim() === '') {
          alert('يرجى إدخال اسم المشروع');
          return;
        }

        var data = {
          name: projectName,
          type: document.getElementById('projectType').value,
          startDate: document.getElementById('startDate').value,
          endDate: document.getElementById('endDate').value,
          notes: document.getElementById('notes').value,
          phases: {
            paper: document.getElementById('phase_PAPER').checked === true,
            fixer: document.getElementById('phase_FIXER').checked === true,
            shootField: document.getElementById('phase_SHOOT_FIELD').checked === true,
            shootInt: document.getElementById('phase_SHOOT_INT').checked === true,
            shootDrama: document.getElementById('phase_SHOOT_DRAMA').checked === true,
            vo: document.getElementById('phase_VO').checked === true,
            animation: document.getElementById('phase_ANIMATION').checked === true,
            infograph: document.getElementById('phase_INFOGRAPH').checked === true,
            montage: document.getElementById('phase_MONTAGE').checked === true,
            archive: document.getElementById('phase_ARCHIVE').checked === true,
            review: document.getElementById('phase_REVIEW').checked === true,
            delivery: document.getElementById('phase_DELIVERY').checked === true
          }
        };

        google.script.run
          .withSuccessHandler(function(result) {
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            alert('حدث خطأ: ' + err.message);
            google.script.host.close();
          })
          .addProject(data);
      }
    </script>
  `).setWidth(500).setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة مشروع جديد');
}

/**
 * عرض ملخص المشاريع
 */
function showProjectsSummary() {
  const stats = getProjectsStats();

  const message = `
ملخص المشاريع
━━━━━━━━━━━━━━━━━━━━━━━━━

📊 إجمالي المشاريع: ${stats.total}

حسب الحالة:
• 🟢 نشط: ${stats.active}
• 🟡 متوقف: ${stats.paused}
• 🔵 منتهي: ${stats.completed}
• ⚫ ملغي: ${stats.cancelled}

⚠️ مشاريع متأخرة: ${stats.delayed}

━━━━━━━━━━━━━━━━━━━━━━━━━

حسب النوع:
${Object.entries(stats.byType).map(([type, count]) => `• ${type}: ${count}`).join('\n')}
  `.trim();

  showInfo(message, 'ملخص المشاريع');
}
