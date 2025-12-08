/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف إدارة المشاريع
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على جميع الدوال المتعلقة بإدارة المشاريع
 * بما في ذلك الإضافة، التعديل، الحذف، والاستعلام
 *
 * ملاحظة مهمة: النظام يعتمد على أسماء الأعمدة وليس أرقامها
 * مما يسمح بإضافة أعمدة جديدة دون التأثير على الكود
 */

// حالات المشروع
const PROJECT_STATUS = {
  ACTIVE: 'نشط',
  PAUSED: 'متوقف',
  COMPLETED: 'منتهي',
  CANCELLED: 'ملغي'
};

// ═══════════════════════════════════════════════════════════════════════════════
// دوال البحث عن الأعمدة بالاسم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على فهرس العمود بناءً على اسم الهيدر
 * @param {Sheet} sheet الشيت
 * @param {string} headerName اسم الهيدر
 * @returns {number} رقم العمود (1-indexed) أو -1 إذا لم يوجد
 */
function getColumnByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(headerName);
  return index >= 0 ? index + 1 : -1;
}

/**
 * الحصول على جميع فهارس الأعمدة
 * @param {Sheet} sheet الشيت
 * @returns {Object} كائن يحتوي على أسماء الأعمدة وفهارسها
 */
function getProjectColumnIndices(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};

  headers.forEach((header, index) => {
    indices[header] = index + 1;
  });

  return indices;
}

/**
 * الحصول على نطاق أعمدة المراحل (checkboxes)
 * ملاحظة: أعمدة المراحل تبدأ من العمود 10 (PHASE_START_COL)
 * لتجنب التطابق الخاطئ مع أعمدة مثل "تاريخ التسليم المتوقع"
 * @param {Sheet} sheet الشيت
 * @returns {Object} { startCol, endCol, count, headers }
 */
function getPhaseColumnsRange(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const phaseHeaders = [];
  let startCol = -1;
  let endCol = -1;

  // البحث عن أعمدة المراحل (تحتوي على أيقونات)
  // مهم: نبدأ البحث من العمود 10 (PHASE_START_COL) فقط
  // لتجنب التطابق مع أعمدة مثل "تاريخ التسليم المتوقع" التي تحتوي كلمة "التسليم"
  headers.forEach((header, index) => {
    const colNum = index + 1;

    // تجاهل الأعمدة قبل PHASE_START_COL (الأعمدة 1-9)
    if (colNum < PHASE_START_COL) return;

    // تجاهل أعمدة النظام (التواريخ)
    if (header && header.includes('تاريخ')) return;

    const stageMatch = Object.values(STAGES).find(s => header.includes(s.icon) || header.includes(s.name));
    if (stageMatch) {
      if (startCol === -1) startCol = colNum;
      endCol = colNum;
      phaseHeaders.push({ header, col: colNum, stage: stageMatch });
    }
  });

  return {
    startCol,
    endCol,
    count: endCol > 0 && startCol > 0 ? endCol - startCol + 1 : 0,
    headers: phaseHeaders
  };
}

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

  // الحصول على فهارس الأعمدة بناءً على الأسماء
  const cols = getProjectColumnIndices(sheet);
  const phaseRange = getPhaseColumnsRange(sheet);

  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return data.map(row => {
    // جمع قيم المراحل ديناميكياً
    const phases = {};
    phaseRange.headers.forEach(ph => {
      const stageKey = ph.stage.id.toLowerCase().replace('_', '');
      // تحويل SHOOT_FIELD إلى shootField
      const camelKey = ph.stage.id.toLowerCase().replace(/_([a-z])/g, (m, p1) => p1.toUpperCase());
      phases[camelKey] = row[ph.col - 1];
    });

    return {
      code: cols[PROJECT_HEADERS.CODE] ? row[cols[PROJECT_HEADERS.CODE] - 1] : '',
      name: row[cols[PROJECT_HEADERS.NAME] - 1],
      type: row[cols[PROJECT_HEADERS.TYPE] - 1],
      startDate: row[cols[PROJECT_HEADERS.START_DATE] - 1],
      endDate: row[cols[PROJECT_HEADERS.END_DATE] - 1],
      status: row[cols[PROJECT_HEADERS.STATUS] - 1],
      channel: row[cols[PROJECT_HEADERS.CHANNEL] - 1],
      program: row[cols[PROJECT_HEADERS.PROGRAM] - 1],
      notes: row[cols[PROJECT_HEADERS.NOTES] - 1],
      phases: phases,
      createdAt: row[cols[PROJECT_HEADERS.CREATED_AT] - 1],
      updatedAt: row[cols[PROJECT_HEADERS.UPDATED_AT] - 1]
    };
  }).filter(project => project.name); // تصفية الصفوف الفارغة
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
 * يستخدم أسماء الأعمدة وليس أرقامها
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
      SpreadsheetApp.getActiveSpreadsheet().toast('كود المشروع موجود مسبقاً', 'خطأ', 3);
      return false;
    }

    // الحصول على فهارس الأعمدة بناءً على الأسماء
    const cols = getProjectColumnIndices(sheet);
    const phaseRange = getPhaseColumnsRange(sheet);

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

    // الحصول على عدد الأعمدة الكلي
    const lastCol = sheet.getLastColumn();

    // إنشاء صف فارغ بحجم الأعمدة
    const rowData = new Array(lastCol).fill('');

    // ملء البيانات الأساسية بناءً على أسماء الأعمدة
    if (cols[PROJECT_HEADERS.CODE]) rowData[cols[PROJECT_HEADERS.CODE] - 1] = code;
    if (cols[PROJECT_HEADERS.NAME]) rowData[cols[PROJECT_HEADERS.NAME] - 1] = projectData.name;
    if (cols[PROJECT_HEADERS.TYPE]) rowData[cols[PROJECT_HEADERS.TYPE] - 1] = projectData.type || PROJECT_TYPES[0];
    if (cols[PROJECT_HEADERS.START_DATE]) rowData[cols[PROJECT_HEADERS.START_DATE] - 1] = startDate;
    if (cols[PROJECT_HEADERS.END_DATE]) rowData[cols[PROJECT_HEADERS.END_DATE] - 1] = endDate;
    if (cols[PROJECT_HEADERS.STATUS]) rowData[cols[PROJECT_HEADERS.STATUS] - 1] = projectData.status || PROJECT_STATUS.ACTIVE;
    if (cols[PROJECT_HEADERS.CHANNEL]) rowData[cols[PROJECT_HEADERS.CHANNEL] - 1] = projectData.channel || '';
    if (cols[PROJECT_HEADERS.PROGRAM]) rowData[cols[PROJECT_HEADERS.PROGRAM] - 1] = projectData.program || '';
    if (cols[PROJECT_HEADERS.NOTES]) rowData[cols[PROJECT_HEADERS.NOTES] - 1] = projectData.notes || '';
    if (cols[PROJECT_HEADERS.CREATED_AT]) rowData[cols[PROJECT_HEADERS.CREATED_AT] - 1] = new Date();
    if (cols[PROJECT_HEADERS.UPDATED_AT]) rowData[cols[PROJECT_HEADERS.UPDATED_AT] - 1] = new Date();

    // ملء قيم المراحل ديناميكياً
    const phases = projectData.phases || {};
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

    // ملء أعمدة المراحل
    phaseRange.headers.forEach(ph => {
      // إيجاد مفتاح المرحلة المناسب
      const phaseKey = Object.keys(phaseMapping).find(key =>
        phaseMapping[key] === ph.stage.id
      );
      if (phaseKey) {
        rowData[ph.col - 1] = Boolean(phases[phaseKey]);
      }
    });

    // إضافة البيانات في الصف المحدد
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    // تطبيق checkbox validation على أعمدة المراحل فقط (من عمود 10 فصاعداً)
    // نستخدم PHASE_START_COL كحد أدنى لضمان عدم إضافة checkboxes للأعمدة 1-9
    const actualStartCol = Math.max(phaseRange.startCol, PHASE_START_COL);
    if (actualStartCol >= PHASE_START_COL && phaseRange.endCol >= actualStartCol) {
      const actualCount = phaseRange.endCol - actualStartCol + 1;
      if (actualCount > 0) {
        const checkboxCells = sheet.getRange(targetRow, actualStartCol, 1, actualCount);
        const checkboxRule = SpreadsheetApp.newDataValidation()
          .requireCheckbox()
          .build();
        checkboxCells.setDataValidation(checkboxRule);
      }
    }

    // التأكد من حفظ البيانات فوراً
    SpreadsheetApp.flush();

    // إظهار رسالة نجاح
    SpreadsheetApp.getActiveSpreadsheet().toast('تم إضافة المشروع: ' + projectData.name, 'تمت الإضافة ✓', 3);

    console.log('تم إضافة المشروع بنجاح: ' + projectData.name + ' في الصف: ' + targetRow);
    return true;

  } catch (error) {
    console.error('Error adding project:', error.toString());
    return false;
  }
}

/**
 * تحديث مشروع موجود
 * يستخدم أسماء الأعمدة وليس أرقامها
 * @param {string} projectName اسم المشروع (المعرف الرئيسي)
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updateProject(projectName, updates) {
  try {
    const sheet = getSheet(SHEETS.PROJECTS);
    if (!sheet) return false;

    // الحصول على فهارس الأعمدة بناءً على الأسماء
    const cols = getProjectColumnIndices(sheet);
    const phaseRange = getPhaseColumnsRange(sheet);

    // البحث عن الصف بناءً على اسم المشروع
    const nameCol = cols[PROJECT_HEADERS.NAME];
    if (!nameCol) {
      console.error('عمود اسم الفيلم غير موجود');
      return false;
    }

    const rowIndex = findRowByValue(SHEETS.PROJECTS, nameCol, projectName);
    if (rowIndex === -1) {
      showError('المشروع غير موجود');
      return false;
    }

    // تحديث الحقول المطلوبة بناءً على أسماء الأعمدة
    if (updates.name !== undefined && cols[PROJECT_HEADERS.NAME]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.NAME]).setValue(updates.name);
    }
    if (updates.type !== undefined && cols[PROJECT_HEADERS.TYPE]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.TYPE]).setValue(updates.type);
    }
    if (updates.startDate !== undefined && cols[PROJECT_HEADERS.START_DATE]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.START_DATE]).setValue(updates.startDate);
    }
    if (updates.endDate !== undefined && cols[PROJECT_HEADERS.END_DATE]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.END_DATE]).setValue(updates.endDate);
    }
    if (updates.status !== undefined && cols[PROJECT_HEADERS.STATUS]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.STATUS]).setValue(updates.status);
    }
    if (updates.channel !== undefined && cols[PROJECT_HEADERS.CHANNEL]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.CHANNEL]).setValue(updates.channel);
    }
    if (updates.program !== undefined && cols[PROJECT_HEADERS.PROGRAM]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.PROGRAM]).setValue(updates.program);
    }
    if (updates.notes !== undefined && cols[PROJECT_HEADERS.NOTES]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.NOTES]).setValue(updates.notes);
    }

    // تحديث المراحل ديناميكياً
    if (updates.phases) {
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

      Object.keys(updates.phases).forEach(phaseKey => {
        const stageId = phaseMapping[phaseKey];
        if (stageId) {
          // البحث عن العمود المناسب في المراحل
          const phaseHeader = phaseRange.headers.find(ph => ph.stage.id === stageId);
          if (phaseHeader) {
            sheet.getRange(rowIndex, phaseHeader.col).setValue(Boolean(updates.phases[phaseKey]));
          }
        }
      });
    }

    // تحديث تاريخ آخر تعديل
    if (cols[PROJECT_HEADERS.UPDATED_AT]) {
      sheet.getRange(rowIndex, cols[PROJECT_HEADERS.UPDATED_AT]).setValue(new Date());
    }

    return true;
  } catch (error) {
    console.error('Error updating project:', error);
    return false;
  }
}

/**
 * تغيير حالة المشروع
 * @param {string} projectName اسم المشروع
 * @param {string} newStatus الحالة الجديدة
 * @returns {boolean} نجاح العملية
 */
function changeProjectStatus(projectName, newStatus) {
  return updateProject(projectName, { status: newStatus });
}

/**
 * إنشاء كود مشروع جديد تلقائياً
 * @returns {string} كود المشروع (مثال: P25001)
 */
function generateProjectCode() {
  const year = new Date().getFullYear().toString().substr(-2);
  const sheet = getSheet(SHEETS.PROJECTS);

  if (!sheet) return `P${year}001`;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return `P${year}001`;

  // الحصول على عمود الكود ديناميكياً
  const cols = getProjectColumnIndices(sheet);
  const codeCol = cols[PROJECT_HEADERS.CODE];
  if (!codeCol) return `P${year}001`;

  // البحث عن آخر كود في نفس السنة
  const codes = sheet.getRange(2, codeCol, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(code => code && code.toString().startsWith(`P${year}`));

  if (codes.length === 0) return `P${year}001`;

  // استخراج أعلى رقم
  const numbers = codes.map(code => parseInt(code.toString().replace(`P${year}`, ''), 10));
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
    (project.code && project.code.toLowerCase().includes(searchTerm)) ||
    (project.channel && project.channel.toLowerCase().includes(searchTerm)) ||
    (project.program && project.program.toLowerCase().includes(searchTerm)) ||
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
 * ترتيب الحقول يتوافق مع ترتيب الأعمدة في الشيت
 */
function showAddProjectDialog() {
  // تحضير خيارات الأنواع
  const typeOptions = PROJECT_TYPES.map(t => '<option value="' + t + '">' + t + '</option>').join('');

  // تحضير خيارات الحالة
  const statusOptions = Object.values(PROJECT_STATUS).map(s =>
    '<option value="' + s + '"' + (s === PROJECT_STATUS.ACTIVE ? ' selected' : '') + '>' + s + '</option>'
  ).join('');

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
      .row { display: flex; gap: 15px; }
      .row .form-group { flex: 1; }
    </style>

    <h3>إضافة مشروع جديد</h3>

    <div class="form-group">
      <label>اسم الفيلم *</label>
      <input type="text" id="projectName" required placeholder="أدخل اسم الفيلم">
    </div>

    <div class="row">
      <div class="form-group">
        <label>نوع الفيلم</label>
        <select id="projectType">
          ${typeOptions}
        </select>
      </div>
      <div class="form-group">
        <label>الحالة</label>
        <select id="projectStatus">
          ${statusOptions}
        </select>
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>تاريخ البداية</label>
        <input type="date" id="startDate" value="${today}">
      </div>
      <div class="form-group">
        <label>تاريخ التسليم المتوقع</label>
        <input type="date" id="endDate">
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>اسم القناة</label>
        <input type="text" id="channel" placeholder="اسم القناة">
      </div>
      <div class="form-group">
        <label>اسم البرنامج</label>
        <input type="text" id="program" placeholder="اسم البرنامج">
      </div>
    </div>

    <div class="form-group">
      <label>ملاحظات</label>
      <textarea id="notes" rows="2" placeholder="أي ملاحظات إضافية..."></textarea>
    </div>

    <div class="form-group">
      <label>المراحل المطلوبة</label>
      <div class="phases-grid">
        ${phaseCheckboxes}
      </div>
    </div>

    <div class="buttons">
      <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>
      <button class="btn btn-primary" onclick="submitForm()">حفظ</button>
    </div>

    <script>
      function submitForm() {
        var projectName = document.getElementById('projectName').value;
        if (!projectName || projectName.trim() === '') {
          alert('يرجى إدخال اسم الفيلم');
          return;
        }

        var data = {
          name: projectName,
          type: document.getElementById('projectType').value,
          startDate: document.getElementById('startDate').value,
          endDate: document.getElementById('endDate').value,
          status: document.getElementById('projectStatus').value,
          channel: document.getElementById('channel').value,
          program: document.getElementById('program').value,
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
  `).setWidth(550).setHeight(700);

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
