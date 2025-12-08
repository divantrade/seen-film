/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف القوائم المنسدلة الديناميكية
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على الدوال المسؤولة عن إنشاء وتحديث القوائم المنسدلة
 * القوائم تعتمد على بعضها البعض (Dependent Dropdowns)
 */

// ═══════════════════════════════════════════════════════════════════════════════
// الثوابت الخاصة بالأعمدة للدروب داون
// ═══════════════════════════════════════════════════════════════════════════════

const DROPDOWN_COLUMNS = {
  // أعمدة شيت الحركة (محدثة)
  MOVEMENT: {
    NUMBER: 1,          // عمود #
    DATE: 2,            // عمود التاريخ
    PROJECT_CODE: 3,    // عمود كود المشروع
    PROJECT_NAME: 4,    // عمود اسم المشروع
    STAGE: 5,           // عمود المرحلة
    SUBTYPE: 6,         // عمود النوع الفرعي
    ELEMENT: 7,         // عمود العنصر
    ACTION: 8,          // عمود الإجراء
    ASSIGNED_TO: 9,     // عمود المسؤول
    STATUS: 10          // عمود الحالة
  }
};

// ═══════════════════════════════════════════════════════════════════════════════
// دالة onEdit - المُشغل الرئيسي
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * دالة تُنفذ تلقائياً عند تعديل أي خلية
 * تتحكم في الدروب داون الديناميكية
 * @param {Object} e حدث التعديل
 */
function onEdit(e) {
  // التحقق من وجود الحدث
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // الشيتات المسموح بها فقط: الحركة والضيوف
  const allowedSheets = [SHEETS.MOVEMENT, SHEETS.GUESTS];
  if (!allowedSheets.includes(sheetName)) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const value = e.value;

  // تجاهل الهيدر
  if (row === 1) return;

  // معالجة حسب الشيت
  switch (sheetName) {
    case SHEETS.MOVEMENT:
      handleMovementEdit(sheet, row, col, value);
      break;
    case SHEETS.GUESTS:
      handleGuestsEdit(sheet, row, col, value);
      break;
  }

  // تلوين الحالة تلقائياً
  autoColorStatus(sheet, row, col, value);
}

// ═══════════════════════════════════════════════════════════════════════════════
// معالجة تعديلات شيت الحركة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * معالجة التعديلات في شيت الحركة
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {number} col رقم العمود
 * @param {string} value القيمة الجديدة
 */
function handleMovementEdit(sheet, row, col, value) {
  const cols = DROPDOWN_COLUMNS.MOVEMENT;

  // عند اختيار كود المشروع ← جلب اسم المشروع وتحديث المراحل
  if (col === cols.PROJECT_CODE && value) {
    const project = getProjectByCode(value);
    if (project) {
      // تعبئة اسم المشروع
      sheet.getRange(row, cols.PROJECT_NAME).setValue(project.name);
      // تحديث المراحل المتاحة
      updateStagesDropdown(sheet, row, project.code);
      // مسح المرحلة والنوع الفرعي السابقين
      sheet.getRange(row, cols.STAGE).clearContent();
      sheet.getRange(row, cols.SUBTYPE).clearContent();
      // إضافة التاريخ تلقائياً
      const dateCell = sheet.getRange(row, cols.DATE);
      if (!dateCell.getValue()) {
        dateCell.setValue(new Date());
      }
    }
  }

  // عند اختيار اسم المشروع ← جلب كود المشروع وتحديث المراحل
  if (col === cols.PROJECT_NAME && value) {
    const project = getProjectByName(value);
    if (project) {
      // تعبئة كود المشروع
      sheet.getRange(row, cols.PROJECT_CODE).setValue(project.code);
      // تحديث المراحل المتاحة
      updateStagesDropdown(sheet, row, project.code || project.name);
      // مسح المرحلة والنوع الفرعي السابقين
      sheet.getRange(row, cols.STAGE).clearContent();
      sheet.getRange(row, cols.SUBTYPE).clearContent();
      // إضافة التاريخ تلقائياً
      const dateCell = sheet.getRange(row, cols.DATE);
      if (!dateCell.getValue()) {
        dateCell.setValue(new Date());
      }
    }
  }

  // عند اختيار المرحلة ← تحديث الأنواع الفرعية
  if (col === cols.STAGE && value) {
    updateSubtypesDropdown(sheet, row, value);
    // مسح النوع الفرعي السابق
    sheet.getRange(row, cols.SUBTYPE).clearContent();
  }
}

/**
 * تحديث قائمة المراحل بناءً على المشروع المختار
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {string} projectCode كود المشروع
 */
function updateStagesDropdown(sheet, row, projectCode) {
  const cols = DROPDOWN_COLUMNS.MOVEMENT;
  const stageCell = sheet.getRange(row, cols.STAGE);

  // الحصول على المراحل المفعلة للمشروع
  const projectPhases = getProjectPhases(projectCode);

  if (projectPhases && projectPhases.length > 0) {
    // إنشاء قائمة منسدلة بالمراحل المفعلة فقط
    const stageNames = projectPhases.map(stage => `${stage.icon} ${stage.name}`);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(stageNames, true)
      .setAllowInvalid(false)
      .build();
    stageCell.setDataValidation(rule);
  } else {
    // إذا لم يوجد مراحل محددة، عرض جميع المراحل
    const allStages = getStageNamesWithIcons();
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allStages, true)
      .setAllowInvalid(false)
      .build();
    stageCell.setDataValidation(rule);
  }
}

/**
 * تحديث قائمة الأنواع الفرعية بناءً على المرحلة المختارة
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {string} stageName اسم المرحلة
 */
function updateSubtypesDropdown(sheet, row, stageName) {
  const cols = DROPDOWN_COLUMNS.MOVEMENT;
  const subtypeCell = sheet.getRange(row, cols.SUBTYPE);

  // استخراج اسم المرحلة بدون الأيقونة
  const cleanStageName = stageName.replace(/^[^\s]+\s/, '');

  // البحث عن المرحلة
  const stage = Object.values(STAGES).find(s => s.name === cleanStageName);

  if (stage && stage.subtypes && stage.subtypes.length > 0) {
    // المرحلة لها أنواع فرعية
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(stage.subtypes, true)
      .setAllowInvalid(false)
      .build();
    subtypeCell.setDataValidation(rule);
  } else {
    // المرحلة ليس لها أنواع فرعية - إزالة القائمة
    subtypeCell.clearDataValidations();
    subtypeCell.setValue('-');
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// معالجة تعديلات الشيتات الأخرى
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * معالجة التعديلات في شيت الضيوف
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {number} col رقم العمود
 * @param {string} value القيمة الجديدة
 */
function handleGuestsEdit(sheet, row, col, value) {
  // عند اختيار المشروع، تحديث الخيارات المتاحة
  if (col === 3 && value) { // عمود المشروع
    // يمكن إضافة منطق إضافي هنا
  }
}

/**
 * معالجة التعديلات في شيت التعليق الصوتي
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {number} col رقم العمود
 * @param {string} value القيمة الجديدة
 */
function handleVoiceOverEdit(sheet, row, col, value) {
  // يمكن إضافة منطق إضافي هنا
}

/**
 * معالجة التعديلات في شيت الرسوم المتحركة
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {number} col رقم العمود
 * @param {string} value القيمة الجديدة
 */
function handleAnimationEdit(sheet, row, col, value) {
  // يمكن إضافة منطق إضافي هنا
}

/**
 * معالجة التعديلات في شيت الأرشيف
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {number} col رقم العمود
 * @param {string} value القيمة الجديدة
 */
function handleArchiveEdit(sheet, row, col, value) {
  // يمكن إضافة منطق إضافي هنا
}

// ═══════════════════════════════════════════════════════════════════════════════
// تلوين الحالة تلقائياً
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تلوين خلية الحالة تلقائياً بناءً على القيمة
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {number} col رقم العمود
 * @param {string} value القيمة الجديدة
 */
function autoColorStatus(sheet, row, col, value) {
  // التحقق من أن العمود هو عمود الحالة
  const sheetName = sheet.getName();
  let statusCol;

  switch (sheetName) {
    case SHEETS.MOVEMENT:
      statusCol = DROPDOWN_COLUMNS.MOVEMENT.STATUS;
      break;
    case SHEETS.GUESTS:
      statusCol = 5; // عمود حالة التواصل
      break;
    default:
      return;
  }

  if (col === statusCol && value) {
    const color = getStatusColor(value);
    sheet.getRange(row, col).setBackground(color);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// إعداد القوائم المنسدلة للشيتات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إعداد جميع القوائم المنسدلة في شيت الحركة
 */
function setupMovementDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MOVEMENT);

  if (!sheet) {
    showError('شيت الحركة غير موجود');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 100);
  const cols = DROPDOWN_COLUMNS.MOVEMENT;

  // قائمة المشاريع النشطة
  const activeProjects = getActiveProjects();
  const projectCodes = activeProjects.map(p => p.code).filter(c => c);
  const projectNames = activeProjects.map(p => p.name);

  // قائمة أكواد المشاريع
  if (projectCodes.length > 0) {
    const codeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(projectCodes, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, cols.PROJECT_CODE, lastRow - 1, 1).setDataValidation(codeRule);
  }

  // قائمة أسماء المشاريع
  if (projectNames.length > 0) {
    const nameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(projectNames, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, cols.PROJECT_NAME, lastRow - 1, 1).setDataValidation(nameRule);
  }

  // قائمة المراحل (الافتراضية - كل المراحل)
  const stageNames = getStageNamesWithIcons();
  const stageRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(stageNames, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, cols.STAGE, lastRow - 1, 1).setDataValidation(stageRule);

  // قائمة المسؤولين (الفريق النشط)
  const teamMembers = getTeamMembers();
  const teamNames = teamMembers.map(t => t.name);

  if (teamNames.length > 0) {
    const teamRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(teamNames, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, cols.ASSIGNED_TO, lastRow - 1, 1).setDataValidation(teamRule);
  }

  // قائمة الحالات
  const statusNames = getStatusNamesWithIcons();
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusNames, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, cols.STATUS, lastRow - 1, 1).setDataValidation(statusRule);

  showToast('تم إعداد القوائم المنسدلة لشيت الحركة', 'نجاح');
}

/**
 * إعداد القوائم المنسدلة في شيت المشاريع
 * يستخدم البحث الديناميكي عن الأعمدة بالاسم
 * ملاحظة: هذه الدالة تضيف dropdown فقط لعمودي النوع والحالة
 * ولا تضيف أي validation للأعمدة الأخرى (التواريخ، القناة، البرنامج، الملاحظات)
 * الـ checkboxes تُضاف فقط من عمود 10 (PHASE_START_COL) وما بعده
 */
function setupProjectsDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PROJECTS);

  if (!sheet) {
    showError('شيت المشاريع غير موجود');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 50);

  // الحصول على فهارس الأعمدة بالاسم
  const cols = getProjectColumnIndices(sheet);
  const phaseRange = getPhaseColumnsRange(sheet);

  // 1. مسح جميع الـ data validations من الأعمدة 1-9 (الأعمدة الأساسية)
  // هذا يضمن عدم وجود أي checkboxes أو validations خاطئة
  for (let col = 1; col <= 9; col++) {
    sheet.getRange(2, col, lastRow - 1, 1).clearDataValidations();
  }

  // 2. إضافة قائمة أنواع المشاريع (عمود 3)
  const typeCol = cols[PROJECT_HEADERS.TYPE];
  if (typeCol && typeCol <= 9) {
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PROJECT_TYPES, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, typeCol, lastRow - 1, 1).setDataValidation(typeRule);
  }

  // 3. إضافة قائمة حالات المشاريع (عمود 6)
  const statusCol = cols[PROJECT_HEADERS.STATUS];
  if (statusCol && statusCol <= 9) {
    const statusValues = ['نشط', 'متوقف', 'منتهي', 'ملغي'];
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusValues, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, lastRow - 1, 1).setDataValidation(statusRule);
  }

  // 4. إضافة checkboxes فقط لأعمدة المراحل (من عمود 10 فصاعداً)
  // نستخدم PHASE_START_COL (10) كحد أدنى لضمان عدم إضافة checkboxes للأعمدة 1-9
  const actualStartCol = Math.max(phaseRange.startCol, PHASE_START_COL);
  if (actualStartCol >= PHASE_START_COL && phaseRange.count > 0) {
    // حساب عدد الأعمدة الفعلي
    const actualCount = phaseRange.endCol - actualStartCol + 1;
    if (actualCount > 0) {
      const checkboxCells = sheet.getRange(2, actualStartCol, lastRow - 1, actualCount);
      checkboxCells.insertCheckboxes();
    }
  }

  showToast('تم إعداد القوائم المنسدلة لشيت المشاريع', 'نجاح');
}

/**
 * إعداد القوائم المنسدلة في شيت الفريق
 * الأعمدة: الكود(1), الاسم(2), الدور(3), البريد(4), الهاتف(5), المراحل(6), الحالة(7), تاريخ الانضمام(8), ملاحظات(9)
 */
function setupTeamDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TEAM);

  if (!sheet) {
    showError('شيت الفريق غير موجود');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 50);

  // قائمة الأدوار - عمود 3
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TEAM_ROLES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(roleRule);

  // قائمة الحالة (نشط/غير نشط) - عمود 7 (وليس 6)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['نشط', 'غير نشط'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, lastRow - 1, 1).setDataValidation(statusRule);

  showToast('تم إعداد القوائم المنسدلة لشيت الفريق', 'نجاح');
}

/**
 * إعداد القوائم المنسدلة في شيت الضيوف
 * الأعمدة: الكود(1), الاسم(2), المشروع(3), نوع المشاركة(4), حالة التواصل(5), حالة الأسئلة(6),
 *          البلد(7), مكان التصوير(8), تاريخ التصوير(9), حالة التصوير(10), يحتاج دوبلاج(11), ...
 */
function setupGuestsDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GUESTS);

  if (!sheet) {
    showError('شيت الضيوف غير موجود');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 50);

  // قائمة المشاريع النشطة - عمود 3
  const activeProjects = getActiveProjects();
  const projectNames = activeProjects.map(p => p.name);

  if (projectNames.length > 0) {
    const projectRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(projectNames, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(projectRule);
  }

  // قائمة نوع المشاركة - عمود 4
  const participationTypes = ['مقابلة', 'دراما', 'تعليق صوتي', 'أخرى'];
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(participationTypes, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, lastRow - 1, 1).setDataValidation(typeRule);

  // قائمة حالة التواصل - عمود 5
  const contactStatuses = ['لم يبدأ', 'جاري التواصل', 'تم التأكيد', 'رفض', 'مؤجل'];
  const contactRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(contactStatuses, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, lastRow - 1, 1).setDataValidation(contactRule);

  // قائمة حالة الأسئلة - عمود 6
  const questionStatuses = ['لم ترسل', 'أُرسلت', 'تم الرد'];
  const questionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(questionStatuses, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, lastRow - 1, 1).setDataValidation(questionRule);

  // قائمة حالة التصوير - عمود 10 (وليس 9)
  const shootStatuses = ['لم يتم', 'مجدول', 'تم', 'ملغي'];
  const shootRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(shootStatuses, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 10, lastRow - 1, 1).setDataValidation(shootRule);

  // قائمة يحتاج دوبلاج - عمود 11 (وليس 10)
  const dubbingRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['نعم', 'لا'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 11, lastRow - 1, 1).setDataValidation(dubbingRule);

  showToast('تم إعداد القوائم المنسدلة لشيت الضيوف', 'نجاح');
}

/**
 * إعداد القوائم المنسدلة في شيت المصورين
 */
function setupPhotographersDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PHOTOGRAPHERS);

  if (!sheet) {
    showError('شيت المصورين غير موجود');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 50);

  // قائمة التخصصات
  const specializations = ['ميداني', 'استوديو', 'درامي', 'جوي (درون)', 'متعدد'];
  const specRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(specializations, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(specRule); // عمود التخصص

  // قائمة الحالة
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['متاح', 'مشغول', 'غير متاح'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, lastRow - 1, 1).setDataValidation(statusRule); // عمود الحالة

  showToast('تم إعداد القوائم المنسدلة لشيت المصورين', 'نجاح');
}

// ═══════════════════════════════════════════════════════════════════════════════
// تحديث جميع القوائم المنسدلة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تحديث جميع القوائم المنسدلة في النظام
 * يُستدعى من القائمة
 */
function refreshAllDropdowns() {
  try {
    setupProjectsDropdowns();
    setupTeamDropdowns();
    setupPhotographersDropdowns();
    setupGuestsDropdowns();
    setupMovementDropdowns();

    showSuccess('تم تحديث جميع القوائم المنسدلة بنجاح');
  } catch (error) {
    showError('حدث خطأ أثناء تحديث القوائم:\n' + error.message);
  }
}

/**
 * تحديث ألوان الحالات في جميع الشيتات
 * يستخدم البحث الديناميكي عن أعمدة الحالة
 */
function refreshStatusColors() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // شيت الحركة - عمود الحالة ثابت
    const movementSheet = ss.getSheetByName(SHEETS.MOVEMENT);
    if (movementSheet) {
      colorStatusColumn(movementSheet, DROPDOWN_COLUMNS.MOVEMENT.STATUS);
    }

    // شيت المشاريع - عمود الحالة ديناميكي
    const projectsSheet = ss.getSheetByName(SHEETS.PROJECTS);
    if (projectsSheet) {
      const cols = getProjectColumnIndices(projectsSheet);
      const statusCol = cols[PROJECT_HEADERS.STATUS];
      if (statusCol) {
        colorStatusColumn(projectsSheet, statusCol);
      }
    }

    // شيت الضيوف - عمود الحالة ثابت
    const guestsSheet = ss.getSheetByName(SHEETS.GUESTS);
    if (guestsSheet) {
      colorStatusColumn(guestsSheet, 5);
    }

    showSuccess('تم تحديث ألوان الحالات بنجاح');
  } catch (error) {
    showError('حدث خطأ أثناء تحديث الألوان:\n' + error.message);
  }
}

/**
 * تلوين عمود الحالة في شيت معين
 * @param {Sheet} sheet الشيت
 * @param {number} statusCol رقم عمود الحالة
 */
function colorStatusColumn(sheet, statusCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    for (let row = 2; row <= lastRow; row++) {
      const cell = sheet.getRange(row, statusCol);
      const value = cell.getValue();
      if (value) {
        const color = getStatusColor(value);
        cell.setBackground(color);
      }
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// إنشاء Trigger للـ onEdit
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء trigger لدالة onEdit
 * يجب تشغيل هذه الدالة مرة واحدة فقط عند إعداد النظام
 */
function createOnEditTrigger() {
  // حذف أي triggers موجودة
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // إنشاء trigger جديد
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  showToast('تم إنشاء الـ Trigger بنجاح', 'نجاح');
}
