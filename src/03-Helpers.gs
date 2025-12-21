/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * الدوال المساعدة
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * الحصول على التاريخ الحالي منسق
 */
function getCurrentDate() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
}

/**
 * الحصول على التاريخ والوقت الحالي منسق
 */
function getCurrentDateTime() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);
}

/**
 * الحصول على شيت بالاسم
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * الحصول على آخر صف يحتوي على بيانات
 */
function getLastRow(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i].some(cell => cell !== '')) {
      return i + 1;
    }
  }
  return 1;
}

/**
 * الحصول على آخر صف في عمود معين
 */
function getLastRowInColumn(sheet, column) {
  const data = sheet.getRange(1, column, sheet.getMaxRows(), 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== '') {
      return i + 1;
    }
  }
  return 1;
}

/**
 * توليد كود المشروع التلقائي
 * الصيغة: P + السنة (آخر رقمين) + رقم تسلسلي (3 أرقام)
 * مثال: P25001
 */
function generateProjectCode() {
  const sheet = getSheet(SHEETS.PROJECTS);
  const year = new Date().getFullYear().toString().slice(-2);
  const prefix = 'P' + year;

  // البحث عن آخر كود
  const lastRow = getLastRowInColumn(sheet, PROJECT_COLS.CODE);

  if (lastRow <= 1) {
    return prefix + '001';
  }

  const codes = sheet.getRange(2, PROJECT_COLS.CODE, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (const [code] of codes) {
    if (code && code.toString().startsWith(prefix)) {
      const num = parseInt(code.toString().slice(-3));
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }

  return prefix + String(maxNum + 1).padStart(3, '0');
}

/**
 * توليد كود عضو الفريق التلقائي
 * الصيغة: رمز الدور + رقم تسلسلي (3 أرقام)
 * مثال: PRD-001
 */
function generateTeamCode(role) {
  const sheet = getSheet(SHEETS.TEAM);
  const prefix = ROLE_CODES[role] || 'OTH';

  // البحث عن آخر كود لهذا الدور
  const lastRow = getLastRowInColumn(sheet, TEAM_COLS.CODE);

  if (lastRow <= 1) {
    return prefix + '-001';
  }

  const codes = sheet.getRange(2, TEAM_COLS.CODE, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (const [code] of codes) {
    if (code && code.toString().startsWith(prefix)) {
      const num = parseInt(code.toString().split('-')[1]);
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }

  return prefix + '-' + String(maxNum + 1).padStart(3, '0');
}

/**
 * الحصول على قائمة المشاريع النشطة
 */
function getActiveProjects() {
  const sheet = getSheet(SHEETS.PROJECTS);
  if (!sheet) return [];

  const lastRow = getLastRowInColumn(sheet, PROJECT_COLS.NAME);

  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, PROJECT_COLS.NOTES).getValues();
  const projects = [];

  for (const row of data) {
    const status = row[PROJECT_COLS.STATUS - 1];
    const name = row[PROJECT_COLS.NAME - 1];

    // إضافة المشروع إذا كان نشطاً أو لم تُحدد حالته
    if (name && (status === 'نشط' || status === '' || !status)) {
      projects.push({
        code: row[PROJECT_COLS.CODE - 1],
        name: name
      });
    }
  }

  return projects;
}

/**
 * الحصول على قائمة جميع المشاريع (بغض النظر عن الحالة)
 */
function getAllProjects() {
  const sheet = getSheet(SHEETS.PROJECTS);
  if (!sheet) return [];

  const lastRow = getLastRowInColumn(sheet, PROJECT_COLS.NAME);

  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, PROJECT_COLS.NOTES).getValues();
  const projects = [];

  for (const row of data) {
    const name = row[PROJECT_COLS.NAME - 1];
    if (name) {
      projects.push({
        code: row[PROJECT_COLS.CODE - 1],
        name: name
      });
    }
  }

  return projects;
}

/**
 * الحصول على قائمة أسماء المشاريع النشطة
 */
function getActiveProjectNames() {
  const activeProjects = getActiveProjects();

  // إذا لم توجد مشاريع نشطة، أرجع كل المشاريع
  if (activeProjects.length === 0) {
    return getAllProjects().map(p => p.name);
  }

  return activeProjects.map(p => p.name);
}

/**
 * الحصول على قائمة أعضاء الفريق النشطين
 */
function getActiveTeamMembers() {
  const sheet = getSheet(SHEETS.TEAM);
  const lastRow = getLastRowInColumn(sheet, TEAM_COLS.NAME);

  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, TEAM_COLS.NOTES).getValues();
  const members = [];

  for (const row of data) {
    if (row[TEAM_COLS.STATUS - 1] === 'نشط') {
      members.push({
        code: row[TEAM_COLS.CODE - 1],
        name: row[TEAM_COLS.NAME - 1],
        role: row[TEAM_COLS.ROLE - 1]
      });
    }
  }

  return members;
}

/**
 * الحصول على قائمة أسماء أعضاء الفريق النشطين
 */
function getActiveTeamNames() {
  return getActiveTeamMembers().map(m => m.name);
}

/**
 * الحصول على الأنواع الفرعية لمرحلة معينة
 */
function getSubtypesForStage(stageName) {
  for (const key in STAGES) {
    if (STAGES[key].name === stageName) {
      return STAGES[key].subtypes;
    }
  }
  return [];
}

/**
 * تلوين الصف حسب الحالة
 */
function colorRowByStatus(sheet, row, status) {
  const statusObj = Object.values(STATUS).find(s =>
    s.name === status || `${s.icon} ${s.name}` === status
  );

  if (statusObj) {
    const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
    range.setBackground(statusObj.color);
  }
}

/**
 * الحصول على لون الحالة
 */
function getStatusColor(status) {
  const statusObj = Object.values(STATUS).find(s =>
    s.name === status || `${s.icon} ${s.name}` === status
  );
  return statusObj ? statusObj.color : COLORS.BACKGROUND;
}

/**
 * التحقق من وجود قيمة في عمود
 */
function valueExistsInColumn(sheet, column, value) {
  const lastRow = getLastRowInColumn(sheet, column);
  if (lastRow <= 1) return false;

  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  return values.some(([v]) => v === value);
}

/**
 * البحث عن صف بقيمة في عمود معين
 */
function findRowByValue(sheet, column, value) {
  const lastRow = getLastRowInColumn(sheet, column);
  if (lastRow <= 1) return -1;

  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === value) {
      return i + 2; // +2 لأن البيانات تبدأ من الصف 2
    }
  }
  return -1;
}

/**
 * الحصول على الأنواع الفرعية لمرحلة معينة من شيت الإعدادات
 */
function getSubtypesFromSettings(stageName) {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return [];

  // قراءة البيانات من الأعمدة E, F (بداية من الصف 6)
  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return [];

  const data = sheet.getRange(6, 5, lastRow - 5, 2).getValues();
  const subtypes = [];

  for (const row of data) {
    if (row[0] === stageName && row[1]) {
      subtypes.push(row[1]);
    }
  }

  return subtypes;
}

/**
 * الحصول على جميع المراحل من شيت الإعدادات
 */
function getStagesFromSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return STAGE_NAMES; // fallback للقيم الافتراضية

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return STAGE_NAMES;

  const data = sheet.getRange(6, 5, lastRow - 5, 1).getValues();
  const stages = new Set();

  for (const row of data) {
    if (row[0]) {
      stages.add(row[0]);
    }
  }

  // إضافة المراحل التي ليس لها أنواع فرعية (مثل التصوير)
  for (const name of STAGE_NAMES) {
    stages.add(name);
  }

  return Array.from(stages);
}

/**
 * عرض رسالة نجاح
 */
function showSuccess(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'نجاح', 3);
}

/**
 * عرض رسالة خطأ
 */
function showError(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'خطأ', 5);
}

/**
 * عرض رسالة معلومات
 */
function showInfo(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'معلومات', 3);
}

/**
 * التحقق من صحة البريد الإلكتروني
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * تنظيف النص من المسافات الزائدة
 */
function cleanText(text) {
  if (!text) return '';
  return text.toString().trim().replace(/\s+/g, ' ');
}
