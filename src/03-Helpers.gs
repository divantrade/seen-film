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
 * تنظيف النصوص ومطابقتها (تجاهل المسافات وحالة الأحرف)
 */
/**
 * تطبيع النصوص - نسخة محسنة جداً للغة العربية
 */
function normalizeString(str) {
  if (str === null || str === undefined) return '';
  let normalized = str.toString().trim().toLowerCase();
  
  // تطبيع الألفات
  normalized = normalized.replace(/[أإآ]/g, 'ا');
  // تطبيع الهاء والتاء المربوطة
  normalized = normalized.replace(/ة/g, 'ه');
  // تطبيع الياء والألف المقصورة
  normalized = normalized.replace(/ى/g, 'ي');
  // إزالة التشكيل
  normalized = normalized.replace(/[\u064B-\u0652]/g, '');
  
  return normalized;
}

/**
 * تجهيز البيانات للإرسال للعميل (Client-Side)
 * يضمن تحويل كافة القيم إلى أنواع بسيطة (Strings/Numbers) لتجنب أخطاء السيرفر
 */
function sanitizeForClient(data) {
  if (data === null || data === undefined) return '';
  
  if (Array.isArray(data)) {
    return data.map(item => sanitizeForClient(item));
  }
  
  if (typeof data === 'object' && !(data instanceof Date)) {
    const sanitized = {};
    for (const key in data) {
      sanitized[key] = sanitizeForClient(data[key]);
    }
    return sanitized;
  }
  
  if (data instanceof Date) {
    if (isNaN(data.getTime())) return '';
    try {
      return data.toISOString();
    } catch (e) {
      return '';
    }
  }
  
  return data;
}

/**
 * تحليل مدخلات التاريخ بصيغ مختلفة وتحويلها إلى Date Object
 * يدعم الفواصل: / و . و -
 */
function parseDateInput_(dateValue) {
  if (!dateValue) return null;
  
  // إذا كان بالفعل Date Object
  if (dateValue instanceof Date) {
    return { dateObj: dateValue, isValid: !isNaN(dateValue.getTime()) };
  }
  
  // إذا كان نص
  const dateStr = dateValue.toString().trim();
  if (!dateStr) return null;
  
  // نمط التاريخ: dd/mm/yyyy أو dd.mm.yyyy أو dd-mm-yyyy
  const regex = /^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{4})$/;
  const match = dateStr.match(regex);
  
  if (match) {
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1; // الأشهر تبدأ من 0
    const year = parseInt(match[3], 10);
    
    const dateObj = new Date(year, month, day);
    
    // التحقق من صحة التاريخ
    if (dateObj.getDate() === day && dateObj.getMonth() === month && dateObj.getFullYear() === year) {
      return { dateObj: dateObj, isValid: true };
    }
  }
  
  // محاولة أخيرة: استخدام Date constructor
  try {
    const dateObj = new Date(dateStr);
    if (!isNaN(dateObj.getTime())) {
      return { dateObj: dateObj, isValid: true };
    }
  } catch (e) {
    console.warn('Could not parse date:', dateStr);
  }
  
  return null;
}

/**
 * تطبيع خلية تاريخ واحدة
 */
function normalizeDateCell_(range, value) {
  const parseResult = parseDateInput_(value);
  
  if (parseResult && parseResult.isValid) {
    range.setValue(parseResult.dateObj);
    range.setNumberFormat('dd/mm/yyyy');
    return true;
  }
  
  return false;
}

/**
 * الحصول على شيت بالاسم (يبحث بالاسم الجديد مع الأيقونة ثم القديم بدونها)
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // جرب الاسم المعطى أولاً
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  // خريطة الأسماء القديمة والجديدة
  const nameMap = {
    'المشاريع': '🎬 المشاريع',
    'الفريق': '👥 الفريق',
    'الحركة': '📋 الحركة',
    'داشبورد': '📊 داشبورد',
    'الإعدادات': '⚙️ الإعدادات',
    'روابط الفولدرات': '📁 روابط الفولدرات',
    'المستخدمين': '🔐 المستخدمين',
    // العكس أيضاً
    '🎬 المشاريع': 'المشاريع',
    '👥 الفريق': 'الفريق',
    '📋 الحركة': 'الحركة',
    '📊 داشبورد': 'داشبورد',
    '⚙️ الإعدادات': 'الإعدادات',
    '📁 روابط الفولدرات': 'روابط الفولدرات',
    '🔐 المستخدمين': 'المستخدمين'
  };

  // جرب الاسم البديل
  if (nameMap[sheetName]) {
    sheet = ss.getSheetByName(nameMap[sheetName]);
    if (sheet) return sheet;
  }

  return null;
}

// ذاكرة مؤقتة للهيدرات لتقليل عمليات الـ API
const _headerCache = {};

/**
 * الحصول على رقم العمود بناءً على اسم الهيدر
 */
function getColumnByHeader(sheet, headerName) {
  if (!sheet) return -1;
  const sheetName = sheet.getName();
  
  // التحقق من الكاش أولاً
  if (!_headerCache[sheetName]) {
    _headerCache[sheetName] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  
  const headers = _headerCache[sheetName];
  const normalizedSearch = normalizeString(headerName);
  
  // قاموس المرادفات الشائعة للهيدرات
  const synonyms = {
    'الاسم': ['الاسم', 'اسم', 'إسم', 'الاسم بالكامل', 'الأسم', 'name'],
    'الحالة': ['الحالة', 'حالة', 'الوضع', 'التوفر', 'status'],
    'المدينة': ['المدينة', 'مدينه', 'مدينة السكن', 'الموقع', 'city'],
    'المسؤول': ['المسؤول', 'المسئول', 'assigned to', 'person'],
    'الكود': ['الكود', 'كود', 'الرقم التعريفى', 'رقم التعريف', 'code'],
    'الدور': ['الدور', 'دور', 'الوظيفة', 'المهنة', 'role', 'position'],
    'الفيلم': ['الفيلم', 'فيلم', 'اسم الفيلم', 'المشروع', 'project'],
    'المرحلة': ['المرحلة', 'مرحلة', 'القسم', 'المرحلة الرئيسية', 'stage'],
    'المرحلة الفرعية': ['المرحلة الفرعية', 'المرحلة الفرعيه', 'النوع الفرعي', 'subtype'],
    'التاريخ': ['التاريخ', 'تاريخ', 'تاريخ الحركة', 'date'],
    'تاريخ التسليم': ['تاريخ التسليم', 'الموعد', 'deadline', 'due date', 'تاريخ الاستحقاق', 'تاريخ المهمة'],
    'العنصر': ['العنصر', 'عنصر', 'المهمة', 'task', 'element']
  };

  const searchList = synonyms[headerName] ? synonyms[headerName].map(normalizeString) : [normalizedSearch];

  for (let i = 0; i < headers.length; i++) {
    const normalizedHeader = normalizeString(headers[i]);
    if (searchList.includes(normalizedHeader)) {
      return i + 1;
    }
  }
  return -1;
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
 * الحصول على قائمة المشاريع النشطة - نسخة ذكية
 */
function getActiveProjects() {
  const sheet = getSheet(SHEETS.PROJECTS);
  if (!sheet) return [];

  const lastRow = getLastRowInColumn(sheet, PROJECT_COLS.NAME);
  if (lastRow <= 1) return [];

  // جلب البيانات حتى عمود الحالة
  const maxCol = Math.max(PROJECT_COLS.NAME, PROJECT_COLS.STATUS, PROJECT_COLS.CODE);
  const data = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  const projects = [];

  for (const row of data) {
    const name = row[PROJECT_COLS.NAME - 1];
    const rawStatus = row[PROJECT_COLS.STATUS - 1];
    const status = normalizeString(rawStatus);

    // إضافة المشروع إذا كان الاسم موجوداً والحالة (نشط) أو (فارغة)
    // نستخدم includes('نشط') للتعامل مع الإيقونات مثل ✅ نشط
    if (name && (status === '' || status.includes('نشط') || status.includes('active'))) {
      projects.push({
        code: row[PROJECT_COLS.CODE - 1],
        name: String(name).trim()
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

  const data = sheet.getRange(2, 1, lastRow - 1, PROJECT_COLS.FOLDER_LINK).getValues();
  const projects = [];

  for (const row of data) {
    const name = row[PROJECT_COLS.NAME - 1];
    if (name) {
      projects.push({
        code: row[PROJECT_COLS.CODE - 1],
        name: name,
        type: row[PROJECT_COLS.TYPE - 1] || '',
        channel: row[PROJECT_COLS.CHANNEL - 1] || '',
        producer: row[PROJECT_COLS.PRODUCER - 1] || '',
        status: row[PROJECT_COLS.STATUS - 1] || 'نشط'
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
  if (!sheet) return [];
  
  const nameCol = getColumnByHeader(sheet, 'الاسم');
  const statusCol = getColumnByHeader(sheet, 'الحالة');
  const codeCol = getColumnByHeader(sheet, 'الكود');
  const roleCol = getColumnByHeader(sheet, 'الدور');

  console.log('Fetching Active Team Members. Columns - Name:', nameCol, 'Status:', statusCol);

  if (nameCol === -1 || statusCol === -1) {
    console.error('Core columns (Name/Status) not found in Team sheet');
    return [];
  }

  const lastRow = getLastRowInColumn(sheet, nameCol);
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const members = [];

  for (const row of data) {
    const rawStatus = row[statusCol - 1];
    const status = normalizeString(rawStatus);
    const name = row[nameCol - 1];
    
    // Include if Name is present AND (Status is active OR Status is blank)
    if (name && (status.includes('نشط') || status === 'active' || status === '')) {
      members.push({
        code: codeCol !== -1 ? row[codeCol - 1] : '',
        name: name,
        role: roleCol !== -1 ? row[roleCol - 1] : ''
      });
    }
  }

  console.log('Found Active Members:', members.length);
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
 * الحصول على جميع المراحل الفرعية من شيت الإعدادات
 */
function getAllSubtypes() {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) {
    // استخدام القيم الافتراضية من STAGES
    const subtypes = [];
    for (const key in STAGES) {
      if (STAGES[key].subtypes) {
        subtypes.push(...STAGES[key].subtypes);
      }
    }
    return subtypes;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) {
    // استخدام القيم الافتراضية
    const subtypes = [];
    for (const key in STAGES) {
      if (STAGES[key].subtypes) {
        subtypes.push(...STAGES[key].subtypes);
      }
    }
    return subtypes;
  }

  const data = sheet.getRange(6, 5, lastRow - 5, 2).getValues();
  const subtypes = new Set();

  for (const row of data) {
    if (row[1]) {
      subtypes.add(row[1]);
    }
  }

  return Array.from(subtypes);
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
  if (!sheet) return STAGE_NAMES;

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return STAGE_NAMES;

  const data = sheet.getRange(6, 5, lastRow - 5, 1).getValues();
  const stages = [...new Set(data.map(row => row[0]).filter(Boolean))];

  return stages.length > 0 ? stages : STAGE_NAMES;
}

/**
 * الحصول على أنواع المشاريع من شيت الإعدادات (العمود A)
 */
function getProjectTypesFromSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return PROJECT_TYPES;

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return PROJECT_TYPES;

  const data = sheet.getRange(6, 1, lastRow - 5, 1).getValues();
  const types = data.map(row => row[0]).filter(Boolean);

  return types.length > 0 ? types : PROJECT_TYPES;
}

/**
 * الحصول على أدوار الفريق من شيت الإعدادات (العمود B)
 */
function getTeamRolesFromSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return TEAM_ROLES;

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return TEAM_ROLES;

  const data = sheet.getRange(6, 2, lastRow - 5, 1).getValues();
  const roles = data.map(row => row[0]).filter(Boolean);

  return roles.length > 0 ? roles : TEAM_ROLES;
}

/**
 * الحصول على المدن من شيت الإعدادات (العمود D)
 */
function getCitiesFromSettings() {
  const settingsSheet = getSheet(SHEETS.SETTINGS);
  if (!settingsSheet) return CONFIG.DEFAULT_CITIES;
  
  const lastRow = getLastRowInColumn(settingsSheet, 4); // Column D
  
  // إذا كان العمود فارغاً تماماً (تحت الهيدر) ارجع الافتراضي
  if (lastRow <= 5) return CONFIG.DEFAULT_CITIES;
  
  const values = settingsSheet.getRange(6, 4, lastRow - 5, 1).getValues();
  const cities = values.map(v => v[0]).filter(v => v);
  
  return cities.length > 0 ? cities : CONFIG.DEFAULT_CITIES;
}

/**
 * الحصول على الحالات من شيت الإعدادات (العمود C)
 */
function getStatusesFromSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return PROJECT_STATUS;

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return PROJECT_STATUS;

  const data = sheet.getRange(6, 3, lastRow - 5, 1).getValues();
  const statuses = data.map(row => row[0]).filter(Boolean);

  return statuses.length > 0 ? statuses : PROJECT_STATUS;
}

/**
 * الحصول على الترجمة الإنجليزية للمرحلة
 */
function getStageTranslation(arabicStage) {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return arabicStage;

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return arabicStage;

  const data = sheet.getRange(6, 5, lastRow - 5, 3).getValues(); // E, F, G

  for (const row of data) {
    if (row[0] === arabicStage && row[2]) {
      return row[2]; // العمود G - Stage
    }
  }

  return arabicStage; // إرجاع الاسم العربي إذا لم توجد ترجمة
}

/**
 * الحصول على الترجمة الإنجليزية للمرحلة الفرعية
 */
function getSubtypeTranslation(arabicStage, arabicSubtype) {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return arabicSubtype;

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return arabicSubtype;

  const data = sheet.getRange(6, 5, lastRow - 5, 4).getValues(); // E, F, G, H

  for (const row of data) {
    if (row[0] === arabicStage && row[1] === arabicSubtype && row[3]) {
      return row[3]; // العمود H - Subtype
    }
  }

  return arabicSubtype; // إرجاع الاسم العربي إذا لم توجد ترجمة
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

/**
 * تطبيع جميع التواريخ في النظام
 * يمسح جميع أعمدة التواريخ ويحولها إلى Date Objects بصيغة dd/mm/yyyy
 */
function normalizeAllDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalNormalized = 0;
  
  try {
    // 1. شيت الحركة - التاريخ وتاريخ التسليم
    const movementSheet = getSheet(SHEETS.MOVEMENT);
    if (movementSheet) {
      const dateCol = getColumnByHeader(movementSheet, 'التاريخ');
      const dueDateCol = getColumnByHeader(movementSheet, 'تاريخ التسليم');
      
      if (dateCol !== -1) {
        totalNormalized += normalizeDateColumn_(movementSheet, dateCol);
      }
      if (dueDateCol !== -1) {
        totalNormalized += normalizeDateColumn_(movementSheet, dueDateCol);
      }
    }
    
    // 2. شيت الفريق - تاريخ الانضمام
    const teamSheet = getSheet(SHEETS.TEAM);
    if (teamSheet) {
      const joinDateCol = getColumnByHeader(teamSheet, 'تاريخ الانضمام');
      if (joinDateCol !== -1) {
        totalNormalized += normalizeDateColumn_(teamSheet, joinDateCol);
      }
    }
    
    // 3. شيت المشاريع - تاريخ البدء والانتهاء
    const projectsSheet = getSheet(SHEETS.PROJECTS);
    if (projectsSheet) {
      if (PROJECT_COLS.START_DATE) {
        totalNormalized += normalizeDateColumn_(projectsSheet, PROJECT_COLS.START_DATE);
      }
      if (PROJECT_COLS.END_DATE) {
        totalNormalized += normalizeDateColumn_(projectsSheet, PROJECT_COLS.END_DATE);
      }
    }
    
    showSuccess(`تم تطبيع ${totalNormalized} تاريخ بنجاح ✅`);
  } catch (e) {
    showError('حدث خطأ أثناء تطبيع التواريخ: ' + e.message);
    console.error('Error in normalizeAllDates:', e);
  }
}

/**
 * تطبيع عمود تاريخ كامل
 */
function normalizeDateColumn_(sheet, colNumber) {
  const lastRow = getLastRowInColumn(sheet, colNumber);
  if (lastRow <= 1) return 0;
  
  const range = sheet.getRange(2, colNumber, lastRow - 1, 1);
  const values = range.getValues();
  let normalized = 0;
  
  for (let i = 0; i < values.length; i++) {
    const value = values[i][0];
    if (value) {
      const cell = sheet.getRange(i + 2, colNumber);
      if (normalizeDateCell_(cell, value)) {
        normalized++;
      }
    }
  }
  
  return normalized;
}

/**
 * الحصول على مشروع بواسطة المعرف
 * @param {string} projectId - كود المشروع
 * @returns {Object|null} بيانات المشروع
 */
function getProjectById(projectId) {
  const sheet = getSheet(SHEETS.PROJECTS);
  if (!sheet) return null;
  
  const row = findRowByValue(sheet, PROJECT_COLS.CODE, projectId);
  if (row === -1) return null;
  
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // تحويل الصف إلى كائن
  const project = {};
  for (const key in PROJECT_COLS) {
    project[PROJECT_COLS[key]] = data[PROJECT_COLS[key] - 1];
  }
  
  return project;
}

