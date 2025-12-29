/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة المستخدمين
 * Users Management System
 * ═══════════════════════════════════════════════════════════════════════════════
 * الإصدار: 1.0.0
 * تاريخ الإنشاء: 2025-12-29
 */

// ═══════════════════════════════════════════════════════════════════════════════
// إعدادات نظام المستخدمين
// ═══════════════════════════════════════════════════════════════════════════════

const USER_SHEET_NAME = 'المستخدمين';

const USER_COLS = {
  ID: 1,              // User ID
  EMAIL: 2,           // البريد الإلكتروني
  NAME: 3,            // الاسم
  ROLE: 4,            // الدور (مدير عام / مدير مشروعات)
  PASSWORD: 5,        // كلمة المرور (مشفرة)
  PROJECTS: 6,        // المشاريع المرتبطة (مفصولة بفاصلة)
  ACTIVE: 7,          // نشط (TRUE/FALSE)
  CREATED_DATE: 8,    // تاريخ الإنشاء
  LAST_LOGIN: 9       // آخر تسجيل دخول
};

const USER_HEADERS = [
  'User ID',
  'البريد الإلكتروني',
  'الاسم',
  'الدور',
  'كلمة المرور',
  'المشاريع المرتبطة',
  'نشط',
  'تاريخ الإنشاء',
  'آخر تسجيل دخول'
];

const USER_ROLES = {
  GENERAL_MANAGER: 'مدير عام',
  PROJECT_MANAGER: 'مدير مشروعات'
};

// ═══════════════════════════════════════════════════════════════════════════════
// إنشاء وإعداد شيت المستخدمين
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إنشاء شيت المستخدمين
 */
function setupUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USER_SHEET_NAME);
  
  // إنشاء الشيت إذا لم يكن موجوداً
  if (!sheet) {
    sheet = ss.insertSheet(USER_SHEET_NAME);
    Logger.log('✅ تم إنشاء شيت المستخدمين');
  }
  
  // إعداد الهيدر
  const headerRange = sheet.getRange(1, 1, 1, USER_HEADERS.length);
  headerRange.setValues([USER_HEADERS]);
  headerRange.setBackground(COLORS.HEADER);
  headerRange.setFontColor(COLORS.HEADER_TEXT);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // تجميد الصف الأول
  sheet.setFrozenRows(1);
  
  // تعيين عرض الأعمدة
  sheet.setColumnWidth(USER_COLS.ID, 100);
  sheet.setColumnWidth(USER_COLS.EMAIL, 200);
  sheet.setColumnWidth(USER_COLS.NAME, 150);
  sheet.setColumnWidth(USER_COLS.ROLE, 120);
  sheet.setColumnWidth(USER_COLS.PASSWORD, 150);
  sheet.setColumnWidth(USER_COLS.PROJECTS, 300);
  sheet.setColumnWidth(USER_COLS.ACTIVE, 80);
  sheet.setColumnWidth(USER_COLS.CREATED_DATE, 150);
  sheet.setColumnWidth(USER_COLS.LAST_LOGIN, 150);
  
  Logger.log('✅ تم إعداد شيت المستخدمين بنجاح');
  return sheet;
}

// ═══════════════════════════════════════════════════════════════════════════════
// تشفير كلمات المرور
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تشفير كلمة المرور
 * @param {string} password - كلمة المرور
 * @returns {string} كلمة المرور المشفرة
 */
function hashPassword(password) {
  const salt = 'SeenFilm2025!@#';
  const combined = password + salt;
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    combined,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(hash);
}

/**
 * التحقق من كلمة المرور
 * @param {string} inputPassword - كلمة المرور المدخلة
 * @param {string} storedHash - كلمة المرور المشفرة المخزنة
 * @returns {boolean}
 */
function verifyPassword(inputPassword, storedHash) {
  const inputHash = hashPassword(inputPassword);
  return inputHash === storedHash;
}

// ═══════════════════════════════════════════════════════════════════════════════
// إدارة المستخدمين - CRUD Operations
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة مستخدم جديد
 * @param {Object} userData - بيانات المستخدم
 * @returns {Object} نتيجة العملية
 */
function addUser(userData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USER_SHEET_NAME);
    
    if (!sheet) {
      sheet = setupUsersSheet();
    }
    
    // التحقق من عدم تكرار البريد الإلكتروني
    const existingUser = getUserByEmail(userData.email);
    if (existingUser) {
      return {
        success: false,
        message: '⚠️ هذا البريد الإلكتروني مسجل مسبقاً'
      };
    }
    
    // توليد User ID
    const userId = generateUserId();
    
    // تشفير كلمة المرور
    const hashedPassword = hashPassword(userData.password);
    
    // إعداد البيانات
    const now = new Date();
    const rowData = [
      userId,
      userData.email,
      userData.name,
      userData.role,
      hashedPassword,
      userData.projects || 'ALL', // ALL للمديرين العامين
      true,
      now,
      ''
    ];
    
    // إضافة الصف
    sheet.appendRow(rowData);
    
    Logger.log(`✅ تم إضافة المستخدم: ${userData.name}`);
    
    return {
      success: true,
      message: `✅ تم إضافة المستخدم: ${userData.name}`,
      userId: userId
    };
    
  } catch (error) {
    Logger.log(`❌ خطأ في إضافة المستخدم: ${error}`);
    return {
      success: false,
      message: `❌ خطأ: ${error.message}`
    };
  }
}

/**
 * توليد User ID فريد
 * @returns {string}
 */
function generateUserId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET_NAME);
  
  if (!sheet) return 'U001';
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'U001';
  
  const lastId = sheet.getRange(lastRow, USER_COLS.ID).getValue();
  const numPart = parseInt(lastId.substring(1)) + 1;
  return 'U' + numPart.toString().padStart(3, '0');
}

/**
 * الحصول على مستخدم بالبريد الإلكتروني
 * @param {string} email
 * @returns {Object|null}
 */
function getUserByEmail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET_NAME);
  
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][USER_COLS.EMAIL - 1] === email) {
      return {
        row: i + 1,
        userId: data[i][USER_COLS.ID - 1],
        email: data[i][USER_COLS.EMAIL - 1],
        name: data[i][USER_COLS.NAME - 1],
        role: data[i][USER_COLS.ROLE - 1],
        passwordHash: data[i][USER_COLS.PASSWORD - 1],
        projects: data[i][USER_COLS.PROJECTS - 1],
        active: data[i][USER_COLS.ACTIVE - 1],
        createdDate: data[i][USER_COLS.CREATED_DATE - 1],
        lastLogin: data[i][USER_COLS.LAST_LOGIN - 1]
      };
    }
  }
  
  return null;
}

/**
 * الحصول على جميع المستخدمين النشطين
 * @returns {Array}
 */
function getAllActiveUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET_NAME);
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][USER_COLS.ACTIVE - 1] === true) {
      users.push({
        userId: data[i][USER_COLS.ID - 1],
        email: data[i][USER_COLS.EMAIL - 1],
        name: data[i][USER_COLS.NAME - 1],
        role: data[i][USER_COLS.ROLE - 1],
        projects: data[i][USER_COLS.PROJECTS - 1]
      });
    }
  }
  
  return users;
}

/**
 * تحديث مشاريع مستخدم
 * @param {string} email
 * @param {string} projects - المشاريع مفصولة بفاصلة
 * @returns {Object}
 */
function updateUserProjects(email, projects) {
  try {
    const user = getUserByEmail(email);
    
    if (!user) {
      return {
        success: false,
        message: '⚠️ المستخدم غير موجود'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET_NAME);
    
    sheet.getRange(user.row, USER_COLS.PROJECTS).setValue(projects);
    
    return {
      success: true,
      message: '✅ تم تحديث المشاريع بنجاح'
    };
    
  } catch (error) {
    return {
      success: false,
      message: `❌ خطأ: ${error.message}`
    };
  }
}

/**
 * تعطيل/تفعيل مستخدم
 * @param {string} email
 * @param {boolean} active
 * @returns {Object}
 */
function toggleUserStatus(email, active) {
  try {
    const user = getUserByEmail(email);
    
    if (!user) {
      return {
        success: false,
        message: '⚠️ المستخدم غير موجود'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET_NAME);
    
    sheet.getRange(user.row, USER_COLS.ACTIVE).setValue(active);
    
    const status = active ? 'تفعيل' : 'تعطيل';
    
    return {
      success: true,
      message: `✅ تم ${status} المستخدم بنجاح`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `❌ خطأ: ${error.message}`
    };
  }
}

/**
 * تحديث آخر تسجيل دخول
 * @param {string} email
 */
function updateLastLogin(email) {
  const user = getUserByEmail(email);
  
  if (user) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET_NAME);
    sheet.getRange(user.row, USER_COLS.LAST_LOGIN).setValue(new Date());
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// نظام تسجيل الدخول
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * التحقق من بيانات تسجيل الدخول
 * @param {string} email
 * @param {string} password
 * @returns {Object}
 */
function authenticateUser(email, password) {
  try {
    const user = getUserByEmail(email);
    
    if (!user) {
      return {
        success: false,
        message: '⚠️ البريد الإلكتروني أو كلمة المرور غير صحيحة'
      };
    }
    
    if (!user.active) {
      return {
        success: false,
        message: '⚠️ هذا الحساب غير نشط. يرجى التواصل مع المدير'
      };
    }
    
    if (!verifyPassword(password, user.passwordHash)) {
      return {
        success: false,
        message: '⚠️ البريد الإلكتروني أو كلمة المرور غير صحيحة'
      };
    }
    
    // تحديث آخر تسجيل دخول
    updateLastLogin(email);
    
    return {
      success: true,
      user: {
        userId: user.userId,
        email: user.email,
        name: user.name,
        role: user.role,
        projects: user.projects
      }
    };
    
  } catch (error) {
    Logger.log(`❌ خطأ في تسجيل الدخول: ${error}`);
    return {
      success: false,
      message: `❌ خطأ: ${error.message}`
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// الحصول على مشاريع المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على المشاريع المسموح بها للمستخدم
 * @param {Object} user
 * @returns {Array}
 */
function getUserAllowedProjects(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(SHEETS.PROJECTS);
  
  if (!projectsSheet) return [];
  
  const data = projectsSheet.getDataRange().getValues();
  const projects = [];
  
  // إذا كان مدير عام، يرى جميع المشاريع
  if (user.role === USER_ROLES.GENERAL_MANAGER || user.projects === 'ALL') {
    for (let i = 1; i < data.length; i++) {
      if (data[i][PROJECT_COLS.CODE - 1]) {
        projects.push({
          code: data[i][PROJECT_COLS.CODE - 1],
          name: data[i][PROJECT_COLS.NAME - 1],
          type: data[i][PROJECT_COLS.TYPE - 1],
          status: data[i][PROJECT_COLS.STATUS - 1]
        });
      }
    }
  } else {
    // مدير مشروعات - يرى مشاريعه فقط
    const allowedProjects = user.projects.split(',').map(p => p.trim());
    
    for (let i = 1; i < data.length; i++) {
      const projectCode = data[i][PROJECT_COLS.CODE - 1];
      if (projectCode && allowedProjects.includes(projectCode)) {
        projects.push({
          code: projectCode,
          name: data[i][PROJECT_COLS.NAME - 1],
          type: data[i][PROJECT_COLS.TYPE - 1],
          status: data[i][PROJECT_COLS.STATUS - 1]
        });
      }
    }
  }
  
  return projects;
}

// ═══════════════════════════════════════════════════════════════════════════════
// تحويل المستخدمين القدامى
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * تحويل بيانات المستخدمين القدامى من شيت الإعدادات
 */
function migrateOldUsers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
    
    if (!settingsSheet) {
      Logger.log('⚠️ شيت الإعدادات غير موجود');
      return;
    }
    
    // إنشاء شيت المستخدمين
    const usersSheet = setupUsersSheet();
    
    const data = settingsSheet.getDataRange().getValues();
    let migratedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const email = data[i][0]; // العمود الأول: البريد الإلكتروني
      const name = data[i][1];  // العمود الثاني: الاسم
      const role = data[i][2];  // العمود الثالث: مستوى الصلاحية
      const projects = data[i][3] || ''; // العمود الرابع: المشاريع (إن وجد)
      
      if (email && name) {
        // تحديد الدور
        let userRole = USER_ROLES.PROJECT_MANAGER;
        if (role && role.includes('مدير')) {
          userRole = USER_ROLES.GENERAL_MANAGER;
        }
        
        // إضافة المستخدم
        const result = addUser({
          email: email,
          name: name,
          role: userRole,
          password: 'Seen2025', // كلمة مرور افتراضية - يجب تغييرها
          projects: projects || (userRole === USER_ROLES.GENERAL_MANAGER ? 'ALL' : '')
        });
        
        if (result.success) {
          migratedCount++;
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(
      '✅ تم تحويل المستخدمين',
      `تم تحويل ${migratedCount} مستخدم بنجاح\n\n` +
      'كلمة المرور الافتراضية: Seen2025\n' +
      'يرجى تغيير كلمة المرور لكل مستخدم',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`❌ خطأ في تحويل المستخدمين: ${error}`);
    SpreadsheetApp.getUi().alert('❌ خطأ في تحويل المستخدمين: ' + error.message);
  }
}
