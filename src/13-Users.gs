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

  // --- إضافة القوائم المنسدلة (Data Validation) ---
  
  // 1. قائمة الأدوار (مدير عام / مدير مشروعات)
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([USER_ROLES.GENERAL_MANAGER, USER_ROLES.PROJECT_MANAGER], true)
    .setAllowInvalid(false)
    .setHelpText('يرجى اختيار الدور من القائمة')
    .build();
  sheet.getRange(2, USER_COLS.ROLE, 1000).setDataValidation(roleRule);
  
  // 2. قائمة الحالة (TRUE / FALSE)
  const activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, USER_COLS.ACTIVE, 1000).setDataValidation(activeRule);
  
  // 3. قائمة المشاريع (تُجلب من شيت المشاريع)
  const projectsSheet = ss.getSheetByName(SHEETS.PROJECTS);
  if (projectsSheet) {
    const lastRow = Math.max(projectsSheet.getLastRow(), 2);
    const rawData = projectsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    // تنسيق العرض: [الكود] اسم الفيلم
    const dropdownList = ['ALL'];
    rawData.forEach(row => {
      if (row[0] && row[1]) {
        dropdownList.push(`[${row[0]}] ${row[1]}`);
      } else if (row[0]) {
        dropdownList.push(row[0]);
      }
    });
    
    const projectRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdownList, true)
      .setAllowInvalid(true) // السماح بإدخال قيم متعددة دون إظهار خطأ صارم
      .setHelpText('اختر المشاريع لإضافتها أو حذفها.')
      .build();
    sheet.getRange(2, USER_COLS.PROJECTS, 1000).setDataValidation(projectRule);
  }
  
  // 4. ملاحظة على عمود المشاريع
  sheet.getRange(1, USER_COLS.PROJECTS).setNote('💡 نظام الاختيار المتعدد الذكي:\n- اختر [كود] اسم لإضافته.\n- اختره ثانيةً لحذفه.\n- سيقوم النظام تلقائياً بتحويل الاختيار لأكواد فقط لكي تعمل في الـ Web App.');
  
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
  const salt = CONFIG.PASSWORD_SALT;
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
    const rowEmail = String(data[i][USER_COLS.EMAIL - 1]).trim().toLowerCase();
    if (rowEmail === email.trim().toLowerCase()) {
      const rawName = data[i][USER_COLS.NAME - 1];
      console.log(`[getUserByEmail] Found user. Email: ${rowEmail}, RawName: "${rawName}"`);
      
      return {
        row: i + 1,
        userId: data[i][USER_COLS.ID - 1],
        email: data[i][USER_COLS.EMAIL - 1],
        name: rawName, // Keep original value to debug
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
    if (!email || !password) {
      return { success: false, message: '⚠️ يرجى إدخل البريد الإلكتروني وكلمة المرور' };
    }
    
    // تنظيف المدخلات
    const cleanEmail = String(email).trim().toLowerCase();
    const user = getUserByEmail(cleanEmail);
    
    if (!user) {
      return {
        success: false,
        message: '⚠️ البريد الإلكتروني غير مسجل في النظام'
      };
    }
    
    // التحقق من الحالة النشطة (سواء كانت Boolean أو String)
    const isActive = String(user.active).toUpperCase() === 'TRUE' || user.active === true;
    if (!isActive) {
      return {
        success: false,
        message: '⚠️ هذا الحساب معطل (نشط = FALSE)'
      };
    }
    
    // التحقق من كلمة المرور مع تنظيفها من أي مسافات بالخطأ
    if (!verifyPassword(String(password).trim(), user.passwordHash)) {
      return {
        success: false,
        message: '⚠️ كلمة المرور التي أدخلتها غير صحيحة'
      };
    }
    
    // تحديث آخر تسجيل دخول (pass row directly to avoid re-scan)
    updateLastLogin(cleanEmail, user.row);
    
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

/**
 * تحديث آخر تسجيل دخول
 * @param {string} email
 * @param {number} rowNumber - Optional row number to skip search
 */
function updateLastLogin(email, rowNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET_NAME);
    
    let targetRow = rowNumber;
    
    // Fallback if row not provided
    if (!targetRow) {
      const user = getUserByEmail(email);
      if (user) targetRow = user.row;
    }
    
    if (targetRow) {
      // Use formatted string for clearer logging/reading
      const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(targetRow, USER_COLS.LAST_LOGIN).setValue(now);
    }
  } catch (e) {
    console.error('Failed to update last login:', e);
    // Don't fail the login just because this fails
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
    // مدير مشروعات - استخراج الأكواد من التنسيق المليء بالأسماء [P25001] الاسم
    let projectStrings = user.projects.split(',').map(p => p.trim());
    let allowedCodes = projectStrings.map(str => {
      const match = str.match(/\[(.*?)\]/);
      return match ? match[1] : str; // إذا وجد قوسين يأخذ ما بينهما، وإلا يأخذ النص كما هو
    });
    
    for (let i = 1; i < data.length; i++) {
      const projectCode = String(data[i][PROJECT_COLS.CODE - 1]).trim();
      if (projectCode && allowedCodes.includes(projectCode)) {
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
 * تحويل بيانات المستخدمين القدامى من شيت الصلاحيات
 */
function migrateOldUsers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    let permSheet = null;
    
    // البحث عن الشيت الذي يحتوي اسمه على كلمة "الصلاحيات"
    for (const s of allSheets) {
      if (s.getName().indexOf('الصلاحيات') > -1) {
        permSheet = s;
        break;
      }
    }
    
    if (!permSheet) {
      SpreadsheetApp.getUi().alert('⚠️ لم يتم العثور على شيت يحتوي على كلمة "الصلاحيات"');
      return;
    }
    
    // إنشاء شيت المستخدمين إذا لم يكن موجوداً
    setupUsersSheet();
    
    const data = permSheet.getDataRange().getValues();
    let migratedCount = 0;
    
    // البحث عن نقطة البداية (تخطي أي هيدرات حتى نجد البريد الإلكتروني)
    let startIndex = 1;
    for (let i = 0; i < data.length; i++) {
        const cell = String(data[i][0]).trim();
        if (cell === 'البريد الإلكتروني' || cell === 'إيميل') {
            startIndex = i + 1;
            break;
        }
    }
    
    for (let i = startIndex; i < data.length; i++) {
      const email = String(data[i][0]).toLowerCase().trim(); // العمود الأول: البريد الإلكتروني
      const name = data[i][1];  // العمود الثاني: الاسم
      const roleText = String(data[i][2]);  // العمود الثالث: مستوى الصلاحية
      const projects = String(data[i][3] || ''); // العمود الرابع: كود المشروع
      
      // التأكد أن هذا صف ببيانات حقيقية (بريد إلكتروني يحتوي على @)
      if (email && email.indexOf('@') > -1) {
        // تحديد الدور
        let userRole = USER_ROLES.PROJECT_MANAGER;
        if (roleText.includes('مدير') || roleText.includes('نظام')) {
          userRole = USER_ROLES.GENERAL_MANAGER;
        }
        
        // إضافة المستخدم
        const result = addUser({
          email: email,
          name: name || email.split('@')[0],
          role: userRole,
          password: CONFIG.DEFAULT_PASSWORD,
          projects: (userRole === USER_ROLES.GENERAL_MANAGER) ? 'ALL' : projects
        });
        
        if (result.success) {
          migratedCount++;
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(
      '✅ تم تحويل المستخدمين',
      `تم تحويل ${migratedCount} مستخدم من شيت الصلاحيات بنجاح\n\n` +
      'كلمة المرور الافتراضية للجميع: Seen2025\n' +
      'يرجى إبلاغ الفريق بتغيير كلمات سرهم قريباً.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`❌ خطأ في تحويل المستخدمين: ${error}`);
    SpreadsheetApp.getUi().alert('❌ خطأ في تحويل المستخدمين: ' + error.message);
  }
}
/**
 * دالة للمسؤول لتعيين كلمة مرور جديدة لأي مستخدم
 */
function adminChangeUserPassword() {
  const ui = SpreadsheetApp.getUi();
  
  // التحقق من أن المستخدم الحالي مدير عام
  const currentUserEmail = Session.getEffectiveUser().getEmail();
  if (!Security.isAdmin(currentUserEmail)) {
    ui.alert('❌ ليس لديك صلاحية لتغيير كلمات مرور المستخدمين. هذه الميزة للمدير العام فقط.');
    return;
  }
  
  // طلب البريد الإلكتروني للمستخدم
  const emailResult = ui.prompt(
    '🔑 إعادة تعيين كلمة مرور',
    'أدخل البريد الإلكتروني للمستخدم المراد تغيير كلمة مروره:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResult.getSelectedButton() !== ui.Button.OK) return;
  const targetEmail = emailResult.getResponseText().trim().toLowerCase();
  
  const user = getUserByEmail(targetEmail);
  if (!user) {
    ui.alert('❌ هذا المستخدم غير مسجل في شيت المستخدمين.');
    return;
  }
  
  // طلب كلمة المرور الجديدة
  const passResult = ui.prompt(
    '🔑 كلمة المرور الجديدة',
    `أدخل كلمة المرور الجديدة للمستخدم: ${user.name}`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (passResult.getSelectedButton() !== ui.Button.OK) return;
  const newPassword = passResult.getResponseText().trim();
  
  if (newPassword.length < 4) {
    ui.alert('⚠️ كلمة المرور قصيرة جداً، يرجى إدخال 4 أحرف على الأقل.');
    return;
  }
  
  // تنفيذ التغيير
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET_NAME);
    const hashedPassword = hashPassword(newPassword);
    
    sheet.getRange(user.row, USER_COLS.PASSWORD).setValue(hashedPassword);
    
    ui.alert('✅ تم تغيير كلمة المرور بنجاح للمستخدم: ' + user.name);
    
  } catch (e) {
    ui.alert('❌ حدث خطأ أثناء تغيير كلمة المرور: ' + e.message);
  }
}

/**
 * إظهار كافة الشيتات للمدير العام (حل طوارئ)
 */
function showAllSheetsForAdmin() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentUserEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim();
  const ownerEmail = ss.getOwner().getEmail().toLowerCase().trim();
  
  // التحقق من أن المستخدم مدير عام (بالإيميل أو بالدور)
  if (Security.isAdmin(currentUserEmail) || currentUserEmail === ownerEmail) {
    const sheets = ss.getSheets();
    sheets.forEach(sheet => sheet.showSheet());
    ui.alert('✅ تم إظهار كافة الشيتات بنجاح يا مدير.');
  } else {
    ui.alert('❌ عذراً، هذه الخاصية للمدير العام فقط.');
  }
}
/**
 * حذف شيت الصلاحيات القديم نهائياً
 */
function deleteOldPermissionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🗑️ حذف الشيت القديم',
    'هل أنت متأكد من حذف شيت "الصلاحيات" نهائياً؟ \n (سيتم الاعتماد كلياً على شيت المستخدمين الجديد)',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const sheetNames = ['🛡️ الصلاحيات', 'الصلاحيات'];
  let sheetFound = null;
  
  for (const name of sheetNames) {
    const s = ss.getSheetByName(name);
    if (s) {
      sheetFound = s;
      break;
    }
  }
  
  if (sheetFound) {
    ss.deleteSheet(sheetFound);
    ui.alert('✅ تم حذف شيت الصلاحيات بنجاح.');
  } else {
    ui.alert('ℹ️ لم يتم العثور على شيت باسم "الصلاحيات" أو "🛡️ الصلاحيات".');
  }
}
