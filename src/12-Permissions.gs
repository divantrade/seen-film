/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * نظام الصلاحيات الكامل
 * ═══════════════════════════════════════════════════════════════════════════════
 * الإصدار: 2.0.0
 * يشمل:
 *   - ثلاث مستويات: مدير، منتج، موظف
 *   - ربط المستخدم بشيت الفريق عبر الإيميل
 *   - تحديد المسؤول عن كل مشروع
 *   - حماية عمليات الحذف والتعديل
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// الإعدادات
// ═══════════════════════════════════════════════════════════════════════════════

const PERMISSIONS_CONFIG = {
  // موقع قائمة المدراء في شيت الإعدادات
  ADMINS_COLUMN: 10,        // العمود J
  ADMINS_HEADER_ROW: 5,     // صف الهيدر
  ADMINS_START_ROW: 6,      // بداية البيانات

  // مستويات الصلاحيات
  LEVELS: {
    ADMIN: 'مدير',
    PRODUCER: 'منتج',
    EMPLOYEE: 'موظف'
  },

  // الأدوار التي تُعتبر "منتج" (صلاحيات أعلى)
  PRODUCER_ROLES: ['منتج', 'مخرج', 'مدير إنتاج'],

  // رسائل الخطأ
  MESSAGES: {
    NO_PERMISSION: '⛔ ليس لديك صلاحية لتنفيذ هذه العملية.',
    ADMIN_ONLY: '🔒 هذه العملية متاحة للمدراء فقط.',
    PRODUCER_ONLY: '🔒 هذه العملية متاحة للمنتجين فقط.',
    NOT_YOUR_PROJECT: '⚠️ هذا المشروع ليس من مشاريعك.',
    NOT_YOUR_TASK: '⚠️ هذه المهمة ليست مسندة إليك.',
    CONTACT_ADMIN: 'تواصل مع أحد المدراء للمساعدة.',
    NO_ADMINS_CONFIGURED: '⚠️ لم يتم تحديد أي مدراء بعد.\n\nاذهب إلى: الإعدادات ← إضافة مدير'
  }
};

// ═══════════════════════════════════════════════════════════════════════════════
// دوال التحقق من الصلاحيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على إيميل المستخدم الحالي
 * @returns {string} إيميل المستخدم
 */
function getCurrentUserEmail() {
  try {
    // محاولة الحصول على الإيميل بعدة طرق
    let email = Session.getActiveUser().getEmail();

    if (!email) {
      email = Session.getEffectiveUser().getEmail();
    }

    return email ? email.toLowerCase().trim() : '';
  } catch (e) {
    console.error('خطأ في الحصول على إيميل المستخدم:', e);
    return '';
  }
}

/**
 * الحصول على قائمة المدراء من شيت الإعدادات
 * @returns {Array<string>} قائمة إيميلات المدراء
 */
function getAdminsList() {
  try {
    const sheet = getSheet(SHEETS.SETTINGS);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < PERMISSIONS_CONFIG.ADMINS_START_ROW) return [];

    const numRows = lastRow - PERMISSIONS_CONFIG.ADMINS_START_ROW + 1;
    const data = sheet.getRange(
      PERMISSIONS_CONFIG.ADMINS_START_ROW,
      PERMISSIONS_CONFIG.ADMINS_COLUMN,
      numRows,
      1
    ).getValues();

    // تنظيف وفلترة الإيميلات
    const admins = data
      .map(row => row[0] ? row[0].toString().toLowerCase().trim() : '')
      .filter(email => email && email.includes('@'));

    return admins;
  } catch (e) {
    console.error('خطأ في قراءة قائمة المدراء:', e);
    return [];
  }
}

/**
 * التحقق إذا كان المستخدم الحالي مدير
 * @returns {boolean}
 */
function isCurrentUserAdmin() {
  const currentEmail = getCurrentUserEmail();

  // إذا لم نستطع تحديد المستخدم، نسمح (لتجنب مشاكل التطوير)
  if (!currentEmail) {
    console.log('تحذير: لم يتم تحديد إيميل المستخدم');
    return true;
  }

  const admins = getAdminsList();

  // إذا لم يتم تحديد أي مدراء، السماح للجميع (النظام جديد)
  if (admins.length === 0) {
    return true;
  }

  return admins.includes(currentEmail);
}

/**
 * التحقق إذا كان إيميل معين مدير
 * @param {string} email - الإيميل للتحقق
 * @returns {boolean}
 */
function isAdmin(email) {
  if (!email) return false;

  const admins = getAdminsList();

  // إذا لم يتم تحديد أي مدراء، السماح للجميع
  if (admins.length === 0) {
    return true;
  }

  return admins.includes(email.toLowerCase().trim());
}

/**
 * طلب صلاحية المدير وعرض رسالة خطأ إذا لم تكن متوفرة
 * @param {string} operationName - اسم العملية
 * @returns {boolean} true إذا كان مدير، false إذا لا
 */
function requireAdmin(operationName) {
  if (isCurrentUserAdmin()) {
    return true;
  }

  // عرض رسالة رفض الصلاحية
  const admins = getAdminsList();
  let message = PERMISSIONS_CONFIG.MESSAGES.ADMIN_ONLY + '\n\n';
  message += 'العملية المطلوبة: ' + operationName + '\n\n';

  if (admins.length > 0) {
    message += 'المدراء المعتمدون:\n';
    message += admins.map(a => '• ' + a).join('\n');
  } else {
    message += PERMISSIONS_CONFIG.MESSAGES.CONTACT_ADMIN;
  }

  SpreadsheetApp.getUi().alert('🔒 صلاحية مرفوضة', message, SpreadsheetApp.getUi().ButtonSet.OK);

  // تسجيل المحاولة في سجل التدقيق
  try {
    logAuditEntry({
      action: 'محاولة مرفوضة',
      sheetName: operationName,
      details: 'المستخدم: ' + getCurrentUserEmail()
    });
  } catch (e) {
    console.error('خطأ في تسجيل المحاولة:', e);
  }

  return false;
}

// ═══════════════════════════════════════════════════════════════════════════════
// إدارة المدراء
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة مدير جديد
 */
function addAdmin() {
  const ui = SpreadsheetApp.getUi();

  // التحقق من الصلاحية (أول مدير يمكن لأي شخص إضافته)
  const admins = getAdminsList();
  if (admins.length > 0 && !isCurrentUserAdmin()) {
    requireAdmin('إضافة مدير');
    return;
  }

  // طلب الإيميل
  const result = ui.prompt(
    '➕ إضافة مدير جديد',
    'أدخل البريد الإلكتروني للمدير الجديد:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const email = result.getResponseText().trim().toLowerCase();

  // التحقق من صحة الإيميل
  if (!email || !email.includes('@')) {
    showError('⚠️ البريد الإلكتروني غير صالح.');
    return;
  }

  // التحقق من عدم التكرار
  if (admins.includes(email)) {
    showInfo('ℹ️ هذا البريد مسجل كمدير بالفعل.');
    return;
  }

  // إضافة المدير
  try {
    const sheet = getSheet(SHEETS.SETTINGS);

    // التأكد من وجود الهيدر
    const headerCell = sheet.getRange(PERMISSIONS_CONFIG.ADMINS_HEADER_ROW, PERMISSIONS_CONFIG.ADMINS_COLUMN);
    if (!headerCell.getValue()) {
      headerCell.setValue('المدراء')
        .setBackground(COLORS.INFO)
        .setFontWeight('bold');
    }

    // إيجاد أول صف فارغ
    const lastRow = sheet.getLastRow();
    let targetRow = PERMISSIONS_CONFIG.ADMINS_START_ROW;

    for (let i = PERMISSIONS_CONFIG.ADMINS_START_ROW; i <= lastRow + 1; i++) {
      const cellValue = sheet.getRange(i, PERMISSIONS_CONFIG.ADMINS_COLUMN).getValue();
      if (!cellValue) {
        targetRow = i;
        break;
      }
      targetRow = i + 1;
    }

    // إضافة الإيميل
    sheet.getRange(targetRow, PERMISSIONS_CONFIG.ADMINS_COLUMN).setValue(email);

    // تسجيل في سجل التدقيق
    logAuditEntry({
      action: 'إضافة مدير',
      sheetName: 'الصلاحيات',
      newValue: email
    });

    showSuccess('✅ تم إضافة المدير: ' + email);

  } catch (e) {
    console.error('خطأ في إضافة المدير:', e);
    showError('❌ حدث خطأ أثناء إضافة المدير.');
  }
}

/**
 * عرض قائمة المدراء الحاليين
 */
function showAdminsList() {
  const admins = getAdminsList();
  const ui = SpreadsheetApp.getUi();

  if (admins.length === 0) {
    ui.alert(
      '📋 قائمة المدراء',
      'لم يتم تحديد أي مدراء بعد.\n\n' +
      'اضغط على "إضافة مدير" لإضافة أول مدير.',
      ui.ButtonSet.OK
    );
    return;
  }

  const currentUser = getCurrentUserEmail();
  let message = 'المدراء المعتمدون:\n\n';

  admins.forEach((admin, index) => {
    const isYou = admin === currentUser ? ' (أنت)' : '';
    message += (index + 1) + '. ' + admin + isYou + '\n';
  });

  message += '\n───────────────────\n';
  message += 'المستخدم الحالي: ' + (currentUser || 'غير محدد');

  ui.alert('📋 قائمة المدراء', message, ui.ButtonSet.OK);
}

/**
 * إزالة مدير
 */
function removeAdmin() {
  // التحقق من الصلاحية
  if (!requireAdmin('إزالة مدير')) {
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const admins = getAdminsList();

  if (admins.length === 0) {
    showInfo('لا يوجد مدراء لإزالتهم.');
    return;
  }

  if (admins.length === 1) {
    showError('⚠️ لا يمكن إزالة المدير الوحيد.\nأضف مديراً آخر أولاً.');
    return;
  }

  // طلب الإيميل
  const result = ui.prompt(
    '➖ إزالة مدير',
    'المدراء الحاليون:\n' + admins.map((a, i) => (i+1) + '. ' + a).join('\n') +
    '\n\nأدخل البريد الإلكتروني للمدير المراد إزالته:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const email = result.getResponseText().trim().toLowerCase();

  if (!admins.includes(email)) {
    showError('⚠️ هذا البريد غير مسجل كمدير.');
    return;
  }

  // منع إزالة النفس إذا كان المدير الأخير
  const currentUser = getCurrentUserEmail();
  if (email === currentUser && admins.length <= 1) {
    showError('⚠️ لا يمكنك إزالة نفسك كمدير وحيد.');
    return;
  }

  // إزالة المدير
  try {
    const sheet = getSheet(SHEETS.SETTINGS);
    const lastRow = sheet.getLastRow();

    for (let i = PERMISSIONS_CONFIG.ADMINS_START_ROW; i <= lastRow; i++) {
      const cell = sheet.getRange(i, PERMISSIONS_CONFIG.ADMINS_COLUMN);
      if (cell.getValue().toString().toLowerCase().trim() === email) {
        cell.clearContent();
        break;
      }
    }

    // تسجيل في سجل التدقيق
    logAuditEntry({
      action: 'إزالة مدير',
      sheetName: 'الصلاحيات',
      oldValue: email
    });

    showSuccess('✅ تم إزالة المدير: ' + email);

  } catch (e) {
    console.error('خطأ في إزالة المدير:', e);
    showError('❌ حدث خطأ أثناء إزالة المدير.');
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// العمليات المحمية (تحتاج صلاحية مدير)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إعادة تهيئة النظام (محمية)
 * تستبدل الدالة الأصلية
 */
function resetSystemProtected() {
  // التحقق من الصلاحية
  if (!requireAdmin('إعادة تهيئة النظام')) {
    return;
  }

  // التحقق والنسخ الاحتياطي
  if (!backupBeforeDangerousOperation('إعادة تهيئة النظام')) {
    return;
  }

  // تسجيل العملية
  logAuditEntry({
    action: 'إعادة تهيئة',
    sheetName: 'النظام',
    details: 'بواسطة: ' + getCurrentUserEmail()
  });

  // تنفيذ العملية الأصلية
  resetSystem();
}

/**
 * حذف مشروع (محمية)
 */
function deleteProjectProtected() {
  // التحقق من الصلاحية
  if (!requireAdmin('حذف مشروع')) {
    return;
  }

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
    showError('لا يوجد مشروع في هذا الصف');
    return;
  }

  const ui = SpreadsheetApp.getUi();

  // تأكيد الحذف
  const result = ui.alert(
    '⚠️ تأكيد الحذف',
    'هل أنت متأكد من حذف المشروع:\n\n' +
    '📁 ' + projectName + ' (' + projectCode + ')\n\n' +
    '⚠️ سيتم حذف الصف فقط، الفولدرات على Drive ستبقى.',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  // تسجيل الحذف
  logAuditEntry({
    action: 'حذف مشروع',
    sheetName: SHEETS.PROJECTS,
    oldValue: projectCode + ' - ' + projectName,
    details: 'بواسطة: ' + getCurrentUserEmail()
  });

  // حذف الصف
  sheet.deleteRow(row);

  // تحديث القوائم المنسدلة
  updateMovementDropdowns();

  showSuccess('✅ تم حذف المشروع: ' + projectName);
}

/**
 * حذف عضو فريق (محمية)
 */
function deleteTeamMemberProtected() {
  // التحقق من الصلاحية
  if (!requireAdmin('حذف عضو فريق')) {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== SHEETS.TEAM) {
    showError('يجب أن تكون في شيت الفريق');
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    showError('اختر عضواً من القائمة');
    return;
  }

  const memberName = sheet.getRange(row, TEAM_COLS.NAME).getValue();
  const memberCode = sheet.getRange(row, TEAM_COLS.CODE).getValue();

  if (!memberName) {
    showError('لا يوجد عضو في هذا الصف');
    return;
  }

  const ui = SpreadsheetApp.getUi();

  // تأكيد الحذف
  const result = ui.alert(
    '⚠️ تأكيد الحذف',
    'هل أنت متأكد من حذف العضو:\n\n' +
    '👤 ' + memberName + ' (' + memberCode + ')',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  // تسجيل الحذف
  logAuditEntry({
    action: 'حذف عضو فريق',
    sheetName: SHEETS.TEAM,
    oldValue: memberCode + ' - ' + memberName,
    details: 'بواسطة: ' + getCurrentUserEmail()
  });

  // حذف الصف
  sheet.deleteRow(row);

  // تحديث القوائم المنسدلة
  updateMovementDropdowns();

  showSuccess('✅ تم حذف العضو: ' + memberName);
}

/**
 * حذف صفوف متعددة (محمية)
 */
function deleteSelectedRowsProtected() {
  // التحقق من الصلاحية
  if (!requireAdmin('حذف صفوف')) {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  // التحقق من الشيت
  const protectedSheets = [SHEETS.PROJECTS, SHEETS.TEAM, SHEETS.MOVEMENT];
  if (!protectedSheets.includes(sheetName)) {
    showError('هذا الشيت غير مدعوم للحذف المحمي');
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow <= 1) {
    showError('لا يمكن حذف صف الهيدر');
    return;
  }

  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    '⚠️ تأكيد الحذف',
    'هل أنت متأكد من حذف ' + numRows + ' صف/صفوف؟',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  // تسجيل الحذف
  logAuditEntry({
    action: 'حذف صفوف',
    sheetName: sheetName,
    oldValue: 'الصفوف ' + startRow + ' - ' + (startRow + numRows - 1),
    details: 'بواسطة: ' + getCurrentUserEmail()
  });

  // حذف الصفوف
  sheet.deleteRows(startRow, numRows);

  showSuccess('✅ تم حذف ' + numRows + ' صف/صفوف');
}

// ═══════════════════════════════════════════════════════════════════════════════
// ربط المستخدم بشيت الفريق
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على بيانات المستخدم الحالي من شيت الفريق
 * @returns {Object|null} بيانات العضو أو null إذا لم يوجد
 */
function getCurrentUserInfo() {
  const email = getCurrentUserEmail();
  if (!email) return null;

  try {
    const sheet = getSheet(SHEETS.TEAM);
    if (!sheet) return null;

    const lastRow = getLastRowInColumn(sheet, TEAM_COLS.NAME);
    if (lastRow < 2) return null;

    // البحث في شيت الفريق عن الإيميل
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

    for (const row of data) {
      const memberEmail = row[TEAM_COLS.EMAIL - 1]; // EMAIL column
      if (memberEmail && memberEmail.toString().toLowerCase().trim() === email) {
        return {
          code: row[TEAM_COLS.CODE - 1],
          name: row[TEAM_COLS.NAME - 1],
          role: row[TEAM_COLS.ROLE - 1],
          status: row[TEAM_COLS.STATUS - 1],
          email: memberEmail
        };
      }
    }

    return null;
  } catch (e) {
    console.error('خطأ في البحث عن المستخدم:', e);
    return null;
  }
}

/**
 * تحديد مستوى صلاحية المستخدم الحالي
 * @returns {string} المستوى: 'مدير', 'منتج', أو 'موظف'
 */
function getCurrentUserLevel() {
  // أولاً: تحقق إذا كان مدير
  if (isCurrentUserAdmin()) {
    return PERMISSIONS_CONFIG.LEVELS.ADMIN;
  }

  // ثانياً: تحقق من دوره في شيت الفريق
  const userInfo = getCurrentUserInfo();

  if (!userInfo) {
    // إذا لم يكن مسجلاً في الفريق، يُعتبر موظف
    return PERMISSIONS_CONFIG.LEVELS.EMPLOYEE;
  }

  // تحقق إذا كان دوره من أدوار المنتجين
  const role = userInfo.role ? userInfo.role.toString().trim() : '';
  const isProducerRole = PERMISSIONS_CONFIG.PRODUCER_ROLES.some(
    r => normalizeString(role).includes(normalizeString(r))
  );

  if (isProducerRole) {
    return PERMISSIONS_CONFIG.LEVELS.PRODUCER;
  }

  return PERMISSIONS_CONFIG.LEVELS.EMPLOYEE;
}

/**
 * الحصول على المشاريع التي يكون فيها المستخدم الحالي هو المنتج المسؤول
 * @returns {Array<Object>} قائمة المشاريع
 */
function getUserProjects() {
  const userInfo = getCurrentUserInfo();
  if (!userInfo) return [];

  const userName = userInfo.name;
  if (!userName) return [];

  try {
    const sheet = getSheet(SHEETS.PROJECTS);
    if (!sheet) return [];

    const lastRow = getLastRowInColumn(sheet, PROJECT_COLS.NAME);
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const projects = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const projectName = row[PROJECT_COLS.NAME - 1];
      const producer = row[PROJECT_COLS.PRODUCER - 1];

      // المنتج المسؤول (عمود المنتج)
      if (projectName && producer) {
        const normalizedProducer = normalizeString(producer);
        const normalizedUser = normalizeString(userName);

        if (normalizedProducer.includes(normalizedUser) || normalizedUser.includes(normalizedProducer)) {
          projects.push({
            row: i + 2,
            code: row[PROJECT_COLS.CODE - 1],
            name: projectName,
            type: row[PROJECT_COLS.TYPE - 1],
            status: row[PROJECT_COLS.STATUS - 1]
          });
        }
      }
    }

    return projects;
  } catch (e) {
    console.error('خطأ في جلب مشاريع المستخدم:', e);
    return [];
  }
}

/**
 * التحقق إذا كان المستخدم مسؤولاً عن مشروع معين
 * @param {string} projectName - اسم المشروع
 * @returns {boolean}
 */
function isUserResponsibleForProject(projectName) {
  if (!projectName) return false;

  // المدراء لهم صلاحية على كل المشاريع
  if (isCurrentUserAdmin()) return true;

  const userProjects = getUserProjects();
  const normalizedTarget = normalizeString(projectName);

  return userProjects.some(p =>
    normalizeString(p.name) === normalizedTarget
  );
}

/**
 * التحقق إذا كان المستخدم يمكنه تعديل مشروع معين
 * @param {string} projectName - اسم المشروع
 * @returns {boolean}
 */
function canEditProject(projectName) {
  // المدراء يستطيعون تعديل كل شيء
  if (isCurrentUserAdmin()) return true;

  // المنتجون يستطيعون تعديل مشاريعهم فقط
  const level = getCurrentUserLevel();
  if (level === PERMISSIONS_CONFIG.LEVELS.PRODUCER) {
    return isUserResponsibleForProject(projectName);
  }

  // الموظفون لا يستطيعون تعديل المشاريع
  return false;
}

/**
 * التحقق إذا كان المستخدم يمكنه تعديل مهمة معينة
 * @param {string} projectName - اسم المشروع
 * @param {string} assignedTo - المسؤول عن المهمة
 * @returns {boolean}
 */
function canEditTask(projectName, assignedTo) {
  // المدراء يستطيعون تعديل كل شيء
  if (isCurrentUserAdmin()) return true;

  // المنتجون يستطيعون تعديل مهام مشاريعهم
  const level = getCurrentUserLevel();
  if (level === PERMISSIONS_CONFIG.LEVELS.PRODUCER) {
    return isUserResponsibleForProject(projectName);
  }

  // الموظفون يستطيعون تعديل المهام المسندة إليهم فقط
  const userInfo = getCurrentUserInfo();
  if (!userInfo || !assignedTo) return false;

  const normalizedAssigned = normalizeString(assignedTo);
  const normalizedUser = normalizeString(userInfo.name);

  return normalizedAssigned.includes(normalizedUser) || normalizedUser.includes(normalizedAssigned);
}

/**
 * طلب صلاحية تعديل مشروع مع عرض رسالة خطأ إذا لم تكن متوفرة
 * @param {string} projectName - اسم المشروع
 * @param {string} operationName - اسم العملية
 * @returns {boolean}
 */
function requireProjectPermission(projectName, operationName) {
  if (canEditProject(projectName)) {
    return true;
  }

  const level = getCurrentUserLevel();
  let message = '';

  if (level === PERMISSIONS_CONFIG.LEVELS.EMPLOYEE) {
    message = PERMISSIONS_CONFIG.MESSAGES.PRODUCER_ONLY + '\n\n';
    message += 'المشروع: ' + projectName + '\n';
    message += 'العملية: ' + operationName + '\n\n';
    message += 'تواصل مع المنتج المسؤول عن هذا المشروع.';
  } else {
    message = PERMISSIONS_CONFIG.MESSAGES.NOT_YOUR_PROJECT + '\n\n';
    message += 'المشروع: ' + projectName + '\n\n';
    message += 'هذا المشروع ليس ضمن مشاريعك المسندة إليك.';
  }

  SpreadsheetApp.getUi().alert('⚠️ صلاحية غير متوفرة', message, SpreadsheetApp.getUi().ButtonSet.OK);

  // تسجيل المحاولة
  try {
    logAuditEntry({
      action: 'محاولة مرفوضة',
      sheetName: operationName,
      details: 'مشروع: ' + projectName + ' | المستخدم: ' + getCurrentUserEmail()
    });
  } catch (e) {
    console.error('خطأ في تسجيل المحاولة:', e);
  }

  return false;
}

/**
 * طلب صلاحية تعديل مهمة مع عرض رسالة خطأ إذا لم تكن متوفرة
 * @param {string} projectName - اسم المشروع
 * @param {string} assignedTo - المسؤول عن المهمة
 * @param {string} operationName - اسم العملية
 * @returns {boolean}
 */
function requireTaskPermission(projectName, assignedTo, operationName) {
  if (canEditTask(projectName, assignedTo)) {
    return true;
  }

  let message = PERMISSIONS_CONFIG.MESSAGES.NOT_YOUR_TASK + '\n\n';
  message += 'المشروع: ' + projectName + '\n';
  message += 'المسؤول: ' + (assignedTo || 'غير محدد') + '\n';
  message += 'العملية: ' + operationName + '\n\n';
  message += 'يمكنك فقط تعديل المهام المسندة إليك.';

  SpreadsheetApp.getUi().alert('⚠️ صلاحية غير متوفرة', message, SpreadsheetApp.getUi().ButtonSet.OK);

  // تسجيل المحاولة
  try {
    logAuditEntry({
      action: 'محاولة مرفوضة',
      sheetName: operationName,
      details: 'مهمة في: ' + projectName + ' | المستخدم: ' + getCurrentUserEmail()
    });
  } catch (e) {
    console.error('خطأ في تسجيل المحاولة:', e);
  }

  return false;
}

// ═══════════════════════════════════════════════════════════════════════════════
// التحقق من حالة الصلاحيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * عرض حالة الصلاحيات للمستخدم الحالي
 */
function showMyPermissions() {
  const ui = SpreadsheetApp.getUi();
  const currentEmail = getCurrentUserEmail();
  const userInfo = getCurrentUserInfo();
  const level = getCurrentUserLevel();
  const userProjects = getUserProjects();
  const admins = getAdminsList();

  // رمز المستوى
  let levelIcon = '👤';
  if (level === PERMISSIONS_CONFIG.LEVELS.ADMIN) levelIcon = '👑';
  else if (level === PERMISSIONS_CONFIG.LEVELS.PRODUCER) levelIcon = '🎬';

  let message = '═══════════════════════════════════\n';
  message += '👤 معلومات المستخدم الحالي\n';
  message += '═══════════════════════════════════\n\n';

  message += '📧 البريد: ' + (currentEmail || 'غير محدد') + '\n';

  if (userInfo) {
    message += '👤 الاسم: ' + userInfo.name + '\n';
    message += '💼 الدور: ' + (userInfo.role || 'غير محدد') + '\n';
  }

  message += levelIcon + ' المستوى: ' + level + '\n\n';

  message += '───────────────────────────────────\n';
  message += 'ما يمكنك فعله:\n\n';

  if (level === PERMISSIONS_CONFIG.LEVELS.ADMIN) {
    message += '✅ إضافة/تعديل/حذف المشاريع\n';
    message += '✅ إضافة/تعديل/حذف أعضاء الفريق\n';
    message += '✅ إضافة/تعديل كل المهام\n';
    message += '✅ إعادة تهيئة النظام\n';
    message += '✅ إدارة المدراء\n';
    message += '✅ عرض سجل التغييرات\n';
  } else if (level === PERMISSIONS_CONFIG.LEVELS.PRODUCER) {
    message += '✅ تعديل مشاريعك\n';
    message += '✅ إضافة/تعديل مهام مشاريعك\n';
    message += '✅ عرض التقارير\n';
    message += '❌ حذف المشاريع\n';
    message += '❌ حذف أعضاء الفريق\n';
    message += '❌ إعادة تهيئة النظام\n';
  } else {
    message += '✅ تعديل المهام المسندة إليك\n';
    message += '✅ تحديث حالة مهامك\n';
    message += '✅ عرض التقارير\n';
    message += '❌ تعديل المشاريع\n';
    message += '❌ حذف أي شيء\n';
  }

  // عرض المشاريع المسندة (للمنتجين)
  if (level === PERMISSIONS_CONFIG.LEVELS.PRODUCER && userProjects.length > 0) {
    message += '\n───────────────────────────────────\n';
    message += '🎬 مشاريعك (' + userProjects.length + '):\n\n';
    userProjects.forEach((p, i) => {
      message += (i + 1) + '. ' + p.name + ' (' + p.code + ')\n';
    });
  }

  message += '\n───────────────────────────────────\n';
  message += 'عدد المدراء المسجلين: ' + admins.length;

  ui.alert('🔐 صلاحياتي', message, ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// حماية التعديل في الـ onEdit
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * التحقق من صلاحية تعديل صف في شيت المشاريع
 * تُستدعى من onProjectEdit
 * @param {Object} e - حدث التعديل
 * @returns {boolean} true إذا مسموح بالتعديل
 */
function validateProjectEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEETS.PROJECTS) return true;

    const row = e.range.getRow();
    if (row <= 1) return true; // الهيدر

    // المدراء يستطيعون تعديل كل شيء
    if (isCurrentUserAdmin()) return true;

    // الحصول على اسم المشروع
    const projectName = sheet.getRange(row, PROJECT_COLS.NAME).getValue();
    if (!projectName) return true; // صف فارغ

    // التحقق من الصلاحية
    if (!canEditProject(projectName)) {
      // إرجاع القيمة القديمة
      const oldValue = e.oldValue;
      if (oldValue !== undefined) {
        e.range.setValue(oldValue);
      }

      // عرض تنبيه
      SpreadsheetApp.getActiveSpreadsheet().toast(
        PERMISSIONS_CONFIG.MESSAGES.NOT_YOUR_PROJECT,
        '⚠️ صلاحية مرفوضة',
        5
      );

      return false;
    }

    return true;
  } catch (error) {
    console.error('خطأ في validateProjectEdit:', error);
    return true; // السماح في حالة الخطأ
  }
}

/**
 * التحقق من صلاحية تعديل صف في شيت الحركة (المهام)
 * تُستدعى من onMovementEdit
 * @param {Object} e - حدث التعديل
 * @returns {boolean} true إذا مسموح بالتعديل
 */
function validateMovementEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEETS.MOVEMENT) return true;

    const row = e.range.getRow();
    if (row <= 1) return true; // الهيدر

    // المدراء يستطيعون تعديل كل شيء
    if (isCurrentUserAdmin()) return true;

    // الحصول على بيانات المهمة
    const projectName = sheet.getRange(row, MOVEMENT_COLS.PROJECT).getValue();
    const assignedTo = sheet.getRange(row, MOVEMENT_COLS.ASSIGNED_TO).getValue();

    if (!projectName) return true; // صف فارغ

    // التحقق من الصلاحية
    if (!canEditTask(projectName, assignedTo)) {
      // إرجاع القيمة القديمة
      const oldValue = e.oldValue;
      if (oldValue !== undefined) {
        e.range.setValue(oldValue);
      }

      // عرض تنبيه
      SpreadsheetApp.getActiveSpreadsheet().toast(
        PERMISSIONS_CONFIG.MESSAGES.NOT_YOUR_TASK,
        '⚠️ صلاحية مرفوضة',
        5
      );

      return false;
    }

    return true;
  } catch (error) {
    console.error('خطأ في validateMovementEdit:', error);
    return true; // السماح في حالة الخطأ
  }
}

/**
 * التحقق من صلاحية تعديل صف في شيت الفريق
 * تُستدعى من onTeamEdit
 * @param {Object} e - حدث التعديل
 * @returns {boolean} true إذا مسموح بالتعديل
 */
function validateTeamEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEETS.TEAM) return true;

    const row = e.range.getRow();
    if (row <= 1) return true; // الهيدر

    // فقط المدراء يستطيعون تعديل شيت الفريق
    if (!isCurrentUserAdmin()) {
      // إرجاع القيمة القديمة
      const oldValue = e.oldValue;
      if (oldValue !== undefined) {
        e.range.setValue(oldValue);
      }

      // عرض تنبيه
      SpreadsheetApp.getActiveSpreadsheet().toast(
        PERMISSIONS_CONFIG.MESSAGES.ADMIN_ONLY,
        '⚠️ صلاحية مرفوضة',
        5
      );

      return false;
    }

    return true;
  } catch (error) {
    console.error('خطأ في validateTeamEdit:', error);
    return true;
  }
}

/**
 * التحقق الموحد من صلاحيات التعديل
 * يُستدعى من onEdit الرئيسية
 * @param {Object} e - حدث التعديل
 * @returns {boolean}
 */
function validateEditPermission(e) {
  const sheetName = e.source.getActiveSheet().getName();

  switch (sheetName) {
    case SHEETS.PROJECTS:
      return validateProjectEdit(e);
    case SHEETS.MOVEMENT:
      return validateMovementEdit(e);
    case SHEETS.TEAM:
      return validateTeamEdit(e);
    default:
      return true;
  }
}
