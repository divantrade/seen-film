/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * نظام الصلاحيات - الدفعة الأولى (الأساسيات)
 * ═══════════════════════════════════════════════════════════════════════════════
 * الإصدار: 1.0.0
 * يشمل:
 *   - تحديد المدراء
 *   - حماية عمليات الحذف
 *   - حماية إعادة التهيئة
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

  // رسائل الخطأ
  MESSAGES: {
    NO_PERMISSION: '⛔ ليس لديك صلاحية لتنفيذ هذه العملية.',
    ADMIN_ONLY: '🔒 هذه العملية متاحة للمدراء فقط.',
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
// التحقق من حالة الصلاحيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * عرض حالة الصلاحيات للمستخدم الحالي
 */
function showMyPermissions() {
  const ui = SpreadsheetApp.getUi();
  const currentEmail = getCurrentUserEmail();
  const isAdmin = isCurrentUserAdmin();
  const admins = getAdminsList();

  let message = '═══════════════════════════════════\n';
  message += '👤 معلومات المستخدم الحالي\n';
  message += '═══════════════════════════════════\n\n';

  message += '📧 البريد: ' + (currentEmail || 'غير محدد') + '\n';
  message += '🔑 الصلاحية: ' + (isAdmin ? '✅ مدير' : '👤 مستخدم عادي') + '\n\n';

  message += '───────────────────────────────────\n';
  message += 'ما يمكنك فعله:\n\n';

  if (isAdmin) {
    message += '✅ إضافة/تعديل/حذف المشاريع\n';
    message += '✅ إضافة/تعديل/حذف أعضاء الفريق\n';
    message += '✅ إضافة/تعديل المهام\n';
    message += '✅ إعادة تهيئة النظام\n';
    message += '✅ إدارة المدراء\n';
    message += '✅ عرض سجل التغييرات\n';
  } else {
    message += '✅ إضافة/تعديل المهام\n';
    message += '✅ تحديث حالة المهام\n';
    message += '✅ عرض التقارير\n';
    message += '❌ حذف المشاريع\n';
    message += '❌ حذف أعضاء الفريق\n';
    message += '❌ إعادة تهيئة النظام\n';
  }

  message += '\n───────────────────────────────────\n';
  message += 'عدد المدراء المسجلين: ' + admins.length;

  ui.alert('🔐 صلاحياتي', message, ui.ButtonSet.OK);
}
