/**
 * مكتبة أمان مركزية لجميع الفحوصات.
 */
const Security = {
  // استيراد مجموعة المديرين من Google Groups (مؤقتاً من شيت الإعدادات)
  getAdminGroupEmails: function() {
    // TODO: استبدال هذا الاستدعاء ب AdminDirectory.Groups.list عندما يتم تفعيل API.
    return getAdminsList(); // من Permissions.gs (قائمة مؤقتة)
  },

  isAdmin: function(email) {
    const admins = this.getAdminGroupEmails();
    return admins.includes(email);
  },

  isOwner: function(email, projectId) {
    const project = getProjectById(projectId); // نفترض وجود هذه الدالة أو سننشئها
    return project && project[PROJECT_COLS.OWNER_EMAIL] === email;
  },

  requirePermission: function(email, requiredRole, projectId) {
    if (requiredRole === 'ADMIN' && this.isAdmin(email)) return true;
    if (requiredRole === 'OWNER' && this.isOwner(email, projectId)) return true;
    return false;
  },

  /**
   * ربط الصلاحية بالعملية.
   * @param {string} operation اسم العملية (مثال: 'تعديل مشروع')
   * @param {string} requiredRole ROLE المطلوبة ('ADMIN' أو 'OWNER')
   * @param {string} [projectId] معرف المشروع إذا كان الدور OWNER.
   */
  enforce: function(operation, requiredRole, projectId) {
    const email = getCurrentUserEmail();
    if (!this.requirePermission(email, requiredRole, projectId)) {
      // محاولة تسجيل العملية في سجل التدقيق إذا كانت الدالة متوفرة
      try {
        if (typeof logAuditEntry === 'function') {
          logAuditEntry({
            action: 'محاولة مرفوضة',
            sheetName: operation,
            details: 'المستخدم: ' + email + (projectId ? ' - المشروع: ' + projectId : '')
          });
        }
      } catch (e) {
        console.error('فشل تسجيل التدقيق:', e);
      }

      SpreadsheetApp.getUi().alert('ليس لديك صلاحية للقيام بـ ' + operation);
      return false;
    }
    return true;
  }
};
