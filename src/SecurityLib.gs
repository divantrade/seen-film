/**
 * مكتبة أمان مركزية لجميع الفحوصات.
 */
const Security = {
  // الحصول على قائمة المدراء العالميين من شيت المستخدمين
  getAdminGroupEmails: function() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('المستخدمين');
      if (!sheet) return [];
      
      const data = sheet.getDataRange().getValues();
      const admins = [];
      
      for (let i = 1; i < data.length; i++) {
        const role = data[i][3]; // عمود الدور D
        const email = data[i][1]; // عمود الإيميل B
        const active = data[i][6]; // عمود النشاط G
        
        if (role === 'مدير عام' && (active === true || String(active).toUpperCase() === 'TRUE')) {
          admins.push(String(email).trim().toLowerCase());
        }
      }
      return admins;
    } catch (e) {
      console.error('Error in getAdminGroupEmails:', e);
      return [];
    }
  },

  isAdmin: function(email) {
    if (!email) {
      email = Session.getEffectiveUser().getEmail();
    }
    const admins = this.getAdminGroupEmails();
    return admins.includes(String(email).trim().toLowerCase());
  },

  isOwner: function(email, projectId) {
    const project = getProjectById(projectId); // نفترض وجود هذه الدالة أو سننشئها
    return project && project[PROJECT_COLS.PRODUCER] === email;
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
  },

  /**
   * التحكم في رؤية التبويبات (Tabs) بناءً على الرتبة
   */
  enforceSheetVisibility: function() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      
      // الحصول على إيميل المستخدم بأكثر من طريقة لضمان النجاح
      let email = "";
      try { email = Session.getActiveUser().getEmail(); } catch(e) {}
      if (!email) {
        try { email = Session.getEffectiveUser().getEmail(); } catch(e) {}
      }
      
      email = (email || "").trim().toLowerCase();
      const ownerEmail = ss.getOwner() ? ss.getOwner().getEmail().trim().toLowerCase() : "";
      
      // إذا فشلنا في جلب الإيميل أو كان صاحب الملف، نفتح كل شيء
      if (email === "" || email === ownerEmail) {
         sheets.forEach(s => {
           try { s.showSheet(); } catch(e) {}
         });
         return;
      }

      const user = getUserByEmail(email);
      const isGeneralManager = user && (String(user.role).trim() === 'مدير عام');
      
      if (isGeneralManager) {
         sheets.forEach(s => {
           try { s.showSheet(); } catch(e) {}
         });
         return;
      }
      
      // لمدير المشروعات: المسموح به فقط
      const allowedNames = ['الحركة', 'داشبورد'];
      sheets.forEach(sheet => {
        try {
          const name = sheet.getName().trim();
          if (allowedNames.includes(name)) {
            sheet.showSheet();
          } else {
            sheet.hideSheet();
          }
        } catch(e) {}
      });

      if (user) {
        this.applyProjectManagerFilter(user);
      }
    } catch (globalError) {
      console.error('Critical failure in enforceSheetVisibility:', globalError);
      // في حالة الفشل الكلي، نحاول إظهار الشيتات لضمان عدم قفل الملف على المستخدم
      try {
        SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(s => s.showSheet());
      } catch(e) {}
    }
  },

  /**
   * تصفية الصفوف ليظهر فقط أفلام الموظف
   */
  applyProjectManagerFilter: function(user) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الحركة');
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    // إظهار الكل أولاً لإعادة الفلترة النظيفة
    sheet.showRows(1, lastRow);

    const rawProjects = user.projects || "";
    // استخراج أسماء الأفلام المسموح بها (بين الأقواس أو النص الكامل)
    const allowedProjects = rawProjects.split(',').map(p => {
       // إذا كان التنسيق [P25001] الكيتاهون، نأخذ "الكيتاهون"
       const match = p.match(/\]\s*(.*)/);
       return match ? match[1].trim() : p.trim();
    }).filter(p => p !== "");

    const data = sheet.getDataRange().getValues();
    const projectCol = getColumnByHeader(sheet, 'الفيلم');
    if (projectCol === -1) return;

    // معالجة الصفوف
    for (let i = 1; i < data.length; i++) {
        const rowProjectName = String(data[i][projectCol - 1]).trim();
        if (rowProjectName === "") continue;

        let isVisible = false;
        for (const allowed of allowedProjects) {
            if (rowProjectName.includes(allowed) || allowed.includes(rowProjectName)) {
                isVisible = true;
                break;
            }
        }

        if (!isVisible) {
            sheet.hideRows(i + 1);
        }
    }
    ss.toast("تم حصر البيانات في مشاريعك فقط 🛡️", "Seen Film Security");
  }
};

/**
 * دالة مساعدة لاستدعاء نظام الرؤية من خارج الكائن
 */
function enforceSheetVisibility() {
  Security.enforceSheetVisibility();
}

/**
 * دالة تشخيص مشكلة إخفاء الشيتات
 * شغلها من القائمة: Run > diagnoseVisibilityIssue
 */
function diagnoseVisibilityIssue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let message = "=== تشخيص نظام الرؤية ===\n\n";

  // 1. إيميل المستخدم الحالي
  let activeEmail = "";
  let effectiveEmail = "";

  try {
    activeEmail = Session.getActiveUser().getEmail();
  } catch(e) {
    activeEmail = "فشل: " + e.message;
  }

  try {
    effectiveEmail = Session.getEffectiveUser().getEmail();
  } catch(e) {
    effectiveEmail = "فشل: " + e.message;
  }

  message += "1. إيميلك (Active): " + (activeEmail || "فارغ") + "\n";
  message += "2. إيميلك (Effective): " + (effectiveEmail || "فارغ") + "\n\n";

  // 2. صاحب الملف
  let ownerEmail = "";
  try {
    const owner = ss.getOwner();
    ownerEmail = owner ? owner.getEmail() : "لا يوجد owner";
  } catch(e) {
    ownerEmail = "فشل: " + e.message;
  }

  message += "3. صاحب الملف: " + ownerEmail + "\n\n";

  // 3. المقارنة
  const normalizedActive = (activeEmail || "").trim().toLowerCase();
  const normalizedOwner = (ownerEmail || "").trim().toLowerCase();
  const isOwner = normalizedActive === normalizedOwner && normalizedActive !== "";

  message += "4. هل أنت صاحب الملف؟ " + (isOwner ? "نعم ✅" : "لا ❌") + "\n";
  message += "   - إيميلك: [" + normalizedActive + "]\n";
  message += "   - صاحب الملف: [" + normalizedOwner + "]\n\n";

  // 4. التحقق من شيت المستخدمين
  let user = null;
  try {
    user = getUserByEmail(normalizedActive);
  } catch(e) {
    message += "5. خطأ في getUserByEmail: " + e.message + "\n";
  }

  message += "5. هل أنت في شيت المستخدمين؟ " + (user ? "نعم ✅" : "لا ❌") + "\n";
  if (user) {
    message += "   - الدور: " + user.role + "\n";
    message += "   - نشط: " + user.active + "\n";
  }

  // 5. قائمة المدراء
  let admins = [];
  try {
    admins = Security.getAdminGroupEmails();
  } catch(e) {}
  message += "\n6. المدراء المسجلين: " + (admins.length > 0 ? admins.join(", ") : "لا يوجد") + "\n";
  message += "   - هل أنت مدير؟ " + (admins.includes(normalizedActive) ? "نعم ✅" : "لا ❌") + "\n";

  ui.alert("تشخيص نظام الرؤية", message, ui.ButtonSet.OK);
}
