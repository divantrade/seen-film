/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * Web App - Backend
 * نظام الويب أب - الخادم
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * فتح Web App مع دعم التوجيه
 */
function doGet(e) {
  // دائماً نبدأ بصفحة الدخول كحاوية رئيسية
  try {
    const template = HtmlService.createTemplateFromFile('Login');
    return template.evaluate()
      .setTitle('Seen Film - نظام إدارة الإنتاج')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput('<h2>⚠️ خطأ في تحميل النظام:</h2><p>' + err.message + '</p>');
  }
}

/**
 * Simple ping for connectivity test
 */
function webAppPing() {
  return "Pong - Server is Reachable";
}

/**
 * دالة لجلب محتوى أي صفحة (View) بدون إعادة تحميل المتصفح
 */
function webAppGetView(viewName) {
  try {
    // معالجة القالب برمجياً لتنفيذ جميع أوامر include والبيانات المدمجة
    const template = HtmlService.createTemplateFromFile(viewName);
    return template.evaluate().getContent();
  } catch (e) {
    return '<h2>❌ تعذر تحميل الواجهة: ' + viewName + '</h2><p>' + e.message + '</p>';
  }
}

/**
 * تضمين ملفات HTML
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return `<!-- Failed to include: ${filename} -->`;
  }
}

/**
 * معالجة تسجيل الدخول من Web App
 */
function webAppLogin(email, password) {
  return authenticateUser(email, password);
}

/**
 * الحصول على مشاريع المستخدم من Web App
 */
function webAppGetUserProjects(userEmail) {
  const user = getUserByEmail(userEmail);
  if (!user) {
    return [];
  }
  return getUserAllowedProjects(user);
}

/**
 * الحصول على بيانات Dashboard للمدير العام
 */
/**
 * الحصول على بيانات Dashboard للمدير العام
 */
/**
 * الحصول على بيانات Dashboard للمدير العام
 */
function webAppGetGeneralManagerData() {
  // ISOLATION MODE: Return immediate mock data purely to test UI connection speed.
  // This bypasses ALL Sheet reading operations.
  return {
    success: true,
    stats: {
      totalProjects: 0,
      activeProjects: 0,
      totalTasks: 0,
      totalTeam: 0 // Mock Data
    },
    recentProjects: [],
    delayedTasks: []
  };
}

/**
 * الحصول على بيانات Dashboard لمدير المشروعات
 */
function webAppGetProjectManagerData(userEmail) {
  try {
    const user = getUserByEmail(userEmail);
    if (!user) return { success: false, message: 'المستخدم غير موجود' };
    
    const allowedProjects = getUserAllowedProjects(user);
    const projectNames = allowedProjects.map(p => normalizeString(p.name));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const movementSheet = ss.getSheetByName(SHEETS.MOVEMENT);
    if (!movementSheet) return { success: true, projects: allowedProjects, projectStats: [], recentTasks: [], delayedTasks: [] };
    
    const mLastRow = movementSheet.getLastRow();
    const movementData = (mLastRow > 1) ? movementSheet.getRange(1, 1, mLastRow, movementSheet.getLastColumn()).getValues() : [];
    
    const userTasks = [];
    const completedCount = {};
    const inProgressCount = {};
    const delayedTasks = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    for (let i = 1; i < movementData.length; i++) {
      const projectNameInSheet = movementData[i][MOVEMENT_COLS.PROJECT - 1];
      if (!projectNameInSheet) continue;
      
      const normProjectName = normalizeString(projectNameInSheet);
      
      if (projectNames.includes(normProjectName)) {
        const project = allowedProjects.find(p => normalizeString(p.name) === normProjectName);
        const projectCode = project ? project.code : projectNameInSheet;
        const status = normalizeString(movementData[i][MOVEMENT_COLS.STATUS - 1]);
        const dueDateRaw = movementData[i][MOVEMENT_COLS.DUE_DATE - 1];
        
        if (!completedCount[projectCode]) completedCount[projectCode] = 0;
        if (!inProgressCount[projectCode]) inProgressCount[projectCode] = 0;
        
        if (status.includes('تم')) completedCount[projectCode]++;
        else if (status.includes('جاري')) inProgressCount[projectCode]++;
        
        if (dueDateRaw && !status.includes('تم') && !status.includes('ملغي')) {
          const dueDate = new Date(dueDateRaw);
          if (dueDate < today) {
            delayedTasks.push({
              project: projectCode,
              projectName: projectNameInSheet,
              stage: movementData[i][MOVEMENT_COLS.STAGE - 1],
              details: movementData[i][MOVEMENT_COLS.DETAILS - 1],
              dueDate: Utilities.formatDate(dueDate, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT),
              assignedTo: movementData[i][MOVEMENT_COLS.ASSIGNED_TO - 1]
            });
          }
        }
        
        if (!status.includes('تم') && !status.includes('ملغي') && userTasks.length < 20) {
          userTasks.push({
            number: movementData[i][MOVEMENT_COLS.NUMBER - 1],
            date: movementData[i][MOVEMENT_COLS.DATE - 1],
            project: projectNameInSheet,
            stage: movementData[i][MOVEMENT_COLS.STAGE - 1],
            subtype: movementData[i][MOVEMENT_COLS.SUBTYPE - 1],
            details: movementData[i][MOVEMENT_COLS.DETAILS - 1],
            assignedTo: movementData[i][MOVEMENT_COLS.ASSIGNED_TO - 1],
            status: movementData[i][MOVEMENT_COLS.STATUS - 1]
          });
        }
      }
    }
    
    const projectStats = allowedProjects.map(project => ({
      code: project.code,
      name: project.name,
      type: project.type,
      status: project.status,
      completedTasks: completedCount[project.code] || 0,
      inProgressTasks: inProgressCount[project.code] || 0
    }));
    
    return sanitizeForClient({
      success: true,
      projects: allowedProjects,
      projectStats: projectStats,
      recentTasks: userTasks,
      delayedTasks: delayedTasks
    });
    
  } catch (error) {
    console.error(`Error in webAppGetProjectManagerData: ${error.message}`);
    return sanitizeForClient({ success: false, message: error.message });
  }
}

/**
 * إدارة المستخدمين - للمديرين العامين فقط
 */
function webAppManageUsers(action, userData, requesterEmail) {
  try {
    // التحقق من أن الطالب هو مدير عام
    const requester = getUserByEmail(requesterEmail);
    if (!requester || requester.role !== USER_ROLES.GENERAL_MANAGER) {
      return {
        success: false,
        message: '⚠️ ليس لديك صلاحية لإدارة المستخدمين'
      };
    }
    
    switch (action) {
      case 'list':
        return {
          success: true,
          users: getAllActiveUsers()
        };
        
      case 'add':
        return addUser(userData);
        
      case 'updateProjects':
        return updateUserProjects(userData.email, userData.projects);
        
      case 'toggle':
        return toggleUserStatus(userData.email, userData.active);
        
      default:
        return {
          success: false,
          message: 'عملية غير معروفة'
        };
    }
    
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * الحصول على قائمة جميع المشاريع (للمديرين العامين)
 */
function webAppGetAllProjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PROJECTS);
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const projects = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][PROJECT_COLS.CODE - 1]) {
      projects.push({
        code: data[i][PROJECT_COLS.CODE - 1],
        name: data[i][PROJECT_COLS.NAME - 1]
      });
    }
  }
  
  return projects;
}
