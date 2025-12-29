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
  const page = e.parameter.page || 'login';
  let template;
  let title = 'Seen Film - نظام إدارة الإنتاج';
  
  switch (page) {
    case 'general-manager':
      template = HtmlService.createTemplateFromFile('webapp/GeneralManagerDashboard');
      title = 'Seen Film - لوحة التحكم - مدير عام';
      break;
      
    case 'project-manager':
      template = HtmlService.createTemplateFromFile('webapp/ProjectManagerDashboard');
      title = 'Seen Film - لوحة التحكم - مدير مشروعات';
      break;
      
    case 'login':
    default:
      template = HtmlService.createTemplateFromFile('webapp/Login');
      title = 'Seen Film - تسجيل الدخول';
      break;
  }
  
  return template.evaluate()
    .setTitle(title)
    .setFaviconUrl('https://www.seenfilm.com/favicon.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * تضمين ملفات HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
function webAppGetGeneralManagerData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // إحصائيات المشاريع
    const projectsSheet = ss.getSheetByName(SHEETS.PROJECTS);
    const projectsData = projectsSheet ? projectsSheet.getDataRange().getValues() : [];
    
    const totalProjects = projectsData.length - 1;
    const activeProjects = projectsData.filter((row, i) => i > 0 && row[PROJECT_COLS.STATUS - 1] === 'نشط').length;
    
    // إحصائيات الحركة
    const movementSheet = ss.getSheetByName(SHEETS.MOVEMENT);
    const movementData = movementSheet ? movementSheet.getDataRange().getValues() : [];
    
    const totalTasks = movementData.length - 1;
    const completedTasks = movementData.filter((row, i) => i > 0 && row[MOVEMENT_COLS.STATUS - 1] === 'تم').length;
    const inProgressTasks = movementData.filter((row, i) => i > 0 && row[MOVEMENT_COLS.STATUS - 1] === 'جاري').length;
    
    // إحصائيات الفريق
    const teamSheet = ss.getSheetByName(SHEETS.TEAM);
    const teamData = teamSheet ? teamSheet.getDataRange().getValues() : [];
    const totalTeam = teamData.length - 1;
    
    // المشاريع الأخيرة
    const recentProjects = [];
    for (let i = 1; i < Math.min(6, projectsData.length); i++) {
      recentProjects.push({
        code: projectsData[i][PROJECT_COLS.CODE - 1],
        name: projectsData[i][PROJECT_COLS.NAME - 1],
        status: projectsData[i][PROJECT_COLS.STATUS - 1],
        type: projectsData[i][PROJECT_COLS.TYPE - 1]
      });
    }
    
    // المهام المتأخرة
    const delayedTasks = [];
    const today = new Date();
    for (let i = 1; i < movementData.length; i++) {
      const dueDate = movementData[i][MOVEMENT_COLS.DUE_DATE - 1];
      const status = movementData[i][MOVEMENT_COLS.STATUS - 1];
      
      if (dueDate && status !== 'تم' && new Date(dueDate) < today) {
        delayedTasks.push({
          project: movementData[i][MOVEMENT_COLS.PROJECT - 1],
          stage: movementData[i][MOVEMENT_COLS.STAGE - 1],
          details: movementData[i][MOVEMENT_COLS.DETAILS - 1],
          dueDate: Utilities.formatDate(new Date(dueDate), CONFIG.TIMEZONE, CONFIG.DATE_FORMAT),
          assignedTo: movementData[i][MOVEMENT_COLS.ASSIGNED_TO - 1]
        });
        
        if (delayedTasks.length >= 10) break;
      }
    }
    
    return {
      success: true,
      stats: {
        totalProjects: totalProjects,
        activeProjects: activeProjects,
        totalTasks: totalTasks,
        completedTasks: completedTasks,
        inProgressTasks: inProgressTasks,
        totalTeam: totalTeam
      },
      recentProjects: recentProjects,
      delayedTasks: delayedTasks
    };
    
  } catch (error) {
    Logger.log(`Error in webAppGetGeneralManagerData: ${error}`);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * الحصول على بيانات Dashboard لمدير المشروعات
 */
function webAppGetProjectManagerData(userEmail) {
  try {
    const user = getUserByEmail(userEmail);
    if (!user) {
      return { success: false, message: 'المستخدم غير موجود' };
    }
    
    const allowedProjects = getUserAllowedProjects(user);
    const projectCodes = allowedProjects.map(p => p.code);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // المهام الخاصة بمشاريع المستخدم
    const movementSheet = ss.getSheetByName(SHEETS.MOVEMENT);
    const movementData = movementSheet ? movementSheet.getDataRange().getValues() : [];
    
    const userTasks = [];
    const completedCount = {};
    const inProgressCount = {};
    const delayedTasks = [];
    const today = new Date();
    
    for (let i = 1; i < movementData.length; i++) {
      const projectCode = movementData[i][MOVEMENT_COLS.PROJECT - 1];
      
      if (projectCodes.includes(projectCode)) {
        const status = movementData[i][MOVEMENT_COLS.STATUS - 1];
        const dueDate = movementData[i][MOVEMENT_COLS.DUE_DATE - 1];
        
        // عد المهام حسب الحالة
        if (!completedCount[projectCode]) completedCount[projectCode] = 0;
        if (!inProgressCount[projectCode]) inProgressCount[projectCode] = 0;
        
        if (status === 'تم') completedCount[projectCode]++;
        if (status === 'جاري') inProgressCount[projectCode]++;
        
        // المهام المتأخرة
        if (dueDate && status !== 'تم' && new Date(dueDate) < today) {
          delayedTasks.push({
            project: projectCode,
            projectName: allowedProjects.find(p => p.code === projectCode)?.name || projectCode,
            stage: movementData[i][MOVEMENT_COLS.STAGE - 1],
            details: movementData[i][MOVEMENT_COLS.DETAILS - 1],
            dueDate: Utilities.formatDate(new Date(dueDate), CONFIG.TIMEZONE, CONFIG.DATE_FORMAT),
            assignedTo: movementData[i][MOVEMENT_COLS.ASSIGNED_TO - 1]
          });
        }
        
        // آخر 20 مهمة
        if (userTasks.length < 20) {
          userTasks.push({
            number: movementData[i][MOVEMENT_COLS.NUMBER - 1],
            date: movementData[i][MOVEMENT_COLS.DATE - 1],
            project: projectCode,
            stage: movementData[i][MOVEMENT_COLS.STAGE - 1],
            subtype: movementData[i][MOVEMENT_COLS.SUBTYPE - 1],
            details: movementData[i][MOVEMENT_COLS.DETAILS - 1],
            assignedTo: movementData[i][MOVEMENT_COLS.ASSIGNED_TO - 1],
            status: status
          });
        }
      }
    }
    
    // إحصائيات لكل مشروع
    const projectStats = allowedProjects.map(project => ({
      code: project.code,
      name: project.name,
      type: project.type,
      status: project.status,
      completedTasks: completedCount[project.code] || 0,
      inProgressTasks: inProgressCount[project.code] || 0
    }));
    
    return {
      success: true,
      projects: allowedProjects,
      projectStats: projectStats,
      recentTasks: userTasks,
      delayedTasks: delayedTasks
    };
    
  } catch (error) {
    Logger.log(`Error in webAppGetProjectManagerData: ${error}`);
    return {
      success: false,
      message: error.message
    };
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
