/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * إدارة الفولدرات التلقائية على Google Drive
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * الحصول على فولدر الإنتاج الرئيسي من الإعدادات
 */
function getMainProductionFolder() {
  const sheet = getSheet(SHEETS.SETTINGS);
  if (!sheet) return null;

  const folderUrl = sheet.getRange('B3').getValue();
  if (!folderUrl || folderUrl === '(أدخل رابط الفولدر هنا)') {
    return null;
  }

  try {
    // استخراج ID الفولدر من الرابط
    const folderId = extractFolderIdFromUrl(folderUrl);
    if (!folderId) return null;

    return DriveApp.getFolderById(folderId);
  } catch (error) {
    console.error('خطأ في الوصول للفولدر الرئيسي:', error);
    return null;
  }
}

/**
 * استخراج ID الفولدر من رابط Google Drive
 */
function extractFolderIdFromUrl(url) {
  if (!url) return null;

  // تنظيف الرابط
  url = url.toString().trim();

  // إذا كان ID مباشرة (بدون /)
  if (!url.includes('/') && !url.includes('?')) {
    return url;
  }

  // استخراج من رابط
  const patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[-\w]{25,}/  // أي ID طويل
  ];

  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match) {
      return match[1] || match[0];
    }
  }

  return null;
}

/**
 * اختبار الفولدر الرئيسي
 */
function testMainFolder() {
  const sheet = getSheet(SHEETS.SETTINGS);
  const folderUrl = sheet.getRange('B3').getValue();

  console.log('الرابط المُدخل:', folderUrl);

  const folderId = extractFolderIdFromUrl(folderUrl);
  console.log('ID المستخرج:', folderId);

  if (!folderId) {
    showError('لم يتم استخراج ID الفولدر. تأكد من صحة الرابط.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    showSuccess('تم الاتصال بالفولدر: ' + folder.getName());
  } catch (e) {
    showError('خطأ في الوصول للفولدر: ' + e.message);
    console.error(e);
  }
}

/**
 * إنشاء هيكل فولدرات لمشروع جديد
 */
function createProjectFolderStructure(projectName, projectCode) {
  const mainFolder = getMainProductionFolder();

  if (!mainFolder) {
    showError('يرجى تحديد فولدر الإنتاج الرئيسي في الإعدادات أولاً');
    return null;
  }

  try {
    // إنشاء فولدر المشروع الرئيسي
    const projectFolderName = projectCode + ' - ' + projectName;
    let projectFolder = findFolderByName(mainFolder, projectFolderName);

    if (!projectFolder) {
      projectFolder = mainFolder.createFolder(projectFolderName);
    }

    // إنشاء الفولدرات الفرعية
    for (const subfolderName of FOLDER_STRUCTURE) {
      if (!findFolderByName(projectFolder, subfolderName)) {
        projectFolder.createFolder(subfolderName);
      }
    }

    return projectFolder.getUrl();

  } catch (error) {
    console.error('خطأ في إنشاء فولدرات المشروع:', error);
    showError('حدث خطأ أثناء إنشاء الفولدرات: ' + error.message);
    return null;
  }
}

/**
 * إنشاء فولدر تصوير لمدينة معينة
 */
function createShootingFolder(projectName, cityName, movementRow) {
  const mainFolder = getMainProductionFolder();

  if (!mainFolder) {
    console.log('فولدر الإنتاج الرئيسي غير محدد');
    return null;
  }

  try {
    // البحث عن فولدر المشروع
    const projectsSheet = getSheet(SHEETS.PROJECTS);
    const projectRow = findRowByValue(projectsSheet, PROJECT_COLS.NAME, projectName);

    if (projectRow === -1) {
      console.log('المشروع غير موجود');
      return null;
    }

    const projectCode = projectsSheet.getRange(projectRow, PROJECT_COLS.CODE).getValue();
    const projectFolderName = projectCode + ' - ' + projectName;

    let projectFolder = findFolderByName(mainFolder, projectFolderName);

    if (!projectFolder) {
      // إنشاء فولدر المشروع إذا لم يكن موجوداً
      const folderUrl = createProjectFolderStructure(projectName, projectCode);
      if (!folderUrl) return null;

      projectFolder = findFolderByName(mainFolder, projectFolderName);
    }

    // البحث عن فولدر التصوير
    const shootingFolder = findFolderByName(projectFolder, '03-التصوير');
    if (!shootingFolder) {
      console.log('فولدر التصوير غير موجود');
      return null;
    }

    // إنشاء فولدر المدينة
    const cityFolderName = 'تصوير ' + cityName;
    let cityFolder = findFolderByName(shootingFolder, cityFolderName);

    if (!cityFolder) {
      cityFolder = shootingFolder.createFolder(cityFolderName);
    }

    const folderUrl = cityFolder.getUrl();

    // تحديث رابط الفولدر في شيت الحركة
    if (movementRow) {
      const movementSheet = getSheet(SHEETS.MOVEMENT);
      movementSheet.getRange(movementRow, MOVEMENT_COLS.LINK).setValue(folderUrl);
    }

    return folderUrl;

  } catch (error) {
    console.error('خطأ في إنشاء فولدر التصوير:', error);
    return null;
  }
}

/**
 * البحث عن فولدر بالاسم داخل فولدر معين
 */
function findFolderByName(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return null;
}

/**
 * إنشاء فولدر لعنصر في الحركة
 */
function createFolderForMovement() {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== SHEETS.MOVEMENT) {
    showError('يجب أن تكون في شيت الحركة');
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    showError('اختر حركة من القائمة');
    return;
  }

  const project = sheet.getRange(row, MOVEMENT_COLS.PROJECT).getValue();
  const stage = sheet.getRange(row, MOVEMENT_COLS.STAGE).getValue();
  const subtype = sheet.getRange(row, MOVEMENT_COLS.SUBTYPE).getValue();
  const element = sheet.getRange(row, MOVEMENT_COLS.ELEMENT).getValue();
  const existingLink = sheet.getRange(row, MOVEMENT_COLS.LINK).getValue();

  if (existingLink) {
    showInfo('يوجد رابط فولدر بالفعل');
    return;
  }

  if (!project) {
    showError('يجب تحديد المشروع أولاً');
    return;
  }

  let folderUrl = null;

  if (stage === 'التصوير' && subtype) {
    folderUrl = createShootingFolder(project, subtype, row);
  } else {
    // إنشاء فولدر عام للعنصر
    folderUrl = createGenericFolder(project, stage, element, row);
  }

  if (folderUrl) {
    showSuccess('تم إنشاء الفولدر بنجاح');
  }
}

/**
 * إنشاء فولدر عام لعنصر
 */
function createGenericFolder(projectName, stageName, elementName, movementRow) {
  const mainFolder = getMainProductionFolder();

  if (!mainFolder) {
    return null;
  }

  try {
    // البحث عن فولدر المشروع
    const projectsSheet = getSheet(SHEETS.PROJECTS);
    const projectRow = findRowByValue(projectsSheet, PROJECT_COLS.NAME, projectName);

    if (projectRow === -1) return null;

    const projectCode = projectsSheet.getRange(projectRow, PROJECT_COLS.CODE).getValue();
    const projectFolderName = projectCode + ' - ' + projectName;

    let projectFolder = findFolderByName(mainFolder, projectFolderName);
    if (!projectFolder) {
      createProjectFolderStructure(projectName, projectCode);
      projectFolder = findFolderByName(mainFolder, projectFolderName);
    }

    if (!projectFolder) return null;

    // تحديد الفولدر الفرعي حسب المرحلة
    let targetFolderName = '01-الأوراق والأبحاث'; // الافتراضي

    const stageToFolder = {
      'الأوراق': '01-الأوراق والأبحاث',
      'التصوير': '03-التصوير',
      'الصوت': '04-الصوت',
      'أنيميشن': '05-الأنيميشن',
      'المونتاج': '07-المونتاج',
      'التسليم': '08-التسليم النهائي'
    };

    if (stageToFolder[stageName]) {
      targetFolderName = stageToFolder[stageName];
    }

    let targetFolder = findFolderByName(projectFolder, targetFolderName);
    if (!targetFolder) {
      targetFolder = projectFolder.createFolder(targetFolderName);
    }

    // إنشاء فولدر العنصر
    const elementFolderName = cleanText(elementName);
    let elementFolder = findFolderByName(targetFolder, elementFolderName);

    if (!elementFolder) {
      elementFolder = targetFolder.createFolder(elementFolderName);
    }

    const folderUrl = elementFolder.getUrl();

    // تحديث الرابط في شيت الحركة
    if (movementRow) {
      const movementSheet = getSheet(SHEETS.MOVEMENT);
      movementSheet.getRange(movementRow, MOVEMENT_COLS.LINK).setValue(folderUrl);
    }

    return folderUrl;

  } catch (error) {
    console.error('خطأ في إنشاء الفولدر:', error);
    return null;
  }
}

/**
 * إنشاء فولدرات المشروع عند إضافته
 */
function onProjectAdd(projectName, projectCode, row) {
  const folderUrl = createProjectFolderStructure(projectName, projectCode);

  if (folderUrl) {
    const sheet = getSheet(SHEETS.PROJECTS);
    sheet.getRange(row, PROJECT_COLS.FOLDER_LINK).setValue(folderUrl);
  }

  return folderUrl;
}

/**
 * فتح فولدر المشروع
 */
function openProjectFolder() {
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

  const folderLink = sheet.getRange(row, PROJECT_COLS.FOLDER_LINK).getValue();

  if (!folderLink) {
    // إنشاء الفولدر إذا لم يكن موجوداً
    const projectName = sheet.getRange(row, PROJECT_COLS.NAME).getValue();
    const projectCode = sheet.getRange(row, PROJECT_COLS.CODE).getValue();

    const newFolderUrl = onProjectAdd(projectName, projectCode, row);

    if (newFolderUrl) {
      showSuccess('تم إنشاء فولدر المشروع. الرابط: ' + newFolderUrl);
    }
  } else {
    showInfo('رابط الفولدر: ' + folderLink);
  }
}

/**
 * التحقق من إعدادات الفولدر الرئيسي
 */
function checkMainFolderSettings() {
  const folder = getMainProductionFolder();

  if (folder) {
    showSuccess('فولدر الإنتاج الرئيسي: ' + folder.getName());
  } else {
    showError('يرجى تحديد فولدر الإنتاج الرئيسي في شيت الإعدادات');
  }
}
