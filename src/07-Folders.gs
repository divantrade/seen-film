/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * إدارة الفولدرات التلقائية على Google Drive
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إدارة روابط الفولدرات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * حفظ رابط فولدر في شيت الروابط
 */
function saveFolderLink(projectCode, projectName, folderType, stage, element, folderId, folderUrl) {
  const sheet = getSheet(SHEETS.FOLDER_LINKS);
  if (!sheet) {
    console.log('شيت روابط الفولدرات غير موجود');
    return false;
  }

  // البحث عن سجل موجود بنفس المعايير
  const existingRow = findFolderLinkRow(projectCode, folderType, stage, element);

  const rowData = [
    projectCode,
    projectName,
    folderType,
    stage || '',
    element || '',
    folderId,
    folderUrl,
    getCurrentDate()
  ];

  if (existingRow > 0) {
    // تحديث السجل الموجود
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // إضافة سجل جديد
    const lastRow = getLastRowInColumn(sheet, FOLDER_LINKS_COLS.PROJECT_CODE);
    const newRow = Math.max(lastRow + 1, 2);
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  }

  return true;
}

/**
 * البحث عن صف رابط فولدر
 */
function findFolderLinkRow(projectCode, folderType, stage, element) {
  const sheet = getSheet(SHEETS.FOLDER_LINKS);
  if (!sheet) return -1;

  const lastRow = getLastRowInColumn(sheet, FOLDER_LINKS_COLS.PROJECT_CODE);
  if (lastRow <= 1) return -1;

  const data = sheet.getRange(2, 1, lastRow - 1, FOLDER_LINKS_COLS.ELEMENT).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][FOLDER_LINKS_COLS.PROJECT_CODE - 1] === projectCode &&
        data[i][FOLDER_LINKS_COLS.FOLDER_TYPE - 1] === folderType &&
        data[i][FOLDER_LINKS_COLS.STAGE - 1] === (stage || '') &&
        data[i][FOLDER_LINKS_COLS.ELEMENT - 1] === (element || '')) {
      return i + 2;
    }
  }

  return -1;
}

/**
 * الحصول على معرف الفولدر الرئيسي للمشروع من شيت الروابط
 */
function getProjectMainFolderId(projectCode) {
  return getFolderIdFromLinks(projectCode, FOLDER_TYPES.MAIN, '', '');
}

/**
 * الحصول على معرف فولدر المرحلة من شيت الروابط
 */
function getStageFolderId(projectCode, stageName) {
  return getFolderIdFromLinks(projectCode, FOLDER_TYPES.STAGE, stageName, '');
}

/**
 * الحصول على معرف فولدر عنصر من شيت الروابط
 */
function getElementFolderId(projectCode, stageName, elementName) {
  return getFolderIdFromLinks(projectCode, FOLDER_TYPES.ELEMENT, stageName, elementName);
}

/**
 * الحصول على معرف فولدر من شيت الروابط
 */
function getFolderIdFromLinks(projectCode, folderType, stage, element) {
  const sheet = getSheet(SHEETS.FOLDER_LINKS);
  if (!sheet) return null;

  const lastRow = getLastRowInColumn(sheet, FOLDER_LINKS_COLS.PROJECT_CODE);
  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, FOLDER_LINKS_COLS.FOLDER_ID).getValues();

  for (const row of data) {
    if (row[FOLDER_LINKS_COLS.PROJECT_CODE - 1] === projectCode &&
        row[FOLDER_LINKS_COLS.FOLDER_TYPE - 1] === folderType &&
        row[FOLDER_LINKS_COLS.STAGE - 1] === (stage || '') &&
        row[FOLDER_LINKS_COLS.ELEMENT - 1] === (element || '')) {
      return row[FOLDER_LINKS_COLS.FOLDER_ID - 1];
    }
  }

  return null;
}

/**
 * الحصول على فولدر من شيت الروابط باستخدام المعرف المحفوظ
 */
function getFolderFromLinks(projectCode, folderType, stage, element) {
  const folderId = getFolderIdFromLinks(projectCode, folderType, stage, element);
  if (!folderId) return null;

  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    console.error('خطأ في الوصول للفولدر:', error);
    return null;
  }
}

/**
 * الحصول على جميع روابط فولدرات مشروع
 */
function getProjectFolderLinks(projectCode) {
  const sheet = getSheet(SHEETS.FOLDER_LINKS);
  if (!sheet) return [];

  const lastRow = getLastRowInColumn(sheet, FOLDER_LINKS_COLS.PROJECT_CODE);
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, FOLDER_LINKS_COLS.CREATED_DATE).getValues();
  const links = [];

  for (const row of data) {
    if (row[FOLDER_LINKS_COLS.PROJECT_CODE - 1] === projectCode) {
      links.push({
        projectCode: row[FOLDER_LINKS_COLS.PROJECT_CODE - 1],
        projectName: row[FOLDER_LINKS_COLS.PROJECT_NAME - 1],
        folderType: row[FOLDER_LINKS_COLS.FOLDER_TYPE - 1],
        stage: row[FOLDER_LINKS_COLS.STAGE - 1],
        element: row[FOLDER_LINKS_COLS.ELEMENT - 1],
        folderId: row[FOLDER_LINKS_COLS.FOLDER_ID - 1],
        folderUrl: row[FOLDER_LINKS_COLS.FOLDER_URL - 1],
        createdDate: row[FOLDER_LINKS_COLS.CREATED_DATE - 1]
      });
    }
  }

  return links;
}

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
 * يحفظ جميع الروابط في شيت روابط الفولدرات
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
    let isNewProject = false;

    if (!projectFolder) {
      projectFolder = mainFolder.createFolder(projectFolderName);
      isNewProject = true;
    }

    // حفظ رابط الفولدر الرئيسي للمشروع
    saveFolderLink(
      projectCode,
      projectName,
      FOLDER_TYPES.MAIN,
      '',
      '',
      projectFolder.getId(),
      projectFolder.getUrl()
    );

    // إنشاء الفولدرات الفرعية وحفظ روابطها
    for (const subfolderName of FOLDER_STRUCTURE) {
      let subfolder = findFolderByName(projectFolder, subfolderName);

      if (!subfolder) {
        subfolder = projectFolder.createFolder(subfolderName);
      }

      // تحديد اسم المرحلة من اسم الفولدر
      const stageName = getStageNameFromFolderName(subfolderName);

      // حفظ رابط فولدر المرحلة
      saveFolderLink(
        projectCode,
        projectName,
        FOLDER_TYPES.STAGE,
        stageName,
        '',
        subfolder.getId(),
        subfolder.getUrl()
      );
    }

    return projectFolder.getUrl();

  } catch (error) {
    console.error('خطأ في إنشاء فولدرات المشروع:', error);
    showError('حدث خطأ أثناء إنشاء الفولدرات: ' + error.message);
    return null;
  }
}

/**
 * استخراج اسم المرحلة من اسم الفولدر
 */
function getStageNameFromFolderName(folderName) {
  const mapping = {
    'Research': 'الأوراق',
    'Shooting': 'التصوير',
    'Sound': 'الصوت',
    'Animation': 'أنيميشن',
    'Editing': 'المونتاج'
  };

  return mapping[folderName] || folderName;
}

/**
 * إنشاء فولدر لعنصر - اسم الفولدر = العنصر (بالإنجليزية)
 * يُوضع في الفولدر الصحيح حسب المرحلة
 * يستخدم شيت الروابط للوصول للفولدرات بدلاً من البحث بالاسم
 */
function createShootingFolder(projectName, subtype, movementRow, elementName) {
  try {
    // البحث عن بيانات المشروع
    const projectsSheet = getSheet(SHEETS.PROJECTS);
    const projectRow = findRowByValue(projectsSheet, PROJECT_COLS.NAME, projectName);

    if (projectRow === -1) {
      console.log('المشروع غير موجود');
      return null;
    }

    const projectCode = projectsSheet.getRange(projectRow, PROJECT_COLS.CODE).getValue();

    // قراءة المرحلة من شيت الحركة
    let stageName = 'التصوير'; // افتراضي
    if (movementRow) {
      const movementSheet = getSheet(SHEETS.MOVEMENT);
      stageName = movementSheet.getRange(movementRow, MOVEMENT_COLS.STAGE).getValue();
    }

    // محاولة الحصول على فولدر المرحلة من شيت الروابط أولاً
    let stageFolder = getFolderFromLinks(projectCode, FOLDER_TYPES.STAGE, stageName, '');

    // إذا لم يكن موجوداً في الروابط، نبحث بالطريقة القديمة ونحفظ الرابط
    if (!stageFolder) {
      const mainFolder = getMainProductionFolder();
      if (!mainFolder) {
        console.log('فولدر الإنتاج الرئيسي غير محدد');
        return null;
      }

      const projectFolderName = projectCode + ' - ' + projectName;
      let projectFolder = findFolderByName(mainFolder, projectFolderName);

      if (!projectFolder) {
        // إنشاء فولدر المشروع إذا لم يكن موجوداً
        const folderUrl = createProjectFolderStructure(projectName, projectCode);
        if (!folderUrl) return null;

        projectFolder = findFolderByName(mainFolder, projectFolderName);
      }

      // البحث عن فولدر المرحلة
      stageFolder = findStageFolderByName(projectFolder, stageName);

      if (stageFolder) {
        // حفظ الرابط للمرات القادمة
        saveFolderLink(
          projectCode,
          projectName,
          FOLDER_TYPES.STAGE,
          stageName,
          '',
          stageFolder.getId(),
          stageFolder.getUrl()
        );
      }
    }

    if (!stageFolder) {
      console.log('فولدر المرحلة غير موجود: ' + stageName);
      showError('فولدر المرحلة غير موجود: ' + stageName);
      return null;
    }

    // تحديد اسم الفولدر الجديد = العنصر
    const folderName = elementName || subtype;

    if (!folderName) {
      console.log('لم يتم تحديد اسم الفولدر');
      return null;
    }

    // التحقق إذا كان الفولدر موجوداً بالفعل في شيت الروابط
    let existingFolderId = getElementFolderId(projectCode, stageName, folderName);
    let newFolder;

    if (existingFolderId) {
      try {
        newFolder = DriveApp.getFolderById(existingFolderId);
      } catch (e) {
        // الفولدر المحفوظ غير موجود، نبحث أو ننشئ جديد
        existingFolderId = null;
      }
    }

    if (!newFolder) {
      // البحث أو إنشاء الفولدر
      newFolder = findFolderByName(stageFolder, folderName);

      if (!newFolder) {
        newFolder = stageFolder.createFolder(folderName);
      }

      // حفظ رابط الفولدر الجديد
      saveFolderLink(
        projectCode,
        projectName,
        FOLDER_TYPES.ELEMENT,
        stageName,
        folderName,
        newFolder.getId(),
        newFolder.getUrl()
      );
    }

    const folderUrl = newFolder.getUrl();

    // تحديث رابط الفولدر في شيت الحركة
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
 * البحث عن فولدر المرحلة داخل فولدر المشروع
 * يبحث بعدة أسماء محتملة
 */
function findStageFolderByName(projectFolder, stageName) {
  // تنظيف اسم المرحلة من المسافات الزائدة
  const cleanStageName = stageName.trim();

  // قائمة الأسماء المحتملة للبحث
  const possibleNames = [];

  // البحث في STAGE_TO_FOLDER أولاً
  if (STAGE_TO_FOLDER[cleanStageName]) {
    possibleNames.push(STAGE_TO_FOLDER[cleanStageName]);
  }

  // إضافة أسماء بديلة للتصوير
  if (cleanStageName === 'تصوير' || cleanStageName === 'التصوير') {
    possibleNames.push('Shooting');
  }

  // إضافة أسماء بديلة للأوراق
  if (cleanStageName === 'الأوراق' || cleanStageName === 'أوراق') {
    possibleNames.push('Research');
  }

  // إضافة أسماء بديلة للصوت
  if (cleanStageName === 'الصوت' || cleanStageName === 'صوت') {
    possibleNames.push('Sound');
  }

  // إضافة أسماء بديلة للأنيميشن
  if (cleanStageName === 'أنيميشن' || cleanStageName === 'انيميشن') {
    possibleNames.push('Animation');
  }

  // إضافة أسماء بديلة للمونتاج
  if (cleanStageName === 'المونتاج' || cleanStageName === 'مونتاج') {
    possibleNames.push('Editing');
  }

  // إضافة أسماء بديلة للتسليم
  if (cleanStageName === 'التسليم' || cleanStageName === 'تسليم') {
    possibleNames.push('Editing');
  }

  // البحث في كل الأسماء المحتملة
  for (const name of possibleNames) {
    const folder = findFolderByName(projectFolder, name);
    if (folder) {
      return folder;
    }
  }

  // محاولة أخيرة: البحث عن أي فولدر يحتوي على اسم المرحلة
  const folders = projectFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = folder.getName().toLowerCase();
    if (folderName.includes(cleanStageName.toLowerCase())) {
      return folder;
    }
  }

  return null;
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

  // التحقق من مرحلة التصوير (بالعربية أو كما في الإعدادات)
  const isShootingStage = stage === 'التصوير' || stage === 'تصوير' || stage.toLowerCase() === 'shooting';

  if (isShootingStage) {
    // اسم الفولدر = العنصر (element) إذا وُجد، وإلا المرحلة الفرعية
    folderUrl = createShootingFolder(project, subtype, row, element);
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
 * يستخدم شيت الروابط للوصول للفولدرات
 */
function createGenericFolder(projectName, stageName, elementName, movementRow) {
  try {
    // البحث عن بيانات المشروع
    const projectsSheet = getSheet(SHEETS.PROJECTS);
    const projectRow = findRowByValue(projectsSheet, PROJECT_COLS.NAME, projectName);

    if (projectRow === -1) return null;

    const projectCode = projectsSheet.getRange(projectRow, PROJECT_COLS.CODE).getValue();

    // محاولة الحصول على فولدر المرحلة من شيت الروابط أولاً
    let stageFolder = getFolderFromLinks(projectCode, FOLDER_TYPES.STAGE, stageName, '');

    // إذا لم يكن موجوداً في الروابط، نبحث بالطريقة القديمة
    if (!stageFolder) {
      const mainFolder = getMainProductionFolder();
      if (!mainFolder) return null;

      const projectFolderName = projectCode + ' - ' + projectName;
      let projectFolder = findFolderByName(mainFolder, projectFolderName);

      if (!projectFolder) {
        createProjectFolderStructure(projectName, projectCode);
        projectFolder = findFolderByName(mainFolder, projectFolderName);
      }

      if (!projectFolder) return null;

      // البحث عن فولدر المرحلة
      stageFolder = findStageFolderByName(projectFolder, stageName);

      if (stageFolder) {
        // حفظ الرابط للمرات القادمة
        saveFolderLink(
          projectCode,
          projectName,
          FOLDER_TYPES.STAGE,
          stageName,
          '',
          stageFolder.getId(),
          stageFolder.getUrl()
        );
      }
    }

    if (!stageFolder) return null;

    // إنشاء فولدر العنصر
    const elementFolderName = cleanText(elementName);

    // التحقق من وجود الفولدر في شيت الروابط
    let existingFolderId = getElementFolderId(projectCode, stageName, elementFolderName);
    let elementFolder;

    if (existingFolderId) {
      try {
        elementFolder = DriveApp.getFolderById(existingFolderId);
      } catch (e) {
        existingFolderId = null;
      }
    }

    if (!elementFolder) {
      elementFolder = findFolderByName(stageFolder, elementFolderName);

      if (!elementFolder) {
        elementFolder = stageFolder.createFolder(elementFolderName);
      }

      // حفظ رابط الفولدر الجديد
      saveFolderLink(
        projectCode,
        projectName,
        FOLDER_TYPES.ELEMENT,
        stageName,
        elementFolderName,
        elementFolder.getId(),
        elementFolder.getUrl()
      );
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

/**
 * تعيين فولدر الإنتاج الرئيسي مباشرة
 * استخدم هذه الدالة إذا كانت هناك مشكلة في الإعدادات
 */
function setMainFolderDirectly() {
  // ضع ID الفولدر هنا مباشرة
  const FOLDER_ID = '17BJ5ZPRX7NaqgVxo4bJBHfva1_UFerVb';

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    showSuccess('تم الاتصال بالفولدر: ' + folder.getName());

    // حفظ في الإعدادات
    const sheet = getSheet(SHEETS.SETTINGS);
    if (sheet) {
      sheet.getRange('B3').setValue('https://drive.google.com/drive/folders/' + FOLDER_ID);
      showSuccess('تم حفظ الرابط في الإعدادات');
    }
  } catch (e) {
    showError('خطأ: ' + e.message);
  }
}

/**
 * تشخيص مشكلة الفولدر
 */
function diagnoseFolderIssue() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSheet(SHEETS.SETTINGS);

  if (!sheet) {
    ui.alert('خطأ', 'شيت الإعدادات غير موجود!\n\nشغّل initializeSystem أولاً.', ui.ButtonSet.OK);
    return;
  }

  const cellA3 = sheet.getRange('A3').getValue();
  const cellB3 = sheet.getRange('B3').getValue();

  let message = 'تشخيص شيت الإعدادات:\n\n';
  message += 'A3: ' + cellA3 + '\n';
  message += 'B3: ' + cellB3 + '\n\n';

  if (!cellB3 || cellB3 === '(أدخل رابط الفولدر هنا)') {
    message += 'المشكلة: لم يتم إدخال رابط الفولدر في B3';
  } else {
    const folderId = extractFolderIdFromUrl(cellB3);
    message += 'ID المستخرج: ' + folderId + '\n\n';

    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        message += 'نجاح! اسم الفولدر: ' + folder.getName();
      } catch (e) {
        message += 'خطأ في الوصول: ' + e.message;
      }
    } else {
      message += 'المشكلة: لم يتم استخراج ID من الرابط';
    }
  }

  ui.alert('تشخيص', message, ui.ButtonSet.OK);
}

/**
 * تشخيص شامل للنظام
 */
function debugSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheets = ss.getSheets();

  let msg = 'الشيتات الموجودة:\n';
  sheets.forEach(s => msg += '- "' + s.getName() + '"\n');

  msg += '\n--- البحث عن شيت الإعدادات ---\n';
  msg += 'الاسم المطلوب: "' + SHEETS.SETTINGS + '"\n\n';

  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (settingsSheet) {
    msg += 'شيت الإعدادات موجود!\n';
    msg += 'A3 = "' + settingsSheet.getRange('A3').getValue() + '"\n';
    msg += 'B3 = "' + settingsSheet.getRange('B3').getValue() + '"';
  } else {
    msg += 'شيت الإعدادات غير موجود!';
  }

  ui.alert('تشخيص النظام', msg, ui.ButtonSet.OK);
}

/**
 * تتبع خطوات الحصول على الفولدر الرئيسي
 * هذه الدالة تشرح بالتفصيل أين تفشل العملية
 */
function traceMainFolder() {
  const ui = SpreadsheetApp.getUi();
  let trace = '=== تتبع خطوات الفولدر الرئيسي ===\n\n';

  // الخطوة 1: البحث عن شيت الإعدادات
  trace += '1️⃣ البحث عن شيت الإعدادات...\n';
  const sheet = getSheet(SHEETS.SETTINGS);

  if (!sheet) {
    trace += '❌ فشل: شيت الإعدادات غير موجود!\n';
    trace += 'الاسم المطلوب: "' + SHEETS.SETTINGS + '"\n';
    ui.alert('نتيجة التتبع', trace, ui.ButtonSet.OK);
    return;
  }
  trace += '✅ نجاح: وجدت شيت "' + sheet.getName() + '"\n\n';

  // الخطوة 2: قراءة B3
  trace += '2️⃣ قراءة الخلية B3...\n';
  const folderUrl = sheet.getRange('B3').getValue();
  trace += 'القيمة: "' + folderUrl + '"\n';

  if (!folderUrl) {
    trace += '❌ فشل: الخلية B3 فارغة!\n';
    ui.alert('نتيجة التتبع', trace, ui.ButtonSet.OK);
    return;
  }

  if (folderUrl === '(أدخل رابط الفولدر هنا)') {
    trace += '❌ فشل: لا يزال النص الافتراضي موجوداً!\n';
    ui.alert('نتيجة التتبع', trace, ui.ButtonSet.OK);
    return;
  }
  trace += '✅ نجاح: الرابط موجود\n\n';

  // الخطوة 3: استخراج ID
  trace += '3️⃣ استخراج ID الفولدر...\n';
  const folderId = extractFolderIdFromUrl(folderUrl);

  if (!folderId) {
    trace += '❌ فشل: لم يتم استخراج ID من الرابط!\n';
    ui.alert('نتيجة التتبع', trace, ui.ButtonSet.OK);
    return;
  }
  trace += '✅ نجاح: ID = "' + folderId + '"\n\n';

  // الخطوة 4: الوصول للفولدر
  trace += '4️⃣ محاولة الوصول للفولدر...\n';
  try {
    const folder = DriveApp.getFolderById(folderId);
    trace += '✅ نجاح التام! اسم الفولدر: "' + folder.getName() + '"\n\n';
    trace += '🎉 كل شيء يعمل بشكل صحيح!\n';
    trace += 'إذا كانت المشكلة مستمرة، جرب إعادة تحميل الصفحة.';
  } catch (e) {
    trace += '❌ فشل: خطأ في الوصول للفولدر!\n';
    trace += 'الخطأ: ' + e.message + '\n';
  }

  ui.alert('نتيجة التتبع', trace, ui.ButtonSet.OK);
}

/**
 * اختبار الدالة الفعلية getMainProductionFolder
 */
function testGetMainFolder() {
  const ui = SpreadsheetApp.getUi();

  try {
    const folder = getMainProductionFolder();

    if (folder) {
      ui.alert('نجاح!',
        'الدالة getMainProductionFolder() تعمل بنجاح!\n\n' +
        'اسم الفولدر: ' + folder.getName() + '\n' +
        'ID: ' + folder.getId(),
        ui.ButtonSet.OK);
    } else {
      ui.alert('فشل!',
        'الدالة getMainProductionFolder() أرجعت null!\n\n' +
        'هذا يعني أن هناك مشكلة في الكود.',
        ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('خطأ!',
      'حدث خطأ عند استدعاء getMainProductionFolder():\n\n' + e.message,
      ui.ButtonSet.OK);
  }
}

/**
 * اختبار إنشاء فولدر مشروع تجريبي
 */
function testCreateProjectFolder() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    'اختبار إنشاء فولدر',
    'أدخل اسم مشروع تجريبي:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const testName = result.getResponseText() || 'مشروع تجريبي';

  try {
    ui.alert('جاري الاختبار...', 'سيتم محاولة إنشاء فولدر للمشروع: ' + testName, ui.ButtonSet.OK);

    const folderUrl = createProjectFolderStructure(testName, 'TEST001');

    if (folderUrl) {
      ui.alert('نجاح!', 'تم إنشاء الفولدر بنجاح!\n\nالرابط: ' + folderUrl, ui.ButtonSet.OK);
    } else {
      ui.alert('فشل!', 'لم يتم إنشاء الفولدر. تحقق من رسائل الخطأ.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('خطأ!', 'حدث خطأ:\n\n' + e.message + '\n\nStack: ' + e.stack, ui.ButtonSet.OK);
  }
}
