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

/**
 * معالج Installable Trigger للتعامل مع Checkbox
 * هذا يعمل بصلاحيات كاملة على عكس onEdit العادي
 */
function onEditInstallable(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // التعامل مع checkbox في شيت الحركة
  if (sheetName === SHEETS.MOVEMENT) {
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // تحقق من أنه checkbox إنشاء الفولدر
    if (col === MOVEMENT_COLS.CREATE_FOLDER && row > 1) {
      const isChecked = e.range.getValue();

      if (isChecked === true) {
        // إنشاء الفولدر
        const project = sheet.getRange(row, MOVEMENT_COLS.PROJECT).getValue();
        const stage = sheet.getRange(row, MOVEMENT_COLS.STAGE).getValue();
        const subtype = sheet.getRange(row, MOVEMENT_COLS.SUBTYPE).getValue();
        const element = sheet.getRange(row, MOVEMENT_COLS.ELEMENT).getValue();
        const existingLink = sheet.getRange(row, MOVEMENT_COLS.LINK).getValue();

        if (existingLink) {
          // فولدر موجود بالفعل، إلغاء التحديد
          e.range.setValue(false);
          SpreadsheetApp.getActiveSpreadsheet().toast('يوجد رابط فولدر بالفعل', 'تنبيه', 3);
          return;
        }

        if (!project) {
          e.range.setValue(false);
          SpreadsheetApp.getActiveSpreadsheet().toast('يجب تحديد المشروع أولاً', 'خطأ', 3);
          return;
        }

        let folderUrl = null;
        const folderName = subtype || element || stage;

        if (stage === 'التصوير' && subtype) {
          folderUrl = createShootingFolder(project, subtype, row);
        } else if (folderName) {
          folderUrl = createGenericFolder(project, stage, folderName, row);
        }

        if (folderUrl) {
          e.range.setValue(false); // إلغاء التحديد بعد الإنشاء
          SpreadsheetApp.getActiveSpreadsheet().toast('تم إنشاء الفولدر بنجاح!', 'نجاح', 3);
        } else {
          e.range.setValue(false);
          SpreadsheetApp.getActiveSpreadsheet().toast('فشل في إنشاء الفولدر', 'خطأ', 3);
        }
      }
    }
  }

  // التعامل مع checkbox في شيت المشاريع (لإنشاء فولدر المشروع)
  if (sheetName === SHEETS.PROJECTS) {
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // يمكن إضافة checkbox لشيت المشاريع لاحقاً إذا أردت
  }
}

/**
 * تثبيت Installable Trigger
 * يجب تشغيل هذه الدالة مرة واحدة فقط
 */
function installEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // حذف الـ triggers القديمة من نفس النوع
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // إنشاء trigger جديد
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('تم تثبيت الـ Trigger بنجاح! الآن يمكنك استخدام ☑ لإنشاء الفولدرات', 'نجاح', 5);
}
