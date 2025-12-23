/**
 * 10-Automations.gs
 * يتعامل مع الأتمتة الذكية للملفات والمجلدات
 */

/**
 * الوظيفة الرئيسية التي يتم استدعاؤها عند تغيير مرحلة المشروع
 * @param {string} projectName اسم المشروع
 * @param {string} newStage المرحلة الجديدة
 */
function onProjectStageChange(projectName, newStage) {
  console.log(`بدء الأتمتة للمشروع: ${projectName} إلى المرحلة: ${newStage}`);
  
  const projectFolder = getFolderByName(projectName); 
  if (!projectFolder) {
    console.error('فولدر المشروع غير موجود');
    return;
  }

  // 1. تأكد من وجود الفولدرات الفرعية المناسبة للمرحلة
  ensureStageSubFolders(projectFolder, newStage);

  // 2. إنشاء الاختصارات (Shortcuts) للملفات المطلوبة
  createStageShortcuts(projectFolder, newStage);

  // 3. توزيع قائمة الضيوف (إذا كنا في مرحلة التحضير)
  if (newStage === STAGES.PRE_PRODUCTION.name) {
    distributeGuestList(projectName);
  }
}

/**
 * التأكد من وجود الفولدرات الفرعية الدقيقة داخل فولدر المرحلة
 */
function ensureStageSubFolders(projectFolder, stageName) {
  // الحصول على فولدر المرحلة
  const stageFolder = findSubFolder(projectFolder, stageName); 
  if (!stageFolder) return;

  // تعريف الفولدرات الفرعية المطلوبة لكل مرحلة
  let subFolders = [];
  
  if (stageName === STAGES.POST_ELEMENTS.name) {
    subFolders = ['الجرافيك', 'التعليق الصوتي', 'المكساج'];
  } else if (stageName === STAGES.EDITING.name) {
    subFolders = ['المسودات (Drafts)', 'التلوين'];
  } else if (stageName === STAGES.PRODUCTION.name) {
    // في الإنتاج، الفولدرات تعتمد على المدن، ويتم إنشاؤها عادة عند إضافة حركة تصوير
    // لكن يمكننا إنشاء فولدر عام
    subFolders = ['Data Management'];
  }

  subFolders.forEach(sub => {
    if (!stageFolder.getFoldersByName(sub).hasNext()) {
      stageFolder.createFolder(sub);
      console.log(`تم إنشاء فولدر فرعي: ${sub}`);
    }
  });
}

/**
 * البحث عن ملف معين في مرحلة "التطوير" وعمل اختصار له في المرحلة الحالية
 */
function createStageShortcuts(projectFolder, stageName) {
  // المصدر دائماً هو فولدر التطوير
  const devFolder = findSubFolder(projectFolder, STAGES.DEVELOPMENT.name);
  if (!devFolder) return;

  // الوجهة هي فولدر المرحلة الحالية
  const targetFolder = findSubFolder(projectFolder, stageName);
  if (!targetFolder) return;

  // قائمة الملفات المطلوب عمل اختصار لها
  let filesToShortcut = [];

  if (stageName === STAGES.PRE_PRODUCTION.name) {
    filesToShortcut = ['أسئلة الضيوف', 'قائمة الضيوف']; 
  } else if (stageName === STAGES.POST_ELEMENTS.name) {
    filesToShortcut = ['مواصفات الجرافيك', 'مواصفات الدراما'];
  }

  filesToShortcut.forEach(fileNamePart => {
    // نبحث عن أي ملف يحتوي اسمه على هذا المقطع
    const files = devFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().includes(fileNamePart)) {
        // التحقق مما إذا كان الملف موجوداً بالفعل في الوجهة لتجنب التكرار
        if (!isFileInFolder(targetFolder, file.getName())) {
            // إنشاء الاختصار (في Google Drive الحديث، إضافة الملف لفولدر آخر يعتبر اختصاراً)
            targetFolder.addFile(file); // هذا هو "الباب الثاني" لنفس الغرفة
            console.log(`تم عمل اختصار للملف: ${file.getName()} في ${stageName}`);
        }
      }
    }
  });
}

/**
 * توزيع قائمة الضيوف على الفيكسرز
 */
function distributeGuestList(projectName) {
  console.log(`بدء توزيع قائمة الضيوف لمشروع: ${projectName}`);
  
  // 1. الحصول على فولدر المشروع
  const projectFolder = getFolderByName(projectName);
  if (!projectFolder) return;

  // 2. البحث عن ملف "قائمة الضيوف" في فولدر التطوير (Development)
  const devFolder = findSubFolder(projectFolder, STAGES.DEVELOPMENT.name);
  if (!devFolder) return;

  const guestListFile = findFileInFolder(devFolder, 'قائمة الضيوف');
  if (!guestListFile) {
    console.log('لم يتم العثور على ملف قائمة الضيوف');
    return;
  }

  // 3. قراءة البيانات من الملف
  let ss;
  try {
    ss = SpreadsheetApp.openById(guestListFile.getId());
  } catch (e) {
    console.error('لا يمكن فتح ملف قائمة الضيوف كـ Spreadsheet');
    return;
  }

  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // لا توجد بيانات

  const headers = data[0];
  const fixerCol = headers.indexOf('الفيكسر المسؤول');
  if (fixerCol === -1) {
    console.error('عمود "الفيكسر المسؤول" غير موجود');
    return;
  }

  // 4. تجميع الضيوف حسب الفيكسر
  const guestsByFixer = {};
  for (let i = 1; i < data.length; i++) {
    const fixerName = data[i][fixerCol];
    if (!fixerName) continue;
    
    if (!guestsByFixer[fixerName]) guestsByFixer[fixerName] = [];
    guestsByFixer[fixerName].push(data[i]);
  }

  // 5. التأكد من وجود فولدر الفيكسرز في التحضير
  const preProdFolder = findSubFolder(projectFolder, STAGES.PRE_PRODUCTION.name);
  if (!preProdFolder) return;

  let fixersFolder = findSubFolder(preProdFolder, 'Fixers');
  if (!fixersFolder) {
    fixersFolder = preProdFolder.createFolder('Fixers');
  }

  // 6. لكل فيكسر، إنشاء/تحديث ملفه الخاص
  for (const fixerName in guestsByFixer) {
    const fileName = `قائمة الضيوف - ${fixerName}`;
    let fixerFile = findFileInFolder(fixersFolder, fileName);
    let fixerSS;

    if (!fixerFile) {
      fixerSS = SpreadsheetApp.create(fileName);
      fixerFile = DriveApp.getFileById(fixerSS.getId());
      fixersFolder.addFile(fixerFile);
      DriveApp.getRootFolder().removeFile(fixerFile); // إزالته من الـ root
    } else {
      fixerSS = SpreadsheetApp.openById(fixerFile.getId());
    }

    const fixerSheet = fixerSS.getSheets()[0];
    fixerSheet.clear();
    fixerSheet.appendRow(headers);
    fixerSheet.getRange(2, 1, guestsByFixer[fixerName].length, headers.length).setValues(guestsByFixer[fixerName]);
    
    // تنسيق بسيط
    fixerSheet.getRange(1, 1, 1, headers.length).setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold');
    fixerSheet.setFrozenRows(1);
    
    console.log(`تم تحديث ملف القائمة للفيكسر: ${fixerName}`);
  }
}

/**
 * البحث عن ملف معين في مرحلة "التطوير" وعمل اختصار له في المرحلة الحالية
 */
function createStageShortcuts(projectFolder, stageName) {
  // المصدر دائماً هو فولدر التطوير
  const devFolder = findSubFolder(projectFolder, STAGES.DEVELOPMENT.name);
  if (!devFolder) return;

  // الوجهة هي فولدر المرحلة الحالية
  const targetFolder = findSubFolder(projectFolder, stageName);
  if (!targetFolder) return;

  // قائمة الملفات المطلوب عمل اختصار لها حسب المرحلة
  let filesToShortcut = [];

  if (stageName === STAGES.PRE_PRODUCTION.name) {
    filesToShortcut = ['أسئلة الضيوف', 'قائمة الضيوف', 'البحث', 'المعالجة']; 
  } else if (stageName === STAGES.POST_ELEMENTS.name) {
    filesToShortcut = ['مواصفات الجرافيك', 'مواصفات الدراما', 'اسكربت أولي'];
  } else if (stageName === STAGES.EDITING.name) {
    filesToShortcut = ['بحث المصادر', 'المعالجة'];
  }

  filesToShortcut.forEach(fileNamePart => {
    const files = devFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().includes(fileNamePart)) {
        if (!isFileInFolder(targetFolder, file.getName())) {
            // إضافة الملف كمختصر (Shortcut)
            targetFolder.addFile(file);
            console.log(`تم عمل اختصار للملف: ${file.getName()} في ${stageName}`);
        }
      }
    }
  });
}

// --- Helpers ---

function findSubFolder(parentFolder, partialName) {
  if (!parentFolder) return null;
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().includes(partialName)) return f;
  }
  return null;
}

function findFileInFolder(folder, partialName) {
  if (!folder) return null;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().includes(partialName)) return file;
  }
  return null;
}

function isFileInFolder(folder, fileName) {
  if (!folder) return false;
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}

/**
 * الحصول على فولدر بالاسم (بحث شامل)
 */
function getFolderByName(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return null;
}
