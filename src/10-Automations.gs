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
 * (هذه وظيفة متقدمة تفترض وجود ملف Google Sheet للضيوف)
 */
function distributeGuestList(projectName) {
  // 1. البحث عن ملف "قائمة الضيوف" في فولدر التطوير
  // 2. قراءة البيانات
  // 3. لكل فيكسر، إنشاء/تحديث شيت خاص به في فولدر التحضير
  
  // سنقوم بتنفيذ الهيكل الأساسي الآن، والتفاصيل عند وجود ملف حقيقي
  console.log('جاري توزيع قائمة الضيوف (Placeholder)...');
}

// --- Helpers ---

function findSubFolder(parentFolder, partialName) {
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().includes(partialName)) return f;
  }
  return null;
}

function isFileInFolder(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}
