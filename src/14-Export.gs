/**
 * ===================================================
 * 14-Export.gs - نظام تصدير شيتات منفصلة
 * ===================================================
 *
 * هذا الملف يحتوي على:
 * - تصدير للمصور (جدول التصوير، المواقع، الضيوف)
 * - تصدير للاستوديو الصوتي (مقاطع التعليق الصوتي)
 * - تصدير لاستوديو الرسوم (المشاهد، الاسكربتات)
 * - تصدير للمونتير (وحدات المونتاج، الملفات)
 * - تصدير تقرير مشروع كامل
 */

// ====================================================
// ثوابت التصدير
// ====================================================

/**
 * أنواع التصدير المتاحة
 */
const EXPORT_TYPES = {
  PHOTOGRAPHER: 'للمصور',
  VOICE_STUDIO: 'للاستوديو الصوتي',
  ANIMATION_STUDIO: 'لاستوديو الرسوم',
  EDITOR: 'للمونتير',
  PROJECT_REPORT: 'تقرير مشروع'
};

// ====================================================
// دوال عرض نوافذ التصدير
// ====================================================

/**
 * عرض نافذة اختيار نوع التصدير
 */
function showExportDialog(type) {
  const html = HtmlService.createHtmlOutput(getExportDialogHtml(type))
    .setWidth(500)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, `تصدير ${EXPORT_TYPES[type] || ''}`);
}

/**
 * عرض نافذة تصدير للمصور
 */
function showExportForPhotographerDialog() {
  showExportDialog('PHOTOGRAPHER');
}

/**
 * عرض نافذة تصدير للاستوديو الصوتي
 */
function showExportForVoiceStudioDialog() {
  showExportDialog('VOICE_STUDIO');
}

/**
 * عرض نافذة تصدير لاستوديو الرسوم
 */
function showExportForAnimationStudioDialog() {
  showExportDialog('ANIMATION_STUDIO');
}

/**
 * عرض نافذة تصدير للمونتير
 */
function showExportForEditorDialog() {
  showExportDialog('EDITOR');
}

/**
 * عرض نافذة تصدير تقرير مشروع
 */
function showExportProjectReportDialog() {
  showExportDialog('PROJECT_REPORT');
}

/**
 * الحصول على HTML لنافذة التصدير
 * @param {string} type - نوع التصدير
 * @returns {string} كود HTML
 */
function getExportDialogHtml(type) {
  let personLabel = '';
  let personFunction = '';
  let showPerson = true;

  switch (type) {
    case 'PHOTOGRAPHER':
      personLabel = 'المصور';
      personFunction = 'getPhotographersList';
      break;
    case 'VOICE_STUDIO':
      personLabel = 'الاستديو';
      personFunction = 'getVoiceStudiosList';
      break;
    case 'ANIMATION_STUDIO':
      personLabel = 'الاستديو';
      personFunction = 'getAnimationStudiosList';
      break;
    case 'EDITOR':
      personLabel = 'المونتير';
      personFunction = 'getEditorsList';
      break;
    case 'PROJECT_REPORT':
      showPerson = false;
      break;
  }

  return `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; direction: rtl; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; font-weight: bold; color: #333; }
        select, input {
          width: 100%;
          padding: 12px;
          border: 1px solid #ddd;
          border-radius: 6px;
          box-sizing: border-box;
          font-size: 14px;
        }
        select:focus, input:focus { border-color: #1a73e8; outline: none; }
        .btn {
          padding: 12px 24px;
          border: none;
          border-radius: 6px;
          cursor: pointer;
          margin-left: 10px;
          font-size: 14px;
          transition: background 0.3s;
        }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-primary:hover { background: #1557b0; }
        .btn-secondary { background: #5f6368; color: white; }
        .btn-secondary:hover { background: #3c4043; }
        .btn-container { text-align: left; margin-top: 30px; }
        .loading { display: none; text-align: center; padding: 20px; }
        .spinner {
          border: 4px solid #f3f3f3;
          border-top: 4px solid #1a73e8;
          border-radius: 50%;
          width: 40px;
          height: 40px;
          animation: spin 1s linear infinite;
          margin: 0 auto 15px;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .success-msg { color: #0f9d58; background: #e8f5e9; padding: 15px; border-radius: 6px; margin-top: 20px; }
        .error-msg { color: #d93025; background: #fce8e6; padding: 15px; border-radius: 6px; margin-top: 20px; }
        h3 { margin-top: 0; color: #202124; }
      </style>
    </head>
    <body>
      <div id="form-container">
        <h3>📤 تصدير ${EXPORT_TYPES[type] || ''}</h3>

        ${showPerson ? `
        <div class="form-group">
          <label>${personLabel} *</label>
          <select id="person" required>
            <option value="">جاري التحميل...</option>
          </select>
        </div>
        ` : ''}

        <div class="form-group">
          <label>المشروع *</label>
          <select id="project" required>
            <option value="">جاري التحميل...</option>
          </select>
        </div>

        <div class="form-group">
          <label>اسم الشيت (اختياري)</label>
          <input type="text" id="sheetName" placeholder="سيتم إنشاء اسم تلقائي إذا تركته فارغاً">
        </div>

        <div class="btn-container">
          <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>
          <button class="btn btn-primary" onclick="doExport()">تصدير</button>
        </div>
      </div>

      <div id="loading" class="loading">
        <div class="spinner"></div>
        <p>جاري إنشاء التصدير...</p>
      </div>

      <div id="result"></div>

      <script>
        const exportType = '${type}';

        // تحميل البيانات عند فتح النافذة
        window.onload = function() {
          ${showPerson ? `
          google.script.run.withSuccessHandler(function(items) {
            const select = document.getElementById('person');
            select.innerHTML = '<option value="">اختر ${personLabel}</option>';
            items.forEach(function(item) {
              select.innerHTML += '<option value="' + item.value + '">' + item.label + '</option>';
            });
          }).${personFunction}();
          ` : ''}

          google.script.run.withSuccessHandler(function(projects) {
            const select = document.getElementById('project');
            select.innerHTML = '<option value="">اختر المشروع</option>';
            projects.forEach(function(p) {
              select.innerHTML += '<option value="' + p.code + '">' + p.code + ' - ' + p.name + '</option>';
            });
          }).getActiveProjectsForExport();
        };

        function doExport() {
          const project = document.getElementById('project').value;
          ${showPerson ? `const person = document.getElementById('person').value;` : `const person = '';`}
          const sheetName = document.getElementById('sheetName').value;

          if (!project ${showPerson ? '|| !person' : ''}) {
            alert('يرجى ملء جميع الحقول المطلوبة');
            return;
          }

          document.getElementById('form-container').style.display = 'none';
          document.getElementById('loading').style.display = 'block';

          google.script.run
            .withSuccessHandler(handleSuccess)
            .withFailureHandler(handleError)
            .processExport(exportType, person, project, sheetName);
        }

        function handleSuccess(result) {
          document.getElementById('loading').style.display = 'none';
          if (result.success) {
            document.getElementById('result').innerHTML =
              '<div class="success-msg">' +
              '<strong>✅ تم التصدير بنجاح!</strong><br><br>' +
              'اسم الشيت: ' + result.sheetName + '<br><br>' +
              '<a href="' + result.url + '" target="_blank">🔗 فتح الشيت</a>' +
              '</div>';
          } else {
            document.getElementById('result').innerHTML =
              '<div class="error-msg">❌ ' + result.error + '</div>';
          }
        }

        function handleError(error) {
          document.getElementById('loading').style.display = 'none';
          document.getElementById('result').innerHTML =
            '<div class="error-msg">❌ حدث خطأ: ' + error.message + '</div>';
        }
      </script>
    </body>
    </html>
  `;
}

// ====================================================
// دوال جلب البيانات للقوائم المنسدلة
// ====================================================

/**
 * جلب قائمة المشاريع النشطة للتصدير
 * @returns {Array} قائمة المشاريع
 */
function getActiveProjectsForExport() {
  const projects = getActiveProjects();
  return projects.map(p => ({
    code: p[PROJECT_COLS.CODE],
    name: p[PROJECT_COLS.NAME]
  }));
}

/**
 * جلب قائمة المصورين
 * @returns {Array} قائمة المصورين
 */
function getPhotographersList() {
  const photographers = getPhotographers();
  return photographers.map(p => ({
    value: p[PHOTOGRAPHER_COLS.CODE],
    label: p[PHOTOGRAPHER_COLS.NAME]
  }));
}

/**
 * جلب قائمة استوديوهات التعليق الصوتي
 * @returns {Array} قائمة الاستوديوهات
 */
function getVoiceStudiosList() {
  // جلب الاستوديوهات من بيانات التعليق الصوتي
  const voiceOvers = getAllVoiceOver();
  const studios = [...new Set(voiceOvers.map(v => v[VO_COLS.STUDIO]).filter(s => s))];
  return studios.map(s => ({ value: s, label: s }));
}

/**
 * جلب قائمة استوديوهات الرسوم المتحركة
 * @returns {Array} قائمة الاستوديوهات
 */
function getAnimationStudiosList() {
  const animations = getAllAnimation();
  const studios = [...new Set(animations.map(a => a[ANIM_COLS.STUDIO]).filter(s => s))];
  return studios.map(s => ({ value: s, label: s }));
}

/**
 * جلب قائمة المونتيرين (من الفريق)
 * @returns {Array} قائمة المونتيرين
 */
function getEditorsList() {
  const team = getTeamMembers();
  // نفترض أن المونتيرين لديهم دور "مونتير" أو يتعاملون مع مرحلة المونتاج
  const editors = team.filter(t => {
    const role = (t[TEAM_COLS.ROLE] || '').toLowerCase();
    const stages = (t[TEAM_COLS.STAGES] || '').toLowerCase();
    return role.includes('مونتير') || role.includes('مونتاج') ||
           stages.includes('مونتاج') || stages.includes('montage');
  });

  return editors.map(e => ({
    value: e[TEAM_COLS.CODE],
    label: e[TEAM_COLS.NAME]
  }));
}

/**
 * جلب جميع التعليقات الصوتية
 * @returns {Array} مصفوفة بجميع التعليقات الصوتية
 */
function getAllVoiceOver() {
  const sheet = getSheet(SHEETS.VOICEOVER);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headers = Object.values(VO_COLS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  return data.map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  }).filter(item => item[VO_COLS.CODE]);
}

/**
 * جلب جميع الرسوم المتحركة
 * @returns {Array} مصفوفة بجميع الرسوم المتحركة
 */
function getAllAnimation() {
  const sheet = getSheet(SHEETS.ANIMATION);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headers = Object.values(ANIM_COLS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  return data.map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  }).filter(item => item[ANIM_COLS.CODE]);
}

// ====================================================
// دوال التصدير الرئيسية
// ====================================================

/**
 * معالجة طلب التصدير
 * @param {string} type - نوع التصدير
 * @param {string} person - كود/اسم الشخص أو الاستوديو
 * @param {string} projectCode - كود المشروع
 * @param {string} customSheetName - اسم مخصص للشيت (اختياري)
 * @returns {Object} نتيجة التصدير
 */
function processExport(type, person, projectCode, customSheetName) {
  try {
    let result;

    switch (type) {
      case 'PHOTOGRAPHER':
        result = exportForPhotographer(person, projectCode, customSheetName);
        break;
      case 'VOICE_STUDIO':
        result = exportForVoiceStudio(person, projectCode, customSheetName);
        break;
      case 'ANIMATION_STUDIO':
        result = exportForAnimationStudio(person, projectCode, customSheetName);
        break;
      case 'EDITOR':
        result = exportForEditor(person, projectCode, customSheetName);
        break;
      case 'PROJECT_REPORT':
        result = exportProjectReport(projectCode, customSheetName);
        break;
      default:
        return { success: false, error: 'نوع تصدير غير معروف' };
    }

    return result;

  } catch (error) {
    console.error('Export Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * تصدير للمصور
 * @param {string} photographerCode - كود المصور
 * @param {string} projectCode - كود المشروع
 * @param {string} customName - اسم مخصص (اختياري)
 * @returns {Object} نتيجة التصدير
 */
function exportForPhotographer(photographerCode, projectCode, customName) {
  // جلب بيانات المصور
  const photographers = getPhotographers();
  const photographer = photographers.find(p => p[PHOTOGRAPHER_COLS.CODE] === photographerCode);
  const photographerName = photographer ? photographer[PHOTOGRAPHER_COLS.NAME] : photographerCode;

  // جلب بيانات المشروع
  const project = getProjectByCode(projectCode);
  const projectName = project ? project[PROJECT_COLS.NAME] : projectCode;

  // جلب الضيوف المجدولين للتصوير
  const guests = getGuestsByProject(projectCode);
  const scheduledGuests = guests.filter(g =>
    g[GUEST_COLS.SHOOT_STATUS] === SHOOT_STATUS.SCHEDULED ||
    g[GUEST_COLS.SHOOT_STATUS] === SHOOT_STATUS.PENDING
  );

  // إعداد البيانات
  const headers = ['اسم الضيف', 'البلد', 'مكان التصوير', 'تاريخ التصوير', 'حالة التصوير', 'نوع المشاركة', 'ملاحظات'];
  const data = scheduledGuests.map(g => [
    g[GUEST_COLS.NAME],
    g[GUEST_COLS.COUNTRY],
    g[GUEST_COLS.SHOOT_LOCATION],
    g[GUEST_COLS.SHOOT_DATE] ? formatDate(g[GUEST_COLS.SHOOT_DATE]) : '',
    g[GUEST_COLS.SHOOT_STATUS],
    g[GUEST_COLS.TYPE],
    g[GUEST_COLS.NOTES]
  ]);

  // إنشاء اسم الشيت
  const sheetName = customName || `تصوير - ${photographerName} - ${projectName}`;

  // إنشاء الشيت
  const result = createExportSheet(sheetName, data, headers, `جدول تصوير للمصور: ${photographerName}`, projectCode);

  // تسجيل التصدير
  if (result.success) {
    logExportRecord('تصدير للمصور', photographerName, projectCode, result.url);
  }

  return result;
}

/**
 * تصدير للاستوديو الصوتي
 * @param {string} studioName - اسم الاستوديو
 * @param {string} projectCode - كود المشروع
 * @param {string} customName - اسم مخصص (اختياري)
 * @returns {Object} نتيجة التصدير
 */
function exportForVoiceStudio(studioName, projectCode, customName) {
  const project = getProjectByCode(projectCode);
  const projectName = project ? project[PROJECT_COLS.NAME] : projectCode;

  // جلب مقاطع التعليق الصوتي للمشروع والاستوديو
  const voiceOvers = getVoiceOverByProject(projectCode);
  const studioVO = voiceOvers.filter(v => v[VO_COLS.STUDIO] === studioName);

  // إعداد البيانات
  const headers = ['الكود', 'النوع', 'رقم المقطع', 'النص', 'المؤدي', 'اللغة', 'الحالة', 'المدة (دقيقة)', 'ملاحظات'];
  const data = studioVO.map(v => [
    v[VO_COLS.CODE],
    v[VO_COLS.TYPE],
    v[VO_COLS.SEGMENT_NUM],
    v[VO_COLS.TEXT],
    v[VO_COLS.PERFORMER],
    v[VO_COLS.LANGUAGE],
    v[VO_COLS.STATUS],
    v[VO_COLS.DURATION],
    v[VO_COLS.NOTES]
  ]);

  const sheetName = customName || `تعليق صوتي - ${studioName} - ${projectName}`;
  const result = createExportSheet(sheetName, data, headers, `مقاطع التعليق الصوتي - استوديو: ${studioName}`, projectCode);

  if (result.success) {
    logExportRecord('تصدير للاستوديو الصوتي', studioName, projectCode, result.url);
  }

  return result;
}

/**
 * تصدير لاستوديو الرسوم المتحركة
 * @param {string} studioName - اسم الاستوديو
 * @param {string} projectCode - كود المشروع
 * @param {string} customName - اسم مخصص (اختياري)
 * @returns {Object} نتيجة التصدير
 */
function exportForAnimationStudio(studioName, projectCode, customName) {
  const project = getProjectByCode(projectCode);
  const projectName = project ? project[PROJECT_COLS.NAME] : projectCode;

  // جلب مشاهد الرسوم المتحركة للمشروع والاستوديو
  const animations = getAnimationByProject(projectCode);
  const studioAnim = animations.filter(a => a[ANIM_COLS.STUDIO] === studioName);

  // إعداد البيانات
  const headers = ['الكود', 'رقم المشهد', 'الوصف', 'النوع', 'المدة (ثانية)', 'رابط السكريبت', 'المحرك', 'الحالة', 'ملاحظات'];
  const data = studioAnim.map(a => [
    a[ANIM_COLS.CODE],
    a[ANIM_COLS.SCENE_NUM],
    a[ANIM_COLS.DESCRIPTION],
    a[ANIM_COLS.TYPE],
    a[ANIM_COLS.DURATION],
    a[ANIM_COLS.SCRIPT_LINK],
    a[ANIM_COLS.ANIMATOR],
    a[ANIM_COLS.STATUS],
    a[ANIM_COLS.NOTES]
  ]);

  const sheetName = customName || `رسوم متحركة - ${studioName} - ${projectName}`;
  const result = createExportSheet(sheetName, data, headers, `مشاهد الرسوم المتحركة - استوديو: ${studioName}`, projectCode);

  if (result.success) {
    logExportRecord('تصدير لاستوديو الرسوم', studioName, projectCode, result.url);
  }

  return result;
}

/**
 * تصدير للمونتير
 * @param {string} editorCode - كود المونتير
 * @param {string} projectCode - كود المشروع
 * @param {string} customName - اسم مخصص (اختياري)
 * @returns {Object} نتيجة التصدير
 */
function exportForEditor(editorCode, projectCode, customName) {
  const team = getTeamMembers();
  const editor = team.find(t => t[TEAM_COLS.CODE] === editorCode);
  const editorName = editor ? editor[TEAM_COLS.NAME] : editorCode;

  const project = getProjectByCode(projectCode);
  const projectName = project ? project[PROJECT_COLS.NAME] : projectCode;

  // جلب مهام المونتاج من شيت الحركة
  const movements = getMovementByProject(projectCode);
  const editingTasks = movements.filter(m => {
    const stage = m[MOVEMENT_COLS.STAGE] || '';
    return stage.includes('مونتاج') || stage.includes('EDITING');
  });

  // جلب التعليقات الصوتية المكتملة
  const voiceOvers = getVoiceOverByProject(projectCode);
  const completedVO = voiceOvers.filter(v => v[VO_COLS.STATUS] === 'مكتمل');

  // جلب الرسوم المتحركة المكتملة
  const animations = getAnimationByProject(projectCode);
  const completedAnim = animations.filter(a => a[ANIM_COLS.STATUS] === 'مكتمل');

  // جلب الأرشيف المرخص
  const archive = getArchiveByProject(projectCode);
  const licensedArchive = archive.filter(a => a[ARCHIVE_COLS.LICENSE_STATUS] === LICENSE_STATUS.LICENSED);

  // إنشاء شيت متعدد الأقسام
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = customName || `مونتاج - ${editorName} - ${projectName}`;

  // التحقق من عدم وجود شيت بنفس الاسم
  let exportSheet = ss.getSheetByName(sheetName);
  if (exportSheet) {
    ss.deleteSheet(exportSheet);
  }
  exportSheet = ss.insertSheet(sheetName);

  let currentRow = 1;

  // قسم معلومات المشروع
  exportSheet.getRange(currentRow, 1).setValue(`📋 تصدير للمونتير: ${editorName}`).setFontWeight('bold').setFontSize(14);
  currentRow++;
  exportSheet.getRange(currentRow, 1).setValue(`المشروع: ${projectCode} - ${projectName}`);
  currentRow++;
  exportSheet.getRange(currentRow, 1).setValue(`تاريخ التصدير: ${formatDate(new Date())}`);
  currentRow += 2;

  // قسم مهام المونتاج
  exportSheet.getRange(currentRow, 1).setValue('🎞️ مهام المونتاج').setFontWeight('bold').setBackground('#E3F2FD');
  currentRow++;
  if (editingTasks.length > 0) {
    const taskHeaders = ['المهمة', 'الحالة', 'الموعد النهائي', 'ملاحظات'];
    exportSheet.getRange(currentRow, 1, 1, taskHeaders.length).setValues([taskHeaders]).setFontWeight('bold');
    currentRow++;
    const taskData = editingTasks.map(t => [t[MOVEMENT_COLS.TASK], t[MOVEMENT_COLS.STATUS], formatDate(t[MOVEMENT_COLS.DEADLINE]), t[MOVEMENT_COLS.NOTES]]);
    exportSheet.getRange(currentRow, 1, taskData.length, 4).setValues(taskData);
    currentRow += taskData.length + 1;
  } else {
    exportSheet.getRange(currentRow, 1).setValue('لا توجد مهام');
    currentRow += 2;
  }

  // قسم التعليقات الصوتية المكتملة
  exportSheet.getRange(currentRow, 1).setValue('🎙️ التعليقات الصوتية المكتملة').setFontWeight('bold').setBackground('#E8F5E9');
  currentRow++;
  if (completedVO.length > 0) {
    const voHeaders = ['الكود', 'النوع', 'المؤدي', 'المدة', 'رابط الملف'];
    exportSheet.getRange(currentRow, 1, 1, voHeaders.length).setValues([voHeaders]).setFontWeight('bold');
    currentRow++;
    const voData = completedVO.map(v => [v[VO_COLS.CODE], v[VO_COLS.TYPE], v[VO_COLS.PERFORMER], v[VO_COLS.DURATION], v[VO_COLS.FILE_LINK]]);
    exportSheet.getRange(currentRow, 1, voData.length, 5).setValues(voData);
    currentRow += voData.length + 1;
  } else {
    exportSheet.getRange(currentRow, 1).setValue('لا توجد تعليقات صوتية مكتملة');
    currentRow += 2;
  }

  // قسم الرسوم المتحركة المكتملة
  exportSheet.getRange(currentRow, 1).setValue('🎨 الرسوم المتحركة المكتملة').setFontWeight('bold').setBackground('#FFF3E0');
  currentRow++;
  if (completedAnim.length > 0) {
    const animHeaders = ['الكود', 'المشهد', 'النوع', 'المدة', 'رابط الملف'];
    exportSheet.getRange(currentRow, 1, 1, animHeaders.length).setValues([animHeaders]).setFontWeight('bold');
    currentRow++;
    const animData = completedAnim.map(a => [a[ANIM_COLS.CODE], a[ANIM_COLS.SCENE_NUM], a[ANIM_COLS.TYPE], a[ANIM_COLS.DURATION], a[ANIM_COLS.FILE_LINK]]);
    exportSheet.getRange(currentRow, 1, animData.length, 5).setValues(animData);
    currentRow += animData.length + 1;
  } else {
    exportSheet.getRange(currentRow, 1).setValue('لا توجد رسوم متحركة مكتملة');
    currentRow += 2;
  }

  // قسم الأرشيف المرخص
  exportSheet.getRange(currentRow, 1).setValue('📁 الأرشيف المرخص').setFontWeight('bold').setBackground('#F3E5F5');
  currentRow++;
  if (licensedArchive.length > 0) {
    const archiveHeaders = ['الكود', 'العنوان', 'النوع', 'المدة', 'رابط الملف'];
    exportSheet.getRange(currentRow, 1, 1, archiveHeaders.length).setValues([archiveHeaders]).setFontWeight('bold');
    currentRow++;
    const archiveData = licensedArchive.map(a => [a[ARCHIVE_COLS.CODE], a[ARCHIVE_COLS.TITLE], a[ARCHIVE_COLS.TYPE], a[ARCHIVE_COLS.DURATION], a[ARCHIVE_COLS.FILE_LINK]]);
    exportSheet.getRange(currentRow, 1, archiveData.length, 5).setValues(archiveData);
  } else {
    exportSheet.getRange(currentRow, 1).setValue('لا توجد مواد أرشيفية مرخصة');
  }

  exportSheet.setRightToLeft(true);
  exportSheet.autoResizeColumns(1, 6);

  const url = ss.getUrl() + '#gid=' + exportSheet.getSheetId();

  logExportRecord('تصدير للمونتير', editorName, projectCode, url);

  return {
    success: true,
    sheetName: sheetName,
    url: url
  };
}

/**
 * تصدير تقرير مشروع كامل
 * @param {string} projectCode - كود المشروع
 * @param {string} customName - اسم مخصص (اختياري)
 * @returns {Object} نتيجة التصدير
 */
function exportProjectReport(projectCode, customName) {
  const summary = generateProjectSummary(projectCode);
  if (!summary) {
    return { success: false, error: 'المشروع غير موجود' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = customName || `تقرير - ${projectCode} - ${summary.project.name}`;

  // حذف الشيت إذا كان موجوداً
  let exportSheet = ss.getSheetByName(sheetName);
  if (exportSheet) {
    ss.deleteSheet(exportSheet);
  }
  exportSheet = ss.insertSheet(sheetName);

  let row = 1;

  // عنوان التقرير
  exportSheet.getRange(row, 1).setValue(`📋 تقرير المشروع: ${projectCode}`).setFontWeight('bold').setFontSize(16);
  exportSheet.getRange(row, 1, 1, 6).merge().setBackground('#1a73e8').setFontColor('white');
  row += 2;

  // معلومات المشروع الأساسية
  exportSheet.getRange(row, 1).setValue('معلومات المشروع').setFontWeight('bold').setFontSize(12);
  row++;
  const projectInfo = [
    ['الاسم', summary.project.name],
    ['العميل', summary.project.client],
    ['الحالة', summary.project.status],
    ['تاريخ البداية', formatDate(summary.project.startDate)],
    ['الموعد النهائي', formatDate(summary.project.deadline)],
    ['الأيام المتبقية', summary.project.daysRemaining !== null ? summary.project.daysRemaining : 'غير محدد']
  ];
  exportSheet.getRange(row, 1, projectInfo.length, 2).setValues(projectInfo);
  row += projectInfo.length + 1;

  // إحصائيات المهام
  exportSheet.getRange(row, 1).setValue('📊 إحصائيات المهام').setFontWeight('bold').setBackground('#E3F2FD');
  row++;
  const taskStats = [
    ['إجمالي المهام', summary.tasks.total],
    ['مكتمل', summary.tasks.completed],
    ['قيد العمل', summary.tasks.inProgress],
    ['متأخر', summary.tasks.overdue],
    ['نسبة الإكمال', summary.tasks.completionRate + '%']
  ];
  exportSheet.getRange(row, 1, taskStats.length, 2).setValues(taskStats);
  row += taskStats.length + 1;

  // إحصائيات الضيوف
  exportSheet.getRange(row, 1).setValue('👥 الضيوف').setFontWeight('bold').setBackground('#E8F5E9');
  row++;
  const guestStats = [
    ['إجمالي الضيوف', summary.guests.total],
    ['تم التواصل', summary.guests.contacted],
    ['مؤكد للتصوير', summary.guests.confirmed],
    ['تم التصوير', summary.guests.shotCompleted]
  ];
  exportSheet.getRange(row, 1, guestStats.length, 2).setValues(guestStats);
  row += guestStats.length + 1;

  // إحصائيات الإنتاج
  exportSheet.getRange(row, 1).setValue('🎬 إحصائيات الإنتاج').setFontWeight('bold').setBackground('#FFF3E0');
  row++;
  const prodStats = [
    ['التعليق الصوتي - إجمالي', summary.voiceOver.total],
    ['التعليق الصوتي - مكتمل', summary.voiceOver.completed],
    ['الرسوم المتحركة - إجمالي', summary.animation.total],
    ['الرسوم المتحركة - مكتمل', summary.animation.completed],
    ['الأرشيف - إجمالي', summary.archive.total],
    ['الأرشيف - مرخص', summary.archive.licensed]
  ];
  exportSheet.getRange(row, 1, prodStats.length, 2).setValues(prodStats);
  row += prodStats.length + 1;

  // التنبيهات
  if (summary.alerts.length > 0) {
    exportSheet.getRange(row, 1).setValue('⚠️ التنبيهات').setFontWeight('bold').setBackground('#FFCDD2');
    row++;
    summary.alerts.forEach(alert => {
      exportSheet.getRange(row, 1).setValue(`${alert.icon} ${alert.message}`);
      row++;
    });
  }

  exportSheet.getRange(row + 1, 1).setValue(`تاريخ التقرير: ${formatDate(new Date())}`).setFontColor('#666');

  exportSheet.setRightToLeft(true);
  exportSheet.autoResizeColumns(1, 6);

  const url = ss.getUrl() + '#gid=' + exportSheet.getSheetId();

  logExportRecord('تقرير مشروع', projectCode, projectCode, url);

  return {
    success: true,
    sheetName: sheetName,
    url: url
  };
}

// ====================================================
// دوال إنشاء شيت التصدير
// ====================================================

/**
 * إنشاء شيت تصدير جديد
 * @param {string} title - عنوان الشيت
 * @param {Array} data - البيانات
 * @param {Array} headers - العناوين
 * @param {string} subtitle - عنوان فرعي (اختياري)
 * @param {string} projectCode - كود المشروع
 * @returns {Object} نتيجة الإنشاء
 */
function createExportSheet(title, data, headers, subtitle, projectCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // التحقق من عدم وجود شيت بنفس الاسم
    let sheet = ss.getSheetByName(title);
    if (sheet) {
      ss.deleteSheet(sheet);
    }

    sheet = ss.insertSheet(title);

    let currentRow = 1;

    // العنوان الرئيسي
    sheet.getRange(currentRow, 1).setValue(subtitle || title)
      .setFontWeight('bold')
      .setFontSize(14);
    sheet.getRange(currentRow, 1, 1, headers.length).merge()
      .setBackground('#1a73e8')
      .setFontColor('white');
    currentRow++;

    // معلومات إضافية
    sheet.getRange(currentRow, 1).setValue(`تاريخ التصدير: ${formatDate(new Date())}`);
    currentRow += 2;

    // العناوين
    sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold')
      .setBackground('#E3F2FD');
    currentRow++;

    // البيانات
    if (data.length > 0) {
      sheet.getRange(currentRow, 1, data.length, headers.length).setValues(data);
    } else {
      sheet.getRange(currentRow, 1).setValue('لا توجد بيانات');
    }

    // تنسيق الشيت
    sheet.setRightToLeft(true);
    sheet.autoResizeColumns(1, headers.length);

    // الحصول على رابط الشيت
    const url = ss.getUrl() + '#gid=' + sheet.getSheetId();

    return {
      success: true,
      sheetName: title,
      url: url
    };

  } catch (error) {
    console.error('Error creating export sheet:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ====================================================
// دوال سجل التصديرات
// ====================================================

/**
 * تسجيل عملية تصدير في السجل
 * @param {string} type - نوع التصدير
 * @param {string} person - الشخص/الاستوديو
 * @param {string} project - كود المشروع
 * @param {string} url - رابط الشيت
 */
function logExportRecord(type, person, project, url) {
  try {
    const sheet = getSheet(SHEETS.EXPORT_LOG);
    if (!sheet) return;

    const id = generateExportId();
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail() || 'غير معروف';

    sheet.appendRow([
      id,
      timestamp,
      type,
      `${person} - ${project}`,
      user,
      url,
      ''
    ]);

  } catch (error) {
    console.error('Error logging export:', error);
  }
}

/**
 * توليد معرف فريد للتصدير
 * @returns {string} معرف التصدير
 */
function generateExportId() {
  const sheet = getSheet(SHEETS.EXPORT_LOG);
  if (!sheet) return 'EX-001';

  const lastRow = sheet.getLastRow();
  return `EX-${(lastRow).toString().padStart(3, '0')}`;
}

/**
 * جلب سجل التصديرات
 * @param {number} limit - عدد السجلات (افتراضي 50)
 * @returns {Array} مصفوفة بسجلات التصدير
 */
function getExportLog(limit = 50) {
  const sheet = getSheet(SHEETS.EXPORT_LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastRow = Math.min(sheet.getLastRow(), limit + 1);
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  return data.map(row => ({
    id: row[0],
    date: row[1],
    type: row[2],
    details: row[3],
    user: row[4],
    url: row[5],
    notes: row[6]
  })).reverse(); // الأحدث أولاً
}

/**
 * عرض سجل التصديرات الأخيرة
 */
function showRecentExports() {
  const exports = getExportLog(10);

  if (exports.length === 0) {
    showInfo('لا توجد تصديرات مسجلة');
    return;
  }

  let message = 'آخر 10 تصديرات:\n\n';
  exports.forEach((exp, index) => {
    message += `${index + 1}. ${exp.type} - ${exp.details}\n   ${formatDate(exp.date)}\n\n`;
  });

  showInfo(message, 'سجل التصديرات');
}
