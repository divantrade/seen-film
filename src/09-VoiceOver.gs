/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف إدارة التعليق الصوتي
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على جميع الدوال المتعلقة بإدارة التعليق الصوتي
 * بما في ذلك الراوي الرئيسي، الاقتباسات، والدوبلاج
 */

// ═══════════════════════════════════════════════════════════════════════════════
// ثوابت أعمدة شيت التعليق الصوتي
// ═══════════════════════════════════════════════════════════════════════════════

const VO_COLS = {
  CODE: 1,           // كود المقطع (VO-001)
  PROJECT: 2,        // المشروع
  TYPE: 3,           // النوع (راوي رئيسي/اقتباس/دوبلاج)
  SEGMENT_NUM: 4,    // رقم المقطع
  TEXT: 5,           // النص/الوصف
  PERFORMER: 6,      // المؤدي
  LANGUAGE: 7,       // اللغة
  STUDIO: 8,         // الاستوديو
  STATUS: 9,         // الحالة
  DURATION: 10,      // المدة (ثانية)
  FILE_LINK: 11,     // رابط الملف
  NOTES: 12,         // ملاحظات
  CREATED_AT: 13,    // تاريخ الإنشاء
  UPDATED_AT: 14     // آخر تحديث
};

// أنواع التعليق الصوتي
const VO_TYPES = {
  NARRATOR: 'راوي رئيسي',
  QUOTE: 'اقتباس',
  DUBBING: 'دوبلاج',
  TRANSLATION: 'ترجمة صوتية'
};

// حالات التعليق الصوتي
const VO_STATUS = {
  SCRIPT_PENDING: 'في انتظار النص',
  SCRIPT_READY: 'النص جاهز',
  RECORDING_SCHEDULED: 'التسجيل مجدول',
  RECORDING_IN_PROGRESS: 'جاري التسجيل',
  EDITING: 'جاري المونتاج',
  COMPLETED: 'مكتمل',
  CANCELLED: 'ملغي'
};

// اللغات المتاحة
const VO_LANGUAGES = [
  'العربية',
  'الإنجليزية',
  'التركية',
  'الفرنسية',
  'الألمانية',
  'أخرى'
];

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الاستعلام عن التعليق الصوتي
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على جميع مقاطع التعليق الصوتي
 * @returns {Array} مصفوفة من كائنات التعليق الصوتي
 */
function getAllVoiceOver() {
  const sheet = getSheet(SHEETS.VOICEOVER);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  return data.map(row => ({
    code: row[VO_COLS.CODE - 1],
    project: row[VO_COLS.PROJECT - 1],
    type: row[VO_COLS.TYPE - 1],
    segmentNum: row[VO_COLS.SEGMENT_NUM - 1],
    text: row[VO_COLS.TEXT - 1],
    performer: row[VO_COLS.PERFORMER - 1],
    language: row[VO_COLS.LANGUAGE - 1],
    studio: row[VO_COLS.STUDIO - 1],
    status: row[VO_COLS.STATUS - 1],
    duration: row[VO_COLS.DURATION - 1],
    fileLink: row[VO_COLS.FILE_LINK - 1],
    notes: row[VO_COLS.NOTES - 1],
    createdAt: row[VO_COLS.CREATED_AT - 1],
    updatedAt: row[VO_COLS.UPDATED_AT - 1]
  })).filter(vo => vo.code);
}

/**
 * الحصول على مقاطع التعليق الصوتي لمشروع معين
 * @param {string} projectCode كود المشروع أو اسمه
 * @returns {Array} مصفوفة المقاطع
 */
function getVoiceOverByProject(projectCode) {
  if (!projectCode) return [];

  const allVO = getAllVoiceOver();
  return allVO.filter(vo =>
    vo.project === projectCode ||
    vo.project.includes(projectCode)
  );
}

/**
 * الحصول على مقاطع التعليق الصوتي حسب النوع
 * @param {string} type النوع (راوي/اقتباس/دوبلاج)
 * @returns {Array} مصفوفة المقاطع
 */
function getVoiceOverByType(type) {
  if (!type) return [];

  const allVO = getAllVoiceOver();
  return allVO.filter(vo => vo.type === type || vo.type.includes(type));
}

/**
 * الحصول على مقطع تعليق صوتي بالكود
 * @param {string} code كود المقطع
 * @returns {Object|null} كائن المقطع أو null
 */
function getVoiceOverByCode(code) {
  if (!code) return null;

  const allVO = getAllVoiceOver();
  return allVO.find(vo => vo.code === code) || null;
}

/**
 * الحصول على المقاطع التي لم تُسجل بعد
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Array} مصفوفة المقاطع المعلقة
 */
function getPendingVoiceOver(projectCode) {
  let voiceOvers = projectCode ? getVoiceOverByProject(projectCode) : getAllVoiceOver();

  const pendingStatuses = [
    VO_STATUS.SCRIPT_PENDING,
    VO_STATUS.SCRIPT_READY,
    VO_STATUS.RECORDING_SCHEDULED
  ];

  return voiceOvers.filter(vo => pendingStatuses.includes(vo.status));
}

/**
 * الحصول على المقاطع الجارية
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Array} مصفوفة المقاطع الجارية
 */
function getInProgressVoiceOver(projectCode) {
  let voiceOvers = projectCode ? getVoiceOverByProject(projectCode) : getAllVoiceOver();

  const inProgressStatuses = [
    VO_STATUS.RECORDING_IN_PROGRESS,
    VO_STATUS.EDITING
  ];

  return voiceOvers.filter(vo => inProgressStatuses.includes(vo.status));
}

/**
 * الحصول على المقاطع المكتملة
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Array} مصفوفة المقاطع المكتملة
 */
function getCompletedVoiceOver(projectCode) {
  let voiceOvers = projectCode ? getVoiceOverByProject(projectCode) : getAllVoiceOver();
  return voiceOvers.filter(vo => vo.status === VO_STATUS.COMPLETED);
}

/**
 * الحصول على إجمالي مدة التعليق الصوتي لمشروع
 * @param {string} projectCode كود المشروع
 * @returns {Object} كائن يحتوي على المدة الكلية والمكتملة
 */
function getTotalVoiceOverDuration(projectCode) {
  const voiceOvers = getVoiceOverByProject(projectCode);

  let totalDuration = 0;
  let completedDuration = 0;

  voiceOvers.forEach(vo => {
    const duration = parseFloat(vo.duration) || 0;
    totalDuration += duration;

    if (vo.status === VO_STATUS.COMPLETED) {
      completedDuration += duration;
    }
  });

  return {
    total: totalDuration,
    completed: completedDuration,
    remaining: totalDuration - completedDuration,
    percentage: totalDuration > 0 ? Math.round((completedDuration / totalDuration) * 100) : 0
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إضافة وتعديل التعليق الصوتي
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة مقطع تعليق صوتي جديد
 * @param {Object} voData بيانات المقطع
 * @returns {boolean} نجاح العملية
 */
function addVoiceOver(voData) {
  try {
    const sheet = getSheet(SHEETS.VOICEOVER);
    if (!sheet) {
      showError('شيت التعليق الصوتي غير موجود');
      return false;
    }

    // التحقق من البيانات المطلوبة
    if (!voData.project) {
      showError('المشروع مطلوب');
      return false;
    }

    if (!voData.type) {
      showError('نوع التعليق الصوتي مطلوب');
      return false;
    }

    // إنشاء الكود تلقائياً
    const code = voData.code || generateVOCode();

    // تجهيز صف البيانات
    const rowData = [
      code,
      voData.project,
      voData.type,
      voData.segmentNum || 1,
      voData.text || '',
      voData.performer || '',
      voData.language || VO_LANGUAGES[0],
      voData.studio || '',
      voData.status || VO_STATUS.SCRIPT_PENDING,
      voData.duration || 0,
      voData.fileLink || '',
      voData.notes || '',
      new Date(),
      new Date()
    ];

    sheet.appendRow(rowData);

    // تلوين الحالة
    const lastRow = sheet.getLastRow();
    applyVOStatusColor(sheet, lastRow, voData.status || VO_STATUS.SCRIPT_PENDING);

    return true;
  } catch (error) {
    console.error('Error adding voice over:', error);
    showError('حدث خطأ أثناء إضافة التعليق الصوتي');
    return false;
  }
}

/**
 * تحديث مقطع تعليق صوتي
 * @param {string} code كود المقطع
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updateVoiceOver(code, updates) {
  try {
    const sheet = getSheet(SHEETS.VOICEOVER);
    if (!sheet) return false;

    const rowIndex = findRowByValue(SHEETS.VOICEOVER, VO_COLS.CODE, code);
    if (rowIndex === -1) {
      showError('المقطع غير موجود');
      return false;
    }

    // تحديث الحقول
    const fieldsMap = {
      project: VO_COLS.PROJECT,
      type: VO_COLS.TYPE,
      segmentNum: VO_COLS.SEGMENT_NUM,
      text: VO_COLS.TEXT,
      performer: VO_COLS.PERFORMER,
      language: VO_COLS.LANGUAGE,
      studio: VO_COLS.STUDIO,
      status: VO_COLS.STATUS,
      duration: VO_COLS.DURATION,
      fileLink: VO_COLS.FILE_LINK,
      notes: VO_COLS.NOTES
    };

    Object.keys(updates).forEach(field => {
      if (fieldsMap[field] && updates[field] !== undefined) {
        sheet.getRange(rowIndex, fieldsMap[field]).setValue(updates[field]);

        // تلوين الحالة إذا تم تحديثها
        if (field === 'status') {
          applyVOStatusColor(sheet, rowIndex, updates[field]);
        }
      }
    });

    // تحديث تاريخ آخر تعديل
    sheet.getRange(rowIndex, VO_COLS.UPDATED_AT).setValue(new Date());

    return true;
  } catch (error) {
    console.error('Error updating voice over:', error);
    return false;
  }
}

/**
 * تطبيق لون الحالة على صف التعليق الصوتي
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {string} status الحالة
 */
function applyVOStatusColor(sheet, row, status) {
  const statusColors = {
    [VO_STATUS.SCRIPT_PENDING]: '#FFCDD2',      // أحمر فاتح
    [VO_STATUS.SCRIPT_READY]: '#FFF9C4',        // أصفر فاتح
    [VO_STATUS.RECORDING_SCHEDULED]: '#BBDEFB', // أزرق فاتح
    [VO_STATUS.RECORDING_IN_PROGRESS]: '#B3E5FC', // أزرق سماوي
    [VO_STATUS.EDITING]: '#FFE0B2',             // برتقالي فاتح
    [VO_STATUS.COMPLETED]: '#C8E6C9',           // أخضر فاتح
    [VO_STATUS.CANCELLED]: '#BDBDBD'            // رمادي
  };

  const color = statusColors[status] || '#FFFFFF';
  sheet.getRange(row, VO_COLS.STATUS).setBackground(color);
}

/**
 * إنشاء كود تعليق صوتي جديد
 * @returns {string} الكود الجديد
 */
function generateVOCode() {
  const sheet = getSheet(SHEETS.VOICEOVER);
  if (!sheet) return 'VO-001';

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'VO-001';

  const codes = sheet.getRange(2, VO_COLS.CODE, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(code => code && code.toString().startsWith('VO-'));

  if (codes.length === 0) return 'VO-001';

  const numbers = codes.map(code => parseInt(code.toString().replace('VO-', ''), 10));
  const maxNum = Math.max(...numbers);

  return 'VO-' + (maxNum + 1).toString().padStart(3, '0');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإحصائيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على إحصائيات التعليق الصوتي
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Object} كائن الإحصائيات
 */
function getVoiceOverStats(projectCode) {
  const voiceOvers = projectCode ? getVoiceOverByProject(projectCode) : getAllVoiceOver();

  const stats = {
    total: voiceOvers.length,
    byType: {},
    byStatus: {},
    byLanguage: {},
    duration: getTotalVoiceOverDuration(projectCode || ''),
    pending: 0,
    inProgress: 0,
    completed: 0
  };

  voiceOvers.forEach(vo => {
    // حسب النوع
    stats.byType[vo.type] = (stats.byType[vo.type] || 0) + 1;

    // حسب الحالة
    stats.byStatus[vo.status] = (stats.byStatus[vo.status] || 0) + 1;

    // حسب اللغة
    stats.byLanguage[vo.language] = (stats.byLanguage[vo.language] || 0) + 1;

    // تصنيف عام
    if ([VO_STATUS.SCRIPT_PENDING, VO_STATUS.SCRIPT_READY, VO_STATUS.RECORDING_SCHEDULED].includes(vo.status)) {
      stats.pending++;
    } else if ([VO_STATUS.RECORDING_IN_PROGRESS, VO_STATUS.EDITING].includes(vo.status)) {
      stats.inProgress++;
    } else if (vo.status === VO_STATUS.COMPLETED) {
      stats.completed++;
    }
  });

  return stats;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال واجهة المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * فتح نموذج إضافة تعليق صوتي
 */
function showAddVoiceOverDialog() {
  const activeProjects = getActiveProjects();
  const projectOptions = activeProjects
    .map(p => `<option value="${p.name}">${p.name}</option>`)
    .join('');

  const typeOptions = Object.values(VO_TYPES)
    .map(t => `<option value="${t}">${t}</option>`)
    .join('');

  const languageOptions = VO_LANGUAGES
    .map(l => `<option value="${l}">${l}</option>`)
    .join('');

  const statusOptions = Object.values(VO_STATUS)
    .map(s => `<option value="${s}">${s}</option>`)
    .join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; direction: rtl; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
      .row { display: flex; gap: 15px; }
      .row .form-group { flex: 1; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px; }
      .btn-primary { background: #1565c0; color: white; }
      .btn-secondary { background: #757575; color: white; }
    </style>

    <h3>إضافة تعليق صوتي جديد</h3>

    <div class="row">
      <div class="form-group">
        <label>المشروع *</label>
        <select id="project" required>
          <option value="">اختر المشروع</option>
          ${projectOptions}
        </select>
      </div>
      <div class="form-group">
        <label>النوع *</label>
        <select id="type" required>${typeOptions}</select>
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>رقم المقطع</label>
        <input type="number" id="segmentNum" value="1" min="1">
      </div>
      <div class="form-group">
        <label>اللغة</label>
        <select id="language">${languageOptions}</select>
      </div>
    </div>

    <div class="form-group">
      <label>النص/الوصف</label>
      <textarea id="text" rows="3"></textarea>
    </div>

    <div class="row">
      <div class="form-group">
        <label>المؤدي</label>
        <input type="text" id="performer">
      </div>
      <div class="form-group">
        <label>الاستوديو</label>
        <input type="text" id="studio">
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>الحالة</label>
        <select id="status">${statusOptions}</select>
      </div>
      <div class="form-group">
        <label>المدة (ثانية)</label>
        <input type="number" id="duration" value="0" min="0">
      </div>
    </div>

    <div class="form-group">
      <label>رابط الملف</label>
      <input type="url" id="fileLink">
    </div>

    <div class="form-group">
      <label>ملاحظات</label>
      <textarea id="notes" rows="2"></textarea>
    </div>

    <button class="btn btn-primary" onclick="submitForm()">حفظ</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>

    <script>
      function submitForm() {
        const data = {
          project: document.getElementById('project').value,
          type: document.getElementById('type').value,
          segmentNum: parseInt(document.getElementById('segmentNum').value) || 1,
          text: document.getElementById('text').value,
          performer: document.getElementById('performer').value,
          language: document.getElementById('language').value,
          studio: document.getElementById('studio').value,
          status: document.getElementById('status').value,
          duration: parseFloat(document.getElementById('duration').value) || 0,
          fileLink: document.getElementById('fileLink').value,
          notes: document.getElementById('notes').value
        };

        if (!data.project || !data.type) {
          alert('المشروع والنوع مطلوبان');
          return;
        }

        google.script.run
          .withSuccessHandler(function() {
            alert('تم إضافة التعليق الصوتي بنجاح');
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            alert('حدث خطأ: ' + err.message);
          })
          .addVoiceOver(data);
      }
    </script>
  `).setWidth(500).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة تعليق صوتي');
}

/**
 * عرض إحصائيات التعليق الصوتي
 */
function showVoiceOverStats() {
  const stats = getVoiceOverStats();

  const typeList = Object.entries(stats.byType)
    .map(([type, count]) => `  • ${type}: ${count}`)
    .join('\n');

  const statusList = Object.entries(stats.byStatus)
    .map(([status, count]) => `  • ${status}: ${count}`)
    .join('\n');

  const message = `
إحصائيات التعليق الصوتي
━━━━━━━━━━━━━━━━━━━━━━━━━

📊 إجمالي المقاطع: ${stats.total}

حسب الحالة:
• ⏳ في الانتظار: ${stats.pending}
• 🔄 جاري التنفيذ: ${stats.inProgress}
• ✅ مكتمل: ${stats.completed}

حسب النوع:
${typeList}

حسب الحالة التفصيلية:
${statusList}

⏱️ المدة الكلية: ${Math.floor(stats.duration.total / 60)} دقيقة
✅ المدة المكتملة: ${Math.floor(stats.duration.completed / 60)} دقيقة
📈 نسبة الإنجاز: ${stats.duration.percentage}%
  `.trim();

  showInfo(message, 'إحصائيات التعليق الصوتي');
}
