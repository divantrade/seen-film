/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف إدارة الرسوم المتحركة
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على جميع الدوال المتعلقة بإدارة الرسوم المتحركة
 * بما في ذلك Animation, Motion Graphics, Infographics
 */

// ═══════════════════════════════════════════════════════════════════════════════
// ثوابت أعمدة شيت الرسوم المتحركة
// ═══════════════════════════════════════════════════════════════════════════════

const ANIM_COLS = {
  CODE: 1,           // كود المشهد (AN-001)
  PROJECT: 2,        // المشروع
  SCENE_NUM: 3,      // رقم المشهد
  DESCRIPTION: 4,    // وصف المشهد
  TYPE: 5,           // النوع (2D/3D/Motion Graphics)
  DURATION: 6,       // المدة (ثانية)
  SCRIPT_LINK: 7,    // رابط الاسكربت
  STUDIO: 8,         // الاستوديو
  ANIMATOR: 9,       // المحرك/المصمم
  STATUS: 10,        // الحالة
  FILE_LINK: 11,     // رابط الملف النهائي
  NOTES: 12,         // ملاحظات
  CREATED_AT: 13,    // تاريخ الإنشاء
  UPDATED_AT: 14     // آخر تحديث
};

// أنواع الرسوم المتحركة
const ANIM_TYPES = {
  TWO_D: '2D Animation',
  THREE_D: '3D Animation',
  MOTION_GRAPHICS: 'Motion Graphics',
  INFOGRAPHIC: 'Infographic',
  CHARACTER: 'Character Animation',
  WHITEBOARD: 'Whiteboard Animation',
  OTHER: 'أخرى'
};

// حالات الرسوم المتحركة
const ANIM_STATUS = {
  CONCEPT: 'مرحلة الفكرة',
  STORYBOARD: 'Storyboard',
  DESIGN: 'التصميم',
  ANIMATION: 'التحريك',
  REVIEW: 'المراجعة',
  REVISION: 'التعديلات',
  COMPLETED: 'مكتمل',
  CANCELLED: 'ملغي'
};

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الاستعلام عن الرسوم المتحركة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على جميع مشاهد الرسوم المتحركة
 * @returns {Array} مصفوفة من كائنات الرسوم
 */
function getAllAnimation() {
  const sheet = getSheet(SHEETS.ANIMATION);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  return data.map(row => ({
    code: row[ANIM_COLS.CODE - 1],
    project: row[ANIM_COLS.PROJECT - 1],
    sceneNum: row[ANIM_COLS.SCENE_NUM - 1],
    description: row[ANIM_COLS.DESCRIPTION - 1],
    type: row[ANIM_COLS.TYPE - 1],
    duration: row[ANIM_COLS.DURATION - 1],
    scriptLink: row[ANIM_COLS.SCRIPT_LINK - 1],
    studio: row[ANIM_COLS.STUDIO - 1],
    animator: row[ANIM_COLS.ANIMATOR - 1],
    status: row[ANIM_COLS.STATUS - 1],
    fileLink: row[ANIM_COLS.FILE_LINK - 1],
    notes: row[ANIM_COLS.NOTES - 1],
    createdAt: row[ANIM_COLS.CREATED_AT - 1],
    updatedAt: row[ANIM_COLS.UPDATED_AT - 1]
  })).filter(anim => anim.code);
}

/**
 * الحصول على مشاهد الرسوم لمشروع معين
 * @param {string} projectCode كود المشروع أو اسمه
 * @returns {Array} مصفوفة المشاهد
 */
function getAnimationByProject(projectCode) {
  if (!projectCode) return [];

  const allAnim = getAllAnimation();
  return allAnim.filter(anim =>
    anim.project === projectCode ||
    anim.project.includes(projectCode)
  );
}

/**
 * الحصول على مشهد رسوم بالكود
 * @param {string} code كود المشهد
 * @returns {Object|null} كائن المشهد أو null
 */
function getAnimationByCode(code) {
  if (!code) return null;

  const allAnim = getAllAnimation();
  return allAnim.find(anim => anim.code === code) || null;
}

/**
 * الحصول على المشاهد التي لم تُنفذ بعد
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Array} مصفوفة المشاهد المعلقة
 */
function getPendingAnimation(projectCode) {
  let animations = projectCode ? getAnimationByProject(projectCode) : getAllAnimation();

  const pendingStatuses = [
    ANIM_STATUS.CONCEPT,
    ANIM_STATUS.STORYBOARD,
    ANIM_STATUS.DESIGN
  ];

  return animations.filter(anim => pendingStatuses.includes(anim.status));
}

/**
 * الحصول على المشاهد الجارية
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Array} مصفوفة المشاهد الجارية
 */
function getInProgressAnimation(projectCode) {
  let animations = projectCode ? getAnimationByProject(projectCode) : getAllAnimation();

  const inProgressStatuses = [
    ANIM_STATUS.ANIMATION,
    ANIM_STATUS.REVIEW,
    ANIM_STATUS.REVISION
  ];

  return animations.filter(anim => inProgressStatuses.includes(anim.status));
}

/**
 * الحصول على المشاهد المكتملة
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Array} مصفوفة المشاهد المكتملة
 */
function getCompletedAnimation(projectCode) {
  let animations = projectCode ? getAnimationByProject(projectCode) : getAllAnimation();
  return animations.filter(anim => anim.status === ANIM_STATUS.COMPLETED);
}

/**
 * الحصول على إجمالي مدة الرسوم المتحركة لمشروع
 * @param {string} projectCode كود المشروع
 * @returns {Object} كائن يحتوي على المدة الكلية والمكتملة
 */
function getTotalAnimationDuration(projectCode) {
  const animations = getAnimationByProject(projectCode);

  let totalDuration = 0;
  let completedDuration = 0;

  animations.forEach(anim => {
    const duration = parseFloat(anim.duration) || 0;
    totalDuration += duration;

    if (anim.status === ANIM_STATUS.COMPLETED) {
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

/**
 * الحصول على المشاهد حسب النوع
 * @param {string} type نوع الرسوم
 * @returns {Array} مصفوفة المشاهد
 */
function getAnimationByType(type) {
  if (!type) return [];

  const allAnim = getAllAnimation();
  return allAnim.filter(anim => anim.type === type);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إضافة وتعديل الرسوم المتحركة
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة مشهد رسوم متحركة جديد
 * @param {Object} animData بيانات المشهد
 * @returns {boolean} نجاح العملية
 */
function addAnimation(animData) {
  try {
    const sheet = getSheet(SHEETS.ANIMATION);
    if (!sheet) {
      showError('شيت الرسوم المتحركة غير موجود');
      return false;
    }

    // التحقق من البيانات المطلوبة
    if (!animData.project) {
      showError('المشروع مطلوب');
      return false;
    }

    // إنشاء الكود تلقائياً
    const code = animData.code || generateAnimationCode();

    // تجهيز صف البيانات
    const rowData = [
      code,
      animData.project,
      animData.sceneNum || 1,
      animData.description || '',
      animData.type || Object.values(ANIM_TYPES)[0],
      animData.duration || 0,
      animData.scriptLink || '',
      animData.studio || '',
      animData.animator || '',
      animData.status || ANIM_STATUS.CONCEPT,
      animData.fileLink || '',
      animData.notes || '',
      new Date(),
      new Date()
    ];

    sheet.appendRow(rowData);

    // تلوين الحالة
    const lastRow = sheet.getLastRow();
    applyAnimStatusColor(sheet, lastRow, animData.status || ANIM_STATUS.CONCEPT);

    return true;
  } catch (error) {
    console.error('Error adding animation:', error);
    showError('حدث خطأ أثناء إضافة مشهد الرسوم');
    return false;
  }
}

/**
 * تحديث مشهد رسوم متحركة
 * @param {string} code كود المشهد
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updateAnimation(code, updates) {
  try {
    const sheet = getSheet(SHEETS.ANIMATION);
    if (!sheet) return false;

    const rowIndex = findRowByValue(SHEETS.ANIMATION, ANIM_COLS.CODE, code);
    if (rowIndex === -1) {
      showError('المشهد غير موجود');
      return false;
    }

    // تحديث الحقول
    const fieldsMap = {
      project: ANIM_COLS.PROJECT,
      sceneNum: ANIM_COLS.SCENE_NUM,
      description: ANIM_COLS.DESCRIPTION,
      type: ANIM_COLS.TYPE,
      duration: ANIM_COLS.DURATION,
      scriptLink: ANIM_COLS.SCRIPT_LINK,
      studio: ANIM_COLS.STUDIO,
      animator: ANIM_COLS.ANIMATOR,
      status: ANIM_COLS.STATUS,
      fileLink: ANIM_COLS.FILE_LINK,
      notes: ANIM_COLS.NOTES
    };

    Object.keys(updates).forEach(field => {
      if (fieldsMap[field] && updates[field] !== undefined) {
        sheet.getRange(rowIndex, fieldsMap[field]).setValue(updates[field]);

        // تلوين الحالة إذا تم تحديثها
        if (field === 'status') {
          applyAnimStatusColor(sheet, rowIndex, updates[field]);
        }
      }
    });

    // تحديث تاريخ آخر تعديل
    sheet.getRange(rowIndex, ANIM_COLS.UPDATED_AT).setValue(new Date());

    return true;
  } catch (error) {
    console.error('Error updating animation:', error);
    return false;
  }
}

/**
 * تطبيق لون الحالة على صف الرسوم المتحركة
 * @param {Sheet} sheet الشيت
 * @param {number} row رقم الصف
 * @param {string} status الحالة
 */
function applyAnimStatusColor(sheet, row, status) {
  const statusColors = {
    [ANIM_STATUS.CONCEPT]: '#FFCDD2',       // أحمر فاتح
    [ANIM_STATUS.STORYBOARD]: '#F8BBD9',    // وردي
    [ANIM_STATUS.DESIGN]: '#E1BEE7',        // بنفسجي فاتح
    [ANIM_STATUS.ANIMATION]: '#BBDEFB',     // أزرق فاتح
    [ANIM_STATUS.REVIEW]: '#FFF9C4',        // أصفر فاتح
    [ANIM_STATUS.REVISION]: '#FFE0B2',      // برتقالي فاتح
    [ANIM_STATUS.COMPLETED]: '#C8E6C9',     // أخضر فاتح
    [ANIM_STATUS.CANCELLED]: '#BDBDBD'      // رمادي
  };

  const color = statusColors[status] || '#FFFFFF';
  sheet.getRange(row, ANIM_COLS.STATUS).setBackground(color);
}

/**
 * إنشاء كود رسوم متحركة جديد
 * @returns {string} الكود الجديد
 */
function generateAnimationCode() {
  const sheet = getSheet(SHEETS.ANIMATION);
  if (!sheet) return 'AN-001';

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'AN-001';

  const codes = sheet.getRange(2, ANIM_COLS.CODE, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(code => code && code.toString().startsWith('AN-'));

  if (codes.length === 0) return 'AN-001';

  const numbers = codes.map(code => parseInt(code.toString().replace('AN-', ''), 10));
  const maxNum = Math.max(...numbers);

  return 'AN-' + (maxNum + 1).toString().padStart(3, '0');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإحصائيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على إحصائيات الرسوم المتحركة
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Object} كائن الإحصائيات
 */
function getAnimationStats(projectCode) {
  const animations = projectCode ? getAnimationByProject(projectCode) : getAllAnimation();

  const stats = {
    total: animations.length,
    byType: {},
    byStatus: {},
    byAnimator: {},
    duration: projectCode ? getTotalAnimationDuration(projectCode) : { total: 0, completed: 0, percentage: 0 },
    pending: 0,
    inProgress: 0,
    completed: 0
  };

  animations.forEach(anim => {
    // حسب النوع
    stats.byType[anim.type] = (stats.byType[anim.type] || 0) + 1;

    // حسب الحالة
    stats.byStatus[anim.status] = (stats.byStatus[anim.status] || 0) + 1;

    // حسب المحرك
    if (anim.animator) {
      stats.byAnimator[anim.animator] = (stats.byAnimator[anim.animator] || 0) + 1;
    }

    // تصنيف عام
    if ([ANIM_STATUS.CONCEPT, ANIM_STATUS.STORYBOARD, ANIM_STATUS.DESIGN].includes(anim.status)) {
      stats.pending++;
    } else if ([ANIM_STATUS.ANIMATION, ANIM_STATUS.REVIEW, ANIM_STATUS.REVISION].includes(anim.status)) {
      stats.inProgress++;
    } else if (anim.status === ANIM_STATUS.COMPLETED) {
      stats.completed++;
    }
  });

  // حساب المدة الإجمالية إذا لم يكن هناك مشروع محدد
  if (!projectCode) {
    let totalDur = 0, completedDur = 0;
    animations.forEach(anim => {
      const dur = parseFloat(anim.duration) || 0;
      totalDur += dur;
      if (anim.status === ANIM_STATUS.COMPLETED) completedDur += dur;
    });
    stats.duration = {
      total: totalDur,
      completed: completedDur,
      percentage: totalDur > 0 ? Math.round((completedDur / totalDur) * 100) : 0
    };
  }

  return stats;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال واجهة المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * فتح نموذج إضافة مشهد رسوم متحركة
 */
function showAddAnimationDialog() {
  const activeProjects = getActiveProjects();
  const projectOptions = activeProjects
    .map(p => `<option value="${p.name}">${p.name}</option>`)
    .join('');

  const typeOptions = Object.values(ANIM_TYPES)
    .map(t => `<option value="${t}">${t}</option>`)
    .join('');

  const statusOptions = Object.values(ANIM_STATUS)
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

    <h3>إضافة مشهد رسوم متحركة</h3>

    <div class="row">
      <div class="form-group">
        <label>المشروع *</label>
        <select id="project" required>
          <option value="">اختر المشروع</option>
          ${projectOptions}
        </select>
      </div>
      <div class="form-group">
        <label>رقم المشهد</label>
        <input type="number" id="sceneNum" value="1" min="1">
      </div>
    </div>

    <div class="form-group">
      <label>وصف المشهد</label>
      <textarea id="description" rows="3"></textarea>
    </div>

    <div class="row">
      <div class="form-group">
        <label>النوع</label>
        <select id="type">${typeOptions}</select>
      </div>
      <div class="form-group">
        <label>المدة (ثانية)</label>
        <input type="number" id="duration" value="0" min="0">
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>الاستوديو</label>
        <input type="text" id="studio">
      </div>
      <div class="form-group">
        <label>المحرك/المصمم</label>
        <input type="text" id="animator">
      </div>
    </div>

    <div class="form-group">
      <label>رابط الاسكربت</label>
      <input type="url" id="scriptLink">
    </div>

    <div class="row">
      <div class="form-group">
        <label>الحالة</label>
        <select id="status">${statusOptions}</select>
      </div>
      <div class="form-group">
        <label>رابط الملف النهائي</label>
        <input type="url" id="fileLink">
      </div>
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
          sceneNum: parseInt(document.getElementById('sceneNum').value) || 1,
          description: document.getElementById('description').value,
          type: document.getElementById('type').value,
          duration: parseFloat(document.getElementById('duration').value) || 0,
          scriptLink: document.getElementById('scriptLink').value,
          studio: document.getElementById('studio').value,
          animator: document.getElementById('animator').value,
          status: document.getElementById('status').value,
          fileLink: document.getElementById('fileLink').value,
          notes: document.getElementById('notes').value
        };

        if (!data.project) {
          alert('المشروع مطلوب');
          return;
        }

        google.script.run
          .withSuccessHandler(function() {
            alert('تم إضافة مشهد الرسوم بنجاح');
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            alert('حدث خطأ: ' + err.message);
          })
          .addAnimation(data);
      }
    </script>
  `).setWidth(500).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة مشهد رسوم متحركة');
}

/**
 * عرض إحصائيات الرسوم المتحركة
 */
function showAnimationStats() {
  const stats = getAnimationStats();

  const typeList = Object.entries(stats.byType)
    .map(([type, count]) => `  • ${type}: ${count}`)
    .join('\n');

  const message = `
إحصائيات الرسوم المتحركة
━━━━━━━━━━━━━━━━━━━━━━━━━

📊 إجمالي المشاهد: ${stats.total}

حسب الحالة:
• ⏳ في الانتظار: ${stats.pending}
• 🔄 جاري التنفيذ: ${stats.inProgress}
• ✅ مكتمل: ${stats.completed}

حسب النوع:
${typeList}

⏱️ المدة الكلية: ${Math.floor(stats.duration.total)} ثانية (${(stats.duration.total / 60).toFixed(1)} دقيقة)
✅ المدة المكتملة: ${Math.floor(stats.duration.completed)} ثانية
📈 نسبة الإنجاز: ${stats.duration.percentage}%
  `.trim();

  showInfo(message, 'إحصائيات الرسوم المتحركة');
}
