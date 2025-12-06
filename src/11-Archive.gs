/**
 * ===================================================
 * 11-Archive.gs - إدارة الأرشيف والمواد المرخصة
 * ===================================================
 *
 * هذا الملف يحتوي على:
 * - أنواع الأرشيف (فيديو، صور، وثائق، صوت)
 * - حالات الترخيص
 * - دوال إدارة الأرشيف
 * - توليد أكواد الأرشيف
 */

// ====================================================
// ثوابت الأرشيف
// ====================================================

/**
 * أعمدة ورقة الأرشيف
 */
const ARCHIVE_COLS = {
  CODE: 'الكود',              // AR-001
  PROJECT: 'المشروع',         // قائمة منسدلة من المشاريع النشطة
  TYPE: 'نوع الأرشيف',        // فيديو / صور / وثائق / صوت
  TITLE: 'العنوان',           // عنوان وصفي للمادة
  DESCRIPTION: 'الوصف',       // وصف تفصيلي
  SOURCE: 'المصدر',           // مصدر المادة (أرشيف خاص، مشتراة، إنترنت، إلخ)
  ORIGINAL_DATE: 'تاريخ المادة الأصلي', // تاريخ المادة الأصلية
  OBTAINED_DATE: 'تاريخ الحصول',      // متى حصلنا عليها
  LICENSE_STATUS: 'حالة الترخيص',     // مرخص / في الانتظار / مرفوض / غير مطلوب
  LICENSE_EXPIRY: 'انتهاء الترخيص',   // تاريخ انتهاء الترخيص
  LICENSE_COST: 'تكلفة الترخيص',      // التكلفة إن وجدت
  FILE_LINK: 'رابط الملف',           // رابط الملف في Drive
  THUMBNAIL: 'صورة مصغرة',           // رابط الصورة المصغرة
  USED_IN: 'مستخدم في',              // أين استخدمت هذه المادة
  DURATION: 'المدة',                 // للفيديو والصوت
  RESOLUTION: 'الدقة',               // للفيديو والصور
  NOTES: 'ملاحظات',
  CREATED_BY: 'أضافه',
  CREATED_AT: 'تاريخ الإضافة',
  UPDATED_AT: 'آخر تحديث'
};

/**
 * حالات الترخيص
 */
const LICENSE_STATUS = {
  LICENSED: 'مرخص',
  PENDING: 'في الانتظار',
  REJECTED: 'مرفوض',
  NOT_REQUIRED: 'غير مطلوب',
  EXPIRED: 'منتهي',
  PUBLIC_DOMAIN: 'ملكية عامة'
};

/**
 * ألوان حالات الترخيص
 */
const LICENSE_COLORS = {
  'مرخص': '#00C853',        // أخضر
  'في الانتظار': '#FFD600', // أصفر
  'مرفوض': '#FF1744',       // أحمر
  'غير مطلوب': '#90A4AE',   // رمادي
  'منتهي': '#FF6D00',       // برتقالي
  'ملكية عامة': '#00BCD4'   // سماوي
};

// ====================================================
// دوال استرجاع الأرشيف
// ====================================================

/**
 * الحصول على جميع مواد الأرشيف
 * @returns {Array} مصفوفة بجميع مواد الأرشيف
 */
function getAllArchive() {
  const sheet = getSheet(SHEETS.ARCHIVE);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headers = Object.values(ARCHIVE_COLS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  return data.map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  }).filter(item => item[ARCHIVE_COLS.CODE]); // تجاهل الصفوف الفارغة
}

/**
 * الحصول على أرشيف مشروع معين
 * @param {string} projectCode - كود المشروع
 * @returns {Array} مصفوفة بمواد أرشيف المشروع
 */
function getArchiveByProject(projectCode) {
  const allArchive = getAllArchive();
  return allArchive.filter(item => item[ARCHIVE_COLS.PROJECT] === projectCode);
}

/**
 * الحصول على الأرشيف حسب النوع
 * @param {string} archiveType - نوع الأرشيف
 * @returns {Array} مصفوفة بالمواد من هذا النوع
 */
function getArchiveByType(archiveType) {
  const allArchive = getAllArchive();
  return allArchive.filter(item => item[ARCHIVE_COLS.TYPE] === archiveType);
}

/**
 * الحصول على الأرشيف حسب حالة الترخيص
 * @param {string} status - حالة الترخيص
 * @returns {Array} مصفوفة بالمواد بهذه الحالة
 */
function getArchiveByLicenseStatus(status) {
  const allArchive = getAllArchive();
  return allArchive.filter(item => item[ARCHIVE_COLS.LICENSE_STATUS] === status);
}

/**
 * الحصول على التراخيص المعلقة (في الانتظار)
 * @returns {Array} مصفوفة بالمواد التي تنتظر الترخيص
 */
function getPendingLicenses() {
  return getArchiveByLicenseStatus(LICENSE_STATUS.PENDING);
}

/**
 * الحصول على التراخيص المنتهية أو قريبة الانتهاء
 * @param {number} daysAhead - عدد الأيام للتحذير المسبق (افتراضي 30)
 * @returns {Array} مصفوفة بالمواد التي انتهى أو سينتهي ترخيصها
 */
function getExpiringLicenses(daysAhead = 30) {
  const allArchive = getAllArchive();
  const today = new Date();
  const warningDate = new Date(today.getTime() + (daysAhead * 24 * 60 * 60 * 1000));

  return allArchive.filter(item => {
    const expiryDate = item[ARCHIVE_COLS.LICENSE_EXPIRY];
    if (!expiryDate || !(expiryDate instanceof Date)) return false;

    // التراخيص المنتهية أو التي ستنتهي خلال الفترة المحددة
    return expiryDate <= warningDate;
  }).sort((a, b) => {
    const dateA = new Date(a[ARCHIVE_COLS.LICENSE_EXPIRY]);
    const dateB = new Date(b[ARCHIVE_COLS.LICENSE_EXPIRY]);
    return dateA - dateB;
  });
}

/**
 * الحصول على الأرشيف غير المستخدم
 * @returns {Array} مصفوفة بالمواد غير المستخدمة
 */
function getUnusedArchive() {
  const allArchive = getAllArchive();
  return allArchive.filter(item => !item[ARCHIVE_COLS.USED_IN] || item[ARCHIVE_COLS.USED_IN].toString().trim() === '');
}

/**
 * البحث في الأرشيف
 * @param {string} searchTerm - كلمة البحث
 * @returns {Array} نتائج البحث
 */
function searchArchive(searchTerm) {
  const allArchive = getAllArchive();
  const term = searchTerm.toLowerCase();

  return allArchive.filter(item => {
    return (
      (item[ARCHIVE_COLS.TITLE] && item[ARCHIVE_COLS.TITLE].toLowerCase().includes(term)) ||
      (item[ARCHIVE_COLS.DESCRIPTION] && item[ARCHIVE_COLS.DESCRIPTION].toLowerCase().includes(term)) ||
      (item[ARCHIVE_COLS.SOURCE] && item[ARCHIVE_COLS.SOURCE].toLowerCase().includes(term)) ||
      (item[ARCHIVE_COLS.NOTES] && item[ARCHIVE_COLS.NOTES].toLowerCase().includes(term))
    );
  });
}

// ====================================================
// دوال إضافة وتعديل الأرشيف
// ====================================================

/**
 * توليد كود أرشيف جديد
 * @returns {string} كود جديد بصيغة AR-XXX
 */
function generateArchiveCode() {
  const sheet = getSheet(SHEETS.ARCHIVE);
  if (!sheet) return 'AR-001';

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'AR-001';

  // الحصول على جميع الأكواد الموجودة
  const codes = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  // إيجاد أعلى رقم
  let maxNum = 0;
  codes.forEach(code => {
    if (code && typeof code === 'string' && code.startsWith('AR-')) {
      const num = parseInt(code.replace('AR-', ''));
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  });

  // توليد الكود الجديد
  const newNum = maxNum + 1;
  return `AR-${newNum.toString().padStart(3, '0')}`;
}

/**
 * إضافة مادة أرشيفية جديدة
 * @param {Object} archiveData - بيانات المادة الأرشيفية
 * @returns {Object} نتيجة العملية
 */
function addArchiveItem(archiveData) {
  try {
    const sheet = getSheet(SHEETS.ARCHIVE);
    if (!sheet) {
      return { success: false, error: 'ورقة الأرشيف غير موجودة' };
    }

    const code = generateArchiveCode();
    const now = new Date();
    const headers = Object.values(ARCHIVE_COLS);

    // تحضير صف البيانات
    const rowData = headers.map(header => {
      switch (header) {
        case ARCHIVE_COLS.CODE:
          return code;
        case ARCHIVE_COLS.CREATED_AT:
        case ARCHIVE_COLS.UPDATED_AT:
          return now;
        default:
          return archiveData[header] || '';
      }
    });

    // إضافة الصف
    sheet.appendRow(rowData);

    // تطبيق لون حالة الترخيص
    const newRow = sheet.getLastRow();
    const licenseStatus = archiveData[ARCHIVE_COLS.LICENSE_STATUS];
    if (licenseStatus && LICENSE_COLORS[licenseStatus]) {
      const licenseColIndex = headers.indexOf(ARCHIVE_COLS.LICENSE_STATUS) + 1;
      sheet.getRange(newRow, licenseColIndex).setBackground(LICENSE_COLORS[licenseStatus]);
    }

    return { success: true, code: code };

  } catch (error) {
    console.error('خطأ في إضافة مادة أرشيفية:', error);
    return { success: false, error: error.message };
  }
}

/**
 * تحديث مادة أرشيفية
 * @param {string} code - كود المادة
 * @param {Object} updateData - البيانات المراد تحديثها
 * @returns {Object} نتيجة العملية
 */
function updateArchiveItem(code, updateData) {
  try {
    const sheet = getSheet(SHEETS.ARCHIVE);
    if (!sheet) {
      return { success: false, error: 'ورقة الأرشيف غير موجودة' };
    }

    // البحث عن الصف
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === code) {
        rowIndex = i + 1; // +1 لأن الصفوف تبدأ من 1
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'المادة الأرشيفية غير موجودة' };
    }

    // تحديث البيانات
    Object.keys(updateData).forEach(key => {
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(updateData[key]);
      }
    });

    // تحديث آخر تعديل
    const updatedAtIndex = headers.indexOf(ARCHIVE_COLS.UPDATED_AT);
    if (updatedAtIndex !== -1) {
      sheet.getRange(rowIndex, updatedAtIndex + 1).setValue(new Date());
    }

    // تحديث لون حالة الترخيص إذا تم تغييرها
    if (updateData[ARCHIVE_COLS.LICENSE_STATUS]) {
      const licenseColIndex = headers.indexOf(ARCHIVE_COLS.LICENSE_STATUS);
      const color = LICENSE_COLORS[updateData[ARCHIVE_COLS.LICENSE_STATUS]];
      if (color) {
        sheet.getRange(rowIndex, licenseColIndex + 1).setBackground(color);
      }
    }

    return { success: true };

  } catch (error) {
    console.error('خطأ في تحديث المادة الأرشيفية:', error);
    return { success: false, error: error.message };
  }
}

/**
 * تحديث حالة استخدام مادة أرشيفية
 * @param {string} code - كود المادة
 * @param {string} usedIn - أين استخدمت (يضاف للقائمة)
 * @returns {Object} نتيجة العملية
 */
function markArchiveAsUsed(code, usedIn) {
  try {
    const sheet = getSheet(SHEETS.ARCHIVE);
    if (!sheet) {
      return { success: false, error: 'ورقة الأرشيف غير موجودة' };
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === code) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'المادة الأرشيفية غير موجودة' };
    }

    // الحصول على القيمة الحالية وإضافة الجديدة
    const usedInColIndex = headers.indexOf(ARCHIVE_COLS.USED_IN);
    const currentValue = values[rowIndex - 1][usedInColIndex] || '';
    const newValue = currentValue ? `${currentValue}, ${usedIn}` : usedIn;

    sheet.getRange(rowIndex, usedInColIndex + 1).setValue(newValue);

    // تحديث آخر تعديل
    const updatedAtIndex = headers.indexOf(ARCHIVE_COLS.UPDATED_AT);
    sheet.getRange(rowIndex, updatedAtIndex + 1).setValue(new Date());

    return { success: true };

  } catch (error) {
    console.error('خطأ في تحديث حالة الاستخدام:', error);
    return { success: false, error: error.message };
  }
}

// ====================================================
// دوال الإحصائيات
// ====================================================

/**
 * الحصول على إحصائيات الأرشيف
 * @param {string} projectCode - كود المشروع (اختياري)
 * @returns {Object} إحصائيات الأرشيف
 */
function getArchiveStats(projectCode = null) {
  let archive = projectCode ? getArchiveByProject(projectCode) : getAllArchive();

  const stats = {
    total: archive.length,
    byType: {},
    byLicenseStatus: {},
    totalCost: 0,
    usedCount: 0,
    unusedCount: 0
  };

  // تهيئة الأنواع
  Object.values(ARCHIVE_TYPES).forEach(type => {
    stats.byType[type] = 0;
  });

  // تهيئة حالات الترخيص
  Object.values(LICENSE_STATUS).forEach(status => {
    stats.byLicenseStatus[status] = 0;
  });

  // حساب الإحصائيات
  archive.forEach(item => {
    // حسب النوع
    const type = item[ARCHIVE_COLS.TYPE];
    if (type && stats.byType.hasOwnProperty(type)) {
      stats.byType[type]++;
    }

    // حسب حالة الترخيص
    const licenseStatus = item[ARCHIVE_COLS.LICENSE_STATUS];
    if (licenseStatus && stats.byLicenseStatus.hasOwnProperty(licenseStatus)) {
      stats.byLicenseStatus[licenseStatus]++;
    }

    // التكلفة الإجمالية
    const cost = parseFloat(item[ARCHIVE_COLS.LICENSE_COST]) || 0;
    stats.totalCost += cost;

    // مستخدم / غير مستخدم
    if (item[ARCHIVE_COLS.USED_IN] && item[ARCHIVE_COLS.USED_IN].toString().trim() !== '') {
      stats.usedCount++;
    } else {
      stats.unusedCount++;
    }
  });

  return stats;
}

/**
 * الحصول على إجمالي تكلفة التراخيص لمشروع
 * @param {string} projectCode - كود المشروع
 * @returns {number} إجمالي التكلفة
 */
function getProjectLicenseCost(projectCode) {
  const archive = getArchiveByProject(projectCode);
  return archive.reduce((total, item) => {
    return total + (parseFloat(item[ARCHIVE_COLS.LICENSE_COST]) || 0);
  }, 0);
}

// ====================================================
// دوال التنبيهات
// ====================================================

/**
 * الحصول على تنبيهات الأرشيف
 * @returns {Array} مصفوفة بالتنبيهات
 */
function getArchiveAlerts() {
  const alerts = [];

  // تراخيص في الانتظار
  const pending = getPendingLicenses();
  if (pending.length > 0) {
    alerts.push({
      type: 'warning',
      message: `${pending.length} مادة في انتظار الترخيص`,
      items: pending.map(p => p[ARCHIVE_COLS.CODE])
    });
  }

  // تراخيص منتهية أو ستنتهي قريباً
  const expiring = getExpiringLicenses(30);
  const expired = expiring.filter(item => new Date(item[ARCHIVE_COLS.LICENSE_EXPIRY]) < new Date());
  const soonExpiring = expiring.filter(item => new Date(item[ARCHIVE_COLS.LICENSE_EXPIRY]) >= new Date());

  if (expired.length > 0) {
    alerts.push({
      type: 'error',
      message: `${expired.length} ترخيص منتهي`,
      items: expired.map(p => p[ARCHIVE_COLS.CODE])
    });
  }

  if (soonExpiring.length > 0) {
    alerts.push({
      type: 'warning',
      message: `${soonExpiring.length} ترخيص سينتهي خلال 30 يوم`,
      items: soonExpiring.map(p => p[ARCHIVE_COLS.CODE])
    });
  }

  return alerts;
}

// ====================================================
// دوال العرض
// ====================================================

/**
 * عرض نموذج إضافة مادة أرشيفية
 */
function showAddArchiveDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; direction: rtl; padding: 15px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
      textarea { height: 80px; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px; }
      .btn-primary { background: #1a73e8; color: white; }
      .btn-secondary { background: #5f6368; color: white; }
      .row { display: flex; gap: 15px; }
      .col { flex: 1; }
    </style>

    <h3>إضافة مادة أرشيفية جديدة</h3>

    <div class="form-group">
      <label>المشروع *</label>
      <select id="project" required></select>
    </div>

    <div class="row">
      <div class="col">
        <div class="form-group">
          <label>نوع الأرشيف *</label>
          <select id="type" required>
            <option value="">اختر النوع</option>
            ${Object.values(ARCHIVE_TYPES).map(t => `<option value="${t}">${t}</option>`).join('')}
          </select>
        </div>
      </div>
      <div class="col">
        <div class="form-group">
          <label>حالة الترخيص *</label>
          <select id="licenseStatus" required>
            <option value="">اختر الحالة</option>
            ${Object.values(LICENSE_STATUS).map(s => `<option value="${s}">${s}</option>`).join('')}
          </select>
        </div>
      </div>
    </div>

    <div class="form-group">
      <label>العنوان *</label>
      <input type="text" id="title" required>
    </div>

    <div class="form-group">
      <label>الوصف</label>
      <textarea id="description"></textarea>
    </div>

    <div class="form-group">
      <label>المصدر</label>
      <input type="text" id="source">
    </div>

    <div class="row">
      <div class="col">
        <div class="form-group">
          <label>تاريخ الحصول</label>
          <input type="date" id="obtainedDate">
        </div>
      </div>
      <div class="col">
        <div class="form-group">
          <label>انتهاء الترخيص</label>
          <input type="date" id="licenseExpiry">
        </div>
      </div>
    </div>

    <div class="form-group">
      <label>رابط الملف</label>
      <input type="url" id="fileLink">
    </div>

    <div class="form-group">
      <label>ملاحظات</label>
      <textarea id="notes"></textarea>
    </div>

    <div style="text-align: left; margin-top: 20px;">
      <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>
      <button class="btn btn-primary" onclick="submitForm()">إضافة</button>
    </div>

    <script>
      // تحميل قائمة المشاريع
      google.script.run.withSuccessHandler(function(projects) {
        const select = document.getElementById('project');
        select.innerHTML = '<option value="">اختر المشروع</option>';
        projects.forEach(function(p) {
          select.innerHTML += '<option value="' + p.code + '">' + p.code + ' - ' + p.name + '</option>';
        });
      }).getActiveProjects();

      function submitForm() {
        const data = {
          '${ARCHIVE_COLS.PROJECT}': document.getElementById('project').value,
          '${ARCHIVE_COLS.TYPE}': document.getElementById('type').value,
          '${ARCHIVE_COLS.LICENSE_STATUS}': document.getElementById('licenseStatus').value,
          '${ARCHIVE_COLS.TITLE}': document.getElementById('title').value,
          '${ARCHIVE_COLS.DESCRIPTION}': document.getElementById('description').value,
          '${ARCHIVE_COLS.SOURCE}': document.getElementById('source').value,
          '${ARCHIVE_COLS.OBTAINED_DATE}': document.getElementById('obtainedDate').value,
          '${ARCHIVE_COLS.LICENSE_EXPIRY}': document.getElementById('licenseExpiry').value,
          '${ARCHIVE_COLS.FILE_LINK}': document.getElementById('fileLink').value,
          '${ARCHIVE_COLS.NOTES}': document.getElementById('notes').value
        };

        if (!data['${ARCHIVE_COLS.PROJECT}'] || !data['${ARCHIVE_COLS.TYPE}'] ||
            !data['${ARCHIVE_COLS.TITLE}'] || !data['${ARCHIVE_COLS.LICENSE_STATUS}']) {
          alert('يرجى ملء جميع الحقول المطلوبة');
          return;
        }

        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              alert('تم إضافة المادة الأرشيفية بنجاح\\nالكود: ' + result.code);
              google.script.host.close();
            } else {
              alert('خطأ: ' + result.error);
            }
          })
          .addArchiveItem(data);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة مادة أرشيفية');
}
