/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف إدارة الضيوف
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على جميع الدوال المتعلقة بإدارة الضيوف والمقابلات
 */

// ═══════════════════════════════════════════════════════════════════════════════
// ثوابت أعمدة شيت الضيوف
// ═══════════════════════════════════════════════════════════════════════════════

const GUEST_COLS = {
  CODE: 1,               // كود الضيف
  NAME: 2,               // الاسم
  PROJECT: 3,            // المشروع
  PARTICIPATION_TYPE: 4, // نوع المشاركة (مقابلة/دراما)
  CONTACT_STATUS: 5,     // حالة التواصل
  QUESTIONS_STATUS: 6,   // حالة الأسئلة
  COUNTRY: 7,            // البلد
  LOCATION: 8,           // مكان التصوير
  SHOOT_DATE: 9,         // تاريخ التصوير
  SHOOT_STATUS: 10,      // حالة التصوير
  NEEDS_DUBBING: 11,     // يحتاج دوبلاج
  EMAIL: 12,             // البريد الإلكتروني
  PHONE: 13,             // رقم الهاتف
  NOTES: 14,             // ملاحظات
  CREATED_AT: 15,        // تاريخ الإنشاء
  UPDATED_AT: 16         // آخر تحديث
};

// أنواع المشاركة
const PARTICIPATION_TYPES = [
  'مقابلة',
  'دراما',
  'تعليق صوتي',
  'أخرى'
];

// حالات التواصل
const CONTACT_STATUS = {
  NOT_STARTED: 'لم يبدأ',
  IN_PROGRESS: 'جاري التواصل',
  CONFIRMED: 'تم التأكيد',
  REFUSED: 'رفض',
  POSTPONED: 'مؤجل'
};

// حالات الأسئلة
const QUESTIONS_STATUS = {
  NOT_SENT: 'لم ترسل',
  SENT: 'أُرسلت',
  REPLIED: 'تم الرد'
};

// حالات التصوير
const SHOOT_STATUS = {
  NOT_DONE: 'لم يتم',
  SCHEDULED: 'مجدول',
  DONE: 'تم',
  CANCELLED: 'ملغي'
};

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الاستعلام عن الضيوف
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على جميع الضيوف
 * @returns {Array} مصفوفة من كائنات الضيوف
 */
function getAllGuests() {
  const sheet = getSheet(SHEETS.GUESTS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

  return data.map(row => ({
    code: row[GUEST_COLS.CODE - 1],
    name: row[GUEST_COLS.NAME - 1],
    project: row[GUEST_COLS.PROJECT - 1],
    participationType: row[GUEST_COLS.PARTICIPATION_TYPE - 1],
    contactStatus: row[GUEST_COLS.CONTACT_STATUS - 1],
    questionsStatus: row[GUEST_COLS.QUESTIONS_STATUS - 1],
    country: row[GUEST_COLS.COUNTRY - 1],
    location: row[GUEST_COLS.LOCATION - 1],
    shootDate: row[GUEST_COLS.SHOOT_DATE - 1],
    shootStatus: row[GUEST_COLS.SHOOT_STATUS - 1],
    needsDubbing: row[GUEST_COLS.NEEDS_DUBBING - 1],
    email: row[GUEST_COLS.EMAIL - 1],
    phone: row[GUEST_COLS.PHONE - 1],
    notes: row[GUEST_COLS.NOTES - 1],
    createdAt: row[GUEST_COLS.CREATED_AT - 1],
    updatedAt: row[GUEST_COLS.UPDATED_AT - 1]
  })).filter(guest => guest.code);
}

/**
 * الحصول على ضيوف مشروع معين
 * @param {string} projectCode كود المشروع أو اسمه
 * @returns {Array} مصفوفة الضيوف
 */
function getGuestsByProject(projectCode) {
  if (!projectCode) return [];

  const allGuests = getAllGuests();
  return allGuests.filter(guest =>
    guest.project === projectCode ||
    guest.project.includes(projectCode)
  );
}

/**
 * الحصول على ضيف بالكود
 * @param {string} code كود الضيف
 * @returns {Object|null} كائن الضيف أو null
 */
function getGuestByCode(code) {
  if (!code) return null;

  const allGuests = getAllGuests();
  return allGuests.find(guest => guest.code === code) || null;
}

/**
 * الحصول على الضيوف الذين يحتاجون متابعة
 * الضيوف الذين:
 * - حالة التواصل "جاري التواصل" أو "لم يبدأ"
 * - أو حالة الأسئلة "لم ترسل" أو "أُرسلت" (لم يتم الرد)
 * - أو حالة التصوير "مجدول" مع اقتراب الموعد
 * @returns {Array} مصفوفة الضيوف المحتاجين متابعة
 */
function getGuestsNeedingFollowup() {
  const allGuests = getAllGuests();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return allGuests.filter(guest => {
    // التحقق من حالة التواصل
    const needsContactFollowup =
      guest.contactStatus === CONTACT_STATUS.NOT_STARTED ||
      guest.contactStatus === CONTACT_STATUS.IN_PROGRESS;

    // التحقق من حالة الأسئلة
    const needsQuestionsFollowup =
      guest.contactStatus === CONTACT_STATUS.CONFIRMED &&
      (guest.questionsStatus === QUESTIONS_STATUS.NOT_SENT ||
        guest.questionsStatus === QUESTIONS_STATUS.SENT);

    // التحقق من اقتراب موعد التصوير
    let needsShootFollowup = false;
    if (guest.shootStatus === SHOOT_STATUS.SCHEDULED && guest.shootDate) {
      const shootDate = new Date(guest.shootDate);
      shootDate.setHours(0, 0, 0, 0);
      const daysUntilShoot = Math.ceil((shootDate - today) / (1000 * 60 * 60 * 24));
      needsShootFollowup = daysUntilShoot <= 7 && daysUntilShoot >= 0;
    }

    return needsContactFollowup || needsQuestionsFollowup || needsShootFollowup;
  });
}

/**
 * الحصول على الضيوف الذين يحتاجون دوبلاج
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Array} مصفوفة الضيوف
 */
function getGuestsNeedingDubbing(projectCode) {
  let guests = projectCode ? getGuestsByProject(projectCode) : getAllGuests();

  return guests.filter(guest =>
    guest.needsDubbing === 'نعم' || guest.needsDubbing === true
  );
}

/**
 * الحصول على الضيوف حسب حالة التواصل
 * @param {string} status حالة التواصل
 * @returns {Array} مصفوفة الضيوف
 */
function getGuestsByContactStatus(status) {
  const allGuests = getAllGuests();
  return allGuests.filter(guest => guest.contactStatus === status);
}

/**
 * الحصول على الضيوف حسب حالة التصوير
 * @param {string} status حالة التصوير
 * @returns {Array} مصفوفة الضيوف
 */
function getGuestsByShootStatus(status) {
  const allGuests = getAllGuests();
  return allGuests.filter(guest => guest.shootStatus === status);
}

/**
 * الحصول على التصويرات المجدولة قريباً
 * @param {number} days عدد الأيام القادمة (افتراضي 7)
 * @returns {Array} مصفوفة الضيوف
 */
function getUpcomingShoots(days) {
  days = days || 7;
  const allGuests = getAllGuests();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return allGuests.filter(guest => {
    if (guest.shootStatus !== SHOOT_STATUS.SCHEDULED || !guest.shootDate) return false;

    const shootDate = new Date(guest.shootDate);
    shootDate.setHours(0, 0, 0, 0);

    const daysUntil = Math.ceil((shootDate - today) / (1000 * 60 * 60 * 24));
    return daysUntil >= 0 && daysUntil <= days;
  }).sort((a, b) => new Date(a.shootDate) - new Date(b.shootDate));
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إضافة وتعديل الضيوف
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة ضيف جديد
 * @param {Object} guestData بيانات الضيف
 * @returns {boolean} نجاح العملية
 */
function addGuest(guestData) {
  try {
    const sheet = getSheet(SHEETS.GUESTS);
    if (!sheet) {
      showError('شيت الضيوف غير موجود');
      return false;
    }

    // التحقق من البيانات المطلوبة
    if (!guestData.name) {
      showError('اسم الضيف مطلوب');
      return false;
    }

    if (!guestData.project) {
      showError('المشروع مطلوب');
      return false;
    }

    // إنشاء كود الضيف تلقائياً
    const code = guestData.code || generateGuestCode();

    // تجهيز صف البيانات
    const rowData = [
      code,
      guestData.name,
      guestData.project,
      guestData.participationType || PARTICIPATION_TYPES[0],
      guestData.contactStatus || CONTACT_STATUS.NOT_STARTED,
      guestData.questionsStatus || QUESTIONS_STATUS.NOT_SENT,
      guestData.country || '',
      guestData.location || '',
      guestData.shootDate || '',
      guestData.shootStatus || SHOOT_STATUS.NOT_DONE,
      guestData.needsDubbing || 'لا',
      guestData.email || '',
      guestData.phone || '',
      guestData.notes || '',
      new Date(),
      new Date()
    ];

    sheet.appendRow(rowData);

    return true;
  } catch (error) {
    console.error('Error adding guest:', error);
    showError('حدث خطأ أثناء إضافة الضيف');
    return false;
  }
}

/**
 * تحديث ضيف
 * @param {string} code كود الضيف
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updateGuest(code, updates) {
  try {
    const sheet = getSheet(SHEETS.GUESTS);
    if (!sheet) return false;

    const rowIndex = findRowByValue(SHEETS.GUESTS, GUEST_COLS.CODE, code);
    if (rowIndex === -1) {
      showError('الضيف غير موجود');
      return false;
    }

    // تحديث الحقول
    const fieldsMap = {
      name: GUEST_COLS.NAME,
      project: GUEST_COLS.PROJECT,
      participationType: GUEST_COLS.PARTICIPATION_TYPE,
      contactStatus: GUEST_COLS.CONTACT_STATUS,
      questionsStatus: GUEST_COLS.QUESTIONS_STATUS,
      country: GUEST_COLS.COUNTRY,
      location: GUEST_COLS.LOCATION,
      shootDate: GUEST_COLS.SHOOT_DATE,
      shootStatus: GUEST_COLS.SHOOT_STATUS,
      needsDubbing: GUEST_COLS.NEEDS_DUBBING,
      email: GUEST_COLS.EMAIL,
      phone: GUEST_COLS.PHONE,
      notes: GUEST_COLS.NOTES
    };

    Object.keys(updates).forEach(field => {
      if (fieldsMap[field] && updates[field] !== undefined) {
        sheet.getRange(rowIndex, fieldsMap[field]).setValue(updates[field]);
      }
    });

    // تحديث تاريخ آخر تعديل
    sheet.getRange(rowIndex, GUEST_COLS.UPDATED_AT).setValue(new Date());

    return true;
  } catch (error) {
    console.error('Error updating guest:', error);
    return false;
  }
}

/**
 * تحديث حالة التواصل للضيف
 * @param {string} code كود الضيف
 * @param {string} newStatus الحالة الجديدة
 * @returns {boolean} نجاح العملية
 */
function updateGuestContactStatus(code, newStatus) {
  return updateGuest(code, { contactStatus: newStatus });
}

/**
 * تحديث حالة الأسئلة للضيف
 * @param {string} code كود الضيف
 * @param {string} newStatus الحالة الجديدة
 * @returns {boolean} نجاح العملية
 */
function updateGuestQuestionsStatus(code, newStatus) {
  return updateGuest(code, { questionsStatus: newStatus });
}

/**
 * تحديث حالة التصوير للضيف
 * @param {string} code كود الضيف
 * @param {string} newStatus الحالة الجديدة
 * @returns {boolean} نجاح العملية
 */
function updateGuestShootStatus(code, newStatus) {
  return updateGuest(code, { shootStatus: newStatus });
}

/**
 * إنشاء كود ضيف جديد
 * @returns {string} كود الضيف
 */
function generateGuestCode() {
  const sheet = getSheet(SHEETS.GUESTS);
  if (!sheet) return 'G001';

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'G001';

  const codes = sheet.getRange(2, GUEST_COLS.CODE, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(code => code && code.startsWith('G'));

  if (codes.length === 0) return 'G001';

  const numbers = codes.map(code => parseInt(code.replace('G', ''), 10));
  const maxNum = Math.max(...numbers);

  return 'G' + (maxNum + 1).toString().padStart(3, '0');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإحصائيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على إحصائيات الضيوف
 * @param {string} projectCode كود المشروع (اختياري)
 * @returns {Object} كائن الإحصائيات
 */
function getGuestsStats(projectCode) {
  const guests = projectCode ? getGuestsByProject(projectCode) : getAllGuests();

  const stats = {
    total: guests.length,
    byContactStatus: {},
    byQuestionsStatus: {},
    byShootStatus: {},
    byParticipationType: {},
    needsDubbing: 0,
    needsFollowup: 0
  };

  guests.forEach(guest => {
    // حسب حالة التواصل
    const contactStatus = guest.contactStatus || 'غير محدد';
    stats.byContactStatus[contactStatus] = (stats.byContactStatus[contactStatus] || 0) + 1;

    // حسب حالة الأسئلة
    const questionsStatus = guest.questionsStatus || 'غير محدد';
    stats.byQuestionsStatus[questionsStatus] = (stats.byQuestionsStatus[questionsStatus] || 0) + 1;

    // حسب حالة التصوير
    const shootStatus = guest.shootStatus || 'غير محدد';
    stats.byShootStatus[shootStatus] = (stats.byShootStatus[shootStatus] || 0) + 1;

    // حسب نوع المشاركة
    const participationType = guest.participationType || 'غير محدد';
    stats.byParticipationType[participationType] = (stats.byParticipationType[participationType] || 0) + 1;

    // يحتاج دوبلاج
    if (guest.needsDubbing === 'نعم' || guest.needsDubbing === true) {
      stats.needsDubbing++;
    }
  });

  // عدد المحتاجين متابعة
  stats.needsFollowup = getGuestsNeedingFollowup().length;

  return stats;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال واجهة المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * فتح نموذج إضافة ضيف جديد
 */
function showAddGuestDialog() {
  // الحصول على المشاريع النشطة
  const activeProjects = getActiveProjects();
  const projectOptions = activeProjects
    .map(p => `<option value="${p.name}">${p.name}</option>`)
    .join('');

  const participationOptions = PARTICIPATION_TYPES
    .map(t => `<option value="${t}">${t}</option>`)
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

    <h3>إضافة ضيف جديد</h3>

    <div class="form-group">
      <label>الاسم *</label>
      <input type="text" id="guestName" required>
    </div>

    <div class="row">
      <div class="form-group">
        <label>المشروع *</label>
        <select id="guestProject" required>
          <option value="">اختر المشروع</option>
          ${projectOptions}
        </select>
      </div>
      <div class="form-group">
        <label>نوع المشاركة</label>
        <select id="participationType">${participationOptions}</select>
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>البلد</label>
        <input type="text" id="country">
      </div>
      <div class="form-group">
        <label>مكان التصوير</label>
        <input type="text" id="location">
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>تاريخ التصوير</label>
        <input type="date" id="shootDate">
      </div>
      <div class="form-group">
        <label>يحتاج دوبلاج</label>
        <select id="needsDubbing">
          <option value="لا">لا</option>
          <option value="نعم">نعم</option>
        </select>
      </div>
    </div>

    <div class="row">
      <div class="form-group">
        <label>البريد الإلكتروني</label>
        <input type="email" id="email">
      </div>
      <div class="form-group">
        <label>رقم الهاتف</label>
        <input type="tel" id="phone">
      </div>
    </div>

    <div class="form-group">
      <label>ملاحظات</label>
      <textarea id="notes" rows="3"></textarea>
    </div>

    <button class="btn btn-primary" onclick="submitForm()">حفظ</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>

    <script>
      function submitForm() {
        const data = {
          name: document.getElementById('guestName').value,
          project: document.getElementById('guestProject').value,
          participationType: document.getElementById('participationType').value,
          country: document.getElementById('country').value,
          location: document.getElementById('location').value,
          shootDate: document.getElementById('shootDate').value,
          needsDubbing: document.getElementById('needsDubbing').value,
          email: document.getElementById('email').value,
          phone: document.getElementById('phone').value,
          notes: document.getElementById('notes').value
        };

        if (!data.name || !data.project) {
          alert('الاسم والمشروع مطلوبان');
          return;
        }

        google.script.run
          .withSuccessHandler(() => {
            alert('تم إضافة الضيف بنجاح');
            google.script.host.close();
          })
          .withFailureHandler((err) => {
            alert('حدث خطأ: ' + err.message);
          })
          .addGuest(data);
      }
    </script>
  `).setWidth(500).setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة ضيف جديد');
}

/**
 * عرض الضيوف المحتاجين متابعة
 */
function showGuestsNeedingFollowup() {
  const guests = getGuestsNeedingFollowup();

  if (guests.length === 0) {
    showInfo('لا يوجد ضيوف يحتاجون متابعة حالياً 🎉', 'المتابعة');
    return;
  }

  const list = guests.map(g => {
    let reason = '';
    if (g.contactStatus === CONTACT_STATUS.NOT_STARTED) reason = '(لم يبدأ التواصل)';
    else if (g.contactStatus === CONTACT_STATUS.IN_PROGRESS) reason = '(جاري التواصل)';
    else if (g.questionsStatus === QUESTIONS_STATUS.NOT_SENT) reason = '(الأسئلة لم ترسل)';
    else if (g.questionsStatus === QUESTIONS_STATUS.SENT) reason = '(بانتظار الرد)';
    else if (g.shootStatus === SHOOT_STATUS.SCHEDULED) reason = '(تصوير قريب)';

    return `• ${g.name} - ${g.project} ${reason}`;
  }).join('\n');

  const message = `
الضيوف المحتاجين متابعة (${guests.length})
━━━━━━━━━━━━━━━━━━━━━━━━━

${list}
  `.trim();

  showInfo(message, 'متابعة الضيوف');
}

/**
 * عرض التصويرات القادمة
 */
function showUpcomingShoots() {
  const guests = getUpcomingShoots(14); // الأسبوعين القادمين

  if (guests.length === 0) {
    showInfo('لا توجد تصويرات مجدولة في الأسبوعين القادمين', 'التصويرات القادمة');
    return;
  }

  const list = guests.map(g => {
    const shootDate = new Date(g.shootDate);
    const dateStr = Utilities.formatDate(shootDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const daysRemaining = Math.ceil((shootDate - new Date()) / (1000 * 60 * 60 * 24));
    return `• ${dateStr} - ${g.name} (${g.project}) - بعد ${daysRemaining} يوم`;
  }).join('\n');

  const message = `
التصويرات القادمة (${guests.length})
━━━━━━━━━━━━━━━━━━━━━━━━━

${list}
  `.trim();

  showInfo(message, 'التصويرات القادمة');
}
