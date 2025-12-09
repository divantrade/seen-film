/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام متابعة الإنتاج الفني - شركة أفلام وثائقية
 * ملف إدارة الفريق والمصورين
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * هذا الملف يحتوي على جميع الدوال المتعلقة بإدارة الفريق والمصورين
 */

// ═══════════════════════════════════════════════════════════════════════════════
// ثوابت أعمدة شيت الفريق
// ═══════════════════════════════════════════════════════════════════════════════

const TEAM_COLS = {
  CODE: 1,           // كود العضو
  NAME: 2,           // الاسم
  ROLE: 3,           // الدور
  EMAIL: 4,          // البريد الإلكتروني
  PHONE: 5,          // رقم الهاتف
  STAGES: 6,         // المراحل المسؤول عنها
  STATUS: 7,         // الحالة (نشط/غير نشط)
  JOIN_DATE: 8,      // تاريخ الانضمام
  NOTES: 9           // ملاحظات
};

// ثوابت أعمدة شيت المصورين
const PHOTOGRAPHER_COLS = {
  CODE: 1,           // كود المصور
  NAME: 2,           // الاسم
  SPECIALIZATION: 3, // التخصص
  EMAIL: 4,          // البريد الإلكتروني
  PHONE: 5,          // رقم الهاتف
  EQUIPMENT: 6,      // المعدات
  STATUS: 7,         // الحالة
  RATE: 8,           // السعر اليومي
  NOTES: 9           // ملاحظات
};

// تخصصات المصورين
const PHOTOGRAPHER_SPECS = [
  'ميداني',
  'استوديو',
  'درامي',
  'جوي (درون)',
  'متعدد'
];

// حالات الفريق
const MEMBER_STATUS = {
  ACTIVE: 'نشط',
  INACTIVE: 'غير نشط'
};

// حالات المصورين
const PHOTOGRAPHER_STATUS = {
  AVAILABLE: 'متاح',
  BUSY: 'مشغول',
  UNAVAILABLE: 'غير متاح'
};

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الاستعلام عن الفريق
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على جميع أعضاء الفريق
 * @returns {Array} مصفوفة من كائنات الأعضاء
 */
function getAllTeamMembers() {
  const sheet = getSheet(SHEETS.TEAM);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  return data.map(row => ({
    code: row[TEAM_COLS.CODE - 1],
    name: row[TEAM_COLS.NAME - 1],
    role: row[TEAM_COLS.ROLE - 1],
    email: row[TEAM_COLS.EMAIL - 1],
    phone: row[TEAM_COLS.PHONE - 1],
    stages: row[TEAM_COLS.STAGES - 1],
    status: row[TEAM_COLS.STATUS - 1],
    joinDate: row[TEAM_COLS.JOIN_DATE - 1],
    notes: row[TEAM_COLS.NOTES - 1]
  })).filter(member => member.code); // تصفية الصفوف الفارغة
}

/**
 * الحصول على أعضاء الفريق النشطين
 * يعتبر الأعضاء نشطين إذا:
 * - الحالة = "نشط"
 * - أو الحالة فارغة (افتراضياً نشط)
 * - أو الحالة ليست "غير نشط"
 * @returns {Array} مصفوفة الأعضاء النشطين
 */
function getTeamMembers() {
  const allMembers = getAllTeamMembers();
  return allMembers.filter(member => {
    const status = (member.status || '').trim();
    // إذا كانت الحالة فارغة، نعتبره نشط
    if (!status) return true;
    // إذا كانت الحالة "نشط" بالضبط
    if (status === MEMBER_STATUS.ACTIVE) return true;
    // استبعاد فقط من حالتهم "غير نشط"
    if (status === MEMBER_STATUS.INACTIVE) return false;
    // أي حالة أخرى تعتبر نشط
    return true;
  });
}

/**
 * الحصول على عضو فريق بالكود
 * @param {string} code كود العضو
 * @returns {Object|null} كائن العضو أو null
 */
function getTeamMemberByCode(code) {
  if (!code) return null;

  const allMembers = getAllTeamMembers();
  return allMembers.find(member => member.code === code) || null;
}

/**
 * الحصول على عضو فريق بالاسم
 * @param {string} name اسم العضو
 * @returns {Object|null} كائن العضو أو null
 */
function getTeamMemberByName(name) {
  if (!name) return null;

  const allMembers = getAllTeamMembers();
  return allMembers.find(member => member.name === name) || null;
}

/**
 * الحصول على أعضاء الفريق حسب الدور
 * @param {string} role الدور المطلوب
 * @returns {Array} مصفوفة الأعضاء المطابقين
 */
function getTeamByRole(role) {
  if (!role) return [];

  const activeMembers = getTeamMembers();
  return activeMembers.filter(member => member.role === role);
}

/**
 * الحصول على أعضاء الفريق حسب المرحلة
 * @param {string} stageName اسم المرحلة
 * @returns {Array} مصفوفة الأعضاء المسؤولين عن المرحلة
 */
function getTeamByStage(stageName) {
  if (!stageName) return [];

  const activeMembers = getTeamMembers();
  return activeMembers.filter(member => {
    if (!member.stages) return false;
    // المراحل مخزنة كنص مفصول بفواصل
    const memberStages = member.stages.split(',').map(s => s.trim());
    return memberStages.includes(stageName);
  });
}

/**
 * الحصول على أسماء أعضاء الفريق النشطين
 * @returns {Array} مصفوفة الأسماء
 */
function getTeamMemberNames() {
  const activeMembers = getTeamMembers();
  return activeMembers.map(member => member.name);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الاستعلام عن المصورين
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على جميع المصورين
 * @returns {Array} مصفوفة من كائنات المصورين
 */
function getAllPhotographers() {
  const sheet = getSheet(SHEETS.PHOTOGRAPHERS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  return data.map(row => ({
    code: row[PHOTOGRAPHER_COLS.CODE - 1],
    name: row[PHOTOGRAPHER_COLS.NAME - 1],
    specialization: row[PHOTOGRAPHER_COLS.SPECIALIZATION - 1],
    email: row[PHOTOGRAPHER_COLS.EMAIL - 1],
    phone: row[PHOTOGRAPHER_COLS.PHONE - 1],
    equipment: row[PHOTOGRAPHER_COLS.EQUIPMENT - 1],
    status: row[PHOTOGRAPHER_COLS.STATUS - 1],
    rate: row[PHOTOGRAPHER_COLS.RATE - 1],
    notes: row[PHOTOGRAPHER_COLS.NOTES - 1]
  })).filter(photographer => photographer.code);
}

/**
 * الحصول على المصورين المتاحين
 * @returns {Array} مصفوفة المصورين المتاحين
 */
function getPhotographers() {
  const allPhotographers = getAllPhotographers();
  return allPhotographers.filter(p => p.status !== PHOTOGRAPHER_STATUS.UNAVAILABLE);
}

/**
 * الحصول على المصورين حسب التخصص
 * @param {string} specialization التخصص
 * @returns {Array} مصفوفة المصورين المطابقين
 */
function getPhotographersBySpecialization(specialization) {
  if (!specialization) return [];

  const availablePhotographers = getPhotographers();
  return availablePhotographers.filter(p =>
    p.specialization === specialization || p.specialization === 'متعدد'
  );
}

/**
 * الحصول على مصور بالكود
 * @param {string} code كود المصور
 * @returns {Object|null} كائن المصور أو null
 */
function getPhotographerByCode(code) {
  if (!code) return null;

  const allPhotographers = getAllPhotographers();
  return allPhotographers.find(p => p.code === code) || null;
}

/**
 * الحصول على أسماء المصورين المتاحين
 * @returns {Array} مصفوفة الأسماء
 */
function getPhotographerNames() {
  const availablePhotographers = getPhotographers();
  return availablePhotographers.map(p => p.name);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إضافة وتعديل الفريق
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة عضو فريق جديد
 * @param {Object} memberData بيانات العضو
 * @returns {boolean} نجاح العملية
 */
function addTeamMember(memberData) {
  try {
    const sheet = getSheet(SHEETS.TEAM);
    if (!sheet) {
      showError('شيت الفريق غير موجود');
      return false;
    }

    // التحقق من البيانات المطلوبة
    if (!memberData.name) {
      showError('اسم العضو مطلوب');
      return false;
    }

    // إنشاء كود العضو تلقائياً
    const code = memberData.code || generateTeamCode();

    // التحقق من عدم تكرار الكود
    if (getTeamMemberByCode(code)) {
      showError('كود العضو موجود مسبقاً');
      return false;
    }

    // تحويل المراحل لنص إذا كانت مصفوفة
    const stages = Array.isArray(memberData.stages)
      ? memberData.stages.join(', ')
      : memberData.stages || '';

    // تجهيز صف البيانات
    const rowData = [
      code,
      memberData.name,
      memberData.role || '',
      memberData.email || '',
      memberData.phone || '',
      stages,
      memberData.status || MEMBER_STATUS.ACTIVE,
      memberData.joinDate || new Date(),
      memberData.notes || ''
    ];

    sheet.appendRow(rowData);

    return true;
  } catch (error) {
    console.error('Error adding team member:', error);
    showError('حدث خطأ أثناء إضافة العضو');
    return false;
  }
}

/**
 * تحديث عضو فريق
 * @param {string} code كود العضو
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updateTeamMember(code, updates) {
  try {
    const sheet = getSheet(SHEETS.TEAM);
    if (!sheet) return false;

    const rowIndex = findRowByValue(SHEETS.TEAM, TEAM_COLS.CODE, code);
    if (rowIndex === -1) {
      showError('العضو غير موجود');
      return false;
    }

    // تحديث الحقول
    if (updates.name !== undefined) {
      sheet.getRange(rowIndex, TEAM_COLS.NAME).setValue(updates.name);
    }
    if (updates.role !== undefined) {
      sheet.getRange(rowIndex, TEAM_COLS.ROLE).setValue(updates.role);
    }
    if (updates.email !== undefined) {
      sheet.getRange(rowIndex, TEAM_COLS.EMAIL).setValue(updates.email);
    }
    if (updates.phone !== undefined) {
      sheet.getRange(rowIndex, TEAM_COLS.PHONE).setValue(updates.phone);
    }
    if (updates.stages !== undefined) {
      const stages = Array.isArray(updates.stages)
        ? updates.stages.join(', ')
        : updates.stages;
      sheet.getRange(rowIndex, TEAM_COLS.STAGES).setValue(stages);
    }
    if (updates.status !== undefined) {
      sheet.getRange(rowIndex, TEAM_COLS.STATUS).setValue(updates.status);
    }
    if (updates.notes !== undefined) {
      sheet.getRange(rowIndex, TEAM_COLS.NOTES).setValue(updates.notes);
    }

    return true;
  } catch (error) {
    console.error('Error updating team member:', error);
    return false;
  }
}

/**
 * إنشاء كود عضو فريق جديد
 * @returns {string} كود العضو
 */
function generateTeamCode() {
  const sheet = getSheet(SHEETS.TEAM);
  if (!sheet) return 'T001';

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'T001';

  const codes = sheet.getRange(2, TEAM_COLS.CODE, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(code => code && code.startsWith('T'));

  if (codes.length === 0) return 'T001';

  const numbers = codes.map(code => parseInt(code.replace('T', ''), 10));
  const maxNum = Math.max(...numbers);

  return 'T' + (maxNum + 1).toString().padStart(3, '0');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال إضافة وتعديل المصورين
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * إضافة مصور جديد
 * @param {Object} photographerData بيانات المصور
 * @returns {boolean} نجاح العملية
 */
function addPhotographer(photographerData) {
  try {
    const sheet = getSheet(SHEETS.PHOTOGRAPHERS);
    if (!sheet) {
      showError('شيت المصورين غير موجود');
      return false;
    }

    // التحقق من البيانات المطلوبة
    if (!photographerData.name) {
      showError('اسم المصور مطلوب');
      return false;
    }

    // إنشاء كود المصور تلقائياً
    const code = photographerData.code || generatePhotographerCode();

    // تجهيز صف البيانات
    const rowData = [
      code,
      photographerData.name,
      photographerData.specialization || PHOTOGRAPHER_SPECS[0],
      photographerData.email || '',
      photographerData.phone || '',
      photographerData.equipment || '',
      photographerData.status || PHOTOGRAPHER_STATUS.AVAILABLE,
      photographerData.rate || '',
      photographerData.notes || ''
    ];

    sheet.appendRow(rowData);

    return true;
  } catch (error) {
    console.error('Error adding photographer:', error);
    showError('حدث خطأ أثناء إضافة المصور');
    return false;
  }
}

/**
 * تحديث مصور
 * @param {string} code كود المصور
 * @param {Object} updates التحديثات
 * @returns {boolean} نجاح العملية
 */
function updatePhotographer(code, updates) {
  try {
    const sheet = getSheet(SHEETS.PHOTOGRAPHERS);
    if (!sheet) return false;

    const rowIndex = findRowByValue(SHEETS.PHOTOGRAPHERS, PHOTOGRAPHER_COLS.CODE, code);
    if (rowIndex === -1) {
      showError('المصور غير موجود');
      return false;
    }

    // تحديث الحقول
    if (updates.name !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.NAME).setValue(updates.name);
    }
    if (updates.specialization !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.SPECIALIZATION).setValue(updates.specialization);
    }
    if (updates.email !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.EMAIL).setValue(updates.email);
    }
    if (updates.phone !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.PHONE).setValue(updates.phone);
    }
    if (updates.equipment !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.EQUIPMENT).setValue(updates.equipment);
    }
    if (updates.status !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.STATUS).setValue(updates.status);
    }
    if (updates.rate !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.RATE).setValue(updates.rate);
    }
    if (updates.notes !== undefined) {
      sheet.getRange(rowIndex, PHOTOGRAPHER_COLS.NOTES).setValue(updates.notes);
    }

    return true;
  } catch (error) {
    console.error('Error updating photographer:', error);
    return false;
  }
}

/**
 * إنشاء كود مصور جديد
 * @returns {string} كود المصور
 */
function generatePhotographerCode() {
  const sheet = getSheet(SHEETS.PHOTOGRAPHERS);
  if (!sheet) return 'PH001';

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'PH001';

  const codes = sheet.getRange(2, PHOTOGRAPHER_COLS.CODE, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(code => code && code.startsWith('PH'));

  if (codes.length === 0) return 'PH001';

  const numbers = codes.map(code => parseInt(code.replace('PH', ''), 10));
  const maxNum = Math.max(...numbers);

  return 'PH' + (maxNum + 1).toString().padStart(3, '0');
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال الإحصائيات
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * الحصول على إحصائيات الفريق
 * @returns {Object} كائن الإحصائيات
 */
function getTeamStats() {
  const allMembers = getAllTeamMembers();
  const allPhotographers = getAllPhotographers();

  const stats = {
    team: {
      total: allMembers.length,
      active: allMembers.filter(m => m.status === MEMBER_STATUS.ACTIVE).length,
      inactive: allMembers.filter(m => m.status === MEMBER_STATUS.INACTIVE).length,
      byRole: {}
    },
    photographers: {
      total: allPhotographers.length,
      available: allPhotographers.filter(p => p.status === PHOTOGRAPHER_STATUS.AVAILABLE).length,
      busy: allPhotographers.filter(p => p.status === PHOTOGRAPHER_STATUS.BUSY).length,
      unavailable: allPhotographers.filter(p => p.status === PHOTOGRAPHER_STATUS.UNAVAILABLE).length,
      bySpecialization: {}
    }
  };

  // إحصاء الفريق حسب الدور
  allMembers.forEach(member => {
    if (member.role) {
      stats.team.byRole[member.role] = (stats.team.byRole[member.role] || 0) + 1;
    }
  });

  // إحصاء المصورين حسب التخصص
  allPhotographers.forEach(photographer => {
    if (photographer.specialization) {
      stats.photographers.bySpecialization[photographer.specialization] =
        (stats.photographers.bySpecialization[photographer.specialization] || 0) + 1;
    }
  });

  return stats;
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال واجهة المستخدم
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * فتح نموذج إضافة عضو فريق
 */
function showAddTeamMemberDialog() {
  const stageOptions = Object.values(STAGES)
    .map(s => `<option value="${s.name}">${s.icon} ${s.name}</option>`)
    .join('');

  const roleOptions = TEAM_ROLES
    .map(r => `<option value="${r}">${r}</option>`)
    .join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; direction: rtl; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
      select[multiple] { height: 150px; }
      .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px; }
      .btn-primary { background: #1565c0; color: white; }
      .btn-secondary { background: #757575; color: white; }
    </style>

    <h3>إضافة عضو فريق جديد</h3>

    <div class="form-group">
      <label>الاسم *</label>
      <input type="text" id="memberName" required>
    </div>

    <div class="form-group">
      <label>الدور</label>
      <select id="memberRole">${roleOptions}</select>
    </div>

    <div class="form-group">
      <label>البريد الإلكتروني</label>
      <input type="email" id="memberEmail">
    </div>

    <div class="form-group">
      <label>رقم الهاتف</label>
      <input type="tel" id="memberPhone">
    </div>

    <div class="form-group">
      <label>المراحل المسؤول عنها (اختر عدة مراحل بالضغط على Ctrl)</label>
      <select id="memberStages" multiple>${stageOptions}</select>
    </div>

    <div class="form-group">
      <label>ملاحظات</label>
      <textarea id="memberNotes" rows="3"></textarea>
    </div>

    <button class="btn btn-primary" onclick="submitForm()">حفظ</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>

    <script>
      function submitForm() {
        const stagesSelect = document.getElementById('memberStages');
        const selectedStages = Array.from(stagesSelect.selectedOptions).map(opt => opt.value);

        const data = {
          name: document.getElementById('memberName').value,
          role: document.getElementById('memberRole').value,
          email: document.getElementById('memberEmail').value,
          phone: document.getElementById('memberPhone').value,
          stages: selectedStages,
          notes: document.getElementById('memberNotes').value
        };

        if (!data.name) {
          alert('الاسم مطلوب');
          return;
        }

        google.script.run
          .withSuccessHandler(() => {
            alert('تم إضافة العضو بنجاح');
            google.script.host.close();
          })
          .withFailureHandler((err) => {
            alert('حدث خطأ: ' + err.message);
          })
          .addTeamMember(data);
      }
    </script>
  `).setWidth(450).setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة عضو فريق');
}

/**
 * عرض إحصائيات الفريق
 */
function showTeamStats() {
  const stats = getTeamStats();

  const message = `
إحصائيات الفريق والمصورين
━━━━━━━━━━━━━━━━━━━━━━━━━

👥 الفريق:
• إجمالي الأعضاء: ${stats.team.total}
• نشط: ${stats.team.active}
• غير نشط: ${stats.team.inactive}

حسب الدور:
${Object.entries(stats.team.byRole).map(([role, count]) => `  • ${role}: ${count}`).join('\n')}

━━━━━━━━━━━━━━━━━━━━━━━━━

📸 المصورين:
• إجمالي المصورين: ${stats.photographers.total}
• متاح: ${stats.photographers.available}
• مشغول: ${stats.photographers.busy}
• غير متاح: ${stats.photographers.unavailable}

حسب التخصص:
${Object.entries(stats.photographers.bySpecialization).map(([spec, count]) => `  • ${spec}: ${count}`).join('\n')}
  `.trim();

  showInfo(message, 'إحصائيات الفريق');
}
