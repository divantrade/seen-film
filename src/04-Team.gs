/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * إدارة الفريق
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * إضافة عضو جديد للفريق عبر نموذج
 */
function showAddTeamMemberForm() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; direction: rtl; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; box-sizing: border-box; }
      button { background: #1565C0; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-left: 10px; }
      button:hover { background: #0D47A1; }
      .cancel { background: #757575; }
    </style>
    <form id="teamForm">
      <div class="form-group">
        <label>الاسم *</label>
        <input type="text" id="name" required>
      </div>
      <div class="form-group">
        <label>الدور *</label>
        <select id="role" required>
          <option value="">اختر الدور</option>
          ${TEAM_ROLES.map(r => `<option value="${r}">${r}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>البريد الإلكتروني</label>
        <input type="email" id="email">
      </div>
      <div class="form-group">
        <label>رقم الهاتف</label>
        <input type="text" id="phone">
      </div>
      <div class="form-group">
        <label>ملاحظات</label>
        <input type="text" id="notes">
      </div>
      <button type="submit">إضافة</button>
      <button type="button" class="cancel" onclick="google.script.host.close()">إلغاء</button>
    </form>
    <script>
      document.getElementById('teamForm').onsubmit = function(e) {
        e.preventDefault();
        const data = {
          name: document.getElementById('name').value,
          role: document.getElementById('role').value,
          email: document.getElementById('email').value,
          phone: document.getElementById('phone').value,
          notes: document.getElementById('notes').value
        };
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .addTeamMember(data);
      };
    </script>
  `)
    .setWidth(400)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة عضو جديد');
}

/**
 * إضافة عضو جديد للفريق
 */
function addTeamMember(data) {
  const sheet = getSheet(SHEETS.TEAM);
  const lastRow = getLastRowInColumn(sheet, TEAM_COLS.NAME);
  const newRow = lastRow + 1;

  // توليد الكود
  const code = generateTeamCode(data.role);

  // إضافة البيانات
  sheet.getRange(newRow, TEAM_COLS.CODE).setValue(code);
  sheet.getRange(newRow, TEAM_COLS.NAME).setValue(cleanText(data.name));
  sheet.getRange(newRow, TEAM_COLS.ROLE).setValue(data.role);
  sheet.getRange(newRow, TEAM_COLS.EMAIL).setValue(data.email || '');
  sheet.getRange(newRow, TEAM_COLS.PHONE).setValue(data.phone || '');
  sheet.getRange(newRow, TEAM_COLS.STATUS).setValue('نشط');
  sheet.getRange(newRow, TEAM_COLS.JOIN_DATE).setValue(getCurrentDate());
  sheet.getRange(newRow, TEAM_COLS.NOTES).setValue(data.notes || '');

  // تحديث القوائم المنسدلة في شيت الحركة
  updateMovementDropdowns();

  showSuccess('تم إضافة ' + data.name + ' بنجاح');
}

/**
 * تحديث حالة عضو الفريق
 */
function toggleTeamMemberStatus() {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== SHEETS.TEAM) {
    showError('يجب أن تكون في شيت الفريق');
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    showError('اختر عضو من القائمة');
    return;
  }

  const currentStatus = sheet.getRange(row, TEAM_COLS.STATUS).getValue();
  const newStatus = currentStatus === 'نشط' ? 'غير نشط' : 'نشط';

  sheet.getRange(row, TEAM_COLS.STATUS).setValue(newStatus);

  // تحديث القوائم المنسدلة
  updateMovementDropdowns();

  showSuccess('تم تغيير الحالة إلى: ' + newStatus);
}

/**
 * تحديث القوائم المنسدلة في شيت الحركة بأسماء الفريق
 */
function updateMovementDropdowns() {
  const movementSheet = getSheet(SHEETS.MOVEMENT);
  const projectsSheet = getSheet(SHEETS.PROJECTS);

  if (!movementSheet) return;

  // تحديث قائمة المسؤولين
  const teamNames = getActiveTeamNames();
  if (teamNames.length > 0) {
    setDropdown(movementSheet, 2, MOVEMENT_COLS.ASSIGNED_TO, 500, teamNames);
  }

  // تحديث قائمة المشاريع
  const projectNames = getActiveProjectNames();
  if (projectNames.length > 0) {
    setDropdown(movementSheet, 2, MOVEMENT_COLS.PROJECT, 500, projectNames);
  }

  // تحديث قائمة المنتجين والمونتيرين في شيت المشاريع
  if (projectsSheet && teamNames.length > 0) {
    setDropdown(projectsSheet, 2, PROJECT_COLS.PRODUCER, 100, teamNames);
    setDropdown(projectsSheet, 2, PROJECT_COLS.EDITOR, 100, teamNames);
  }
}

/**
 * Trigger عند تعديل شيت الفريق
 */
function onTeamEdit(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== SHEETS.TEAM) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // إذا تم تعديل الدور، قم بتحديث الكود
  if (col === TEAM_COLS.ROLE && row > 1) {
    const newRole = e.value;
    const currentCode = sheet.getRange(row, TEAM_COLS.CODE).getValue();

    // فقط إذا كان الكود فارغاً
    if (!currentCode) {
      const newCode = generateTeamCode(newRole);
      sheet.getRange(row, TEAM_COLS.CODE).setValue(newCode);
    }
  }

  // إذا تم تعديل الحالة، قم بتحديث القوائم
  if (col === TEAM_COLS.STATUS) {
    updateMovementDropdowns();
  }
}

/**
 * البحث عن عضو بالاسم
 */
function findTeamMemberByName(name) {
  const sheet = getSheet(SHEETS.TEAM);
  const row = findRowByValue(sheet, TEAM_COLS.NAME, name);

  if (row === -1) return null;

  const data = sheet.getRange(row, 1, 1, TEAM_COLS.NOTES).getValues()[0];

  return {
    code: data[TEAM_COLS.CODE - 1],
    name: data[TEAM_COLS.NAME - 1],
    role: data[TEAM_COLS.ROLE - 1],
    email: data[TEAM_COLS.EMAIL - 1],
    phone: data[TEAM_COLS.PHONE - 1],
    status: data[TEAM_COLS.STATUS - 1]
  };
}

/**
 * الحصول على إحصائيات الفريق
 */
function getTeamStats() {
  const members = getActiveTeamMembers();
  const stats = {};

  for (const member of members) {
    if (!stats[member.role]) {
      stats[member.role] = 0;
    }
    stats[member.role]++;
  }

  return stats;
}
