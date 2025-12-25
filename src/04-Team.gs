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
          ${getTeamRolesFromSettings().map(r => `<option value="${r}">${r}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>المدينة *</label>
        <select id="city" required>
          <option value="">اختر المدينة</option>
          ${getCitiesFromSettings().map(c => `<option value="${c}">${c}</option>`).join('')}
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
          city: document.getElementById('city').value,
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
  sheet.getRange(newRow, TEAM_COLS.CITY).setValue(data.city || '');
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
  const teamSheet = getSheet(SHEETS.TEAM);

  if (!movementSheet) return;

  const teamNames = getActiveTeamNames();
  const cities = getCitiesFromSettings();

  // 1. تحديث قائمة المسؤولين في شيت الحركة
  if (teamNames.length > 0) {
    const assignedToCol = getColumnByHeader(movementSheet, 'المسؤول');
    if (assignedToCol !== -1) {
      setDropdown(movementSheet, 2, assignedToCol, 500, teamNames);
    }
  }

  // 2. تحديث قائمة المدن في شيت الفريق
  if (teamSheet && cities.length > 0) {
    const cityCol = getColumnByHeader(teamSheet, 'المدينة');
    if (cityCol !== -1) {
      setDropdown(teamSheet, 2, cityCol, 500, cities);
    }
  }

  // 3. تحديث قائمة المشاريع في شيت الحركة
  const projectNames = getActiveProjectNames();
  if (projectNames.length > 0) {
    const projectCol = getColumnByHeader(movementSheet, 'الفيلم');
    if (projectCol !== -1) {
      setDropdown(movementSheet, 2, projectCol, 500, projectNames);
    }
  }

  // 4. تحديث قائمة المنتجين والمونتيرين في شيت المشاريع
  if (projectsSheet && teamNames.length > 0) {
    const producerCol = getColumnByHeader(projectsSheet, 'المنتج المسؤول');
    const editorCol = getColumnByHeader(projectsSheet, 'المونتير');
    
    if (producerCol !== -1) setDropdown(projectsSheet, 2, producerCol, 500, teamNames);
    if (editorCol !== -1) setDropdown(projectsSheet, 2, editorCol, 500, teamNames);
  }
}

/**
 * Trigger عند تعديل شيت الفريق
 */
function onTeamEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEETS.TEAM) return;

  const range = e.range;
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const col = range.getColumn();

  const nameCol = getColumnByHeader(sheet, 'الاسم');
  const roleCol = getColumnByHeader(sheet, 'الدور');
  const codeCol = getColumnByHeader(sheet, 'الكود');
  const statusCol = getColumnByHeader(sheet, 'الحالة');
  const joinDateCol = getColumnByHeader(sheet, 'تاريخ الانضمام');

  // معالجة كل صف في النطاق المعدل (لدعم النسخ واللصق)
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    if (currentRow <= 1) continue;

    // 0. تطبيع التواريخ تلقائياً إلى dd/mm/yyyy
    if (col === joinDateCol && joinDateCol !== -1) {
      const cell = sheet.getRange(currentRow, col);
      const value = cell.getValue();
      if (value) {
        normalizeDateCell_(cell, value);
      }
    }

    // إذا تم تعديل الدور أو الاسم، تأكد من وجود كود
    if (col === roleCol || col === nameCol) {
      const role = (roleCol !== -1) ? sheet.getRange(currentRow, roleCol).getValue() : '';
      const currentCode = (codeCol !== -1) ? sheet.getRange(currentRow, codeCol).getValue() : '';

      if (!currentCode && role) {
        const newCode = generateTeamCode(role);
        if (codeCol !== -1) {
          sheet.getRange(currentRow, codeCol).setValue(newCode);
        }
      }
    }
  }

  // إذا تم تعديل الحالة أو الدور أو الاسم، قم بتحديث القوائم المنسدلة
  if (col === statusCol || col === roleCol || col === nameCol) {
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
    city: data[TEAM_COLS.CITY - 1],
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
