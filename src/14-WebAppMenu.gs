/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * دوال القائمة الخاصة بنظام المستخدمين و Web App
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * فتح شيت المستخدمين
 */
function openUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USER_SHEET_NAME);
  
  if (!sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'شيت المستخدمين غير موجود',
      'هل تريد إنشاء شيت المستخدمين الآن؟',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      sheet = setupUsersSheet();
      SpreadsheetApp.setActiveSheet(sheet);
      showSuccess('تم إنشاء شيت المستخدمين بنجاح!');
    }
    return;
  }
  
  SpreadsheetApp.setActiveSheet(sheet);
}

/**
 * نموذج إضافة مستخدم جديد
 */
function showAddUserForm() {
  // التحقق من أن المستخدم الحالي هو مدير أو هو صاحب الشيت
  const userEmail = Session.getActiveUser().getEmail();
  const ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner() ? SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() : '';
  
  if (!isCurrentUserAdmin() && userEmail !== ownerEmail) {
    showError('⚠️ فقط المدراء يمكنهم إضافة مستخدمين جدد');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(SHEETS.PROJECTS);
  const projectsData = projectsSheet ? projectsSheet.getDataRange().getValues() : [];
  
  // قائمة المشاريع
  let projectOptions = '';
  for (let i = 1; i < projectsData.length; i++) {
    const code = projectsData[i][PROJECT_COLS.CODE - 1];
    const name = projectsData[i][PROJECT_COLS.NAME - 1];
    if (code && name) {
      projectOptions += `<option value="${code}">${code} - ${name}</option>`;
    }
  }
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { 
        font-family: Arial, sans-serif; 
        padding: 20px; 
        direction: rtl; 
        background: #f5f7fa;
      }
      .form-container {
        background: white;
        padding: 25px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      }
      .form-group { 
        margin-bottom: 20px; 
      }
      label { 
        display: block; 
        margin-bottom: 8px; 
        font-weight: 600;
        color: #333;
      }
      input, select { 
        width: 100%; 
        padding: 10px 12px; 
        box-sizing: border-box; 
        border: 2px solid #e0e0e0;
        border-radius: 6px;
        font-size: 14px;
        transition: border-color 0.3s;
      }
      input:focus, select:focus {
        outline: none;
        border-color: #667eea;
      }
      .hint {
        font-size: 12px;
        color: #666;
        margin-top: 5px;
      }
      .button-group {
        display: flex;
        gap: 10px;
        margin-top: 25px;
      }
      button { 
        flex: 1;
        background: #667eea; 
        color: white; 
        padding: 12px 20px; 
        border: none; 
        cursor: pointer;
        border-radius: 6px;
        font-weight: 600;
        transition: background 0.3s;
      }
      button:hover { 
        background: #5568d3; 
      }
      .cancel { 
        background: #757575; 
      }
      .cancel:hover {
        background: #616161;
      }
      #projectsContainer {
        display: none;
        margin-top: 10px;
      }
      .projects-select {
        min-height: 150px;
      }
    </style>
    <div class="form-container">
      <h2 style="margin-top: 0; color: #333;">➕ إضافة مستخدم جديد</h2>
      <form id="userForm">
        <div class="form-group">
          <label>البريد الإلكتروني *</label>
          <input type="email" id="email" required placeholder="example@seenfilm.com">
        </div>
        
        <div class="form-group">
          <label>الاسم الكامل *</label>
          <input type="text" id="name" required placeholder="أدخل اسم المستخدم">
        </div>
        
        <div class="form-group">
          <label>كلمة المرور *</label>
          <input type="password" id="password" required placeholder="••••••••">
          <div class="hint">الحد الأدنى 6 أحرف</div>
        </div>
        
        <div class="form-group">
          <label>تأكيد كلمة المرور *</label>
          <input type="password" id="confirmPassword" required placeholder="••••••••">
        </div>
        
        <div class="form-group">
          <label>الدور *</label>
          <select id="role" required onchange="toggleProjects()">
            <option value="">اختر الدور</option>
            <option value="مدير عام">مدير عام</option>
            <option value="مدير مشروعات">مدير مشروعات</option>
          </select>
        </div>
        
        <div id="projectsContainer" class="form-group">
          <label>المشاريع المرتبطة</label>
          <select id="projects" multiple class="projects-select">
            ${projectOptions}
          </select>
          <div class="hint">اضغط Ctrl/Cmd لاختيار عدة مشاريع</div>
        </div>
        
        <div class="button-group">
          <button type="submit">إضافة المستخدم</button>
          <button type="button" class="cancel" onclick="google.script.host.close()">إلغاء</button>
        </div>
      </form>
    </div>
    
    <script>
      function toggleProjects() {
        const role = document.getElementById('role').value;
        const container = document.getElementById('projectsContainer');
        
        if (role === 'مدير مشروعات') {
          container.style.display = 'block';
        } else {
          container.style.display = 'none';
        }
      }
      
      document.getElementById('userForm').onsubmit = function(e) {
        e.preventDefault();
        
        const password = document.getElementById('password').value;
        const confirmPassword = document.getElementById('confirmPassword').value;
        
        if (password !== confirmPassword) {
          alert('⚠️ كلمة المرور غير متطابقة');
          return;
        }
        
        if (password.length < 6) {
          alert('⚠️ يجب أن تكون كلمة المرور 6 أحرف على الأقل');
          return;
        }
        
        const role = document.getElementById('role').value;
        let projects = 'ALL';
        
        if (role === 'مدير مشروعات') {
          const selected = Array.from(document.getElementById('projects').selectedOptions)
            .map(opt => opt.value);
          
          if (selected.length === 0) {
            alert('⚠️ يجب اختيار مشروع واحد على الأقل لمدير المشروعات');
            return;
          }
          
          projects = selected.join(', ');
        }
        
        const data = {
          email: document.getElementById('email').value,
          name: document.getElementById('name').value,
          password: password,
          role: role,
          projects: projects
        };
        
        google.script.run
          .withSuccessHandler(function(result) {
            alert(result.message);
            if (result.success) {
              google.script.host.close();
            }
          })
          .withFailureHandler(function(error) {
            alert('❌ خطأ: ' + error.message);
          })
          .addUser(data);
      };
    </script>
  `)
    .setWidth(550)
    .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'إضافة مستخدم جديد');
}

/**
 * الحصول على إحصائيات لوحة التحكم للويب أب
 */
function webAppGetDashboardStats() {
  const stats = getMovementStats();
  const projects = getActiveProjects();

  return sanitizeForClient({
    totalProjects: projects.length,
    activeProjects: projects.length, // حالياً نعتبر النشط هو المتاح
    totalTasks: stats.total,
    delayedTasks: stats.delayed,
    teamCount: getActiveTeamMembers().length
  });
}

/**
 * فتح Web App
 */
function openWebApp() {
  const url = getWebAppUrl();
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        padding: 30px;
        direction: rtl;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        text-align: center;
      }
      .container {
        background: white;
        color: #333;
        padding: 30px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
      }
      h2 {
        margin-top: 0;
        color: #667eea;
      }
      .url-box {
        background: #f5f7fa;
        padding: 15px;
        border-radius: 8px;
        margin: 20px 0;
        word-break: break-all;
        font-family: monospace;
        font-size: 13px;
      }
      button {
        background: #667eea;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 8px;
        cursor: pointer;
        font-size: 15px;
        font-weight: 600;
        margin: 5px;
        transition: background 0.3s;
      }
      button:hover {
        background: #5568d3;
      }
      .info {
        background: #e3f2fd;
        padding: 15px;
        border-radius: 8px;
        border-right: 4px solid #2196f3;
        margin-top: 20px;
        text-align: right;
      }
    </style>
    <div class="container">
      <h2>🌐 رابط نظام Web App</h2>
      <div class="url-box">${url}</div>
      <button onclick="window.open('${url}', '_blank')">فتح Web App</button>
      <button onclick="copyToClipboard()">نسخ الرابط</button>
      <div class="info">
        <strong>📌 ملاحظة:</strong><br>
        شارك هذا الرابط مع فريقك للوصول إلى النظام.<br>
        كل شخص سيسجل دخول بالبريد وكلمة المرور الخاصة به.
      </div>
    </div>
    <script>
      function copyToClipboard() {
        const url = '${url}';
        navigator.clipboard.writeText(url).then(function() {
          alert('✅ تم نسخ الرابط بنجاح!');
        }, function() {
          alert('⚠️ فشل نسخ الرابط. استخدم Ctrl+C');
        });
      }
    </script>
  `)
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'رابط Web App');
}

/**
 * نسخ رابط Web App
 */
function copyWebAppUrl() {
  const url = getWebAppUrl();
  
  SpreadsheetApp.getUi().alert(
    '📋 رابط Web App',
    'الرابط:\\n\\n' + url + '\\n\\n' +
    'تم نسخه إلى الحافظة (إن أمكن)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * الحصول على رابط Web App
 */
function getWebAppUrl() {
  // استخدام معرف النسخة المستقرة الموثوقة (Deployment ID)
  // هذا الرابط مفتوح للجميع (ANYONE) ويعمل بصلاحيات المالك (USER_DEPLOYING)
  const stableId = 'AKfycbxE_2BD4ljd1ci2rmaXlrpZF4M5mK7MD-5riWgy44PQR3lnFdjdyKOiOrwHRj_2gehn';
  return `https://script.google.com/macros/s/${stableId}/exec`;
}

/**
 * إعادة تعيين كلمة مرور المستخدم الحالي (إصلاح سريع للوصول)
 */
function resetMyPassword() {
  const email = Session.getActiveUser().getEmail();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('🔑 إعادة تعيين كلمة المرور', 
    `سيتم تعيين كلمة مرور جديدة للحساب: ${email}\n\nالرجاء إدخال كلمة المرور (6 أحرف على الأقل):`, 
    ui.ButtonSet.OK_CANCEL);
    
  if (response.getSelectedButton() === ui.Button.OK) {
    const newPassword = response.getResponseText().trim();
    if (newPassword.length < 5) {
      ui.alert('⚠️ كلمة المرور قصيرة جداً');
      return;
    }
    
    // تشفير
    const hashedPassword = hashPassword(newPassword);
    const user = getUserByEmail(email);
    
    if (!user) {
      ui.alert('⚠️ لم يتم العثور على بريدك في شيت "المستخدمين".\nيرجى التأكد من إضافة بريدك الإلكتروني في الشيت أولاً.');
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET_NAME);
    sheet.getRange(user.row, USER_COLS.PASSWORD).setValue(hashedPassword);
    
    ui.alert('✅ تم تحديث كلمة المرور بنجاح!\nيمكنك الآن تسجيل الدخول في الويب أب بكلمة السر الجديدة.');
  }
}
