/**
 * ===================================================
 * 15-Notifications.gs - نظام الإشعارات بالإيميل
 * ===================================================
 *
 * هذا الملف يحتوي على:
 * - إرسال تقارير يومية وأسبوعية
 * - تذكيرات بالمواعيد النهائية
 * - إشعارات المهام المتأخرة
 * - إدارة الـ Triggers
 * - إعدادات الإشعارات
 */

// ====================================================
// ثوابت الإشعارات
// ====================================================

/**
 * مفاتيح خصائص الإشعارات
 */
const NOTIFICATION_PROPS = {
  ENABLED_DAILY: 'notification_daily_enabled',
  ENABLED_WEEKLY: 'notification_weekly_enabled',
  RECIPIENTS: 'notification_recipients',
  DAILY_TIME: 'notification_daily_time',
  WEEKLY_DAY: 'notification_weekly_day'
};

/**
 * الإعدادات الافتراضية
 */
const DEFAULT_NOTIFICATION_SETTINGS = {
  dailyEnabled: false,
  weeklyEnabled: false,
  recipients: '',
  dailyTime: '09:00',
  weeklyDay: 'SUNDAY'
};

// ====================================================
// دوال إرسال التقارير
// ====================================================

/**
 * إرسال التقرير اليومي
 * يحتوي على: المهام المتأخرة، المواعيد خلال 3 أيام، ملخص المشاريع
 */
function sendDailyReport() {
  try {
    const settings = getNotificationSettings();
    if (!settings.recipients) {
      console.log('لا توجد عناوين بريد للإرسال');
      return;
    }

    const recipients = settings.recipients.split(',').map(e => e.trim());
    const subject = `📋 تقرير يومي - نظام الإنتاج - ${formatDate(new Date())}`;
    const htmlBody = generateDailyReportHtml();

    recipients.forEach(email => {
      if (email && isValidEmail(email)) {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody
        });
      }
    });

    console.log('تم إرسال التقرير اليومي بنجاح');

  } catch (error) {
    console.error('خطأ في إرسال التقرير اليومي:', error);
  }
}

/**
 * إرسال التقرير الأسبوعي
 * يحتوي على: نسب إنجاز المشاريع، الإنجازات، الخطط، التنبيهات
 */
function sendWeeklyReport() {
  try {
    const settings = getNotificationSettings();
    if (!settings.recipients) {
      console.log('لا توجد عناوين بريد للإرسال');
      return;
    }

    const recipients = settings.recipients.split(',').map(e => e.trim());
    const weekStart = new Date();
    weekStart.setDate(weekStart.getDate() - weekStart.getDay());
    const subject = `📊 تقرير أسبوعي - نظام الإنتاج - أسبوع ${formatDate(weekStart)}`;
    const htmlBody = generateWeeklyReportHtml();

    recipients.forEach(email => {
      if (email && isValidEmail(email)) {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody
        });
      }
    });

    console.log('تم إرسال التقرير الأسبوعي بنجاح');

  } catch (error) {
    console.error('خطأ في إرسال التقرير الأسبوعي:', error);
  }
}

/**
 * إرسال تذكير بالمواعيد النهائية
 * @param {number} days - عدد الأيام القادمة (افتراضي 3)
 */
function sendDeadlineReminder(days = 3) {
  try {
    const settings = getNotificationSettings();
    if (!settings.recipients) return;

    const upcoming = getUpcomingDeadlines(days);
    if (upcoming.length === 0) return;

    const recipients = settings.recipients.split(',').map(e => e.trim());
    const subject = `⏰ تذكير: ${upcoming.length} مهمة تستحق خلال ${days} أيام`;
    const htmlBody = generateDeadlineReminderHtml(upcoming, days);

    recipients.forEach(email => {
      if (email && isValidEmail(email)) {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody
        });
      }
    });

  } catch (error) {
    console.error('خطأ في إرسال تذكير المواعيد:', error);
  }
}

/**
 * إرسال مهام شخص معين إليه
 * @param {string} personEmail - البريد الإلكتروني للشخص
 * @param {string} personName - اسم الشخص
 */
function sendPersonTasks(personEmail, personName) {
  try {
    if (!personEmail || !isValidEmail(personEmail)) {
      console.log('بريد إلكتروني غير صالح');
      return { success: false, error: 'بريد إلكتروني غير صالح' };
    }

    const report = generatePersonReport(personName);
    const subject = `📋 مهامك - ${formatDate(new Date())}`;
    const htmlBody = generatePersonTasksHtml(report);

    MailApp.sendEmail({
      to: personEmail,
      subject: subject,
      htmlBody: htmlBody
    });

    return { success: true };

  } catch (error) {
    console.error('خطأ في إرسال مهام الشخص:', error);
    return { success: false, error: error.message };
  }
}

/**
 * إشعار بالمهام المتأخرة
 */
function notifyOverdueTasks() {
  try {
    const settings = getNotificationSettings();
    if (!settings.recipients) return;

    const overdue = getOverdueTasks();
    if (overdue.length === 0) return;

    const recipients = settings.recipients.split(',').map(e => e.trim());
    const subject = `🔴 تنبيه: ${overdue.length} مهمة متأخرة`;
    const htmlBody = generateOverdueTasksHtml(overdue);

    recipients.forEach(email => {
      if (email && isValidEmail(email)) {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody
        });
      }
    });

  } catch (error) {
    console.error('خطأ في إرسال إشعار المهام المتأخرة:', error);
  }
}

// ====================================================
// دوال توليد HTML للإيميلات
// ====================================================

/**
 * الحصول على نمط CSS للإيميلات
 * @returns {string} كود CSS
 */
function getEmailStyles() {
  return `
    <style>
      body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; direction: rtl; background: #f5f5f5; margin: 0; padding: 20px; }
      .container { max-width: 700px; margin: 0 auto; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
      .header { background: linear-gradient(135deg, #1a73e8, #0d47a1); color: white; padding: 30px; text-align: center; }
      .header h1 { margin: 0; font-size: 24px; }
      .header p { margin: 10px 0 0; opacity: 0.9; }
      .content { padding: 30px; }
      .section { margin-bottom: 30px; }
      .section-title { font-size: 18px; font-weight: bold; color: #202124; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #e0e0e0; }
      .section-title span { margin-left: 8px; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; }
      th { background: #f8f9fa; padding: 12px; text-align: right; font-weight: 600; color: #5f6368; border-bottom: 2px solid #e0e0e0; }
      td { padding: 12px; border-bottom: 1px solid #f0f0f0; color: #3c4043; }
      tr:hover { background: #f8f9fa; }
      .badge { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 500; }
      .badge-danger { background: #fce8e6; color: #c5221f; }
      .badge-warning { background: #fef7e0; color: #b06000; }
      .badge-success { background: #e6f4ea; color: #137333; }
      .badge-info { background: #e8f0fe; color: #1967d2; }
      .stat-box { display: inline-block; background: #f8f9fa; border-radius: 8px; padding: 15px 25px; margin: 5px; text-align: center; }
      .stat-number { font-size: 28px; font-weight: bold; color: #1a73e8; }
      .stat-label { font-size: 12px; color: #5f6368; margin-top: 5px; }
      .alert { padding: 15px; border-radius: 8px; margin-bottom: 15px; }
      .alert-danger { background: #fce8e6; border-right: 4px solid #c5221f; }
      .alert-warning { background: #fef7e0; border-right: 4px solid #f9ab00; }
      .alert-info { background: #e8f0fe; border-right: 4px solid #1967d2; }
      .footer { background: #f8f9fa; padding: 20px; text-align: center; color: #5f6368; font-size: 12px; }
      .progress-bar { height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; }
      .progress-fill { height: 100%; border-radius: 4px; }
      .progress-high { background: #34a853; }
      .progress-medium { background: #fbbc04; }
      .progress-low { background: #ea4335; }
    </style>
  `;
}

/**
 * توليد HTML للتقرير اليومي
 * @returns {string} كود HTML
 */
function generateDailyReportHtml() {
  const today = new Date();
  const overdue = getOverdueTasks();
  const upcoming = getUpcomingDeadlines(3);
  const projects = getAllProjectsSummary();

  return `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <meta charset="UTF-8">
      ${getEmailStyles()}
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>📋 التقرير اليومي</h1>
          <p>${formatDate(today)}</p>
        </div>

        <div class="content">
          <!-- ملخص سريع -->
          <div class="section">
            <div style="text-align: center;">
              <div class="stat-box">
                <div class="stat-number">${overdue.length}</div>
                <div class="stat-label">مهام متأخرة</div>
              </div>
              <div class="stat-box">
                <div class="stat-number">${upcoming.length}</div>
                <div class="stat-label">تستحق خلال 3 أيام</div>
              </div>
              <div class="stat-box">
                <div class="stat-number">${projects.length}</div>
                <div class="stat-label">مشاريع نشطة</div>
              </div>
            </div>
          </div>

          <!-- المهام المتأخرة -->
          ${overdue.length > 0 ? `
          <div class="section">
            <div class="section-title"><span>🔴</span> المهام المتأخرة</div>
            <table>
              <tr>
                <th>المشروع</th>
                <th>المهمة</th>
                <th>المسؤول</th>
                <th>الموعد</th>
                <th>التأخير</th>
              </tr>
              ${overdue.slice(0, 10).map(task => {
                const daysOverdue = Math.ceil((today - new Date(task.dueDate)) / (1000 * 60 * 60 * 24));
                return `
                  <tr>
                    <td>${task.projectCode || task.projectName || '-'}</td>
                    <td>${task.element || task.action || '-'}</td>
                    <td>${task.assignedTo || '-'}</td>
                    <td>${formatDate(task.dueDate)}</td>
                    <td><span class="badge badge-danger">${daysOverdue} يوم</span></td>
                  </tr>
                `;
              }).join('')}
            </table>
            ${overdue.length > 10 ? `<p style="color: #5f6368; margin-top: 10px;">... و ${overdue.length - 10} مهمة أخرى</p>` : ''}
          </div>
          ` : `
          <div class="section">
            <div class="alert alert-info">✅ لا توجد مهام متأخرة - ممتاز!</div>
          </div>
          `}

          <!-- المواعيد القادمة -->
          ${upcoming.length > 0 ? `
          <div class="section">
            <div class="section-title"><span>📅</span> تستحق خلال 3 أيام</div>
            <table>
              <tr>
                <th>المشروع</th>
                <th>المهمة</th>
                <th>المسؤول</th>
                <th>الموعد</th>
                <th>متبقي</th>
              </tr>
              ${upcoming.slice(0, 10).map(task => {
                const daysLeft = Math.ceil((new Date(task.dueDate) - today) / (1000 * 60 * 60 * 24));
                const badgeClass = daysLeft === 0 ? 'badge-danger' : daysLeft <= 1 ? 'badge-warning' : 'badge-info';
                const daysText = daysLeft === 0 ? 'اليوم!' : `${daysLeft} يوم`;
                return `
                  <tr>
                    <td>${task.projectCode || task.projectName || '-'}</td>
                    <td>${task.element || task.action || '-'}</td>
                    <td>${task.assignedTo || '-'}</td>
                    <td>${formatDate(task.dueDate)}</td>
                    <td><span class="badge ${badgeClass}">${daysText}</span></td>
                  </tr>
                `;
              }).join('')}
            </table>
          </div>
          ` : ''}

          <!-- ملخص المشاريع -->
          <div class="section">
            <div class="section-title"><span>📁</span> المشاريع النشطة</div>
            <table>
              <tr>
                <th>المشروع</th>
                <th>العميل</th>
                <th>نسبة الإنجاز</th>
                <th>التنبيهات</th>
              </tr>
              ${projects.map(p => {
                const progressClass = p.progress >= 75 ? 'progress-high' : p.progress >= 50 ? 'progress-medium' : 'progress-low';
                return `
                  <tr>
                    <td><strong>${p.code}</strong><br><small>${p.name}</small></td>
                    <td>${p.client || '-'}</td>
                    <td>
                      <div style="display: flex; align-items: center;">
                        <div class="progress-bar" style="width: 80px; margin-left: 10px;">
                          <div class="progress-fill ${progressClass}" style="width: ${p.progress}%;"></div>
                        </div>
                        <span>${p.progress}%</span>
                      </div>
                    </td>
                    <td>${p.alertsCount > 0 ? `<span class="badge badge-warning">⚠️ ${p.alertsCount}</span>` : '<span class="badge badge-success">✓</span>'}</td>
                  </tr>
                `;
              }).join('')}
            </table>
          </div>
        </div>

        <div class="footer">
          <p>تم إنشاء هذا التقرير تلقائياً من نظام متابعة الإنتاج الفني</p>
          <p>${formatDate(new Date())} - ${new Date().toLocaleTimeString('ar-SA')}</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * توليد HTML للتقرير الأسبوعي
 * @returns {string} كود HTML
 */
function generateWeeklyReportHtml() {
  const today = new Date();
  const weekStart = new Date(today);
  weekStart.setDate(today.getDate() - today.getDay());

  const projects = getAllProjectsSummary();
  const weeklyReport = generateWeeklyReport();

  // حساب الإحصائيات الكلية
  const totalTasks = projects.reduce((sum, p) => sum + p.totalTasks, 0);
  const completedTasks = projects.reduce((sum, p) => sum + p.completedTasks, 0);
  const overallProgress = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;

  return `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <meta charset="UTF-8">
      ${getEmailStyles()}
    </head>
    <body>
      <div class="container">
        <div class="header" style="background: linear-gradient(135deg, #0d47a1, #1565c0);">
          <h1>📊 التقرير الأسبوعي</h1>
          <p>أسبوع ${formatDate(weekStart)} - ${formatDate(today)}</p>
        </div>

        <div class="content">
          <!-- ملخص الأسبوع -->
          <div class="section">
            <div class="section-title"><span>📈</span> ملخص الأسبوع</div>
            <div style="text-align: center;">
              <div class="stat-box">
                <div class="stat-number">${weeklyReport.summary.completed}</div>
                <div class="stat-label">مهام أُكملت</div>
              </div>
              <div class="stat-box">
                <div class="stat-number">${weeklyReport.summary.new}</div>
                <div class="stat-label">مهام جديدة</div>
              </div>
              <div class="stat-box">
                <div class="stat-number">${overallProgress}%</div>
                <div class="stat-label">الإنجاز الكلي</div>
              </div>
              <div class="stat-box">
                <div class="stat-number">${weeklyReport.summary.overdue}</div>
                <div class="stat-label">متأخرة</div>
              </div>
            </div>
          </div>

          <!-- حالة المشاريع -->
          <div class="section">
            <div class="section-title"><span>📁</span> حالة المشاريع</div>
            ${projects.map(p => {
              const progressClass = p.progress >= 75 ? 'progress-high' : p.progress >= 50 ? 'progress-medium' : 'progress-low';
              return `
                <div style="margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 8px;">
                  <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                    <strong>${p.code} - ${p.name}</strong>
                    <span class="badge ${p.alertsCount > 0 ? 'badge-warning' : 'badge-success'}">${p.status}</span>
                  </div>
                  <div style="display: flex; justify-content: space-between; align-items: center;">
                    <span style="color: #5f6368;">المهام: ${p.completedTasks}/${p.totalTasks}</span>
                    <div style="display: flex; align-items: center;">
                      <div class="progress-bar" style="width: 150px; margin-left: 10px;">
                        <div class="progress-fill ${progressClass}" style="width: ${p.progress}%;"></div>
                      </div>
                      <strong>${p.progress}%</strong>
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>

          <!-- المهام المخططة للأسبوع القادم -->
          ${weeklyReport.upcomingNextWeek.length > 0 ? `
          <div class="section">
            <div class="section-title"><span>📅</span> الأسبوع القادم</div>
            <table>
              <tr>
                <th>المشروع</th>
                <th>المهمة</th>
                <th>المسؤول</th>
                <th>الموعد</th>
              </tr>
              ${weeklyReport.upcomingNextWeek.slice(0, 10).map(task => `
                <tr>
                  <td>${task.projectCode || task.projectName || '-'}</td>
                  <td>${task.element || task.action || '-'}</td>
                  <td>${task.assignedTo || '-'}</td>
                  <td>${formatDate(task.dueDate)}</td>
                </tr>
              `).join('')}
            </table>
          </div>
          ` : ''}

          <!-- تنبيهات مهمة -->
          ${weeklyReport.summary.overdue > 0 ? `
          <div class="section">
            <div class="alert alert-danger">
              <strong>⚠️ تنبيه:</strong> يوجد ${weeklyReport.summary.overdue} مهمة متأخرة تحتاج إلى متابعة فورية
            </div>
          </div>
          ` : ''}
        </div>

        <div class="footer">
          <p>تم إنشاء هذا التقرير تلقائياً من نظام متابعة الإنتاج الفني</p>
          <p>${formatDate(new Date())} - ${new Date().toLocaleTimeString('ar-SA')}</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * توليد HTML لتذكير المواعيد
 * @param {Array} tasks - المهام
 * @param {number} days - عدد الأيام
 * @returns {string} كود HTML
 */
function generateDeadlineReminderHtml(tasks, days) {
  const today = new Date();

  return `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <meta charset="UTF-8">
      ${getEmailStyles()}
    </head>
    <body>
      <div class="container">
        <div class="header" style="background: linear-gradient(135deg, #f9ab00, #ff8f00);">
          <h1>⏰ تذكير بالمواعيد النهائية</h1>
          <p>${tasks.length} مهمة تستحق خلال ${days} أيام</p>
        </div>

        <div class="content">
          <table>
            <tr>
              <th>المشروع</th>
              <th>المهمة</th>
              <th>المسؤول</th>
              <th>الموعد</th>
              <th>متبقي</th>
            </tr>
            ${tasks.map(task => {
              const daysLeft = Math.ceil((new Date(task.dueDate) - today) / (1000 * 60 * 60 * 24));
              const badgeClass = daysLeft === 0 ? 'badge-danger' : daysLeft <= 1 ? 'badge-warning' : 'badge-info';
              const daysText = daysLeft === 0 ? 'اليوم!' : `${daysLeft} يوم`;
              return `
                <tr>
                  <td>${task.projectCode || task.projectName || '-'}</td>
                  <td>${task.element || task.action || '-'}</td>
                  <td>${task.assignedTo || '-'}</td>
                  <td>${formatDate(task.dueDate)}</td>
                  <td><span class="badge ${badgeClass}">${daysText}</span></td>
                </tr>
              `;
            }).join('')}
          </table>
        </div>

        <div class="footer">
          <p>تم إنشاء هذا التذكير تلقائياً</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * توليد HTML للمهام المتأخرة
 * @param {Array} tasks - المهام المتأخرة
 * @returns {string} كود HTML
 */
function generateOverdueTasksHtml(tasks) {
  const today = new Date();

  return `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <meta charset="UTF-8">
      ${getEmailStyles()}
    </head>
    <body>
      <div class="container">
        <div class="header" style="background: linear-gradient(135deg, #c5221f, #d93025);">
          <h1>🔴 تنبيه: مهام متأخرة</h1>
          <p>${tasks.length} مهمة تجاوزت موعدها النهائي</p>
        </div>

        <div class="content">
          <div class="alert alert-danger">
            <strong>⚠️ يرجى المتابعة الفورية!</strong><br>
            المهام التالية تجاوزت مواعيدها النهائية وتحتاج إلى إجراء عاجل.
          </div>

          <table>
            <tr>
              <th>المشروع</th>
              <th>المهمة</th>
              <th>المسؤول</th>
              <th>الموعد</th>
              <th>التأخير</th>
            </tr>
            ${tasks.map(task => {
              const daysOverdue = Math.ceil((today - new Date(task.dueDate)) / (1000 * 60 * 60 * 24));
              return `
                <tr>
                  <td>${task.projectCode || task.projectName || '-'}</td>
                  <td>${task.element || task.action || '-'}</td>
                  <td>${task.assignedTo || '-'}</td>
                  <td>${formatDate(task.dueDate)}</td>
                  <td><span class="badge badge-danger">${daysOverdue} يوم</span></td>
                </tr>
              `;
            }).join('')}
          </table>
        </div>

        <div class="footer">
          <p>تم إنشاء هذا التنبيه تلقائياً</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * توليد HTML لمهام الشخص
 * @param {Object} report - تقرير الشخص
 * @returns {string} كود HTML
 */
function generatePersonTasksHtml(report) {
  return `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <meta charset="UTF-8">
      ${getEmailStyles()}
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>📋 مهامك</h1>
          <p>${report.person} - ${formatDate(new Date())}</p>
        </div>

        <div class="content">
          <!-- ملخص -->
          <div class="section">
            <div style="text-align: center;">
              <div class="stat-box">
                <div class="stat-number">${report.summary.totalTasks}</div>
                <div class="stat-label">إجمالي</div>
              </div>
              <div class="stat-box">
                <div class="stat-number" style="color: #34a853;">${report.summary.completed}</div>
                <div class="stat-label">مكتمل</div>
              </div>
              <div class="stat-box">
                <div class="stat-number" style="color: #fbbc04;">${report.summary.inProgress}</div>
                <div class="stat-label">قيد العمل</div>
              </div>
              <div class="stat-box">
                <div class="stat-number" style="color: #ea4335;">${report.summary.overdue}</div>
                <div class="stat-label">متأخر</div>
              </div>
            </div>
          </div>

          <!-- المهام -->
          ${report.movements.length > 0 ? `
          <table>
            <tr>
              <th>المشروع</th>
              <th>المرحلة</th>
              <th>المهمة</th>
              <th>الحالة</th>
              <th>الموعد</th>
            </tr>
            ${report.movements.map(m => {
              const statusBadge = m.taskStatus === 'overdue' ? 'badge-danger' :
                                  m.taskStatus === 'inProgress' ? 'badge-warning' :
                                  m.taskStatus === 'completed' ? 'badge-success' : 'badge-info';
              return `
                <tr style="background: ${m.color};">
                  <td>${m.projectCode || m.projectName || '-'}</td>
                  <td>${m.stage || '-'}</td>
                  <td>${m.element || m.action || '-'}</td>
                  <td><span class="badge ${statusBadge}">${m.status || '-'}</span></td>
                  <td>${m.dueDate ? formatDate(m.dueDate) : '-'}</td>
                </tr>
              `;
            }).join('')}
          </table>
          ` : '<p>لا توجد مهام مسندة إليك حالياً</p>'}
        </div>

        <div class="footer">
          <p>تم إنشاء هذا التقرير تلقائياً</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

// ====================================================
// إدارة الـ Triggers
// ====================================================

/**
 * إعداد trigger يومي
 */
function setupDailyTrigger() {
  try {
    // حذف الـ triggers القديمة للتقرير اليومي
    removeTriggerByFunction('sendDailyReport');

    const settings = getNotificationSettings();
    const time = settings.dailyTime || '09:00';
    const [hours, minutes] = time.split(':').map(Number);

    // إنشاء trigger جديد
    ScriptApp.newTrigger('sendDailyReport')
      .timeBased()
      .atHour(hours)
      .nearMinute(minutes)
      .everyDays(1)
      .create();

    // حفظ حالة التفعيل
    saveNotificationSetting(NOTIFICATION_PROPS.ENABLED_DAILY, 'true');

    showToast('تم تفعيل الإشعارات اليومية', 'نجاح');
    return { success: true };

  } catch (error) {
    console.error('Error setting up daily trigger:', error);
    showError('خطأ في إعداد الإشعارات اليومية: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * إعداد trigger أسبوعي
 */
function setupWeeklyTrigger() {
  try {
    // حذف الـ triggers القديمة للتقرير الأسبوعي
    removeTriggerByFunction('sendWeeklyReport');

    const settings = getNotificationSettings();
    const day = settings.weeklyDay || 'SUNDAY';

    // إنشاء trigger جديد
    ScriptApp.newTrigger('sendWeeklyReport')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay[day])
      .atHour(9)
      .create();

    // حفظ حالة التفعيل
    saveNotificationSetting(NOTIFICATION_PROPS.ENABLED_WEEKLY, 'true');

    showToast('تم تفعيل الإشعارات الأسبوعية', 'نجاح');
    return { success: true };

  } catch (error) {
    console.error('Error setting up weekly trigger:', error);
    showError('خطأ في إعداد الإشعارات الأسبوعية: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * تفعيل جميع الإشعارات التلقائية
 */
function enableAutoNotifications() {
  setupDailyTrigger();
  setupWeeklyTrigger();
  showSuccess('تم تفعيل الإشعارات التلقائية');
}

/**
 * إيقاف جميع الإشعارات التلقائية
 */
function disableAutoNotifications() {
  removeTriggers();
  saveNotificationSetting(NOTIFICATION_PROPS.ENABLED_DAILY, 'false');
  saveNotificationSetting(NOTIFICATION_PROPS.ENABLED_WEEKLY, 'false');
  showSuccess('تم إيقاف الإشعارات التلقائية');
}

/**
 * حذف جميع الـ triggers
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const functionName = trigger.getHandlerFunction();
    if (functionName === 'sendDailyReport' || functionName === 'sendWeeklyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * حذف trigger بناءً على اسم الدالة
 * @param {string} functionName - اسم الدالة
 */
function removeTriggerByFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ====================================================
// إدارة إعدادات الإشعارات
// ====================================================

/**
 * الحصول على إعدادات الإشعارات
 * @returns {Object} إعدادات الإشعارات
 */
function getNotificationSettings() {
  const props = PropertiesService.getScriptProperties();

  return {
    dailyEnabled: props.getProperty(NOTIFICATION_PROPS.ENABLED_DAILY) === 'true',
    weeklyEnabled: props.getProperty(NOTIFICATION_PROPS.ENABLED_WEEKLY) === 'true',
    recipients: props.getProperty(NOTIFICATION_PROPS.RECIPIENTS) || '',
    dailyTime: props.getProperty(NOTIFICATION_PROPS.DAILY_TIME) || '09:00',
    weeklyDay: props.getProperty(NOTIFICATION_PROPS.WEEKLY_DAY) || 'SUNDAY'
  };
}

/**
 * حفظ إعداد إشعار
 * @param {string} key - مفتاح الإعداد
 * @param {string} value - القيمة
 */
function saveNotificationSetting(key, value) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(key, value);
}

/**
 * حفظ جميع إعدادات الإشعارات
 * @param {Object} settings - الإعدادات
 */
function saveNotificationSettings(settings) {
  const props = PropertiesService.getScriptProperties();

  props.setProperty(NOTIFICATION_PROPS.ENABLED_DAILY, settings.dailyEnabled ? 'true' : 'false');
  props.setProperty(NOTIFICATION_PROPS.ENABLED_WEEKLY, settings.weeklyEnabled ? 'true' : 'false');
  props.setProperty(NOTIFICATION_PROPS.RECIPIENTS, settings.recipients || '');
  props.setProperty(NOTIFICATION_PROPS.DAILY_TIME, settings.dailyTime || '09:00');
  props.setProperty(NOTIFICATION_PROPS.WEEKLY_DAY, settings.weeklyDay || 'SUNDAY');

  return { success: true };
}

/**
 * عرض نافذة إعدادات الإشعارات
 */
function showNotificationSettings() {
  const settings = getNotificationSettings();

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; direction: rtl; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; font-weight: bold; }
        input[type="text"], input[type="time"], select {
          width: 100%;
          padding: 10px;
          border: 1px solid #ddd;
          border-radius: 6px;
          box-sizing: border-box;
        }
        .checkbox-group { display: flex; align-items: center; margin-bottom: 15px; }
        .checkbox-group input { margin-left: 10px; width: 20px; height: 20px; }
        .btn {
          padding: 12px 24px;
          border: none;
          border-radius: 6px;
          cursor: pointer;
          margin-left: 10px;
          font-size: 14px;
        }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-secondary { background: #5f6368; color: white; }
        .btn-container { text-align: left; margin-top: 30px; }
        .hint { font-size: 12px; color: #666; margin-top: 5px; }
        h3 { margin-top: 0; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px; }
      </style>
    </head>
    <body>
      <h3>⚙️ إعدادات الإشعارات</h3>

      <div class="form-group">
        <label>البريد الإلكتروني للإشعارات *</label>
        <input type="text" id="recipients" value="${settings.recipients}" placeholder="email@example.com">
        <div class="hint">يمكنك إدخال عدة عناوين مفصولة بفاصلة</div>
      </div>

      <div class="form-group">
        <div class="checkbox-group">
          <input type="checkbox" id="dailyEnabled" ${settings.dailyEnabled ? 'checked' : ''}>
          <label for="dailyEnabled" style="margin-bottom: 0;">تفعيل التقرير اليومي</label>
        </div>
      </div>

      <div class="form-group">
        <label>وقت إرسال التقرير اليومي</label>
        <input type="time" id="dailyTime" value="${settings.dailyTime}">
      </div>

      <div class="form-group">
        <div class="checkbox-group">
          <input type="checkbox" id="weeklyEnabled" ${settings.weeklyEnabled ? 'checked' : ''}>
          <label for="weeklyEnabled" style="margin-bottom: 0;">تفعيل التقرير الأسبوعي</label>
        </div>
      </div>

      <div class="form-group">
        <label>يوم إرسال التقرير الأسبوعي</label>
        <select id="weeklyDay">
          <option value="SUNDAY" ${settings.weeklyDay === 'SUNDAY' ? 'selected' : ''}>الأحد</option>
          <option value="MONDAY" ${settings.weeklyDay === 'MONDAY' ? 'selected' : ''}>الاثنين</option>
          <option value="TUESDAY" ${settings.weeklyDay === 'TUESDAY' ? 'selected' : ''}>الثلاثاء</option>
          <option value="WEDNESDAY" ${settings.weeklyDay === 'WEDNESDAY' ? 'selected' : ''}>الأربعاء</option>
          <option value="THURSDAY" ${settings.weeklyDay === 'THURSDAY' ? 'selected' : ''}>الخميس</option>
          <option value="FRIDAY" ${settings.weeklyDay === 'FRIDAY' ? 'selected' : ''}>الجمعة</option>
          <option value="SATURDAY" ${settings.weeklyDay === 'SATURDAY' ? 'selected' : ''}>السبت</option>
        </select>
      </div>

      <div class="btn-container">
        <button class="btn btn-secondary" onclick="google.script.host.close()">إلغاء</button>
        <button class="btn btn-primary" onclick="saveSettings()">حفظ</button>
      </div>

      <script>
        function saveSettings() {
          const settings = {
            recipients: document.getElementById('recipients').value,
            dailyEnabled: document.getElementById('dailyEnabled').checked,
            dailyTime: document.getElementById('dailyTime').value,
            weeklyEnabled: document.getElementById('weeklyEnabled').checked,
            weeklyDay: document.getElementById('weeklyDay').value
          };

          if (!settings.recipients) {
            alert('يرجى إدخال البريد الإلكتروني');
            return;
          }

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                alert('تم حفظ الإعدادات بنجاح');
                google.script.host.close();
              } else {
                alert('خطأ: ' + result.error);
              }
            })
            .saveAndApplyNotificationSettings(settings);
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(450)
  .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'إعدادات الإشعارات');
}

/**
 * حفظ وتطبيق إعدادات الإشعارات
 * @param {Object} settings - الإعدادات
 * @returns {Object} نتيجة العملية
 */
function saveAndApplyNotificationSettings(settings) {
  try {
    // حفظ الإعدادات
    saveNotificationSettings(settings);

    // تطبيق التغييرات على الـ triggers
    if (settings.dailyEnabled) {
      setupDailyTrigger();
    } else {
      removeTriggerByFunction('sendDailyReport');
    }

    if (settings.weeklyEnabled) {
      setupWeeklyTrigger();
    } else {
      removeTriggerByFunction('sendWeeklyReport');
    }

    return { success: true };

  } catch (error) {
    console.error('Error saving notification settings:', error);
    return { success: false, error: error.message };
  }
}

// ====================================================
// دوال مساعدة
// ====================================================

/**
 * التحقق من صحة البريد الإلكتروني
 * @param {string} email - البريد الإلكتروني
 * @returns {boolean} هل البريد صالح
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * إرسال التقرير اليومي الآن (يدوياً)
 */
function sendDailyReportNow() {
  const settings = getNotificationSettings();
  if (!settings.recipients) {
    showError('يرجى إعداد البريد الإلكتروني أولاً من إعدادات الإشعارات');
    return;
  }

  sendDailyReport();
  showSuccess('تم إرسال التقرير اليومي');
}

/**
 * إرسال التقرير الأسبوعي الآن (يدوياً)
 */
function sendWeeklyReportNow() {
  const settings = getNotificationSettings();
  if (!settings.recipients) {
    showError('يرجى إعداد البريد الإلكتروني أولاً من إعدادات الإشعارات');
    return;
  }

  sendWeeklyReport();
  showSuccess('تم إرسال التقرير الأسبوعي');
}
