/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام إدارة الإنتاج - Seen Film
 * نظام التقارير (Reports System)
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * 1. عرض تقرير الخط الزمني للفيلم
 */
function showFilmTimelineReport() {
  const html = HtmlService.createTemplateFromFile('reports/FilmTimeline.html')
    .evaluate()
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'تقرير الخط الزمني للمشروع');
}

/**
 * 2. عرض التقارير المجمعة (أبحاث، فكسز، تصوير)
 */
function showCompanyReport() {
  const html = HtmlService.createTemplateFromFile('reports/CompanyReports.html')
    .evaluate()
    .setWidth(1000)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'تقارير الشركة المجمعة');
}

/**
 * 3. إنشاء تقرير تفصيلي (Sheet)
 */
function createDetailedFilmReport() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.DASHBOARD);
  let projectName = sheet.getRange('B3').getValue();

  if (!projectName || projectName === 'الكل') {
    // محاولة الحصول على المشروع من الصف النشط إذا كنا في شيت المشاريع أو الحركة
    const activeSheet = SpreadsheetApp.getActiveSheet();
    const activeRow = activeSheet.getActiveCell().getRow();
    
    if (activeSheet.getName() === SHEETS.PROJECTS && activeRow > 1) {
       projectName = activeSheet.getRange(activeRow, PROJECT_COLS.NAME).getValue();
    } else if (activeSheet.getName() === SHEETS.MOVEMENT && activeRow > 1) {
       projectName = activeSheet.getRange(activeRow, MOVEMENT_COLS.PROJECT).getValue();
    } else {
       ui.alert('الرجاء اختيار فيلم من الداشبورد أو الوقوف على صف الفيلم في شيت المشاريع.');
       return;
    }
  }
  
  if(!projectName) {
     ui.alert('لم يتم تحديد مشروع.');
     return;
  }

  showInfo('جاري إعداد التقرير التفصيلي لـ ' + projectName + '...');
  const url = generateDetailedFilmReport(projectName);
  
  if (url) {
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial; direction: rtl; text-align: center; padding: 20px;">
        <h3>✅ تم إنشاء التقرير بنجاح!</h3>
        <p>تم حفظ التقرير في ملف منفصل.</p>
        <a href="${url}" target="_blank" style="background: #1565C0; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">فتح التقرير</a>
      </div>
    `).setWidth(400).setHeight(200);
    ui.showModalDialog(html, 'التقرير التفصيلي');
  }
}

/**
 * Core Logic: Generate Detailed Spreadsheet Report
 */
function generateDetailedFilmReport(projectName) {
  const allMovements = getProjectMovements(projectName);
  
  // Create new Spreadsheet
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE || 'GMT', 'yyyy-MM-dd');
  const ss = SpreadsheetApp.create(`تقرير - ${projectName} - ${dateStr}`);
  const sheet = ss.getActiveSheet();
  sheet.setRightToLeft(true);
  
  // Header
  sheet.getRange('A1:E1').merge().setValue(`تقرير إنتاج تفصيلي: ${projectName}`)
    .setBackground(COLORS.HEADER).setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(16);
  
  sheet.getRange('A2:E2').merge().setValue(`تاريخ التقرير: ${dateStr}`)
    .setBackground(COLORS.BACKGROUND).setHorizontalAlignment('center');

  // Headers
  const headers = ['المرحلة', 'الخطوة (1-16)', 'المهمة/العنصر', 'المسؤول', 'الحالة'];
  sheet.getRange('A4:E4').setValues([headers])
    .setBackground('#E0E0E0').setFontWeight('bold').setBorder(true, true, true, true, true, true);

  let row = 5;
  
  // Define standard workflow order
  const workflow = [
    { phase: 'التطوير', step: 'الفكرة' },
    { phase: 'التطوير', step: 'البحث' },
    { phase: 'التطوير', step: 'المعالجة' },
    { phase: 'التطوير', step: 'اسكربت أولي' },
    { phase: 'التحضير', step: 'قائمة الضيوف' },
    { phase: 'التحضير', step: 'الفكسز' },
    { phase: 'التحضير', step: 'إعداد الأسئلة' },
    { phase: 'التحضير', step: 'تنسيق المدن' },
    { phase: 'التحضير', step: 'تنسيق الدراما' },
    { phase: 'الإنتاج', step: 'التصوير' },
    { phase: 'ما بعد التصوير', step: 'اسكربت نهائي' },
    { phase: 'ما بعد التصوير', step: 'تجهيز الأرشيف' },
    { phase: 'عناصر ما بعد الإنتاج', step: 'جرافيك' },
    { phase: 'عناصر ما بعد الإنتاج', step: 'مشاهد دراما' },
    { phase: 'عناصر ما بعد الإنتاج', step: 'الصوت' },
    { phase: 'المونتاج', step: 'المونتاج' } // Includes versions
  ];
  
  // Group movements by Subtype (Step)
  // Subtype in Config matches 'step' above mostly, but user data might vary. 
  // We check mapping.
  
  workflow.forEach(item => {
    // Header for the Step
    sheet.getRange(row, 1, 1, 5).setBackground('#F3F3F3');
    sheet.getRange(row, 1).setValue(item.phase).setFontWeight('bold');
    sheet.getRange(row, 2).setValue(item.step).setFontWeight('bold').setFontColor('#1565C0');
    
    // Find matching movements (Subtype == Step OR Stage == Phase if Step is generic)
    const tasks = allMovements.filter(m => {
      // Loose matching for flexibility
      const matchSubtype = m.subtype.includes(item.step) || (item.step === 'المونتاج' && m.stage === 'المونتاج');
      
      // Special check for Production types
      if (item.phase === 'الإنتاج' && item.step === 'التصوير') {
          return m.stage === 'الإنتاج'; // Include all production tasks under "Shooting" step in report for now
      }
      return matchSubtype;
    });

    if (tasks.length === 0) {
      // ⚠️ COVERAGE CHECK: If this is "City Coordination" (Planning) or "Shooting" (Execution), we might want to flag mismatch
      sheet.getRange(row, 3).setValue('⚠️ لا يوجد بيانات مسجلة').setFontColor('#E65100');
      row++;
    } else {
      // Print tasks
      
      // ... (Existing printing logic)
      
      tasks.forEach(t => {
         // ... (Printing rows)
         sheet.getRange(row, 1).setValue(''); 
         sheet.getRange(row, 2).setValue(''); 
         sheet.getRange(row, 3).setValue(t.element);
         sheet.getRange(row, 4).setValue(t.assignedTo);
         sheet.getRange(row, 5).setValue(t.status);
         
         if(t.status.includes('تم')) sheet.getRange(row, 5).setFontColor('green');
         if(t.status.includes('متأخر')) sheet.getRange(row, 5).setFontColor('red');
         
         row++;
      });
    }
  });

  // ═══════════════════════════════════════════════════════════════════════════════
  // تغطية المدن (City Coverage Check)
  // ═══════════════════════════════════════════════════════════════════════════════
  
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge().setValue('📊 تحليل تغطية المدن (التخطيط vs التنفيذ)')
       .setBackground(COLORS.HEADER).setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  row++;
  
  const headerRow = row;
  sheet.getRange(row, 1).setValue('المدينة');
  sheet.getRange(row, 2).setValue('حالة التخطيط (التحضير)');
  sheet.getRange(row, 3).setValue('حالة التنفيذ (الإنتاج)');
  sheet.getRange(row, 4).setValue('المطابقة');
  sheet.getRange(row, 1, 1, 5).setBackground('#EEE').setFontWeight('bold');
  row++;

  // 1. Get Planned Cities (From Pre-Production > City Coordination)
  const plannedCities = new Set();
  allMovements.filter(m => m.stage === 'التحضير' && m.subtype.includes('تنسيق المدن')).forEach(m => {
      // Assuming Element contains City Name
      if(m.element) plannedCities.add(m.element.trim());
  });

  // 2. Get Executed Cities (From Production > City Shoot)
  const executedCities = new Set();
  allMovements.filter(m => m.stage === 'الإنتاج' && m.subtype.includes('تصوير مدينة')).forEach(m => {
      if(m.element) executedCities.add(m.element.trim());
  });

  // Union of all cities
  const allCities = new Set([...plannedCities, ...executedCities]);
  
  if (allCities.size === 0) {
      sheet.getRange(row, 1, 1, 5).merge().setValue('لا توجد بيانات مدن مسجلة بعد.');
  } else {
      allCities.forEach(city => {
          const isPlanned = plannedCities.has(city);
          const isExecuted = executedCities.has(city);
          
          sheet.getRange(row, 1).setValue(city);
          sheet.getRange(row, 2).setValue(isPlanned ? '✅ موجود' : '❌ غير مخطط');
          sheet.getRange(row, 3).setValue(isExecuted ? '✅ تم التصوير' : '⏳ لم يتم بعد');
          
          let matchStatus = '';
          let matchColor = 'black';
          
          if (isPlanned && isExecuted) { matchStatus = '✅ متطابق'; matchColor = 'green'; }
          else if (isPlanned && !isExecuted) { matchStatus = '⚠️ باقي للتنفيذ'; matchColor = '#EF6C00'; } // Orange
          else if (!isPlanned && isExecuted) { matchStatus = '❓ غير مخطط (Ad-hoc)'; matchColor = 'purple'; }
          
          sheet.getRange(row, 4).setValue(matchStatus).setFontColor(matchColor);
          row++;
      });
  }

  // Formatting columns
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);

  return ss.getUrl();
}

// ... (Rest of existing functions getFilmTimelineData, etc.) ...

/**
 * API: جلب بيانات الخط الزمني لفيلم محدد (Updated for new Stages)
 */
function getFilmTimelineData(projectName) {
  if (!projectName) return null;

  const allMovements = getProjectMovements(projectName);
  
  // تعريف ترتيب المراحل الجديد
  // ترتيب ثابت للمراحل الستة
  const PHASE_ORDER = [
    'التطوير',
    'التحضير',
    'الإنتاج',
    'ما بعد التصوير',
    'عناصر ما بعد الإنتاج',
    'المونتاج',
    'التسليم'
  ];

  // تجميع البيانات للمخطط الزمني
  const timeline = PHASE_ORDER.map(phaseName => {
    // تصفية المهام الخاصة بهذه المرحلة
    const tasks = allMovements.filter(m => m.stage === phaseName);
    
    // إذا لم تكن هناك مهام، نعيد هيكل فارغ ولكن باسم المرحلة للحفاظ على التسلسل
    // أو نتخطاها إذا أردنا إخفاء المراحل الفارغة (حسب رغبة المستخدم)
    // هنا سنظهر المراحل لتوضيح خارطة الطريق
    
    const completedTasks = tasks.filter(m => m.status.includes('تم')).length;
    let stageStatus = 'pending';
    
    if (tasks.length > 0) {
        if (completedTasks === tasks.length) stageStatus = 'completed';
        else if (tasks.some(m => m.status.includes('جاري'))) stageStatus = 'active';
        else if (tasks.some(m => m.status.includes('متأخر'))) stageStatus = 'delayed';
    }

    return {
      name: phaseName,
      status: stageStatus,
      progress: tasks.length > 0 ? Math.round((completedTasks / tasks.length) * 100) : 0,
      tasks: tasks.map(t => ({
        element: `[${t.subtype || 'عام'}] ${t.element}`, // إضافة النوع الفرعي للعرض
        status: t.status,
        date: formatDate(t.date),
        assignedTo: t.assignedTo
      }))
    };
  });

  // حساب النسبة الكلية
  let totalTasks = allMovements.length;
  let completedTotal = allMovements.filter(m => m.status.includes('تم')).length;
  let completionPercentage = totalTasks > 0 ? Math.round((completedTotal / totalTasks) * 100) : 0;

  // الخطوة القادمة
  let nextStep = "لم يتم تحديد مهام بعد";
  const firstPending = allMovements.find(m => !m.status.includes('تم') && !m.status.includes('ملغي'));
  if (firstPending) {
    nextStep = `${firstPending.stage} > ${firstPending.subtype || ''} : ${firstPending.element}`;
  }

  return {
    projectName: projectName,
    timeline: timeline,
    overallProgress: completionPercentage,
    nextStep: nextStep
  };
}

// ... (Existing helper functions groupBy, formatDate) ...
function groupBy(array, key) {
  return array.reduce((result, currentValue) => {
    const k = currentValue[key] || 'غير محدد';
    (result[k] = result[k] || []).push(currentValue);
    return result;
  }, {});
}

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return d.toISOString().split('T')[0];
}

// ... (Existing getResearchAndFixingData, getFilmingLogisticsData functions should be kept or updated if stage names changed) ...
// Since we changed Stage Names in Config, we MUST update the filters here too.

function getResearchAndFixingData() {
  const allData = getAllMovements();
  
  // Update Filter: Research is now stage 'التطوير' and subtypes 'بحث', 'الأوراق' etc
  const researchData = allData.filter(m => 
    m.stage === 'التطوير' // New stage name
  );
  
  const researchByPerson = groupBy(researchData, 'assignedTo');
  
  // Update Filter: Fixing is stage 'التحضير' subtype 'الفكسز' or similar
  const fixingData = allData.filter(m => 
    m.stage === 'التحضير' || 
    m.subtype?.includes('تصريح') || 
    m.subtype?.includes('موافقة')
  );
  
  return {
    research: researchByPerson,
    fixing: fixingData
  };
}

function getFilmingLogisticsData() {
  const allData = getAllMovements();
  // Filter for Production Stage
  // We want to capture anything related to shooting
  const filmingData = allData.filter(m => m.stage === 'الإنتاج');
  
  const cityGroups = {};
  
  filmingData.forEach(task => {
    // If subtype is 'تصوير مدينة', the Element IS the city.
    // If subtype is 'تصوير دراما' or others, user might mention city in details or element.
    // For Matrix, we primarily look at 'تصوير مدينة' subtype or fallback to 'General' if not specified.
    
    let city = 'غير محدد';
    if (task.subtype.includes('مدينة') || task.subtype.includes('ميداني')) {
       city = task.element || 'غير محدد';
    } else {
       // For Drama/Inserts, we try to see if it's grouped. 
       // For now, put them in a separate bucket or 'General'
       city = 'أخرى / دراما / انسرتات';
    }
    
    if (!cityGroups[city]) {
      cityGroups[city] = {
        name: city,
        tasks: [],
        projects: new Set(),
        startDate: null,
        endDate: null
      };
    }
    
    cityGroups[city].tasks.push(task);
    cityGroups[city].projects.add(task.project);
    
    // Date Range Logic
    const dates = [];
    if (task.date) dates.push(new Date(task.date));
    if (task.dueDate) dates.push(new Date(task.dueDate));
    
    dates.forEach(d => {
       if (!cityGroups[city].startDate || d < cityGroups[city].startDate) cityGroups[city].startDate = d;
       if (!cityGroups[city].endDate || d > cityGroups[city].endDate) cityGroups[city].endDate = d;
    });
  });

  // Sort by date soonest first
  const reportData = Object.values(cityGroups).map(g => ({
    city: g.name,
    projectCount: g.projects.size,
    projectNames: Array.from(g.projects).join(', '),
    tasksCount: g.tasks.length,
    startDate: g.startDate ? formatDate(g.startDate) : 'غير محدد',
    endDate: g.endDate ? formatDate(g.endDate) : 'غير محدد',
    tasks: g.tasks
  }));
  
  reportData.sort((a, b) => {
     if(a.startDate === 'غير محدد') return 1;
     if(b.startDate === 'غير محدد') return -1;
     return new Date(a.startDate) - new Date(b.startDate);
  });
  
  return reportData;
}
