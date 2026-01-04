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
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE || 'GMT', 'dd-MM-yyyy');
  const ss = SpreadsheetApp.create(`تقرير - ${projectName} - ${dateStr}`);
  const sheet = ss.getActiveSheet();
  sheet.setRightToLeft(true);
  
  // Try to move report to project's Reports folder
  try {
    const projectsSheet = getSheet(SHEETS.PROJECTS);
    const projectRow = findRowByValue(projectsSheet, PROJECT_COLS.NAME, projectName);
    
    if (projectRow !== -1) {
      const projectCode = projectsSheet.getRange(projectRow, PROJECT_COLS.CODE).getValue();
      const mainFolderId = getProjectMainFolderId(projectCode);
      
      if (mainFolderId) {
        const mainFolder = DriveApp.getFolderById(mainFolderId);
        
        // Create or get Reports subfolder
        let reportsFolder = findFolderByName(mainFolder, 'Reports');
        if (!reportsFolder) {
          reportsFolder = mainFolder.createFolder('Reports');
        }
        
        // Move the spreadsheet to Reports folder
        const file = DriveApp.getFileById(ss.getId());
        file.moveTo(reportsFolder);
      }
    }
  } catch (e) {
    console.error('Could not move report to project folder:', e);
    // Continue anyway - report is still created
  }
  
  // Header
  sheet.getRange('A1:E1').merge().setValue(`تقرير إنتاج تفصيلي: ${projectName}`)
    .setBackground(COLORS.HEADER).setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(16);
  
  sheet.getRange('A2:E2').merge().setValue(`تاريخ التقرير: ${dateStr}`)
    .setBackground(COLORS.BACKGROUND).setHorizontalAlignment('center');

  // Headers
  const headers = ['المرحلة', 'النوع الفرعي', 'المهمة/العنصر', 'المسؤول', 'الحالة'];
  sheet.getRange('A4:E4').setValues([headers])
    .setBackground('#E0E0E0').setFontWeight('bold').setBorder(true, true, true, true, true, true);

  let row = 5;
  
  // Group all movements by stage
  const stageOrder = [
    STAGES.DEVELOPMENT,
    STAGES.PRE_PRODUCTION,
    STAGES.PRODUCTION,
    STAGES.POST_PAPERWORK,
    STAGES.POST_ELEMENTS,
    STAGES.EDITING,
    STAGES.COLORING,
    STAGES.DELIVERY
  ];
  
  stageOrder.forEach(stageObj => {
    const stageName = stageObj.name;
    const stageIcon = stageObj.icon;
    
    // Filter tasks for this stage
    const tasks = allMovements.filter(m => normalizeString(m.stage) === normalizeString(stageName));
    
    if (tasks.length === 0) {
      // Skip empty stages
      return;
    }
    
    // Stage Header
    sheet.getRange(row, 1, 1, 5).merge()
      .setValue(`${stageIcon} ${stageName}`)
      .setBackground(COLORS.HEADER)
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setFontSize(14);
    row++;
    
    // Print tasks
    tasks.forEach(t => {
      sheet.getRange(row, 1).setValue(t.subtype || '-');
      sheet.getRange(row, 2).setValue('');
      sheet.getRange(row, 3).setValue(t.element);
      sheet.getRange(row, 4).setValue(t.assignedTo);
      const tStatus = String(t.status || '');
      sheet.getRange(row, 5).setValue(tStatus);

      if(tStatus.includes('تم')) sheet.getRange(row, 5).setFontColor('green');
      if(tStatus.includes('متأخر')) sheet.getRange(row, 5).setFontColor('red');
      if(tStatus.includes('جاري')) sheet.getRange(row, 5).setFontColor('blue');
      
      row++;
    });
    
    row++; // Add spacing between stages
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
  const preProdName = normalizeString(STAGES.PRE_PRODUCTION.name);
  const plannedCities = new Set();
  allMovements.filter(m => normalizeString(m.stage) === preProdName && normalizeString(m.subtype).includes('تنسيق المدن')).forEach(m => {
      if(m.element) plannedCities.add(m.element.trim());
  });

  // 2. Get Executed Cities (From Production > City Shoot) with status tracking
  const prodName = normalizeString(STAGES.PRODUCTION.name);
  const cityStatus = new Map(); // Map<cityName, status>
  
  allMovements.filter(m => normalizeString(m.stage) === prodName && normalizeString(m.subtype).includes('تصوير')).forEach(m => {
      if(m.element) {
        const city = m.element.trim();
        const status = normalizeString(m.status);
        
        // Priority: تم > جاري > لم يبدأ
        if (!cityStatus.has(city) || status.includes('تم')) {
          cityStatus.set(city, m.status);
        } else if (status.includes('جاري') && !normalizeString(cityStatus.get(city)).includes('تم')) {
          cityStatus.set(city, m.status);
        }
      }
  });

  // Union of all cities
  const allCities = new Set([...plannedCities, ...cityStatus.keys()]);
  
  if (allCities.size === 0) {
      sheet.getRange(row, 1, 1, 5).merge().setValue('لا توجد بيانات مدن مسجلة بعد.');
  } else {
      allCities.forEach(city => {
          const isPlanned = plannedCities.has(city);
          const executionStatus = cityStatus.get(city);
          
          sheet.getRange(row, 1).setValue(city);
          sheet.getRange(row, 2).setValue(isPlanned ? '✅ موجود' : '❌ غير مخطط');
          
          // Execution status based on actual task status
          let execText = '⏳ لم يتم بعد';
          let execColor = 'gray';
          
          if (executionStatus) {
            const normStatus = normalizeString(executionStatus);
            if (normStatus.includes('تم')) {
              execText = '✅ تم التصوير';
              execColor = 'green';
            } else if (normStatus.includes('جاري')) {
              execText = '🔄 جاري التصوير';
              execColor = 'blue';
            } else if (normStatus.includes('لم يبدأ')) {
              execText = '⏳ مخطط للتصوير';
              execColor = 'orange';
            } else {
              execText = executionStatus; // Show actual status
              execColor = 'black';
            }
          }
          
          sheet.getRange(row, 3).setValue(execText).setFontColor(execColor);
          
          // Match status
          let matchStatus = '';
          let matchColor = 'black';
          
          if (isPlanned && executionStatus && normalizeString(executionStatus).includes('تم')) {
            matchStatus = '✅ متطابق';
            matchColor = 'green';
          } else if (isPlanned && executionStatus && normalizeString(executionStatus).includes('جاري')) {
            matchStatus = '🔄 جاري التنفيذ';
            matchColor = 'blue';
          } else if (isPlanned && !executionStatus) {
            matchStatus = '⚠️ باقي للتنفيذ';
            matchColor = '#EF6C00';
          } else if (!isPlanned && executionStatus) {
            matchStatus = '❓ غير مخطط (Ad-hoc)';
            matchColor = 'purple';
          }
          
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
  const PHASE_ORDER = Object.values(STAGES).map(s => s.name);

  // تجميع البيانات للمخطط الزمني
  const timeline = PHASE_ORDER.map(phaseName => {
    const normalizedPhase = normalizeString(phaseName);
    // تصفية المهام الخاصة بهذه المرحلة
    const tasks = allMovements.filter(m => normalizeString(m.stage) === normalizedPhase);
    
    // إذا لم تكن هناك مهام، نعيد هيكل فارغ ولكن باسم المرحلة للحفاظ على التسلسل
    // أو نتخطاها إذا أردنا إخفاء المراحل الفارغة (حسب رغبة المستخدم)
    // هنا سنظهر المراحل لتوضيح خارطة الطريق
    
    const completedTasks = tasks.filter(m => String(m.status || '').includes('تم')).length;
    let stageStatus = 'pending';

    if (tasks.length > 0) {
        if (completedTasks === tasks.length) stageStatus = 'completed';
        else if (tasks.some(m => String(m.status || '').includes('جاري'))) stageStatus = 'active';
        else if (tasks.some(m => String(m.status || '').includes('متأخر'))) stageStatus = 'delayed';
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
  let completedTotal = allMovements.filter(m => String(m.status || '').includes('تم')).length;
  let completionPercentage = totalTasks > 0 ? Math.round((completedTotal / totalTasks) * 100) : 0;

  // الخطوة القادمة
  let nextStep = "لم يتم تحديد مهام بعد";
  const firstPending = allMovements.find(m => {
    const s = String(m.status || '');
    return !s.includes('تم') && !s.includes('ملغي');
  });
  if (firstPending) {
    nextStep = `${firstPending.stage} > ${firstPending.subtype || ''} : ${firstPending.element}`;
  }

  // تطهير البيانات قبل الإرسال لمنع أخطاء السيرفر
  return sanitizeForClient({
    projectName: projectName,
    timeline: timeline,
    overallProgress: completionPercentage,
    nextStep: nextStep
  });
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
  try {
    console.log('Fetching Research and Fixing Data...');
    const allData = getAllMovements();
    if (!allData || allData.length === 0) {
      console.warn('No movements found for research report.');
      return { research: {}, fixing: [] };
    }

    const devStage = normalizeString(STAGES.DEVELOPMENT.name);
    const preProdStage = normalizeString(STAGES.PRE_PRODUCTION.name);
    
    const researchData = allData.filter(m => 
      normalizeString(m.stage) === devStage
    );
    
    const researchByPerson = groupBy(researchData, 'assignedTo');
    
    const fixingData = allData.filter(m => {
      const stage = normalizeString(m.stage);
      const subtype = normalizeString(m.subtype);
      return stage === preProdStage || 
             subtype.includes('تصريح') || 
             subtype.includes('موافقة');
    });
    
    console.log('Finished Fetching Research/Fixing. Research Count:', Object.keys(researchByPerson).length, 'Fixing Count:', fixingData.length);
    return sanitizeForClient({
      research: researchByPerson,
      fixing: fixingData
    });
  } catch (e) {
    console.error('Error in getResearchAndFixingData:', e);
    throw e;
  }
}

function getFilmingLogisticsData() {
  try {
    console.log('Fetching Filming Logistics Data...');
    const allData = getAllMovements();
    const prodStage = normalizeString(STAGES.PRODUCTION.name);
    const filmingData = allData.filter(m => normalizeString(m.stage) === prodStage);
    
    console.log('Filming Movements Found:', filmingData.length);
    const cityGroups = {};
    
    filmingData.forEach(task => {
      const subtype = normalizeString(task.subtype);
      let city = 'غير محدد';
      if (subtype.includes('مدينة') || subtype.includes('ميداني')) {
         city = task.element || 'غير محدد';
      } else {
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
      
      const dates = [];
      if (task.date) dates.push(new Date(task.date));
      if (task.dueDate) dates.push(new Date(task.dueDate));
      
      dates.forEach(d => {
         if (!cityGroups[city].startDate || d < cityGroups[city].startDate) cityGroups[city].startDate = d;
         if (!cityGroups[city].endDate || d > cityGroups[city].endDate) cityGroups[city].endDate = d;
      });
    });

    const reportData = Object.values(cityGroups).map(g => ({
      city: g.name,
      projectCount: g.projects.size,
      projectNames: Array.from(g.projects).join(', '),
      tasksCount: g.tasks.length,
      startDate: g.startDate ? formatDate(g.startDate) : 'غير محدد',
      endDate: g.endDate ? formatDate(g.endDate) : 'غير محدد',
      tasks: g.tasks.map(t => ({
          ...t,
          date: t.date ? formatDate(t.date) : '',
          dueDate: t.dueDate ? formatDate(t.dueDate) : ''
      }))
    }));
    
    reportData.sort((a, b) => {
       if(a.startDate === 'غير محدد') return 1;
       if(b.startDate === 'غير محدد') return -1;
       return new Date(a.startDate) - new Date(b.startDate);
    });
    
    console.log('Finished Filming Logistics. City Groups:', reportData.length);
    return sanitizeForClient(reportData);
  } catch (e) {
    console.error('Error in getFilmingLogisticsData:', e);
    throw e;
  }
}

/**
 * جلب بيانات مراحل ما بعد الإنتاج (الصوت، الجرافيك، المونتاج)
 */
function getPostProductionData() {
  try {
    console.log('Fetching Post-Production Data...');
    const allData = getAllMovements();
    const postElementsStage = normalizeString(STAGES.POST_ELEMENTS.name);
    const editingStage = normalizeString(STAGES.EDITING.name);
    
    const soundData = allData.filter(m => 
      normalizeString(m.stage) === postElementsStage && normalizeString(m.subtype).includes('صوت')
    );

    const graphicsData = allData.filter(m => 
      normalizeString(m.stage) === postElementsStage && normalizeString(m.subtype).includes('جرافيك')
    );

    const editingData = allData.filter(m => 
      normalizeString(m.stage) === editingStage
    );

    console.log('Post-Production Found - Sound:', soundData.length, 'Graphics:', graphicsData.length, 'Editing:', editingData.length);
    return sanitizeForClient({
      sound: groupBy(soundData, 'project'),
      graphics: groupBy(graphicsData, 'project'),
      editing: groupBy(editingData, 'project')
    });
  } catch (e) {
    console.error('Error in getPostProductionData:', e);
    throw e;
  }
}

/**
 * جلب بيانات التسليم والأرشيف
 */
function getDeliveryData() {
  try {
    console.log('Fetching Delivery Data...');
    const allData = getAllMovements();
    const postPaperworkStage = normalizeString(STAGES.POST_PAPERWORK.name);
    const deliveryStage = normalizeString(STAGES.DELIVERY.name);
    
    const archiveData = allData.filter(m => 
      normalizeString(m.stage) === postPaperworkStage && normalizeString(m.subtype).includes('أرشيف')
    );

    const deliveryData = allData.filter(m => 
      normalizeString(m.stage) === deliveryStage
    );

    console.log('Delivery Found - Archive:', archiveData.length, 'Delivery:', deliveryData.length);
    return sanitizeForClient({
      archive: archiveData,
      delivery: groupBy(deliveryData, 'project')
    });
  } catch (e) {
    console.error('Error in getDeliveryData:', e);
    throw e;
  }
}

/**
 * جلب قائمة أعضاء الفريق (للأزرار)
 */
function getAllTeamMembers() {
   const sheet = getSheet(SHEETS.TEAM);
   const lastRow = getLastRowInColumn(sheet, TEAM_COLS.NAME);
   if (lastRow <= 1) return [];
   
   const names = sheet.getRange(2, TEAM_COLS.NAME, lastRow - 1, 1).getValues().flat();
   return names.filter(n => n); // Remove empty
}

/**
 * جلب بيانات عضو محدد (Workload)
 */
function getMemberWorkloadData(memberName) {
  if(!memberName) return [];
  const allData = getAllMovements();
  // Filter by assignedTo
  const tasks = allData.filter(m => {
    const s = String(m.status || '');
    return m.assignedTo === memberName && !s.includes('تم') && !s.includes('ملغي');
  });
  
  // Sort by date/deadline
  tasks.sort((a,b) => {
     const dateA = a.dueDate ? new Date(a.dueDate) : (a.date ? new Date(a.date) : new Date(8640000000000000));
     const dateB = b.dueDate ? new Date(b.dueDate) : (b.date ? new Date(b.date) : new Date(8640000000000000));
     return dateA - dateB;
  });
  
  return sanitizeForClient(tasks);
}

/**
 * إنشاء شيت للطباعة (Export Printable Sheet)
 */
function createPrintableWorkloadSheet(memberName) {
  const tasks = getMemberWorkloadData(memberName);
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE || 'GMT', 'yyyy-MM-dd');
  
  // Create spreadsheet
  const ss = SpreadsheetApp.create(`مهام - ${memberName} - ${dateStr}`);
  const sheet = ss.getActiveSheet();
  sheet.setRightToLeft(true);
  
  // Header
  sheet.getRange('A1:E1').merge().setValue(`تقرير مهام: ${memberName}`)
       .setBackground('#1565C0').setFontColor('white').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  
  sheet.getRange('A2:E2').merge().setValue(`تاريخ الاستخراج: ${dateStr}`)
       .setHorizontalAlignment('center').setBackground('#E3F2FD');

  // Table Headers
  const headers = ['المشروع', 'المهمة / العنصر', 'التاريخ / الموعد', 'الحالة', 'ملاحظات'];
  sheet.getRange('A4:E4').setValues([headers])
       .setBackground('#EEEEEE').setFontWeight('bold').setBorder(true, true, true, true, true, true);
  
  // Data
  if (tasks.length > 0) {
    const rows = tasks.map(t => [
      t.project,
      `[${t.stage} > ${t.subtype}] ${t.element}`,
      t.dueDate ? formatDate(t.dueDate) : formatDate(t.date),
      t.status,
      t.details || ''
    ]);
    
    sheet.getRange(5, 1, rows.length, 5).setValues(rows);
    sheet.getRange(5, 1, rows.length, 5).setBorder(true, true, true, true, true, true);
  } else {
    sheet.getRange('A5:E5').merge().setValue('لا توجد مهام نشطة حالياً.').setHorizontalAlignment('center');
  }

  // Formatting
  sheet.setColumnWidth(1, 150); // Project
  sheet.setColumnWidth(2, 300); // Task
  sheet.setColumnWidth(3, 100); // Date
  sheet.setColumnWidth(4, 100); // Status
  sheet.setColumnWidth(5, 200); // Notes
  
  return ss.getUrl();
}
