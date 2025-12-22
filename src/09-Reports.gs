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
 * API: جلب بيانات الخط الزمني لفيلم محدد
 */
function getFilmTimelineData(projectName) {
  if (!projectName) return null;

  const allMovements = getProjectMovements(projectName);
  
  // ترتيب الحركات حسب التاريخ والمرحلة
  // نفترض وجود ترتيب للمراحل في Config، سنستخدم الترتيب الافتراضي للقائمة
  const stageOrder = STAGE_NAMES.reduce((acc, stage, index) => {
    acc[stage] = index;
    return acc;
  }, {});

  allMovements.sort((a, b) => {
    // أولاً حسب ترتيب المرحلة
    const stageDiff = (stageOrder[a.stage] || 99) - (stageOrder[b.stage] || 99);
    if (stageDiff !== 0) return stageDiff;
    // ثم حسب التاريخ
    return new Date(a.date) - new Date(b.date);
  });

  // تحليل الحالة الحالية
  let currentStageIndex = 0;
  let nextStep = "لم يبدأ المشروع بعد";
  let completionPercentage = 0;
  let totalTasks = allMovements.length;
  let completedTasks = 0;

  // تجميع البيانات للمخطط الزمني
  const timeline = STAGE_NAMES.map(stage => {
    const tasks = allMovements.filter(m => m.stage === stage);
    if (tasks.length === 0) return null; // تخطي المراحل الفارغة

    const completedInStage = tasks.filter(m => m.status.includes('تم')).length;
    completedTasks += completedInStage;
    
    // حالة المرحلة الكلية
    let stageStatus = 'pending';
    if (completedInStage === tasks.length && tasks.length > 0) stageStatus = 'completed';
    else if (tasks.some(m => m.status.includes('جاري'))) stageStatus = 'active';
    else if (tasks.some(m => m.status.includes('متأخر'))) stageStatus = 'delayed';

    return {
      name: stage,
      status: stageStatus,
      progress: Math.round((completedInStage / tasks.length) * 100),
      tasks: tasks.map(t => ({
        element: t.element,
        status: t.status,
        date: formatDate(t.date),
        assignedTo: t.assignedTo
      }))
    };
  }).filter(t => t !== null);

  if (totalTasks > 0) {
    completionPercentage = Math.round((completedTasks / totalTasks) * 100);
  }

  // تحديد الخطوة القادمة (أول مهمة غير مكتملة)
  const firstPending = allMovements.find(m => !m.status.includes('تم') && !m.status.includes('ملغي'));
  if (firstPending) {
    nextStep = `[${firstPending.stage}] ${firstPending.element}`;
    if (firstPending.assignedTo) nextStep += ` (${firstPending.assignedTo})`;
  } else if (totalTasks > 0 && completedTasks === totalTasks) {
    nextStep = "تم تسليم المشروع بالكامل! 🎉";
  }

  return {
    projectName: projectName,
    timeline: timeline,
    overallProgress: completionPercentage,
    nextStep: nextStep
  };
}

/**
 * API: جلب بيانات تقرير الأبحاث والفكسز
 */
function getResearchAndFixingData() {
  const allData = getAllMovements();
  
  // 1. تقرير الأبحاث (مرحلة الأوراق)
  const researchData = allData.filter(m => m.stage === 'الأوراق' || m.stage === 'أوراق');
  const researchByPerson = groupBy(researchData, 'assignedTo');
  
  // 2. تقرير الفكسز (أي مهمة تابعة للفيكسر أو مرحلة التصاريح إذا وجدت)
  // هنا سنفترض أنها المهام التي يقوم بها شخص دوره "فيكسر" أو نوعها "تصريح"
  // أو يمكننا البحث في مرحلة "التصوير" عن عناصر فرعية معينة.
  // للتبسيط، سنأخذ كل ما هو نوعه الفرعي "تصاريح" أو "لوجستيات" أو في الأوراق ونوعه "موافقات"
  const fixingData = allData.filter(m => 
    m.subtype?.includes('تصريح') || 
    m.subtype?.includes('موافقة') || 
    m.element?.includes('تصريح')
  );
  const fixingByStatus = groupBy(fixingData, 'status');

  return {
    research: researchByPerson,
    fixing: fixingData // سنرسل القائمة كاملة والفرونت إند يجمعها
  };
}

/**
 * API: جلب بيانات لوجستيات التصوير (مجمعة بالمدينة)
 */
function getFilmingLogisticsData() {
  const allData = getAllMovements();
  const filmingData = allData.filter(m => m.stage === 'التصوير' || m.stage === 'تصوير');
  
  // تجميع حسب المدينة (النوع الفرعي في التصوير هو المدينة عادة)
  const cityGroups = {};
  
  filmingData.forEach(task => {
    // المدينة هي الـ Subtype في التصوير
    const city = task.subtype || 'غير محدد'; // fallback if subtype is empty
    
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
    
    // حساب المدى الزمني
    // نعتمد هنا على تاريخ الاستحقاق كتاريخ نهائي، وتاريخ البداية (تاريخ الإدخال) كبداية
    // أو إذا كان هناك تاريخ محدد في التفاصيل. للتبسيط سنستخدم DueDate
    if (task.dueDate) {
      const d = new Date(task.dueDate);
      if (!cityGroups[city].endDate || d > cityGroups[city].endDate) cityGroups[city].endDate = d;
      // للبداية، سنفترض أنها قبل الموعد بأسبوع إذا لم يوجد تاريخ آخر، أو نستخدم تاريخ الإدخال
      // هنا سنستخدم تاريخ المهمة نفسه (date)
      const start = new Date(task.date);
      if (!cityGroups[city].startDate || start < cityGroups[city].startDate) cityGroups[city].startDate = start;
    }
  });

  // تحويل المجاميع إلى مصفوفة وتنسيق التواريخ
  return Object.values(cityGroups).map(g => ({
    city: g.name,
    projectCount: g.projects.size,
    projectNames: Array.from(g.projects).join(', '),
    tasksCount: g.tasks.length,
    startDate: g.startDate ? formatDate(g.startDate) : 'غير محدد',
    endDate: g.endDate ? formatDate(g.endDate) : 'غير محدد',
    tasks: g.tasks // التفاصيل
  }));
}

// --- Helpers internal to this file ---

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
  return d.toISOString().split('T')[0]; // yyyy-mm-dd
}
