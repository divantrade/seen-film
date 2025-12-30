/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * نظام التقارير الذكي (Smart Explorer Engines)
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * جلب البيانات الشاملة لجميع العمليات مع دمج بيانات المشاريع (القناة، البرنامج)
 */
function webAppGetSmartExplorerData() {
  const logPrefix = '[SmartExplorer] ';
  try {
    console.time(logPrefix + 'TotalTime');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. جلب بيانات الحركة
    console.time(logPrefix + 'MovementsFetch');
    const movements = getAllMovements();
    console.timeEnd(logPrefix + 'MovementsFetch');
    
    if (!movements) throw new Error('فشل جلب بيانات الحركة');

    // 2. جلب وتخزين بيانات المشاريع
    console.time(logPrefix + 'ProjectsFetch');
    const projectSheet = ss.getSheetByName(SHEETS.PROJECTS);
    const projectMap = {};
    
    if (projectSheet) {
      const pLastRow = projectSheet.getLastRow();
      const pLastCol = projectSheet.getLastColumn();
      
      if (pLastRow > 1) {
        const projectData = projectSheet.getRange(1, 1, pLastRow, pLastCol).getValues();
        
        // تحديد الأعمدة ديناميكياً
        const headers = projectData[0];
        const getIdx = (header) => {
          const search = normalizeString(header);
          return headers.findIndex(h => normalizeString(h) === search);
        };
        
        const pIdx = {
          name: getIdx('اسم الفيلم'),
          channel: getIdx('القناة'),
          program: getIdx('اسم البرنامج'),
          producer: getIdx('المنتج'),
          editor: getIdx('المونتير'),
          code: getIdx('الكود')
        };
        
        if (pIdx.name === -1) pIdx.name = pIdx.code; // Backup

        for (let i = 1; i < projectData.length; i++) {
          const row = projectData[i];
          const name = pIdx.name !== -1 ? row[pIdx.name] : '';
          if (name) {
            const normalizedName = normalizeString(name);
            projectMap[normalizedName] = {
              channel: pIdx.channel !== -1 ? row[pIdx.channel] || '-' : '-',
              program: pIdx.program !== -1 ? row[pIdx.program] || '-' : '-',
              producer: pIdx.producer !== -1 ? row[pIdx.producer] || '-' : '-',
              editor: pIdx.editor !== -1 ? row[pIdx.editor] || '-' : '-'
            };
          }
        }
      }
    }
    console.timeEnd(logPrefix + 'ProjectsFetch');
    
    // 3. دمج البيانات
    console.time(logPrefix + 'DataEnrichment');
    const enrichedData = movements.map(m => {
      const normalizedProject = normalizeString(m.project);
      const pInfo = projectMap[normalizedProject] || { channel: '-', program: '-', producer: '-', editor: '-' };
      return {
        ...m,
        channel: pInfo.channel,
        program: pInfo.program,
        producer: pInfo.producer,
        editor: pInfo.editor
      };
    });
    console.timeEnd(logPrefix + 'DataEnrichment');
    
    // 4. جلب القوائم الفرعية للفلاتر
    const filters = {
      projects: [...new Set(movements.map(m => m.project).filter(p => p))].sort(),
      cities: [...new Set(movements.map(m => m.city).filter(c => c))].sort(),
      members: [...new Set(movements.map(m => m.assignedTo).filter(a => a))].sort(),
      channels: [...new Set(Object.values(projectMap).map(p => p.channel).filter(c => c && c !== '-'))].sort(),
      stages: STAGE_NAMES || []
    };
    
    console.timeEnd(logPrefix + 'TotalTime');
    
    // تطهير البيانات قبل الإرسال لمنع أخطاء السيرفر
    const result = {
      data: enrichedData,
      filters: filters
    };
    
    return sanitizeForClient(result);
    
  } catch (e) {
    console.error('CRITICAL ERROR in webAppGetSmartExplorerData:', e);
    // نرمي الخطأ ليظهر في واجهة المستخدم
    throw new Error('فشل نظام التقارير: ' + e.message);
  }
}
