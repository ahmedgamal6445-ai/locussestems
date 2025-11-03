/********************************************************************************
 * Locus Finance - Main Entry Point (V3 - Fully Modular)
 * التاريخ: 2024
 * الوصف: هذا الملف هو نقطة الدخول والتوجيه فقط.
 ********************************************************************************/

const SP = PropertiesService.getScriptProperties();
const CACHE = CacheService.getScriptCache();
const SS_APP = SpreadsheetApp;
const DRIVE_APP = DriveApp;

/**
 * يعالج طلبات GET ويوجه المستخدم إلى الواجهة المناسبة بناءً على دوره.
 * @param {object} e - The event parameter.
 * @returns {HtmlOutput} The HTML page to render.
 */
function doGet(e) {
  if (e.parameter.page === 'dashboard') {
    const token = e.parameter.token;
    try {
      const user = getSession(token);
      let templateFile;

      switch (user.role) {
        case 'Admin': templateFile = 'Admin_Dashboard'; break;
        case 'Owner': templateFile = 'Owner_Dashboard'; break;
        case 'Finance Manager': templateFile = 'FinanceManager_Dashboard'; break;
        case 'HR Manager': templateFile = 'HR_Dashboard'; break;
        case 'Branch Manager': templateFile = 'BranchManager_Dashboard'; break;
        case 'Operation Manager': templateFile = 'OperationManager_Dashboard'; break;
        case 'Sales Manager': templateFile = 'SalesManager_Dashboard'; break;
        case 'Sales': templateFile = 'Sales_Dashboard'; break;
        case 'Sales Follow Up': templateFile = 'SalesFollowUp_Dashboard'; break;
        case 'Business Developer': templateFile = 'BusinessDeveloper_Dashboard'; break;
        default: templateFile = 'Dashboard'; break; // For 'Employee' or other roles
      }

      const template = HtmlService.createTemplateFromFile(templateFile);
      return template.evaluate()
        .setTitle(`Locus - ${user.role} Dashboard`)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } catch (error) {
      Logger.log(`Session error in doGet: ${error.message}`);
      return HtmlService.createTemplateFromFile('Login').evaluate().setTitle('Locus - Login').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
  return HtmlService.createTemplateFromFile('Login').evaluate().setTitle('Locus - Login').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * يعالج طلبات POST (تستخدم حالياً للتحقق من التوكن بين التطبيقات).
 * @param {object} e - The event parameter.
 * @returns {ContentService.TextOutput} A JSON response.
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    if (params.action === 'verifyHandshakeToken') {
      const userData = verifyHandshakeToken(params.token);
      return ContentService.createTextOutput(JSON.stringify({ user: userData })).setMimeType(ContentService.MimeType.JSON);
    }
    throw new Error("Unknown action in doPost.");
  } catch (error) {
    Logger.log(`Error in doPost: ${error.message}`);
    return ContentService.createTextOutput(JSON.stringify({ error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================================================================
// SETUP & WRAPPER FUNCTIONS (RUNNABLE FROM EDITOR)
// ==================================================================

function initSystemProperties() {
  const properties = {
    MASTER_SETTING_FILE_ID: "1xsWLVw9UwElhcjcl-QzdMZyArMdkOAYAZ34fOwU_CnU",
    EMPLOYEES_FILE_ID: "1NQpwAJICSnPQyo5FsRXh_8fwE90SLzNjfP9sI-EDN6A",
    LOCUS_CUSTOMER_FILE_ID: "1b_DgOxAYY-gRdfL2h_rKCHYOT1ugQuExgCUUYa8vd0",
    LOCUS_CONTRACTS_SHEET_NAME: "Contracts",
    ROOT_BRANCHES_FOLDER_ID: "1ZzEu_o2vXDPXw21csRfadvkH7c_omdT3",
    HR_FILE_ID: "1aEmncgXL0_B0P8DYgiKpHUwi1nbMDx1oCX-KuABXQso",
    ATTENDANCE_FILE_ID: "1X40wKEXc16E8RbS2L95UgFoe9zlhblOo1wFN8JHCb6Y",
    FINANCE_DB_FILE_ID: "19sQhJMqtDvdGcQm02agrIZFWfzzlJT47a17iowkdIQA",
    CONTRACT_APP_URL: "https://script.google.com/macros/s/AKfycbzxwmjdtd3dcFGT-a1uzW-C53lYtrCF8RGXvG3AIY1HqBp1CI0JEZy35ulz0WRwY0dO_g/exec"
  };
  SP.setProperties(properties, true);
  Logger.log('✅ تم تهيئة متغيرات النظام بنجاح.');
  _ensureFinanceDB();
  Logger.log('✅ تم فحص وتهيئة قاعدة البيانات Finance_DB.');
}

function installTriggers(){
  const triggers = ScriptApp.getProjectTriggers();
  // حذف التريجرز القديمة لمنع التكرار
  triggers.forEach(t => {
    if (['syncFromLocus', 'dailyKpiSnapshot', 'hrRecalcPayroll', 'cleanCache', 'processDailyAttendance'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('syncFromLocus').timeBased().atHour(3).everyDays(1).create();
  ScriptApp.newTrigger('dailyKpiSnapshot').timeBased().atHour(4).everyDays(1).create();
  ScriptApp.newTrigger('hrRecalcPayroll').timeBased().onMonthDay(1).atHour(5).create();
  ScriptApp.newTrigger('cleanCache').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('processDailyAttendance').timeBased().atHour(2).everyDays(1).create();
  
  Logger.log('✅ تم تثبيت/تحديث التريجرز بنجاح.');
  return { success: true, message: 'تم تثبيت/تحديث التريجرز بنجاح.' };
}

function cleanCache(){ 
  CACHE.removeAll([]); 
  Logger.log('✅ تم مسح الكاش.'); 
  return true; 
}

function projectSelfTest(){
  const required = ['MASTER_SETTING_FILE_ID','EMPLOYEES_FILE_ID','LOCUS_CUSTOMER_FILE_ID','LOCUS_CONTRACTS_SHEET_NAME','ROOT_BRANCHES_FOLDER_ID','HR_FILE_ID','ATTENDANCE_FILE_ID','FINANCE_DB_FILE_ID'];
  const missing = [];
  const sp = PropertiesService.getScriptProperties();
  required.forEach(k=>{ if(!sp.getProperty(k)) missing.push(k); });
  if (missing.length === 0) return { ok: true, message: 'جميع خصائص النظام الأساسية موجودة.', props: sp.getProperties() };
  else return { ok: false, message: `خصائص النظام التالية مفقودة: ${missing.join(', ')}. يرجى تشغيل initSystemProperties.`, missing: missing, props: sp.getProperties() };
}
