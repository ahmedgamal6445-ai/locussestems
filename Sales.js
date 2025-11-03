/********************************************************************************
 * Locus Finance - Sales & Follow-up Module
 * الوصف: يحتوي على كل الدوال المتعلقة بعمليات المبيعات والمتابعة والتهيئة.
 ********************************************************************************/

// ==================================================================
// CONFIGURATION
// ==================================================================

// !!! هام: الرجاء وضع ID ملف جوجل شيت الخاص بالمبيعات هنا
const SALES_SHEET_ID = "1ftJE_LDi3p3Yw3LtNRcc7udm-02sl3cnlT8NGMiwx0M";
// ==================================================================
// CLIENT-SIDE FUNCTIONS (Called from HTML)
// ==================================================================

/**
 * يجلب البيانات الأولية اللازمة لواجهات المبيعات (قوائم منسدلة وبيانات المستخدم).
 * @param {string} token The user's session token.
 * @returns {object} An object containing user data and sales master data.
 */
function getSalesInitialData(token) {
  const user = getSession(token);
  const msId = SP.getProperty('MASTER_SETTING_FILE_ID');
  const ss = _getSSById(msId);
  const salesSheet = _getSheet(ss, 'Sales');
  const branchesSheet = _getSheet(ss, 'Branch List');

  // تعديل: إضافة قوائم جديدة
  const masterData = { services: [], leadSources: [], branches: [], qualities: [], deals: [] };

  _sheetDataToObjects(salesSheet).forEach(row => {
    if (row.Service) masterData.services.push(row.Service);
    if (row.Lead_Source) masterData.leadSources.push(row.Lead_Source);
    if (row.Quality) masterData.qualities.push(row.Quality); // <-- إضافة Quality
    if (row.Deal) masterData.deals.push(row.Deal);           // <-- إضافة Deal
  });

  _sheetDataToObjects(branchesSheet).forEach(row => {
    if (row['اسم الفرع']) masterData.branches.push(row['اسم الفرع']);
  });
  
  // إزالة التكرار
  masterData.services = [...new Set(masterData.services)].sort();
  masterData.leadSources = [...new Set(masterData.leadSources)].sort();
  masterData.branches = [...new Set(masterData.branches)].sort();
  masterData.qualities = [...new Set(masterData.qualities)].sort();
  masterData.deals = [...new Set(masterData.deals)].sort();

  return { user, masterData };
}

/**
 * [Sales] يسجل عميل محتمل جديد في شيت المبيعات.
 * @param {string} token The user's session token.
 * @param {object} leadData بيانات العميل الجديد من الواجهة.
 * @returns {object} رسالة نجاح.
 */
function addNewLead(token, leadData) {
  try {
    const user = getSession(token); // user يحتوي على user.code و user.name
    if (!['Sales', 'Sales Manager', 'Admin'].includes(user.role)) {
      throw new Error('غير مصرح لك بتنفيذ هذه العملية.');
    }

    const salesSS = _getSSById(SALES_SHEET_ID);
    const salesSheet = _getSheet(salesSS, "SalesandFollowup");
    
    const now = new Date();
    // تعديل: إضافة الحقول الجديدة
    const newLead = {
      'Lead_ID': _generateId('LEAD', leadData.branch, now, salesSheet),
      'Timestamp': now,
      'Customer_Name': leadData.customerName,
      'Customer_Mobile': leadData.customerMobile,
      'Customer_National_ID': leadData.customerNationalId,
      'Branch': leadData.branch,
      'Service': leadData.service,
      'Lead_Source': leadData.leadSource,
      'Sales_Employee_Code': user.code, // <-- كود الموظف من الجلسة
      'Sales_Employee_Name': user.name, // <-- اسم الموظف من الجلسة
      'Quality': leadData.quality,      // <-- الحقل الجديد
      'Deal': leadData.deal,            // <-- الحقل الجديد
      'Sales_Feedback': leadData.feedback,
      'Deal_Status': 'Pending',
      'FollowUp_Needed': true,
    };

    _appendObjectAsRow(salesSheet, newLead);
    return { success: true, message: 'تم تسجيل العميل بنجاح.' };

  } catch (e) {
    Logger.log(`Error in addNewLead: ${e.message}`);
    throw new Error(`فشل تسجيل العميل: ${e.message}`);
  }
}
/**
 * [Follow-up] يجلب قائمة العملاء الذين يحتاجون إلى متابعة.
 * @param {string} token The user's session token.
 * @returns {Array<object>} قائمة العملاء للمتابعة.
 */
function getLeadsForFollowUp(token) {
  getSession(token);
  
  const salesSS = _getSSById(SALES_SHEET_ID);
  const allLeads = _sheetDataToObjects(_getSheet(salesSS, "SalesandFollowup"));
  
  const contractsSheet = _getSSById(SP.getProperty('LOCUS_CUSTOMER_FILE_ID')).getSheetByName(SP.getProperty('LOCUS_CONTRACTS_SHEET_NAME'));
  const contractedIDs = new Set(_sheetDataToObjects(contractsSheet).map(c => String(c['رقم اثبات الشخصية'] || '').trim()));

  const twoDaysAgo = new Date();
  twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);

  const leadsToFollow = allLeads.filter(lead => {
    const leadDate = new Date(lead.Timestamp);
    const nationalId = String(lead.Customer_National_ID || '').trim();
    
    return lead.Deal_Status === 'Pending' && leadDate <= twoDaysAgo && (!nationalId || !contractedIDs.has(nationalId));
  });

  return leadsToFollow;
}

/**
 * [Follow-up] تحديث بيانات عميل بعد إجراء مكالمة المتابعة.
 * @param {string} token The user's session token.
 * @param {string} leadId The ID of the lead to update.
 * @param {string} feedback The feedback from the follow-up call.
 * @returns {object} A success message.
 */
function updateFollowUp(token, leadId, feedback) {
  try {
    const user = getSession(token);
    if (!['Sales Follow Up', 'Sales Manager', 'Admin'].includes(user.role)) {
      throw new Error('غير مصرح لك.');
    }

    const salesSS = _getSSById(SALES_SHEET_ID);
    const sheet = _getSheet(salesSS, "SalesandFollowup");
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idCol = headers.indexOf('Lead_ID');
    const feedbackCol = headers.indexOf('FollowUp_Feedback');
    const empCodeCol = headers.indexOf('FollowUp_Employee_Code');
    const dateCol = headers.indexOf('FollowUp_Date');

    for (let i = 0; i < data.length; i++) {
      if (data[i][idCol] === leadId) {
        sheet.getRange(i + 2, feedbackCol + 1).setValue(feedback);
        sheet.getRange(i + 2, empCodeCol + 1).setValue(user.code);
        sheet.getRange(i + 2, dateCol + 1).setValue(new Date());
        return { success: true, message: 'تم تحديث المتابعة بنجاح.' };
      }
    }
    throw new Error('لم يتم العثور على العميل.');
  } catch (e) {
    Logger.log(`Error in updateFollowUp: ${e.message}`);
    throw new Error(`فشل تحديث المتابعة: ${e.message}`);
  }
}


// ==================================================================
// SETUP FUNCTION (Runnable from Editor)
// ==================================================================

/**
 * [SETUP] - دالة لتهيئة شيتات المبيعات والإعدادات الخاصة بها.
 * يتم تشغيلها يدوياً من محرر الأكواد عند الحاجة.
 */
function setupSalesEnvironment() {
  try {
    // --- الخطوة 1: تهيئة شيت الإعدادات الرئيسية (Master Setting) ---
    const msId = SP.getProperty('MASTER_SETTING_FILE_ID');
    if (!msId) throw new Error("MASTER_SETTING_FILE_ID is not defined in Script Properties.");
    const masterSS = _getSSById(msId);

    let salesSettingsSheet = masterSS.getSheetByName('Sales');
    if (!salesSettingsSheet) {
      salesSettingsSheet = masterSS.insertSheet('Sales');
      Logger.log('✅ تم إنشاء تاب "Sales" في Master Setting.');
    }
    
    // هذا الشرط يضمن عدم الكتابة فوق بياناتك الحالية
    if (salesSettingsSheet.getLastRow() < 2) {
      Logger.log('تاب "Sales" فارغة، سيتم إضافة الأعمدة وبيانات مبدئية...');
      const salesHeaders = ['Service', 'Lead_Source', 'Quality', 'Deal'];
      salesSettingsSheet.getRange(1, 1, 1, salesHeaders.length).setValues([salesHeaders]).setFontWeight('bold');
      
      const sampleData = [
        ['تركيب تقويم', 'Facebook', 'Excellent', 'Interested'],
        ['زراعة أسنان', 'Website', 'Good', 'Not Interested'],
        ['تبييض ليزر', 'Phone Call', 'Poor', 'Call Back Later'],
        ['حشو عصب', 'Instagram', '', ''],
        ['تنظيف جير', 'Walk-in', '', '']
      ];
      salesSettingsSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
      Logger.log('✅ تم إضافة بيانات مبدئية في تاب "Sales".');
    } else {
      Logger.log('✅ تاب "Sales" تحتوي على بيانات بالفعل، تم التخطي للحفاظ عليها.');
    }

    // --- الخطوة 1.5: التحقق من وجود تاب "Branch List" (للقراءة فقط) ---
    const branchListSheet = masterSS.getSheetByName('Branch List');
    if (!branchListSheet) {
      throw new Error('تاب "Branch List" غير موجودة في ملف Master Setting. هذه التاب ضرورية لعمل النظام، يرجى التأكد من وجودها.');
    }
    Logger.log('✅ تم العثور على تاب "Branch List" بنجاح.');


    // --- الخطوة 2: تهيئة شيت بيانات المبيعات (SalesandFollowup) ---
    if (!SALES_SHEET_ID || SALES_SHEET_ID === "YOUR_SALES_AND_FOLLOWUP_SHEET_ID") {
        throw new Error("Please set the SALES_SHEET_ID constant in Sales.gs first.");
    }
    const salesSS = _getSSById(SALES_SHEET_ID);
    
    let salesDataSheet = salesSS.getSheetByName('SalesandFollowup');
    if (!salesDataSheet) {
      salesDataSheet = salesSS.insertSheet('SalesandFollowup');
      Logger.log('✅ تم إنشاء تاب "SalesandFollowup".');
    }

    // هذا الجزء سيقوم دائماً بإعادة ضبط الأعمدة لضمان الهيكل الصحيح
    salesDataSheet.clear(); 
    const headers = [
      'Lead_ID', 'Timestamp', 'Customer_Name', 'Customer_Mobile', 'Customer_National_ID',
      'Branch', 'Service', 'Lead_Source', 
      'Sales_Employee_Code', 'Sales_Employee_Name',
      'Quality', 'Deal',
      'Sales_Feedback',
      'Deal_Status', 'FollowUp_Needed', 'FollowUp_Date', 'FollowUp_Employee_Code', 'FollowUp_Feedback'
    ];
    salesDataSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    salesDataSheet.setFrozenRows(1);
    
    Logger.log('✅ تم تهيئة الأعمدة في شيت "SalesandFollowup" بنجاح.');

    SpreadsheetApp.flush();
    Browser.msgBox('🎉 تم تهيئة بيئة المبيعات بنجاح!');
    return 'Success';

  } catch (e) {
    Logger.log(`🛑 ERROR in setupSalesEnvironment: ${e.message}`);
    Browser.msgBox(`فشل التهيئة: ${e.message}`);
    return `Error: ${e.message}`;
  }
}