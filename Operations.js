/********************************************************************************
 * Locus Finance - Operations Module
 * الوصف: يحتوي على دوال تنفيذ العمليات مثل إضافة دخل، تكلفة،
 * رفع مرفقات، والمزامنة.
 ********************************************************************************/

/**
 * يضيف سجل دخل جديد.
 * @param {string} token The user's session token.
 * @param {object} record The income record data.
 * @returns {object} A success message.
 */
function addIncome(token, record) {
  try {
    const user = getSession(token);
    const sheet = _getSheet(_getSSById(SP.getProperty('FINANCE_DB_FILE_ID')), 'Income');
    const entryDate = new Date(record.Date);
    entryDate.setHours(0, 0, 0, 0);

    const dayOfWeek = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), 'EEEE');
    const arabicDays = { 'Sunday': 'الأحد', 'Monday': 'الاثنين', 'Tuesday': 'الثلاثاء', 'Wednesday': 'الأربعاء', 'Thursday': 'الخميس', 'Friday': 'الجمعة', 'Saturday': 'السبت' };

    const newRowObject = {
      Entry_ID: _generateId('INC', record.Branch, entryDate, sheet),
      Branch: record.Branch, Date: entryDate, Day: arabicDays[dayOfWeek] || dayOfWeek,
      Service_Name: record.Service_Name, Amount: Number(record.Amount) || 0,
      Payment_Method: record.Payment_Method, Doctor_Transfer_No: Number(record.Doctor_Transfer_No) || 0,
      Service_Details: record.Service_Details || '', Notes: record.Notes || '',
      EmployeeName: user.name, EmployeeCode: user.code, Status: 'Pending', Timestamp: new Date()
    };

    _appendObjectAsRow(sheet, newRowObject);
    Logger.log(`New income entry [${newRowObject.Entry_ID}] added by ${user.name}.`);
    return { success: true, message: 'تم حفظ الدخل بنجاح.' };
  } catch (e) {
    Logger.log(`Error in addIncome: ${e.message}`);
    throw new Error(`فشل حفظ الدخل: ${e.message}`);
  }
}

/**
 * يضيف سجل تكلفة جديد.
 * @param {string} token The user's session token.
 * @param {object} record The cost record data.
 * @returns {object} A success message.
 */
function addCost(token, record) {
    try {
        const user = getSession(token);
        const sheet = _getSheet(_getSSById(SP.getProperty('FINANCE_DB_FILE_ID')), 'cost');
        const entryDate = new Date(record.Date);
        entryDate.setHours(0, 0, 0, 0);

        const dayOfWeek = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), 'EEEE');
        const arabicDays = { 'Sunday': 'الأحد', 'Monday': 'الاثنين', 'Tuesday': 'الثلاثاء', 'Wednesday': 'الأربعاء', 'Thursday': 'الخميس', 'Friday': 'الجمعة', 'Saturday': 'السبت' };

        const newRowObject = {
          Entry_ID: _generateId('CST', record.Branch, entryDate, sheet),
          Branch: record.Branch, Date: entryDate, Day: arabicDays[dayOfWeek] || dayOfWeek,
          Cost_Name: record.Cost_Name, Amount: Number(record.Amount) || 0,
          Payment_Method: record.Payment_Method, Cost_Details: record.Cost_Details || '',
          Notes: record.Notes || '', EmployeeCode: user.code, EmployeeName: user.name,
          Attachment_URL: record.Attachment_URL || '', Status: 'Pending', Timestamp: new Date()
        };

        _appendObjectAsRow(sheet, newRowObject);
        Logger.log(`New cost entry [${newRowObject.Entry_ID}] added by ${user.name}.`);
        return { success: true, message: 'تم حفظ التكلفة بنجاح.' };
    } catch (e) {
        Logger.log(`Error in addCost: ${e.message}`);
        throw new Error(`فشل حفظ التكلفة: ${e.message}`);
    }
}

/**
 * يرفع مرفقاً إلى جوجل درايف.
 * @param {string} token The user's session token.
 * @param {string} branch The branch name.
 * @param {string} fileName The name of the file.
 * @param {string} base64Data The file data encoded in Base64.
 * @returns {object} An object with the file URL and ID.
 */
function uploadCostAttachment(token, branch, fileName, base64Data) {
  const user = getSession(token);
  const allowedRoles = ['Admin', 'Owner', 'Finance Manager', 'Branch Manager'];
  if (!allowedRoles.includes(user.role)) throw new Error('غير مصرح لك.');
 
  const rootFolder = DRIVE_APP.getFolderById(SP.getProperty('ROOT_BRANCHES_FOLDER_ID'));
  const branchFolder = _ensureSubFolder(rootFolder, branch);
  const costsFolder = _ensureSubFolder(branchFolder, 'Costs');
  const now = new Date();
  const yearFolder = _ensureSubFolder(costsFolder, String(now.getFullYear()));
  const monthFolder = _ensureSubFolder(yearFolder, ('0' + (now.getMonth() + 1)).slice(-2));

  const decodedData = Utilities.base64Decode(base64Data);
  const extension = fileName.split('.').pop();
  const mimeType = _getMimeType(extension);
  const blob = Utilities.newBlob(decodedData, mimeType, fileName);
  const file = monthFolder.createFile(blob);
  
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { url: file.getUrl(), id: file.getId() };
}

/**
 * يزامن بيانات العقود من نظام Locus Customer.
 * @returns {object} A summary of the sync operation.
 */
function syncFromLocus() {
  const incomeSh = _getSheet(_getSSById(SP.getProperty('FINANCE_DB_FILE_ID')), 'Income');
  const existingContracts = new Set(_sheetDataToObjects(incomeSh).map(obj => String(obj.Contract_No || '').trim()).filter(Boolean));
  const allContractObjects = _sheetDataToObjects(_getSheet(_getSSById(SP.getProperty('LOCUS_CUSTOMER_FILE_ID')), SP.getProperty('LOCUS_CONTRACTS_SHEET_NAME')));
  
  const newIncomeRows = [];
  allContractObjects.forEach(contract => {
    const contractNo = String(contract.رقم_العقد || '').trim();
    if (!contractNo || existingContracts.has(contractNo)) return;

    const branch = String(contract['اسم الفرع'] || '').trim();
    const amount = Number(contract['قيمة العقد'] || 0);
    if (!branch || amount <= 0) return;

    const entryDate = new Date();
    entryDate.setHours(0, 0, 0, 0);
    const dayOfWeek = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), 'EEEE');
    const arabicDays = { 'Sunday': 'الأحد', 'Monday': 'الاثنين', 'Tuesday': 'الثلاثاء', 'Wednesday': 'الأربعاء', 'Thursday': 'الخميس', 'Friday': 'الجمعة', 'Saturday': 'السبت' };

    newIncomeRows.push({
      Entry_ID: _generateId('INC', branch, entryDate, incomeSh),
      Branch: branch, Date: entryDate, Day: arabicDays[dayOfWeek] || dayOfWeek,
      Service_Name: 'Contract Income', Amount: amount, Payment_Method: 'Locus Sync',
      Doctor_Transfer_No: 0, Service_Details: 'Imported from Locus Contracts', Notes: '',
      EmployeeName: String(contract.EmployeeName || '').trim(), Status: 'Pending', Timestamp: new Date()
    });
  });

  if (newIncomeRows.length > 0) {
    newIncomeRows.forEach(obj => _appendObjectAsRow(incomeSh, obj));
  }

  return { success: true, inserted: newIncomeRows.length, message: `تمت مزامنة ${newIncomeRows.length} عقد جديد من Locus.` };
}

/**
 * يأخذ لقطة يومية من مؤشرات الأداء الرئيسية.
 */
function dailyKpiSnapshot() {
  const kpiSheet = _getSheet(_getSSById(SP.getProperty('FINANCE_DB_FILE_ID')), 'KPI_Snapshots');
  const today = new Date();
  today.setHours(0,0,0,0);

  const globalStats = getGlobalStats(null, today.toISOString().split('T')[0], today.toISOString().split('T')[0]);

  _appendObjectAsRow(kpiSheet, {
    Date: today, Revenue: globalStats.revenue, Cost: globalStats.costs,
    Net: globalStats.net, Timestamp: new Date()
  });
  Logger.log('Daily KPI snapshot taken.');
}
