/********************************************************************************
 * Locus Finance - Finance Module
 * الوصف: يحتوي على كل الدوال المتعلقة بالعمليات المالية والتقارير.
 ********************************************************************************/

/**
 * يجلب البيانات الأولية اللازمة لبدء تشغيل الواجهة.
 * @param {string} token The user's session token.
 * @returns {object} An object containing user data, master data, and app URLs.
 */
function getInitialData(token) {
    const user = getSession(token);
    const masterData = getMasterFinanceData();
    const contractAppUrl = SP.getProperty('CONTRACT_APP_URL');
    return { user, masterData, contractAppUrl };
}

/**
 * يجلب البيانات الرئيسية من شيت الإعدادات.
 * @returns {object} An object with master data lists (services, costs, etc.).
 */
function getMasterFinanceData() {
  const msId = SP.getProperty('MASTER_SETTING_FILE_ID');
  const ss = _getSSById(msId);
  const result = { services: [], costCats: [], payMethods: [], salesChannels: [], branches: [] };

  const financeSheet = ss.getSheetByName('Finance');
  if (financeSheet) {
    const data = financeSheet.getDataRange().getValues().slice(1);
    data.forEach(row => {
      if (row[0]) result.services.push(String(row[0]));
      if (row[2]) result.costCats.push(String(row[2]));
      if (row[3]) result.payMethods.push(String(row[3]));
      if (row[4]) result.salesChannels.push(String(row[4]));
    });
    result.services = [...new Set(result.services)];
    result.costCats = [...new Set(result.costCats)];
    result.payMethods = [...new Set(result.payMethods)];
    result.salesChannels = [...new Set(result.salesChannels)];
  }

  const allItemsSet = new Set([...result.services, ...result.costCats]);
  allItemsSet.add('Marketing'); 
  result.allItems = Array.from(allItemsSet).sort();

  const branchSheet = ss.getSheetByName('Branch List');
  if (branchSheet) {
    const data = branchSheet.getRange("A2:A" + branchSheet.getLastRow()).getValues(); 
    result.branches = data.map(row => String(row[0])).filter(Boolean);
  }

  return result;
}

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
 * يحسب الإحصائيات العامة (الدخل، التكاليف، الربح).
 * @param {string} token The user's session token.
 * @param {string} startDateStr The start date for the filter.
 * @param {string} endDateStr The end date for the filter.
 * @returns {object} An object with global KPIs.
 */
function getGlobalStats(token, startDateStr, endDateStr) {
    if (token) { // Token might be null if called internally by a trigger
      const user = getSession(token);
      if (!['Owner', 'Admin'].includes(user.role)) throw new Error('غير مصرح لك بعرض هذه البيانات.');
    }

    let startDate, endDate;
    if (startDateStr && endDateStr) {
        startDate = new Date(startDateStr);
        endDate = new Date(endDateStr);
    } else {
        const now = new Date();
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    }
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);

    const dbId = SP.getProperty('FINANCE_DB_FILE_ID');
    const ss = _getSSById(dbId);
    const incomeObjects = _sheetDataToObjects(_getSheet(ss, 'Income'));
    const costObjects = _sheetDataToObjects(_getSheet(ss, 'cost'));

    const filteredIncome = incomeObjects.filter(entry => new Date(entry.Date) >= startDate && new Date(entry.Date) <= endDate);
    const filteredCost = costObjects.filter(entry => new Date(entry.Date) >= startDate && new Date(entry.Date) <= endDate);

    const revenue = filteredIncome.reduce((acc, entry) => acc + (Number(entry.Amount) || 0), 0);
    const costs = filteredCost.reduce((acc, entry) => acc + (Number(entry.Amount) || 0), 0);
    const net = revenue - costs;
    const marketingCost = filteredCost
        .filter(entry => String(entry.Cost_Name || '').trim().toLowerCase() === 'marketing')
        .reduce((acc, entry) => acc + (Number(entry.Amount) || 0), 0);
    
    const grossROI = costs > 0 ? (revenue / costs) : 0;
    const netROI = (costs + marketingCost) > 0 ? (revenue / (costs + marketingCost)) : 0;
    
    return { revenue, costs, net, grossROI: grossROI.toFixed(2), netROI: netROI.toFixed(2) };
}

/**
 * يجلب تقارير مفصلة للدخل والتكاليف لكل فرع.
 * @param {string} token The user's session token.
 * @param {string} filterItem The item to filter by.
 * @param {string} startDateStr The start date.
 * @param {string} endDateStr The end date.
 * @returns {object} A report object.
 */
function getDetailedReports(token, filterItem, startDateStr, endDateStr) {
    getSession(token);

    let startDate, endDate;
    if (startDateStr && endDateStr) {
        startDate = new Date(startDateStr);
        endDate = new Date(endDateStr);
    } else {
        const now = new Date();
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    }
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);

    const dbId = SP.getProperty('FINANCE_DB_FILE_ID');
    const ss = _getSSById(dbId);
    const filteredIncome = _sheetDataToObjects(_getSheet(ss, 'Income')).filter(e => new Date(e.Date) >= startDate && new Date(e.Date) <= endDate);
    const filteredCost = _sheetDataToObjects(_getSheet(ss, 'cost')).filter(e => new Date(e.Date) >= startDate && new Date(e.Date) <= endDate);

    const reports = {};
    const useFilter = filterItem && filterItem !== 'all';
    const processEntries = (entries, type) => {
      entries.forEach(entry => {
        const branch = entry.Branch;
        if (!branch) return;
        const itemName = type === 'income' ? entry.Service_Name : entry.Cost_Name;
        if (useFilter && itemName !== filterItem) return;
        if (!reports[branch]) reports[branch] = { income: 0, cost: 0 };
        reports[branch][type] += Number(entry.Amount || 0);
      });
    };

    processEntries(filteredIncome, 'income');
    processEntries(filteredCost, 'cost');
    
    return reports;
}

/**
 * يجهز البيانات اللازمة للرسوم البيانية.
 * @param {string} token The user's session token.
 * @param {string} startDateStr The start date.
 * @param {string} endDateStr The end date.
 * @returns {object} Data formatted for Google Charts.
 */
function getChartData(token, startDateStr, endDateStr) {
    const user = getSession(token);
    if (!['Owner', 'Admin'].includes(user.role)) throw new Error('غير مصرح لك بعرض هذه البيانات.');

    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);

    const dbId = SP.getProperty('FINANCE_DB_FILE_ID');
    const ss = _getSSById(dbId);
    const incomeObjects = _sheetDataToObjects(_getSheet(ss, 'Income'));
    const costObjects = _sheetDataToObjects(_getSheet(ss, 'cost'));

    const timeSeriesData = {};
    incomeObjects.forEach(entry => {
        const entryDate = new Date(entry.Date);
        if (entryDate >= startDate && entryDate <= endDate) {
            const dateKey = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!timeSeriesData[dateKey]) timeSeriesData[dateKey] = { income: 0, cost: 0 };
            timeSeriesData[dateKey].income += Number(entry.Amount || 0);
        }
    });
    costObjects.forEach(entry => {
        const entryDate = new Date(entry.Date);
        if (entryDate >= startDate && entryDate <= endDate) {
            const dateKey = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!timeSeriesData[dateKey]) timeSeriesData[dateKey] = { income: 0, cost: 0 };
            timeSeriesData[dateKey].cost += Number(entry.Amount || 0);
        }
    });
    const timeSeriesArray = [['التاريخ', 'الدخل', 'التكاليف'], ...Object.keys(timeSeriesData).sort().map(date => [date, timeSeriesData[date].income, timeSeriesData[date].cost])];

    const costBreakdownData = {};
    costObjects.forEach(entry => {
        const entryDate = new Date(entry.Date);
        if (entryDate >= startDate && entryDate <= endDate) {
            const costName = String(entry.Cost_Name || 'غير محدد').trim();
            costBreakdownData[costName] = (costBreakdownData[costName] || 0) + Number(entry.Amount || 0);
        }
    });
    const costBreakdownArray = [['بند التكلفة', 'المبلغ'], ...Object.keys(costBreakdownData).map(item => [item, costBreakdownData[item]])];

    return { timeSeries: timeSeriesArray, costBreakdown: costBreakdownArray };
}

/**
 * يجلب المعاملات المعلقة للمراجعة والاعتماد.
 * @param {string} token The user's session token.
 * @param {string} startDateStr The start date.
 * @param {string} endDateStr The end date.
 * @returns {object} An object containing pending income and cost entries.
 */
function getPendingEntries(token, startDateStr, endDateStr) {
    const user = getSession(token);
    if (!['Finance Manager', 'Admin', 'Owner'].includes(user.role)) throw new Error('غير مصرح لك.');

    let startDate, endDate;
    if (startDateStr && endDateStr) {
        startDate = new Date(startDateStr);
        endDate = new Date(endDateStr);
    } else {
        const now = new Date();
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    }
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);

    const dbId = SP.getProperty('FINANCE_DB_FILE_ID');
    const ss = _getSSById(dbId);
    const incomeObjects = _sheetDataToObjects(_getSheet(ss, 'Income'));
    const costObjects = _sheetDataToObjects(_getSheet(ss, 'cost'));

    const pending = { income: [], cost: [] };

    pending.income = incomeObjects
      .filter(entry => entry.Status === 'Pending' && new Date(entry.Date) >= startDate && new Date(entry.Date) <= endDate)
      .map(entry => ({
        id: entry.Entry_ID, branch: entry.Branch, date: Utilities.formatDate(new Date(entry.Date), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        service: entry.Service_Name, amount: entry.Amount, employee: entry.EmployeeName
      }));

    pending.cost = costObjects
      .filter(entry => entry.Status === 'Pending' && new Date(entry.Date) >= startDate && new Date(entry.Date) <= endDate)
      .map(entry => ({
        id: entry.Entry_ID, branch: entry.Branch, date: Utilities.formatDate(new Date(entry.Date), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        item: entry.Cost_Name, amount: entry.Amount, attachment: entry.Attachment_URL, employee: entry.EmployeeName
      }));

    return pending;
}

/**
 * يقوم باعتماد معاملة معلقة.
 * @param {string} token The user's session token.
 * @param {string} entryId The ID of the entry to approve.
 * @param {string} entryType The type of entry ('income' or 'cost').
 * @returns {object} A success message.
 */
function approveEntry(token, entryId, entryType) {
    try {
        const user = getSession(token);
        if (!['Finance Manager', 'Admin', 'Owner'].includes(user.role)) throw new Error('غير مصرح لك.');

        const sheetName = entryType === 'income' ? 'Income' : 'cost';
        const sheet = _getSheet(_getSSById(SP.getProperty('FINANCE_DB_FILE_ID')), sheetName);
        
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const entryIdColIndex = headers.indexOf('Entry_ID');
        const statusColIndex = headers.indexOf('Status');

        if (entryIdColIndex === -1 || statusColIndex === -1) throw new Error('لم يتم العثور على أعمدة المعرف أو الحالة.');

        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][entryIdColIndex] === entryId) {
                sheet.getRange(i + 1, statusColIndex + 1).setValue('Approved');
                Logger.log(`Entry [${entryId}] in sheet [${sheetName}] approved by ${user.name}.`);
                return { success: true, message: `تم اعتماد ${entryType === 'income' ? 'الدخل' : 'التكلفة'}.` };
            }
        }
        throw new Error('لم يتم العثور على المعاملة.');
    } catch (e) {
        Logger.log(`Error in approveEntry: ${e.message}`);
        throw new Error(`فشل الاعتماد: ${e.message}`);
    }
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
