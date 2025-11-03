/********************************************************************************
 * Locus Finance - Data & Reports Module (V3.1 - Timestamp Filter)
 * الوصف: يحتوي على دوال جلب البيانات الرئيسية، بيانات لوحات التحكم،
 * والتقارير المختلفة.
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
 * يجلب البيانات المجمعة للوحة التحكم بناءً على دور المستخدم.
 * @param {string} token The user's session token.
 * @returns {object} Data for the dashboard.
 */
function getDashboardData(token) {
  const user = getSession(token);
  const response = { role: user.role, kpis: null, branchStats: null, detailedReports: null, pendingEntries: null };

  if (['Owner', 'Admin'].includes(user.role)) {
    response.kpis = getGlobalStats(token);
    response.detailedReports = getDetailedReports(token);
  }
  if (user.role === 'Finance Manager') {
    response.pendingEntries = getPendingEntries(token);
  }
  return response;
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

    // *** التعديل: تم التغيير من entry.Date إلى entry.Timestamp ***
    const filteredIncome = incomeObjects.filter(entry => entry.Timestamp && new Date(entry.Timestamp) >= startDate && new Date(entry.Timestamp) <= endDate);
    const filteredCost = costObjects.filter(entry => entry.Timestamp && new Date(entry.Timestamp) >= startDate && new Date(entry.Timestamp) <= endDate);

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
    
    // *** التعديل: تم التغيير من e.Date إلى e.Timestamp ***
    const filteredIncome = _sheetDataToObjects(_getSheet(ss, 'Income')).filter(e => e.Timestamp && new Date(e.Timestamp) >= startDate && new Date(e.Timestamp) <= endDate);
    const filteredCost = _sheetDataToObjects(_getSheet(ss, 'cost')).filter(e => e.Timestamp && new Date(e.Timestamp) >= startDate && new Date(e.Timestamp) <= endDate);

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
    // --- التحصين ضد البيانات الخاطئة ---
    if (!startDateStr || !endDateStr) {
        Logger.log(`getChartData failed: Received invalid date strings. Start: ${startDateStr}, End: ${endDateStr}`);
        // إرجاع بيانات فارغة لتجنب كسر الواجهة
        return { 
            timeSeries: [['التاريخ', 'الدخل', 'التكاليف']], 
            costBreakdown: [['بند التكلفة', 'المبلغ']] 
        };
    }

    try {
        const user = getSession(token);
        if (!['Owner', 'Admin'].includes(user.role)) throw new Error('غير مصرح لك بعرض هذه البيانات.');

        const startDate = new Date(startDateStr);
        startDate.setHours(0, 0, 0, 0);
        
        const endDate = new Date(endDateStr);
        endDate.setHours(23, 59, 59, 999);

        // التحقق من صحة التواريخ بعد الإنشاء
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
            throw new Error(`Invalid date format received. Start: ${startDateStr}, End: ${endDateStr}`);
        }

        const dbId = SP.getProperty('FINANCE_DB_FILE_ID');
        const ss = _getSSById(dbId);
        const incomeObjects = _sheetDataToObjects(_getSheet(ss, 'Income'));
        const costObjects = _sheetDataToObjects(_getSheet(ss, 'cost'));

        const timeSeriesData = {};
        const costBreakdownData = {};

        const processEntry = (entry, type) => {
            if (!entry.Timestamp || !entry.Amount) return;

            const entryTimestamp = new Date(entry.Timestamp);
            if (isNaN(entryTimestamp.getTime())) return; // تجاهل التواريخ غير الصالحة في البيانات

            if (entryTimestamp >= startDate && entryTimestamp <= endDate) {
                const dateKey = Utilities.formatDate(entryTimestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                
                if (!timeSeriesData[dateKey]) {
                    timeSeriesData[dateKey] = { income: 0, cost: 0 };
                }

                if (type === 'income') {
                    timeSeriesData[dateKey].income += Number(entry.Amount);
                } else { // cost
                    timeSeriesData[dateKey].cost += Number(entry.Amount);
                    const costName = String(entry.Cost_Name || 'غير محدد').trim();
                    costBreakdownData[costName] = (costBreakdownData[costName] || 0) + Number(entry.Amount);
                }
            }
        };

        incomeObjects.forEach(entry => processEntry(entry, 'income'));
        costObjects.forEach(entry => processEntry(entry, 'cost'));

        const timeSeriesArray = [['التاريخ', 'الدخل', 'التكاليف'], ...Object.keys(timeSeriesData).sort().map(date => [date, timeSeriesData[date].income, timeSeriesData[date].cost])];
        const costBreakdownArray = [['بند التكلفة', 'المبلغ'], ...Object.keys(costBreakdownData).map(item => [item, costBreakdownData[item]])];

        Logger.log("SUCCESS: Generated Time Series Array for Chart: " + JSON.stringify(timeSeriesArray));

        return { timeSeries: timeSeriesArray, costBreakdown: costBreakdownArray };

    } catch (e) {
        Logger.log(`CRITICAL ERROR in getChartData: ${e.message} \nStack: ${e.stack}`);
        // في حالة حدوث أي خطأ آخر، أرجع رسالة خطأ واضحة
        throw new Error(`فشل تحميل بيانات الرسم البياني: ${e.message}`);
    }
}/**
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

    // *** التعديل: تم التغيير من entry.Date إلى entry.Timestamp ***
    pending.income = incomeObjects
      .filter(entry => entry.Status === 'Pending' && entry.Timestamp && new Date(entry.Timestamp) >= startDate && new Date(entry.Timestamp) <= endDate)
      .map(entry => ({
        id: entry.Entry_ID, branch: entry.Branch, date: Utilities.formatDate(new Date(entry.Date), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        service: entry.Service_Name, amount: entry.Amount, employee: entry.EmployeeName
      }));

    // *** التعديل: تم التغيير من entry.Date إلى entry.Timestamp ***
    pending.cost = costObjects
      .filter(entry => entry.Status === 'Pending' && entry.Timestamp && new Date(entry.Timestamp) >= startDate && new Date(entry.Timestamp) <= endDate)
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
