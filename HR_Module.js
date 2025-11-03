/********************************************************************************
 * Locus Finance - HR & Attendance Module
 * التاريخ: 2024
 * الوصف: يحتوي على كل الدوال المتعلقة بإدارة الموظفين، الرواتب، الحضور والانصراف،
 * والتقارير الخاصة بالموارد البشرية.
 ********************************************************************************/

// ==================================================================
// SECTION 1: EMPLOYEE & PAYROLL MANAGEMENT
// ==================================================================

/**
 * [HR] يجلب قائمة الموظفين مع بياناتهم المالية الأساسية.
 * @param {string} token The user's session token.
 * @returns {Array<Object>} A list of employee objects.
 */
function getEmployeeListForHR(token) {
  const user = getSession(token);
  if (user.role !== 'HR Manager' && user.role !== 'Admin') throw new Error('Unauthorized.');

  const empSS = _getSSById(SP.getProperty('EMPLOYEES_FILE_ID'));
  const employeesData = _sheetDataToObjects(_getSheet(empSS, 'Employees'));

  const hrSS = _getSSById(SP.getProperty('HR_FILE_ID'));
  const payrollData = _sheetDataToObjects(_getSheet(hrSS, 'Payroll_Base'));
  const profilesData = _sheetDataToObjects(_getSheet(hrSS, 'Employee_Profiles'));
  const positionsData = _sheetDataToObjects(_getSheet(hrSS, 'Employee_Positions'));

  const payrollMap = payrollData.reduce((map, item) => { map[String(item.EmployeeCode).trim()] = item; return map; }, {});
  const profileMap = profilesData.reduce((map, item) => { map[String(item.EmployeeCode).trim()] = item; return map; }, {});
  
  const positionsMap = positionsData.reduce((map, item) => {
    const code = String(item.EmployeeCode).trim();
    if (!map[code]) map[code] = { Basic_Salary: 0, Allowance: 0 };
    if (String(item.IsActive).toLowerCase() === 'yes') {
        if (item.SalaryType === 'Basic') map[code].Basic_Salary += Number(item.Amount) || 0;
        if (item.SalaryType === 'Allowance' || item.SalaryType === 'Fixed_Bonus') map[code].Allowance += Number(item.Amount) || 0;
    }
    return map;
  }, {});

  const combinedEmployees = employeesData.map(emp => {
    const empCode = String(emp.Code).trim();
    const payrollInfo = payrollMap[empCode] || {};
    const profileInfo = profileMap[empCode] || {};
    const positionInfo = positionsMap[empCode] || { Basic_Salary: 0, Allowance: 0 };

    return {
      Code: emp.Code, Name: emp.Name, Branch: emp.Branch, Role: emp.Role, IsActive: emp.IsActive,
      Basic_Salary: positionInfo.Basic_Salary,
      Allowance: positionInfo.Allowance,
      Commission_Pct: payrollInfo.Commission_Pct || 0,
      Phone: profileInfo.Phone || '',
      National_ID: profileInfo.National_ID || '',
      Qualification: profileInfo.Qualification || ''
    };
  });

  return combinedEmployees.filter(emp => emp.Role !== 'Owner');
}

/**
 * [HR] يضيف موظفاً جديداً إلى النظام.
 * @param {string} token The user's session token.
 * @param {object} newEmployeeData The new employee's data.
 * @returns {object} Success message and new employee credentials.
 */
function addEmployeeHR(token, newEmployeeData) {
  const user = getSession(token);
  if (user.role !== 'HR Manager' && user.role !== 'Admin') throw new Error('Unauthorized.');
  if (newEmployeeData.Role === 'Owner') throw new Error('Cannot assign "Owner" role.');

  const employeeCode = _generateNextEmployeeCode();
  const initialPassword = '000000';

  const empSS = _getSSById(SP.getProperty('EMPLOYEES_FILE_ID'));
  _appendObjectAsRow(_getSheet(empSS, 'Employees'), {
    Code: employeeCode, Password: initialPassword, Name: newEmployeeData.Name,
    Branch: newEmployeeData.Branch, Role: newEmployeeData.Role, IsActive: 'No'
  });

  const hrSS = _getSSById(SP.getProperty('HR_FILE_ID'));
  _appendObjectAsRow(_getSheet(hrSS, 'Payroll_Base'), {
    EmployeeCode: employeeCode, EmployeeName: newEmployeeData.Name,
    Commission_Pct: newEmployeeData.Commission_Pct || 0
  });
  
  if (newEmployeeData.Basic_Salary !== undefined) {
      addEmployeePosition(token, {
          EmployeeCode: employeeCode,
          PositionTitle: 'راتب أساسي',
          SalaryType: 'Basic',
          Amount: newEmployeeData.Basic_Salary,
          Notes: 'Initial salary on creation'
      });
  }

  _appendObjectAsRow(_getSheet(hrSS, 'Employee_Profiles'), {
    EmployeeCode: employeeCode, EmployeeName: newEmployeeData.Name,
    Phone: newEmployeeData.Phone || '', National_ID: newEmployeeData.National_ID || '',
    Qualification: newEmployeeData.Qualification || '', Onboarding_Status: 'Pending', Timestamp: new Date()
  });

  return { success: true, message: `Employee added successfully.`, newEmployee: { code: employeeCode, password: initialPassword } };
}

/**
 * [HR] يحدث بيانات موظف موجود.
 * @param {string} token The user's session token.
 * @param {string} employeeCode The code of the employee to update.
 * @param {object} updates The data to update.
 * @returns {object} A success message.
 */
function updateEmployeeHR(token, employeeCode, updates) {
  const user = getSession(token);
  if (user.role !== 'HR Manager' && user.role !== 'Admin') throw new Error('Unauthorized.');
  if (updates.Role === 'Owner') throw new Error('Cannot assign "Owner" role.');

  let updatedCount = 0;
  const findAndUpdateRow = (sheet, identifierCol, identifierValue, updateFields) => {
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idColIndex = headers.indexOf(identifierCol);
    if (idColIndex === -1) return;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][idColIndex]).trim() === String(identifierValue).trim()) {
        Object.keys(updateFields).forEach(key => {
          const colIndex = headers.indexOf(key);
          if (colIndex !== -1) {
            sheet.getRange(i + 2, colIndex + 1).setValue(updateFields[key]);
            updatedCount++;
          }
        });
      }
    }
  };

  const employeeCoreUpdates = {};
  if (updates.Name !== undefined) employeeCoreUpdates.Name = updates.Name;
  if (updates.Branch !== undefined) employeeCoreUpdates.Branch = updates.Branch;
  if (updates.Role !== undefined) employeeCoreUpdates.Role = updates.Role;
  if (updates.IsActive !== undefined) employeeCoreUpdates.IsActive = updates.IsActive;
  if (Object.keys(employeeCoreUpdates).length > 0) {
    findAndUpdateRow(_getSheet(_getSSById(SP.getProperty('EMPLOYEES_FILE_ID')), 'Employees'), 'Code', employeeCode, employeeCoreUpdates);
  }

  const hrSS = _getSSById(SP.getProperty('HR_FILE_ID'));
  const payrollUpdates = {};
  if (updates.Name !== undefined) payrollUpdates.EmployeeName = updates.Name;
  if (updates.Commission_Pct !== undefined) payrollUpdates.Commission_Pct = updates.Commission_Pct;
  if (Object.keys(payrollUpdates).length > 0) {
    findAndUpdateRow(_getSheet(hrSS, 'Payroll_Base'), 'EmployeeCode', employeeCode, payrollUpdates);
  }

  if (updates.Name !== undefined) {
    findAndUpdateRow(_getSheet(hrSS, 'Employee_Positions'), 'EmployeeCode', employeeCode, { EmployeeName: updates.Name });
  }

  const profileUpdates = {};
  if (updates.Name !== undefined) profileUpdates.EmployeeName = updates.Name;
  if (updates.Phone !== undefined) profileUpdates.Phone = updates.Phone;
  if (updates.National_ID !== undefined) profileUpdates.National_ID = updates.National_ID;
  if (updates.Qualification !== undefined) profileUpdates.Qualification = updates.Qualification;
  if (Object.keys(profileUpdates).length > 0) {
    findAndUpdateRow(_getSheet(hrSS, 'Employee_Profiles'), 'EmployeeCode', employeeCode, profileUpdates);
  }

  if (updatedCount === 0) throw new Error('Employee not found or no updates applied.');
  return { success: true, message: 'Employee data updated.' };
}

/**
 * [HR] يضيف بند راتب جديد للموظف.
 * @param {string} token The user's session token.
 * @param {object} positionData The position data.
 * @returns {object} A success message.
 */
function addEmployeePosition(token, positionData) {
  const user = getSession(token);
  if (user.role !== 'HR Manager' && user.role !== 'Admin') throw new Error('Unauthorized access.');
  if (!positionData.EmployeeCode || !positionData.PositionTitle || !positionData.SalaryType || positionData.Amount === undefined) throw new Error('Missing required position data.');

  const hrSS = _getSSById(SP.getProperty('HR_FILE_ID'));
  const positionsSheet = _getSheet(hrSS, 'Employee_Positions');
  const allPositions = _sheetDataToObjects(positionsSheet);
  const employeePositions = allPositions.filter(p => String(p.EmployeeCode || '').trim() === String(positionData.EmployeeCode).trim());
  
  const lastSequence = employeePositions.reduce((max, pos) => {
    const parts = String(pos.Position_ID || '').split('-');
    const num = parts.length > 1 ? parseInt(parts[1], 10) : 0;
    return (!isNaN(num) && num > max) ? num : max;
  }, 0);
  
  const employeeNumber = String(positionData.EmployeeCode).replace(/\D/g, ''); 
  const newPositionId = `P${employeeNumber}-${lastSequence + 1}`;

  const newPosition = {
    Position_ID: newPositionId, EmployeeCode: positionData.EmployeeCode,
    EmployeeName: _getEmployeeNameByCode(positionData.EmployeeCode), PositionTitle: positionData.PositionTitle,
    SalaryType: positionData.SalaryType, Amount: positionData.Amount,
    IsActive: 'Yes', Notes: positionData.Notes || ''
  };

  _appendObjectAsRow(positionsSheet, newPosition);
  return { success: true, message: 'Salary component added successfully.' };
}

/**
 * [HR] يحدث بند راتب موجود.
 * @param {string} token The user's session token.
 * @param {string} positionId The ID of the position to update.
 * @param {object} updates The data to update.
 * @returns {object} A success message.
 */
function updateEmployeePosition(token, positionId, updates) {
  const user = getSession(token);
  if (user.role !== 'HR Manager' && user.role !== 'Admin') throw new Error('Unauthorized access.');

  const sheet = _getSheet(_getSSById(SP.getProperty('HR_FILE_ID')), 'Employee_Positions');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idColIndex = headers.indexOf('Position_ID');
  if (idColIndex === -1) throw new Error('Position_ID column not found.');

  for (let i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === positionId) {
      Object.keys(updates).forEach(key => {
        const colIndex = headers.indexOf(key);
        if (colIndex !== -1) {
          sheet.getRange(i + 2, colIndex + 1).setValue(updates[key]);
        }
      });
      return { success: true, message: 'Salary component updated successfully.' };
    }
  }
  throw new Error('Position ID not found.');
}

/**
 * [HR] يضيف خصماً على موظف.
 * @param {string} token The user's session token.
 * @param {object} deductionData The deduction data.
 * @returns {object} A success message.
 */
function addDeduction(token, deductionData) {
  if (token) {
    const user = getSession(token);
    if (user.role !== 'HR Manager' && user.role !== 'Admin') throw new Error('Unauthorized.');
  }
  const deductionsSheet = _getSheet(_getSSById(SP.getProperty('HR_FILE_ID')), 'Deductions');
  const newDeduction = {
    Date: new Date(deductionData.Date), EmployeeCode: deductionData.EmployeeCode,
    EmployeeName: _getEmployeeNameByCode(deductionData.EmployeeCode),
    Amount: Number(deductionData.Amount) || 0, Reason: deductionData.Reason || '',
    Notes: deductionData.Notes || '', Timestamp: new Date()
  };
  _appendObjectAsRow(deductionsSheet, newDeduction);
  return { success: true, message: 'Deduction added.' };
}

// ==================================================================
// SECTION 2: ATTENDANCE
// ==================================================================

/**
 * [HR] يسجل عملية حضور أو انصراف للموظف.
 * @param {string} token The user's session token.
 * @param {string} type The attendance type ('Check-in' or 'Check-out').
 * @param {string} notes Any notes.
 * @param {number} clientLat Latitude from client.
 * @param {number} clientLon Longitude from client.
 * @param {string} clientLocationStatus Status of geolocation from client.
 * @returns {object} A success message and status.
 */
function logAttendance(token, type, notes, clientLat, clientLon, clientLocationStatus) {
  try {
    const user = getSession(token);
    if (!user || !user.code) throw new Error('Invalid session.');

    const now = new Date();
    let locationStatus = 'Unverified - Client Error';
    let isLocationVerified = false;

    if (clientLocationStatus === 'Granted' && clientLat && clientLon) {
      const allBranches = _sheetDataToObjects(_getSheet(_getSSById(SP.getProperty('MASTER_SETTING_FILE_ID')), 'Branch List'));
      let closestBranchDistance = Infinity;

      for (const branch of allBranches) {
        if (branch.Latitude && branch.Longitude && branch.Location_Radius_Meters) {
          const distance = _getDistanceBetweenPoints(clientLat, clientLon, Number(branch.Latitude), Number(branch.Longitude));
          if (distance < closestBranchDistance) closestBranchDistance = distance;
          if (distance <= Number(branch.Location_Radius_Meters)) {
            isLocationVerified = true;
            break; 
          }
        }
      }
      locationStatus = isLocationVerified ? 'Verified' : `Mismatch - ${Math.round(closestBranchDistance)}m away`;
    } else {
      locationStatus = `Unverified - ${clientLocationStatus}`;
    }

    const logSheet = _getSheet(_getSSById(SP.getProperty('ATTENDANCE_FILE_ID')), 'Attendance_Log');
    _appendObjectAsRow(logSheet, {
      DateTime: now, EmployeeCode: user.code, EmployeeName: user.name, Type: type, 
      Source: 'WebApp', Notes: notes || '', Latitude: clientLat || '', Longitude: clientLon || '',
      Location_Status: locationStatus, Timestamp: new Date()
    });

    const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'hh:mm:ss a');
    let message = `Successfully logged ${type} at ${formattedTime}. (${locationStatus})`;

    return { success: true, message: message, locationStatus: locationStatus };

  } catch (e) {
    Logger.log(`!!! [ERROR] in logAttendance: ${e.message} at ${e.stack}`);
    throw new Error(`Failed to log attendance: ${e.message}`);
  }
}

/**
 * [HR] يجلب سجل الحضور الخاص بالموظف الحالي.
 * @param {string} token The user's session token.
 * @param {string} startDateStr Start date.
 * @param {string} endDateStr End date.
 * @returns {Array<Object>} A list of attendance log objects.
 */
function getMyAttendanceLog(token, startDateStr, endDateStr) {
  const user = getSession(token);
  if (!user || !user.code) throw new Error('Invalid session.');

  const allLogs = _sheetDataToObjects(_getSheet(_getSSById(SP.getProperty('ATTENDANCE_FILE_ID')), 'Attendance_Log'));
  const startDate = new Date(startDateStr); startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(endDateStr); endDate.setHours(23, 59, 59, 999);

  return allLogs.filter(log => {
    const logDate = new Date(log.DateTime);
    return String(log.EmployeeCode).trim() === String(user.code).trim() && logDate >= startDate && logDate <= endDate;
  }).map(log => {
    log.FormattedDateTime = Utilities.formatDate(new Date(log.DateTime), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a');
    return log;
  }).sort((a, b) => new Date(b.DateTime) - new Date(a.DateTime));
}

/**
 * [HR] يجلب سجل الحضور لجميع الموظفين (للمدير فقط).
 * @param {string} token The user's session token.
 * @param {string} startDateStr Start date.
 * @param {string} endDateStr End date.
 * @returns {Array<Object>} A list of all attendance logs.
 */
function getAllAttendanceLogs(token, startDateStr, endDateStr) {
  const user = getSession(token);
  if (user.role !== 'HR Manager' && user.role !== 'Admin') throw new Error('Unauthorized.');

  const allLogs = _sheetDataToObjects(_getSheet(_getSSById(SP.getProperty('ATTENDANCE_FILE_ID')), 'Attendance_Log'));
  const startDate = new Date(startDateStr); startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(endDateStr); endDate.setHours(23, 59, 59, 999);

  return allLogs.filter(log => {
    const logDate = new Date(log.DateTime);
    return logDate >= startDate && logDate <= endDate;
  }).map(log => {
    log.EmployeeName = _getEmployeeNameByCode(String(log.EmployeeCode).trim());
    log.FormattedDateTime = Utilities.formatDate(new Date(log.DateTime), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a');
    return log;
  }).sort((a, b) => new Date(b.DateTime) - new Date(a.DateTime));
}

// ==================================================================
// SECTION 3: SETUP & BOOTSTRAP FUNCTIONS (RUNNABLE FROM EDITOR)
// ==================================================================

/**
 * [SETUP] يهيئ كل ما يتعلق بالموارد البشرية (شيتات + تريجرز).
 * @returns {object} A success message.
 */
function HR_bootstrapAll(){
  ensureHRSheets();
  ensureAttendanceSheets();
  installHRTriggers();
  return { success: true, message: 'تم تهيئة HR و Attendance وتثبيت التريجرز.' };
}

/**
 * [SETUP] يقوم بتثبيت أو تحديث التريجرز الخاصة بالموارد البشرية.
 * @returns {object} A success message.
 */
function installHRTriggers(){
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (['hrRecalcPayroll', 'processDailyAttendance'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Recalculate payroll on the 1st of every month at 5 AM.
  ScriptApp.newTrigger('hrRecalcPayroll').timeBased().onMonthDay(1).atHour(5).create();
  
  // Process previous day's attendance every day at 2 AM.
  ScriptApp.newTrigger('processDailyAttendance').timeBased().atHour(2).everyDays(1).create();
  
  Logger.log('✅ تم تثبيت/تحديث تريجرز الموارد البشرية بنجاح.');
  return { success: true, message: 'تم تثبيت/تحديث تريجرز الموارد البشرية بنجاح.' };
}

/**
 * [SETUP] يتأكد من وجود جميع الشيتات اللازمة للموارد البشرية.
 * @returns {object} A success message.
 */
function ensureHRSheets(){
  const hrId = SP.getProperty('HR_FILE_ID');
  if (!hrId) throw new Error("HR_FILE_ID is not set.");
  const ss = _getSSById(hrId);

  const sheetsToEnsure = {
    'Employee_Positions': ['Position_ID', 'EmployeeCode', 'EmployeeName', 'PositionTitle', 'SalaryType', 'Amount', 'IsActive', 'Notes'],
    'Payroll_Base': ['EmployeeCode', 'EmployeeName', 'Commission_Pct'],
    'Deductions': ['Date', 'EmployeeCode', 'EmployeeName', 'Amount', 'Reason', 'Notes', 'Timestamp'],
    'Payroll_Result': ['Month', 'EmployeeCode', 'EmployeeName', 'Branch', 'Basic', 'Allowance', 'Commission', 'Overtime', 'Deductions', 'NetPay', 'Timestamp'],
    'Employee_Profiles': ['EmployeeCode', 'EmployeeName', 'Phone', 'National_ID', 'Qualification', 'Onboarding_Status', 'National_ID_Front_URL', 'National_ID_Back_URL', 'Birth_Certificate_URL', 'Qualification_Certificate_URL', 'Other_Documents_URL', 'Timestamp'],
  };

  Object.entries(sheetsToEnsure).forEach(([sheetName, headers]) => {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setFrozenRows(1);
    }
  });
  return { success: true, message: 'HR sheets checked.' };
}

/**
 * [SETUP] يتأكد من وجود جميع الشيتات اللازمة للحضور والانصراف.
 * @returns {object} A success message.
 */
function ensureAttendanceSheets(){
  const attId = SP.getProperty('ATTENDANCE_FILE_ID');
  if (!attId) throw new Error("ATTENDANCE_FILE_ID is not set.");
  const ss = _getSSById(attId);

  const sheetsToEnsure = {
    'Attendance_Log': ['DateTime', 'EmployeeCode', 'EmployeeName', 'Type', 'Source', 'Notes', 'Latitude', 'Longitude', 'Location_Status', 'Timestamp'],
    'Attendance_Daily': ['Date', 'EmployeeCode', 'EmployeeName', 'Branch', 'ShiftName', 'Expected_In', 'Expected_Out', 'FirstIn', 'LastOut', 'WorkHours', 'Lateness_Minutes', 'EarlyDeparture_Minutes', 'Status', 'Verified', 'Deduction_Reason', 'Deduction_Amount', 'Timestamp'],
    'Official_Holidays': ['Date', 'Description', 'IsPaid', 'Timestamp']
  };

  Object.entries(sheetsToEnsure).forEach(([sheetName, headers]) => {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setFrozenRows(1);
    }
  });
  return { success: true, message: 'Attendance sheets checked.' };
}

// ==================================================================
// SECTION 4: PRIVATE HELPER FUNCTIONS FOR HR MODULE
// ==================================================================

/**
 * يولد كود موظف جديد بشكل تسلسلي.
 * @private
 * @returns {string} The next employee code (e.g., 'emp001').
 */
function _generateNextEmployeeCode() {
  const employeesSheet = _getSheet(_getSSById(SP.getProperty('EMPLOYEES_FILE_ID')), 'Employees');
  const lastRow = employeesSheet.getLastRow();
  if (lastRow < 2) return 'emp001';
  
  const allCodes = employeesSheet.getRange("A2:A" + lastRow).getValues().flat();
  let maxNumber = 0;
  allCodes.forEach(code => {
    if (code && String(code).toLowerCase().startsWith('emp')) {
      const numberPart = parseInt(String(code).substring(3), 10);
      if (!isNaN(numberPart) && numberPart > maxNumber) maxNumber = numberPart;
    }
  });
  const newNumber = maxNumber + 1;
  return `emp${String(newNumber).padStart(3, '0')}`;
}

/**
 * يحسب المسافة بين نقطتين جغرافيتين.
 * @private
 */
function _getDistanceBetweenPoints(lat1, lon1, lat2, lon2) {
  const R = 6371e3; // Radius of Earth in meters
  const φ1 = lat1 * Math.PI/180, φ2 = lat2 * Math.PI/180;
  const Δφ = (lat2-lat1) * Math.PI/180, Δλ = (lon2-lon1) * Math.PI/180;
  const a = Math.sin(Δφ/2) * Math.sin(Δφ/2) + Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ/2) * Math.sin(Δλ/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}
