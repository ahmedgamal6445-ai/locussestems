/********************************************************************************
 * Locus Finance - Utilities Module
 * الوصف: يحتوي على دوال مساعدة عامة تستخدم في جميع أنحاء المشروع.
 ********************************************************************************/

/**
 * يحصل على نوع MIME من امتداد الملف.
 * @param {string} extension The file extension.
 * @returns {string} The corresponding MimeType.
 */
function _getMimeType(extension) {
  const ext = extension.toLowerCase();
  const mimeTypes = {
    'jpg': MimeType.JPEG, 'jpeg': MimeType.JPEG, 'png': MimeType.PNG, 'gif': MimeType.GIF,
    'pdf': MimeType.PDF, 'doc': MimeType.MICROSOFT_WORD, 'docx': MimeType.MICROSOFT_WORD,
    'xls': MimeType.MICROSOFT_EXCEL, 'xlsx': MimeType.MICROSOFT_EXCEL,
    'ppt': MimeType.MICROSOFT_POWERPOINT, 'pptx': MimeType.MICROSOFT_POWERPOINT
  };
  return mimeTypes[ext] || MimeType.OCTET_STREAM;
}

/**
 * يفتح جدول بيانات بواسطة ID.
 * @param {string} id The Spreadsheet ID.
 * @returns {Spreadsheet} The Spreadsheet object.
 */
function _getSSById(id) { if (!id) throw new Error("Sheet ID is undefined."); return SS_APP.openById(id); }

/**
 * يحصل على شيت معين بالاسم من جدول بيانات.
 * @param {Spreadsheet} ss The Spreadsheet object.
 * @param {string} name The name of the sheet.
 * @returns {Sheet} The Sheet object.
 */
function _getSheet(ss, name) { const sh = ss.getSheetByName(name); if (!sh) throw new Error(`Sheet not found: ${name}`); return sh; }

/**
 * يتأكد من وجود مجلد فرعي، وينشئه إذا لم يكن موجوداً.
 * @param {Folder} parentFolder The parent folder.
 * @param {string} name The name of the subfolder.
 * @returns {Folder} The subfolder object.
 */
function _ensureSubFolder(parentFolder, name) { const it = parentFolder.getFoldersByName(name); return it.hasNext() ? it.next() : parentFolder.createFolder(name); }

/**
 * تحول بيانات الشيت إلى مصفوفة من الكائنات بناءً على أسماء الأعمدة.
 * @param {Sheet} sheet - الشيت المراد قراءة بياناته.
 * @returns {Array<Object>} مصفوفة من الكائنات، كل كائن يمثل صفًا.
 */
function _sheetDataToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data.shift().map(h => String(h).trim());

  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      if (header) {
        obj[header] = row[i];
      }
    });
    return obj;
  });
}

/**
 * تضيف صفًا جديدًا إلى الشيت باستخدام كائن.
 * @param {Sheet} sheet - الشيت المراد الكتابة فيه.
 * @param {Object} obj - الكائن الذي يحتوي على البيانات.
 */
function _appendObjectAsRow(sheet, obj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => {
    if (obj[header] instanceof Date) return obj[header];
    return obj[header] !== undefined && obj[header] !== null ? obj[header] : '';
  });
  sheet.appendRow(row);
}

/**
 * يولد معرفًا تسلسليًا فريدًا.
 * @param {string} prefix - بادئة نوع الإدخال.
 * @param {string} branch - اسم الفرع.
 * @param {Date} date - تاريخ الإدخال.
 * @param {Sheet} sheet - الشيت للبحث فيه.
 * @returns {string} المعرف الجديد.
 */
function _generateId(prefix, branch, date, sheet) {
  const branchCode = (branch || 'XXX').substring(0, 3).toUpperCase();
  const datePart = Utilities.formatDate(date, Session.getScriptTimeZone(), 'ddMMyy');
  const idPrefix = `${branchCode}-${datePart}-`;

  const allIDs = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat();
  let lastNumber = 0;

  allIDs.forEach(id => {
    if (id && String(id).startsWith(idPrefix)) {
      const numberPart = parseInt(String(id).substring(idPrefix.length), 10);
      if (!isNaN(numberPart) && numberPart > lastNumber) {
        lastNumber = numberPart;
      }
    }
  });

  const newNumber = lastNumber + 1;
  return `${idPrefix}${String(newNumber).padStart(3, '0')}`;
}

/**
 * يتأكد من وجود وهيكلة شيتات قاعدة البيانات المالية.
 * @returns {string} The ID of the finance database spreadsheet.
 */
function _ensureFinanceDB() {
  const dbId = SP.getProperty('FINANCE_DB_FILE_ID');
  if (!dbId) throw new Error("FINANCE_DB_FILE_ID is not set.");
  
  const ss = _getSSById(dbId);
  
  const INCOME_HEADERS = ['Entry_ID', 'Branch', 'Date', 'Day', 'Service_Name', 'Amount', 'Payment_Method', 'Doctor_Transfer_No', 'Service_Details', 'Notes', 'EmployeeCode', 'EmployeeName', 'Status', 'Timestamp'];
  const COST_HEADERS = ['Entry_ID', 'Branch', 'Date', 'Day', 'Cost_Name', 'Amount', 'Payment_Method', 'Cost_Details', 'Notes', 'EmployeeCode', 'EmployeeName', 'Attachment_URL', 'Status', 'Timestamp'];
  const KPI_SNAPSHOTS_HEADERS = ['Date','Revenue','Cost','Net','Timestamp'];

  const sheetsToEnsure = {
    'Income': INCOME_HEADERS,
    'cost': COST_HEADERS,
    'KPI_Snapshots': KPI_SNAPSHOTS_HEADERS
  };

  Object.entries(sheetsToEnsure).forEach(([sheetName, headers]) => {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setFrozenRows(1);
    }
  });
  return ss.getId();
}

const employeeNameCache = {};
/**
 * يجلب اسم الموظف بواسطة الكود الخاص به مع استخدام الكاش.
 * @param {string} employeeCode The employee's code.
 * @returns {string} The employee's name.
 */
function _getEmployeeNameByCode(employeeCode) {
    if (employeeNameCache[employeeCode]) {
        return employeeNameCache[employeeCode];
    }

    const empSS = _getSSById(SP.getProperty('EMPLOYEES_FILE_ID'));
    const sh = empSS.getSheets()[0];
    const users = _sheetDataToObjects(sh);

    users.forEach(user => {
        employeeNameCache[String(user.Code).trim()] = String(user.Name).trim();
    });

    return employeeNameCache[employeeCode] || 'Unknown Employee';
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
