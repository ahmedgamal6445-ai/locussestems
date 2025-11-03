/********************************************************************************
 * Locus Finance - Authentication & Session Module
 * الوصف: يحتوي على دوال تسجيل الدخول، إدارة الجلسات، تغيير كلمة المرور،
 * والتواصل الآمن بين التطبيقات.
 ********************************************************************************/

/**
 * يتحقق من بيانات المستخدم ويقوم بإنشاء جلسة له.
 * @param {string} code The user's code.
 * @param {string} password The user's password.
 * @returns {object} An object containing the session token and user data.
 */
function login(code, password) {
  const empSS = _getSSById(SP.getProperty('EMPLOYEES_FILE_ID'));
  const sh = empSS.getSheets()[0];
  const users = _sheetDataToObjects(sh);

  const activeUser = users.find(user =>
    String(user.Code).trim() === String(code).trim() &&
    String(user.Password).trim() === String(password).trim() &&
    String(user.IsActive).toLowerCase() === 'yes'
  );

  if (!activeUser) throw new Error('بيانات الدخول غير صحيحة أو المستخدم غير مُفعّل.');

  const userSession = {
    code: String(activeUser.Code),
    name: String(activeUser.Name),
    branch: String(activeUser.Branch),
    role: String(activeUser.Role),
  };

  const token = Utilities.getUuid();
  CACHE.put(token, JSON.stringify(userSession), 3600); // Session lasts for 1 hour
  return { token, user: userSession };
}

/**
 * يسترجع بيانات جلسة المستخدم باستخدام التوكن.
 * @param {string} token The session token.
 * @returns {object} The user session data.
 */
function getSession(token) {
  if (!token) throw new Error('جلسة غير صالحة. يرجى تسجيل الدخول مرة أخرى.');
  const sessionData = CACHE.get(token);
  if (!sessionData) throw new Error('انتهت صلاحية الجلسة. يرجى تسجيل الدخول مرة أخرى.');
  return JSON.parse(sessionData);
}

/**
 * يسمح للمستخدم الحالي بتغيير كلمة المرور الخاصة به.
 * @param {string} token - توكن جلسة المستخدم.
 * @param {string} oldPassword - كلمة المرور الحالية للتحقق.
 * @param {string} newPassword - كلمة المرور الجديدة.
 * @returns {Object} رسالة نجاح.
 */
function changePassword(token, oldPassword, newPassword) {
  try {
    const user = getSession(token);
    if (!user) {
      throw new Error('جلسة غير صالحة. يرجى تسجيل الدخول مرة أخرى.');
    }

    if (!oldPassword || !newPassword) {
      throw new Error('يجب إدخال كلمة المرور القديمة والجديدة.');
    }
    
    if (newPassword.length < 6) {
      throw new Error('يجب أن لا تقل كلمة المرور الجديدة عن 6 أحرف.');
    }

    const empSS = _getSSById(SP.getProperty('EMPLOYEES_FILE_ID'));
    const sheet = _getSheet(empSS, 'Employees'); 

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); 
    const codeColIndex = headers.indexOf('Code');
    const passwordColIndex = headers.indexOf('Password');

    if (codeColIndex === -1 || passwordColIndex === -1) {
      throw new Error('خطأ في النظام: لم يتم العثور على أعمدة البيانات الأساسية.');
    }

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][codeColIndex]).trim() === String(user.code).trim()) {
        const storedPassword = String(data[i][passwordColIndex]).trim();

        if (storedPassword !== String(oldPassword).trim()) {
          throw new Error('كلمة المرور القديمة غير صحيحة.');
        }

        sheet.getRange(i + 2, passwordColIndex + 1).setValue(newPassword);
        
        Logger.log(`Password changed successfully for user ${user.code}.`);
        return { success: true, message: 'تم تغيير كلمة المرور بنجاح.' };
      }
    }

    throw new Error('لم يتم العثور على حسابك في النظام.');

  } catch (e) {
    Logger.log(`Error in changePassword: ${e.message}`);
    throw new Error(`فشل تغيير كلمة المرور: ${e.message}`);
  }
}

/**
 * ينشئ توكن قصير الأمد للتواصل مع تطبيق آخر (مثل تطبيق العقود).
 * @param {string} token The user's main session token.
 * @returns {object} An object containing the new handshake token.
 */
function generateHandshakeToken(token) {
  const user = getSession(token); 

  const handshakeToken = Utilities.getUuid();
  const handshakeData = {
    user: user,
    used: false 
  };

  CACHE.put(`handshake_${handshakeToken}`, JSON.stringify(handshakeData), 60);

  Logger.log(`Generated handshake token for user ${user.code}`);
  return { handshakeToken: handshakeToken };
}

/**
 * يتحقق من توكن المصافحة القادم من تطبيق آخر.
 * @param {string} handshakeToken The token sent from the other app.
 * @returns {Object|null} The user session data if the token is valid, otherwise null.
 */
function verifyHandshakeToken(handshakeToken) {
  const cacheKey = `handshake_${handshakeToken}`;
  const handshakeDataJSON = CACHE.get(cacheKey);

  if (!handshakeDataJSON) {
    Logger.log(`Handshake verification failed: Token not found or expired.`);
    return null; 
  }

  const handshakeData = JSON.parse(handshakeDataJSON);

  if (handshakeData.used) {
    Logger.log(`Handshake verification failed: Token already used.`);
    CACHE.remove(cacheKey);
    return null;
  }

  handshakeData.used = true;
  CACHE.put(cacheKey, JSON.stringify(handshakeData), 60); 

  Logger.log(`Handshake verification successful for user ${handshakeData.user.code}`);
  return handshakeData.user; 
}
