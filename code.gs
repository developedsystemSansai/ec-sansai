var DRIVE_FOLDER_NAME = 'EC_Uploads';
var CUSTOM_DOMAIN = 'https://developedsystemSansai.github.io/ec-sansai';
var _configCache = {};
function getConfig_(key) {
  if (!_configCache[key]) {
    var val = PropertiesService.getScriptProperties().getProperty(key);
    if (!val) throw new Error('[CONFIG] Script Property "' + key + '" ไม่ได้ตั้งค่า กรุณาตั้งค่าใน Project Settings → Script Properties');
    _configCache[key] = val;
  }
  return _configCache[key];
}
function getSheetId_() { return getConfig_('SHEET_ID'); }
function getDriveFolderId_() { return getConfig_('DRIVE_FOLDER_ID'); }
var SHEET_ID = '';
var DRIVE_FOLDER_ID = '';
var CSRF_TOKEN_EXPIRY = 30 * 60;
function generateCsrfToken_(sessionId) {
  try {
    var token = Utilities.getUuid();
    var cache = CacheService.getScriptCache();
    cache.put('csrf_' + sessionId, token, CSRF_TOKEN_EXPIRY);
    return token;
  } catch (e) {
    Logger.log('generateCsrfToken_ error: ' + e.message);
    return '';
  }
}
function validateCsrfToken_(sessionId, token) {
  try {
    if (!sessionId || !token) return false;
    var cache = CacheService.getScriptCache();
    var stored = cache.get('csrf_' + sessionId);
    if (!stored) return false;
    return stored === token;
  } catch (e) {
    Logger.log('validateCsrfToken_ error: ' + e.message);
    return false;
  }
}
function getCsrfToken(sessionId) {
  if (!sessionId) return { success: false };
  var token = generateCsrfToken_(sessionId);
  return token ? { success: true, token: token } : { success: false };
}
var MAX_FILE_SIZE_MB = 20;
var ALLOWED_MIME_TYPES = [
  'application/pdf',
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'image/jpeg',
  'image/png'
];
var MAGIC_BYTES = {
  'application/pdf': ['25504446'],
  'image/jpeg': ['FFD8FF'],
  'image/png': ['89504E47'],
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['504B0304']
};
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.resetToken) {
      var token = e.parameter.resetToken;
      var validation = validateResetToken(token);
      if (validation.success) {
        var tpl = HtmlService.createTemplateFromFile('reset-password');
        tpl.resetToken = token;
        return tpl.evaluate()
          .setTitle('Reset Password - EC Sansai')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      } else {
        return HtmlService.createHtmlOutput(
          '<html><head><meta charset="UTF-8"><title>ลิงก์ไม่ถูกต้อง</title>' +
          '<style>body{font-family:sans-serif;text-align:center;padding:50px;}</style></head>' +
          '<body><h2>❌ ลิงก์ไม่ถูกต้องหรือหมดอายุแล้ว</h2>' +
          '<p>กรุณาขอรีเซ็ตรหัสผ่านใหม่</p>' +
          '<a href="' + CUSTOM_DOMAIN + '" style="display:inline-block;padding:10px 20px;background:#2d8a52;color:#fff;text-decoration:none;border-radius:5px;margin-top:20px;">กลับหน้าเข้าสู่ระบบ</a>' +
          '</body></html>'
        );
      }
    }
    return HtmlService.createHtmlOutputFromFile('index')
  .setTitle('EC Online Submission - Sansai')
  .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    Logger.log('doGet error: ' + err.message + ' | stack: ' + err.stack);
    try { writeAuditLog_('SYSTEM_ERROR', 'doGet: ' + err.message, 'system', 'ERROR'); } catch(e2) { Logger.log('AuditLog unavailable: ' + e2.message); }
    return HtmlService.createHtmlOutput(
      '<html><head><meta charset="UTF-8"><title>ระบบขัดข้อง</title>' +
      '<style>body{font-family:sans-serif;text-align:center;padding:60px;background:#f5f9f6;}' +
      'h2{color:#1a6b3a;} p{color:#555;} a{color:#2d8a52;font-weight:600;}</style></head>' +
      '<body><h2>&#9888;&#65039; ระบบขัดข้องชั่วคราว</h2>' +
      '<p>กรุณาลองใหม่อีกครั้ง หากปัญหายังคงอยู่ กรุณาติดต่อผู้ดูแลระบบ</p>' +
      '<a href="' + CUSTOM_DOMAIN + '">กลับหน้าหลัก</a>' +
      '</body></html>'
    );
  }
}
function getSheet(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(getSheetId_());
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log('Created new sheet: ' + sheetName);
    }
    return sheet;
  } catch (e) {
    Logger.log('Error accessing sheet ' + sheetName + ': ' + e.message);
    return null;
  }
}
function sheetToObjects(sheetName) {
  try {
    var sheet = getSheet(sheetName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var isEmpty = row.every(function(cell) { return cell === '' || cell === null || cell === undefined; });
      if (isEmpty) continue;
      var obj = {};
      headers.forEach(function(h, j) { obj[h] = j < row.length ? row[j] : ''; });
      obj['_row'] = i + 1;
      result.push(obj);
    }
    return result;
  } catch (e) {
    Logger.log('Error in sheetToObjects(' + sheetName + '): ' + e.message);
    return [];
  }
}
function appendRow(sheetName, obj) {
  try {
    var sheet = getSheet(sheetName);
    if (!sheet) return false;
    if (sheet.getLastRow() === 0) {
      var initHeaders = Object.keys(obj);
      sheet.appendRow(initHeaders);
      sheet.getRange(1, 1, 1, initHeaders.length)
           .setBackground('#2d8a52').setFontColor('#ffffff').setFontWeight('bold');
    }
    var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = sheetHeaders.map(function(h) {
      var key = String(h || '').trim().toLowerCase();
      if (obj[key] !== undefined) return obj[key];
      if (obj[String(h || '').trim()] !== undefined) return obj[String(h || '').trim()];
      return '';
    });
    sheet.appendRow(row);
    return true;
  } catch (e) {
    Logger.log('Error in appendRow(' + sheetName + '): ' + e.message);
    return false;
  }
}
function updateRowByField(sheetName, fieldName, fieldValue, updates) {
  try {
    var sheet = getSheet(sheetName);
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return false;
    var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var colIdx = headers.indexOf(fieldName.toLowerCase());
    if (colIdx === -1) return false;
    var targetValue = String(fieldValue || '').trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colIdx] || '').trim() === targetValue) {
        Object.keys(updates).forEach(function(key) {
          var updateCol = headers.indexOf(key.toLowerCase());
          if (updateCol !== -1) sheet.getRange(i + 1, updateCol + 1).setValue(updates[key]);
        });
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log('Error in updateRowByField: ' + e.message);
    return false;
  }
}
function updateRowByField2(sheetName, field1, val1, field2, val2, updates) {
  try {
    var sheet = getSheet(sheetName);
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return false;
    var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var c1 = headers.indexOf(field1.toLowerCase());
    var c2 = headers.indexOf(field2.toLowerCase());
    if (c1 === -1 || c2 === -1) return false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][c1] || '').trim() === String(val1 || '').trim() &&
          String(data[i][c2] || '').trim().toLowerCase() === String(val2 || '').trim().toLowerCase()) {
        Object.keys(updates).forEach(function(key) {
          var col = headers.indexOf(key.toLowerCase());
          if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(updates[key]);
        });
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log('updateRowByField2 error: ' + e.message);
    return false;
  }
}
function writeLog(action, detail, user) {
  try {
    var sheet = getSheet('Logs');
    if (!sheet) return;
    if (sheet.getLastRow() === 0) sheet.appendRow(['timestamp', 'username', 'action', 'detail']);
    sheet.appendRow([new Date(), user || 'system', action, detail]);
  } catch (e) { Logger.log('Log write failed: ' + e.message); }
}
function validateEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').toLowerCase());
}
function sanitizeEmailField_(text) {
  return String(text || '')
    .replace(/[\r\n]/g, ' ')
    .replace(/[<>]/g, '')
    .trim()
    .substring(0, 500);
}
function sendEmail_(to, subject, body) {
  if (!to || !validateEmail_(to)) return false;
  subject = sanitizeEmailField_(subject);
  body = String(body || '').substring(0, 10000);
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body + '\n\n─────────────────────\nระบบ EC Online Submission\nโรงพยาบาลสันทราย',
      noReply: true
    });
    return true;
  } catch (e) {
    Logger.log('Email send failed: ' + e.message);
    return false;
  }
}
function sendEmailWithAttachments_(to, subject, body, attachments) {
  try {
    if (!to || !validateEmail_(to)) return false;
    var options = {
      to: to,
      subject: subject,
      body: body + '\n\n─────────────────────\nระบบ EC Online Submission\nโรงพยาบาลสันทราย'
    };
    if (attachments && attachments.length > 0) {
      options.attachments = attachments;
    }
    MailApp.sendEmail(options);
    return true;
  } catch (e) {
    Logger.log('sendEmailWithAttachments_ failed: ' + e.message);
    return false;
  }
}
function notifyStaff_(subject, body) {
  try {
    var users = sheetToObjects('Users');
    users.filter(function(u) {
      return (String(u['roles'] || '').indexOf('staff') !== -1 ||
              String(u['roles'] || '').indexOf('admin') !== -1) &&
             String(u['status'] || 'active').toLowerCase() === 'active';
    }).forEach(function(s) {
      if (s['email'] && validateEmail_(s['email'])) sendEmail_(s['email'], subject, body);
    });
    return true;
  } catch (e) {
    Logger.log('Staff notification failed: ' + e.message);
    return false;
  }
}
function login(username, password) {
  try {
    if (!username || !password) return { success: false, message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
    var rateCheck = checkRateLimit_(username);
    if (rateCheck.blocked) return { success: false, message: rateCheck.message };
    ensureUsersSheet();
    var users = sheetToObjects('Users');
    var u = users.find(function(u) { return String(u['username'] || '').trim() === String(username).trim(); });
    if (!u) {
      recordFailedLogin_(username);
      return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
    }
    var storedHash = String(u['password_hash'] || '').trim();
    if (!storedHash) {
      writeLog('LOGIN_BLOCKED_NO_HASH', String(u['username'] || ''), 'system');
      recordFailedLogin_(username);
      return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
    }
    var passwordOk = verifyPassword_(password, storedHash);
    if (!passwordOk) {
      recordFailedLogin_(username);
      return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
    }
    if (isLegacyHash_(storedHash)) {
      try {
        var upgradedHash = hashPassword_(password);
        updateRowByField('Users', 'username', String(u['username'] || '').trim(), {
          password_hash: upgradedHash
        });
        writeAuditLog_('PASSWORD_HASH_UPGRADED', 'HMAC→PBKDF2 auto-upgrade on login', String(u['username'] || ''), 'INFO');
      } catch (upgradeErr) {
        Logger.log('[login] Hash upgrade failed (non-fatal): ' + upgradeErr.message);
      }
    }
    var status = String(u['status'] || 'active').toLowerCase();
    if (status === 'inactive') return { success: false, message: 'บัญชีของท่านถูกระงับการใช้งาน' };
    if (status === 'pending')  return { success: false, message: 'บัญชีของท่านรออนุมัติจากผู้ดูแลระบบ' };
    var roles = String(u['roles'] || 'researcher').split(',').map(function(r) {
      r = r.trim().replace('pending_', '');
      return r === 'user' ? 'researcher' : r;
    }).filter(Boolean);
    if (!roles.length) roles = ['researcher'];
    clearFailedLogins_(username);
    writeLog('LOGIN', 'สำเร็จ', username);
    var sessionId = '';
    try {
      sessionId = createSession_(String(u['username'] || '').trim(), '', '');
    } catch (sessionErr) {
      Logger.log('createSession_ error (non-fatal): ' + sessionErr.message);
    }
    return {
      success:    true,
      sessionId:  sessionId,
      username:   String(u['username']   || '').trim(),
      name:       String(u['name']       || username),
      roles:      roles
    };
  } catch (e) {
    Logger.log('login error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}
function clearFailedLogins_(username) {
  try {
    CacheService.getScriptCache().remove('login_' + username);
    var sheet = getSheet('FailedLogins');
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h||'').trim().toLowerCase(); });
    var uCol = headers.indexOf('username');
    var cCol = headers.indexOf('count');
    var tCol = headers.indexOf('locked_until');
    if (uCol === -1) return;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uCol]||'').trim() === String(username).trim()) {
        if (cCol !== -1) sheet.getRange(i+1, cCol+1).setValue(0);
        if (tCol !== -1) sheet.getRange(i+1, tCol+1).setValue(0);
        return;
      }
    }
  } catch (e) {
    Logger.log('clearFailedLogins_ error: ' + e.message);
  }
}
function recordFailedLogin_(username) {
  try {
    var cache = CacheService.getScriptCache();
    var key   = 'login_' + username;
    var raw   = cache.get(key);
    var data  = raw ? JSON.parse(raw) : { count: 0, lockedUntil: 0 };
    data.count++;
    if (data.count >= MAX_LOGIN_ATTEMPTS) {
      data.lockedUntil = Date.now() + (LOCKOUT_DURATION_MINUTES * 60 * 1000);
      writeLog('ACCOUNT_LOCKED', username + ' | ' + data.count + ' failed attempts', 'system');
    }
    cache.put(key, JSON.stringify(data), LOCKOUT_DURATION_MINUTES * 60 * 2);
    recordFailedLoginSheet_(username, data.count, data.lockedUntil);
  } catch (e) {
    Logger.log('recordFailedLogin_ error: ' + e.message);
  }
}
function recordFailedLoginSheet_(username, count, lockedUntil) {
  try {
    var sheet = getSheet('FailedLogins');
    if (!sheet) return;
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['username', 'count', 'locked_until', 'last_attempt']);
      sheet.getRange(1,1,1,4).setBackground('#d93025').setFontColor('#fff').setFontWeight('bold');
    }
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h||'').trim().toLowerCase(); });
    var uCol = headers.indexOf('username');
    var cCol = headers.indexOf('count');
    var tCol = headers.indexOf('locked_until');
    var lCol = headers.indexOf('last_attempt');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uCol]||'').trim() === String(username).trim()) {
        if (cCol !== -1) sheet.getRange(i+1, cCol+1).setValue(count);
        if (tCol !== -1) sheet.getRange(i+1, tCol+1).setValue(lockedUntil);
        if (lCol !== -1) sheet.getRange(i+1, lCol+1).setValue(Date.now());
        return;
      }
    }
    sheet.appendRow([username, count, lockedUntil, Date.now()]);
  } catch (e) {
    Logger.log('recordFailedLoginSheet_ error: ' + e.message);
  }
}
var MAX_LOGIN_ATTEMPTS = 5;
var LOCKOUT_DURATION_MINUTES = 15;
function checkRateLimit_(username) {
  try {
    var cache = CacheService.getScriptCache();
    var key   = 'login_' + username;
    var raw   = cache.get(key);
    if (!raw) {
      return checkRateLimitSheet_(username);
    }
    var data  = JSON.parse(raw);
    var now   = Date.now();
    if (data.lockedUntil && data.lockedUntil > now) {
      var remaining = Math.ceil((data.lockedUntil - now) / 60000);
      return { blocked: true, message: 'บัญชีถูกล็อคชั่วคราว กรุณารอ ' + remaining + ' นาที' };
    }
    return { blocked: false };
  } catch (e) {
    Logger.log('checkRateLimit_ error: ' + e.message);
    writeLog('RATE_LIMIT_ERROR', e.message, username || 'unknown');
    return { blocked: true, message: 'ระบบไม่สามารถตรวจสอบได้ชั่วคราว กรุณาลองใหม่ในอีกสักครู่' };
  }
}
function checkRateLimitSheet_(username) {
  try {
    var sheet = getSheet('FailedLogins');
    if (!sheet || sheet.getLastRow() < 2) return { blocked: false };
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h||'').trim().toLowerCase(); });
    var uCol = headers.indexOf('username');
    var cCol = headers.indexOf('count');
    var tCol = headers.indexOf('locked_until');
    if (uCol === -1) return { blocked: false };
    var now = Date.now();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uCol]||'').trim() === String(username).trim()) {
        var lockedUntil = tCol !== -1 ? Number(data[i][tCol]||0) : 0;
        if (lockedUntil > now) {
          var remaining = Math.ceil((lockedUntil - now) / 60000);
          return { blocked: true, message: 'บัญชีถูกล็อคชั่วคราว กรุณารอ ' + remaining + ' นาที' };
        }
        return { blocked: false };
      }
    }
    return { blocked: false };
  } catch (e) {
    Logger.log('checkRateLimitSheet_ error: ' + e.message);
    return { blocked: false };
  }
}

function ensureUsersSheet() {
  try {
    var ss = SpreadsheetApp.openById(getSheetId_());
    var sheet = ss.getSheetByName('Users');
    if (!sheet) {
      sheet = ss.insertSheet('Users');
      var headers = ['username', 'password', 'password_hash', 'name', 'name_en', 'email', 'phone',
                     'faculty', 'department', 'position', 'roles', 'status', 'registered_date'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setBackground('#2d8a52')
           .setFontColor('#fff')
           .setFontWeight('bold');
    }
    Logger.log('Users sheet ready (no test users added)');
  } catch (e) {
    Logger.log('Error ensuring users sheet: ' + e.message);
  }
}
function isAdmin_(username) {
  try {
    var u = sheetToObjects('Users').find(function(u) {
      return String(u['username'] || '').trim() === String(username || '').trim();
    });
    return u ? String(u['roles'] || '').indexOf('admin') !== -1 : false;
  } catch (e) { return false; }
}
function changePassword(sessionId, username, oldPassword, newPassword) {
  try {
    var auth = requireValidSession_(sessionId, username);
    if (!auth.ok) return { success: false, message: auth.message };
    var safeUsername = auth.username;
    if (!oldPassword || !newPassword) {
      return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };
    }
    if (newPassword.length < 8) {
      return { success: false, message: 'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร' };
    }
    if (!/(?=.*[a-z])(?=.*[A-Z])(?=.*\d)/.test(newPassword)) {
      return { success: false, message: 'รหัสผ่านต้องประกอบด้วย ตัวพิมพ์เล็ก ตัวพิมพ์ใหญ่ และตัวเลข' };
    }
    var user = sheetToObjects('Users').find(function(u) {
      return String(u['username'] || '').trim() === safeUsername;
    });
    if (!user) return { success: false, message: 'ไม่พบผู้ใช้' };
    var storedHash = String(user['password_hash'] || '').trim();
    if (!storedHash) return { success: false, message: 'ไม่สามารถตรวจสอบรหัสผ่านได้ กรุณาติดต่อผู้ดูแล' };
    var isOldCorrect = verifyPassword_(oldPassword, storedHash);
    if (!isOldCorrect) {
      return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
    }
    var newHash = hashPassword_(newPassword);
    var ok = updateRowByField('Users', 'username', safeUsername, {
      password_hash: newHash,
      password: ''
    });
    if (!ok) return { success: false, message: 'ไม่สามารถเปลี่ยนรหัสผ่านได้' };
    writeLog('CHANGE_PASSWORD', safeUsername, safeUsername);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function updateProfile(sessionId, username, payload) {
  var auth = requireValidSession_(sessionId, username);
  if (!auth.ok) return { success: false, message: auth.message };
  var safeUsername = auth.username;
  var ALLOWED_FIELDS = ['name', 'email', 'phone', 'faculty', 'department', 'position'];
  var updates = {};
  Object.keys(payload).forEach(function(key) {
    if (ALLOWED_FIELDS.indexOf(key) !== -1) {
      updates[key] = sanitizeInput_(payload[key]);
    } else {
      Logger.log('⚠️ Blocked field update attempt: ' + key);
    }
  });
  if (!Object.keys(updates).length) {
    return { success: false, message: 'ไม่มีฟิลด์ที่อนุญาต' };
  }
  return updateRowByField('Users', 'username', safeUsername, updates);
}
function sanitizeInput_(text) {
  if (!text) return '';
  return String(text)
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;')
    .replace(/\\/g, '&#x5C;')
    .trim();
}
function sanitizeEmail_(email) {
  if (!email) return '';
  var cleaned = String(email).trim().toLowerCase();
  if (!validateEmail_(cleaned)) return '';
  return cleaned;
}
function registerUser(payload) {
  try {
    var required = ['username', 'email', 'password', 'name', 'phone', 'faculty', 'department', 'position', 'roles'];
    for (var i = 0; i < required.length; i++) {
      if (!payload[required[i]] || !String(payload[required[i]]).trim()) {
        return { success: false, message: 'กรุณากรอก "' + required[i] + '" ให้ครบ' };
      }
    }
    if (!/^[a-zA-Z0-9_]{4,20}$/.test(payload.username)) {
      return { success: false, message: 'Username ต้องเป็นภาษาอังกฤษหรือตัวเลข 4-20 ตัว' };
    }
    if (payload.password.length < 8) {
      return { success: false, message: 'รหัสผ่านต้องมีอย่างน้อย 8 ตัว' };
    }
    if (!/(?=.*[a-z])(?=.*[A-Z])(?=.*\d)/.test(payload.password)) {
      return {
        success: false,
        message: 'รหัสผ่านต้องประกอบด้วย ตัวพิมพ์เล็ก ตัวพิมพ์ใหญ่ และตัวเลข'
      };
    }
    var email = sanitizeEmail_(payload.email);
    if (!email) {
      return { success: false, message: 'รูปแบบอีเมลไม่ถูกต้อง' };
    }
    var users = sheetToObjects('Users');
    if (users.find(function(u) {
      return String(u['username'] || '').toLowerCase() === payload.username.toLowerCase();
    })) {
      return { success: false, message: 'ชื่อผู้ใช้นี้ถูกใช้งานแล้ว' };
    }
    if (users.find(function(u) {
      return String(u['email'] || '').toLowerCase() === email;
    })) {
      return { success: false, message: 'อีเมลนี้ถูกใช้งานแล้ว' };
    }
    var passwordHash = hashPassword_(payload.password);
    var name = sanitizeInput_(payload.name);
    var nameEn = sanitizeInput_(payload.nameEn || '');
    var idCard = sanitizeInput_(payload.idCard || '');
    var phone = sanitizeInput_(payload.phone);
    var faculty = sanitizeInput_(payload.faculty);
    var department = sanitizeInput_(payload.department);
    var position = sanitizeInput_(payload.position);
    appendRow('Users', {
      username: payload.username,
      password: '',
      password_hash: passwordHash,
      name: name,
      name_en: nameEn,
      email: email,
      id_card: idCard,
      phone: phone,
      faculty: faculty,
      department: department,
      position: position,
      roles: 'pending_' + payload.roles,
      status: 'pending',
      registered_date: new Date()
    });
    writeLog('REGISTER', payload.username + ' | ' + email, payload.username);
    notifyStaff_(
      '[EC Sansai] มีผู้ลงทะเบียนใหม่รอการอนุมัติ',
      'ผู้ใช้ใหม่รอการอนุมัติ:\nชื่อ: ' + name + '\nUsername: ' + payload.username +
      '\nอีเมล: ' + email
    );
    sendEmail_(
      email,
      '[EC Sansai] รับการลงทะเบียนเรียบร้อย',
      'คุณ ' + name + '\n\nระบบได้รับการลงทะเบียนของท่านแล้ว\n' +
      'กรุณารอการอนุมัติจากผู้ดูแลระบบ\n\nUsername: ' + payload.username
    );
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
var RESET_TOKEN_EXPIRY = 30 * 60;
function requestPasswordReset(email) {
  try {
    if (!email || !validateEmail_(email)) {
      return { success: false, message: 'กรุณากรอกอีเมลให้ถูกต้อง' };
    }
    var users = sheetToObjects('Users');
    var user = users.find(function(u) {
      return String(u['email'] || '').toLowerCase() === String(email).toLowerCase();
    });
    if (!user) {
      Logger.log('Password reset requested for non-existent email: ' + email);
      return {
        success: true,
        message: 'หากอีเมลนี้มีในระบบ เราจะส่งลิงก์รีเซ็ตรหัสผ่านไปให้'
      };
    }
    var username = String(user['username'] || '').trim();
    var name = String(user['name'] || username);
    var resetToken = Utilities.getUuid();
    var expiresAt = Date.now() + (RESET_TOKEN_EXPIRY * 1000);
    var cache = CacheService.getScriptCache();
    cache.put('reset_' + resetToken, JSON.stringify({
      username: username,
      email: email,
      expiresAt: expiresAt
    }), RESET_TOKEN_EXPIRY);
    var baseUrl = CUSTOM_DOMAIN;
    var resetLink = baseUrl + '?resetToken=' + encodeURIComponent(resetToken);
    var emailBody =
      'สวัสดีคุณ ' + name + '\n\n' +
      'เราได้รับคำขอรีเซ็ตรหัสผ่านสำหรับบัญชีของคุณ (Username: ' + username + ')\n\n' +
      'กรุณาคลิกลิงก์ด้านล่างเพื่อตั้งรหัสผ่านใหม่:\n' +
      resetLink + '\n\n' +
      'ลิงก์นี้จะหมดอายุใน 30 นาที\n\n' +
      'หากคุณไม่ได้ขอรีเซ็ตรหัสผ่าน กรุณาละเลยอีเมลนี้\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'ระบบ EC Online Submission\n' +
      'โรงพยาบาลสันทราย';
    var emailSent = sendEmail_(
      email,
      '[EC Sansai] คำขอรีเซ็ตรหัสผ่าน',
      emailBody
    );
    if (emailSent) {
      writeLog('PASSWORD_RESET_REQUEST', 'ส่งลิงก์รีเซ็ตรหัสผ่านไปยัง: ' + email, 'system');
    } else {
      Logger.log('Failed to send password reset email to: ' + email);
    }
    return {
      success: true,
      message: 'หากอีเมลนี้มีในระบบ เราจะส่งลิงก์รีเซ็ตรหัสผ่านไปให้'
    };
  } catch (e) {
    Logger.log('requestPasswordReset error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด กรุณาลองอีกครั้ง' };
  }
}
function validateResetToken(token) {
  try {
    if (!token) return { success: false };
    var cache = CacheService.getScriptCache();
    var data = cache.get('reset_' + token);
    if (!data) return { success: false };
    var resetData = JSON.parse(data);
    var now = Date.now();
    if (resetData.expiresAt < now) {
      cache.remove('reset_' + token);
      return { success: false, message: 'ลิงก์หมดอายุแล้ว กรุณาขอใหม่อีกครั้ง' };
    }
    return {
      success: true,
      username: resetData.username,
      email: resetData.email
    };
  } catch (e) {
    Logger.log('validateResetToken error: ' + e.message);
    return { success: false };
  }
}
function resetPassword(token, newPassword, confirmPassword) {
  try {
    if (!token || !newPassword || !confirmPassword) {
      return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };
    }
    if (newPassword !== confirmPassword) {
      return { success: false, message: 'รหัสผ่านไม่ตรงกัน' };
    }
    if (newPassword.length < 8) {
      return { success: false, message: 'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร' };
    }
    if (!/(?=.*[a-z])(?=.*[A-Z])(?=.*\d)/.test(newPassword)) {
      return {
        success: false,
        message: 'รหัสผ่านต้องประกอบด้วย ตัวพิมพ์เล็ก ตัวพิมพ์ใหญ่ และตัวเลข'
      };
    }
    var validation = validateResetToken(token);
    if (!validation.success) {
      return {
        success: false,
        message: validation.message || 'ลิงก์ไม่ถูกต้องหรือหมดอายุแล้ว'
      };
    }
    var users = sheetToObjects('Users');
    var user = users.find(function(u) {
      return String(u['username'] || '').trim() === validation.username;
    });
    if (!user) {
      return { success: false, message: 'ไม่พบผู้ใช้' };
    }
    var newHash = hashPassword_(newPassword);
    var updated = updateRowByField('Users', 'username', validation.username, {
      password_hash: newHash,
      password: ''
    });
    if (!updated) {
      return { success: false, message: 'ไม่สามารถเปลี่ยนรหัสผ่านได้ กรุณาลองอีกครั้ง' };
    }
    var cache = CacheService.getScriptCache();
    cache.remove('reset_' + token);
    writeLog('PASSWORD_RESET_SUCCESS', 'รีเซ็ตรหัสผ่านสำเร็จ: ' + validation.username, validation.username);
    sendEmail_(
      validation.email,
      '[EC Sansai] การเปลี่ยนรหัสผ่านสำเร็จ',
      'รหัสผ่านบัญชี ' + validation.username + ' ของคุณถูกเปลี่ยนเรียบร้อยแล้ว\n\n' +
      'หากคุณไม่ได้เป็นผู้ดำเนินการนี้ กรุณาติดต่อเจ้าหน้าที่ทันที'
    );
    return { success: true };
  } catch (e) {
    Logger.log('resetPassword error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด กรุณาลองอีกครั้ง' };
  }
}
function migratePasswordsToHash() {
  try {
    var sheet = getSheet('Users');
    if (!sheet) {
      return '❌ Users sheet not found';
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return '❌ No users to migrate';
    }
    var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var pwCol = headers.indexOf('password');
    var pwHashCol = headers.indexOf('password_hash');
    if (pwCol === -1) {
      return '❌ password column not found';
    }
    if (pwHashCol === -1) {
      pwHashCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, pwHashCol).setValue('password_hash')
        .setBackground('#2d8a52').setFontColor('#fff').setFontWeight('bold');
    }
    var migrated = 0;
    var skipped = 0;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var plainPassword = String(row[pwCol] || '').trim();
      var existingHash = String(row[pwHashCol] || '').trim();
      if (!plainPassword || existingHash) {
        skipped++;
        continue;
      }
      var hash = hashPassword_(plainPassword);
      sheet.getRange(i + 1, pwHashCol + 1).setValue(hash);
      sheet.getRange(i + 1, pwCol + 1).setValue('');
      migrated++;
    }
    writeLog('PASSWORD_MIGRATION', 'Migrated: ' + migrated + ', Skipped: ' + skipped, 'system');
    return '✅ Password migration complete\n' +
           'Migrated: ' + migrated + ' users\n' +
           'Skipped: ' + skipped + ' users\n\n' +
           '⚠️ IMPORTANT: Plain text passwords are still in sheet for backup.\n' +
           'After verifying all users can login, manually clear the "password" column.';
  } catch (e) {
    return '❌ Migration failed: ' + e.message;
  }
}
function migrateEncryptUsers() {
  var FIELDS_TO_ENCRYPT = ['email', 'phone', 'id_card'];
  var ss     = SpreadsheetApp.openById(getSheetId_());
  var sheet  = ss.getSheetByName('Users');
  var data   = sheet.getDataRange().getValues();
  var headers= data[0];
  var updated = 0;
  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var changed = false;
    FIELDS_TO_ENCRYPT.forEach(function(fieldName) {
      var colIdx = headers.indexOf(fieldName);
      if (colIdx === -1) return;
      var val = row[colIdx];
      if (val && !isEncrypted_(val)) {
        row[colIdx] = encryptField_(val);
        changed = true;
      }
    });
    if (changed) {
      sheet.getRange(i+1, 1, 1, row.length).setValues([row]);
      updated++;
    }
  }
  Logger.log('✅ เข้ารหัสแล้ว ' + updated + ' แถว ใน Users');
  writeAuditLog_('ENCRYPT_MIGRATION', 'Users: '+updated+' rows', 'admin');
}
function createSession_(username, ipAddress, userAgent) {
  ensureSessionsSheet_();
  var sessionId = Utilities.getUuid();
  var now = new Date();
  var expiresAt = new Date(now.getTime() + (8 * 60 * 60 * 1000));
  appendRow('Sessions', {
    session_id: sessionId,
    username: username,
    created_at: now,
    expires_at: expiresAt,
    ip_address: ipAddress,
    user_agent: userAgent,
    is_active: true
  });
  return sessionId;
}
function validateSession(sessionId, usernameFromClient) {
  if (!sessionId) return { success: false, message: 'No session' };
  var sessions = sheetToObjects('Sessions');
  var session = sessions.find(function(s) {
    return s.session_id === sessionId &&
           (s.is_active === true || String(s.is_active).toLowerCase() === 'true');
  });
  if (!session) return { success: false, message: 'Invalid session' };
  var trustedUsername = String(session['username'] || '').trim();
  var expiresAt = new Date(session.expires_at);
  if (expiresAt < new Date()) {
    updateRowByField('Sessions', 'session_id', sessionId, { is_active: false });
    return { success: false, message: 'Session expired' };
  }
  var u = sheetToObjects('Users').find(function(u) {
    return String(u['username'] || '').trim() === trustedUsername;
  });
  if (!u) return { success: false, message: 'User not found' };
  var roles = String(u['roles'] || 'researcher').split(',').map(function(r) {
    r = r.trim().replace('pending_', '');
    return r === 'user' ? 'researcher' : r;
  }).filter(Boolean);
  if (!roles.length) roles = ['researcher'];
  return {
    success:    true,
    sessionId:  sessionId,
    username:   trustedUsername,
    name:       String(u['name'] || trustedUsername),
    roles:      roles
  };
}
function getUserProfile(sessionId, username) {
  try {
    var auth = requireValidSession_(sessionId, username);
    if (!auth.ok) return { success: false, message: auth.message };
    var users = sheetToObjects('Users');
    var u = users.find(function(u) {
      return String(u['username'] || '').trim() === auth.username;
    });
    if (!u) return { success: false, message: 'User not found' };
    return {
      success:    true,
      name:       String(u['name']       || ''),
      email:      decryptField_(String(u['email']      || '')),
      phone:      decryptField_(String(u['phone']      || '')),
      faculty:    String(u['faculty']    || ''),
      department: String(u['department'] || ''),
      position:   String(u['position']  || '')
    };
  } catch (e) {
    Logger.log('getUserProfile error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}
function sanitizeEmailContent_(text, maxLength) {
  maxLength = maxLength || 2000;
  return String(text || '')
    .replace(/[\r\n]{3,}/g, '\n\n')
    .replace(/[^\S\r\n]{10,}/g, ' ')
    .trim()
    .substring(0, maxLength);
}
function getAllUsers(sessionId, callerUsername) {
  try {
    var auth = requireValidSession_(sessionId, callerUsername);
    if (!auth.ok) return [];
    var safeUsername = auth.username;
    var user = sheetToObjects('Users').find(function(u) {
      return String(u['username'] || '').trim() === safeUsername;
    });
    if (!user || String(user['roles'] || '').indexOf('admin') === -1) {
      writeAuditLog_('AUTH_GUARD_FAIL', 'getAllUsers non-admin attempt', safeUsername, 'WARN');
      return [];
    }
    return sheetToObjects('Users').map(function(u) {
      return {
        username:        String(u['username']        || ''),
        name:            String(u['name']            || ''),
        email:           String(u['email']           || ''),
        phone:           String(u['phone']           || ''),
        faculty:         String(u['faculty']         || ''),
        department:      String(u['department']      || ''),
        position:        String(u['position']        || ''),
        roles:           String(u['roles']           || ''),
        status:          String(u['status']          || 'active'),
        registered_date: u['registered_date'] ? String(u['registered_date']) : ''
      };
    });
  } catch (e) {
    Logger.log('getAllUsers error: ' + e.message);
    return [];
  }
}
function createUser(payload, adminUsername) {
  try {
    if (!isAdmin_(adminUsername)) return { success: false, message: 'คุณไม่มีสิทธิ์สร้างผู้ใช้งาน' };
    var username = String(payload.username||'').trim();
    var password = String(payload.password||'');
    var name     = String(payload.name||'').trim();
    var email    = String(payload.email||'').trim().toLowerCase();
    var roles    = String(payload.roles||'').trim();
    if (!username||!password||!name||!email||!roles) return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };
    if (!/^[a-zA-Z0-9_]{4,20}$/.test(username)) return { success: false, message: 'Username ไม่ถูกรูปแบบ' };
    if (password.length < 8) return { success: false, message: 'รหัสผ่านต้องมีอย่างน้อย 8 ตัว' };
    if (!/(?=.*[a-z])(?=.*[A-Z])(?=.*\d)/.test(password)) return { success: false, message: 'รหัสผ่านต้องประกอบด้วย ตัวพิมพ์เล็ก ตัวพิมพ์ใหญ่ และตัวเลข' };
    if (!validateEmail_(email)) return { success: false, message: 'รูปแบบอีเมลไม่ถูกต้อง' };
    var users = sheetToObjects('Users');
    if (users.find(function(u){ return String(u['username']||'').toLowerCase() === username.toLowerCase(); }))
      return { success: false, message: 'Username "' + username + '" ถูกใช้งานแล้ว' };
    if (users.find(function(u){ return String(u['email']||'').toLowerCase() === email; }))
      return { success: false, message: 'อีเมล "' + email + '" ถูกใช้งานแล้ว' };
    var passwordHash = hashPassword_(password);
    appendRow('Users', {
      username:      username,
      password:      '',
      password_hash: passwordHash,
      name:          name,
      name_en:       '',
      email:         email,
      id_card:       '',
      phone:         payload.phone||'',
      faculty:       payload.faculty||'',
      department:    payload.department||'',
      position:      payload.position||'',
      roles:         roles,
      status:        payload.status||'active',
      registered_date: new Date()
    });
    writeLog('CREATE_USER', 'สร้าง: ' + username + ' | roles: ' + roles, adminUsername);
    if ((payload.status||'active') === 'active') {
      sendEmail_(email, '[EC Sansai] บัญชีผู้ใช้งานของท่านพร้อมใช้งาน',
        'คุณ ' + name + '\n\nผู้ดูแลระบบได้สร้างบัญชีให้ท่านแล้ว\n' +
        'Username: ' + username + '\nสิทธิ์: ' + roles + '\n\nกรุณาเข้าสู่ระบบและเปลี่ยนรหัสผ่านทันที');
    }
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function assertOwnsProject_(projectCode, callerUsername) {
  var proj = sheetToObjects('Projects').find(function(p) {
    return String(p['project_code'] || '').trim() === String(projectCode).trim();
  });
  if (!proj) throw new Error('ไม่พบโครงการ');
  var baseCode = String(projectCode).replace(/_fix\d+$/, '');
  var ownsBase = sheetToObjects('Projects').find(function(p) {
    return String(p['project_code'] || '').trim() === baseCode &&
           String(p['submitted_by'] || '').trim() === String(callerUsername).trim();
  });
  var ownsExact = String(proj['submitted_by'] || '').trim() === String(callerUsername).trim();
  if (!ownsExact && !ownsBase) {
    throw new Error('Forbidden: ไม่มีสิทธิ์จัดการโครงการนี้');
  }
  return proj;
}
function updateUser(payload, adminUsername) {
  try {
    if (!isAdmin_(adminUsername)) return { success: false, message: 'คุณไม่มีสิทธิ์แก้ไขผู้ใช้' };
    var username = String(payload.username||'').trim();
    if (!username) return { success: false, message: 'ไม่ระบุ username' };
    var users = sheetToObjects('Users');
    if (!users.find(function(u){ return String(u['username']||'').trim() === username; }))
      return { success: false, message: 'ไม่พบผู้ใช้ "' + username + '"' };
    if (payload.password && String(payload.password).length < 6)
      return { success: false, message: 'รหัสผ่านต้องมีอย่างน้อย 6 ตัว' };
    var updates = {};
    ['name','name_en','email','phone','faculty','department','position','roles','status'].forEach(function(f) {
      if (payload[f] !== undefined) updates[f] = payload[f];
    });
    if (payload.password && String(payload.password).trim()) {
  updates['password_hash'] = hashPassword_(payload.password);
  updates['password'] = '';
}
    if (!updateRowByField('Users', 'username', username, updates)) return { success: false, message: 'ไม่สามารถอัพเดทได้' };
    writeLog('UPDATE_USER', 'แก้ไข: ' + username, adminUsername);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteUser(sessionId, targetUsername, adminUsername) {
  try {
    var auth = requireValidSession_(sessionId, adminUsername);
    if (!auth.ok) return { success: false, message: auth.message };
    var safeAdmin = auth.username;
    if (!isAdmin_(safeAdmin)) return { success: false, message: 'คุณไม่มีสิทธิ์ลบผู้ใช้' };
    if (String(targetUsername).trim() === safeAdmin)
      return { success: false, message: 'ไม่สามารถลบบัญชีของตัวเองได้' };
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });
    var uCol = headers.indexOf('username');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uCol]).trim() === String(targetUsername).trim()) {
        sheet.deleteRow(i + 1);
        writeLog('DELETE_USER', 'ลบ: ' + targetUsername, safeAdmin);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (e) { return { success: false, message: e.message }; }
}
function approveUser(sessionId, targetUsername, adminUsername) {
  try {
    var auth = requireValidSession_(sessionId, adminUsername);
    if (!auth.ok) return { success: false, message: auth.message };
    var safeAdmin = auth.username;
    if (!isAdmin_(safeAdmin)) return { success: false, message: 'คุณไม่มีสิทธิ์อนุมัติผู้ใช้' };
    var user = sheetToObjects('Users').find(function(u){ return String(u['username']||'').trim() === targetUsername.trim(); });
    if (!user) return { success: false, message: 'ไม่พบผู้ใช้' };
    var newRoles = String(user['roles']||'').replace(/pending_/g,'').trim();
    if (!updateRowByField('Users','username',targetUsername,{roles:newRoles,status:'active'}))
      return { success: false, message: 'ไม่สามารถอนุมัติได้' };
    writeLog('APPROVE_USER', targetUsername + ' → ' + newRoles, safeAdmin);
    if (user['email'] && validateEmail_(user['email'])) {
      sendEmail_(user['email'], '[EC Sansai] บัญชีของท่านได้รับการอนุมัติแล้ว',
        'คุณ ' + (user['name']||targetUsername) + '\n\nบัญชีของท่านได้รับการอนุมัติเรียบร้อยแล้ว\n' +
        'Username: ' + targetUsername + '\nสิทธิ์: ' + newRoles + '\n\nท่านสามารถเข้าสู่ระบบได้ทันที');
    }
    return { success: true, roles: newRoles };
  } catch (e) { return { success: false, message: e.message }; }
}
function generateProjectCode_() {
  try {
    var year = new Date().getFullYear() + 543;
    var sheet = getSheet('Projects');
    var allProjects = sheetToObjects('Projects');
    var mainProjects = allProjects.filter(function(p) {
      return String(p['project_code']||'').indexOf('_fix') === -1;
    });
    var count = mainProjects.length;
    return 'Sansai-' + year + '-' + String(count + 1).padStart(4, '0');
  } catch (e) { return 'Sansai-' + new Date().getTime(); }
}
function generateRevisionProjectCode_(originalCode, revRound) {
  var baseCode = String(originalCode).replace(/_fix\d+$/, '');
  return baseCode + '_fix' + revRound;
}
function ensureProjectsSheet() {
  var ss = SpreadsheetApp.openById(getSheetId_());
  var sheet = ss.getSheetByName('Projects');
  if (!sheet) sheet = ss.insertSheet('Projects');
  if (sheet.getLastRow() === 0) {
    var headers = [
      'project_code','title_th','title_en','project_type','pi_name',
      'pi_dept','pi_phone','pi_email','objectives','methodology',
      'expected_benefits','budget','duration','gcp_cert','status',
      'submitted_by','created_date','updated_date',
      'final_result','result_note','revision_deadline','revision_round',
      'result_sent_date','parent_code','doc_verified','doc_verified_by','doc_verified_date',
      'cert_file_id',
      'revision_file_id',
      'skip_committee_reeval'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setBackground('#2d8a52').setFontColor('#fff').setFontWeight('bold');
  }
  return sheet;
}
function getAllProjectsRaw(callerUsername) {
  requireRole_(callerUsername, ['admin', 'staff']);
  try {
    var sheet = getSheet('Projects');
    if (!sheet) return { success: false, message: 'ไม่พบ Projects sheet', data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.every(function(c){ return c===''||c===null||c===undefined; })) continue;
      var obj = {};
      headers.forEach(function(h,j){ obj[h] = String(row[j]!==null&&row[j]!==undefined?row[j]:''); });
      result.push({
        project_code:         obj['project_code']||'',
        title_th:             obj['title_th']||'(ไม่มีชื่อ)',
        title_en:             obj['title_en']||'',
        project_type:         obj['project_type']||'',
        pi_name:              obj['pi_name']||'',
        pi_dept:              obj['pi_dept']||'',
        pi_phone:             obj['pi_phone']||'',
        pi_email:             obj['pi_email']||'',
        status:               obj['status']||'รอพิจารณา',
        submitted_by:         obj['submitted_by']||'',
        created_date:         obj['created_date']||'',
        updated_date:         obj['updated_date']||'',
        budget:               obj['budget']||'0',
        duration:             obj['duration']||'',
        doc_verified:         obj['doc_verified']||'',
        doc_verified_by:      obj['doc_verified_by']||'',
        doc_verified_date:    obj['doc_verified_date']||'',
        final_result:         obj['final_result']||'',
        revision_file_id:     obj['revision_file_id']||'',
        cert_file_id:         obj['cert_file_id']||'',
        skip_committee_reeval: obj['skip_committee_reeval'] || ''
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
function getMyProjects(sessionId, username) {
  try {
    var auth = requireValidSession_(sessionId, username);
    if (!auth.ok) return { success: false, message: auth.message, data: [] };
    var safeUsername = auth.username;
    var rows = sheetToObjects('Projects').filter(function(r){
      return String(r['submitted_by']||'').trim() === safeUsername;
    });
    return {
      success: true,
      data: rows.map(function(r){
        return {
          project_code: String(r['project_code']||''),
          title_th:     String(r['title_th']||''),
          project_type: String(r['project_type']||''),
          pi_name:      String(r['pi_name']||''),
          status:       String(r['status']||'รอพิจารณา'),
          submitted_by: String(r['submitted_by']||''),
          created_date: r['created_date'] ? String(r['created_date']) : ''
        };
      })
    };
  } catch (e) { return { success: false, message: e.message, data: [] }; }
}
function getMyProjectsWithRevision(sessionId, username) {
  try {
    var auth = requireValidSession_(sessionId, username);
    if (!auth.ok) return { success: false, message: auth.message, data: [] };
    var safeUsername = auth.username;
    var user = sheetToObjects('Users').find(function(u) {
    return String(u['username'] || '').trim() === safeUsername;
  });
  if (!user || user.status !== 'active') {
    return { success: false, message: 'User inactive', data: [] };
  }
  var rows = sheetToObjects('Projects').filter(function(r) {
    return String(r['submitted_by'] || '').trim() === safeUsername;
  });
    return {
      success: true,
      data: rows.map(function(r) {
        var finalResult = String(r['final_result'] || '').trim();
        if (!finalResult) {
          var status = String(r['status'] || '');
          if (status.indexOf('รับรองมีข้อแก้ไข') !== -1) finalResult = 'รับรองมีข้อแก้ไข';
          else if (status.indexOf('รับรอง') !== -1) finalResult = 'รับรอง';
          else if (status.indexOf('ไม่รับรอง') !== -1) finalResult = 'ไม่รับรอง';
        }
        return {
          project_code: String(r['project_code'] || ''),
          title_th: String(r['title_th'] || ''),
          project_type: String(r['project_type'] || ''),
          pi_name: String(r['pi_name'] || ''),
          status: String(r['status'] || 'รอพิจารณา'),
          final_result: finalResult,
          result_note: String(r['result_note'] || ''),
          revision_deadline: String(r['revision_deadline'] || ''),
          revision_round: String(r['revision_round'] || ''),
          result_sent_date: String(r['result_sent_date'] || ''),
          submitted_by: String(r['submitted_by'] || ''),
          created_date: r['created_date'] ? String(r['created_date']) : '',
          revision_file_id: String(r['revision_file_id'] || ''),
cert_file_id: String(r['cert_file_id'] || ''),
          updated_date: r['updated_date'] ? String(r['updated_date']) : ''
        };
      })
    };
  } catch (e) {
    Logger.log('getMyProjectsWithRevision error: ' + e.message);
    return { success: false, message: e.message, data: [] };
  }
}
function saveOrUpdateProject(projectData, isUpdate) {
  try {
    ensureProjectsSheet();
    if (!projectData) return { success: false, message: 'ไม่มีข้อมูลโครงการ' };
    var ts = new Date();
    if (isUpdate) {
      var code = projectData.project_code || projectData.projectCode;
      if (!code) return { success: false, message: 'ไม่ระบุรหัสโครงการ' };
      updateRowByField('Projects', 'project_code', code, {
        title_th: projectData.title_th||'', title_en: projectData.title_en||'',
        objectives: projectData.objectives||'', methodology: projectData.methodology||'',
        expected_benefits: projectData.expectedBenefits||'', budget: projectData.budget||'0',
        duration: projectData.duration||'', status: 'รอพิจารณา', updated_date: ts
      });
      if (projectData.fileIds && projectData.fileIds.length) updateUploadedFilesProjectCode_(projectData.fileIds, code);
      writeLog('UPDATE_PROJECT', code, projectData.submittedBy||'system');
      return { success: true, projectId: code };
    } else {
      var projectCode = generateProjectCode_();
      var piName = (String(projectData.title_name||'') + String(projectData.first_name||'') + ' ' + String(projectData.last_name||'')).trim();
      appendRow('Projects', {
        project_code: projectCode, title_th: projectData.title_th||'', title_en: projectData.title_en||'',
        project_type: projectData.type==='1'?'new':'cont', pi_name: piName,
        pi_dept: projectData.location||'', pi_phone: projectData.tel||'', pi_email: projectData.email||'',
        objectives: projectData.objectives||'', methodology: projectData.methodology||'',
        expected_benefits: projectData.expectedBenefits||'', budget: projectData.budget||'0',
        duration: projectData.duration||'', gcp_cert: projectData.cert_date||'',
        status: 'รอพิจารณา', submitted_by: projectData.submittedBy||'',
        created_date: ts, updated_date: ts,
        doc_verified: '', doc_verified_by: '', doc_verified_date: '', parent_code: ''
      });
      if (projectData.fileIds && projectData.fileIds.length) updateUploadedFilesProjectCode_(projectData.fileIds, projectCode);
      addNotification_('staff', '', projectCode,
        '📥 มีโครงการวิจัยใหม่รอตรวจสอบเอกสาร: ' + projectCode + ' — ' + (projectData.title_th||''));
      notifyStaff_('[EC Sansai] โครงการใหม่รอตรวจสอบเอกสาร: ' + projectCode,
        'มีโครงการวิจัยใหม่ส่งเข้าระบบ:\nรหัส: ' + projectCode + '\nชื่อ: ' + (projectData.title_th||'-') +
        '\nหัวหน้าโครงการ: ' + piName + '\nวันที่: ' + ts.toLocaleString('th-TH') +
        '\n\n⚠ กรุณาตรวจสอบเอกสารก่อนมอบหมายให้กรรมการ');
      writeLog('SUBMIT_PROJECT', projectCode + ' | ' + (projectData.title_th||''), projectData.submittedBy);
      return { success: true, projectId: projectCode };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function updateUploadedFilesProjectCode_(fileIds, projectCode) {
  try {
    var sheet = getSheet('UploadedFiles');
    if (!sheet || !fileIds || !fileIds.length) return;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    var headers = data[0].map(function(h){ return String(h).trim(); });
    var fidCol  = headers.indexOf('file_id');
    var codeCol = headers.indexOf('project_code');
    if (fidCol === -1 || codeCol === -1) return;
    fileIds.forEach(function(fid) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][fidCol]).trim() === String(fid).trim()) {
          sheet.getRange(i+1, codeCol+1).setValue(projectCode);
        }
      }
    });
  } catch (e) { Logger.log('updateUploadedFilesProjectCode_ error: ' + e.message); }
}
function updateProjectStatus(projectCode, newStatus, staffUsername, note) {
  try {
    var ok = updateRowByField('Projects','project_code',projectCode,{status:newStatus,updated_date:new Date()});
    if (ok) {
      writeLog('UPDATE_STATUS', projectCode + ' → ' + newStatus + (note?' | '+note:''), staffUsername);
      return { success: true };
    }
    return { success: false, message: 'ไม่พบโครงการ' };
  } catch (e) { return { success: false, message: e.message }; }
}
function verifyProjectDocuments(projectCode, decision, staffUsername, note) {
  try {
    if (!projectCode || !decision) return { success: false, message: 'ข้อมูลไม่ครบ' };
    var ts = new Date();
    var newStatus = decision === 'approved' ? 'ผ่านตรวจเอกสาร' : 'เอกสารไม่ครบ';
    var ok = updateRowByField('Projects', 'project_code', projectCode, {
      status: newStatus,
      doc_verified: decision,
      doc_verified_by: staffUsername,
      doc_verified_date: ts,
      updated_date: ts
    });
    if (!ok) return { success: false, message: 'ไม่พบโครงการ' };
    var proj = sheetToObjects('Projects').find(function(p) {
      return String(p['project_code']||'').trim() === projectCode;
    });
    if (proj) {
      var submittedBy = String(proj['submitted_by']||'');
      var resUser = submittedBy ? sheetToObjects('Users').find(function(u) {
        return String(u['username']||'').trim() === submittedBy;
      }) : null;
      var piEmail = (resUser && resUser['email']) ? String(resUser['email']).trim() : String(proj['pi_email']||'').trim();
      if (decision === 'approved') {
        addNotification_('researcher', submittedBy, projectCode,
          '✅ เอกสารโครงการ ' + projectCode + ' ผ่านการตรวจสอบแล้ว รออยู่ในขั้นตอนมอบหมายกรรมการ');
        if (piEmail && validateEmail_(piEmail)) {
          sendEmail_(piEmail, '[EC Sansai] ✅ เอกสารโครงการผ่านการตรวจสอบ: ' + projectCode,
            'เรียนคุณ ' + (proj['pi_name']||submittedBy) + '\n\nเอกสารโครงการ ' + projectCode + ' — ' + (proj['title_th']||'') +
            '\nผ่านการตรวจสอบเรียบร้อยแล้ว\n\nขณะนี้อยู่ในขั้นตอนมอบหมายให้คณะกรรมการพิจารณา\n' +
            (note ? '\nหมายเหตุจากเจ้าหน้าที่: ' + note : ''));
        }
      } else {
        addNotification_('researcher', submittedBy, projectCode,
          '⚠️ เอกสารโครงการ ' + projectCode + ' ไม่ครบถ้วน กรุณาตรวจสอบและส่งเพิ่มเติม');
        if (piEmail && validateEmail_(piEmail)) {
          sendEmail_(piEmail, '[EC Sansai] ⚠️ เอกสารโครงการยังไม่ครบถ้วน: ' + projectCode,
            'เรียนคุณ ' + (proj['pi_name']||submittedBy) + '\n\nเอกสารโครงการ ' + projectCode + ' — ' + (proj['title_th']||'') +
            '\nยังไม่ครบถ้วน กรุณาตรวจสอบและส่งเอกสารเพิ่มเติม\n' +
            (note ? '\nเอกสารที่ขาด / หมายเหตุ:\n' + note : '') +
            '\n\nกรุณาเข้าสู่ระบบ EC Online Submission เพื่อแนบเอกสารเพิ่มเติม');
        }
      }
    }
    writeLog('VERIFY_DOCS', projectCode + ' → ' + decision + (note?' | '+note:''), staffUsername);
    return { success: true, newStatus: newStatus };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function ensureProjectAssignmentsSheet_() {
  var ss = SpreadsheetApp.openById(getSheetId_());
  var sheet = ss.getSheetByName('ProjectAssignments');
  if (!sheet) {
    sheet = ss.insertSheet('ProjectAssignments');
    var headers = ['assignment_id','project_code','project_type','committee_username',
                   'committee_name','assigned_date','acceptance_status','eval_status','deadline'];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setBackground('#2d8a52').setFontColor('#fff').setFontWeight('bold');
  }
  return sheet;
}
function getAssignedProjects(sessionId, committeeUsername) {
  try {
    var auth = requireValidSession_(sessionId, committeeUsername);
    if (!auth.ok) return { success: false, message: auth.message, data: [] };
    var trimUser = auth.username.toLowerCase();
    if (!trimUser) return { success: false, message: 'ไม่ระบุชื่อผู้ใช้', data: [] };
    var assignSheet = getSheet('ProjectAssignments');
    if (!assignSheet) return { success: false, message: 'ไม่พบ sheet ProjectAssignments', data: [] };
    var assignData = assignSheet.getDataRange().getValues();
    if (assignData.length < 2) return { success: true, data: [] };
    var aH = {};
    assignData[0].forEach(function(h, i) { aH[String(h || '').trim().toLowerCase()] = i; });
    if (aH['committee_username'] === undefined) {
      return { success: false, message: 'ไม่พบคอลัมน์ committee_username ใน sheet', data: [] };
    }
    var projMap = {};
    var projSheet = getSheet('Projects');
    if (projSheet) {
      var projData = projSheet.getDataRange().getValues();
      if (projData.length >= 2) {
        var pH = {};
        projData[0].forEach(function(h, i) { pH[String(h || '').trim().toLowerCase()] = i; });
        for (var p = 1; p < projData.length; p++) {
          var pr = projData[p];
          var pc = String(pr[pH['project_code']] || '').trim();
          if (pc) {
            var skipReeval = String(pr[pH['skip_committee_reeval']] || '').toLowerCase() === 'true';
            var projectStatus = String(pr[pH['status']] || '').trim();
            projMap[pc] = {
              title_th:     String(pr[pH['title_th']]     || '').trim(),
              project_type: String(pr[pH['project_type']] || '').trim(),
              pi_name:      String(pr[pH['pi_name']]      || '').trim(),
              skip_committee_reeval: skipReeval,
              status: projectStatus
            };
          }
        }
      }
    }
    var result = [];
    for (var i = 1; i < assignData.length; i++) {
      var row = assignData[i];
      if (!row[aH['project_code']]) continue;
      var rowUser = String(row[aH['committee_username']] || '').trim().toLowerCase();
      if (rowUser !== trimUser) continue;
      var code   = String(row[aH['project_code']]      || '').trim();
      var proj   = projMap[code] || {};
      if (proj.skip_committee_reeval) continue;
      var finalStatuses = [
        'รับรอง',
        'ไม่รับรอง',
        'ปิดโครงการแล้ว',
        'รับรอง (แอดมินตรวจ)',
        'รับรองแล้ว (แก้ไขแล้ว)',
        'รับรอง (ข้ามกรรมการ)'
      ];
      var currentProjStatus = proj.status || '';
      var isFinal = finalStatuses.some(function(s) {
        return currentProjStatus.indexOf(s) !== -1;
      });
      if (isFinal) continue;
      var accSt  = String(row[aH['acceptance_status']] || '').trim() || 'รอตอบรับ';
      var evalSt = String(row[aH['eval_status']]       || '').trim() || 'รอประเมิน';
      var pType  = String(row[aH['project_type']]      || proj.project_type || 'new').trim();
      result.push({
        project_code:      code,
        title_th:          proj.title_th  || code,
        project_type:      pType,
        pi_name:           proj.pi_name   || '',
        acceptance_status: accSt,
        eval_status:       evalSt,
        deadline:          String(row[aH['deadline']]      || '').trim(),
        assigned_date:     String(row[aH['assigned_date']] || '').trim(),
        assignment_id:     String(row[aH['assignment_id']] || '').trim()
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
function getProjectAssignments(projectCode) {
  try {
    var rows = sheetToObjects('ProjectAssignments').filter(function(a){
      return String(a['project_code']||'').trim() === String(projectCode||'').trim();
    });
    return { success: true, data: rows };
  } catch (e) { return { success: false, data: [] }; }
}
function getCommitteeUsers(callerUsername) {
  requireRole_(callerUsername, ['admin', 'staff']);
  try {
    var committees = sheetToObjects('Users').filter(function(u) {
      var roles = String(u['roles'] || '');
      var status = String(u['status'] || 'active').toLowerCase();
      return roles.indexOf('committee') !== -1 && status === 'active';
    });
    return {
      success: true,
      data: committees.map(function(u){
        return {
          username:   String(u['username']  || ''),
          name:       String(u['name']      || u['username'] || ''),
          email:      String(u['email']     || ''),
          department: String(u['department']|| ''),
          position:   String(u['position']  || '')
        };
      })
    };
  } catch (e) {
    return { success: false, data: [] };
  }
}
function assignProjectWithEmail(payload, staffUsername) {
  try {
    if (!payload || !payload.projectCode) return { success: false, message: 'ไม่ระบุรหัสโครงการ' };
    if (!payload.committeeUsername) return { success: false, message: 'ไม่ระบุกรรมการ' };
    var projects = sheetToObjects('Projects');
    var proj = projects.find(function(p) {
      return String(p['project_code']||'').trim() === String(payload.projectCode).trim();
    });
    if (!proj) return { success: false, message: 'ไม่พบโครงการ "' + payload.projectCode + '"' };
    var docVerified = String(proj['doc_verified']||'').trim();
    var currentStatus = String(proj['status']||'').trim();
    if (docVerified !== 'approved' && currentStatus !== 'ผ่านตรวจเอกสาร') {
      return {
        success: false,
        message: 'โครงการนี้ยังไม่ผ่านการตรวจสอบเอกสาร\nสถานะปัจจุบัน: ' + currentStatus + '\nกรุณาตรวจสอบเอกสารก่อน'
      };
    }
    var existingAssign = sheetToObjects('ProjectAssignments').filter(function(a) {
      return String(a['project_code']||'').trim() === String(payload.projectCode).trim() &&
             String(a['committee_username']||'').trim() === String(payload.committeeUsername).trim();
    });
    if (existingAssign.length > 0) {
      return { success: false, message: 'กรรมการ "' + payload.committeeUsername + '" ได้รับมอบหมายโครงการนี้แล้ว' };
    }
    var allUsers = sheetToObjects('Users');
    var committee = allUsers.find(function(u) {
      return String(u['username']||'').trim() === String(payload.committeeUsername).trim();
    });
    if (!committee) return { success: false, message: 'ไม่พบกรรมการ "' + payload.committeeUsername + '"' };
    var projectType  = proj ? String(proj['project_type']||'new') : (payload.projectType||'new');
    var projectTitle = payload.projectTitle || (proj ? String(proj['title_th']||'') : payload.projectCode);
    var ts = new Date();
    var assignId = 'ASSIGN-' + ts.getTime();
    ensureProjectAssignmentsSheet_();
    var appendOk = appendRow('ProjectAssignments', {
      assignment_id:      assignId,
      project_code:       payload.projectCode,
      project_type:       projectType,
      committee_username: String(committee['username']||payload.committeeUsername).trim(),
      committee_name:     String(committee['name']||payload.committeeName||payload.committeeUsername),
      assigned_date:      ts,
      acceptance_status:  'รอตอบรับ',
      eval_status:        'รอประเมิน',
      deadline:           payload.deadline||''
    });
    if (!appendOk) return { success: false, message: 'ไม่สามารถบันทึกการมอบหมายได้' };
    updateRowByField('Projects','project_code',payload.projectCode,{status:'รอตอบรับ',updated_date:ts});
    var emailSent   = false;
    var committeeEmail = String(committee['email']||'').trim();
    if (committeeEmail && validateEmail_(committeeEmail)) {
      var projectFiles = getFileBlobsForProject_(payload.projectCode);
      var deadlineText = payload.deadline ? '\nกำหนดส่งผล: ' + payload.deadline : '';
      var noteText = payload.note
      ? '\nหมายเหตุ: ' + sanitizeEmailContent_(payload.note, 500)
      : '';
      var fileCountText = projectFiles.length > 0 ? '\n\n📎 เอกสารโครงการแนบมาด้วย (' + projectFiles.length + ' ไฟล์)' : '';
      var emailBody =
        'คุณ ' + String(committee['name']||payload.committeeUsername) + '\n\n' +
        'ท่านได้รับมอบหมายให้พิจารณาโครงการวิจัย:\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        'รหัสโครงการ : ' + payload.projectCode + '\n' +
        'ชื่อโครงการ  : ' + projectTitle + '\n' +
        'ประเภท      : ' + (projectType === 'new' ? 'โครงการใหม่' : 'โครงการต่อเนื่อง') +
        deadlineText + noteText + '\n' +
        'มอบหมายโดย : ' + (staffUsername||'เจ้าหน้าที่') + '\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━' +
        fileCountText + '\n\n' +
        'กรุณาเข้าสู่ระบบ EC Online Submission เพื่อตอบรับและส่งแบบประเมิน';
      emailSent = sendEmailWithAttachments_(committeeEmail,
        '[EC Sansai] มอบหมายโครงการ: ' + payload.projectCode, emailBody, projectFiles);
    }
    addNotification_('committee', String(committee['username']||'').trim(), payload.projectCode,
      '📋 ท่านได้รับมอบหมายให้พิจารณาโครงการ ' + payload.projectCode + ' — ' + projectTitle);
    writeLog('ASSIGN_PROJECT', payload.projectCode + ' → ' + payload.committeeUsername +
      ' | email:' + (emailSent ? 'sent' : 'skip') + ' | files:' + (emailSent ? 'attached' : '0'), staffUsername);
    return { success: true, emailSent: emailSent, assignId: assignId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function getFileBlobsForProject_(projectCode) {
  var blobs = [];
  try {
    var uploadedFiles = sheetToObjects('UploadedFiles').filter(function(f) {
      return String(f['project_code']||'').trim() === String(projectCode).trim() &&
             String(f['status']||'active') !== 'deleted';
    });
    uploadedFiles.forEach(function(fileRec) {
      var driveId = String(fileRec['drive_file_id']||fileRec['file_id']||'').trim();
      if (!driveId) return;
      try {
        var driveFile = DriveApp.getFileById(driveId);
        blobs.push(driveFile.getBlob());
      } catch (e) {
        Logger.log('getFileBlobsForProject_: cannot get file ' + driveId + ': ' + e.message);
      }
    });
  } catch (e) {
    Logger.log('getFileBlobsForProject_ error: ' + e.message);
  }
  return blobs;
}
function assignProjectMultiCommittee(payload, staffUsername) {
  try {
    if (!payload || !payload.projectCode || !payload.committeeList || !payload.committeeList.length) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
    }
    var proj = sheetToObjects('Projects').find(function(p) {
      return String(p['project_code']||'').trim() === String(payload.projectCode).trim();
    });
    if (!proj) return { success: false, message: 'ไม่พบโครงการ' };
    var docVerified = String(proj['doc_verified']||'').trim();
    var currentStatus = String(proj['status']||'').trim();
    if (docVerified !== 'approved' && currentStatus !== 'ผ่านตรวจเอกสาร') {
      return {
        success: false,
        message: 'โครงการยังไม่ผ่านการตรวจสอบเอกสาร (สถานะ: ' + currentStatus + ')\nกรุณากด "ตรวจเอกสาร" ก่อนทำการมอบหมาย'
      };
    }
    var results = [];
    payload.committeeList.forEach(function(comm) {
      var r = assignProjectWithEmail({
        projectCode:       payload.projectCode,
        projectTitle:      payload.projectTitle || payload.projectCode,
        committeeUsername: comm.username,
        committeeName:     comm.name || comm.username,
        deadline:          payload.deadline || '',
        note:              payload.note || ''
      }, staffUsername);
      results.push({
        username:  comm.username,
        name:      comm.name || comm.username,
        success:   r.success,
        emailSent: r.emailSent || false,
        message:   r.message || ''
      });
    });
    var successCount = results.filter(function(r){ return r.success; }).length;
    var emailCount   = results.filter(function(r){ return r.emailSent; }).length;
    return { success: successCount > 0, assignCount: successCount, emailCount: emailCount, results: results };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function submitRevision(payload) {
  try {
    if (!payload || !payload.projectCode || !payload.callerUsername) {
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    }
    var projectCode    = String(payload.projectCode).trim();
    var callerUsername = String(payload.callerUsername).trim();
    try {
      assertOwnsProject_(projectCode, callerUsername);
    } catch(e) {
      writeLog('SECURITY', 'submitRevision unauthorized: ' + callerUsername + ' on ' + projectCode, callerUsername);
      return { success: false, message: e.message };
    }
    var proj = sheetToObjects('Projects').find(function(p) {
      return String(p['project_code'] || '').trim() === projectCode;
    });
    if (!proj) return { success: false, message: 'ไม่พบโครงการ' };
    var currentStatus = String(proj['status'] || '');
    if (currentStatus !== 'รับรองมีข้อแก้ไข') {
      return { success: false, message: 'โครงการนี้ไม่อยู่ในสถานะที่สามารถยื่นแก้ไขได้ (สถานะปัจจุบัน: ' + currentStatus + ')' };
    }
    var ts           = new Date();
    var revRound     = parseInt(String(proj['revision_round'] || '1')) || 1;
    var skipCommittee = String(proj['skip_committee_reeval'] || '').toLowerCase() === 'true';
    var newProjectCode = generateRevisionProjectCode_(projectCode, revRound);
    var revId        = 'REV-' + newProjectCode + '-' + ts.getTime();
    appendRow('ProjectRevisions', {
      revision_id:    revId,
      project_code:   newProjectCode,
      original_code:  projectCode,
      round:          revRound,
      submitted_by:   callerUsername,
      submitted_date: ts,
      revision_note:  payload.revisionNote || '',
      file_ids:       (payload.fileIds || []).join(','),
      status:         skipCommittee ? 'รอตรวจโดยแอดมิน' : 'รอตรวจสอบ'
    });
    if (payload.fileIds && payload.fileIds.length) {
      updateUploadedFilesProjectCode_(payload.fileIds, newProjectCode);
    }
    appendRow('Projects', {
      project_code:   newProjectCode,
      title_th:       String(proj['title_th']||''),
      title_en:       String(proj['title_en']||''),
      project_type:   String(proj['project_type']||'new'),
      pi_name:        String(proj['pi_name']||''),
      pi_dept:        String(proj['pi_dept']||''),
      pi_phone:       String(proj['pi_phone']||''),
      pi_email:       String(proj['pi_email']||''),
      objectives:     String(proj['objectives']||''),
      methodology:    String(proj['methodology']||''),
      expected_benefits: String(proj['expected_benefits']||''),
      budget:         String(proj['budget']||'0'),
      duration:       String(proj['duration']||''),
      status:         skipCommittee ? 'รอตรวจโดยแอดมิน R' + revRound : 'รอประเมินแก้ไข R' + revRound,
      submitted_by:   callerUsername,
      created_date:   ts,
      updated_date:   ts,
      parent_code:    projectCode,
      doc_verified:   'approved',
      doc_verified_by: 'auto',
      doc_verified_date: ts,
      revision_round: revRound,
      skip_committee_reeval: skipCommittee ? 'true' : 'false'
    });
    updateRowByField('Projects', 'project_code', projectCode, {
      status: 'ส่งแก้ไขแล้ว (ดู ' + newProjectCode + ')',
      updated_date: ts
    });
    var emailCount  = 0;
    var projectTitle = String(proj['title_th']||projectCode);
    var revisionFiles = payload.fileIds && payload.fileIds.length
      ? getFileBlobsForEmail_(payload.fileIds) : [];
    if (skipCommittee) {
      notifyStaff_(
        '[EC Sansai] 🔔 ผู้วิจัยส่งแก้ไขครั้งที่ ' + revRound + ' (รอตรวจโดยแอดมิน): ' + newProjectCode,
        '⚠️ โครงการนี้ข้ามการประเมินโดยกรรมการ — ต้องตรวจโดยแอดมินโดยตรง\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        'รหัสโครงการใหม่ : ' + newProjectCode + '\n' +
        'รหัสเดิม         : ' + projectCode + '\n' +
        'ชื่อโครงการ       : ' + projectTitle + '\n' +
        'แก้ไขครั้งที่      : ' + revRound + '\n' +
        'ผู้วิจัย          : ' + callerUsername + '\n' +
        'หมายเหตุผู้วิจัย  : ' + (payload.revisionNote || '-') + '\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
        '📌 กรุณาเข้าระบบเพื่อตรวจสอบและแจ้งผลต่อผู้วิจัย'
      );
      addNotification_('staff', '', newProjectCode,
        '🔔 ผู้วิจัยส่งแก้ไขครั้งที่ ' + revRound + ' (รอตรวจโดยแอดมิน) — ' + newProjectCode);
      writeLog('SUBMIT_REVISION_SKIP_COMMITTEE',
        newProjectCode + '|original:' + projectCode + '|round:' + revRound + '|skip_committee:true',
        callerUsername);
      return {
        success: true, revisionId: revId, round: revRound, emailCount: 0,
        newProjectCode: newProjectCode, skippedCommittee: true,
        message: 'ส่งแก้ไขสำเร็จ — รอตรวจโดยแอดมินโดยตรง (ข้ามกรรมการ)'
      };
    } else {
      var assignSheet = getSheet('ProjectAssignments');
      var assignData  = assignSheet ? assignSheet.getDataRange().getValues() : [];
      var assignedCommittees = [];
      if (assignData.length >= 2) {
        var aH = {};
        assignData[0].forEach(function(h, i) { aH[String(h||'').trim().toLowerCase()] = i; });
        for (var i = 1; i < assignData.length; i++) {
          var row = assignData[i];
          if (!row[aH['project_code']]) continue;
          if (String(row[aH['project_code']]||'').trim() !== projectCode) continue;
          var commUser  = String(row[aH['committee_username']]||'').trim();
          var commName  = String(row[aH['committee_name']]||'').trim();
          var accStatus = String(row[aH['acceptance_status']]||'').trim();
          appendRow('ProjectAssignments', {
            assignment_id:      'ASSIGN-REV-' + ts.getTime() + '-' + commUser,
            project_code:       newProjectCode,
            project_type:       String(row[aH['project_type']]||'new'),
            committee_username: commUser,
            committee_name:     commName,
            assigned_date:      ts,
            acceptance_status:  'รอตอบรับ',
            eval_status:        'รอประเมิน',
            deadline:           String(row[aH['deadline']]||'')
          });
          var commUser_ = sheetToObjects('Users').find(function(u) {
            return String(u['username']||'').trim() === commUser;
          });
          assignedCommittees.push({
            username: commUser,
            name: commName,
            email: commUser_ ? String(commUser_['email']||'').trim() : ''
          });
        }
      }
      assignedCommittees.forEach(function(comm) {
        if (!comm.email || !validateEmail_(comm.email)) return;
        var fileLinks = payload.fileIds && payload.fileIds.length
          ? '\n\n📎 ไฟล์เอกสารแก้ไขแนบมาด้วย (' + payload.fileIds.length + ' ไฟล์)' : '';
        var emailBody =
          'คุณ ' + comm.name + '\n\n' +
          '✏️ ผู้วิจัยส่งเอกสารแก้ไขครั้งที่ ' + revRound + ' สำหรับโครงการ:\n\n' +
          '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
          'รหัสโครงการใหม่ : ' + newProjectCode + '\n' +
          'รหัสเดิม         : ' + projectCode + '\n' +
          'ชื่อโครงการ       : ' + projectTitle + '\n' +
          'แก้ไขครั้งที่      : ' + revRound + '\n' +
          'หมายเหตุผู้วิจัย  : ' + (payload.revisionNote || '-') + '\n' +
          '━━━━━━━━━━━━━━━━━━━━━━━━━━━━' + fileLinks + '\n\n' +
          'กรุณาเข้าสู่ระบบ EC Online Submission เพื่อตอบรับและประเมินเอกสารที่แก้ไข';
        var sent = sendEmailWithAttachments_(
          comm.email,
          '[EC Sansai] ✏️ ผู้วิจัยส่งแก้ไขครั้งที่ ' + revRound + ': ' + newProjectCode,
          emailBody, revisionFiles
        );
        if (sent) emailCount++;
        addNotification_('committee', comm.username, newProjectCode,
          '✏️ ผู้วิจัยส่งแก้ไขครั้งที่ ' + revRound + ' — ' + newProjectCode + ' รอตอบรับ/ประเมิน');
      });
      notifyStaff_(
        '[EC Sansai] ผู้วิจัยส่งแก้ไขครั้งที่ ' + revRound + ': ' + newProjectCode,
        'โครงการ ' + newProjectCode + ' (แก้ไขจาก ' + projectCode + ') — ' + projectTitle +
        '\nผู้วิจัย: ' + callerUsername + ' ส่งแก้ไขครั้งที่ ' + revRound +
        '\nหมายเหตุ: ' + (payload.revisionNote || '-') +
        '\n\n📌 ระบบได้มอบหมายให้กรรมการ ' + assignedCommittees.length + ' ท่านเรียบร้อยแล้ว'
      );
      addNotification_('staff', '', newProjectCode,
        '📝 ผู้วิจัยส่งแก้ไขครั้งที่ ' + revRound + ' — ' + newProjectCode + ' (มอบหมายกรรมการ ' + assignedCommittees.length + ' ท่าน)');
      writeLog('SUBMIT_REVISION',
        newProjectCode + '|original:' + projectCode + '|round:' + revRound + '|committees:' + assignedCommittees.length,
        callerUsername);
      return {
        success: true,
        revisionId: revId,
        round: revRound,
        emailCount: emailCount,
        newProjectCode: newProjectCode,
        skippedCommittee: false,
        committeesCount: assignedCommittees.length
      };
    }
  } catch (e) {
    Logger.log('submitRevision ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}
function getFileBlobsForEmail_(fileIds) {
  var blobs = [];
  try {
    if (!fileIds || !fileIds.length) return blobs;
    var files = sheetToObjects('UploadedFiles');
    fileIds.forEach(function(fid) {
      var rec = files.find(function(f) {
        return String(f['file_id'] || f['drive_file_id'] || '') === String(fid).trim();
      });
      if (!rec) return;
      var driveId = String(rec['drive_file_id'] || rec['file_id'] || '').trim();
      if (driveId) {
        try {
          var file = DriveApp.getFileById(driveId);
          blobs.push(file.getBlob());
        } catch (e) {}
      }
    });
  } catch (e) {}
  return blobs;
}
function submitAcceptance(payload) {
  try {
    if (!payload || !payload.projectCode || !payload.decision)
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    var VALID_ROLES = ['1st', '2nd', '3rd'];
    var reviewerRole = String(payload.reviewerRole || '1st').trim();
    if (VALID_ROLES.indexOf(reviewerRole) === -1) reviewerRole = '1st';
    var ts = new Date();
    appendRow('CommitteeReviews', {
      timestamp:          ts,
      project_code:       payload.projectCode,
      committee_username: payload.committeeUsername,
      committee_name:     payload.committeeName,
      decision:           payload.decision,
      reviewer_role:      reviewerRole,
      note:               payload.note || '',
      status:             payload.decision === 'accept' ? 'ตอบรับแล้ว' : 'ปฏิเสธ',
      eval_status:        payload.decision === 'accept' ? 'รอประเมิน' : '-'
    });
    updateRowByField2('ProjectAssignments', 'project_code', payload.projectCode,
      'committee_username', payload.committeeUsername,
      { acceptance_status: payload.decision === 'accept' ? 'ตอบรับ' : 'ปฏิเสธ' });
    var newProjStatus;
    if (payload.decision === 'accept') {
      if (reviewerRole === '3rd') {
        newProjStatus = 'รอประเมิน (ครบ 3 คน)';
      } else {
        newProjStatus = 'รอประเมิน';
      }
    } else {
      newProjStatus = 'หาผู้ทดแทน';
    }
    updateRowByField('Projects', 'project_code', payload.projectCode, {
      status: newProjStatus, updated_date: ts
    });
    var roleLabel = reviewerRole + ' Reviewer';
    addNotification_('staff', '', payload.projectCode,
      'กรรมการ ' + payload.committeeName + ' [' + roleLabel + '] ' +
      (payload.decision === 'accept' ? 'ตอบรับ' : 'ปฏิเสธ') + ' โครงการ ' + payload.projectCode);
    notifyStaff_(
      '[EC Sansai] กรรมการ' + (payload.decision === 'accept' ? 'ตอบรับ' : 'ปฏิเสธ') + ': ' + payload.projectCode,
      'กรรมการ ' + payload.committeeName + ' [' + roleLabel + '] ได้' +
      (payload.decision === 'accept' ? 'ตอบรับ' : 'ปฏิเสธ') +
      'การพิจารณาโครงการ ' + payload.projectCode
    );
    writeLog('ACCEPTANCE', payload.decision + '|' + reviewerRole + '|' + payload.projectCode, payload.committeeUsername);
    return { success: true, reviewerRole: reviewerRole };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function submitEvaluation(payload) {
  try {
    if (!payload || !payload.projectCode || !payload.finalDecision)
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    var auth = requireValidSession_(payload.sessionId, payload.committeeUsername);
    if (!auth.ok) return { success: false, message: auth.message };
    var safeCommittee = auth.username;
    var ts = new Date();
    appendRow('CommitteeReviews', {
      timestamp:          ts,
      project_code:       payload.projectCode,
      committee_username: safeCommittee,
      committee_name:     payload.committeeName,
      project_type:       payload.projectType || 'new',
      eval_answers:       JSON.stringify(payload.answers || {}),
      final_decision:     payload.finalDecision,
      comment:            payload.comment || '',
      eval_status:        'ส่งประเมินแล้ว',
      status:             'รอผลพิจารณา'
    });
    updateRowByField('Projects', 'project_code', payload.projectCode, {
      status: 'รอผลพิจารณา', updated_date: ts
    });
    updateRowByField2('ProjectAssignments','project_code',payload.projectCode,
      'committee_username', safeCommittee,
      { eval_status: 'ส่งประเมินแล้ว' });
    addNotification_('staff', '', payload.projectCode,
      'กรรมการ ' + payload.committeeName + ' ส่งแบบประเมิน ' + payload.projectCode +
      ' [' + payload.finalDecision + ']');
    notifyStaff_(
      '[EC Sansai] ส่งแบบประเมิน: ' + payload.projectCode,
      'กรรมการ ' + payload.committeeName + ' ส่งแบบประเมินโครงการ ' + payload.projectCode +
      '\nมติที่เสนอ: ' + payload.finalDecision
    );
    writeLog('EVALUATION', payload.finalDecision + '|' + payload.projectCode, safeCommittee);
    checkEvaluationComplete(payload.projectCode);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function checkEvaluationComplete(projectCode) {
  try {
    var assignSheet = getSheet('ProjectAssignments');
    if (!assignSheet) return { complete: false };
    var data = assignSheet.getDataRange().getValues();
    if (data.length < 2) return { complete: false };
    var headers = {};
    data[0].forEach(function(h, i) { headers[String(h || '').trim().toLowerCase()] = i; });
    var projectRows = [];
    for (var i = 1; i < data.length; i++) {
      var code = String(data[i][headers['project_code']] || '').trim();
      if (code === String(projectCode).trim()) projectRows.push(data[i]);
    }
    if (!projectRows.length) return { complete: false };
    var evalCol = headers['eval_status'];
    var accCol  = headers['acceptance_status'];
    var accepted = projectRows.filter(function(r) {
      var acc = String(r[accCol] || '').trim();
      return acc === 'ตอบรับ' || acc === 'auto';
    });
    var evaluated = accepted.filter(function(r) {
      return String(r[evalCol] || '').trim() === 'ส่งประเมินแล้ว';
    });
    var complete = accepted.length > 0 && evaluated.length === accepted.length;
    if (complete) {
      var proj = sheetToObjects('Projects').find(function(p) {
        return String(p['project_code'] || '').trim() === String(projectCode).trim();
      });
      var currentStatus = proj ? String(proj['status'] || '') : '';
      if (currentStatus !== 'รอแจ้งผล' && currentStatus !== 'รับรอง' && currentStatus !== 'ไม่รับรอง') {
        updateRowByField('Projects', 'project_code', projectCode, {
          status: 'รอแจ้งผล', updated_date: new Date()
        });
        notifyStaff_('[EC Sansai] กรรมการประเมินครบแล้ว: ' + projectCode,
          'โครงการ ' + projectCode + ' กรรมการทั้ง ' + evaluated.length + ' ท่านส่งประเมินครบแล้ว');
        addNotification_('staff', '', projectCode,
          '✅ กรรมการ ' + evaluated.length + ' ท่านส่งประเมินครบ — ' + projectCode + ' รอแจ้งผล');
        writeLog('EVAL_COMPLETE', projectCode + '|' + evaluated.length + ' evaluators done', 'system');
      }
    }
    return { complete: complete, total: projectRows.length, accepted: accepted.length, done: evaluated.length };
  } catch (e) {
    Logger.log('checkEvaluationComplete error: ' + e.message);
    return { complete: false };
  }
}
function getEvaluationHistory(sessionId, committeeUsername) {
  try {
    var auth = requireValidSession_(sessionId, committeeUsername);
    if (!auth.ok) return { success: false, data: [] };
    var safeUsername = auth.username;
    var rows = sheetToObjects('CommitteeReviews').filter(function(r){
      return String(r['committee_username']||'').trim() === safeUsername &&
             (String(r['eval_status']||'')==='ส่งประเมินแล้ว'||String(r['eval_status']||'')==='ตรวจสอบมติแล้ว');
    });
    return { success: true, data: rows };
  } catch (e) { return { success: false, data: [] }; }
}
function getCommitteeEvalSummary(projectCode) {
  try {
    var reviews = sheetToObjects('CommitteeReviews').filter(function(r) {
      return String(r['project_code'] || '').trim() === String(projectCode).trim() &&
             String(r['eval_status']  || '') === 'ส่งประเมินแล้ว';
    });
    var files = sheetToObjects('UploadedFiles').filter(function(f) {
      return String(f['project_code'] || '').trim() === String(projectCode).trim() &&
             String(f['status'] || 'active') !== 'deleted';
    });
    return {
      success:  true,
      reviews:  reviews.map(function(r) {
        return {
          committee_name: String(r['committee_name']  || ''),
          final_decision: String(r['final_decision']  || ''),
          comment:        String(r['comment']         || ''),
          eval_status:    String(r['eval_status']     || ''),
          timestamp:      r['timestamp'] ? String(r['timestamp']) : ''
        };
      }),
      files: files.map(function(f) {
        return {
          file_name:    String(f['file_name']    || ''),
          file_url:     String(f['file_url']     || ''),
          download_url: String(f['download_url'] || ''),
          file_type:    String(f['file_type']    || ''),
          uploaded_by:  String(f['uploaded_by']  || '')
        };
      })
    };
  } catch (e) {
    return { success: false, message: e.message, reviews: [], files: [] };
  }
}
function getProjectsReadyToNotify() {
  try {
    var projects = sheetToObjects('Projects').filter(function(p) {
      return String(p['status'] || '') === 'รอแจ้งผล';
    });
    return {
      success: true,
      data: projects.map(function(p) {
        return {
          project_code: String(p['project_code']  || ''),
          title_th:     String(p['title_th']       || ''),
          pi_name:      String(p['pi_name']        || ''),
          pi_email:     String(p['pi_email']       || ''),
          submitted_by: String(p['submitted_by']   || ''),
          status:       String(p['status']         || ''),
          created_date: String(p['created_date']   || '')
        };
      })
    };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
function sendResultToResearcher(payload) {
  try {
    if (!payload || !payload.projectCode || !payload.finalResult) {
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    }
    var projectCode    = String(payload.projectCode).trim();
    var finalResult    = String(payload.finalResult).trim();
    var certFileId     = String(payload.certFileId     || '');
    var revisionFileId = String(payload.revisionFileId || '');
    var rejectionFileId = String(payload.rejectionFileId || '');
    var otherFileIds   = payload.otherFileIds || [];
    var skipCommittee  = payload.skipCommitteeReeval === true;
    var staffUsername  = String(payload.staffUsername  || 'staff');
    var ts = new Date();
    var proj = sheetToObjects('Projects').find(function(p) {
      return String(p['project_code'] || '').trim() === projectCode;
    });
    if (!proj) return { success: false, message: 'ไม่พบโครงการ' };
    var resultNote   = sanitizeEmailContent_(payload.resultNote, 1000);
    var revDeadline  = sanitizeEmailField_(payload.revisionDeadline || '');
    var piName       = sanitizeEmailField_(proj['pi_name']  || '');
    var projectTitle = sanitizeEmailField_(proj['title_th'] || projectCode);
    var piEmail     = String(proj['pi_email']     || '').trim();
    var submittedBy = String(proj['submitted_by'] || '').trim();
    var researcherEmail = piEmail;
    if (submittedBy) {
      var resUser = sheetToObjects('Users').find(function(u) {
        return String(u['username'] || '').trim() === submittedBy;
      });
      if (resUser && resUser['email']) researcherEmail = String(resUser['email']).trim();
    }
    var newStatus = finalResult === 'รับรอง'             ? 'รับรอง'
                  : finalResult === 'รับรองมีข้อแก้ไข'  ? 'รับรองมีข้อแก้ไข'
                  : finalResult === 'ไม่รับรอง'          ? 'ไม่รับรอง'
                  : finalResult;
    if (skipCommittee && projectCode.indexOf('_fix') !== -1) {
      newStatus = finalResult === 'รับรอง' ? 'รับรอง (แอดมินตรวจ)' : newStatus;
    }
    var projUpdates = {
      status:                newStatus,
      final_result:          finalResult,
      result_note:           resultNote,
      result_sent_date:      ts,
      updated_date:          ts,
      cert_file_id:          certFileId,
      revision_file_id:      revisionFileId,
      skip_committee_reeval: skipCommittee ? 'true' : 'false'
    };
    if (revDeadline && finalResult === 'รับรองมีข้อแก้ไข') {
      projUpdates['revision_deadline'] = revDeadline;
      projUpdates['revision_round']    = 1;
    }
    updateRowByField('Projects', 'project_code', projectCode, projUpdates);
    appendRow('ProjectResults', {
      timestamp:             ts,
      project_code:          projectCode,
      final_result:          finalResult,
      result_note:           resultNote,
      revision_deadline:     revDeadline,
      sent_by:               staffUsername,
      researcher_email:      researcherEmail,
      cert_file_id:          certFileId,
      revision_file_id:      revisionFileId,
      skip_committee_reeval: skipCommittee ? 'true' : 'false'
    });
    var notificationMsg = '📋 ผลการพิจารณาโครงการ ' + projectCode + ': ' + finalResult;
    if (revDeadline) notificationMsg += ' | กำหนดแก้ไขภายใน: ' + revDeadline;
    if (certFileId) notificationMsg += ' | 📜 หนังสือรับรองพร้อมแล้ว';
    if (skipCommittee) notificationMsg += ' (แอดมินตรวจสอบเอง)';
    addNotification_('researcher', submittedBy, projectCode, notificationMsg);
    var emailSent = false;
    if (researcherEmail && validateEmail_(researcherEmail)) {
      var resultIcon = finalResult === 'รับรอง'            ? '✅'
                     : finalResult === 'รับรองมีข้อแก้ไข'  ? '✏️'
                     : '❌';
      var adminNote = skipCommittee ? '\n\n👤 (ตรวจสอบโดยแอดมิน)' : '';
      var revisionSection = (finalResult === 'รับรองมีข้อแก้ไข' && revDeadline)
        ? '\n\n📌 ขั้นตอนถัดไป:\n   แก้ไขโครงการตามข้อเสนอแนะและส่งกลับภายในวันที่ ' + revDeadline
        : '';
      var certSection = '';
      if (certFileId) {
        certSection = '\n\n📜 หนังสือรับรอง:\n   ดาวน์โหลด: https://drive.google.com/uc?export=download&id=' + certFileId;
      }
      var emailBody =
        'คุณ ' + piName + '\n\n' +
        resultIcon + ' ผลการพิจารณาโครงการวิจัยของท่าน' + adminNote + '\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        'รหัสโครงการ : ' + projectCode + '\n' +
        'ชื่อโครงการ  : ' + projectTitle + '\n' +
        '📋 ผลการพิจารณา: ' + resultIcon + ' ' + finalResult + '\n' +
        (resultNote ? '📝 ข้อเสนอแนะ:\n' + resultNote + '\n' : '') +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━' +
        revisionSection + certSection;
      var attachmentBlobs = [];
      var attachedFiles = [];
      if (certFileId) {
        try {
          var certFile = DriveApp.getFileById(certFileId);
          attachmentBlobs.push(certFile.getBlob());
          attachedFiles.push('📜 หนังสือรับรอง: ' + certFile.getName());
        } catch (e) {
          Logger.log('Cannot attach cert file: ' + e.message);
        }
      }
      if (revisionFileId && finalResult === 'รับรองมีข้อแก้ไข') {
        try {
          var revisionFile = DriveApp.getFileById(revisionFileId);
          attachmentBlobs.push(revisionFile.getBlob());
          attachedFiles.push('✏️ ข้อเสนอแนะการแก้ไข: ' + revisionFile.getName());
        } catch (e) {
          Logger.log('Cannot attach revision file: ' + e.message);
        }
      }
      if (rejectionFileId && finalResult === 'ไม่รับรอง') {
        try {
          var rejectionFile = DriveApp.getFileById(rejectionFileId);
          attachmentBlobs.push(rejectionFile.getBlob());
          attachedFiles.push('❌ เหตุผลการไม่รับรอง: ' + rejectionFile.getName());
        } catch (e) {
          Logger.log('Cannot attach rejection file: ' + e.message);
        }
      }
      if (otherFileIds && otherFileIds.length) {
        otherFileIds.forEach(function(fileId) {
          try {
            var otherFile = DriveApp.getFileById(fileId);
            attachmentBlobs.push(otherFile.getBlob());
            attachedFiles.push('📎 เอกสารเพิ่มเติม: ' + otherFile.getName());
          } catch (e) {
            Logger.log('Cannot attach other file: ' + e.message);
          }
        });
      }
      if (attachedFiles.length > 0) {
        emailBody += '\n\n📎 ไฟล์ที่แนบมาด้วย:\n';
        attachedFiles.forEach(function(fileInfo) {
          emailBody += '   ' + fileInfo + '\n';
        });
      }
      emailBody += '\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
                   'ระบบ EC Online Submission\n' +
                   'โรงพยาบาลสันทราย';
      emailSent = sendEmailWithAttachments_(
        researcherEmail,
        '[EC Sansai] ' + resultIcon + ' ผลการพิจารณา: ' + projectCode + ' — ' + finalResult,
        emailBody,
        attachmentBlobs
      );
    }
    writeLog('SEND_RESULT',
      projectCode + '→' + finalResult + '|email:' + (emailSent?'sent':'skip') +
      '|cert:' + (certFileId?'yes':'no') +
      '|revision:' + (revisionFileId?'yes':'no') +
      '|rejection:' + (rejectionFileId?'yes':'no') +
      '|other:' + otherFileIds.length +
      '|skipCommittee:' + skipCommittee,
      staffUsername);
    return {
      success: true,
      emailSent: emailSent,
      newStatus: newStatus,
      certUploaded: !!certFileId,
      filesAttached: attachmentBlobs.length
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function addNotification_(toRole, toUsername, projectCode, message) {
  try {
    var ts = new Date();
    if (toUsername) {
      appendRow('Notifications', {
        timestamp:    ts,
        to_username:  toUsername,
        to_role:      toRole || '',
        project_code: projectCode || '',
        message:      message || '',
        is_read:      'false'
      });
    }
    if (toRole && toRole !== 'researcher') {
      appendRow('Notifications', {
        timestamp:    ts,
        to_username:  '',
        to_role:      toRole,
        project_code: projectCode || '',
        message:      message || '',
        is_read:      'false'
      });
    }
    if (!toUsername && toRole === 'researcher') {
      appendRow('Notifications', {
        timestamp:    ts,
        to_username:  '',
        to_role:      'researcher',
        project_code: projectCode || '',
        message:      message || '',
        is_read:      'false'
      });
    }
  } catch (e) {
    Logger.log('addNotification_ error: ' + e.message);
  }
}
function getNotificationsForUser(sessionId, username, roles) {
  try {
    var auth = requireValidSession_(sessionId, username);
    if (!auth.ok) return [];
    var safeUsername = auth.username;
    var userRoles  = Array.isArray(roles) ? roles : (roles ? [roles] : []);
    var trimUser   = safeUsername.toLowerCase();
    var allNotifs  = sheetToObjects('Notifications');
    var seen = {};
    var filtered = [];
    allNotifs.forEach(function(r) {
      var toUser = String(r['to_username'] || '').trim().toLowerCase();
      var toRole = String(r['to_role']     || '').trim().toLowerCase();
      var msg    = String(r['message']     || '');
      var ts     = String(r['timestamp']   || '');
      var key    = ts + '|' + msg;
      if (seen[key]) return;
      var match = false;
      if (toUser && toUser === trimUser) match = true;
      if (!toUser && toRole) {
        if (toRole === 'all') match = true;
        userRoles.forEach(function(role) {
          if (toRole === role.toLowerCase()) match = true;
        });
      }
      if (match) {
        seen[key] = true;
        filtered.push(r);
      }
    });
    var sorted = filtered.slice(-100).reverse().slice(0, 60);
    return {
      success: true,
      data: sorted.map(function(r) {
        var isRead = String(r['is_read'] || 'false').toLowerCase();
        return {
          id:           String(r['_row'] || ''),
          timestamp:    r['timestamp'] ? String(r['timestamp']) : '',
          project_code: String(r['project_code'] || ''),
          message:      String(r['message']       || ''),
          is_read:      isRead === 'true' || isRead === '1' || isRead === 'yes',
          to_role:      String(r['to_role']       || ''),
          to_username:  String(r['to_username']   || '')
        };
      })
    };
  } catch (e) {
    return { success: true, data: [] };
  }
}
function markNotificationsRead(rowNumbers) {
  try {
    if (!rowNumbers || !rowNumbers.length) return { success: true };
    var sheet = getSheet('Notifications');
    if (!sheet) return { success: false };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var readCol = headers.map(function(h){ return String(h).trim().toLowerCase(); }).indexOf('is_read');
    if (readCol === -1) return { success: false };
    rowNumbers.forEach(function(rowNum) {
      sheet.getRange(rowNum, readCol + 1).setValue('true');
    });
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function getMeetings(sessionId, callerUsername) {
  var auth = requireValidSession_(sessionId, callerUsername);
  if (!auth.ok) return { success: false, data: [] };
  requireRole_(auth.username, ['admin', 'staff', 'committee']);
  try {
    return { success: true, data: sheetToObjects('Meetings') };
  } catch (e) { return { success: false, data: [] }; }
}
function getDashboardStats(sessionId, committeeUsername) {
  try {
    var auth = requireValidSession_(sessionId, committeeUsername);
    if (!auth.ok) return { success: false, pendingMeeting: 0 };
    var meetings = sheetToObjects('Meetings');
    var pendingMeeting = meetings.filter(function(m){
      return String(m['status']||'').indexOf('รอตรวจสอบ') !== -1;
    }).length;
    return { success: true, pendingMeeting: pendingMeeting };
  } catch (e) { return { success: true, pendingMeeting: 0 }; }
}
function getStaffDashboardStats(sessionId, callerUsername) {
  try {
    var auth = requireValidSession_(sessionId, callerUsername);
    if (!auth.ok) return { success: false, message: auth.message };
    var projects = sheetToObjects('Projects');
    var meetings = sheetToObjects('Meetings');
    var mainProjects = projects.filter(function(p) {
      var s = String(p['status']||'');
      return s !== 'ส่งแก้ไขแล้ว (ดู ' && s.indexOf('ส่งแก้ไขแล้ว') === -1;
    });
    return {
      success: true, total: mainProjects.length,
      pending:       projects.filter(function(p){ return String(p['status']||'')==='รอพิจารณา'; }).length,
      needVerify:    projects.filter(function(p){ var s=String(p['status']||''); return s==='รอพิจารณา'&&String(p['doc_verified']||'')!=='approved'; }).length,
      verified:      projects.filter(function(p){ return String(p['status']||'')==='ผ่านตรวจเอกสาร'; }).length,
      inReview:      projects.filter(function(p){ var s=String(p['status']||''); return s.indexOf('รอประเมิน')!==-1||s.indexOf('รอตอบรับ')!==-1; }).length,
      waitingResult: projects.filter(function(p){ return String(p['status']||'')==='รอผลพิจารณา'||String(p['status']||'')==='รอแจ้งผล'; }).length,
      approved:      projects.filter(function(p){ return String(p['status']||'')==='รับรอง'; }).length,
      pendingMeetings: meetings.filter(function(m){ return String(m['status']||'').indexOf('รอตรวจสอบ')!==-1; }).length
    };
  } catch (e) { return { success: false, message: e.message }; }
}
function saveMeeting(payload, staffUsername) {
  try {
    if (!payload||!payload.title||!payload.meetingDate) return { success: false, message: 'ข้อมูลไม่ครบ' };
    var ts = new Date();
    if (payload.meeting_id) {
      updateRowByField('Meetings','meeting_id',payload.meeting_id,{
        title:payload.title,meeting_date:payload.meetingDate,
        location:payload.location||'',status:payload.status||'รอตรวจสอบ',note:payload.note||''
      });
    } else {
      var mid = 'MTG-'+ts.getFullYear()+'-'+String(ts.getMonth()+1).padStart(2,'0')+'-'+String(ts.getTime()).slice(-4);
      appendRow('Meetings',{meeting_id:mid,title:payload.title,meeting_date:payload.meetingDate,
        status:'รอตรวจสอบ',location:payload.location||'',note:payload.note||''});
      payload.meeting_id = mid;
    }
    writeLog('SAVE_MEETING', payload.meeting_id, staffUsername);
    return { success: true, meetingId: payload.meeting_id };
  } catch (e) { return { success: false, message: e.message }; }
}
function submitMeetingReview(payload) {
  try {
    if (!payload||!payload.meetingId||!payload.agendaItems) return { success: false, message: 'ข้อมูลไม่ครบ' };
    var ts = new Date();
    payload.agendaItems.forEach(function(item){
      appendRow('CommitteeReviews',{
        timestamp:ts, project_code:item.projectCode||'',
        committee_username:payload.committeeUsername, committee_name:payload.committeeName,
        meeting_id:payload.meetingId, agenda_id:item.agendaId, agenda_title:item.agendaTitle,
        resolution:item.resolution, reviewer_no:payload.reviewerNo||1, eval_status:'ตรวจสอบมติแล้ว'
      });
    });
    updateRowByField('Meetings','meeting_id',payload.meetingId,{status:'ตรวจสอบแล้ว_'+(payload.reviewerNo||1)});
    if ((payload.reviewerNo||1) >= 3) {
      notifyStaff_('[EC Sansai] ตรวจสอบมติครบแล้ว','การประชุม '+payload.meetingId+' ผ่านการตรวจสอบมติครบ 3 ท่านแล้ว');
      updateRowByField('Meetings','meeting_id',payload.meetingId,{status:'เสร็จสิ้น'});
    }
    writeLog('MEETING_REVIEW','meetingId:'+payload.meetingId,payload.committeeUsername);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function getUploadedFiles(sessionId, projectCode, callerUsername) {
  try {
    var auth = requireValidSession_(sessionId, callerUsername);
    if (!auth.ok) return { success: false, message: auth.message, data: [] };
    var safeUsername = auth.username;
    if (!projectCode) return { success: false, data: [] };
    var user = requireRole_(safeUsername, ['admin', 'staff', 'committee', 'researcher']);
    var roles = String(user['roles'] || '');
    var isPrivileged = roles.indexOf('admin') !== -1 ||
                       roles.indexOf('staff') !== -1 ||
                       roles.indexOf('committee') !== -1;
    if (!isPrivileged) {
      var owned = sheetToObjects('Projects').find(function(p) {
        return String(p['project_code'] || '').trim() === String(projectCode).trim() &&
               String(p['submitted_by'] || '').trim() === safeUsername;
      });
      if (!owned) {
        var baseCode = String(projectCode).replace(/_fix\d+$/, '');
        owned = sheetToObjects('Projects').find(function(p) {
          return String(p['project_code'] || '').trim() === baseCode &&
                 String(p['submitted_by'] || '').trim() === safeUsername;
        });
      }
      if (!owned) return { success: false, message: 'ไม่มีสิทธิ์เข้าถึงไฟล์นี้', data: [] };
    }
    var codeStr = String(projectCode).trim();
    var uploadedFiles = sheetToObjects('UploadedFiles');
    Logger.log('Total uploaded files: ' + uploadedFiles.length);
    var files = uploadedFiles.filter(function(f) {
      var fc = String(f['project_code'] || '').trim();
      var status = String(f['status'] || 'active').toLowerCase();
      var isActive = status !== 'deleted';
      var isExact = fc === codeStr;
      var isFixChild = /^.+_fix\d+$/.test(fc) && fc.replace(/_fix\d+$/, '') === codeStr;
      var isRptChild = fc.indexOf(codeStr + '_rpt_') === 0;
      return (isExact || isFixChild || isRptChild) && isActive;
    });
    Logger.log('Filtered files for ' + projectCode + ': ' + files.length);
    if (files.length === 0) {
      var baseCode2 = codeStr.replace(/_fix\d+$/, '');
      if (baseCode2 !== codeStr) {
        files = uploadedFiles.filter(function(f) {
          var fc = String(f['project_code'] || '').trim();
          var status = String(f['status'] || 'active').toLowerCase();
          return (fc === baseCode2 || fc === codeStr) && status !== 'deleted';
        });
      }
    }
    var result = files.map(function(f) {
      var fileId = String(f['file_id'] || f['drive_file_id'] || '');
      var fileUrl = String(f['file_url'] || '');
      var downloadUrl = String(f['download_url'] || '');
      if (!downloadUrl && fileId) {
        downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
      }
      if (!fileUrl && fileId) {
        fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
      }
      return {
        file_id: fileId,
        project_code: String(f['project_code'] || ''),
        file_name: String(f['file_name'] || ''),
        file_url: fileUrl,
        download_url: downloadUrl,
        file_type: String(f['file_type'] || 'other'),
        mime_type: String(f['mime_type'] || ''),
        uploaded_by: String(f['uploaded_by'] || ''),
        uploaded_date: f['uploaded_date'] ? String(f['uploaded_date']) : '',
        file_size: f['file_size'] ? String(f['file_size']) : '',
        status: String(f['status'] || 'active')
      };
    });
    return { success: true, data: result };
  } catch (e) {
    Logger.log('getUploadedFiles error: ' + e.message);
    return { success: false, message: e.message, data: [] };
  }
}
function uploadFileWithType(base64Data, fileName, mimeType, projectCode, username, fileType) {
  try {
    if (!base64Data || !fileName) {
      throw new Error('ข้อมูลไฟล์ไม่ถูกต้อง');
    }
    fileName = sanitizeInput_(fileName);
    if (!fileName) throw new Error('ชื่อไฟล์ไม่ถูกต้อง');
    var estimatedSize = (base64Data.length * 3) / 4;
    var sizeInMB = estimatedSize / (1024 * 1024);
    if (sizeInMB > MAX_FILE_SIZE_MB) {
      return {
        success: false,
        message: 'ไฟล์มีขนาดใหญ่เกิน ' + MAX_FILE_SIZE_MB + ' MB (ขนาดจริง: ' +
                 sizeInMB.toFixed(2) + ' MB)'
      };
    }
    if (mimeType && ALLOWED_MIME_TYPES.indexOf(mimeType) === -1) {
      return {
        success: false,
        message: 'ประเภทไฟล์ไม่ได้รับอนุญาต: ' + mimeType
      };
    }
    if (projectCode) {
      projectCode = sanitizeInput_(projectCode);
    }
    if (mimeType && MAGIC_BYTES[mimeType]) {
      if (!verifyFileMagicBytes_(base64Data, mimeType)) {
        return {
          success: false,
          message: 'เนื้อหาไฟล์ไม่ตรงกับประเภทที่ระบุ (' + mimeType + ') — อาจเป็นไฟล์ที่เปลี่ยนนามสกุล'
        };
      }
    }
    var folder;
    try {
      folder = DriveApp.getFolderById(getDriveFolderId_());
    } catch (err) {
      throw new Error('ไม่สามารถเข้าถึงโฟลเดอร์: ' + err.message);
    }
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var fileUrl = file.getUrl();
    var downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
    appendRow('UploadedFiles', {
      file_id: fileId,
      project_code: projectCode || '',
      file_name: fileName,
      file_url: fileUrl,
      download_url: downloadUrl,
      drive_file_id: fileId,
      file_type: fileType || 'other',
      mime_type: mimeType || '',
      uploaded_by: username || '',
      uploaded_date: new Date(),
      file_size: Math.round(estimatedSize),
      is_required: false,
      status: 'active'
    });
    writeLog('UPLOAD_FILE', fileName + '|type:' + fileType + '|size:' + sizeInMB.toFixed(2) + 'MB|project:' + (projectCode || 'TEMP'), username);
    return {
      success: true,
      fileId: fileId,
      fileUrl: fileUrl,
      downloadUrl: downloadUrl,
      fileName: fileName,
      fileSize: estimatedSize
    };
  } catch (e) {
    Logger.log('Upload failed: ' + e.message);
    return { success: false, message: e.message };
  }
}
function verifyFileMagicBytes_(base64Data, declaredMimeType) {
  var decoded = Utilities.base64Decode(base64Data);
  var hex = '';
  for (var i = 0; i < Math.min(4, decoded.length); i++) {
    hex += ('0' + (decoded[i] & 0xFF).toString(16)).slice(-2);
  }
  hex = hex.toUpperCase();
  var expectedMagic = MAGIC_BYTES[declaredMimeType];
  if (!expectedMagic) return true;
  return expectedMagic.some(function(magic) {
    return hex.indexOf(magic) === 0;
  });
}
function deleteUploadedFile(sessionId, fileId, username, deleteDriveFile) {
  try {
    var auth = requireValidSession_(sessionId, username);
    if (!auth.ok) return { success: false, message: auth.message };
    var safeUsername = auth.username;
    if (!fileId) return { success: false, message: 'ไม่ระบุ file ID' };
    if (deleteDriveFile) {
      try { DriveApp.getFileById(fileId).setTrashed(true); }
      catch (e) { Logger.log('Could not delete from Drive: ' + e.message); }
    }
    if (!updateRowByField('UploadedFiles','file_id',fileId,{status:'deleted'}))
      return { success: false, message: 'ไม่พบไฟล์ในระบบ' };
    writeLog('DELETE_FILE', fileId, safeUsername);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function setupSheets() {
  try {
    var ss = SpreadsheetApp.openById(getSheetId_());
    var schemas = {
  'Users':    ['username','password','name','name_en','email','phone',
               'faculty','department','position','roles','status','registered_date'],
  'Sessions': ['session_id','username','created_at','expires_at','ip_address','user_agent','is_active'],
      'Projects':           ['project_code','title_th','title_en','project_type','pi_name','pi_dept','pi_phone','pi_email','objectives','methodology','expected_benefits','budget','duration','gcp_cert','status','submitted_by','created_date','updated_date','final_result','result_note','revision_deadline','revision_round','result_sent_date','parent_code','doc_verified','doc_verified_by','doc_verified_date','cert_file_id','skip_committee_reeval'],
      'ProjectAssignments': ['assignment_id','project_code','project_type','committee_username','committee_name','assigned_date','acceptance_status','eval_status','deadline'],
      'CommitteeReviews':   ['timestamp','project_code','committee_username','committee_name','project_type','decision','reviewer_role','note','eval_answers','final_decision','comment','eval_status','status','meeting_id','agenda_id','agenda_title','resolution','reviewer_no'],
      'UploadedFiles':      ['file_id','project_code','file_name','file_url','download_url','drive_file_id','file_type','mime_type','uploaded_by','uploaded_date','file_size','is_required','status'],
      'Notifications':      ['timestamp','to_username','to_role','project_code','message','is_read'],
      'Logs':               ['timestamp','username','action','detail'],
      'Meetings':           ['meeting_id','title','meeting_date','status','location','note'],
      'ProjectResults': ['timestamp','project_code','final_result','result_note',
                   'revision_deadline','sent_by','researcher_email',
                   'cert_file_id','revision_file_id','skip_committee_reeval'],
      'ProjectRevisions':   ['revision_id','project_code','original_code','round','submitted_by','submitted_date','revision_note','file_ids','status']
    };
    var results = [];
    Object.keys(schemas).forEach(function(name) {
      var headers = schemas[name];
      var sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
        sheet.appendRow(headers);
        results.push('✓ สร้าง: ' + name);
      } else {
        var existH = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
          .map(function(h){ return String(h||'').trim().toLowerCase(); });
        var added = [];
        headers.forEach(function(col) {
          if (existH.indexOf(col.toLowerCase()) === -1) {
            var nextCol = sheet.getLastColumn() + 1;
            sheet.getRange(1, nextCol).setValue(col)
                 .setBackground('#2d8a52').setFontColor('#fff').setFontWeight('bold');
            added.push(col);
          }
        });
        results.push('✓ มีอยู่: ' + name + (added.length ? ' (+' + added.join(',') + ')' : ''));
      }
      if (sheet.getLastRow() >= 1)
        sheet.getRange(1,1,1,sheet.getLastColumn()).setBackground('#2d8a52').setFontColor('#fff').setFontWeight('bold');
    });
    ensureUsersSheet();
    results.push('✓ ผู้ใช้เริ่มต้นพร้อม');
    return '✅ Setup เสร็จสมบูรณ์\n\n' + results.join('\n');
  } catch (e) { return '❌ Error: ' + e.message; }
}
function ensureMeetingMinutesSheet() {
  var ss = SpreadsheetApp.openById(getSheetId_());
  var sheet = ss.getSheetByName('MeetingMinutes');
  if (!sheet) {
    sheet = ss.insertSheet('MeetingMinutes');
    var headers = ['minute_id','meeting_id','meeting_date','title',
                   'uploaded_by','uploaded_date','file_id','file_url',
                   'download_url','file_name','status','access_roles'];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setBackground('#2d8a52').setFontColor('#fff').setFontWeight('bold');
  }
  return sheet;
}
function ensureSessionsSheet_() {
  var ss = SpreadsheetApp.openById(getSheetId_());
  var sheet = ss.getSheetByName('Sessions');
  if (!sheet) {
    sheet = ss.insertSheet('Sessions');
    var headers = ['session_id','username','created_at','expires_at','ip_address','user_agent','is_active'];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setBackground('#2d8a52').setFontColor('#fff').setFontWeight('bold');
  }
  return sheet;
}
function uploadMeetingMinutes(payload, adminUsername) {
  try {
    if (!isAdmin_(adminUsername)) {
      return { success: false, message: 'คุณไม่มีสิทธิ์อัพโหลดรายงานการประชุม (เฉพาะแอดมิน)' };
    }
    if (!payload || !payload.base64Data || !payload.fileName) {
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    }
    ensureMeetingMinutesSheet();
    var folder;
    try { folder = DriveApp.getFolderById(getDriveFolderId_()); }
    catch (e) { return { success: false, message: 'ไม่สามารถเข้าถึงโฟลเดอร์: ' + e.message }; }
    var decoded  = Utilities.base64Decode(payload.base64Data);
    var blob     = Utilities.newBlob(decoded, payload.mimeType || 'application/pdf', payload.fileName);
    var file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId   = file.getId();
    var fileUrl  = file.getUrl();
    var downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
    var ts = new Date();
    var minuteId = 'MIN-' + ts.getTime();
    appendRow('MeetingMinutes', {
      minute_id:    minuteId,
      meeting_id:   payload.meetingId   || '',
      meeting_date: payload.meetingDate || '',
      title:        payload.title       || payload.fileName,
      uploaded_by:  adminUsername,
      uploaded_date: ts,
      file_id:      fileId,
      file_url:     fileUrl,
      download_url: downloadUrl,
      file_name:    payload.fileName,
      status:       'active',
      access_roles: 'admin,committee'
    });
    var committees = sheetToObjects('Users').filter(function(u) {
      return String(u['roles']||'').indexOf('committee') !== -1 &&
             String(u['status']||'active').toLowerCase() === 'active';
    });
    committees.forEach(function(c) {
      addNotification_('committee', String(c['username']||''), '',
        '📄 มีรายงานการประชุมใหม่: ' + (payload.title || payload.fileName) +
        (payload.meetingDate ? ' วันที่ ' + payload.meetingDate : ''));
    });
    writeLog('UPLOAD_MINUTES', minuteId + '|' + payload.fileName, adminUsername);
    return { success: true, minuteId: minuteId, fileUrl: fileUrl, downloadUrl: downloadUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function getMeetingMinutes(username) {
  try {
    var user = sheetToObjects('Users').find(function(u) {
      return String(u['username']||'').trim() === String(username||'').trim();
    });
    if (!user) return { success: false, message: 'ไม่พบผู้ใช้', data: [] };
    var roles = String(user['roles']||'');
    var isAdminOrCommittee = roles.indexOf('admin') !== -1 || roles.indexOf('committee') !== -1;
    if (!isAdminOrCommittee) {
      return { success: false, message: 'คุณไม่มีสิทธิ์เข้าถึงรายงานการประชุม', data: [] };
    }
    var minutes = sheetToObjects('MeetingMinutes').filter(function(m) {
      return String(m['status']||'active') !== 'deleted';
    });
    return {
      success: true,
      data: minutes.map(function(m) {
        return {
          minute_id:    String(m['minute_id']    || ''),
          meeting_id:   String(m['meeting_id']   || ''),
          meeting_date: String(m['meeting_date'] || ''),
          title:        String(m['title']        || ''),
          uploaded_by:  String(m['uploaded_by']  || ''),
          uploaded_date: String(m['uploaded_date'] || ''),
          file_url:     String(m['file_url']     || ''),
          download_url: String(m['download_url'] || ''),
          file_name:    String(m['file_name']    || ''),
          status:       String(m['status']       || 'active')
        };
      })
    };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
function deleteMeetingMinute(minuteId, adminUsername) {
  try {
    if (!isAdmin_(adminUsername)) return { success: false, message: 'ไม่มีสิทธิ์' };
    var ok = updateRowByField('MeetingMinutes', 'minute_id', minuteId, { status: 'deleted' });
    if (!ok) return { success: false, message: 'ไม่พบรายการ' };
    writeLog('DELETE_MINUTES', minuteId, adminUsername);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function ensureResearchReportsSheet() {
  var ss = SpreadsheetApp.openById(getSheetId_());
  var sheet = ss.getSheetByName('ResearchReports');
  if (!sheet) {
    sheet = ss.insertSheet('ResearchReports');
    var headers = ['report_id','project_code','report_type','report_round',
                   'title','submitted_by','submitted_date','file_ids',
                   'status','review_status','reviewed_by','reviewed_date',
                   'review_note','expiry_date'];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold');
  }
  return sheet;
}
function submitResearchReport(payload) {
  try {
    if (!payload || !payload.projectCode || !payload.reportType || !payload.callerUsername) {
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    }
    var callerUsername = String(payload.callerUsername).trim();
    try {
      assertOwnsProject_(payload.projectCode, callerUsername);
    } catch(e) {
      writeLog('SECURITY', 'submitResearchReport unauthorized: ' + callerUsername + ' on ' + payload.projectCode, callerUsername);
      return { success: false, message: e.message };
    }
    var proj = sheetToObjects('Projects').find(function(p) {
      return String(p['project_code']||'').trim() === String(payload.projectCode).trim();
    });
    if (!proj) return { success: false, message: 'ไม่พบโครงการ' };
    var status = String(proj['status']||'');
    var validStatuses = ['รับรอง', 'รายงานความก้าวหน้า', 'ต่ออายุแล้ว'];
    var isValid = validStatuses.some(function(s) { return status.indexOf(s) !== -1; });
    if (!isValid) {
      return { success: false, message: 'โครงการต้องได้รับการรับรองก่อน (สถานะปัจจุบัน: ' + status + ')' };
    }
    ensureResearchReportsSheet();
    var ts = new Date();
    var existingReports = sheetToObjects('ResearchReports').filter(function(r) {
      return String(r['project_code']||'').trim() === String(payload.projectCode).trim() &&
             String(r['report_type']||'').trim() === String(payload.reportType).trim();
    });
    var round = existingReports.length + 1;
    var reportId = 'RPT-' + payload.projectCode + '-' + payload.reportType.toUpperCase().slice(0,3) + '-' + round + '-' + ts.getTime();
    var reportTypeLabels = {
      progress:  'รายงานความก้าวหน้า',
      amendment: 'รายงานการแก้ไขโครงการ',
      sae:       'รายงานเหตุการณ์ไม่พึงประสงค์ (SAE)',
      deviation: 'รายงานการเบี่ยงเบนโครงการ',
      closure:   'รายงานปิดโครงการ',
      renewal:   'รายงานต่ออายุโครงการ'
    };
    var reportLabel = reportTypeLabels[payload.reportType] || payload.reportType;
    appendRow('ResearchReports', {
      report_id:      reportId,
      project_code:   payload.projectCode,
      report_type:    payload.reportType,
      report_round:   round,
      title:          reportLabel + ' ครั้งที่ ' + round + ' — ' + (proj['title_th']||payload.projectCode),
      submitted_by:   callerUsername,
      submitted_date: ts,
      file_ids:       (payload.fileIds || []).join(','),
      status:         'รอตรวจสอบ',
      review_status:  'pending',
      reviewed_by:    '',
      reviewed_date:  '',
      review_note:    payload.note || '',
      expiry_date:    payload.expiryDate || ''
    });
    if (payload.fileIds && payload.fileIds.length) {
      updateUploadedFilesProjectCode_(payload.fileIds, payload.projectCode + '_rpt_' + payload.reportType);
    }
    if (payload.reportType === 'closure') {
      updateRowByField('Projects', 'project_code', payload.projectCode, {
        status: 'รอปิดโครงการ', updated_date: ts
      });
    } else if (payload.reportType === 'renewal') {
      updateRowByField('Projects', 'project_code', payload.projectCode, {
        status: 'รอต่ออายุ', updated_date: ts
      });
    }
    addNotification_('staff', '', payload.projectCode,
      '📊 ' + reportLabel + ' ครั้งที่ ' + round + ' — โครงการ ' + payload.projectCode);
    notifyStaff_(
      '[EC Sansai] ' + reportLabel + ': ' + payload.projectCode,
      'ผู้วิจัย ' + callerUsername + ' ยื่น' + reportLabel + ' ครั้งที่ ' + round +
      '\nโครงการ: ' + payload.projectCode + ' — ' + (proj['title_th']||'') +
      '\nหมายเหตุ: ' + (payload.note||'-')
    );
    writeLog('SUBMIT_REPORT', reportId + '|type:' + payload.reportType + '|round:' + round, callerUsername);
    return { success: true, reportId: reportId, round: round };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function getResearchReports(projectCode, username) {
  try {
    var user = sheetToObjects('Users').find(function(u) {
      return String(u['username']||'').trim() === String(username||'').trim();
    });
    if (!user) return { success: false, data: [] };
    var roles = String(user['roles']||'');
    var isStaffOrAdmin = roles.indexOf('admin') !== -1 || roles.indexOf('staff') !== -1;
    var reports = sheetToObjects('ResearchReports').filter(function(r) {
      var codeMatch = !projectCode || String(r['project_code']||'').trim() === String(projectCode).trim();
      var ownerMatch = isStaffOrAdmin || String(r['submitted_by']||'').trim() === String(username).trim();
      return codeMatch && ownerMatch;
    });
    return {
      success: true,
      data: reports.map(function(r) {
        return {
          report_id:      String(r['report_id']     || ''),
          project_code:   String(r['project_code']  || ''),
          report_type:    String(r['report_type']   || ''),
          report_round:   String(r['report_round']  || '1'),
          title:          String(r['title']         || ''),
          submitted_by:   String(r['submitted_by']  || ''),
          submitted_date: String(r['submitted_date']|| ''),
          status:         String(r['status']        || ''),
          review_status:  String(r['review_status'] || ''),
          reviewed_by:    String(r['reviewed_by']   || ''),
          reviewed_date:  String(r['reviewed_date'] || ''),
          review_note:    String(r['review_note']   || ''),
          expiry_date:    String(r['expiry_date']   || ''),
          file_ids:       String(r['file_ids']      || '')
        };
      }).sort(function(a,b) { return b.submitted_date.localeCompare(a.submitted_date); })
    };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
function reviewResearchReport(payload, staffUsername) {
  try {
    if (!payload || !payload.reportId || !payload.decision) {
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    }
    var ts = new Date();
    var newStatus = payload.decision === 'approved' ? 'ผ่านการพิจารณา' : 'ต้องแก้ไข';
    var ok = updateRowByField('ResearchReports', 'report_id', payload.reportId, {
      status:        newStatus,
      review_status: payload.decision,
      reviewed_by:   staffUsername,
      reviewed_date: ts,
      review_note:   payload.note || ''
    });
    if (!ok) return { success: false, message: 'ไม่พบรายงาน' };
    var report = sheetToObjects('ResearchReports').find(function(r) {
      return String(r['report_id']||'').trim() === String(payload.reportId).trim();
    });
    if (report) {
      var submittedBy = String(report['submitted_by']||'');
      addNotification_('researcher', submittedBy, String(report['project_code']||''),
        (payload.decision === 'approved' ? '✅ รายงานผ่านการพิจารณา' : '⚠️ รายงานต้องแก้ไข') +
        ': ' + String(report['title']||''));
      if (payload.decision === 'approved' && String(report['report_type']||'') === 'renewal') {
        var newExpiry = payload.newExpiryDate || '';
        if (newExpiry) {
          updateRowByField('Projects', 'project_code', String(report['project_code']||''), {
            status: 'ต่ออายุแล้ว', updated_date: ts
          });
        }
      }
      if (payload.decision === 'approved' && String(report['report_type']||'') === 'closure') {
        updateRowByField('Projects', 'project_code', String(report['project_code']||''), {
          status: 'ปิดโครงการแล้ว', updated_date: ts
        });
      }
      var resUser = submittedBy ? sheetToObjects('Users').find(function(u) {
        return String(u['username']||'').trim() === submittedBy;
      }) : null;
      if (resUser && resUser['email'] && validateEmail_(resUser['email'])) {
        var icon = payload.decision === 'approved' ? '✅' : '⚠️';
        sendEmail_(String(resUser['email']).trim(),
          '[EC Sansai] ' + icon + ' ผลการตรวจสอบรายงาน: ' + String(report['project_code']||''),
          'คุณ ' + (resUser['name']||submittedBy) + '\n\n' +
          icon + ' ' + newStatus + '\n' +
          'รายงาน: ' + String(report['title']||'') + '\n' +
          (payload.note ? '\nหมายเหตุ:\n' + payload.note : '')
        );
      }
    }
    writeLog('REVIEW_REPORT', payload.reportId + '→' + payload.decision, staffUsername);
    return { success: true, newStatus: newStatus };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function getApprovedProjectsForReporting(username) {
  try {
    var allProjects = sheetToObjects('Projects');
    var validStatuses = ['รับรอง', 'ต่ออายุแล้ว', 'รายงานความก้าวหน้า', 'รอปิดโครงการ', 'ปิดโครงการแล้ว'];
    var myProjects = allProjects.filter(function(p) {
      var status = String(p['status']||'');
      var isOwner = String(p['submitted_by']||'').trim() === String(username||'').trim();
      var isApproved = validStatuses.some(function(s) { return status.indexOf(s) !== -1; });
      return isOwner && isApproved;
    });
    return {
      success: true,
      data: myProjects.map(function(p) {
        return {
          project_code: String(p['project_code'] || ''),
          title_th: String(p['title_th'] || ''),
          status: String(p['status'] || ''),
          duration: String(p['duration'] || '')
        };
      })
    };
  } catch (e) {
    return { success: false, data: [] };
  }
}
function getAllResearchReports(callerUsername) {
  try {
    if (!callerUsername) return { success: false, message: 'ไม่ระบุผู้ใช้', data: [] };
    var user = sheetToObjects('Users').find(function(u) {
      return String(u['username'] || '').trim() === String(callerUsername).trim();
    });
    if (!user) return { success: false, message: 'ไม่พบผู้ใช้', data: [] };
    var roles = String(user['roles'] || '');
    if (roles.indexOf('admin') === -1 && roles.indexOf('staff') === -1) {
      return { success: false, message: 'คุณไม่มีสิทธิ์เข้าถึงข้อมูลนี้', data: [] };
    }
    var reports  = sheetToObjects('ResearchReports');
    var projects = sheetToObjects('Projects');
    var projMap  = {};
    projects.forEach(function(p) {
      projMap[String(p['project_code'] || '')] = {
        title_th: String(p['title_th'] || ''),
        pi_name:  String(p['pi_name']  || '')
      };
    });
    return {
      success: true,
      data: reports.map(function(r) {
        var code = String(r['project_code'] || '');
        var proj = projMap[code] || {};
        return {
          report_id:      String(r['report_id']      || ''),
          project_code:   code,
          title_th:       proj.title_th || '',
          pi_name:        proj.pi_name  || '',
          report_type:    String(r['report_type']    || ''),
          report_round:   String(r['report_round']   || ''),
          title:          String(r['title']          || ''),
          submitted_by:   String(r['submitted_by']   || ''),
          submitted_date: String(r['submitted_date'] || ''),
          status:         String(r['status']         || ''),
          review_status:  String(r['review_status']  || ''),
          reviewed_by:    String(r['reviewed_by']    || ''),
          expiry_date:    String(r['expiry_date']    || '')
        };
      }).sort(function(a, b) { return b.submitted_date.localeCompare(a.submitted_date); })
    };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
function getClosureReportFiles() {
  try {
    var reports = sheetToObjects('ResearchReports');
    var files = sheetToObjects('UploadedFiles');
    var closureReports = reports.filter(function(r) {
      return String(r['report_type']||'').toLowerCase() === 'closure' &&
             String(r['status']||'').includes('ผ่าน');
    });
    var result = [];
    closureReports.forEach(function(rep) {
      var projCode = String(rep['project_code']||'');
      var fileList = files.filter(function(f) {
        return String(f['project_code']||'').trim() === projCode &&
               String(f['status']||'active') !== 'deleted';
      });
      fileList.forEach(function(f) {
        result.push({
          project_code: projCode,
          title_th: rep['title'] || '',
          file_name: String(f['file_name']||''),
          uploaded_date: String(f['uploaded_date']||''),
          download_url: String(f['download_url']||'')
        });
      });
    });
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function reviewRevisionByAdmin(payload, staffUsername) {
  try {
    if (!payload || !payload.projectCode || !payload.decision) {
      return { success: false, message: 'ข้อมูลไม่ครบ' };
    }
    var projectCode = String(payload.projectCode).trim();
    var decision = String(payload.decision);
    var note = String(payload.note || '');
    var skipCommittee = payload.skipCommitteeReeval === true;
    var ts = new Date();
    var proj = sheetToObjects('Projects').find(function(p) {
      return String(p['project_code'] || '').trim() === projectCode;
    });
    if (!proj) return { success: false, message: 'ไม่พบโครงการแก้ไข' };
    var newStatus = decision === 'approved'
      ? (skipCommittee ? 'รับรอง (ข้ามกรรมการ)' : 'รับรองมีข้อแก้ไขแล้ว')
      : 'ต้องแก้ไขเพิ่ม';
    updateRowByField('Projects', 'project_code', projectCode, {
      status: newStatus,
      final_result: decision === 'approved' ? 'รับรอง' : 'ไม่รับรอง',
      result_note: note,
      result_sent_date: ts,
      updated_date: ts,
      skip_committee_reeval: skipCommittee ? 'true' : 'false'
    });
    if (proj['parent_code']) {
      updateRowByField('Projects', 'project_code', proj['parent_code'], {
        status: decision === 'approved' ? 'รับรองแล้ว (แก้ไขแล้ว)' : 'รับรองมีข้อแก้ไข (ยังไม่ผ่าน)',
        updated_date: ts
      });
    }
    writeLog('REVIEW_REVISION', projectCode + ' → ' + decision + (skipCommittee ? ' (skip committee)' : ''), staffUsername);
    var piEmail = String(proj['pi_email'] || '');
    if (piEmail && validateEmail_(piEmail)) {
      var icon = decision === 'approved' ? '✅' : '⚠️';
      sendEmail_(piEmail,
        '[EC Sansai] ' + icon + ' ผลการตรวจสอบการแก้ไข: ' + projectCode,
        'เรียนคุณ ' + (proj['pi_name'] || 'ผู้วิจัย') + '\n\n' +
        icon + ' การแก้ไขโครงการ ' + projectCode + ' ได้รับการ ' + (decision === 'approved' ? 'อนุมัติ' : 'ปฏิเสธ') +
        '\n\nหมายเหตุจากแอดมิน:\n' + (note || '-') +
        (decision === 'approved' ? '\n\n✅ โครงการของคุณได้รับการรับรองแล้ว' : '')
      );
    }
    return { success: true, newStatus: newStatus, skippedCommittee: skipCommittee };
  } catch (e) {
    Logger.log('reviewRevisionByAdmin error: ' + e.message);
    return { success: false, message: e.message };
  }
}
function requireRole_(callerUsername, allowedRoles) {
  var user = sheetToObjects('Users').find(function(u) {
    return String(u['username']||'').trim() === String(callerUsername||'').trim();
  });
  if (!user || String(user['status']||'') !== 'active')
    throw new Error('Unauthorized');
  var userRoles = String(user['roles']||'').split(',');
  var allowed = allowedRoles.some(function(r) {
    return userRoles.indexOf(r) !== -1;
  });
  if (!allowed) throw new Error('Forbidden: requires ' + allowedRoles.join(' or '));
  return user;
}
function ensureNewSheets() { return setupSheets(); }
function testSheetConnection() {
  try {
    var ss    = SpreadsheetApp.openById(getSheetId_());
    var users = sheetToObjects('Users');
    return {
      success: true, spreadsheetName: ss.getName(),
      sheets: ss.getSheets().map(function(s){ return {name:s.getName(),rows:s.getLastRow()}; }),
      userCount: users.length
    };
  } catch (e) { return { success: false, message: e.message }; }
}
function authorizeDriveAccess() {
  try {
    var folder   = DriveApp.getFolderById(getDriveFolderId_());
    var testBlob = Utilities.newBlob('test', 'text/plain', 'test.txt');
    var testFile = folder.createFile(testBlob);
    testFile.setTrashed(true);
    return '✅ Drive permissions OK: ' + folder.getName();
  } catch (e) { return '❌ Drive error: ' + e.message; }
}
function writeAuditLog_(action, detail, username, severity, extra) {
  try {
    var ss = SpreadsheetApp.openById(getSheetId_());
    var sheet = ss.getSheetByName('AuditLog');
    if (!sheet) {
      sheet = ss.insertSheet('AuditLog');
      var headers = ['timestamp', 'username', 'action', 'detail', 'severity', 'extra'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setBackground('#b71c1c').setFontColor('#ffffff').setFontWeight('bold');
    }
    var sev = severity || 'INFO';
    var rowColor = sev === 'CRITICAL' ? '#ffcdd2' : sev === 'WARN' ? '#fff9c4' : '#ffffff';
    var lastRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      new Date(),
      String(username || 'system'),
      String(action || ''),
      String(detail || '').substring(0, 1000),
      sev,
      String(extra || '').substring(0, 500)
    ]);
    if (sev !== 'INFO') {
      sheet.getRange(lastRow, 1, 1, 6).setBackground(rowColor);
    }
  } catch (e) {
    Logger.log('[AuditLog write failed] ' + e.message);
  }
}
function validateResetTokenSecure_(token) {
  var result = { success: false };
  try {
    if (!token || typeof token !== 'string') {
      writeAuditLog_('RESET_TOKEN_INVALID', 'Token missing or wrong type', 'anonymous', 'WARN');
      return result;
    }
    if (token.length > 100) {
      writeAuditLog_('RESET_TOKEN_SUSPICIOUS', 'Token length: ' + token.length, 'anonymous', 'WARN');
      return result;
    }
    if (!/^[a-zA-Z0-9\-]+$/.test(token)) {
      writeAuditLog_('RESET_TOKEN_INJECTION', 'Invalid chars in token', 'anonymous', 'CRITICAL');
      return result;
    }
    var cache = CacheService.getScriptCache();
    var data = cache.get('reset_' + token);
    if (!data) {
      writeAuditLog_('RESET_TOKEN_NOT_FOUND', 'Token not in cache', 'anonymous', 'WARN');
      return result;
    }
    var resetData = JSON.parse(data);
    if (resetData.expiresAt < Date.now()) {
      cache.remove('reset_' + token);
      writeAuditLog_('RESET_TOKEN_EXPIRED', 'Token expired', resetData.username || 'unknown', 'WARN');
      return result;
    }
    if (resetData.used === true) {
      writeAuditLog_('RESET_TOKEN_REUSE', 'Token already used', resetData.username || 'unknown', 'CRITICAL');
      return result;
    }
    result = { success: true, username: resetData.username, email: resetData.email };
  } catch (e) {
    Logger.log('validateResetTokenSecure_ error: ' + e.message);
  }
  return result;
}
function markResetTokenUsed_(token) {
  try {
    var cache = CacheService.getScriptCache();
    var data = cache.get('reset_' + token);
    if (data) {
      var resetData = JSON.parse(data);
      resetData.used = true;
      cache.put('reset_' + token, JSON.stringify(resetData), 60);
    }
  } catch (e) {
    Logger.log('markResetTokenUsed_ error: ' + e.message);
  }
}
function resetPasswordSecure(token, newPassword, confirmPassword) {
  var callerInfo = 'resetPasswordSecure';
  try {
    var validation = validateResetTokenSecure_(token);
    if (!validation.success) {
      return { success: false, message: 'ลิงก์ไม่ถูกต้องหรือหมดอายุแล้ว' };
    }
    markResetTokenUsed_(token);
    var result = resetPassword(token, newPassword, confirmPassword);
    if (result.success) {
      writeAuditLog_('PASSWORD_RESET_SUCCESS', 'User: ' + validation.username, validation.username, 'INFO');
    } else {
      writeAuditLog_('PASSWORD_RESET_FAILED', result.message + ' | User: ' + validation.username, validation.username, 'WARN');
    }
    return result;
  } catch (e) {
    Logger.log(callerInfo + ' error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด กรุณาลองอีกครั้ง' };
  }
}
var SESSION_INACTIVITY_MINUTES = 480;
function touchSession(sessionId, username) {
  try {
    if (!sessionId || !username) return { success: false };
    var cache = CacheService.getScriptCache();
    var key = 'sess_activity_' + sessionId;
    cache.put(key, JSON.stringify({
      username: username,
      lastActive: Date.now()
    }), SESSION_INACTIVITY_MINUTES * 60);
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}
function validateSessionSecure(sessionId, username) {
  try {
    var baseResult = validateSession(sessionId, username);
    if (!baseResult.success) {
      writeAuditLog_('SESSION_INVALID', 'sessionId: ' + String(sessionId || '').substring(0, 8) + '...', username, 'WARN');
      return baseResult;
    }
    var cache = CacheService.getScriptCache();
    var key = 'sess_activity_' + sessionId;
    var raw = cache.get(key);
    if (raw) {
      var data = JSON.parse(raw);
      var idleMs = Date.now() - data.lastActive;
      var idleMinutes = idleMs / 60000;
      if (idleMinutes > SESSION_INACTIVITY_MINUTES) {
        updateRowByField('Sessions', 'session_id', sessionId, { is_active: false });
        cache.remove(key);
        writeAuditLog_('SESSION_TIMEOUT', 'Idle: ' + Math.round(idleMinutes) + 'min', username, 'INFO');
        return { success: false, message: 'Session หมดอายุจากการไม่ใช้งาน' };
      }
      data.lastActive = Date.now();
      cache.put(key, JSON.stringify(data), SESSION_INACTIVITY_MINUTES * 60);
    }
    return { success: true };
  } catch (e) {
    Logger.log('validateSessionSecure error: ' + e.message);
    return validateSession(sessionId, username);
  }
}
function logoutSecure(sessionId, username) {
  try {
    if (!sessionId) return { success: false };
    updateRowByField('Sessions', 'session_id', sessionId, { is_active: false });
    var cache = CacheService.getScriptCache();
    cache.remove('sess_activity_' + sessionId);
    writeLog('LOGOUT', 'session: ' + String(sessionId).substring(0, 8) + '...', username);
    writeAuditLog_('LOGOUT', 'Explicit logout', username, 'INFO');
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}
function sanitizeForHtml_(text, maxLen) {
  maxLen = maxLen || 500;
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;')
    .replace(/\\/g, '&#x5C;')
    .replace(/`/g, '&#x60;')
    .replace(/=/g, '&#x3D;')
    .trim()
    .substring(0, maxLen);
}
function sanitizeProjectCode_(code) {
  if (!code) return '';
  var cleaned = String(code).trim();
  if (!/^[a-zA-Zก-๙0-9\-_]{3,50}$/.test(cleaned)) {
    Logger.log('⚠️ Invalid project code format: ' + cleaned);
    return '';
  }
  return cleaned;
}
function sanitizeUsername_(username) {
  if (!username) return '';
  var cleaned = String(username).trim();
  if (!/^[a-zA-Z0-9_]{4,20}$/.test(cleaned)) return '';
  return cleaned;
}
var RATE_LIMITS = {
  'upload':      { max: 10,  windowMin: 5,  lockMin: 15 },
  'register':    { max: 3,   windowMin: 60, lockMin: 60 },
  'resetReq':    { max: 3,   windowMin: 30, lockMin: 30 },
  'submitProj':  { max: 5,   windowMin: 10, lockMin: 10 },
  'default':     { max: 30,  windowMin: 1,  lockMin: 5  }
};
function checkActionRateLimit_(identifier, action) {
  try {
    var limit = RATE_LIMITS[action] || RATE_LIMITS['default'];
    var cache = CacheService.getScriptCache();
    var key = 'ratelimit_' + action + '_' + String(identifier || 'anon');
    var raw = cache.get(key);
    var data = raw ? JSON.parse(raw) : { count: 0, lockedUntil: 0, windowStart: Date.now() };
    var now = Date.now();
    if (data.lockedUntil && data.lockedUntil > now) {
      var remaining = Math.ceil((data.lockedUntil - now) / 60000);
      writeAuditLog_('RATE_LIMIT_BLOCKED', action + ' | ' + identifier, identifier, 'WARN');
      return { blocked: true, message: 'คำขอถูกจำกัดชั่วคราว กรุณารอ ' + remaining + ' นาที' };
    }
    if ((now - data.windowStart) > (limit.windowMin * 60 * 1000)) {
      data = { count: 0, lockedUntil: 0, windowStart: now };
    }
    data.count++;
    if (data.count > limit.max) {
      data.lockedUntil = now + (limit.lockMin * 60 * 1000);
      writeAuditLog_('RATE_LIMIT_TRIGGERED', action + ' count:' + data.count, identifier, 'WARN');
    }
    var ttl = Math.max(limit.lockMin, limit.windowMin) * 60 * 2;
    cache.put(key, JSON.stringify(data), ttl);
    if (data.lockedUntil > now) {
      var rem = Math.ceil((data.lockedUntil - now) / 60000);
      return { blocked: true, message: 'คำขอมากเกินไป กรุณารอ ' + rem + ' นาที' };
    }
    return { blocked: false };
  } catch (e) {
    Logger.log('checkActionRateLimit_ error: ' + e.message);
    writeAuditLog_('RATE_LIMIT_ERROR', action + ' | ' + e.message, String(identifier || 'anon'), 'WARN');
    return { blocked: true, message: 'ระบบไม่สามารถตรวจสอบได้ชั่วคราว กรุณาลองใหม่ในอีกสักครู่' };
  }
}
function uploadFileSecure(base64Data, fileName, mimeType, projectCode, username, fileType) {
  try {
    var rl = checkActionRateLimit_(username || 'anon', 'upload');
    if (rl.blocked) return { success: false, message: rl.message };
    var safeFileName = sanitizeForHtml_(fileName, 200);
    var safeProjectCode = sanitizeProjectCode_(projectCode);
    var safeUsername = sanitizeUsername_(username);
    var safeFileType = String(fileType || 'other').replace(/[^a-zA-Z0-9_-]/g, '').substring(0, 50);
    if (!safeFileName) return { success: false, message: 'ชื่อไฟล์ไม่ถูกต้อง' };
    if (projectCode && !safeProjectCode) return { success: false, message: 'รหัสโครงการไม่ถูกต้อง' };
    if (!safeUsername) return { success: false, message: 'ไม่ระบุผู้ใช้' };
    if (mimeType && ALLOWED_MIME_TYPES.indexOf(mimeType) === -1) {
      writeAuditLog_('UPLOAD_BLOCKED_MIME', 'File: ' + safeFileName + ' | MIME: ' + mimeType, safeUsername, 'WARN');
      return { success: false, message: 'ประเภทไฟล์ไม่ได้รับอนุญาต' };
    }
    var ext = safeFileName.split('.').pop().toLowerCase();
    var dangerousExts = ['php', 'js', 'html', 'exe', 'bat', 'sh', 'py', 'rb', 'jsp', 'asp', 'aspx', 'cgi'];
    var allParts = safeFileName.toLowerCase().split('.');
    var hasDanger = allParts.some(function(p) { return dangerousExts.indexOf(p) !== -1; });
    if (hasDanger) {
      writeAuditLog_('UPLOAD_BLOCKED_EXT', 'File: ' + safeFileName, safeUsername, 'CRITICAL');
      return { success: false, message: 'ชื่อไฟล์มีนามสกุลที่ไม่อนุญาต' };
    }
    writeAuditLog_('UPLOAD_ATTEMPT', safeFileName + ' | ' + mimeType + ' | project:' + (safeProjectCode || 'TEMP'), safeUsername, 'INFO');
    var result = uploadFileWithType(base64Data, safeFileName, mimeType, safeProjectCode, safeUsername, safeFileType);
    if (!result.success) {
      writeAuditLog_('UPLOAD_FAILED', result.message + ' | ' + safeFileName, safeUsername, 'WARN');
    }
    return result;
  } catch (e) {
    Logger.log('uploadFileSecure error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาดในการอัปโหลด' };
  }
}
function registerUserSecure(payload) {
  try {
    var identifier = String(payload && payload.username ? payload.username : 'anon');
    var rl = checkActionRateLimit_(identifier, 'register');
    if (rl.blocked) return { success: false, message: rl.message };
    writeAuditLog_('REGISTER_ATTEMPT', 'Username: ' + identifier, identifier, 'INFO');
    var result = registerUser(payload);
    if (result.success) {
      writeAuditLog_('REGISTER_SUCCESS', 'Username: ' + identifier, identifier, 'INFO');
    } else {
      writeAuditLog_('REGISTER_FAILED', result.message + ' | ' + identifier, identifier, 'WARN');
    }
    return result;
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}
function requestPasswordResetSecure(email) {
  try {
    var identifier = String(email || 'anon').toLowerCase();
    var rl = checkActionRateLimit_(identifier, 'resetReq');
    if (rl.blocked) return { success: true, message: 'หากอีเมลนี้มีในระบบ เราจะส่งลิงก์รีเซ็ตรหัสผ่านไปให้' };
    writeAuditLog_('PASSWORD_RESET_REQUEST', 'Email: ' + identifier, 'anonymous', 'INFO');
    return requestPasswordReset(email);
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}
function saveOrUpdateProjectSecure(projectData, isUpdate) {
  try {
    var username = String(projectData && projectData.submittedBy ? projectData.submittedBy : 'anon');
    var rl = checkActionRateLimit_(username, 'submitProj');
    if (rl.blocked) return { success: false, message: rl.message };
    if (projectData) {
      if (projectData.title_th)  projectData.title_th  = sanitizeForHtml_(projectData.title_th, 500);
      if (projectData.title_en)  projectData.title_en  = sanitizeForHtml_(projectData.title_en, 500);
      if (projectData.objectives) projectData.objectives = sanitizeForHtml_(projectData.objectives, 3000);
      if (projectData.methodology) projectData.methodology = sanitizeForHtml_(projectData.methodology, 3000);
    }
    writeAuditLog_('PROJECT_SUBMIT', (isUpdate ? 'UPDATE' : 'NEW') + ' | by: ' + username, username, 'INFO');
    return saveOrUpdateProject(projectData, isUpdate);
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}
var SAFE_ERROR_MESSAGES = {
  'not found':      'ไม่พบข้อมูลที่ต้องการ',
  'permission':     'คุณไม่มีสิทธิ์ดำเนินการนี้',
  'unauthorized':   'กรุณาเข้าสู่ระบบใหม่',
  'forbidden':      'คุณไม่มีสิทธิ์ดำเนินการนี้',
  'timeout':        'คำขอใช้เวลานานเกินไป กรุณาลองใหม่',
  'quota':          'ระบบกำลังยุ่ง กรุณาลองใหม่ในภายหลัง',
  'default':        'เกิดข้อผิดพลาด กรุณาลองอีกครั้ง'
};
function sanitizeErrorMessage_(errorMsg) {
  var msg = String(errorMsg || '').toLowerCase();
  for (var key in SAFE_ERROR_MESSAGES) {
    if (key !== 'default' && msg.indexOf(key) !== -1) {
      return SAFE_ERROR_MESSAGES[key];
    }
  }
  if (/[\u0E00-\u0E7F]/.test(errorMsg) && errorMsg.length < 200) {
    return errorMsg;
  }
  return SAFE_ERROR_MESSAGES['default'];
}
function addSecurityHeaders_(htmlOutput) {
  try {
    htmlOutput.addMetaTag('Content-Security-Policy',
      "default-src 'self' https://script.google.com; " +
      "script-src 'self' 'unsafe-inline' https://script.google.com; " +
      "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; " +
      "font-src https://fonts.gstatic.com; " +
      "img-src 'self' data: https:; " +
      "connect-src https://script.google.com;"
    );
    htmlOutput.addMetaTag('X-Content-Type-Options', 'nosniff');
    htmlOutput.addMetaTag('Referrer-Policy', 'strict-origin-when-cross-origin');
  } catch (e) {
    Logger.log('addSecurityHeaders_ error: ' + e.message);
  }
  return htmlOutput;
}
var RETENTION_POLICIES = {
  'Logs':         365,
  'AuditLog':     730,
  'Sessions':     60,
  'Notifications': 90
};
function runDataRetention() {
  var results = [];
  try {
    Object.keys(RETENTION_POLICIES).forEach(function(sheetName) {
      var days = RETENTION_POLICIES[sheetName];
      var cutoff = new Date(Date.now() - (days * 24 * 60 * 60 * 1000));
      var deleted = deleteOldRows_(sheetName, 'timestamp', cutoff);
      results.push(sheetName + ': ลบ ' + deleted + ' แถว (เก่ากว่า ' + days + ' วัน)');
    });
    writeAuditLog_('DATA_RETENTION', results.join(' | '), 'system', 'INFO');
    Logger.log('Data Retention: ' + results.join(', '));
  } catch (e) {
    Logger.log('runDataRetention error: ' + e.message);
    writeAuditLog_('DATA_RETENTION_ERROR', e.message, 'system', 'WARN');
  }
  return results;
}
function deleteOldRows_(sheetName, timestampCol, cutoffDate) {
  var deleted = 0;
  try {
    var ss = SpreadsheetApp.openById(getSheetId_());
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var tsCol = headers.indexOf(timestampCol.toLowerCase());
    if (tsCol === -1) return 0;
    for (var i = data.length - 1; i >= 1; i--) {
      var cellVal = data[i][tsCol];
      if (!cellVal) continue;
      var rowDate = new Date(cellVal);
      if (isNaN(rowDate.getTime())) continue;
      if (rowDate < cutoffDate) {
        sheet.deleteRow(i + 1);
        deleted++;
      }
    }
  } catch (e) {
    Logger.log('deleteOldRows_ error [' + sheetName + ']: ' + e.message);
  }
  return deleted;
}
var ALLOWED_ROLES = ['admin', 'staff', 'committee', 'researcher'];
function assertValidRole_(username, claimedRoles) {
  try {
    if (!username) throw new Error('No username');
    var user = sheetToObjects('Users').find(function(u) {
      return String(u['username'] || '').trim() === String(username).trim();
    });
    if (!user) throw new Error('User not found');
    if (String(user['status'] || '') !== 'active') throw new Error('User not active');
    var realRoles = String(user['roles'] || '').split(',').map(function(r) { return r.trim(); });
    if (Array.isArray(claimedRoles)) {
      claimedRoles.forEach(function(cr) {
        if (ALLOWED_ROLES.indexOf(cr) !== -1 && realRoles.indexOf(cr) === -1) {
          writeAuditLog_('PRIVILEGE_ESCALATION_ATTEMPT',
            'User: ' + username + ' claimed: ' + cr + ' but has: ' + realRoles.join(','),
            username, 'CRITICAL');
          throw new Error('Privilege escalation detected');
        }
      });
    }
    return { valid: true, realRoles: realRoles };
  } catch (e) {
    return { valid: false, message: e.message };
  }
}
function runSecurityHealthCheck(adminUsername) {
  try {
    if (adminUsername && !isAdmin_(adminUsername)) {
      return { success: false, message: 'เฉพาะ admin เท่านั้น' };
    }
    var report = [];
    var warnings = [];
    var users = sheetToObjects('Users');
    var plainPwUsers = users.filter(function(u) {
      return String(u['password'] || '').trim().length > 0;
    });
    if (plainPwUsers.length > 0) {
      warnings.push('⚠️ พบ ' + plainPwUsers.length + ' ผู้ใช้ที่ยังมี plain text password — รัน migratePasswordsToHash() ด้วย');
    } else {
      report.push('✅ ไม่พบ plain text password');
    }
    var sessions = sheetToObjects('Sessions');
    var expiredActive = sessions.filter(function(s) {
      var isActive = s['is_active'] === true || String(s['is_active']).toLowerCase() === 'true';
      var expiry = new Date(s['expires_at']);
      return isActive && expiry < new Date();
    });
    if (expiredActive.length > 0) {
      warnings.push('⚠️ พบ ' + expiredActive.length + ' Session หมดอายุแต่ยัง active — ควรรัน cleanExpiredSessions()');
    } else {
      report.push('✅ Sessions สะอาด');
    }
    var admins = users.filter(function(u) {
      return String(u['roles'] || '').indexOf('admin') !== -1 &&
             String(u['status'] || '') === 'active';
    });
    report.push('ℹ️ Admin accounts ที่ active: ' + admins.length + ' บัญชี');
    if (admins.length > 3) {
      warnings.push('⚠️ มี admin มากกว่า 3 บัญชี — ควรทบทวนสิทธิ์');
    }
    var ss = SpreadsheetApp.openById(getSheetId_());
    var auditSheet = ss.getSheetByName('AuditLog');
    if (!auditSheet) {
      warnings.push('⚠️ ยังไม่มี AuditLog sheet — รัน writeAuditLog_ สักครั้งเพื่อสร้าง');
    } else {
      report.push('✅ AuditLog sheet มีอยู่ (' + (auditSheet.getLastRow() - 1) + ' records)');
    }
    var cache = CacheService.getScriptCache();
    report.push('✅ CacheService พร้อมใช้งาน');
    var summary = {
      success: true,
      checkedAt: new Date().toISOString(),
      warnings: warnings.length,
      report: report,
      warningDetails: warnings
    };
    writeAuditLog_('SECURITY_HEALTH_CHECK',
      'Warnings: ' + warnings.length + ' | ' + warnings.join('; '),
      adminUsername || 'system',
      warnings.length > 0 ? 'WARN' : 'INFO');
    return summary;
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function cleanExpiredSessions() {
  try {
    var sheet = getSheet('Sessions');
    if (!sheet || sheet.getLastRow() < 2) return '✅ ไม่มี session ให้ทำความสะอาด';
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var expiryCol = headers.indexOf('expires_at');
    var activeCol = headers.indexOf('is_active');
    if (expiryCol === -1) return '❌ ไม่พบคอลัมน์ expires_at';
    var now = new Date();
    var cleaned = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      var expiry = new Date(data[i][expiryCol]);
      if (!isNaN(expiry.getTime()) && expiry < now) {
        if (activeCol !== -1) sheet.getRange(i + 1, activeCol + 1).setValue(false);
        cleaned++;
      }
    }
    writeAuditLog_('CLEAN_SESSIONS', 'Cleaned: ' + cleaned, 'system', 'INFO');
    return '✅ ล้าง expired sessions: ' + cleaned + ' รายการ';
  } catch (e) {
    return '❌ Error: ' + e.message;
  }
}
function setupTimeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var name = t.getHandlerFunction();
    if (name === 'cleanExpiredSessions' || name === 'runDataRetention') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('cleanExpiredSessions')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();
  ScriptApp.newTrigger('runDataRetention')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .create();
  Logger.log('✅ Triggers ตั้งค่าเรียบร้อย');
}
function recordPdpaConsent(payload) {
  try {
    var ss = SpreadsheetApp.openById(getSheetId_());
    var sheet = ss.getSheetByName('PdpaConsents');
    if (!sheet) {
      sheet = ss.insertSheet('PdpaConsents');
      var headers = ['timestamp','username','ip_address','user_agent','consent_version','action','page','consent_required','consent_email','consent_audit'];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length)
           .setBackground('#1a6b3a').setFontColor('#ffffff').setFontWeight('bold');
    }
    var action = payload.action || 'accept';
    var consentEmail = true;
    var consentAudit = true;
    if (action.indexOf('custom:') === 0) {
      var parts = action.replace('custom:','').split(',');
      parts.forEach(function(p) {
        var kv = p.split('=');
        if (kv[0] === 'email') consentEmail = kv[1] === 'true';
        if (kv[0] === 'audit') consentAudit = kv[1] === 'true';
      });
      action = 'custom';
    }
    sheet.appendRow([
      new Date(),
      payload.username || 'anonymous',
      payload.ip        || '',
      payload.userAgent || '',
      payload.version   || '1.0',
      action,
      payload.page      || 'login',
      true,
      consentEmail,
      consentAudit
    ]);
    writeAuditLog_('PDPA_CONSENT', action + ' v' + (payload.version||'1.0') + ' email=' + consentEmail + ' audit=' + consentAudit, payload.username || 'anonymous', 'info');
    return { success: true };
  } catch(e) {
    Logger.log('recordPdpaConsent error: ' + e.message);
    return { success: false, message: e.message };
  }
}
function requireValidSession_(sessionId, username) {
  if (!sessionId || !username) {
    writeAuditLog_('AUTH_GUARD_FAIL', 'missing sessionId or username', String(username || 'anonymous'), 'WARN');
    return { ok: false, message: 'กรุณาเข้าสู่ระบบก่อน' };
  }
  try {
    var v = validateSession(sessionId, username);
    if (!v.success) {
      writeAuditLog_('AUTH_GUARD_FAIL', v.message, String(username), 'WARN');
      return { ok: false, message: v.message };
    }
    return { ok: true, username: v.username };
  } catch (e) {
    Logger.log('requireValidSession_ error: ' + e.message);
    return { ok: false, message: 'เกิดข้อผิดพลาดในการตรวจสอบ session' };
  }
}
function checkKeyAge_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var keySetAt = props.getProperty('AES_KEY_SET_AT');
    if (!keySetAt) {
      props.setProperty('AES_KEY_SET_AT', new Date().toISOString());
      Logger.log('[checkKeyAge_] AES_KEY_SET_AT initialized: ' + new Date().toISOString());
      return;
    }
    var ageDays = (Date.now() - new Date(keySetAt).getTime()) / 86400000;
    Logger.log('[checkKeyAge_] AES key age: ' + Math.floor(ageDays) + ' days');
    if (ageDays > 90) {
      MailApp.sendEmail(
        'sireetorn.wa@sansaihospital.go.th',
        '[EC System] ⚠️ AES Key อายุเกิน 90 วัน — ควร Rotate ด่วน',
        'AES_SECRET_KEY ถูกตั้งเมื่อ: ' + keySetAt + '\n' +
        'อายุปัจจุบัน: ' + Math.floor(ageDays) + ' วัน\n\n' +
        'วิธีการ Rotate Key:\n' +
        '1. สร้าง key ใหม่ด้วย Utilities.getUuid() + Utilities.getUuid()\n' +
        '2. รัน Script เพื่อ re-encrypt ข้อมูลใน Sheet ด้วย key ใหม่\n' +
        '3. อัปเดต AES_SECRET_KEY ใน GAS Script Properties\n' +
        '4. อัปเดต AES_KEY_SET_AT เป็น: ' + new Date().toISOString() + '\n\n' +
        'หากมีข้อสงสัยติดต่อผู้ดูแลระบบ'
      );
      writeAuditLog_('KEY_ROTATION_ALERT', 'AES key age: ' + Math.floor(ageDays) + ' days', 'system', 'WARN');
    }
  } catch (e) {
    Logger.log('[checkKeyAge_] error: ' + e.message);
  }
}
function dailySecurityCheck_() {
  checkKeyAge_();
  writeAuditLog_('DAILY_SECURITY_CHECK', 'completed', 'system', 'INFO');
}
function setupSecurityTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailySecurityCheck_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('dailySecurityCheck_')
    .timeBased().everyDays(1).atHour(1).create();
  Logger.log('[setupSecurityTriggers] ✅ dailySecurityCheck_ trigger created');
}
function generateMyKey() {
  var key = Utilities.getUuid() + Utilities.getUuid();
  Logger.log('KEY: ' + key);
}
function debugDoGet() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('index');
    Logger.log('✅ index โหลดได้ ความยาว: ' + html.getContent().length + ' chars');
  } catch(e) {
    Logger.log('❌ index ERROR: ' + e.message);
  }
  
  try {
    var props = PropertiesService.getScriptProperties();
    Logger.log('SHEET_ID: ' + (props.getProperty('SHEET_ID') ? '✅' : '❌ ไม่มี'));
    Logger.log('DRIVE_FOLDER_ID: ' + (props.getProperty('DRIVE_FOLDER_ID') ? '✅' : '❌ ไม่มี'));
    Logger.log('AES_SECRET_KEY: ' + (props.getProperty('AES_SECRET_KEY') ? '✅' : '❌ ไม่มี'));
  } catch(e) {
    Logger.log('Props ERROR: ' + e.message);
  }
}
function debugNow() {
  try {
    HtmlService.createHtmlOutputFromFile('index');
    Logger.log('✅ index.html พบ');
  } catch(e) {
    Logger.log('❌ index error: ' + e.message);
  }
  
  var p = PropertiesService.getScriptProperties();
  Logger.log('SHEET_ID = ' + (p.getProperty('SHEET_ID') || 'ไม่มี'));
  Logger.log('DRIVE_FOLDER_ID = ' + (p.getProperty('DRIVE_FOLDER_ID') || 'ไม่มี'));
  Logger.log('AES_SECRET_KEY = ' + (p.getProperty('AES_SECRET_KEY') ? 'มี' : 'ไม่มี'));
}
