const CONFIG = {
  SHEET_ID: '1NJPMpHlHWBwwb7F2x1AuXflaqpw5A7lJYlGWVgQGRVw',
  COMPANY_DOMAIN: 'hoptrisummit.com',
  COMPANY_NAME: 'Hợp Trí Summit',
  SHEET_USERS: 'Users',
  SHEET_PERMISSIONS: 'Permissions',
  SHEET_FOLDERS: 'Folders',
  TOKEN_EXPIRY_MINUTES: 15,
  SESSION_EXPIRY_DAYS: 30,
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.COMPANY_NAME + ' - Tài liệu nội bộ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Không tìm thấy sheet: ' + sheetName + '. Vui lòng kiểm tra lại tên sheet hoặc chạy hàm setupSheets() trước.');
  }
  return sheet;
}

function readSheetAsObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .map((row, idx) => {
      const obj = { _row: idx + 2 };
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    // Fix: Kiểm tra cột đầu tiên thay vì gán chết cột Email
    .filter(obj => obj[headers[0]] !== undefined && String(obj[headers[0]]).trim() !== '');
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function isCompanyEmail_(email) {
  const domain = normalizeEmail_(email).split('@')[1];
  return domain === CONFIG.COMPANY_DOMAIN.toLowerCase();
}

function generateToken_(length) {
  length = length || 32;
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let token = '';
  const randomBytes = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  for (let i = 0; i < length; i++) {
    const idx = parseInt(randomBytes.substr(i * 2, 2), 16) % chars.length;
    token += chars.charAt(idx);
  }
  return token;
}

function generateOtp_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function hashPassword_(password, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + '::' + salt,
    Utilities.Charset.UTF_8
  );
  return raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function nowPlusMinutes_(minutes) {
  return new Date(Date.now() + minutes * 60 * 1000);
}

function nowPlusDays_(days) {
  return new Date(Date.now() + days * 24 * 60 * 60 * 1000);
}

function isExpired_(dateValue) {
  if (!dateValue) return true;
  const d = new Date(dateValue);
  return isNaN(d.getTime()) || d.getTime() < Date.now();
}
