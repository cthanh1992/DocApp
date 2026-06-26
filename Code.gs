/****************************************************************
 * CẤU HÌNH CHUNG - SỬA CÁC GIÁ TRỊ NÀY CHO PHÙ HỢP VỚI CÔNG TY BẠN
 ****************************************************************/
const CONFIG = {
  // ID của Google Sheet dùng làm database (lấy từ URL sheet, đoạn giữa /d/ và /edit)
  SHEET_ID: '1NJPMpHlHWBwwb7F2x1AuXflaqpw5A7lJYlGWVgQGRVw',

  // Tên miền email công ty (dùng để chặn đăng ký email ngoài công ty)
  // Ví dụ: nếu email công ty là user@mycompany.com thì điền 'mycompany.com'
  COMPANY_DOMAIN: 'hoptrisummit.com',

  // Tên công ty hiển thị trên giao diện và trong email
  COMPANY_NAME: 'Hợp Trí Summit',

  // Tên 2 sheet (tab) trong Google Sheet, không cần đổi nếu dùng đúng template đi kèm
  SHEET_USERS: 'Users',
  SHEET_PERMISSIONS: 'Permissions',
  SHEET_FOLDERS: 'Folders',

  // Thời hạn của token đăng ký / quên mật khẩu (phút)
  TOKEN_EXPIRY_MINUTES: 15,

  // Thời hạn của session đăng nhập (ngày) - "ghi nhớ đăng nhập"
  SESSION_EXPIRY_DAYS: 30,

  // Email gửi đi sẽ có tên hiển thị (From name) là tên công ty
  // (Apps Script sẽ gửi bằng địa chỉ Gmail của người sở hữu script)
};

/****************************************************************
 * ENTRY POINT - HIỂN THỊ TRANG WEB
 ****************************************************************/
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.COMPANY_NAME + ' - Tài liệu nội bộ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Cho phép include file HTML/CSS/JS riêng vào trong template chính
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/****************************************************************
 * TRUY CẬP SHEET - HELPER
 ****************************************************************/
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Không tìm thấy sheet: ' + sheetName + '. Vui lòng kiểm tra lại tên sheet hoặc chạy hàm setupSheets() trước.');
  }
  return sheet;
}

// Đọc toàn bộ dữ liệu của 1 sheet thành mảng object, dùng dòng đầu làm tên cột
function readSheetAsObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .map((row, idx) => {
      const obj = { _row: idx + 2 }; // số dòng thật trong sheet (idx+2 vì header ở dòng 1)
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .filter(obj => obj.Email && String(obj.Email).trim() !== ''); // bỏ dòng trống
}

/****************************************************************
 * UTILITY
 ****************************************************************/
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

// Sinh OTP 6 số (dùng cho token email - dễ nhập hơn token dài)
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
