/****************************************************************
 * SETUP.GS - CHẠY 1 LẦN DUY NHẤT ĐỂ KHỞI TẠO CẤU TRÚC GOOGLE SHEET
 *
 * CÁCH DÙNG:
 * 1. Tạo 1 Google Sheet mới (trống), copy ID của sheet (trong URL,
 *    đoạn giữa /d/ và /edit) vào CONFIG.SHEET_ID trong file Code.gs
 * 2. Trong Apps Script Editor, chọn hàm "setupSheets" ở thanh công cụ trên
 * 3. Bấm nút Run (▶) - lần đầu sẽ yêu cầu cấp quyền, bấm Cho phép (Allow)
 * 4. Mở Google Sheet lên kiểm tra đã có đủ 3 tab: Users, Permissions, Folders
 ****************************************************************/
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // --- Sheet Users ---
  let usersSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(CONFIG.SHEET_USERS);
  }
  const userHeaders = [
    'Email', 'PasswordHash', 'Salt', 'Status',
    'RegToken', 'RegTokenExpiry',
    'SessionToken', 'SessionExpiry',
    'ResetToken', 'ResetTokenExpiry',
    'CreatedAt'
  ];
  setHeaderIfEmpty_(usersSheet, userHeaders);

  // --- Sheet Folders (khai báo các folder Drive sẽ hiển thị trong app) ---
  let foldersSheet = ss.getSheetByName(CONFIG.SHEET_FOLDERS);
  if (!foldersSheet) {
    foldersSheet = ss.insertSheet(CONFIG.SHEET_FOLDERS);
  }
  const folderHeaders = ['FolderKey', 'FolderName', 'FolderId', 'Description'];
  setHeaderIfEmpty_(foldersSheet, folderHeaders);
  if (foldersSheet.getLastRow() < 2) {
    foldersSheet.appendRow(['hr', 'Tài liệu Nhân sự', 'DÁN_FOLDER_ID_DRIVE_VÀO_ĐÂY', 'Quy định, chính sách nhân sự']);
    foldersSheet.appendRow(['finance', 'Tài liệu Tài chính', 'DÁN_FOLDER_ID_DRIVE_VÀO_ĐÂY', 'Báo cáo, hoá đơn, ngân sách']);
  }

  // --- Sheet Permissions (phân quyền email nào xem folder nào) ---
  let permSheet = ss.getSheetByName(CONFIG.SHEET_PERMISSIONS);
  if (!permSheet) {
    permSheet = ss.insertSheet(CONFIG.SHEET_PERMISSIONS);
  }
  const permHeaders = ['Email', 'FolderKey'];
  setHeaderIfEmpty_(permSheet, permHeaders);
  if (permSheet.getLastRow() < 2) {
    permSheet.appendRow(['vidu@' + CONFIG.COMPANY_DOMAIN, 'hr']);
  }

  // Xoá sheet mặc định "Sheet1" nếu nó trống và không phải sheet ta vừa tạo
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.flush();
  Logger.log('Đã khởi tạo xong cấu trúc Sheet. Vui lòng mở Google Sheet để: ' +
    '1) Điền FolderId thật vào sheet "Folders", ' +
    '2) Khai báo phân quyền email vào sheet "Permissions".');
}

function setHeaderIfEmpty_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f5');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Hàm phụ trợ: dùng để admin lấy nhanh ID của 1 folder Drive
 * khi đang xem folder đó trên trình duyệt (copy URL folder, dán vào đây).
 * Chạy thử trong Apps Script Editor (Run), xem kết quả trong Logger.
 */
function helper_getFolderIdFromUrl(folderUrl) {
  const match = folderUrl.match(/[-\w]{25,}/);
  Logger.log(match ? match[0] : 'Không tìm thấy ID trong URL.');
  return match ? match[0] : null;
}
