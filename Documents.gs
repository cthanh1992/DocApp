/****************************************************************
 * DOCUMENTS.GS - LẤY DANH MỤC FOLDER/TÀI LIỆU THEO PHÂN QUYỀN
 ****************************************************************/

/**
 * Trả về danh sách folder mà email này được phép truy cập
 */
function api_getAccessibleFolders(email, sessionToken) {
  const session = api_checkSession(email, sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.', folders: [] };
  }
  email = normalizeEmail_(email);

  const permissions = readSheetAsObjects_(CONFIG.SHEET_PERMISSIONS);
  const allFolders = readSheetAsObjects_(CONFIG.SHEET_FOLDERS);

  const allowedKeys = permissions
    .filter(p => normalizeEmail_(p.Email) === email)
    .map(p => String(p.FolderKey).trim());

  const folders = allFolders
    .filter(f => allowedKeys.includes(String(f.FolderKey).trim()))
    .map(f => ({
      key: f.FolderKey,
      name: f.FolderName,
      folderId: f.FolderId,
      description: f.Description || ''
    }));
  return { success: true, folders: folders };
}

/**
 * Trả về danh sách toàn bộ file trong folder và các folder con
 */
function api_getFilesInFolder(email, sessionToken, folderKey) {
  const session = api_checkSession(email, sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.', files: [] };
  }
  email = normalizeEmail_(email);

  // Kiểm tra quyền truy cập folderKey này
  const permissions = readSheetAsObjects_(CONFIG.SHEET_PERMISSIONS);
  const hasAccess = permissions.some(p =>
    normalizeEmail_(p.Email) === email && String(p.FolderKey).trim() === String(folderKey).trim()
  );
  if (!hasAccess) {
    return { success: false, message: 'Bạn không có quyền truy cập danh mục này.', files: [] };
  }

  const allFolders = readSheetAsObjects_(CONFIG.SHEET_FOLDERS);
  const folderInfo = allFolders.find(f => String(f.FolderKey).trim() === String(folderKey).trim());
  if (!folderInfo) {
    return { success: false, message: 'Không tìm thấy danh mục.', files: [] };
  }

  try {
    const folder = DriveApp.getFolderById(folderInfo.FolderId);
    const files = [];
    
    // Thực hiện quét đệ quy lấy toàn bộ file ở thư mục gốc và thư mục con
    getFilesRecursive_(folder, files);

    // Sắp xếp danh sách file theo thứ tự bảng chữ cái câu chữ tiếng Việt
    files.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
    return { success: true, folderName: folderInfo.FolderName, files: files };
  } catch (err) {
    return { success: false, message: 'Không thể đọc danh mục này. Vui lòng kiểm tra lại Folder ID hoặc quyền chia sẻ trên Drive. Chi tiết lỗi: ' + err.message, files: [] };
  }
}
/**
 * Hàm phụ trợ: Duyệt đệ quy để gom tất cả file từ thư mục cha vào các thư mục con
 */
function getFilesRecursive_(folder, files) {
  // 1. Quét tất cả file ở cấp thư mục hiện tại
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    
    // Xử lý an toàn: Lấy dung lượng file, bỏ qua lỗi nếu là file Google Docs/Sheets
    let fileSize = 0;
    try {
      fileSize = file.getSize();
    } catch (e) {
      fileSize = 0; 
    }

    files.push({
      id: file.getId(),
      name: file.getName(),
      mimeType: file.getMimeType(),
      sizeBytes: fileSize, // Sử dụng biến an toàn
      updatedAt: file.getLastUpdated().toISOString(), // Đã sửa lỗi: Thêm .toISOString()
      previewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/preview',
      openUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
      iconUrl: null
    });
  }
  
  // 2. Tìm các thư mục con và tiếp tục đào sâu để quét file bên trong chúng
  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    const subFolder = folderIterator.next();
    getFilesRecursive_(subFolder, files);
  }
}
