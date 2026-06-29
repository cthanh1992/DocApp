function api_getAccessibleFolders(email, sessionToken) {
  const session = api_checkSession(email, sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.', folders: [] };
  }
  email = normalizeEmail_(email);

  const permissions = readSheetAsObjects_(CONFIG.SHEET_PERMISSIONS);
  const allFolders = readSheetAsObjects_(CONFIG.SHEET_FOLDERS);

  // --- CẬP NHẬT: Lấy danh sách FolderKey từ 1 dòng, cắt bằng dấu phẩy ---
  let allowedKeys = [];
  const userPermissionRow = permissions.find(p => normalizeEmail_(p.Email) === email);
  if (userPermissionRow && userPermissionRow.FolderKey) {
    allowedKeys = String(userPermissionRow.FolderKey)
      .split(',')                     // Cắt chuỗi theo dấu phẩy
      .map(k => k.trim())             // Loại bỏ khoảng trắng 2 đầu mỗi key
      .filter(k => k !== '');         // Bỏ qua các giá trị rỗng (ví dụ người dùng gõ dư dấu phẩy)
  }
  // ----------------------------------------------------------------------

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

function api_getFilesInFolder(email, sessionToken, folderKey, subFolderId) {
  const session = api_checkSession(email, sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.', files: [] };
  }
  email = normalizeEmail_(email);

  const permissions = readSheetAsObjects_(CONFIG.SHEET_PERMISSIONS);
  
  // --- CẬP NHẬT: Kiểm tra quyền truy cập theo mảng sau khi cắt dấu phẩy ---
  let hasAccess = false;
  const userPermissionRow = permissions.find(p => normalizeEmail_(p.Email) === email);
  if (userPermissionRow && userPermissionRow.FolderKey) {
    const allowedKeys = String(userPermissionRow.FolderKey)
      .split(',')
      .map(k => k.trim());
    hasAccess = allowedKeys.includes(String(folderKey).trim());
  }
  // ------------------------------------------------------------------------

  if (!hasAccess) {
    return { success: false, message: 'Bạn không có quyền truy cập danh mục này.', files: [] };
  }

  const allFolders = readSheetAsObjects_(CONFIG.SHEET_FOLDERS);
  const folderInfo = allFolders.find(f => String(f.FolderKey).trim() === String(folderKey).trim());
  if (!folderInfo) {
    return { success: false, message: 'Không tìm thấy danh mục.', files: [] };
  }

  try {
    const targetFolderId = subFolderId || folderInfo.FolderId;
    const folder = DriveApp.getFolderById(targetFolderId);
    const files = [];
    
    // 1. Quét Thư mục con
    const folderIterator = folder.getFolders();
    while (folderIterator.hasNext()) {
      const subFolder = folderIterator.next();
      files.push({
        id: subFolder.getId(),
        name: subFolder.getName(),
        mimeType: 'application/vnd.google-apps.folder',
        sizeBytes: 0,
        updatedAt: subFolder.getLastUpdated().toISOString(),
        previewUrl: '',
        openUrl: subFolder.getUrl(),
        iconUrl: null
      });
    }

    // 2. Quét File
    const fileIterator = folder.getFiles();
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      let fileSize = 0;
      try { fileSize = file.getSize(); } catch (e) { fileSize = 0; }

      files.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        sizeBytes: fileSize,
        updatedAt: file.getLastUpdated().toISOString(),
        previewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/preview',
        openUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
        iconUrl: null
      });
    }

    files.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
    const currentFolderName = subFolderId ? folder.getName() : folderInfo.FolderName;
    return { success: true, folderName: currentFolderName, files: files };
  } catch (err) {
    return { success: false, message: 'Không thể đọc danh mục này. Lỗi: ' + err.message, files: [] };
  }
}
