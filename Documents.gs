/****************************************************************
 * DOCUMENTS.GS - LẤY DANH MỤC FOLDER/TÀI LIỆU THEO PHÂN QUYỀN
 ****************************************************************/

/**
 * Trả về danh sách folder mà email này được phép truy cập,
 * dựa trên sheet Permissions (cột Email, FolderId/FolderKey).
 * Thông tin tên + ID folder thật lấy từ sheet Folders.
 */
function api_getAccessibleFolders(email, sessionToken) {
  const session = api_checkSession(email, sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.', folders: [] };
  }
  email = normalizeEmail_(email);

  const permissions = readSheetAsObjects_(CONFIG.SHEET_PERMISSIONS);
  const allFolders = readSheetAsObjects_(CONFIG.SHEET_FOLDERS);

  // Lấy danh sách FolderKey mà user này được cấp quyền
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
 * Trả về danh sách file (tên, id, loại, link xem/tải) trong 1 folder,
 * SAU KHI đã kiểm tra email có quyền truy cập folder đó hay không.
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
    const it = folder.getFiles();
    while (it.hasNext()) {
      const file = it.next();
      files.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        sizeBytes: file.getSize(),
        updatedAt: file.getLastUpdated(),
        // Link xem trực tiếp (embed) trên Drive - hỗ trợ xem, in, tải tuỳ quyền share của file
        previewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/preview',
        // Link mở trực tiếp trên tab mới (đầy đủ tính năng xem/tải/in của Drive viewer)
        openUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
        iconUrl: file.getThumbnail() ? null : null
      });
    }

    // Sắp xếp theo tên
    files.sort((a, b) => a.name.localeCompare(b.name, 'vi'));

    return { success: true, folderName: folderInfo.FolderName, files: files };
  } catch (err) {
    return { success: false, message: 'Không thể đọc danh mục này. Vui lòng kiểm tra lại Folder ID hoặc quyền chia sẻ trên Drive. Chi tiết lỗi: ' + err.message, files: [] };
  }
}
