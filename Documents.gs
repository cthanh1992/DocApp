/****************************************************************
 * DOCUMENTS.GS - LẤY DANH MỤC FOLDER/TÀI LIỆU THEO PHÂN QUYỀN
 ****************************************************************/

/**
 * 1. Hàm lấy quyền ĐỌC: Được cấp quyền thư mục mẹ -> tự động có quyền đọc thư mục con
 */
function getReadAccessKeys_(email, allFolders) {
  const permissions = readSheetAsObjects_(CONFIG.SHEET_PERMISSIONS);
  
  let baseKeys = [];
  const userPermissionRow = permissions.find(p => normalizeEmail_(p.Email) === normalizeEmail_(email));
  if (userPermissionRow && userPermissionRow.FolderKey) {
    baseKeys = String(userPermissionRow.FolderKey)
      .split(',')
      .map(k => k.trim())
      .filter(k => k !== '');
  }

  let expandedKeys = new Set(baseKeys);
  let added = true;
  while (added) {
    added = false;
    allFolders.forEach(f => {
      const parentKey = f.ParentKey ? String(f.ParentKey).trim() : '';
      const folderKey = f.FolderKey ? String(f.FolderKey).trim() : '';
      // Nếu có quyền đọc thư mục cha -> tự động cấp quyền đọc thư mục con
      if (parentKey && expandedKeys.has(parentKey) && !expandedKeys.has(folderKey)) {
        expandedKeys.add(folderKey);
        added = true;
      }
    });
  }
  return Array.from(expandedKeys);
}

/**
 * 2. Hàm lấy quyền HIỂN THỊ: Được cấp quyền thư mục con -> tự động vẽ thư mục mẹ để tạo hình cây
 */
function getDisplayFolderKeys_(readAccessKeys, allFolders) {
  let displayKeys = new Set(readAccessKeys);
  let added = true;
  while (added) {
    added = false;
    allFolders.forEach(f => {
      const parentKey = f.ParentKey ? String(f.ParentKey).trim() : '';
      const folderKey = f.FolderKey ? String(f.FolderKey).trim() : '';
      // Nếu có hiển thị thư mục con -> bắt buộc phải hiển thị thư mục cha chứa nó
      if (parentKey && displayKeys.has(folderKey) && !displayKeys.has(parentKey)) {
        displayKeys.add(parentKey);
        added = true;
      }
    });
  }
  return Array.from(displayKeys);
}

/**
 * Trả về danh sách folder để vẽ lên giao diện
 */
function api_getAccessibleFolders(email, sessionToken) {
  const session = api_checkSession(email, sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.', folders: [] };
  }
  email = normalizeEmail_(email);

  const allFolders = readSheetAsObjects_(CONFIG.SHEET_FOLDERS);
  const readKeys = getReadAccessKeys_(email, allFolders);
  // Dùng displayKeys để quyết định xem những thư mục nào sẽ hiện trên thanh bên trái
  const displayKeys = getDisplayFolderKeys_(readKeys, allFolders);

  const folders = allFolders
    .filter(f => displayKeys.includes(String(f.FolderKey).trim()))
    .map(f => ({
      key: f.FolderKey,
      name: f.FolderName,
      folderId: f.FolderId,
      description: f.Description || '',
      parentKey: f.ParentKey ? String(f.ParentKey).trim() : ''
    }));
  return { success: true, folders: folders };
}

/**
 * Trả về danh sách file và thư mục con khi click vào
 */
function api_getFilesInFolder(email, sessionToken, folderKey, subFolderId) {
  const session = api_checkSession(email, sessionToken);
  if (!session.valid) {
    return { success: false, message: 'Phiên đăng nhập không hợp lệ, vui lòng đăng nhập lại.', files: [] };
  }
  email = normalizeEmail_(email);

  const allFolders = readSheetAsObjects_(CONFIG.SHEET_FOLDERS);
  // Chú ý: Tại đây chỉ dùng readKeys. Nếu cố click vào thư mục mẹ (chỉ có displayKeys), sẽ bị chặn!
  const readKeys = getReadAccessKeys_(email, allFolders);

  const hasAccess = readKeys.includes(String(folderKey).trim());
  if (!hasAccess) {
    return { success: false, message: 'Thư mục này chỉ hiển thị để xem cấu trúc, bạn không có quyền xem tài liệu bên trong.', files: [] };
  }

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
