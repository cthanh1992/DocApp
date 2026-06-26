/****************************************************************
 * AUTH.GS - ĐĂNG KÝ / ĐĂNG NHẬP / QUÊN MẬT KHẨU / SESSION
 * Toàn bộ hàm bắt đầu bằng "api_" được gọi từ giao diện qua
 * google.script.run (xem file JS_Auth.html)
 ****************************************************************/

/**
 * Khóa xử lý đồng thời để tránh 2 request cùng sửa Sheet 1 lúc
 * gây trùng/lệch dữ liệu (race condition) khi nhiều người dùng
 * thao tác cùng lúc.
 */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  const ok = lock.tryLock(10000); // chờ tối đa 10s
  if (!ok) {
    return { success: false, message: 'Hệ thống đang xử lý yêu cầu khác, vui lòng thử lại sau giây lát.' };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/****************************************************************
 * BƯỚC 1: ĐĂNG KÝ - GỬI TOKEN XÁC NHẬN
 ****************************************************************/
function api_requestRegister(email) {
  return withLock_(() => {
    email = normalizeEmail_(email);

    if (!email) {
      return { success: false, message: 'Vui lòng nhập email.' };
    }
    if (!isCompanyEmail_(email)) {
      return { success: false, message: 'Chỉ chấp nhận email công ty (@' + CONFIG.COMPANY_DOMAIN + ').' };
    }

    const sheet = getSheet_(CONFIG.SHEET_USERS);
    const users = readSheetAsObjects_(CONFIG.SHEET_USERS);
    const existing = users.find(u => normalizeEmail_(u.Email) === email);

    if (existing && existing.Status === 'ACTIVE') {
      return { success: false, message: 'Email này đã đăng ký tài khoản. Vui lòng đăng nhập hoặc dùng chức năng quên mật khẩu.' };
    }

    const otp = generateOtp_();
    const expiry = nowPlusMinutes_(CONFIG.TOKEN_EXPIRY_MINUTES);

    if (existing) {
      // Email đã từng bắt đầu đăng ký nhưng chưa hoàn tất -> cập nhật lại token
      sheet.getRange(existing._row, headerIndex_(sheet, 'RegToken') + 1).setValue(otp);
      sheet.getRange(existing._row, headerIndex_(sheet, 'RegTokenExpiry') + 1).setValue(expiry);
      sheet.getRange(existing._row, headerIndex_(sheet, 'Status') + 1).setValue('PENDING');
    } else {
      // Thêm dòng user mới ở trạng thái PENDING (chưa có mật khẩu)
      appendUserRow_(sheet, {
        Email: email,
        PasswordHash: '',
        Salt: '',
        Status: 'PENDING',
        RegToken: otp,
        RegTokenExpiry: expiry,
        SessionToken: '',
        SessionExpiry: '',
        ResetToken: '',
        ResetTokenExpiry: '',
        CreatedAt: new Date()
      });
    }

    sendMail_(email, 'Mã xác nhận đăng ký tài khoản',
      mailTemplate_('Xác nhận đăng ký tài khoản',
        `Bạn (hoặc ai đó dùng email của bạn) vừa yêu cầu đăng ký tài khoản truy cập hệ thống tài liệu nội bộ của <b>${CONFIG.COMPANY_NAME}</b>.<br><br>
         Mã xác nhận của bạn là:<br>
         <div style="font-size:28px;font-weight:bold;letter-spacing:4px;background:#f1f3f5;padding:12px 20px;border-radius:8px;display:inline-block;margin:10px 0;">${otp}</div><br>
         Mã có hiệu lực trong <b>${CONFIG.TOKEN_EXPIRY_MINUTES} phút</b>. Nếu bạn không thực hiện yêu cầu này, vui lòng bỏ qua email này.`
      )
    );

    return { success: true, message: 'Đã gửi mã xác nhận đến email của bạn. Vui lòng kiểm tra hộp thư.' };
  });
}

/****************************************************************
 * BƯỚC 2: XÁC NHẬN TOKEN + ĐẶT MẬT KHẨU -> HOÀN TẤT ĐĂNG KÝ
 ****************************************************************/
function api_completeRegister(email, otp, password) {
  return withLock_(() => {
    email = normalizeEmail_(email);
    otp = String(otp || '').trim();

    if (!email || !otp || !password) {
      return { success: false, message: 'Vui lòng nhập đầy đủ thông tin.' };
    }
    if (password.length < 8) {
      return { success: false, message: 'Mật khẩu cần tối thiểu 8 ký tự.' };
    }

    const sheet = getSheet_(CONFIG.SHEET_USERS);
    const users = readSheetAsObjects_(CONFIG.SHEET_USERS);
    const user = users.find(u => normalizeEmail_(u.Email) === email);

    if (!user) {
      return { success: false, message: 'Không tìm thấy yêu cầu đăng ký cho email này.' };
    }
    if (user.Status === 'ACTIVE') {
      return { success: false, message: 'Email này đã được đăng ký trước đó.' };
    }
    if (String(user.RegToken) !== otp) {
      return { success: false, message: 'Mã xác nhận không đúng.' };
    }
    if (isExpired_(user.RegTokenExpiry)) {
      return { success: false, message: 'Mã xác nhận đã hết hạn. Vui lòng đăng ký lại để nhận mã mới.' };
    }

    const salt = generateToken_(16);
    const hash = hashPassword_(password, salt);

    setRow_(sheet, user._row, {
      PasswordHash: hash,
      Salt: salt,
      Status: 'ACTIVE',
      RegToken: '',
      RegTokenExpiry: ''
    });

    sendMail_(email, 'Đăng ký tài khoản thành công',
      mailTemplate_('Đăng ký thành công',
        `Tài khoản của bạn tại hệ thống tài liệu nội bộ <b>${CONFIG.COMPANY_NAME}</b> đã được tạo thành công với email: <b>${email}</b>.<br><br>
         Bạn có thể đăng nhập ngay bây giờ bằng mật khẩu bạn đã đặt.<br><br>
         Nếu bạn không thực hiện việc đăng ký này, vui lòng liên hệ bộ phận IT ngay.`
      )
    );

    return { success: true, message: 'Đăng ký thành công! Bạn có thể đăng nhập ngay.' };
  });
}

/****************************************************************
 * ĐĂNG NHẬP
 ****************************************************************/
function api_login(email, password, remember) {
  return withLock_(() => {
    email = normalizeEmail_(email);

    if (!email || !password) {
      return { success: false, message: 'Vui lòng nhập email và mật khẩu.' };
    }

    const sheet = getSheet_(CONFIG.SHEET_USERS);
    const users = readSheetAsObjects_(CONFIG.SHEET_USERS);
    const user = users.find(u => normalizeEmail_(u.Email) === email);

    if (!user || user.Status !== 'ACTIVE') {
      return { success: false, message: 'Email hoặc mật khẩu không đúng.' };
    }

    const hash = hashPassword_(password, user.Salt);
    if (hash !== user.PasswordHash) {
      return { success: false, message: 'Email hoặc mật khẩu không đúng.' };
    }

    const sessionToken = generateToken_(40);
    const sessionExpiry = nowPlusDays_(CONFIG.SESSION_EXPIRY_DAYS);

    setRow_(sheet, user._row, {
      SessionToken: sessionToken,
      SessionExpiry: sessionExpiry
    });

    return {
      success: true,
      message: 'Đăng nhập thành công.',
      sessionToken: sessionToken,
      email: email
    };
  });
}

/****************************************************************
 * KIỂM TRA SESSION (gọi mỗi khi load lại trang để tự đăng nhập)
 ****************************************************************/
function api_checkSession(email, sessionToken) {
  email = normalizeEmail_(email);
  if (!email || !sessionToken) {
    return { valid: false };
  }

  const users = readSheetAsObjects_(CONFIG.SHEET_USERS);
  const user = users.find(u => normalizeEmail_(u.Email) === email);

  if (!user || user.Status !== 'ACTIVE') {
    return { valid: false };
  }
  if (String(user.SessionToken) !== String(sessionToken)) {
    return { valid: false };
  }
  if (isExpired_(user.SessionExpiry)) {
    return { valid: false };
  }

  return { valid: true, email: email };
}

/****************************************************************
 * ĐĂNG XUẤT - xoá session token trong Sheet
 ****************************************************************/
function api_logout(email, sessionToken) {
  return withLock_(() => {
    email = normalizeEmail_(email);
    const sheet = getSheet_(CONFIG.SHEET_USERS);
    const users = readSheetAsObjects_(CONFIG.SHEET_USERS);
    const user = users.find(u => normalizeEmail_(u.Email) === email);

    if (user && String(user.SessionToken) === String(sessionToken)) {
      setRow_(sheet, user._row, { SessionToken: '', SessionExpiry: '' });
    }
    return { success: true };
  });
}

/****************************************************************
 * QUÊN MẬT KHẨU - BƯỚC 1: GỬI TOKEN
 ****************************************************************/
function api_requestPasswordReset(email) {
  return withLock_(() => {
    email = normalizeEmail_(email);
    if (!email) {
      return { success: false, message: 'Vui lòng nhập email.' };
    }

    const sheet = getSheet_(CONFIG.SHEET_USERS);
    const users = readSheetAsObjects_(CONFIG.SHEET_USERS);
    const user = users.find(u => normalizeEmail_(u.Email) === email);

    // Thông báo chung để tránh lộ thông tin email nào tồn tại trong hệ thống
    const genericMessage = 'Nếu email tồn tại trong hệ thống, mã đặt lại mật khẩu đã được gửi đến email đó.';

    if (!user || user.Status !== 'ACTIVE') {
      return { success: true, message: genericMessage };
    }

    const otp = generateOtp_();
    const expiry = nowPlusMinutes_(CONFIG.TOKEN_EXPIRY_MINUTES);

    setRow_(sheet, user._row, {
      ResetToken: otp,
      ResetTokenExpiry: expiry
    });

    sendMail_(email, 'Mã đặt lại mật khẩu',
      mailTemplate_('Đặt lại mật khẩu',
        `Bạn vừa yêu cầu đặt lại mật khẩu cho tài khoản <b>${email}</b> tại hệ thống tài liệu nội bộ ${CONFIG.COMPANY_NAME}.<br><br>
         Mã đặt lại mật khẩu của bạn là:<br>
         <div style="font-size:28px;font-weight:bold;letter-spacing:4px;background:#f1f3f5;padding:12px 20px;border-radius:8px;display:inline-block;margin:10px 0;">${otp}</div><br>
         Mã có hiệu lực trong <b>${CONFIG.TOKEN_EXPIRY_MINUTES} phút</b>. Nếu bạn không thực hiện yêu cầu này, vui lòng bỏ qua email này hoặc liên hệ IT nếu nghi ngờ có người khác cố truy cập tài khoản của bạn.`
      )
    );

    return { success: true, message: genericMessage };
  });
}

/****************************************************************
 * QUÊN MẬT KHẨU - BƯỚC 2: XÁC NHẬN TOKEN + ĐẶT MẬT KHẨU MỚI
 ****************************************************************/
function api_completePasswordReset(email, otp, newPassword) {
  return withLock_(() => {
    email = normalizeEmail_(email);
    otp = String(otp || '').trim();

    if (!email || !otp || !newPassword) {
      return { success: false, message: 'Vui lòng nhập đầy đủ thông tin.' };
    }
    if (newPassword.length < 8) {
      return { success: false, message: 'Mật khẩu cần tối thiểu 8 ký tự.' };
    }

    const sheet = getSheet_(CONFIG.SHEET_USERS);
    const users = readSheetAsObjects_(CONFIG.SHEET_USERS);
    const user = users.find(u => normalizeEmail_(u.Email) === email);

    if (!user || user.Status !== 'ACTIVE') {
      return { success: false, message: 'Yêu cầu không hợp lệ.' };
    }
    if (String(user.ResetToken) !== otp || !otp) {
      return { success: false, message: 'Mã xác nhận không đúng.' };
    }
    if (isExpired_(user.ResetTokenExpiry)) {
      return { success: false, message: 'Mã xác nhận đã hết hạn. Vui lòng yêu cầu lại.' };
    }

    const salt = generateToken_(16);
    const hash = hashPassword_(newPassword, salt);

    setRow_(sheet, user._row, {
      PasswordHash: hash,
      Salt: salt,
      ResetToken: '',
      ResetTokenExpiry: '',
      SessionToken: '', // Bắt đăng nhập lại trên mọi thiết bị sau khi đổi mật khẩu
      SessionExpiry: ''
    });

    sendMail_(email, 'Mật khẩu đã được thay đổi',
      mailTemplate_('Mật khẩu đã thay đổi',
        `Mật khẩu cho tài khoản <b>${email}</b> vừa được thay đổi thành công.<br><br>
         Nếu bạn không thực hiện thay đổi này, vui lòng liên hệ bộ phận IT ngay lập tức.`
      )
    );

    return { success: true, message: 'Đổi mật khẩu thành công. Vui lòng đăng nhập lại.' };
  });
}

/****************************************************************
 * HELPER GHI/ĐỌC SHEET THEO TÊN CỘT (để code rõ ràng, dễ bảo trì)
 ****************************************************************/
function headerIndex_(sheet, columnName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(columnName);
  if (idx === -1) {
    throw new Error('Không tìm thấy cột "' + columnName + '" trong sheet ' + sheet.getName());
  }
  return idx;
}

function appendUserRow_(sheet, rowObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => (h in rowObj ? rowObj[h] : ''));
  sheet.appendRow(row);
}

function setRow_(sheet, rowNumber, fieldsObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Object.keys(fieldsObj).forEach(key => {
    const colIdx = headers.indexOf(key);
    if (colIdx !== -1) {
      sheet.getRange(rowNumber, colIdx + 1).setValue(fieldsObj[key]);
    }
  });
}

/****************************************************************
 * GỬI EMAIL
 ****************************************************************/
function sendMail_(toEmail, subject, htmlBody) {
  MailApp.sendEmail({
    to: toEmail,
    subject: '[' + CONFIG.COMPANY_NAME + '] ' + subject,
    htmlBody: htmlBody,
    name: CONFIG.COMPANY_NAME
  });
}

function mailTemplate_(title, contentHtml) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px;border:1px solid #e9ecef;border-radius:12px;">
      <h2 style="color:#1a1a1a;margin-top:0;">${title}</h2>
      <p style="color:#444;font-size:15px;line-height:1.6;">${contentHtml}</p>
      <hr style="border:none;border-top:1px solid #e9ecef;margin:24px 0;">
      <p style="color:#999;font-size:12px;">Email này được gửi tự động từ hệ thống tài liệu nội bộ ${CONFIG.COMPANY_NAME}. Vui lòng không trả lời email này.</p>
    </div>
  `;
}
