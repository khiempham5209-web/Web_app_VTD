/**
 * CanhBaoLoi_VNPT.gs — cùng project với VanHanh_VNPT.gs và API_VNPT.gs.
 * Không tự gửi mail khi dán code. Mail chạy khi hàm vận hành lỗi; cấp quyền bằng vnptAuthorizeAlerts().
 * Cùng 1 lỗi chỉ gửi lại sau cooldownMinutes.
 */
const VNPT_ALERT = {to: 'khiempham5209@gmail.com', cooldownMinutes: 60};

function vnptNotifyError(context, error) {
  const msg = String(error && error.message || error);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, msg, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
  const prop = 'VNPT_ALERT_' + digest, props = PropertiesService.getScriptProperties();
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10000)) { console.error('Chưa gửi cảnh báo do khóa bận: ' + msg); return; }
  try {
    const now = Date.now(), previous = Number(props.getProperty(prop) || 0);
    if (previous && now - previous < VNPT_ALERT.cooldownMinutes * 60000) return;
    if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Đã hết quota gửi email cảnh báo.');
    const ss = SpreadsheetApp.openById(VNPT.id);
    const time = Utilities.formatDate(new Date(now), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    MailApp.sendEmail({
      to: VNPT_ALERT.to,
      subject: '[VNPT] Lỗi vận hành: ' + String(context).slice(0, 100),
      body: [
        'File: ' + ss.getName(),
        'Link: ' + ss.getUrl(),
        'Thời điểm: ' + time,
        'Luồng/hàm: ' + context,
        'Lỗi: ' + msg,
        '',
        'Luồng đã dừng tại bước gặp lỗi. Những bước ghi trước đó có thể đã hoàn tất; không có cơ chế rollback toàn bộ.',
        'Sửa đúng header/dữ liệu rồi chạy lại luồng tương ứng. Số liệu được tính tuyệt đối nên chạy lại không cộng trùng.',
        '',
        'Chi tiết kỹ thuật:',
        String(error && error.stack || error)
      ].join('\n')
    });
    props.setProperty(prop, String(now));
    // Dọn dấu chống trùng cũ; không xóa property của API/vận hành.
    const old = props.getProperties();
    Object.keys(old).filter(k => k.indexOf('VNPT_ALERT_') === 0 && Number(old[k]) < now - 7 * 86400000)
      .forEach(k => props.deleteProperty(k));
  } finally { lock.releaseLock(); }
}

/** Xin quyền MailApp, không gửi email. */
function vnptAuthorizeAlerts() {
  console.log('Email nhận cảnh báo: ' + VNPT_ALERT.to + ' | Quota còn lại: ' + MailApp.getRemainingDailyQuota());
}

/** Chỉ chạy thủ công khi muốn nhận một email test. */
function vnptTestAlertEmail() {
  vnptNotifyError('TEST THỦ CÔNG', new Error('Email kiểm tra cảnh báo VNPT — không phải lỗi dữ liệu tồn kho.'));
}
