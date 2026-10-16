/**
 * Theme_Image.gs — ảnh nền theme lấy qua Apps Script thay vì app tải thẳng từ Drive.
 *
 * Vì sao: trình duyệt và APK bị Google chặn tải file Drive trực tiếp (CORS), còn link Drive
 * chỉ hiện trên máy đang đăng nhập Google có quyền xem file. Script chạy bằng quyền chủ file
 * nên đọc được ảnh, rồi trả về cho app dạng data URL.
 *
 * Cài đặt (project VTD):
 *   1. Thêm file này.
 *   2. Thêm 1 dòng vào map trong vtdApp_apiDispatch_:   themeImage: vtdApp_themeImage,
 *   3. Deploy → Manage deployments → Edit → New version.
 *
 * Tự cập nhật khi đổi ảnh:
 *   - Đổi sang link ảnh khác trong Admin → Giao diện → "Áp dụng cho mọi máy": fileId đổi → các máy tải ảnh mới.
 *   - Thay nội dung chính file Drive cũ (Quản lý phiên bản → Tải lên phiên bản mới): ngày sửa file đổi → các máy tải lại.
 * App gửi kèm phiên bản đang có; ảnh không đổi thì script chỉ trả {unchanged: true}, không gửi lại ảnh.
 */

const THEME_IMAGE_MAX_BYTES = 8 * 1024 * 1024;   // ảnh nền nên nén dưới 1 MB; quá 8 MB thì từ chối

function vtdApp_themeImage(params) {
  params = params || {};
  const fileId = String(params.fileId || "").trim();
  if (!/^[\w-]{20,}$/.test(fileId)) return vtdApp_fail_("Thiếu mã file ảnh nền.");

  // Action công khai (màn đăng nhập cũng cần ảnh) → chỉ cho tải ảnh đang nằm trong theme hệ thống
  const theme = vtdApp_publicSystemConfig_().themeConfig || {};
  if (JSON.stringify(theme).indexOf(fileId) < 0) return vtdApp_fail_("Ảnh này không thuộc theme hệ thống.");

  let file;
  try { file = DriveApp.getFileById(fileId); }
  catch (err) { return vtdApp_fail_("Script không mở được file ảnh trên Drive. Kiểm tra tài khoản chạy script có quyền xem file."); }

  const size = file.getSize();
  const version = fileId + ":" + file.getLastUpdated().getTime() + ":" + size;
  if (String(params.knownVersion || "") === version) return vtdApp_ok_({unchanged: true, version: version});
  if (size > THEME_IMAGE_MAX_BYTES) return vtdApp_fail_("Ảnh nền " + Math.round(size / 1048576) + " MB, quá lớn. Nén ảnh dưới 1 MB rồi tải lên lại.");

  const blob = file.getBlob();
  const mime = String(blob.getContentType() || file.getMimeType() || "");
  if (mime.indexOf("image/") !== 0) return vtdApp_fail_("File Drive không phải ảnh (" + mime + ").");
  return vtdApp_ok_({
    version: version,
    mimeType: mime,
    bytes: size,
    dataUrl: "data:" + mime + ";base64," + Utilities.base64Encode(blob.getBytes())
  });
}
