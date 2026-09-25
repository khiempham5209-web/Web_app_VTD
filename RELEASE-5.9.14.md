# 5.9.14 Hàng đợi API, chống mất phản hồi và chống mail lỗi lặp

Nguyên nhân (log 25/09): server xử lý xong trong 0,5–2 giây nhưng máy chờ 17–35 giây rồi nhận 404 "Không tìm thấy trang" hoặc "Load failed". Lỗi nằm ở bước Google chuyển kết quả về máy (redirect sang script.googleusercontent.com), xảy ra khi một máy giữ 14–18 request song song ngay sau đăng nhập.

App (web + APK):
- Hàng đợi chung cho API VTD, Thao tác CT, Nhập TH: tối đa 4 request cùng lúc; request nền tối đa 3, báo cáo/admin tối đa 2. Luôn chừa chỗ cho lưu đơn, upload ảnh, đăng nhập và thao tác người dùng bấm. Timeout chỉ tính từ lúc request thật sự được gửi.
- Thứ tự: thao tác người dùng và lưu đơn → dữ liệu nhập đơn (Master SKU, khách hàng, cấu hình, đơn hôm nay, dashboard) → báo cáo và mục admin.
- Lệnh đọc bị mất phản hồi (404/HTML, Load failed, RESPONSE_LOST) tự gửi lại 1 lần sau 1,5 giây. Lệnh ghi (lưu đơn, upload ảnh, xóa, đổi pass...) không tự gửi lại.
- Cùng một lệnh đọc nền đang chạy thì lần gọi sau dùng chung kết quả.
- Báo tiến độ cache: bước trung gian gộp, tối đa 1 lần/phút mỗi loại dữ liệu; kết quả cuối done/error gửi ngay.
- Màn Cài đặt: hết vòng lặp gọi lại forgotPassRequests/queueErrors/lookupCacheStatus khi màn vẽ lại; lỗi thì chờ 60 giây mới tự tải lại.
- Queue ảnh: ảnh vừa ghi chưa kịp gắn vào đơn không còn bị dọn nhầm khi lưu đơn PN cùng lúc (lỗi readImage "parameter 1 is not of type 'Blob'"). Lượt sync khác đã upload ảnh thì bỏ qua, không báo lỗi. Ảnh thật sự mất có thông báo rõ cần bổ sung.

Server Code_chinh_VTD.gs:
- queueErrors: staff có quyền sync (Chạy lại sync, Chạy lệnh sync, xem sync/vận hành) xem đơn lỗi của chính mình; admin xem tất cả.
- doGet: GET mang _diagRequestId không có action (Google chuyển hướng nhầm) trả JSON RESPONSE_LOST thay cho trang HTML.
- Mail lỗi cache: cùng user + máy + loại dữ liệu + nội dung lỗi chỉ gửi 1 lần/30 phút; bỏ qua số đếm thay đổi (pending=, syncing=, 100/103).

Triển khai: paste Code_chinh_VTD.gs vào project VTD, Deploy → Manage deployments → Edit → New version (giữ URL). PN/TH API không đổi. Cài đè VTD-ChungTu-v5.9.14.apk, không gỡ app. Web tải lại trang.

Kiểm thử: test-release 30/30, test-network-queue 5/5 (giới hạn song song, ưu tiên lưu đơn, dùng chung request, thử lại chỉ lệnh đọc, gộp báo cáo cache), test-request-diagnostics, test-json-diagnostics, test-evidence, test-preauth, test-cache-fix. Chữ ký APK trùng chứng chỉ cũ (SHA-256 77b8abed…0ae4). Chưa kiểm thử trên máy thật; cần xem log sau khi triển khai: concurrentRequests ≤ 4, không còn GET trùng mã với POST.
