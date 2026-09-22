# Bộ đối chiếu lỗi API — 5.9.11

Mục đích: thu bằng chứng cho lỗi HTML/JSON, lỗi mạng, phản hồi sai cấu trúc và lỗi hàm xử lý. Đây là bản bổ sung chẩn đoán, chưa phải kết luận hoặc sửa nguyên nhân HTML/JSON.

## Cài đặt

Trong project Apps Script VTD đang dùng:

1. Thay nội dung code chính bằng `Code_chinh_VTD.gs` của bản này.
2. Thay nội dung file nhận log cũ bằng `VTD_Web_Diagnostics.gs`. Không giữ hai hàm `vtdDiagnosticsReceive` cùng tên. Bản cũ người dùng gửi có whitelist thiếu các trường JSON nên làm mất chi tiết chẩn đoán.
3. Thêm file mới `API_Request_Diagnostics.gs`.
4. Cập nhật deployment đang dùng sang phiên bản mới, giữ URL `/exec` hiện tại.

Trong project Apps Script PN:

1. Thay nội dung API bằng `ChungTuFF_WebApi.gs` của bản này.
2. Thêm file mới `API_Request_Diagnostics.gs` giống VTD.
3. Cập nhật deployment đang dùng sang phiên bản mới, giữ URL hiện tại.

Thay đổi trong hai code chính chỉ là điểm gọi bộ ghi log quanh doGet/doPost đang hoạt động. File riêng không thể tự quan sát những hàm này nếu không có điểm gọi. Không thêm một doGet/doPost khác vào project. Không sửa logic lưu đơn.

Web cần nạp bản 5.9.11; Android cài đè `VTD-ChungTu-v5.9.11.apk`, không gỡ app. File `request-diagnostics.js` được đóng gói trong cả web và APK. Chưa build điểm gọi phía server TH vì không có code server TH trong bản sửa này; các request TH qua fetch được ghi phía máy như các API khác.

## Xem log

- Sheet VTD, tab `_VTD_WEB_DIAGNOSTICS`: log trên máy, gồm `request_diagnostics_loaded` phiên bản 2, `request_started`, `request_headers`, `response_inspected`, `json_diagnostic`, `request_network_error`; tiếp tục giữ log đăng nhập, lỗi JavaScript và lỗi ghi bộ nhớ cũ.
- Sheet VTD và Sheet PN, tab `_API_REQUEST_DIAGNOSTICS`: request thực sự vào từng API, method GET/POST, action, thời gian xử lý, định dạng đầu ra, trạng thái nghiệp vụ, nhóm lỗi và một số stack frame nếu API trả về stack.
- Apps Script Executions: log `received` và kết thúc với cùng `requestId`.

Lấy `requestId` từ Detail của `json_diagnostic`, tìm cùng mã trong tab server tương ứng:

| Bằng chứng | Kết luận được phép |
|---|---|
| Server `response_ready`, `responseFormat=json`; máy nhận HTML cùng ID | Hàm xử lý đã tạo JSON; cần điều tra đường trả kết quả/redirect Google và phía máy. Chưa đủ để khẳng định quota. |
| Server ghi GET cho request máy gửi POST | Request quan sát ở entry point là GET. Kiểm tra redirect/định tuyến; không suy đoán rằng doGet và doPost tự xung đột. |
| Server `handler_threw` hoặc `serverOk=false` | Có lỗi trong quá trình xử lý/kiểm tra nghiệp vụ; xem nhóm lỗi và Executions. |
| Máy `request_network_error` | Fetch không nhận được phản hồi đọc được. Không có HTTP status để kết luận 404/quota. |
| `pageCategory=google_login`, `access_denied`, `quota_or_rate_limit` | Nội dung phản hồi khớp dấu hiệu tương ứng; là phân loại chẩn đoán, cần đối chiếu server. |
| Không có dòng server | Chưa đủ kết luận request không đến API: có thể deployment chưa cập nhật, ghi log lỗi/khóa bận hoặc execution bị ngắt. Xem Executions và dấu `API_DIAG_WRITE_*`. |

## Giới hạn và tải phát sinh

- Không đọc/lưu nguyên văn body phản hồi, đơn hàng, ảnh, mật khẩu, token hay query URL vào log. ID endpoint là mã băm dùng để so sánh cấu hình giữa máy.
- Không gửi lại request lưu đơn để thử lỗi. Log phía máy được gom theo lô qua bộ nhận hiện có; log server phát sinh một lần ghi Sheet cho request có ID. Bộ này có chi phí xử lý, dùng để điều tra rồi tháo khi đủ bằng chứng.
- Server giữ tối đa 5.000 dòng log luân phiên mỗi project; dòng cũ sẽ bị ghi đè. Không xóa dữ liệu nghiệp vụ.
- Log máy vẫn phụ thuộc mạng và phiên đăng nhập của bộ nhận hiện có. Khi API log không truy cập được, chưa thể hứa có log ngay trên Sheet. Hàng đợi hiện giữ tối đa 80 sự kiện; lỗi kéo dài có thể mất sự kiện cũ. Apps Script Executions cung cấp bằng chứng bổ sung.
- Native Android có ID trong body; nếu một redirect làm mất body trước entry point thì có thể không ghép được log server. Không suy diễn từ việc thiếu dòng log.

## Kiểm tra tự động

`node test-request-diagnostics.cjs .`

Kiểm tra ghép ID web/native, HTML 404, lỗi mạng, không gọi lặp request, không ghi bí mật, lỗi Sheet không thay đổi kết quả/exception nghiệp vụ, body null và GET trả HTML.

`node test-json-diagnostics.cjs json-diagnostics.js`

`node test-release.cjs`
