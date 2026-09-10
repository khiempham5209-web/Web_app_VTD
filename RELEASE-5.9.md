# Bản 5.9 — PN và MASTER SKU theo mã

## Thứ tự cập nhật

1. Trong Apps Script PN đang dùng, thay nội dung bằng **ChungTuFF_WebApi.gs** trong bản này. File đã đầy đủ các hàm, không cần ghép file JS khác. Tab **Hoàn sản phẩm** cần có header **Phân loại** và **Hạn sử dụng** ở bất kỳ vị trí nào.
2. Chọn **Deploy → Manage deployments → Edit → New version → Deploy** trên deployment hiện có để giữ URL `/exec` đang cấu hình trong app.
3. Làm tương tự với Apps Script VTD, sử dụng **Code_chinh_VTD.gs**. Hai API vẫn giữ action cũ cho app 5.8 và hàng đợi thao tác cũ; các trường báo cáo mới được bổ sung vào tab báo cáo hiện tại.
4. Cài **VTD-ChungTu-v5.9.apk** đè lên 5.8 trên một máy kiểm thử. Không gỡ app hoặc xóa dữ liệu. APK dùng cùng chứng chỉ ký với 5.8, versionCode 5900.
5. Mở app, chờ tải ban đầu xong. Đối soát chạy nền sau tối thiểu khoảng 30 giây, đợi các API đang chạy, hoạt động upload và nhập liệu tạm lắng. Kiểm tra báo cáo theo đúng user và device: **Đã đủ**, thiếu 0, cần cập nhật 0, không có dòng thiếu/trùng mã, có thời điểm đối soát thành công mới.
6. Thử quét số đơn hàng và mã GHTK, đơn thu hồi chỉ có mã GHTK; thử nhập một SKU với Phân loại và ghi chú nhiều lựa chọn. Kiểm tra hai cột mới trong Sheet. Sau kiểm thử máy thật mới phát hành rộng.

Đưa file lên GitHub không tự deploy Apps Script. Bản này chưa được xác minh với deployment thật hoặc thiết bị letuyet; kiểm thử hiện tại dùng dữ liệu giả lập và IndexedDB thật trong trình duyệt.

## Cơ chế dữ liệu

- PN: `ORDER:<số đơn hàng>`; thiếu số đơn hàng thì `GHTK:<mã GHTK/thu hồi>`. Row chỉ định vị. Cache mới nằm trong IndexedDB riêng, không sửa/xóa hàng đợi ảnh và thao tác.
- Giữ tải nối tiếp PN lúc mở app. Đối soát nền đọc danh sách mã + row + fingerprint nội dung; tải tối đa 100 bản ghi thiếu/thay đổi mỗi yêu cầu. Danh sách nhẹ được phân trang 1.000 mục. Đổi row đơn thuần chỉ cập nhật map.
- MASTER SKU VTD: `MATERIAL:<mã vật tư>`, lưu IndexedDB, đối soát theo nội dung. Action mới không bị giới hạn tổng 5.000 dòng. Giữ action `listProducts` cũ để tương thích app cũ.
- Lần nâng cấp đầu, dữ liệu cache cũ chưa có fingerprint có thể cần tải lại để xác minh. Các lượt đã đối soát đầy đủ chỉ tải bản ghi mới/thiếu/thay đổi.
- Mỗi yêu cầu đối soát đọc nguồn thực tế nên phát hiện sửa tay, không chỉ thay đổi do API save. Nếu nguồn đổi giữa lượt tải, dừng xác nhận và thử lại; không đánh dấu đủ từ dữ liệu chắp ghép nhiều phiên bản.
- Khi đối soát hoàn chỉnh, loại cache không còn trên nguồn khỏi danh sách tra cứu/chọn mới. Chỉ thực hiện khi danh sách nguồn đầy đủ và không có mã trống/trùng; lịch sử và queue giữ nguyên.
- Lặp nền khoảng 10 phút; khi lỗi thử lại có giãn cách 1–10 phút. Đây là kiểm tra theo từng thời điểm, không phải đồng bộ tức thời khi Sheet vừa sửa.
- Dòng không có khóa hoặc trùng khóa được báo lỗi, không tự coi là đủ. Dòng hoàn toàn trống không tính là đơn/SKU.
- Báo cáo admin có matched/missing/changed, invalidCount/duplicateCount, sourceVersion, verifiedAt và protocol. Báo cáo từ app cũ có chú thích chưa xác minh theo mã.

## Nhập SKU PN

- Áp dụng cho **Sản phẩm** và **Sản phẩm + chứng từ**.
- **Phân loại** ở trước Tình trạng sản phẩm: Hoàn trả/Hàng thu hồi; mặc định trống, bắt buộc chọn.
- Ghi chú chọn nhiều: Bình thường, Hàng cận date, Xẹp hơi, Hết hạn sử dụng, Vỏ bị xé/hở, Khác. Bình thường không trộn mục khác. Khác cần nội dung bổ sung.
- Hạn sử dụng không bắt buộc. Chỉ hiện với Hàng cận date/Hết hạn sử dụng. Nhập DD/MM/YYYY (tự chèn dấu `/` khi gõ số); ngày không hợp lệ bị chặn. Bỏ các ghi chú liên quan sẽ xóa giá trị hạn sử dụng đang ẩn.
- Các trường mới được giữ khi thêm/sửa SKU, lưu queue offline và đọc lại lịch sử; ghi theo tên header vào Hoàn sản phẩm. Đọc PN/DS SKU PN không còn giới hạn A:W/A:J.

## Kiểm thử đã thực hiện

- `node test-release.cjs`: 16 nhóm kiểm thử gồm header đổi vị trí ra ngoài W, bổ sung GHTK thủ công, đơn thu hồi, mã trống/trùng, master 5.030 SKU, race tra cứu, ghi lỗi không được báo đủ, và đối soát 2.000 + 30 đơn (tải 30, lượt tiếp tải 0, sửa một đơn tải 1).
- `node qa-server.cjs`: fixture localhost, chặn mọi kết nối ngoài. IndexedDB thực: ghi/đọc, rollback khi put lỗi, đổi row không nhân đôi, lưu trường SKU, chuyển legacy cache 5.8. Form kiểm tra tại 390×844, chọn nhiều ghi chú và hai loại hoàn.
- Kiểm tra cú pháp hai API và script app, `git diff --check`.
- Android SDK build và apksigner verify đạt v1/v2/v3. Chứng chỉ SHA-256 trùng APK 5.8: `77b8abedfd8035841637d57050dce9add7c3a33c3735468fe3d2bafc129c0ae4`.
- HTML trong APK đã đối chiếu trùng source, chỉ đóng gói một asset HTML hiện tại.

APK SHA-256: `063172A7CAB6FE606AA02B4A00B3FCA92400581541463A0B15D515A2997C54EE`.
