# VTD Chứng từ FF 5.9.27

Gồm bản sửa theme ảnh nền (link Drive cho nền trong app, theme màn đăng nhập tải lại mỗi lần mở app) và các bổ sung màn Nhập TH dưới đây. Các màn VTD, PN giữ nguyên.

## Nhập TH
- Thêm ô **Loại hoàn** ngay sau Mã đơn, trước Tên sản phẩm. Mặc định Hoàn trả, chọn Ecom khi cần. Cả đơn dùng chung 1 Loại hoàn.
- Quét QR ra link `https://i.ghtk.vn/1598188396` thì app chỉ lấy 10 số cuối làm mã đơn. API cũng tự tách khi ghi sheet. Chỉ áp dụng ở màn TH.
- DS SKU TH tải giống SKU PN:
  - Lưu IndexedDB, mỗi ngày đối soát 1 lần, chỉ tải SKU mới hoặc đã đổi (khóa theo SUB SKU).
  - Có dòng trạng thái trong "Dữ liệu trên máy này".
  - Lỗi (vd thiếu SUB SKU) được báo về admin qua mail báo cáo cache.
- **Sửa / xóa đơn trong ngày**:
  - Thao tác: Trang chủ → bấm vào đơn TH → Sửa → sửa thông tin đơn và SKU, thêm SKU, Xóa SKU, Xóa đơn.
  - Không có nút xóa ngoài màn chi tiết.
  - Mất mạng vẫn lưu được, có mạng app tự gửi.
  - Đơn còn chờ gửi vẫn sửa được như đơn thường.
  - Mỗi lần sửa/xóa được ghi vào tab "Log sửa TH" và gửi mail báo admin.
- Quyền mới trong Cài đặt → user: **Sửa đơn TH**, **Xóa đơn TH**.
- Màn **Tìm kiếm** thêm lựa chọn **TH Truemilk**:
  - Tìm trên dữ liệu đã tải về máy, lọc theo Mã đơn, Loại hoàn, Tình trạng, Tên sản phẩm, User, Ngày.
  - Dữ liệu được tải khi mở app/đăng nhập và khi bấm tải lại trang chủ.
- Bảng TH trong ngày hiện "Chờ gửi" cho đơn/lượt sửa chưa gửi xong. Đơn đã xóa hoặc đổi mã trên sheet không còn sót trong cache máy.

## API Nhập TH (Code_api_TH.gs)
- Bỏ hàm `returnTh_listSku_` thừa (bản cũ bị bản sau đè, không được dùng).
- Thêm action `skuSync`, `keySync`, `updateReturnTh`, `deleteReturnTh`. `listToday` trả thêm ID từng SKU.
- Các action ghi dữ liệu chạy trong LockService để 2 máy lưu cùng lúc không ghi chồng dòng.
- Cấu hình mới: `logSheet` (mặc định "Log sửa TH"), `alertEmail` (mail nhận báo sửa/xóa).
- **Cần Deploy phiên bản mới** cho Web App TH. Lần đầu sẽ phải cấp thêm quyền gửi mail (MailApp).
