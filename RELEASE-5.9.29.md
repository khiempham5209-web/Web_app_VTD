# VTD Chứng từ FF 5.9.29

## Nhập VNPT (mới)
- Màn **Nhập VNPT** giống Nhập TH: một đơn nhiều SKU, Loại hoàn, tách mã từ link GHTK, chọn hoặc tạo SKU mới, chụp ảnh, mất mạng vẫn lưu.
- **Quét ICCID** từng SIM: mỗi mã 1 ô, bấm × để bỏ mã quét nhầm, báo trùng. Có quét thì số ICCID phải bằng Số lượng; hàng không phải SIM bỏ trống.
- Sửa / xóa đơn VNPT trong ngày (sửa được cả ICCID).
- Tìm kiếm: thêm shop VNPT, tìm theo mã đơn hoặc ICCID.
- Trang chủ: bảng Chi tiết VNPT trong ngày (dưới TH) và dòng VNPT trong tổng nhập.
- Quyền riêng: màn Nhập VNPT, Lưu / Sửa / Xóa đơn VNPT (admin cấp ở Config user).
- Hướng dẫn trong app: thêm mục Nhập VNPT, Sửa / xóa đơn VNPT, Tìm kiếm VNPT.

## Cài đặt
- Cài đặt cá nhân → Vị trí các màn: chọn màn nằm ở thanh dưới hay trong Cài đặt, nút ↑↓ đổi thứ tự. Màn đặt trong Cài đặt hiện ở mục "Màn khác".

## Dashboard
- Thêm nguồn VNPT (tab Hoàn trả). Nhãn khoảng ngày đúng khi chọn Tùy chọn; trên điện thoại thanh lọc không dính đầu màn hình, có nút "↑ Bộ lọc" (`475e5d0`).

## Tìm kiếm Phạm Nguyên
- Ô tìm mã tìm được cả Mã ecom CT (nhập một phần cũng được), nhãn "Mã đơn/Số đơn hàng" (`13db3d5`, `8963f4c`).

## Cách build
- Giữ nguyên classes.dex / manifest / resources của 5.9.28, chỉ thay assets (index.html, dashboard.html, các file diagnostics) và nâng versionCode 5929 / versionName 5.9.29. Ký cùng keystore VTD, cài đè được lên bản cũ.

Cần deploy trước khi phát hành: **API VNPT** (API_VNPT.gs có ICCID), **API VTD** (Code_chinh_VTD.gs), **Dashboard_VTD.gs**.
