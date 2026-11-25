# Log Thao tác Phạm Nguyên

## Giai đoạn 1 — Apps Script cho APK đang dùng

1. Trong project **PJ_Phạm Nguyên**, thay toàn bộ nội dung file `01_API.gs` và `11_API_FullSave.gs` bằng bản trong `update-code-1-10/apps-script/`. Thêm file `16_PN_AppLog.gs` cùng thư mục.
2. Tạo version Apps Script mới và cập nhật deployment web app đang dùng để giữ nguyên URL `/exec`.
3. Test một đơn chứng từ, một đơn sản phẩm nhiều SKU, một đơn kết hợp và một lượt gửi lại trùng `clientId`. Kiểm tra cả tab đích lẫn `PN_LOG_APP`. Tab log được tạo khi có request chứa hoặc đối chiếu được mã GHTK.
4. Test cùng một mã GHTK ở hai lượt độc lập. Bảng có một dòng, các cột tóm tắt theo lượt mới nhất, JSON giữ cả hai lượt. Có thể rollback bằng cách trỏ deployment về version cũ.

`Update 4–8` dành cho mốc phía API, `Update 1–3` và `Update 9` dành cho dữ liệu của bản app sau. `Vị trí ghi` là một chuỗi trên một dòng. Số dòng là vị trí tại thời điểm quan sát, mã bản ghi nằm trong JSON. Khi đồng bộ `Chứng từ_FF` được xếp hàng, log ghi `MIRROR_PENDING`. VHFF ghi `VHFF_SCHEDULED` khi luồng lưu gọi trình lập lịch. Apps Script có thể bị dừng giữa chừng hoặc lỗi ghi log. Khi đó dùng mốc cuối cùng đã lưu, đối chiếu `_PN_REQUESTS` và Apps Script Executions. Log lỗi không làm API báo lưu đơn thất bại.

## Giai đoạn 2 — Chờ duyệt

`phase-2.patch` chứa thay đổi frontend, điểm nhận sự kiện bên PN và endpoint kiểm tra phiên bên VTD. Bộ API này kiểm tra phiên đăng nhập VTD có hiệu lực. App gắn mã riêng cho sự kiện và gửi theo lô. App dùng chung `clientId` với request save hiện tại. Các sự kiện cần có mã GHTK đã xác định, vẫn lưu trên thiết bị khi mất mạng và được gửi sau khi đăng nhập lại bằng đúng tài khoản. Nếu máy vẫn offline, Sheet chưa thấy log phía app.

Khi được duyệt, từ thư mục gốc của repo chạy:

```sh
git apply --check pending-review/pn-app-log/phase-2.patch
git apply pending-review/pn-app-log/phase-2.patch
```

1. Deploy `Code_chinh_VTD.gs` đã áp bản vá lên deployment VTD hiện tại. Trong Script Properties của project PN, đặt `PN_LOG_VTD_API_URL` thành URL `/exec` của deployment VTD đó.
2. Deploy lại Apps Script PN với `01_API.gs` và `16_PN_AppLog.gs` đã áp bản vá. Kiểm tra API từ chối phiên thiếu/hết hạn. Thử log từ tài khoản nhân viên thật trên bản web kiểm thử trước khi build.
3. Kiểm tra ca thiếu ảnh, hàng chờ offline, lỗi đọc ảnh, mất phản hồi sau khi Sheet đã ghi, retry nhiều lần và log đến trễ. Xác nhận mốc app xuất hiện ở Update 1–3 và 9, JSON có `at` và `receivedAt`, mã GHTK chỉ có một dòng. Kiểm tra thử đăng xuất rồi đăng nhập tài khoản khác trên cùng máy.
4. Nâng `APP_VERSION`, `versionName` và `versionCode` lên số tiếp theo so với APK mới nhất tại ngày build. Xem `PENDING-APK.md` và release mới nhất. Đóng gói lại assets, ký cùng keystore hiện tại và test cài đè khi còn đơn trong hàng chờ.
