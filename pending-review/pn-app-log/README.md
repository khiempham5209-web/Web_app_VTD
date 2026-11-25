# Log Thao tác Phạm Nguyên

## Giai đoạn 1 — Apps Script cho APK đang dùng

1. Trong project **PJ_Phạm Nguyên**, nếu đã cài bản log đầu tiên, thay `02_Reject.gs`, `11_API_FullSave.gs`, `12_SyncEngine.gs` và `16_PN_AppLog.gs` bằng các tệp tương ứng. Nếu cài lần đầu, cập nhật thêm `01_API.gs`.
2. Tạo version Apps Script mới và cập nhật deployment web app đang dùng để giữ nguyên URL `/exec`.
3. Chạy tay `pnRebuildAppLogHistory` trong Apps Script để hiển thị lại lịch sử sẵn có. Chạy thêm nếu kết quả trả `complete: false`. Cột R là nhật ký đọc được, cột S (ẩn) giữ JSON gốc. Lịch sử vượt dung lượng ô được chuyển sang `PN_LOG_HISTORY`.
4. Kiểm tra một đơn chứng từ thông thường, một đơn sản phẩm, một ca thu hồi khớp Mã ecom CT và một đơn tích Đã tạo sv từ Sheet. Các bước ghi chung GHTK nằm trên một hàng, mỗi mốc xuống một dòng trong ô Update. Kiểm tra app trả thành công nếu chỉ luồng đồng bộ báo lỗi và phần ghi nguồn đã thành công.

`Update 4–8` dành cho mốc phía API. `Update 1–3` và `Update 9` dành cho bản app sau. `Vị trí ghi` là một dòng. `APP_SAVE`, `SYNC_FF`, `SYNC_PRODUCT_VHFF` và `SYNC_SV_VHFF` có kết quả riêng trong Log. `PRODUCT_SYNC_PENDING` xác nhận đã xếp lịch, kết quả worker xem trong Apps Script Executions. `SV_CHECKED` cho biết luồng Sheet đọc được checkbox TRUE ở thời điểm chạy. Bản hiện tại không dựng lại các mốc SV trước thời điểm cập nhật. Nếu cần hoàn tất lượt lưu cũ còn treo, gửi lại đúng mục hàng chờ đang có để dùng lại `clientId`.

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
