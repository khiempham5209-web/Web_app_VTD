# Thay đổi chờ build APK

Danh sách thay đổi **đã có trên web (Vercel) nhưng chưa có trong APK**. Bản APK mới nhất: **5.9.29** (build từ commit `beb2e64` + nâng APP_VERSION).

Mỗi khi sửa `index.html` hoặc `dashboard.html` mà chưa build APK, ghi thêm 1 dòng vào bảng dưới.
Khi build APK mới: đưa các dòng này vào `RELEASE-<phiên bản>.md`, rồi xóa khỏi bảng và cập nhật dòng "Bản APK mới nhất" ở trên.

| Commit | File | Thay đổi | Cần làm gì khác |
|---|---|---|---|

## Cách build (cho Claude trên laptop)

```
Pull main mới nhất. Build APK phiên bản kế tiếp như bản trước:
- nâng APP_VERSION trong index.html;
- đóng gói cả index.html và dashboard.html vào assets;
- ký đúng keystore VTD (cài đè được lên bản cũ), không commit keystore/mật khẩu;
- tạo RELEASE-<phiên bản>.md từ bảng "Thay đổi chờ build APK" trong PENDING-APK.md, rồi làm trống bảng đó;
- kiểm tra: unzip -l APK phải có assets/index.html và assets/dashboard.html.
```
