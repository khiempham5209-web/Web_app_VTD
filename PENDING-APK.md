# Thay đổi chờ build APK

Danh sách thay đổi **đã có trên web (Vercel) nhưng chưa có trong APK**. Bản APK mới nhất: **5.9.28** (build từ commit `800b6bf`).

Mỗi khi sửa `index.html` hoặc `dashboard.html` mà chưa build APK, ghi thêm 1 dòng vào bảng dưới.
Khi build APK mới: đưa các dòng này vào `RELEASE-<phiên bản>.md`, rồi xóa khỏi bảng và cập nhật dòng "Bản APK mới nhất" ở trên.

| Commit | File | Thay đổi | Cần làm gì khác |
|---|---|---|---|
| `475e5d0` | dashboard.html | Dashboard: nhãn khoảng ngày đúng khi chọn Tùy chọn (bảng Hiệu suất nhân viên không còn ghi "hôm nay" sai); trên điện thoại thanh lọc không dính đầu màn hình, thêm nút "↑ Bộ lọc" | Không |
| `13db3d5`, `8963f4c` | index.html | Tìm kiếm Phạm Nguyên: ô tìm mã tìm được cả Mã ecom CT (nhập một phần cũng được); đơn khớp theo mã ecom hiện "ecom &lt;mã&gt;" ở dòng kết quả; nhãn ô đổi thành "Mã đơn/Số đơn hàng", placeholder "Nhập/quét mã đơn GHTK, số đh, ecom"; cập nhật hướng dẫn trong app | Không |

## Cách build (cho Claude trên laptop)

```
Pull main mới nhất. Build APK phiên bản kế tiếp như bản trước:
- nâng APP_VERSION trong index.html;
- đóng gói cả index.html và dashboard.html vào assets;
- ký đúng keystore VTD (cài đè được lên bản cũ), không commit keystore/mật khẩu;
- tạo RELEASE-<phiên bản>.md từ bảng "Thay đổi chờ build APK" trong PENDING-APK.md, rồi làm trống bảng đó;
- kiểm tra: unzip -l APK phải có assets/index.html và assets/dashboard.html.
```
