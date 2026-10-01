# Bàn giao để tiếp tục nâng cấp

## Trạng thái ngày 01/10/2026

Bộ code ứng viên đã được tạo local. Chưa triển khai API/app, chưa chuyển dữ liệu hoặc clear sheet thật. Việc lưu lên GitHub không triển khai hệ thống. Không coi trạng thái này là production.

Repo đích: khiempham5209-web/Web_app_VTD. Khi đưa bộ này vào repo, HEAD gốc là 4f96fed12b822bf94f4489e8039b6fdabc3ef2ba. index.html ở gốc repo khác bản frontend nguồn của bộ ứng viên (SHA256 lần lượt 707895c02a4b1b7922f309ea576434e513da7ab11d898a215c9c26a47ee351ec và f17523d5bd4d50873c9847fe344b2e57d623ac50f1df2a4bb5b9553fdd02df26). Vì vậy trước khi phát hành app phải tích hợp các thay đổi đồng bộ từ patch-frontend.cjs vào phiên bản đang phát hành và kiểm thử lại; không ghi đè mù index.html gốc bằng ứng viên. Lần upload này chỉ thêm thư mục bàn giao update-code-1-10, không thay app ở gốc repo.

## Đọc theo thứ tự

1. INSTALL.md: bảng tên 9 code cũ, 3 file mới và các bước cài.
2. README.md: nghiệp vụ đã chốt, giới hạn, nghiệm thu và rollback.
3. originals/ và SOURCE_MANIFEST.json: bản gốc và checksum.
4. apps-script/, frontend/: mã ứng viên cần triển khai.
5. CHANGES.patch, FORMULA_AUDIT.json, TEST_RESULTS.txt: thay đổi và bằng chứng kiểm tra.

## Các quyết định cần giữ

- API đọc/ghi Chứng từ_full; tìm đơn bằng khóa ổn định, không tin rowNumber cũ.
- Chính giữ tháng hiện tại; Chứng từ_FF giữ thêm tháng cũ chưa nhận. Shop hủy OD tháng cũ chỉ ở full sau dọn an toàn.
- Booking/File đơn giữ mọi dòng ở full; File đơn có mã trùng thực tế, không dedup tùy ý.
- Chỉ dọn khi full khớp, không còn yêu cầu pending và sync VHFF đạt. Không deleteRows làm dịch pivot.
- Giữ clientId và hàng đợi offline; không xóa cache để chữa lỗi đồng bộ.
- 12 file chung project để dùng chung ScriptLock. Lock không chặn người sửa tay sheet.
- Full cùng workbook làm gọn tab chính, không giảm tổng dung lượng workbook; chưa đo hiệu năng/quota thực tế.

## Phát triển trên máy khác

Cài Node.js và Git, mở thư mục này rồi chạy:

```sh
node build.cjs
node patch-frontend.cjs
node tests/run.cjs
node review-diff.cjs
```

Không cần thư viện npm. Build dùng originals/ đã đóng gói. build.cjs tái tạo 01–09, patch-frontend.cjs tái tạo frontend/index.html: nếu sửa các file sinh ra, phải cập nhật logic patch tương ứng để lần build sau không mất sửa đổi. Các helper 00, 10, 11 sửa trực tiếp. Không thay originals/ bằng mã mới.

Các test hiện dùng mô phỏng Apps Script/Sheets. Chưa kiểm thử Google/điện thoại thật, quyền OAuth, thời gian thực thi hay quota. Nguồn frontend gốc là bản github-publish/index.html trên máy cũ; cần xác nhận quy trình build/phát hành thực tế trước khi triển khai.

Không đưa token đăng nhập, .env hoặc dữ liệu xuất từ sheet vào repo. Bộ này lưu mã và tài liệu, không phải bản sao dữ liệu nghiệp vụ đầy đủ.
