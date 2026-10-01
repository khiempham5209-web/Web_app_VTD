# Sửa migration: 4 dòng Booking chưa khớp

## Bổ sung: dòng 2, cột P chính trống

Đã đọc trực tiếp Booking!A2:AE2 và Booking_full!A2:AE2: P chính trống, P full = 2. Repair hiện cho phép thay đổi chỉ bên chính khi fingerprint của full (sau phục hồi kiểu Date/number tương đương) vẫn bằng checkpoint chung. Không ghi đè giá trị chính; đưa vào pendingMigration với reason primaryChangedOnly. Migration tiếp theo tính lại các cột công thức và đối soát như bình thường. Nếu full cũng thay đổi, vẫn dừng. Thay duy nhất FullStore.gs, chạy repair → migration → audit, không xóa lô nhập bù ngày 30/9. Tổng 28 kiểm thử mô phỏng.

## Bổ sung: repair báo thiếu ID dòng 2242

Đọc live cho thấy Booking dòng 2242 chưa có AD:AE và khác đơn với Booking_full dòng 2242. Không ghép theo số dòng. Bản mới để các dòng chưa có cả ID/checkpoint trong `pendingMigration`; chỉ phục hồi định dạng dòng đã xác định bằng ID. ID có sẵn mà mất ở full, hoặc mất ID nhưng còn checkpoint, vẫn dừng.

Thay lại duy nhất FullStore.gs bằng apps-script/00_FullStore.gs mới. Chạy pnRepairMigrationFormats(). Kết quả `ok:true, scope:formatRepair` nghĩa là bước sửa định dạng đạt; `migrationReady:false` và audit chưa khớp các dòng pending là dự kiến. Sau đó chạy pnMigrateFull() để đối soát các dòng đó theo khóa, rồi pnAuditFull(). Chỉ migration/audit cuối cùng `ok:true` mới cho phép tiếp tục triển khai. Không xóa full/AD:AE hoặc tự đánh dấu READY. 27 kiểm thử mô phỏng đạt sau bổ sung.

Log thực tế: Booking 2273/2273, lỗi dòng 742, 747, 1053, 1104; File đơn 2543/2543 và Chứng từ_FF 2329/2329 khớp.

Đã đọc đúng 4 dòng ở Booking và Booking_full: giá trị và ID giống nhau; cột K bên chính có định dạng DATE `d.m`, bên full có NUMBER `#,##0`. Code cũ lấy định dạng dòng 2 áp toàn bảng full, làm getValues trả Date ở chính nhưng number ở full. Không phải thiếu 4 hồ sơ. Chưa xác định giá trị nghiệp vụ Số Thùng ban đầu trước khi Sheets chuyển thành ngày; bản sửa không đoán hay đổi số thùng.

## Cách tiếp tục

1. Trong Apps Script, thay toàn bộ nội dung **FullStore.gs** (file trước đây lấy từ **00_FullStore.gs**) bằng bản mới tại apps-script/00_FullStore.gs. Không tạo thêm file trùng hàm. Các file khác không cần thay cho lỗi này.
2. Lưu. Giữ nguyên full, AD:AE và dữ liệu chính. Trong lúc phục hồi không nhập/sửa dữ liệu hoặc để trigger cũ chạy.
3. Chạy **pnRepairMigrationFormats()**. Hàm kiểm tra toàn bộ trước khi sửa, chỉ phục hồi định dạng/validation từ đúng dòng chính sang đúng dòng full. Không ghi lại dữ liệu nghiệp vụ hoặc ID, không clear. Nếu dữ liệu thực sự đã đổi so với checkpoint sẽ dừng.
4. Khi trả ok=true, chạy lại **pnMigrateFull()**, rồi **pnAuditFull()**. Chỉ tiếp tục cài đặt khi cả hai đạt.
5. Nếu còn lỗi, giữ nguyên và cung cấp log mới; không xóa full, không bỏ điều kiện audit.

Bản sửa không hạ độ chặt của fingerprint. Các lượt ghi mới copy định dạng theo từng nhóm dòng nguồn/đích tương ứng, không áp dòng 2 lên toàn bộ lịch sử. Hàm phục hồi chỉ cho migration chưa hoàn tất, không dùng sau khi đã bật nguồn full.

Kiểm thử bổ sung mô phỏng khác biệt Date/number do định dạng, phục hồi lượt migration dở và từ chối sửa khi giá trị thực khác. Chưa chạy hàm phục hồi trực tiếp trên Apps Script của người dùng.
