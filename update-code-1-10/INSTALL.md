# Thay code nào và cài như thế nào?

Đây là bộ ứng viên chưa triển khai lên Google. Không chỉ dán code rồi bật clear.

## Thay toàn bộ nội dung 9 file cũ

Giữ tên file đang có trong Apps Script. Lấy nội dung mới từ apps-script/, không lấy originals/. Không giữ hai bản định nghĩa trùng hàm.

| Tên code người dùng gửi | File mới thay nội dung |
|---|---|
| Code API | apps-script/01_API.gs |
| Sync_chứng từ_VHFF | apps-script/02_Reject.gs |
| Sync_booking_to_chungtuff | apps-script/03_Booking.gs |
| Bangiaochungtu | apps-script/04_Handover.gs |
| DS SKU_and_TT Nhap | apps-script/05_SKU.gs |
| Sync CT+SP_VHFF | apps-script/06_VHFF.gs |
| Sync hàng lỗi shop_GHTK_to_VHFF | apps-script/07_Defect.gs |
| SYNC sự vụ | apps-script/08_Incident.gs |
| Huyện | apps-script/09_Area.gs |

## Tạo thêm 3 file mới

- apps-script/00_FullStore.gs: quản lý full, đối soát, chuyển dữ liệu, dọn và phục hồi.
- apps-script/10_Formulas_Triggers.gs: tính thay công thức, lịch chạy và điều phối sync.
- apps-script/11_API_FullSave.gs: bổ sung luồng lưu API và chống trùng.

Tất cả 12 file phải ở cùng project API hiện tại. Nếu script cũ nằm ở nhiều project, cần tập hợp và tắt trigger tương ứng ở các project cũ; installer không làm được việc này xuyên project/tài khoản.

## Thứ tự triển khai

1. Sao lưu cả sheet nguồn và sheet VHFF, lưu phiên bản deployment. Tạm ngừng thao tác khi chuyển lần đầu.
2. Thay 9 file, thêm 3 file; tắt các trigger cũ liên quan.
3. Chạy pnMigrateFull(): chuyển dữ liệu và 4 công thức; chưa clear chính.
4. Chạy pnAuditFull(): chỉ tiếp tục khi ok=true và không có errors.
5. Chạy pnScheduledReconcile(), kiểm tra dữ liệu VHFF và Script Properties PN_VHFF_ERROR / PN_VHFF_DIRTY theo README.
6. Cập nhật New version cho deployment API hiện tại, giữ URL /exec, quyền và cách thực thi.
7. Tích hợp thay đổi frontend qua quy trình app hiện tại. Đây là mã app, không dán vào Apps Script. index.html gốc repo GitHub khác bản nguồn của ứng viên: cần áp dụng thay đổi đồng bộ vào bản đang phát hành và kiểm thử lại, không ghi đè nguyên file ứng viên. Xem HANDOFF.md. Giữ cache/hàng đợi chưa gửi.
8. Chạy pnInstallTriggers(). Dọn tự động vẫn tắt.
9. Kiểm thử thật các ca trong README, chạy pnCleanupPreview().
10. Khi kết quả đạt mới chạy pnEnableCleanup(). Không clear thủ công.

## Cần kiểm tra/điều chỉnh gì?

- Đúng spreadsheet nguồn/đích, tên tab và quyền Drive/Sheets của tài khoản chạy script.
- Đúng project API hiện tại; không còn trigger cũ chạy song song ở project khác.
- Đúng bản frontend đang được phát hành. Giao diện/nghiệp vụ giữ nguyên nhưng mã đồng bộ app có thay đổi.
- Pivot/báo cáo ngoài 9 code cần lịch sử phải đổi nguồn sang full. Pivot File đơn!T:W hiện vẫn theo chính.
- Không tự xóa ARRAYFORMULA trước migration; giữ công thức TT Nhập!L3 và DS BC.
- Không chỉnh AD:AE, nhật ký _PN_REQUESTS hoặc các ID phụ trợ.

Không cần tự sửa mapping nghiệp vụ chỉ vì thêm 3 tab nếu môi trường đúng như đã kiểm tra. Tuy nhiên, kiểm thử mô phỏng không thay thế nghiệm thu thật.
