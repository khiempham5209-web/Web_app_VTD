# Cài bộ đồng bộ V2 (theo tên cột, chỉ xử lý dòng thay đổi)

Áp dụng cho project Apps Script PN đã chạy bộ full (đã có `PN_FULL_READY`). Không cần chạy lại migration.

## File thay đổi

| File | Việc cần làm |
|---|---|
| `00_FullStore.gs` | Thay toàn bộ nội dung |
| `10_Formulas_Triggers.gs` | Thay toàn bộ nội dung |
| `11_API_FullSave.gs` | Thay toàn bộ nội dung |
| `12_SyncEngine.gs` | **Thêm file mới** |
| `01`–`09` | Giữ nguyên |

## Các bước

1. Dán 4 file trên vào project. Lưu.
2. Chạy hàm **`pnInstallTriggers`** một lần: xóa mọi trigger cũ (kể cả trigger 5 phút, 10 phút, 1 giờ sáng), chỉ cài 2 trigger theo sự kiện: **khi sửa Sheet** và **khi mở Sheet**. Không còn trigger chạy theo giờ.
3. Chạy hàm **`pnMirrorAll`** một lần: lượt đồng bộ đầu của bản mới. Kết quả phải có `"ok": true`.
   - Lượt này tự xác định đúng bên vừa sửa từ dấu vân tay kiểu cũ, rồi chuyển sang kiểu mới.
   - Nếu có tab `_PN_CONFLICTS` (ẩn) thì xem các dòng trong đó: đó là bản bị thay khi hai bên cùng sửa.
4. Chạy **`pnCleanupPreview`** để xem các dòng tháng cũ sẽ rời tab chính trong lần dọn tới.
   Muốn dọn ngay thì chạy **`pnCleanupDaily`**.
5. Deploy → Manage deployments → Edit → **New version** (giữ URL) để API app dùng code mới.

## Không có trigger theo giờ

- Sửa tay: đồng bộ đúng dòng vừa sửa.
- App lưu đơn: API ghi Chứng từ_full rồi đẩy đúng dòng đó sang Chứng từ_FF.
- Dọn tháng cũ: tự chạy 1 lần/ngày ở lần sửa hoặc mở Sheet đầu tiên trong ngày.
- VHFF: khi có thay đổi, mỗi lần sửa Sheet làm 1 phần. Cần cập nhật ngay thì chạy tay `pnRunVhffNow`.
- Dữ liệu vào bằng script/nguồn khác (không qua sửa tay, không qua app): chạy tay `pnScheduledReconcile`.

## Quy tắc vận hành

- Cột tìm theo **tên tiêu đề**. Đổi chỗ cột thoải mái. Tab chính: cột nghiệp vụ là dãy tiêu đề **liền nhau từ cột A**
  (bảng pivot/ghi chú đặt sau một cột trống không bị đụng tới). Thêm cột mới ở tab chính: tab full tự có cột cùng tên.
- **Đổi tên cột** = cột mới (cột tên cũ ở full giữ nguyên lịch sử). Hai cột trùng tên trong một tab: đồng bộ dừng và báo lỗi.
- Không xóa/sửa các cột ẩn `__PN_ID`, `__PN_BASE`, cột `Tháng` của full.
- Một dòng là hồ sơ khi có mã: Booking = Row Labels; File đơn = Mã đơn; Chứng từ = Số đơn hàng hoặc Mã đơn GHTK.
  Dòng trống/chỉ có giá trị mặc định không được đồng bộ.
- Xóa một dòng tháng hiện tại ở tab chính: dòng đó được chép lại từ full (full là lịch sử). Muốn bỏ hẳn thì xóa ở cả hai tab.
- Hai bên cùng sửa một dòng: bên vừa sửa tay thắng; không biết thì tab chính thắng. Bản bị thay lưu ở `_PN_CONFLICTS`.
- Full bị mất dòng: không chặn đồng bộ, báo ở Script Property `PN_FULL_LOSS`, tự chép lại các dòng còn ở tab chính.

## Hàm kiểm tra

- `pnAuditFull`: đối chiếu từng dòng tab chính với full (phải `ok: true`).
- `pnListNonRecordRows`: liệt kê dòng "rác" bản cũ đã tạo (có ID nhưng không có mã). Chỉ liệt kê, không xóa.
- `pnCleanupPreview`: các dòng sẽ được dọn.

## Quay lại bản trước

Dán lại 3 file `00`, `10`, `11` của commit trước, xóa `12_SyncEngine.gs`, chạy `pnInstallTriggers`.
Dữ liệu không bị đổi cấu trúc. Lưu ý: bản cũ không hiểu `__PN_BASE` kiểu mới; dòng nào đang khớp hai bên thì
bản cũ chạy bình thường, dòng đang lệch sẽ bị bản cũ báo "Xung đột" và dừng như trước.
