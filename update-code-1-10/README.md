# update code 1/10

Bộ cập nhật 3 tab full — Phạm Nguyên, ngày 01/10/2026. Thư mục GitHub: `update-code-1-10` (dấu `/` trong tên yêu cầu được thay bằng `-` để không tạo thư mục lồng nhau).

## Trạng thái bàn giao

Đã tạo bộ mã ứng viên riêng, dựa trên đúng 9 đoạn code người dùng gửi. Chưa sửa mã đang triển khai, chưa chạy migration, chưa xóa công thức hay clear dữ liệu trên Google Sheets thật. Không đổi URL API.

24 bài kiểm thử tự động đã chạy trên bộ mô phỏng Apps Script/Sheets; toàn bộ 12 file .gs và JavaScript inline trong bản app mới đều qua kiểm tra cú pháp. Đây chưa phải kiểm thử trên dịch vụ Google hoặc trên điện thoại thật. Cần chạy trình tự nghiệm thu dưới đây trước khi bật dọn tự động.

Đã đọc trực tiếp công thức và đối chiếu 2.240 dòng Booking, 2.426 dòng File đơn ngày 30/09/2026. Kết quả tương đương; File đơn dòng 1342 đang có lỗi tra kho “Lộc Nga” trong DS BC, hai ô P/Q là #N/A. Bộ mã không tự gán quản lý/khu vực thay thế. Xem FORMULA_AUDIT.json.

## Các thư mục

Đọc [INSTALL.md](INSTALL.md) để biết chính xác file nào thay code cũ và thứ tự cài. Đọc [HANDOFF.md](HANDOFF.md) khi tiếp tục nâng cấp ở máy khác.

- originals/: nguyên bản 9 file và frontend, giữ nguyên nội dung; SHA256 nằm ở SOURCE_MANIFEST.json.
- apps-script/: 9 file đã chỉnh các điểm tích hợp + 3 file bổ sung 00, 10, 11.
- frontend/index.html: bản app ứng viên, dựa trên D:/Projeck_FF/PJ_new_api/github-publish/index.html. Không ghi đè file nguồn.
- build.cjs: tái tạo 9 file nghiệp vụ từ nguyên bản đính kèm; không ghi đè 3 helper.
- patch-frontend.cjs: tái tạo bản app riêng bằng các thay thế có kiểm tra điểm neo.
- tests/run.cjs: mô phỏng, kiểm tra hồi quy và các ca lỗi.

## Hành vi được chốt

| Nguồn | Giữ ở chính | Giữ ở full |
|---|---|---|
| Booking | Tháng hiện tại; dữ liệu chưa xác định được tháng giữ lại để kiểm tra | Tất cả dòng đã lưu |
| File đơn | Tháng hiện tại; dữ liệu chưa xác định được tháng giữ lại để kiểm tra | Tất cả dòng đã lưu |
| Chứng từ_FF | Tháng hiện tại mọi trạng thái; tháng cũ chưa nhận | Mọi tháng, mọi trạng thái |

Chứng từ tháng cũ có “Đã nhận chứng từ” hoặc “Shop hủy OD” được dọn sau đối soát. Tháng cũ có trạng thái lạ/trống không tự dọn. Dữ liệu tương lai cũng không tự dọn.

API dùng Chứng từ_full cho init/page/keySync/lookup/today/save. App chỉ tải bản thiếu hoặc thay đổi theo khóa; không dựa vào việc số dòng tab chính tăng hay giảm. Khi đổi nguồn, có sourceEpoch=pn-full-v1 để đối soát cache. Hàng đợi thao tác/ảnh chưa gửi được giữ nguyên.

API tìm lại hồ sơ bằng syncKey hoặc mã đơn duy nhất trước khi lưu; bỏ qua rowNumber đã cũ. Cùng clientId không chạy lại thao tác đã hoàn tất. Thay sessionToken hoặc rowNumber không biến một yêu cầu gửi lại thành yêu cầu mới.

Lưu app cập nhật full rồi phản chiếu chính. Đơn cũ đã nhận/hủy chỉ có ở full không tự thêm về chính; đổi thành chưa nhận thì trở lại chính. Đơn cũ đang ở chính vừa được nhận vẫn còn đến lượt dọn hằng ngày.

Đơn mới/sửa tay chính được phản chiếu bằng trigger. Sửa full được kéo về chính nếu bản chính còn tồn tại. Hai bên cùng sửa kể từ lần đối soát trước thì báo xung đột, không chọn một bản để ghi đè tùy ý. Hoạt động nhập bằng script không kích hoạt onEdit sẽ được lượt đối soát 5 phút xử lý.

## Nghiệp vụ giữ lại và phần bổ sung

| File | Thay đổi |
|---|---|
| 01_API.gs | Đổi nguồn cố định sang full; thêm epoch; bao luồng save bằng nhật ký; ghi chứng từ theo giai đoạn; chống trùng dòng hoàn bằng clientItemId. Các trường, kiểm tra SKU, cách tải ảnh và tính CBM giữ từ bản gốc. |
| 02_Reject.gs | Đọc chứng từ/full và File đơn/full; giữ điều kiện Đã tạo sv, khóa GHTK và hai cột Ecom nhập tay. Chặn batch có GHTK trùng thay vì để boolean true bị dùng nhầm như số dòng đích. |
| 03_Booking.gs | Nguồn booking có thể chính/full; chống tạo trùng dựa vào chứng từ_full; tra mã GHTK từ File đơn_full. Giữ mapping, ngày và các giá trị mặc định gốc. |
| 04_Handover.gs | Báo cáo/dropdown theo Chứng từ_full, giữ mẫu và thống kê hiện tại. |
| 05_SKU.gs | Giữ nghiệp vụ DS SKU/TT Nhập; chỉ thống nhất trigger/lock. |
| 06_VHFF.gs | Đọc chứng từ_full, lấy ngày booking từ Booking_full; giữ điều kiện ngày bàn giao và các cột output. Ghi dữ liệu mới trước, sau đó mới xóa phần đuôi cũ. |
| 07_Defect.gs | Giữ mapping Hàng lỗi/DS SKU; thống nhất trigger/lock; xóa đuôi sau ghi thành công. |
| 08_Incident.gs | Tra File đơn_full; giữ nghiệp vụ Sự vụ; xóa đuôi sau ghi thành công. |
| 09_Area.gs | Tra File đơn_full; dispatcher áp dụng cho Booking chính và full. Giữ quy tắc phân loại nội thành/huyện. |
| 00_FullStore.gs | Di chuyển ban đầu, ID, đối soát hai bản, Tháng, filter, backup/dọn/phục hồi. |
| 10_Formulas_Triggers.gs | Bốn công thức thay bằng code, dispatcher, đối soát 5 phút, xác nhận sync VHFF. |
| 11_API_FullSave.gs | Nhật ký idempotency và các bước lưu có thể tiếp tục. |

Các file nghiệp vụ vẫn chứa phần bình luận hướng dẫn cài bản cũ để dễ đối chiếu. Khi dùng bộ full, làm theo tài liệu này.

## Công thức chuyển thành code

- Booking!P: giữ đúng ROUNDUP(Tổng bánh * TT Nhập!L3, 0), làm tròn ra xa 0, điều kiện Ngày trống thì để trống.
- File đơn!P: XLOOKUP Kho đích qua DS BC!G, trả DS BC!E.
- File đơn!Q: lookup tương tự, trả DS BC!O.
- File đơn!R: định dạng Thời gian tạo đơn thành dd/MM/yyyy.

Giữ nguyên TT Nhập!L3 và công thức trong DS BC. Khi TT Nhập/DS BC thay đổi, phép tính trên chính/full được cập nhật; có lịch 5 phút bù thay đổi do script. Giữ cách XLOOKUP chọn kết quả đầu tiên, so khớp không phân biệt hoa thường. Không bỏ dấu hay tự đổi tên kho.

pnMigrateFull kiểm tra đúng công thức đã quan sát và đối chiếu kết quả trước khi bỏ các ô công thức gốc. Nếu khác sẽ dừng. Các cột đó trở thành giá trị do code quản lý; không paste lại ARRAYFORMULA vào chúng.

## Dữ liệu phụ trợ

- Tháng thêm cuối vùng nghiệp vụ mỗi full: Booking cột Q, File đơn cột S, Chứng từ cột R.
- AD:AE ở 6 tab là ID và dấu vết đối soát, được ẩn. Không sửa/xóa/paste đè các cột này.
- Filter của full bao gồm đến AE để khi sort qua filter, ID đi cùng dữ liệu. Các cột kỹ thuật và khoảng trống được ẩn. Không sort riêng một vài cột của bảng.
- _PN_REQUESTS: nhật ký yêu cầu và kết quả các bước, ẩn.
- _PN_BACKUP_*: snapshot trước lần dọn gần nhất của mỗi tab, ẩn. Không thay thế backup toàn file.
- __PN_ITEM_ID: cột ẩn thêm cuối Hoàn sản phẩm để chống thêm trùng khi gửi lại.

File đơn có mã lặp nên không gộp/xóa các dòng đó tự động. Lần migration giữ từng dòng và cấp ID riêng; chạy lại không thêm bản trùng do migration. Cách lookup theo mã vẫn theo nghiệp vụ gốc. Khi nhập một lô hoàn toàn mới có cùng mã, code không tự suy đoán đó là dòng cũ cần xóa.

## Điều kiện triển khai bắt buộc

Tất cả 12 file .gs phải nằm chung project Apps Script đang phục vụ URL API hiện tại. Không tạo deployment URL mới. Lý do: ScriptLock chỉ phối hợp được trong cùng project và các lời gọi nghiệp vụ phải có mặt trong project đó.

Nếu 9 script đang chia nhiều project: tắt các trigger cũ tương ứng trong từng project sau khi sao lưu, trước khi bật bộ mới. pnInstallTriggers chỉ xóa/thay thế các handler đã biết trong chính project đang chạy, không thể tắt trigger ở project khác/tài khoản khác.

Không để hai bản cùng định nghĩa một hàm (ví dụ cả Code.gs cũ lẫn 01_API.gs mới) trong một project. Thay file hiện hữu bằng nội dung tương ứng; nguyên bản vẫn có trong originals và lịch sử deployment.

Các công thức/pivot/báo cáo bên ngoài 9 đoạn code nếu còn trỏ trực tiếp tab chính chỉ nhìn được vùng đang vận hành. Bảng pivot bên phải File đơn được giữ nguyên; nó tiếp tục phản ánh tab chính. Nếu cần báo cáo toàn lịch sử, chủ động đổi nguồn pivot đó sang File đơn_full. Không tự đổi nghiệp vụ báo cáo ngoài phạm vi đã xác nhận.

## Trình tự thay mã và chuyển dữ liệu

1. Tạo bản sao spreadsheet nguồn và VHFF, lưu phiên bản deployment API hiện tại. Chọn một khoảng bảo trì ngắn, ngừng chỉnh sheet và gửi app trong lúc chuyển lần đầu.
2. Tắt trigger cũ liên quan ở các project khác. Thay 9 file tương ứng trong project API, thêm 00_FullStore.gs, 10_Formulas_Triggers.gs, 11_API_FullSave.gs. Không dán nguyên bản và bản mới cùng lúc.
3. Chạy pnMigrateFull(). Hàm chép/đối soát, chuyển bốn công thức, thêm metadata/filter; KHÔNG dọn tab chính và KHÔNG ghi VHFF. Hàm chạy lại được nếu gián đoạn.
4. Chạy pnAuditFull(). Chỉ tiếp tục nếu ok=true, errors rỗng ở cả ba cặp. Kiểm tra các số liệu full và ảnh/ngày/checkbox; không chỉ nhìn số dòng.
5. Chạy pnScheduledReconcile() để cập nhật phần còn thiếu và đồng bộ VHFF. Kiểm tra PN_VHFF_ERROR không còn, PN_VHFF_DIRTY đã được xóa và số liệu file đích đúng.
6. Trong Deploy → Manage deployments, sửa deployment hiện tại, chọn New version. Giữ nguyên deployment ID/URL /exec, quyền truy cập và cách thực thi hiện tại.
7. Phát hành bản frontend/index.html qua đúng quy trình xuất bản app đang sử dụng. File nguồn dùng để tạo bản này là github-publish/index.html; cần xác nhận đây là bản được build vào app đang phát hành, không thay nhầm một bản lưu cũ. Không clear localStorage/IndexedDB hoặc hàng đợi ảnh.
8. Chạy pnInstallTriggers(). Kiểm tra có pnHandleEdit, pnOpen, pnScheduledReconcile (5 phút), pnCleanupDaily (khoảng 01 giờ, Asia/Saigon). Dọn tự động vẫn chưa được bật.
9. Thực hiện các ca nghiệm thu thật bên dưới. Chạy pnCleanupPreview() để xem số dòng sẽ dọn, audit và pendingRequests.
10. Sau khi nghiệm thu đạt, chạy pnEnableCleanup(). Đây là bước bật dọn; không cần clear thủ công. Có thể chạy pnCleanupDaily() một lần trong thời gian bảo trì để kiểm tra kết quả, sau đó để lịch chạy.

## Nghiệm thu trên Google/điện thoại thật

- Nhập một Booking mới: chứng từ chính/full cùng có đơn, app tải được.
- Sửa ngày Booking: ngày chứng từ/full cập nhật; không thêm trùng đơn.
- Import File đơn sau: mã GHTK/khu vực được bổ sung cả cho hồ sơ lịch sử liên quan.
- App lưu chứng từ và sản phẩm: kiểm tra ảnh, số lượng hàng hoàn, đúng đơn ở full/chính.
- Gửi lại cùng clientId: không thêm ảnh/dòng hàng hoàn.
- Đơn cũ chỉ còn full: tra được, cập nhật được, không tự xuất hiện lại chính nếu vẫn đã nhận/hủy.
- Đơn cũ chuyển chưa nhận: trở lại chính.
- Nhập ngày bàn giao ở full: cập nhật báo cáo Bàn giao chứng từ và VHFF Đã bàn giao CT.
- Kiểm tra hai cột Ecom nhập tay vẫn nguyên sau sync.
- Kiểm tra Hoàn sản phẩm vẫn lấy ngày booking khi Booking chính đã dọn.
- Giữ một thao tác offline qua lượt dọn rồi gửi: lưu đúng mã dù rowNumber cũ.
- Chặn quyền file VHFF trên bản thử: thao tác đã lưu còn nguyên, dọn bị chặn, sync chạy lại được khi phục hồi quyền.
- Kiểm tra pivot File đơn!T:W không bị dọn.

Không thể gọi các kiểm thử mô phỏng là kiểm thử production. Chưa có đo latency/quota trên deployment thật.

## Cách xử lý lỗi, không nhập lại dữ liệu

- pnDisableCleanup(): tắt dọn, không xóa dữ liệu.
- pnAuditFull(): xem bản chính/full còn lệch ở đâu.
- pnScheduledReconcile(): chạy lại đồng bộ, bao gồm VHFF.
- PN_VHFF_ERROR trong Script Properties: tên bước lỗi và nguyên nhân. Không đánh dấu app thất bại chỉ vì file đích đang lỗi.
- Có PN_CLEANUP_PENDING: API/mirror tạm dừng, giữ cache/hàng đợi. Chạy pnResumeCleanup() để hoàn tất từ backup. Nếu có sửa tay sau lúc lỗi, hàm từ chối ghi đè và cần đối chiếu thủ công.
- Upload có trạng thái RUNNING nhưng chưa có kết quả: không tự upload lại vì có thể Drive đã tạo file. Kiểm tra file đã tạo trước khi cập nhật checkpoint tương ứng trong _PN_REQUESTS; không xóa clientId rồi gửi lại mù.
- Xung đột hai bản: đối chiếu dòng được báo, xác định giá trị đúng, làm cho hai bản khớp rồi chạy lại. Không xóa ID/base để ép hệ thống ghi đè.
- Full bị giảm số hồ sơ: dừng để kiểm tra bản backup; không hạ bộ đếm PN_FULL_COUNT_* để bỏ qua mất dữ liệu.

## Dọn có những bảo vệ gì?

Chỉ dọn sau khi chính/full khớp, không còn yêu cầu app PENDING và VHFF đã xác nhận. Có snapshot và dấu vết phục hồi trước khi dồn dữ liệu. Chỉ sửa vùng nghiệp vụ (Booking A:P, File đơn A:R, Chứng từ A:Q) và AD:AE; không deleteRows cả sheet nên File đơn!T:W không bị dịch/xóa.

ScriptLock không chặn thao tác tay trên Google Sheets. Chọn giờ ít người sửa; mã kiểm tra thay đổi trước/sau dọn và giữ trạng thái phục hồi nếu phát hiện lệch. Không nên nhập tay đồng thời với việc dồn bảng.

Giữ full đầy đủ trong cùng spreadsheet giúp gọn vùng thao tác nhưng không giảm tổng dữ liệu workbook. Khi lịch sử tăng nhiều, full và các báo cáo vẫn cần được theo dõi thời gian chạy.

## Rollback

Trước khi bật dọn: tắt trigger bộ mới, chuyển deployment về phiên bản cũ và phục hồi công thức/bản sao sheet nếu cần. Không bật hai bộ trigger song song.

Sau khi đã dọn: không chỉ rollback API cũ vì API cũ không đọc full. Tắt dọn và trigger mới, khôi phục đủ dữ liệu tab chính từ bản sao/backup đã đối soát trước khi phục hồi API cũ. Không xóa full hoặc _PN_REQUESTS trong quá trình này.

## Chạy kiểm thử local

    node build.cjs
    node patch-frontend.cjs
    node tests/run.cjs

build.cjs và patch-frontend.cjs đọc bản gốc trong originals/ bằng đường dẫn tương đối. Tải toàn bộ thư mục về máy khác có Node.js là có thể build/test; không cần thư mục đính kèm của cuộc trò chuyện. Không cần build lại chỉ để dùng các file .gs đã xuất. Chạy node review-diff.cjs sau thay đổi để cập nhật diff và checksum (cần Git).
