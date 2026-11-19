# TH Truemilk (FFPL | 3PL_TH Truemilk_FF) — Apps Script

Bản trên máy: `D:\API_APP Chứng từ\TH truemilk\` (.txt). Giữ 2 nơi giống nhau.

| File GitHub | File trên máy |
|---|---|
| Vanhanh.gs | Vanhanh.txt |
| Code_API.gs (cũng chép ra `../Code_api_TH.gs`) | Code API.txt |
| Sync_hoan.gs | Sync hoàn.txt |
| Mail_loi.gs | Mail_loi.txt |
| appsscript.json | appscript.json.txt |

Vận hành (2026-11-19): mỗi tab một luồng riêng, chỉ trigger sửa (`handleSpreadsheetEdit`);
không trigger 5 phút; khóa chỉ lúc ghi Tổng hợp all/STATE/Kiểm kê/BC/SCUTT.
App nhập hoàn → API hẹn 1 lượt `ffplReturnsJob` (tính HOÀN + sync SCUTT).
Sau khi dán code: chạy `ffplInstallTriggers()` một lần.
