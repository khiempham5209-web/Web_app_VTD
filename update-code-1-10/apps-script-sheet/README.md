# Project Apps Script GẮN VỚI Google Sheet [FFPL]_Phạm Nguyên (project OCR)

Khác với `../apps-script/` (project độc lập `PJ_Phạm Nguyên`, 00 → 14). Project này gắn trực tiếp vào Sheet
nên mới thêm được menu trên thanh menu của Sheet.

| File trong Apps Script | Nội dung | Bản trong máy |
|---|---|---|
| `Ocr_img.gs` | Menu "OCR > Nhập từ ảnh": đọc ảnh phiếu nhập, ghi tab TT Nhập | `D:\API_APP Chứng từ\Code_OCR_TTNhap.txt` |
| `Call_delete.gs` | Menu "Delete > Xóa dữ liệu…": hộp thoại xóa Booking / Booking_full / Chứng từ_FF / Chứng từ_full | `D:\API_APP Chứng từ\PN_Menu_Xoa.txt` |

## Thư viện (bắt buộc cho `Call_delete.gs`)
- Thư viện: `PJ_Phạm Nguyên`, Script ID `1lti2ZBkm2tf3wcItV9hq5xugAwuafKo_RW9ItmbG6sfCrA1TsLCfE6CF`
- **Mã nhận dạng: `PN`** (hộp thoại gọi `PN.pnMenuXoaXemTruoc` / `PN.pnMenuXoaThucHien` trong `../apps-script/14_Menu_Xoa.gs`).
- Phiên bản: bản triển khai mới nhất của `PJ_Phạm Nguyên` (hoặc HEAD). Sửa `14_Menu_Xoa.gs` xong phải triển khai bản mới
  rồi đổi phiên bản thư viện ở đây (trừ khi dùng HEAD).
