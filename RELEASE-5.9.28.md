# VTD Chứng từ FF 5.9.28

## Hướng dẫn trong app
- Bổ sung hướng dẫn Nhập TH mới: Loại hoàn, tách mã từ link GHTK, Sửa / xóa đơn TH, Tìm kiếm TH.

## Tìm kiếm
- Ô chọn shop: VTD / Phạm Nguyên / TH Truemilk nằm trên 1 hàng.
- Tìm kiếm TH: Tên sản phẩm là ô chọn, lấy theo DS SKU TH (tự cập nhật khi DS SKU tải xong).

## Nhập TH
- DS SKU TH luôn đủ: API TH chưa có skuSync thì tự lấy bằng listSku. Lỗi tải thì lần mở app sau thử lại.

## Sửa đơn VTD
- SKU cũ: hiện ảnh sản phẩm hiện có, chụp / chọn thêm ảnh (ảnh mới được thêm vào SKU), quét Mã vật tư và Serial.
- **+ Thêm SKU**: form đầy đủ như màn nhập, gồm tìm hoặc tạo mới SKU, Type, quét Mã vật tư (tự điền tên + type), Tình trạng, SL, Mã lô/HSD, quét Serial, Ghi chú. Bắt buộc ít nhất 1 ảnh sản phẩm.
- Ảnh chứng từ khi sửa chụp được trên APK.
- Lưu sửa đi qua hàng đợi: mất mạng vẫn lưu, có mạng tự gửi và đánh dấu cần đồng bộ lại. Không có thay đổi thì không gửi gì.
- SKU của lượt nhập khác cùng mã đơn chỉ hiển thị, không bị ghi đè.
- Sửa lỗi đang sửa dở bị đẩy về màn chi tiết khi chi tiết đơn tải xong.

## API VTD (Code_chinh_VTD.gs), cần deploy lại
- Sửa lỗi **mất link ảnh sản phẩm** khi lưu sửa đơn. Trước đây mỗi lần sửa, SKU bị ghi lại không có ảnh.
- SKU trả về app kèm ID, RAW_ID, link ảnh.
- updateRecord: không gửi danh sách SKU thì giữ nguyên SKU. Ảnh sản phẩm gắn đúng SKU của chính RAW_ID. Gửi lại cùng một lần sửa thì không xử lý lần 2.

Cần deploy cả **API TH** (từ 5.9.27) và **API VTD** trước khi phát hành APK.
