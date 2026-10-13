# Thao tác tập tin: đổi tên, xóa, xuất

Chương này nói về các thao tác tập tin dùng chung cho mọi trình soạn thảo; các lựa chọn xuất riêng của từng trình soạn thảo nằm trong chương của nó.

Menu **⋯** của một hàng tập tin (rê chuột lên hàng tập tin) gom các thao tác này:

![Menu ⋯ của một hàng tập tin](img/file-ops.png)

## Đổi tên

Hai lối vào, một bộ kiểm tra duy nhất:

- **⋯ ▸ Đổi tên** trên hàng của Trang chủ.
- **Nhấp đúp vào một thẻ tập tin** để đổi tên tại chỗ (xem [Quản lý thẻ và cửa sổ](help://tabs-and-windows)).

Quy tắc: phần mở rộng được giữ tự động; ký tự không hợp lệ, dấu chấm ở cuối và tên dành riêng (CON/NUL và các tên tương tự) bị từ chối kèm thông báo; trùng tên trong cùng thư mục cũng bị chặn. Tập tin thật trên đĩa được đổi tên, và danh sách gần đây cùng các gắn sao đi theo.

## Xóa

- **⋯ ▸ Xóa** ở Trang chủ: hỏi những tập tin nào, rồi chuyển chúng vào **thùng rác hệ thống**, nơi hệ điều hành có thể khôi phục. GenOffice không giữ thao tác hoàn tác riêng cho việc này — khôi phục là việc của thùng rác, không phải của thông báo.

## Nhân bản

**⋯ ▸ Tạo bản sao** tạo một bản `<name> bản sao` trong cùng thư mục; khi trùng tên, một số thứ tự được thêm vào tự động. Bản sao được thêm vào **Gần đây** thay vì mở sẵn cho bạn — cách một cú nhấp, chứ không nằm ngay trước mặt.

## Lưu và Lưu dưới dạng

- **⌘S / ctrl+S** lưu tập tin hiện tại; một tập tin chưa có tiêu đề sẽ hỏi vị trí và tên trước.
- **Lưu dưới dạng** ghi một tập tin mới và để nguyên tệp gốc; các lần sửa sau đó nhắm vào tập tin mới.
- Mỗi lần lưu đều là nguyên tử (tập tin tạm + đổi tên); thoát giữa chừng lúc đang ghi không làm hỏng tệp.
- Tự động lưu chỉ bắt đầu sau lần lưu thủ công đầu tiên (xem [Bắt đầu nhanh](help://getting-started)).

## Xuất dưới dạng PDF

- **Docs**: File ▸ Xuất dưới dạng PDF (hoặc nút trên dải ruy-bâng), theo đúng cách ngắt trang hiện tại.
- **Slides**: việc xuất raster hóa từng trang, có thanh tiến trình cho các bộ lớn.
- **Sheets**: việc xuất theo cách ngắt trang in.
- Việc xuất được kết xuất trong một cửa sổ ẩn và lưu vào nơi bạn chọn.

## Xuất sang Word / hình ảnh

- **PDF ▸ Xuất dưới dạng Word…**: biến PDF thành .docx (chuyển đổi tại máy; bố cục phức tạp được xử lý ở mức cố gắng hết sức).
- **Docs** có thể xuất các trang thành hình ảnh (PNG từng trang).

## In

File ▸ In trong từng trình soạn thảo (⌘P/ctrl+P) mở hộp thoại in của hệ thống; PDF in theo thứ tự trang và góc xoay hiện tại.

## Tập tin chưa có tiêu đề nằm ở đâu

Vị trí bạn chọn ở lần lưu đầu tiên chính là "nhà" của nó; trước đó tài liệu chỉ tồn tại trong bộ nhớ. Tự động lưu chỉ tiếp quản sau lần lưu đầu tiên ấy.
