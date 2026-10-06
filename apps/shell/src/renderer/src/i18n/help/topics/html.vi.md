# Trình soạn thảo HTML

Trình soạn thảo HTML mở .html / .htm ở hai chế độ: **Xem trước** (trang đã kết xuất) và **Mã nguồn**.

- **Xem trước**: kết xuất thật; các bảng định kiểu và hình ảnh tương đối được nạp cạnh tệp.
- **Trình kiểm tra xem trước**: nhấp để chọn một phần tử, nhấp đúp để sửa văn bản tại chỗ, xóa bằng thanh công cụ, và hỏi AI về phần đã chọn.
- **Chế độ mã nguồn**: sửa HTML; ctrl+F để tìm, Thay thế tất cả lưu lại đánh dấu đã viết lại.
- **Lưu**: trung thành từng byte (giữ BOM/CRLF/dấu xuống dòng cuối); lưu mà không sửa gì sẽ không ghi lại.
- **Thu phóng**: ctrl+lăn chuột / véo tay để phóng bản xem trước; ctrl+Z bên trong bản xem trước hoàn tác chỉnh sửa gần nhất.

## Thanh công cụ

Nhấp vào bất kỳ phần tử nào trong bản xem trước, một thanh công cụ sẽ nổi lên phía trên:

![Thanh công cụ nổi trên một phần tử đang chọn](img/html-toolbar.png)

- **Tệp và lịch sử**: Lưu, Lưu dưới dạng, Hoàn tác, Làm lại, Tìm kiếm; công tắc **Tự động lưu** ghi thay đổi theo định kỳ.
- Chuyển đổi **Xem trước / Mã nguồn**; **Toàn màn hình** hiển thị trang ở chế độ toàn màn hình.
- **Định dạng**: in đậm, in nghiêng, tăng/giảm cỡ chữ; **bảng kiểu** cho phần tử đang chọn (màu sắc và hơn thế).
- **Chèn**: tiêu đề, đoạn văn, bảng, hình ảnh (theo liên kết), và hơn thế.
- **Thao tác hình ảnh** (khi đang chọn một hình ảnh): cắt, **Xóa nền**, thay thế, khóa tỷ lệ khung hình.
- **Thao tác phần tử** (khi chọn một phần tử trong trình kiểm tra xem trước): xóa, nhân bản, di chuyển lên/dưới.
- **Nút AI**: mở bảng AI; có thể hỏi trực tiếp về phần tử đang chọn.

## Xuất

Trình đơn Tệp, tất cả đều cục bộ và đều hỏi lưu kết quả ở đâu:

- **Xuất dưới dạng Word…** và **Xuất dưới dạng PDF…** ghi ra tệp .docx hoặc .pdf thật.
- **Xuất dưới dạng HTML một tệp…** ghi ra một tệp .html duy nhất có nhúng hình ảnh bên trong. Nó không ghi đè tệp bạn đang mở và cho biết có bao nhiêu hình ảnh không nhúng được.

## Chèn khung xương

Với một trang trắng, **Chèn ▸ Chèn khung xương** sẽ ghi ra một tài liệu tối giản ở chế độ tiêu chuẩn:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Mỗi phần đều có lý do, và vì vậy đây là một lệnh chứ không phải thứ gõ tay:

- **doctype**, nếu không có nó thì bản xem trước chạy ở chế độ quirks, nơi kích thước khung và bố cục bảng tuân theo những quy tắc khác với điều bạn mong đợi;
- **`lang`**, nếu không có nó thì trình đọc màn hình chẳng có ngôn ngữ nào để đọc trang, và trình duyệt sẽ chọn phông chữ cùng công cụ kiểm tra chính tả cho một ngôn ngữ sai;
- **charset**, nếu không có nó thì một trang có văn bản phi chữ Latin có thể hiện ra thành ký tự hỏng.

Thẻ viewport cố ý vắng mặt: nội dung này kết xuất trong một khung làm việc nền, không có khung xem di động nào để nó tác động.

`lang` bám theo ngôn ngữ giao diện của ứng dụng, nên khung xương bạn chèn vào chính là khung mà công cụ của bạn đã được thiết lập cho. Sau đó cứ tùy ý mà sửa.

Mục này chỉ hiện ở chế độ chỉnh sửa, và chỉ khi tài liệu còn trống — đã có nội dung thì không còn gì để chèn khung xương _vào_ nữa.
