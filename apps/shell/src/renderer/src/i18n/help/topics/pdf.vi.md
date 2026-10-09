# PDF: đọc, chú thích và che nội dung

Trình soạn thảo PDF có năm thẻ ruy-bâng: **Trang chủ / Chú thích / Chỉnh sửa / Trang / Xem**. Nó vừa đọc vừa ghi: văn bản sửa được, nội dung có thể che, có thể ký và có thể điền biểu mẫu.

## Đọc và điều hướng

- Thanh bên trái: **hình thu nhỏ** (nhấp để nhảy, vùng đang hiện được đánh dấu) hoặc **mục lục** (bookmark, nếu có).
- Thu phóng: bộ điều khiển tỉ lệ ở góc dưới bên phải; ctrl+lăn chuột thay đổi từng bậc.
- Xoay: theo từng trang hoặc tất cả các trang từ menu Trang; góc xoay được ghi lại khi lưu.
- Tìm kiếm: ctrl+F tìm trong toàn văn, tô sáng mọi kết quả.
- PDF có mã hóa: một lời nhắc mật khẩu (một cửa sổ nhỏ riêng) mở tệp; mật khẩu chỉ dùng cho phiên làm việc này.

## Chọn văn bản và đánh dấu (Chú thích)

Hãy thử trên bất kỳ đoạn nào:

1. **Kéo chuột qua một câu** — khi thả, một thanh chú thích nổi lên phía trên:

![Thanh chú thích sau khi chọn văn bản](img/pdf-highlight.png)

2. Chọn **Tô sáng** (ô vuông vàng mở bảng màu), **Gạch chân** hoặc **Gạch ngang**; **Hỏi AI** gửi phần đã chọn kèm câu hỏi của bạn tới bảng AI.
3. Để bỏ một dấu đánh, hãy kéo chọn lại đoạn văn đó rồi nhấp nút đang bật trên thanh (một nút bật/tắt kiểu Word), hoặc chọn rồi nhấn Delete.

Lưu ý:

- Kéo trên văn bản sẽ hiện một thanh bật lên: **Tô sáng / Gạch chân / Gạch ngang / Sao chép / Hỏi AI**.
- Màu lấy từ bảng màu; **áp dụng lại cùng một dấu đánh lên một đoạn đã đánh dấu là gỡ bỏ dấu đó** (nút bật/tắt kiểu Word).
- Những dấu đánh đã lưu vào tệp cũng có thể chọn rồi xóa (menu ⋯ hoặc Delete).
- **Lưu ý**: khi một công cụ vẽ đang được trang bị, lớp văn bản không thể chọn — công cụ tự ngắm sau mỗi lần đặt, nên bạn trở lại chế độ chọn cho thao tác tiếp theo.

## Công cụ vẽ (Chú thích)

Sáu công cụ: **Vẽ, Hình chữ nhật, Hình elip, Mũi tên, Ghi chú**, cộng với **hộp che nội dung** ở thẻ Chú thích.

- Mỗi công cụ đều là một nút bật/tắt: nhấp để trang bị; **nó tự ngắm sau khi một hình được đặt xuống** (nhấp lại vào công cụ nếu muốn vẽ tiếp); nhấp vào công cụ đang bật cũng để tắt.
- Nét vẽ bám theo độ rộng bút; hình chữ nhật/elip/mũi tên được kéo ra; màu lấy từ bảng màu vẽ.
- Hình đã đặt có thể chọn, xóa, kéo, và (chữ nhật/elip) đổi kích thước.
- **Che nội dung, quy trình đầy đủ** (để giấu một dòng văn bản):

  1. Thẻ Chú thích ▸ nhấp **Che vùng nội dung** (công cụ được trang bị).
  2. **Kéo một khung lên nội dung cần che** — phần bên trong được phủ bằng dấu gạch chéo, và thanh công cụ có thêm hai nút **Xóa các dấu đã đánh / Áp dụng che nội dung**:

  ![Trang sau khi đánh dấu che nội dung](img/pdf-redact.png)

  3. Nhấp **Áp dụng che nội dung** và xác nhận — kết quả là một bản làm việc trong đó văn bản và hình ảnh bị che đã bị gỡ bỏ hẳn (không phải phủ lên) và không thể hoàn tác; tài liệu gốc hoàn toàn không đổi.

  Làm nhầm? Nút xóa các dấu đã đánh sẽ xoá dấu hiện tại để bạn vẽ lại.

## Ghi chú dán và chuỗi bình luận

- **Công cụ ghi chú** đặt một ghim và mở một thẻ ở lề để viết văn bản (tên tác giả có thể đặt); xác nhận thì lưu thành chú thích văn bản PDF tiêu chuẩn.
- Nhấp vào một ghim để mở chuỗi: **Trả lời** (chuỗi phẳng kiểu WPS/Acrobat), sửa bình luận của bạn, **Xóa** một bình luận hoặc cả chuỗi.
- Nội dung đang sửa vẫn còn cho tới khi lưu ghi lại văn bản mới vào chính chú thích đó trong tệp, nên chuỗi trả lời được giữ nguyên.

## Chỉnh sửa nội dung PDF (Chỉnh sửa)

- **Sửa văn bản**: nhấp vào văn bản để sửa theo từng khối (bộ xử lý pdfium; cố gắng khớp phông chữ).
- **Chèn văn bản**: đặt văn bản có thể tìm kiếm với phông chữ/cỡ/màu tuỳ chọn.
- **Chèn hình ảnh / dấu đóng dấu**.
- Biểu mẫu: điền trực tiếp vào các trường AcroForm; giá trị được ghi khi lưu.

## Chữ ký

- **Chữ ký viết tay**: vẽ ra; có thể gắn vào một trường chữ ký trong biểu mẫu.
- **Chữ ký dạng ảnh**: đặt một hình ảnh làm chữ ký.
- Chữ ký đã lưu có thể dùng lại.

## Thao tác trang (Trang)

- **Xoay / xóa / sắp xếp lại**: kéo hình thu nhỏ để đổi thứ tự; xóa sẽ hỏi xác nhận.
- **Nhập trang**: kéo các trang từ một tệp PDF khác vào tài liệu. **Chèn trang trắng** thêm một trang trắng.
- **Thay thế trang** thay cả một đoạn bằng các trang từ nơi khác; **Cắt xén trang** cắt bỏ phần thừa, với tùy chọn áp dụng cho tất cả các trang.
- **Kích thước trang** đưa mọi trang về cùng một khổ giấy; **Đảo thứ tự** lật ngược tài liệu từ cuối về đầu.
- **Trích xuất trang**: xuất các trang đã chọn ra một tệp PDF mới.
- **Tách PDF**: có hai dạng — tách thành nhiều tệp theo từng đoạn, hoặc cắt mỗi trang thành lưới các trang nhỏ hơn.
- **Gộp PDF**: có hai dạng — nối thêm các PDF khác, hoặc gộp nhiều trang lên cùng một tờ. Kích thước được cộng lại **trước khi bất kỳ nội dung nào được đọc**, và **tổng trên 1 GiB sẽ bị từ chối** kèm thông báo dễ hiểu (để giữ mức dùng bộ nhớ có kiểm soát).
- Thay đổi ở mức trang được ghi lại ở lần lưu tiếp theo; Lưu dưới dạng giữ nguyên tệp gốc.

## Xuất và in

- **Xuất dưới dạng Word… / PowerPoint… / Excel…** trong menu Tệp, hoặc lấy đúng ba mục đó từ **Chuyển đổi PDF** trên dải ruy-bâng — tất cả đều tại máy, không tải lên. Tệp .pptx ra một slide mỗi trang và tệp .xlsx ra một bảng tính mỗi trang. Mỗi mục đều hỏi lưu ở đâu.
- **In**: qua hộp thoại hệ thống, theo thứ tự trang và góc xoay hiện tại; có thể chọn phạm vi trang.

## Lưu

- Lưu thường / tự động lưu ghi lại chú thích và chỉnh sửa vào tệp (ghi nguyên tử).
- **Che nội dung luôn đi qua quy trình "Áp dụng" của nó** để tạo ra một bản sao, giữ nguyên tệp gốc, nhờ đó nội dung nhạy cảm không nằm lại trong đó.
