# Trình soạn thảo Markdown

Trình soạn thảo Markdown mở .md / .markdown với trải nghiệm "mã nguồn + bản xem trước đã kết xuất".

- **Mở**: từ Trang chủ hoặc File ▸ Mở; dòng lệnh cũng dùng được.
- **Chỉnh sửa**: chỉnh sửa văn bản thuần; các mở rộng GFM (bảng, danh sách công việc, gạch ngang, liên kết tự động) hiển thị trong bản xem trước.
- **Xem trước**: theo thời gian thực; các tài nguyên tương đối như hình ảnh được phân giải cạnh tài liệu.
- **Lưu**: trung thành từng byte — BOM, CRLF và việc có dấu xuống dòng ở cuối hay không đều được giữ; lưu mà không sửa gì sẽ không ghi lại tệp.
- **Tìm và thay thế**: ctrl+F tìm trong mã nguồn; Thay thế tất cả ghi lại vào tệp.
- **AI**: các nút mẫu cho phép trợ lý viết lại, mở rộng hoặc dịch tài liệu.

## Thanh công cụ

Một hàng nút phía trên trình soạn thảo ( rê chuột để xem chú thích):

![Thanh công cụ Markdown](img/md-toolbar.png)

- **Tệp và lịch sử**: Lưu, Lưu dưới dạng, Hoàn tác, Làm lại, Tìm kiếm; công tắc **Tự động lưu** ở bên phải ghi thay đổi xuống đĩa theo định kỳ.
- **Nút AI**: mở bảng AI; bên cạnh là các nút mẫu viết lại / mở rộng / dịch.
- **Kiểu đoạn văn** (thả xuống): chuyển giữa văn bản thường và các cấp tiêu đề.
- **Định dạng trong dòng**: **In đậm**, _In nghiêng_, ~~Gạch ngang~~, `Mã nội dòng`, liên kết.
- **Danh sách**: dấu đầu dòng, đánh số, danh sách công việc.
- **Chèn**: bảng, hình ảnh, đường phân cách.
- **Thuộc tính**: chèn hoặc nhảy tới khối YAML front matter ở đầu tệp.
- **Mục lục**: nhảy theo thứ bậc tiêu đề.
- **Chính tả**: bật/tắt kiểm tra chính tả cho tài liệu này.

Ba ví dụ nhanh:

- **Tiêu đề**: đặt con trỏ vào dòng đó ▸ thả xuống kiểu đoạn văn ▸ chọn "Tiêu đề 1".
- **Bảng**: nhấp **Chèn bảng** ▸ kéo để chọn số hàng và cột ▸ gõ vào các ô; bản xem trước kết xuất ngay lập tức.
- **Danh sách công việc**: chọn vài dòng ▸ nhấp **Danh sách công việc** ▸ mỗi dòng thành `- [ ]`, hiển thị thành ô tích trong bản xem trước.

## Xuất

Trình đơn Tệp, tất cả đều cục bộ và đều hỏi lưu kết quả ở đâu:

- **Xuất dưới dạng Word…** và **Xuất dưới dạng PDF…** ghi ra tệp .docx hoặc .pdf thật.
- **Xuất dưới dạng hình ảnh…** ghi một tệp PNG cho mỗi trang vào thư mục bạn chọn.
- **Chuyển đổi và mở trong Docs** chuyển sang .docx rồi mở trong thẻ Docs tích hợp ngay trong ứng dụng này — đây không phải chuyển giao cho thứ gì trên đám mây, và bản đã chuyển đổi nằm trong một thư mục bộ nhớ đệm được dọn dẹp sau khoảng một tuần.

## Chế độ mã nguồn

Trên dải ruy-bâng có công tắc **Mã nguồn** (được bản địa hóa cùng ứng dụng). Bật lên, trình soạn thảo được thay bằng Markdown thô: đúng văn bản mà một lần lưu sẽ ghi ra, không làm đẹp, không chuẩn hóa gì bên dưới bạn.

- **Soạn thảo trung thành từng byte.** Lưu từ chế độ mã nguồn tạo ra đúng những byte mà lưu từ trình soạn thảo tạo ra — BOM, CRLF và việc có dấu xuống dòng ở cuối đều được giữ nguyên.
- **Đây vẫn là cùng một tài liệu.** Chuyển qua lại tùy ý; mã nguồn chính là văn bản của trình soạn thảo, không phải một bản sao phải điều hợp.
- **Thanh định dạng không dùng được** khi chế độ xem này đang mở, vì phần lớn các nút đó chèn những cấu trúc riêng của trình soạn thảo mà chỉ có ý nghĩa ở phần đã kết xuất. Nó trở lại khi bạn đóng chế độ xem.
- **JSON và các tệp ở chế độ mã nguồn khác** mở thẳng tại đây: không có gì để kết xuất, nên mã nguồn _chính là_ tài liệu.
