# Bảng điều khiển trợ lý AI

Mọi trình soạn thảo đều có thể gọi bảng AI: chọn một nội dung, đưa ra chỉ dẫn, xem kết quả hiện dần.

## Mở và sử dụng

![Bảng AI trong Docs](img/ai-panel.png)

- Các lối vào: **nút AI** trên dải ruy-bâng của từng trình soạn thảo, **Hỏi AI** trong menu ngữ cảnh, hoặc Hỏi AI trên thanh đánh dấu.
- Mô tả công việc bằng ngôn ngữ tự nhiên (viết lại đoạn này / biến cột này thành phần trăm / bố cục lại trang này...) rồi nhấn Enter.
- Câu trả lời hiển thị **dạng luồng**; khi AI cần công cụ (đọc tài liệu, sửa tài liệu) thì nó thực thi và tiếp tục cho tới khi xong.
- **Dừng**: có thể ngắt lượt hiện tại bất cứ lúc nào.
- **Hàng đợi**: khi phản hồi đang truyền, Enter đưa tin nhắn của bạn vào hàng đợi thay vì bỏ đi — dải phía trên ô nhập đếm chúng; mở ra để sửa, bỏ, đổi thứ tự, tạm dừng hay dọn sạch. Tin tiếp theo tự gửi khi lượt kết thúc; Esc vẫn dừng lượt như cũ.

## AI có thể làm gì

- **Docs**: viết lại/mở rộng/dịch/tóm tắt, chèn bảng và hình ảnh, điều chỉnh định dạng; mỗi lượt đều chụp ảnh trước.
- **Sheets**: công thức, điền dữ liệu, biến đổi hàng loạt, định dạng.
- **Slides**: tạo cả bộ trang chiếu, điều chỉnh bố cục, viết lại nội dung.
- **PDF**: hỏi đáp và tóm tắt dựa trên văn bản hoặc trang đã chọn.
- **Markdown / HTML**: viết lại, mở rộng, dịch.
- **Trang chủ ▸ Ask AI** trả lời từ các tệp bạn đánh sao và trích dẫn chúng — bảng cơ sở tri thức có mục riêng trong chương Trang chủ, và bảng AI của Docs cũng mang kỹ năng ấy.

## Quay lui và an toàn

- Bảng của Docs giữ một **danh sách phiên bản**: mỗi lượt một ảnh chụp, quay lui về bất kỳ ảnh nào, và bản thân thao tác quay lui cũng quay lui được bằng Ctrl+Z. Ảnh chụp vẫn còn sau khi mở lại tài liệu.
- Các chỉnh sửa của AI đi qua đúng quy trình chỉnh sửa như chỉnh sửa thủ công (hoàn tác được, phải qua xác nhận lưu của bạn) — không có gì lách qua bước xác nhận lưu.

## Ẩn vùng chọn khỏi AI

- Trong Slides, HTML và Markdown, một vùng chọn có thể được giấu khỏi mắt mô hình: chuột phải ▸ **Ẩn vùng chọn khỏi AI**, đặt tên chỗ giữ chỗ, và mô hình đọc `{{tên}}` thay cho chữ. (Markdown bật mục này bằng công tắc mắt-gạch-chéo trên ruy-băng; của HTML nằm ở bảng mã nguồn.)
- Tệp giữ nguyên văn thật — lưu, xuất và in không bị ảnh hưởng, dấu còn sống qua lần mở lại. Chính mục đó hủy việc ẩn (Markdown có thêm mục riêng **Hiện lại vùng chọn cho AI**).
- HTML ẩn nội dung phần tử, giá trị thuộc tính và hằng chuỗi trong script; trang đã render không bao giờ đổi.
- Slides còn ẩn cả hình, video và âm thanh trọn vẹn, và một slide có nội dung ẩn không bao giờ được gửi dạng chụp màn hình: kiểm tra bố cục tự động bỏ qua trang đó và nói rõ.
- Điều này chỉ dành cho AI trong ứng dụng — CLI, công cụ MCP và xuất headless đọc tệp trực tiếp nên vẫn thấy văn bản.

## Quyền riêng tư

- Chỉ dẫn và nội dung tài liệu liên quan được gửi tới **dịch vụ mô hình mà bạn đã cấu hình** (Genspark đám mây hoặc một điểm cuối tùy chỉnh, xem chương sau); không cấu hình thì không gửi gì.
- Tập tin cục bộ không được tải lên nơi khác; khóa BYOK được lưu trong tệp cài đặt của ứng dụng trên máy này và chỉ được gửi trong phần đầu yêu cầu.
