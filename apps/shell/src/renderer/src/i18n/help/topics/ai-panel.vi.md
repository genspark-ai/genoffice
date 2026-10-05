# Bảng điều khiển trợ lý AI

Mọi trình soạn thảo đều có thể gọi bảng AI: chọn một nội dung, đưa ra chỉ dẫn, xem kết quả hiện dần.

## Mở và sử dụng

![Bảng AI trong Docs](img/ai-panel.png)

- Các lối vào: **nút AI** trên dải ruy-bâng của từng trình soạn thảo, **Hỏi AI** trong menu ngữ cảnh, hoặc Hỏi AI trên thanh đánh dấu.
- Mô tả công việc bằng ngôn ngữ tự nhiên (viết lại đoạn này / biến cột này thành phần trăm / bố cục lại trang này...) rồi nhấn Enter.
- Câu trả lời hiển thị **dạng luồng**; khi AI cần công cụ (đọc tài liệu, sửa tài liệu, chạy script) thì nó thực thi và tiếp tục cho tới khi xong.
- **Dừng**: có thể ngắt lượt hiện tại bất cứ lúc nào.

## AI có thể làm gì

- **Docs**: viết lại/mở rộng/dịch/tóm tắt, chèn bảng và hình ảnh, điều chỉnh định dạng; mỗi lượt đều chụp ảnh trước.
- **Sheets**: công thức, điền dữ liệu, biến đổi hàng loạt, định dạng.
- **Slides**: tạo cả bộ trang chiếu, điều chỉnh bố cục, viết lại nội dung.
- **PDF**: hỏi đáp và tóm tắt dựa trên văn bản hoặc trang đã chọn.
- **Markdown / HTML**: viết lại, mở rộng, dịch.

## Quay lui và an toàn

- Bảng của Docs giữ một **danh sách phiên bản**: mỗi lượt một ảnh chụp, quay lui về bất kỳ ảnh nào, và bản thân thao tác quay lui cũng quay lui được bằng Ctrl+Z. Ảnh chụp vẫn còn sau khi mở lại tài liệu.
- Các chỉnh sửa của AI đi qua đúng quy trình chỉnh sửa như chỉnh sửa thủ công (hoàn tác được, phải qua xác nhận lưu của bạn) — không có gì lách qua bước xác nhận lưu.

## Quyền riêng tư

- Chỉ dẫn và nội dung tài liệu liên quan được gửi tới **dịch vụ mô hình mà bạn đã cấu hình** (Genspark đám mây hoặc một điểm cuối tùy chỉnh, xem chương sau); không cấu hình thì không gửi gì.
- Tập tin cục bộ không được tải lên nơi khác; khóa BYOK chỉ nằm trong phần đầu yêu cầu — không bao giờ nằm trên đĩa hay trong nhật ký.
