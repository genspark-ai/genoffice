# Mô hình AI và cài đặt

## Nhà cung cấp và mô hình

Mô hình và khóa được cấu hình trong Cài đặt (hàng tài khoản ở góc dưới bên trái của Trang chủ):

![Cửa sổ Cài đặt](img/settings-general.png)

- **Genspark đám mây**: đăng nhập (luồng mã thiết bị) rồi dùng — không phải cấu hình gì.
- **Điểm cuối tùy chỉnh (BYOK)**: Cài đặt ▸ AI nhận một URL cơ sở và một khóa API cho từng giao thức — tương thích OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen) và hơn thế. Khóa được lưu trong tệp cài đặt của ứng dụng trên máy này và chỉ được gửi trong phần đầu yêu cầu.
- Có thể chọn mô hình khác nhau cho từng năng lực: trò chuyện/tạo sinh, tạo hình ảnh, phân tích hình ảnh.
- **Kiểm tra kết nối**: xác minh điểm cuối truy cập được và mô hình hiển thị trước khi lưu.
- URL cơ sở có thể kèm đường dẫn và chuỗi truy vấn (kiểu cổng kết nối); các đường dẫn điểm cuối được nối thêm một cách chính xác.

## Tích hợp CLI (kiểu Codex)

- Cài đặt nhận một đường dẫn tới chương trình CLI cục bộ (thư mục home không phải ASCII và tiền tố ~ đều dùng được; ~ được mở rộng tự động); phát hiện mô hình sẽ dò các mô hình mà CLI cung cấp.
- Kiểm tra chỉ xác nhận sự tồn tại — không hạn chế bảng ký tự.

## Thay đổi có hiệu lực khi nào

- Thay đổi mô hình và điểm cuối có hiệu lực ngay lập tức; một cuộc trò chuyện đang diễn ra vẫn dùng cấu hình cũ cho tới lượt kế tiếp của nó.
