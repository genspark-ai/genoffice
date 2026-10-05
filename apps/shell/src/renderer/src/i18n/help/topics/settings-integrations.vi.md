# Cài đặt, ngôn ngữ, chủ đề và tích hợp MCP

## Mở cài đặt

Hàng tài khoản ở góc dưới bên trái của Trang chủ mở bảng cài đặt (khi bạn chưa đăng nhập, hàng đó ghi "Đăng nhập"); các tuỳ chọn liên quan tới AI nằm trong mục Mô hình AI của bảng đó

![Cửa sổ Cài đặt](img/settings-integrations.png) — việc cấu hình mô hình được nói ở chương Mô hình AI và cài đặt.

## Ngôn ngữ

- Cài đặt có **21 ngôn ngữ giao diện**: tiếng Anh, Trung giản thể, Nhật, Hàn, Pháp, Đức, Tây Ban Nha, Thái, Indonesia, Nga, Ả Rập, Bồ Đào Nha, Ý, Ba Lan, Séc, Hà Lan, Mã Lai, Hebrew, Hindi, Trung Quốc Phồn thể, Việt.
- Việc đổi có hiệu lực ngay và được ghi nhớ; dải menu gốc được dựng lại theo ngôn ngữ đó.

## Chủ đề

Sáng / Tối / Theo hệ thống. Khi chọn Theo hệ thống, giao diện bám theo vẻ ngoài của hệ điều hành, và vùng trình soạn thảo đổi màu đồng bộ mà không nhấp nháy.

## Đăng ký ứng dụng mặc định

Cài đặt có thể đăng ký GenOffice làm ứng dụng xử lý .docx / .xlsx / .pptx / .pdf và các định dạng liên quan (đăng ký ứng dụng mặc định ở cấp nền tảng; xác nhận khi được hỏi).

## Thông báo bên thứ ba và cập nhật

- Trợ giúp ▸ Thông báo của bên thứ ba: danh mục giấy phép nguồn mở đầy đủ đi kèm ứng dụng.
- Trợ giúp ▸ Kiểm tra bản cập nhật: kích hoạt một lần kiểm tra thủ công; nếu có bản mới sẽ nhắc cài đặt.

## Đăng nhập Genspark

- Lối đăng nhập (trong cài đặt hoặc danh sách dự án đám mây) dùng luồng **mã thiết bị**: GenOffice hiển thị một mã và mở trang đăng nhập trong trình duyệt; xong thì nó tự tiếp tục.
- Đăng nhập chỉ dùng cho: danh sách dự án đám mây và các mô hình đám mây của Genspark. Không đăng nhập thì mọi tính năng cục bộ và mô hình tùy chỉnh vẫn hoạt động bình thường.
- Đăng xuất chỉ cần một cú nhấp trong cài đặt.

## Tích hợp MCP (dành cho người dùng nâng cao / ứng dụng AI)

GenOffice tích hợp sẵn một **máy chủ MCP** cục bộ để các ứng dụng AI bên ngoài (Claude Desktop, Cursor, ...) có thể đọc và ghi tài liệu của bạn trực tiếp:

- Khởi động: chạy `genoffice mcp` ở dòng lệnh (cổng và mã xác thực có thể cấu hình; mặc định chỉ lắng nghe trên loopback).
- Năng lực: tạo/mở/sửa docx, xlsx và pptx, đọc nội dung, chuyển đổi định dạng, xuất PDF và hơn thế — cùng bộ công cụ mà các ứng dụng máy tính dùng.
- Bảo mật: xác thực bằng mã là tuỳ chọn nhưng nên bật; trình lắng nghe mặc định không rời khỏi máy của bạn; xem `genoffice mcp --help`.

## Bảng tra nhanh dòng lệnh

| Lệnh                | Tác dụng                    |
| ------------------- | --------------------------- |
| `genoffice <file>`  | mở một tập tin              |
| `genoffice mcp`     | khởi động máy chủ MCP cục bộ |
| `genoffice --help`  | mọi lệnh và tuỳ chọn        |
