# Cài đặt, ngôn ngữ, chủ đề và tích hợp MCP

## Mở cài đặt

Hàng tài khoản ở góc dưới bên trái của Trang chủ mở bảng cài đặt (khi bạn chưa đăng nhập, hàng đó ghi "Đăng nhập"). Bảng đó có sáu mục: Tài khoản, Mô hình AI, Phương tiện AI & Tìm kiếm, Chung, Tích hợp và Giới thiệu.

![Cài đặt ▸ Chung, nơi có ngôn ngữ, chủ đề, tự động lưu và công tắc thống kê sử dụng](img/settings-general.png)

Việc cấu hình mô hình có bài riêng; trong **Phương tiện AI & Tìm kiếm**, bạn bật tạo hình ảnh, phân tích hình ảnh, phân tích video, tìm kiếm web và tìm kiếm tệp cục bộ cho từng nhà cung cấp.

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

**Tích hợp** là bảng nối GenOffice với một tác nhân lập trình, và nó có bài riêng: Kết nối tác nhân lập trình. Nói ngắn gọn — chọn một lối (dòng lệnh, hoặc MCP), làm theo phần đó, rồi mở một cuộc trò chuyện mới và hỏi.

![Cài đặt ▸ Tích hợp: ba bước, rồi đến các hàng kỹ năng và tuỳ chọn MCP](img/settings-integrations.png)

Dưới **Máy chủ HTTP cục bộ**, ứng dụng cũng có thể tự chạy máy chủ đó — công tắc bật và cổng —, còn **Nâng cao** thêm URL kiểm tra tình trạng và tệp nhật ký, thay vì để dành cho trợ lý. Nó chỉ lắng nghe trên localhost.

## Bảng tra nhanh dòng lệnh

| Lệnh               | Tác dụng                     |
| ------------------ | ---------------------------- |
| `genoffice <file>` | mở một tập tin               |
| `genoffice mcp`    | khởi động máy chủ MCP cục bộ |
| `genoffice --help` | mọi lệnh và tuỳ chọn         |
