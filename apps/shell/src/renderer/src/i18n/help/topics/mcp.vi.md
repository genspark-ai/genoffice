# Kết nối tác nhân lập trình

GenOffice nói chuyện theo Model Context Protocol, nên một tác nhân lập trình có thể đọc, ghi và kết xuất tài liệu của bạn qua đúng những engine mà ứng dụng dùng. Tác nhân không phải đoán mò một định dạng tệp: nó nhận lược đồ thao tác có kiểu từ đúng những định nghĩa mà bộ thực thi dùng để kiểm tra.

## Đăng ký

Trường hợp thông thường chỉ cần một lệnh:

```sh
genoffice mcp install all
```

Lệnh này tìm các tác nhân lập trình có trên máy này — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — rồi ghi mục máy chủ stdio vào cấu hình riêng của từng tác nhân, giữ nguyên phần còn lại của tệp đó như cũ.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Một tác nhân bạn cài ở chỗ không mấy quen thuộc cần `--dir <path>`; còn `--force` ghi đè một mục đã có sẵn.

## Tự chạy nó

Với một ứng dụng trên máy khác, hãy phục vụ qua HTTP thay thế:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` đổi nơi nó lắng nghe. Hãy truyền cùng mã thông báo đó cho ứng dụng.

Qua HTTP, tệp cũng di chuyển: `PUT /files/<name>` tải lên một tệp, mọi công cụ đều nhận một URL `http(s)` thay cho đường dẫn, và kết quả trả về dưới dạng URL tải xuống — và, khi đủ nhỏ, dưới dạng tài nguyên nhúng. Các thao tác, đặc tả và Markdown đều được truyền nội tuyến trong cả hai trường hợp.

## Tác nhân nhận được gì

Mỗi lệnh là một công cụ. Những cái đáng chú ý:

- **`docs`, `sheet`, `slides`** — đọc và sửa một tệp qua lối ghi của chính ứng dụng, mỗi lần một **thao tác** (op). Một bộ trình chiếu mới đi theo `deck_start`, `deck_page`, `deck_build`.
- **`render`** — một PNG cho mỗi trang, được bộ kết xuất của ứng dụng bố cục, để tác nhân có thể nhìn một trang chiếu thay vì đoán mò.
- **`pdf`** — lớp văn bản của PDF, từng trang một, không cần tiến trình ứng dụng.
- **`info`** — siêu dữ liệu và phần tóm tắt cấu trúc, thường là lời gọi rẻ nhất khi gặp một tệp lạ.
- **`search`, `image`, `media`** — các nhà cung cấp đã cấu hình trong ứng dụng, nên tác nhân không cần khóa riêng.
- **`merge`** — điền một mẫu `{{key}}`.

## Lược đồ, và một hạn mức nhỏ hơn

`apply` và `create` công bố các tham số `ops`, `cells` và `data` kèm lược đồ có kiểu cho từng thao tác, sinh từ `genoffice guide <domain> --json`. Nó chính xác và không hề nhỏ. Một ứng dụng có cửa sổ ngữ cảnh chật hẹp có thể xin các mảng thuần thay vì thế:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Kỹ năng

Một tác nhân không biết từ vựng thao tác sẽ đoán mò. `genoffice skill` cài một kỹ năng GenOffice vào những tác nhân nó tìm thấy, mang theo tài liệu tham chiếu và các hướng dẫn thiết kế — đúng những tài liệu mà `genoffice guide` in ra.

## Nó không phải là gì

Máy chủ MCP đọc và ghi tệp. Nó không phải cửa sổ: không có bảng AI, và hộp thoại cập nhật trong ứng dụng không áp dụng. Nếu một bước nào đó cần cửa sổ, hãy mở tệp ra.