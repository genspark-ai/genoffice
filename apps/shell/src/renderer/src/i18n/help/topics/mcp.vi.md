# Kết nối tác nhân lập trình

GenOffice nói chuyện theo Model Context Protocol, nên một tác nhân lập trình có thể đọc, ghi và kết xuất tài liệu của bạn qua đúng những engine mà ứng dụng dùng. Tác nhân không phải đoán mò một định dạng tệp: nó nhận lược đồ thao tác có kiểu từ đúng những định nghĩa mà bộ thực thi dùng để kiểm tra.

## Đăng ký từ trong ứng dụng

Việc này làm ở **Cài đặt ▸ Tích hợp**. Bảng đó có hai nửa, và bạn có thể dùng một nửa hoặc cả hai.

**Kỹ năng.** Mỗi tác nhân lập trình tìm thấy trên máy này một dòng — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, mỗi dòng có **Cài đặt**, **Cập nhật** và **Gỡ cài đặt**, cộng thêm **Cài đặt vào một thư mục khác…**, **Tải xuống kỹ năng (zip)** và **Sao chép đường dẫn**. Nếu trợ lý của bạn không có trong danh sách, hãy chỉ cho GenOffice thư mục mà nó đọc `SKILL.md`, hoặc lưu tệp zip lại rồi bảo trợ lý tự cài đặt. Kỹ năng và MCP có thể nằm cạnh nhau: trợ lý chọn một trong hai, và chúng làm đúng những việc giống nhau.

**MCP.** Có hai đường: **Được khởi động bởi trợ lý (khuyến nghị)**, ở đó bạn thêm cấu hình được hiển thị vào phía khách của mình và trợ lý tự khởi chạy máy chủ, và **Máy chủ HTTP cục bộ**, do ứng dụng chạy thay bạn. Dù chọn đường nào, trợ lý vẫn kết nối với GenOffice và bạn không phải gõ lệnh nào.

## Đăng ký từ dòng lệnh

Cũng làm được từ terminal — đây là đường nâng cao, dùng khi tác nhân nằm ở nơi bảng kia không tìm thấy:

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

Mọi thứ mà máy chủ nhận đều nằm trong một màn hình — các dạng đăng ký, gỡ bỏ và liệt kê, phục vụ qua HTTP, cùng hai cờ lược đồ:

![Kết quả thật của genoffice mcp --help: các dạng install, uninstall và list, kèm các tùy chọn --http, --host, --token, --compact-schemas, --dir và --force](img/mcp.png)

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

## Vì sao cần kỹ năng

Một tác nhân không biết từ vựng thao tác sẽ đoán mò. Kỹ năng mang theo tài liệu tham chiếu và các hướng dẫn thiết kế — đúng những tài liệu mà `genoffice guide` in ra —, nên trợ lý viết ra thao tác mà nó đã thực sự đọc đặc tả. Cài nó từ bảng kia ở trên, hoặc bằng `genoffice skill` trong terminal.

## Nó chạm được gì trong ứng dụng

Máy chủ không chỉ dừng ở các tệp trên đĩa. Chừng nào GenOffice còn chạy, tác nhân cũng có thể làm việc qua cửa sổ:

- **`open_in_genoffice`** mở một tệp trong một thẻ và đưa thẻ đó lên tiền cảnh.
- **`open_documents`** liệt kê mọi tài liệu bạn đang mở — id, loại, đường dẫn và có thay đổi chưa lưu hay không —, rồi đọc nội dung hiện tại của một tài liệu hoặc đóng nó lại, lưu trước trừ khi bạn bảo nó bỏ thay đổi.
- **Các công cụ nội dung** nhận id (hoặc đường dẫn) đó làm tham số `document`, nên một thay đổi sẽ rơi vào đúng thẻ bạn đang mở và cửa sổ chuyển sang hiển thị nó.

Có hai thứ vẫn nằm ngoài tầm với: không có bảng AI, và hộp thoại cập nhật trong ứng dụng không áp dụng.

## Bảng Máy chủ HTTP cục bộ

Dưới **Máy chủ HTTP cục bộ**, ứng dụng tự chạy máy chủ thay vì để trợ lý lo: công tắc bật, ô **Cổng**, chỉ báo **Đang chạy / Không chạy** và **Tạo ngầm dưới nền** (ghi thẳng tài liệu ra một đường dẫn mà không mở giao diện), cùng **Ví dụ cấu hình máy khách** để sao chép. Mở **Nâng cao** sẽ thêm hai URL kết nối — Streamable HTTP và URL SSE cũ —, URL **Kiểm tra tình trạng** và công tắc **Ghi nhật ký** ghi lại hoạt động của máy chủ và các công cụ vào một tệp cục bộ mà bạn có thể **Mở**, **Làm mới** hoặc **Xóa** ngay tại đó. Nó chỉ lắng nghe trên localhost.
