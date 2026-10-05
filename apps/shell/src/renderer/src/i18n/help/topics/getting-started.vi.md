# Bắt đầu nhanh: giao diện và những điều cơ bản

GenOffice là bộ ứng dụng văn phòng chạy hoàn toàn trên máy của bạn: một cửa sổ, một hàng thẻ, chứa sáu trình soạn thảo — Docs (xử lý văn bản), Sheets (bảng tính), Slides (trình chiếu), PDF, Markdown và HTML. Tập tin đều là .docx / .xlsx / .pptx / .pdf thật, hoàn toàn tương thích với Word, Excel và PowerPoint. Không cần mạng.

## Tổng quan giao diện

![Màn hình Trang chủ](img/home-screen.png)

Cửa sổ gồm ba phần:

- **Dải thẻ (phía trên)**: mỗi tập tin đang mở là một thẻ. Thẻ Trang chủ ngoài cùng bên trái luôn tồn tại và không thể đóng; các thẻ còn lại là tài liệu của bạn. Nhấp đúp vào một thẻ để đổi tên tập tin ngay tại chỗ.
- **Vùng nội dung**: trình soạn thảo (hoặc Trang chủ) thuộc thẻ đang hoạt động.
- **Dải menu**: trên dải menu hệ thống của macOS, ở trên cửa sổ trên Windows/Linux. Các menu File/Edit/View đổi theo trình soạn thảo đang hoạt động.

## Tạo tài liệu

Bất kỳ cách nào sau đây:

- Nhấp vào một thẻ tạo nhanh trong mục [Bắt đầu nhanh](help://getting-started) của **Trang chủ** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **File ▸ Mới**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML hoặc PDF.
- Kéo một tập tin vào cửa sổ, hoặc nhấp đúp vào nó trong trình quản lý tập tin (nếu GenOffice là ứng dụng mặc định).

Tài liệu mới mở ra chưa có tiêu đề; tập tin trên đĩa chỉ được tạo ở lần lưu đầu tiên.

## Mở tập tin

- Menu **File ▸ Mở** (⌘O/ctrl+O) mở bộ chọn của hệ thống: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Nhấp vào bất kỳ mục nào trong danh sách **Gần đây** của Trang chủ.
- `genoffice <file>` từ terminal cũng mở được tập tin.

## Mô hình lưu

- **Lưu thủ công**: ⌘S/ctrl+S, hoặc File ▸ Lưu / Lưu dưới dạng. Lần lưu đầu tiên sẽ hỏi vị trí và tên.
- **Tự động lưu** chỉ bật sau khi bạn đã lưu tay tập tin đó ít nhất một lần — một tập tin PDF bạn chỉ đọc sẽ không bao giờ bị ghi lại trong im lặng. Tự động lưu chạy ngay sau khi nội dung thay đổi.
- Đóng một thẻ có thay đổi chưa lưu sẽ hỏi Lưu / Bỏ / Hủy trước.
- Mọi lần ghi đều diễn ra theo kiểu nguyên tử (tập tin tạm + đổi tên), nên mất điện không thể để lại một tập tin dở dang.

## Phím tắt thường dùng

| Thao tác            | macOS | Windows / Linux |
| ------------------ | ----- | --------------- |
| Tài liệu mới        | ⌘N    | ctrl+N          |
| Mở                 | ⌘O    | ctrl+O          |
| Lưu                | ⌘S    | ctrl+S          |
| Đóng thẻ           | ⌘W    | ctrl+W          |
| Mở sổ tay này      | F1    | F1              |

Các phím tắt bên trong từng trình soạn thảo (công cụ định dạng, tìm và thay thế, thao tác bảng, ...) nằm trong chương tương ứng; Docs còn có hộp thoại phím tắt bàn phím có thể tìm kiếm (xem chương của nó).

## Nên đọc tiếp ở đâu

- Tập tin nằm ở đâu: [Màn hình Trang chủ](help://home-screen).
- Quản lý nhiều tập tin đang mở: [Quản lý thẻ và cửa sổ](help://tabs-and-windows).
- Để AI làm việc giúp: [Bảng điều khiển trợ lý AI](help://ai-panel).
- Ngôn ngữ, chủ đề, ứng dụng mặc định: [Cài đặt, ngôn ngữ, chủ đề và tích hợp MCP](help://settings-integrations).
