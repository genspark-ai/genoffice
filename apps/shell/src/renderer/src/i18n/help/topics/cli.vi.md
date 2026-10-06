# Dòng lệnh và tác nhân

Mọi bản cài đặt đều kèm lệnh `genoffice` điều khiển đúng những engine mà cửa sổ dùng — cùng bộ phân tích, cùng bộ ghi, cùng bộ kết xuất. Một tệp do ứng dụng lưu và một tệp do lệnh ghi là cùng một tệp, và một kiểm tra mà bảng AI của ứng dụng vượt qua thì lệnh cũng vượt qua.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

Đó là toàn bộ bề mặt trong một màn hình — mọi lệnh đều có một dòng giải thích nó làm gì, rồi đến các tùy chọn toàn cục và mã thoát:

![Kết quả thực tế của genoffice --help: biểu ngữ phiên bản, từng lệnh kèm mô tả một dòng, cùng các tùy chọn toàn cục và mã thoát](img/cli.png)

## Lấy lệnh

macOS và Windows đặt nó bên trong gói ứng dụng. Muốn gọi theo tên, hãy chạy `genoffice install-cli` một lần: lệnh này tạo liên kết tượng trỏ tới tệp nhị phân đi kèm trong `/usr/local/bin`, hoặc trong `PATH` của người dùng trên Windows.

## Những lệnh đáng biết

| Lệnh              | Tác dụng                                                                                                                                                                                                                     |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Mở một tài liệu trong ứng dụng; khởi động ứng dụng nếu nó chưa chạy.                                                                                                                                                         |
| `convert`         | Chuyển đổi giữa các định dạng bằng chính engine của ứng dụng.                                                                                                                                                                |
| `create`          | Tạo tài liệu từ nội dung có cấu trúc.                                                                                                                                                                                        |
| `render`          | Một PNG cho mỗi trang, đúng bố cục do bộ kết xuất dàn ra.                                                                                                                                                                    |
| `pdf`             | Đọc lớp văn bản của PDF từng trang một, không cần tiến trình ứng dụng.                                                                                                                                                       |
| `info`            | Siêu dữ liệu và phần tóm tắt cấu trúc của tài liệu.                                                                                                                                                                          |
| `search`          | Tìm kiếm web hoặc hình ảnh qua nhà cung cấp đã cấu hình trong ứng dụng.                                                                                                                                                      |
| `image` / `media` | Tạo hình ảnh, hoặc mô tả và đặt câu hỏi về một tệp hình ảnh, video hay âm thanh.                                                                                                                                             |
| `merge`           | Điền các vùng chờ `{{key}}` trong một mẫu `.docx`, `.pptx` hoặc `.xlsx`.                                                                                                                                                     |
| `capabilities`    | Báo cáo những tính năng đám mây nào đã được cấu hình trên máy này.                                                                                                                                                           |
| `guide`           | Tra cứu thao tác và các hướng dẫn thiết kế, sinh ra từ đúng những định nghĩa mà bộ thực thi dùng để kiểm tra — nên nó không thể lệch khỏi những gì `apply` chấp nhận. `--json` trả về tra cứu kèm lược đồ của từng thao tác. |
| `install-cli`     | Đưa `genoffice` vào `PATH`.                                                                                                                                                                                                  |
| `skill`           | Liệt kê các tác nhân lập trình tìm thấy trên máy này rồi cài hoặc cập nhật kỹ năng GenOffice trong chúng.                                                                                                                    |
| `mcp`             | Phục vụ mọi lệnh như một công cụ của Model Context Protocol. Xem **Kết nối tác nhân lập trình**.                                                                                                                             |

## Soạn thảo: tài liệu, bảng tính, bản trình chiếu

`genoffice docs`, `genoffice sheet` và `genoffice slides` đọc và sửa qua đúng lối ghi mà ứng dụng dùng, và chúng chia sẻ một từ vựng: một **thao tác** (op) là một lần sửa, còn một **đặc tả** (spec) là danh sách các thao tác áp dụng theo thứ tự.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` cho biết một lô sẽ làm gì và không ghi gì cả, đây là cách rẻ nhất để kiểm tra một đặc tả trước khi nó áp dụng. Bảng AI trong ứng dụng chạy đúng trên những thao tác này, nên bất cứ điều gì bạn nhờ nó làm thì bạn cũng viết kịch bản được.

## Model Context Protocol

`genoffice mcp` phục vụ mọi lệnh như một công cụ MCP, và `genoffice mcp install <agent|all>` đăng ký nó vào cấu hình riêng của một tác nhân lập trình. Xem **Kết nối tác nhân lập trình** cho phần bên kia.

## Điều lệnh không làm

Nó đọc tệp và ghi tệp. Nó không phải ứng dụng: không có cửa sổ, và hộp thoại cập nhật trong ứng dụng không áp dụng. Bất cứ thứ gì cần cửa sổ — bảng AI, lượt kiểm tra chất lượng trên một trang chiếu đã kết xuất — đều phải chờ bạn mở tệp ra. `genoffice render` đưa cho bạn các điểm ảnh mà không cần cửa sổ, và `genoffice slides` tự kiểm tra bố cục của một bộ trình chiếu.
