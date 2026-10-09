# Màn hình Trang chủ: tập tin của bạn ở đâu

Trang chủ là trang khởi đầu của GenOffice: thanh điều hướng bên trái, danh sách tập tin và các thẻ tạo nhanh bên phải.

![Màn hình Trang chủ](img/home-screen.png)

## Điều hướng trên thanh bên

- **Gần đây**: các tập tin bạn đã mở gần đây. Mỗi hàng đều đóng dấu thời điểm — hôm nay, hôm qua hoặc ngày.
- **Đã gắn sao**: các tập tin bạn đã gắn sao. Rê chuột lên một hàng tập tin rồi nhấp vào ngôi sao để thêm hoặc bỏ.
- **Hướng dẫn sử dụng**: mở sổ tay này.
- **Genspark Projects**: sau khi đăng nhập tài khoản Genspark, mục này hiển thị các dự án bạn đã tạo bằng Genspark AI trên web; nhấp vào một dự án để tiếp tục chỉnh sửa trong trình duyệt. Có hỗ trợ tìm kiếm, sắp xếp theo thời gian, làm mới và tải thêm.
- **Thư mục**: các thư mục bạn thêm vào thanh bên bằng **Thêm thư mục…**, hoặc kéo vào đó từ trình quản lý tập tin. Mỗi thư mục trở thành một gốc mà bạn có thể mở, tạo thư mục con bên trong, đổi tên và gỡ bỏ; thư mục nào mất kết nối sẽ hiển thị là không khả dụng và có thể gỡ khỏi danh sách. **Thư mục mới** tạo thêm một thư mục nữa.

Ở đây không có mục Thùng rác. Các tập tin đã xóa đi vào thùng rác của hệ thống, và việc khôi phục là của hệ điều hành.

## Danh sách tập tin

Mỗi hàng hiển thị biểu tượng, tên tập tin, thời gian sửa đổi và các thông tin khác. **Menu ⋯** của hàng cung cấp:

- **Mở**, và **Hiển thị trong thư mục** để xác định vị trí tập tin trong trình quản lý tập tin của bạn.
- **Sao chép đường dẫn**.
- **Di chuyển vào thư mục…**: mở hộp chọn thư mục và thực sự chuyển tập tin đi; nếu nơi đến đã có tập tin trùng tên, bạn có thể bỏ qua, ghi đè hoặc đổi tên.
- **Đổi tên**: tại chỗ, phần mở rộng được giữ tự động.
- **Gắn sao / Bỏ gắn sao** — dấu sao vẫn còn sau khi khởi động lại và đi theo tập tin khi bạn đổi tên nó.
- **Tạo bản sao**: tạo một bản sao trong cùng thư mục.
- **Xóa**: chuyển tập tin vào thùng rác hệ thống — không xóa vĩnh viễn.
- **Xóa khỏi danh sách**, trong mục **Gần đây** cấp cao nhất, để gỡ một mục mà không đụng tới tập tin.

### Nhiều tập tin cùng lúc

Chọn ô tick của một hàng, hoặc nhấp ⌘/ctrl, để tạo một vùng chọn; ô tick ở tiêu đề chọn tất cả những gì đang được liệt kê, và một thanh phía trên danh sách cho biết đã chọn bao nhiêu (**Đã chọn {n}**) kèm **Di chuyển vào thư mục…** và **Xóa tệp** cho toàn bộ nhóm. Bạn cũng có thể kéo một vùng chọn nhiều vào một thư mục trên thanh bên.

## Tìm kiếm

Ô tìm kiếm ở trên khớp với hai thứ cùng lúc:

- **Tên tập tin**: lọc nhanh theo tên.
- **Nội dung tập tin**: GenOffice lập chỉ mục cho tập tin của bạn ở nền (văn bản bên trong docx/xlsx/pptx/pdf/md/html), nên tìm trong phần chữ cũng ra tập tin. Phạm vi và các công tắc nằm trong cài đặt tìm kiếm.

## Thẻ bắt đầu nhanh

Các thẻ phía trên danh sách tạo tài liệu mới trong một bước. Nhấp vào một thẻ sẽ tạo tập tin thuộc loại đó và mở trình soạn thảo tương ứng — bạn có thể viết ngay, hoặc để AI soạn thảo giúp (mỗi trình soạn thảo đều có **nút AI** trên dải ruy-băng, và **Hỏi AI** trong menu ngữ cảnh của vùng đã chọn).

Tập tin mới nằm trong thư mục đang được chọn ở thanh bên; nếu không chọn gì thì nó vào thư mục mặc định.

Từng thẻ dùng để làm gì:

- **AI Docs** (.docx): tài liệu văn bản trống trong trình soạn thảo Docs. Tập tin chỉ thực sự được ghi xuống đĩa ở **lần lưu đầu tiên**; tài liệu mới mở với bảng AI đang mở rộng (tắt trong Cài đặt → "Mở bảng AI trong tài liệu mới").
- **AI Sheets** (.xlsx): bảng tính trống trong trình soạn thảo Sheets. Trước khi bạn lưu, trên đĩa chưa có tập tin nào — tên được dành sẵn cho lần lưu đầu tiên; sau lần AI sinh nội dung đầu tiên, tên tập tin còn có thể được đổi tự động theo nội dung.
- **AI Slides** (.pptx): bản trình bày trống trong trình soạn thảo Slides.
- **AI Markdown** (.md): tài liệu Markdown trống trong trình soạn thảo Markdown.
- **AI HTML** (.html): trang web trống trong trình soạn thảo HTML.
- **AI PDF** (.pdf): khác với những thẻ còn lại — nó **ngay lập tức** tạo một tập tin PDF trắng thật một trang trong thư mục đích và mở nó như một tập tin bình thường (trình soạn thảo PDF làm việc trên tập tin thật). Phù hợp để thêm chú thích, che nội dung hoặc thêm văn bản; tên tập tin có thể được đổi tự động theo nội dung ở lần lưu đầu tiên.
- **Mở tệp cục bộ**: bộ chọn tập tin của hệ thống cho Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) và trang web (.html/.htm). Có thể chọn nhiều; mỗi tập tin mở một thẻ riêng.

> Gợi ý: File ▸ Mới trên dải menu tạo ra cùng những loại tài liệu đó (⌘N/Ctrl+N mặc định tạo tài liệu văn bản); kéo một tập tin vào cửa sổ nghĩa là mở nó.

## Dự án trên đám mây (Genspark Projects)

- Lần dùng đầu tiên cần đăng nhập tài khoản Genspark (luồng mã thiết bị: GenOffice hiển thị một mã, bạn hoàn tất đăng nhập trong trình duyệt).
- Danh sách dự án đồng bộ với web; Mở trong trình duyệt đưa bạn tới đó để tiếp tục.
- Không đăng nhập không ảnh hưởng tới bất kỳ tính năng cục bộ nào.
