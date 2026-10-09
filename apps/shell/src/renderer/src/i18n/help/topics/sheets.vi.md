# Sheets: bảng tính

Sheets là trình soạn thảo giống Excel; việc tính toán chạy trong một tiến trình công cụ Rust riêng (sự cố ở đó không bao giờ làm sập ứng dụng). Nó mở và lưu .xlsx thật; .csv và .tsv mở ra dưới dạng bảng.

## Giao diện

- **Dải ruy-bâng**: tám thẻ, được giải thích từng thẻ bên dưới.
- **Thanh công thức**: hiển thị và chỉnh sửa công thức của ô đang hoạt động; có hỗ trợ các hàm thông dụng.
- **Thẻ trang tính** (dưới cùng): thêm / đổi tên / xóa / di chuyển trang tính.
- **Sửa ô**: nhấp đúp hoặc chỉ cần gõ; Enter xác nhận và xuống dưới, Tab sang phải, Escape hủy (thói quen của Excel).
- **Phím tắt**: theo họ Excel (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Các thẻ trên dải ruy-bâng

- **Trang đầu**: phông chữ, tô màu, viền, định dạng số (tiền tệ/phần trăm/hàng nghìn, tăng giảm số thập phân), căn chỉnh, gộp, chèn hàng/cột và đổi kích thước, định dạng có điều kiện, định dạng dưới dạng bảng, kiểu ô, bảng tạm và công cụ định dạng, sắp xếp & lọc.
- **Chèn**: hình ảnh, hình dạng, hộp văn bản, liên kết, nhận xét, hộp kiểm tra, đầu trang & chân trang, ký hiệu, phương trình; **chín loại biểu đồ** (cột, thanh, đường, vùng, tròn, phân tán, radar, vành khuyên, và kết hợp cột + đường) cùng **Biểu đồ đề xuất**, vốn xếp hạng các loại theo vùng đang chọn và cho xem trước; **PivotChart** tạo từ ô bạn đang đứng; **Biểu đồ mini (Sparklines)** (đường, cột, thắng/thua); **Slicer** và **Timeline** để lọc một PivotTable.
  - Biểu đồ, biểu đồ pivot và biểu đồ mini là các đối tượng thật trong sổ làm việc đã lưu.
  - **Slicer** và **Timeline** là điều khiển theo phiên: thao tác lọc của chúng được lưu lại và Excel hiển thị đúng PivotTable đã lọc đó, nhưng bản thân nút slicer không nằm trong tệp.
- **Bố trí Trang**: màu và phông chủ đề, công tắc in đường lưới/tiêu đề, xem trước ngắt trang.
- **Công thức**: Tính tổng tự động và chèn hàm, đặt tên (kể cả từ vùng chọn), dò ô tiền nhiệm/ô phụ thuộc, cửa sổ theo dõi, tính toán lại trang tính/sổ làm việc.
- **Dữ liệu**: sắp xếp & lọc (gồm bộ lọc nâng cao, xóa bộ lọc), chuyển văn bản thành cột, hợp nhất sổ làm việc, làm mới tất cả.
- **Xem lại**: duyệt nhận xét (hiện, trước/sau), dịch và nhóm **Bảo vệ** — **Bảo vệ trang tính**, **Bảo vệ sổ làm việc** và **Cho phép chỉnh sửa các vùng**.
  - Bảo vệ được ghi vào tệp .xlsx và những gì ứng dụng này áp dụng không có mật khẩu, nên cùng nút đó thành **Bỏ bảo vệ…** và hoàn tác lại. Bảo vệ đến từ chương trình khác _thực sự_ có mật khẩu thì không thể gỡ khỏi đây.
  - Hai nút bảo vệ có hiệu lực ngay — không có hộp thoại nào để hủy, chỉ có ghi chú trên thanh trạng thái rằng nó sẽ được ghi khi lưu.
  - **Cho phép chỉnh sửa các vùng** đánh dấu những ô vẫn chỉnh sửa được trong khi phần còn lại của trang tính bị khóa.
- **Xem**: công tắc đường lưới, **thanh công thức**, tiêu đề và làm nổi bật hàng & cột hiện hoạt; thu phóng; **Bình thường** và **Xem trước ngắt trang**. Đường lưới và tiêu đề được lưu cùng trang tính; làm nổi bật là tùy chọn riêng của bạn.
- **Thiết kế Biểu đồ**: xuất hiện khi đang chọn một biểu đồ — loại biểu đồ, kiểu và màu, chỉnh sửa vùng dữ liệu.

Thẻ "Dữ liệu", từng nút (theo ảnh, từ trái sang phải):

![Thẻ Dữ liệu](img/sheets-data.png)

- **PivotTable**: dựng PivotTable từ vùng hiện tại; kéo các trường để tổng hợp.
- **Làm mới**: tính toán lại dữ liệu của PivotTable hiện tại.
- **Từ Văn bản/CSV**: nhập một tập tin .csv/.txt thành trang tính mới, tách theo dấu phân cách.
- **Hợp nhất sổ làm việc**: lấy các trang tính từ tập tin .xlsx khác vào tập tin này.
- **Làm mới tất cả**: tính toán lại mọi PivotTable và tập dữ liệu bên ngoài.
- **Sắp xếp** (thả xuống): tăng dần / giảm dần / sắp xếp tùy chỉnh (quy tắc nhiều cột).
- **Bộ lọc**: thêm mũi tên ▼ vào hàng tiêu đề; đánh dấu các giá trị muốn giữ.
- Các nút nhỏ xếp chồng bên cạnh: **Xóa** (đưa mọi hàng trở lại), **Áp dụng lại** (chạy lại bộ lọc hiện tại), **Nâng cao** (lọc bằng một vùng tiêu chí).
- **Chuyển văn bản thành cột** (thả xuống): tách một cột thành nhiều cột theo dấu phân cách hoặc theo độ rộng cố định.
- **Flash Fill**: đưa ra một ví dụ, phần còn lại của cột sẽ điền theo ví dụ đó (ctrl+E).
- **Xóa các bản sao trùng lặp**: bỏ các hàng trùng theo những cột đang chọn.
- **Xác thực dữ liệu** (thả xuống): quy tắc nhập cho vùng chọn (danh sách thả xuống, khoảng số...).
- **Hợp nhất dữ liệu**: gom nhiều vùng vào một chỗ theo danh mục.
- **Phân tích What-If** (thả xuống): Tìm kiếm mục tiêu (Goal Seek) — giải một ô đầu vào để ô công thức đạt giá trị mục tiêu.
- **Nhóm hàng / Bỏ nhóm hàng** (thả xuống): nhóm hàng hoặc cột, kèm thu gọn và mở rộng.
- **Tổng phụ**: chèn các hàng tổng phụ theo danh mục.

Thẻ "Công thức", từng nút:

![Thẻ Công thức](img/sheets-formulas.png)

- **Chèn hàm** (fx): tìm hàm kèm trình hướng dẫn tham số.
- **Tính tổng tự động (AutoSum)** (thả xuống): SUM một cú nhấp, cùng trung bình/đếm/lớn nhất/nhỏ nhất.
- **Dùng gần đây / Tài chính / Logic / Văn bản / Ngày & Giờ / Tra cứu & Tham chiếu / Toán & Lượng giác / Thêm hàm khác**: duyệt và chèn hàm theo danh mục.
- **Trình quản lý tên**: xem, tạo và xóa các vùng đã đặt tên.
- **Đặt tên** (thả xuống): đặt tên cho vùng chọn; **Sử dụng trong công thức** chèn một tên đã có; **Tạo từ vùng chọn** tạo tên từ hàng tiêu đề/cột tiêu đề của vùng đó.
- **Dò ô tiền nhiệm / Dò ô phụ thuộc**: các mũi tên xanh cho biết dữ liệu của công thức đến từ đâu và chảy đi đâu; **Xóa mũi tên** xoá chúng.
- **Hiện công thức**: các ô hiển thị chính công thức thay vì kết quả.
- **Kiểm tra lỗi**: tìm và giải thích lỗi trong công thức.
- **Cửa sổ theo dõi**: ghim những ô bạn quan tâm và theo dõi giá trị trực tiếp của chúng.
- **Tùy chọn tính toán** (thả xuống): tính toán lại tự động hay thủ công; ở chế độ thủ công, **Tính toán ngay / Tính toán trang tính** kích hoạt thủ công.

## Số và định dạng

- Định dạng số: chung, số, tiền tệ, phần trăm, ngày/giờ, phân số, khoa học và hơn thế.
- Căn chỉnh, xuống dòng, ô đã gộp, viền và màu tô.
- Kéo để đổi chiều cao hàng và độ rộng cột; nhấp đúp vào đường viền để tự vừa.

## Dữ liệu

**Sắp xếp & lọc** (ví dụ, giảm dần theo một cột):

1. Nhấp vào **ô bất kỳ trong cột đó** (không cần chọn cả cột).
2. Thẻ Trang đầu ▸ **Sắp xếp & Lọc** ▸ **Z → A**; cả hàng sẽ sắp xếp lại cùng nhau (vùng đó được coi là một khối).
3. Muốn dùng quy tắc riêng (nhiều cột, theo màu): đi cùng đường đó, chọn **Sắp xếp tùy chỉnh**.
4. Bộ lọc: chọn hàng tiêu đề rồi nhấp **Sắp xếp & Lọc ▸ Bộ lọc** — mỗi tiêu đề có mũi tên ▼ để đánh dấu giá trị muốn giữ; xóa bộ lọc để đưa tất cả trở lại.

- Sắp xếp và lọc.
- Cố định ngăn.
- .csv / .tsv: mở thẳng dưới dạng bảng (tsv phân tách bằng tab được phân tích thành một bảng); khi lưu sẽ ghi lại đúng định dạng gốc.

## Menu chuột phải

- **Trong lưới**: menu sẵn có của chính trình soạn thảo (Univer) — cắt/sao chép/dán, chèn và xóa hàng/cột, ẩn, gộp ô, cố định ngăn và các mục thường ngày khác.
- **Trên dải trạng thái dưới cùng**: chọn số liệu thống kê nào hiển thị trên thanh trạng thái (trung bình / đếm / tổng, ...); lựa chọn được ghi nhớ.
- **Trên một thẻ trang tính ở dưới**: thêm / đổi tên / xóa / tô màu / ẩn trang tính (menu thẻ của Univer).
- Menu ngữ cảnh của dải thẻ phía trên được nói ở [Quản lý thẻ và cửa sổ](help://tabs-and-windows).

## AI

- Bảng AI bên cạnh: chọn một vùng rồi ra lệnh bằng ngôn ngữ tự nhiên (đổi định dạng, tạo dữ liệu, viết công thức).
- Gắn tệp vào câu lệnh bằng nút 📎, hoặc kéo tệp thả vào bảng; chúng đi cùng câu hỏi và ảnh trở về dưới dạng hình thu nhỏ.
- Một câu trả lời có thể trích dẫn một ô — bấm vào tham chiếu là lưới nhảy tới ô đó.
- Có thể quay lui các thay đổi của AI ngay trên bảng này.

## Lưu và xuất

- Lưu .xlsx (công thức và định dạng được giữ); Lưu dưới dạng; xuất PDF theo cách ngắt trang in.
- Tự động lưu tuân theo quy tắc chung (bật sau lần lưu thủ công đầu tiên).

## Độ ổn định

- Bộ công cụ tính toán Rust được tách khỏi giao diện ở mức tiến trình: nếu dữ liệu cực đại làm nó chết, bạn chỉ nhận được một thông báo và một lần thử khôi phục phiên làm việc — chứ không phải ứng dụng sập.
