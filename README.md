# 1. Phân tích các tham số của hàm VLOOKUP

Hàm VLOOKUP được sử dụng để tra cứu thông tin từ Sheet Danh_Muc và tự động điền vào Sheet Don_Hang dựa trên mã linh kiện. Phương pháp này giúp giảm thời gian nhập liệu thủ công, hạn chế sai sót và bảo đảm tính thống nhất của dữ liệu.

Cú pháp hàm:

`=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

Trong bài toán này, hàm VLOOKUP có 4 tham số:

**Tham số 1 – Lookup_value (Giá trị cần tìm):** Là mã linh kiện cần tra cứu trong Sheet Don_Hang, chẳng hạn LK01 hoặc LK02. Ví dụ, nếu mã linh kiện nằm tại ô A2 thì giá trị cần tìm là A2.

**Tham số 2 – Table_array (Bảng tra cứu):** Là vùng dữ liệu chứa mã linh kiện, tên linh kiện và đơn giá trong Sheet Danh_Muc. Vùng này cần được cố định bằng dấu `$` để không thay đổi khi sao chép công thức xuống các hàng bên dưới. Ví dụ, nếu bảng danh mục nằm trong vùng A2:C101 thì sử dụng `$A$2:$C$101`.

**Tham số 3 – Col_index_num (Số thứ tự cột trả về):** Là vị trí của cột chứa thông tin cần lấy trong vùng bảng tra cứu. Nếu cột A chứa mã linh kiện, cột B chứa tên linh kiện và cột C chứa đơn giá thì số thứ tự cột trả về tên linh kiện là 2, còn đơn giá là 3.

**Tham số 4 – Range_lookup (Chế độ tìm kiếm):** Sử dụng `FALSE` để yêu cầu tìm kiếm chính xác mã linh kiện. Điều này giúp tránh trường hợp Excel trả về thông tin của một mã linh kiện khác khi không tìm thấy kết quả khớp chính xác.

# 2. Triển khai công thức VLOOKUP

## 2.1. Tra cứu tên linh kiện

Giả sử Sheet Don_Hang có mã linh kiện tại ô A2. Trong Sheet Danh_Muc, cột A chứa mã linh kiện, cột B chứa tên linh kiện và cột C chứa đơn giá, dữ liệu bắt đầu từ hàng 2 đến hàng 101.

Tại ô B2 của Sheet Don_Hang, nhập công thức:

`=IFERROR(VLOOKUP(TRIM(A2),Danh_Muc!$A$2:$C$101,2,FALSE),"Không tìm thấy mã linh kiện")`

Công thức sử dụng TRIM() để loại bỏ khoảng trắng thông thường ở đầu và cuối mã linh kiện cần tìm. IFERROR() giúp hiển thị thông báo dễ hiểu thay vì lỗi #N/A khi không tìm thấy kết quả.

## 2.2. Tra cứu đơn giá

Tại ô C2 của Sheet Don_Hang, nhập công thức:

`=IFERROR(VLOOKUP(TRIM(A2),Danh_Muc!$A$2:$C$101,3,FALSE),"Không tìm thấy mã linh kiện")`

Tham số thứ ba được đổi từ 2 thành 3 để lấy dữ liệu từ cột Đơn giá trong bảng danh mục.

Sau khi nhập công thức, kéo công thức xuống các hàng còn lại để tra cứu tự động toàn bộ 100 mã linh kiện. Dấu `$` giúp cố định vùng tra cứu, trong khi tham chiếu A2 tự động thay đổi thành A3, A4 và các hàng tiếp theo.

# 3. Xử lý bẫy dữ liệu và lỗi #N/A

**Bẫy 1: Quên cố định vùng tra cứu bằng dấu `$`.**

Nếu sử dụng vùng tham chiếu tương đối như `Danh_Muc!A2:C101`, vùng dữ liệu có thể dịch chuyển khi kéo công thức xuống, dẫn đến kết quả sai hoặc lỗi #N/A. Cần sử dụng `Danh_Muc!$A$2:$C$101` để cố định toàn bộ vùng tra cứu.

**Bẫy 2: Mã linh kiện có khoảng trắng thừa.**

Hàm TRIM() trong công thức giúp loại bỏ khoảng trắng thông thường ở đầu và cuối mã linh kiện cần tìm. Tuy nhiên, nếu mã linh kiện trong Sheet Danh_Muc cũng chứa khoảng trắng thừa, chỉ sử dụng TRIM(A2) là chưa đủ. Cần làm sạch cột mã linh kiện trong Sheet Danh_Muc bằng TRIM() trước khi tra cứu.

Nếu dữ liệu chứa khoảng trắng đặc biệt như ký tự không ngắt dòng, cần xử lý bổ sung bằng SUBSTITUTE() hoặc phương pháp làm sạch dữ liệu phù hợp.

# 4. Kết luận

Việc sử dụng VLOOKUP kết hợp với IFERROR() giúp tự động tra cứu tên linh kiện và đơn giá từ Sheet Danh_Muc sang Sheet Don_Hang. Việc cố định vùng tham chiếu bằng dấu `$`, sử dụng chế độ tìm kiếm chính xác FALSE và làm sạch mã linh kiện giúp hạn chế lỗi, tăng độ chính xác và tiết kiệm đáng kể thời gian xử lý dữ liệu.
