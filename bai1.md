# Bài tập 1

## Phần 1:
### Phân tích: Xác định Record và Filed trong bảng Excel giao dịch.

* **Field (Trường / Cột):** Là các thuộc tính đơn lẻ mô tả đặc điểm của giao dịch. Trong bảng đối soát, các Field tiêu biểu bao gồm: 
    * `Mã giao dịch` (Transaction ID)
    * `Mã người dùng` (User ID)
    * `Tên người dùng` (User Name - được tra cứu qua hàm VLOOKUP)
    * `Số tiền giao dịch` (Transaction Amount)
    * `Phí giao dịch` (Transaction Fee)
    * `Trạng thái giao dịch` (Status: thành công/ thất bại/ đang xử lý)
    * `Thời gian giao dịch` (timestamp)
* **Record (Bản ghi / Hàng):** là một tập hợp các giá trị của các Field liên kết với nhau, đại diện cho một giao dịch hoàn chỉnh cụ thể được thực hiện trên hệ thống ví điện tử PayKK tại một thời điểm nhất định.



## Phần 2:
### 1. Tài liệu Word: Cấu trúc SRS dài 5 trang
*   **Thiết kế cấu trúc:** Tài liệu được phân cấp rõ ràng bằng Heading Styles (`Heading 1`, `Heading 2`) nhằm tự động hóa mục lục (`Auto TOC`).
*   **Gợi ý phân chia nội dung theo 5 trang:**
    *   **Trang 1:** Trang bìa (Tên dự án "PayKX", Tên nhóm/Trưởng dự án, Ngày tháng) và Mục lục tự động (`Auto TOC`).
    *   **Trang 2:** 
        *   `Heading 1: 1. Giới thiệu chung`
            *   `Heading 2: 1.1 Mục đích dự án ví điện tử PayKX`
            *   `Heading 2: 1.2 Phạm vi tài liệu`
    *   **Trang 3:** 
        *   `Heading 1: 2. Mô tả tính năng hoàn tiền (Cashback)`
            *   `Heading 2: 2.1 Quy chế tích điểm và hoàn tiền`
            *   `Heading 2: 2.2 Điều kiện áp dụng cho sinh viên`
    *   **Trang 4:** 
        *   `Heading 1: 3. Yêu cầu chức năng hệ thống`
            *   `Heading 2: 3.1 Giao diện đối soát giao dịch`
            *   `Heading 2: 3.2 Quản lý dữ liệu người dùng`
    *   **Trang 5:** 
        *   `Heading 1: 4. Phụ lục và Kiểm thử`
            *   `Heading 2: 4.1 Xử lý lỗi hệ thống`
            *   `Heading 2: 4.2 Kế hoạch bàn giao`



### 2. Bảng tính Excel: Xử lý dữ liệu và Tra cứu VLOOKUP
*   **Làm sạch dữ liệu:** Trước khi dùng hàm tra cứu, áp dụng hàm `TRIM` để loại bỏ các khoảng trắng thừa (ví dụ: `=TRIM(A2)`) nhằm tránh lỗi không khớp dữ liệu do khoảng trắng vô hình.
*   **Công thức VLOOKUP tra cứu tên người dùng:**
    *   Giả sử bảng dữ liệu giao dịch nằm ở cột A-B và bảng danh mục người dùng (mã và tên) nằm ở vùng `D$2:E$100`.
    *   Công thức tại ô chứa tên người dùng:
        ```excel
        =VLOOKUP(TRIM(A2), D$2:E$100, 2, FALSE)
        ```
    *   *Lưu ý quan trọng theo ràng buộc:* Tuyệt đối **không Merge Cell** trong vùng dữ liệu và **không để hàng trống** giữa các bản ghi để đảm bảo hàm tra cứu hoạt động chính xác tuyệt đối.



### 3. Bài thuyết trình PowerPoint: Báo cáo doanh thu theo Rule 10/20/30
*   **Quy tắc Assertion Title:** Tiêu đề slide không dùng từ khóa chung chung (như "Báo cáo doanh thu") mà phải là một câu khẳng định đúc kết insight (Ví dụ: *"Doanh thu tăng 15% nhờ ưu đãi nạp điện thoại"*).
*   **Định dạng tuân thủ ràng buộc:**
    *   Sử dụng font chữ lớn, tối thiểu **30pt** cho nội dung chính.
    *   Trực quan hóa số liệu bằng **biểu đồ cột/đường** thay vì copy bảng số liệu thô nhiều chữ số gây rối mắt người xem.




## Phần 3. Kiểm thử: Xử lý lỗi #N/A trong VLOOKUP (Quy trình 4 điểm)

Khi bảng tra cứu trả về lỗi `#N/A` (nghĩa là không tìm thấy giá trị dò tìm trong bảng tham chiếu), hãy thực hiện tuần tự quy trình 4 điểm sau để tìm nguyên nhân và khắc phục:

1.  **Kiểm tra khoảng trắng thừa (Whitespace Error):**
    *   *Nguyên nhân:* Dữ liệu ô cần tìm hoặc bảng tra cứu chứa khoảng trắng ở đầu hoặc cuối (ví dụ: `"KH01 "` khác với `"KH01"`).
    *   *Cách khắc phục:* Đảm bảo đã bọc hàm `TRIM` ở giá trị dò tìm (`=VLOOKUP(TRIM(A2),...)`) hoặc sử dụng công cụ *Find & Replace* để xóa khoảng trắng thừa.
2.  **Kiểm tra định dạng dữ liệu (Data Type Mismatch):**
    *   *Nguyên nhân:* Một bên là định dạng kiểu Số (`Number`) nhưng bên kia lại ở dạng Văn bản (`Text`) (ví dụ mã giao dịch `1002` lưu thành chữ thay vì số).
    *   *Cách khắc phục:* Đồng bộ định dạng dữ liệu của cả cột giá trị dò tìm và cột đầu tiên của bảng tham chiếu về chung một định dạng (thường là General hoặc Text).
3.  **Kiểm tra phạm vi bảng tham chiếu (Table Array Reference):**
    *   *Nguyên nhân:* Vùng dữ liệu tham chiếu chưa được cố định tuyệt đối bằng ký tự đô la (`$`), dẫn đến việc khi kéo công thức xuống dưới, phạm vi vùng tìm kiếm bị lệch đi.
    *   *Cách khắc phục:* Cố định vùng tham chiếu (ví dụ đổi từ `D2:E100` thành `D$2:E$100`).
4.  **Kiểm tra tính chính xác của tham số cách thức dò tìm (Index / Exact Match):**
    *   *Nguyên nhân:* Quên đặt tham số cuối cùng là `0` hoặc `FALSE`, khiến hàm mặc định tìm kiếm tương đối (`TRUE`) dẫn đến lỗi kết quả hoặc trả về `#N/A` khi dữ liệu chưa được sắp xếp.
    *   *Cách khắc phục:* Đảm bảo tham số thứ 4 luôn là `FALSE` (hoặc `0`) để bắt buộc khớp đúng tuyệt đối.