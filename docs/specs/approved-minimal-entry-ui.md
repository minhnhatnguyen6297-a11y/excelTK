# Đặc tả đã phê duyệt: giao diện nhập liệu tối giản

> Trạng thái: đã chốt để làm cơ sở triển khai workbook/VBA.  
> Cập nhật: 26/08/2026.

## 1. Bối cảnh và mục tiêu

Workbook phục vụ nhập một hồ sơ và xuất các biểu mẫu Word. Giao diện nhập chính phải là nơi thao tác nhanh, giống một cơ sở dữ liệu nhỏ, thay vì cố gắng hiển thị toàn bộ quan hệ và quy tắc nghiệp vụ trên cùng một màn hình.

Mục tiêu của giao diện chính là nhập đúng dữ liệu gốc, dễ nhìn khi có nhiều người và nhiều trường tài sản. Các dữ liệu phái sinh, quan hệ hoặc cấu hình kỹ thuật không làm cản trở việc nhập liệu hàng ngày.

Đặc tả này cụ thể hóa phần giao diện của [đặc tả tái cấu trúc workbook](excel-restructure.md) và giữ nguyên mô hình dữ liệu thừa kế trong [đặc tả hồ sơ thừa kế](inheritance-workflow.md).

## 2. Bản demo đã phê duyệt

Bản tham chiếu giao diện là [ho-so-thua-ke-demo.html](../../prototypes/ho-so-thua-ke-demo.html).

Demo này là minh họa bố cục và thao tác; không phải dữ liệu thật, không ghi vào workbook, và không thay thế cho thiết kế VBA.

## 3. Những quyết định đã chốt

1. Sheet nhập chính chỉ có **hai vùng nhập liệu**:
   - `Thông tin người`.
   - `Thông tin tài sản`.
2. Không hiển thị ở sheet nhập chính: tiến độ hồ sơ, sơ đồ nhánh thừa kế, bảng phân chia, cảnh báo chi tiết, chú thích giải thích nghiệp vụ, hay các bảng kỹ thuật.
3. Các thông tin phái sinh được đặt ở sheet phụ. Khi thật sự cần sửa, người dùng có thể vào sheet phụ thay vì phải có thêm vùng điều khiển trên sheet nhập chính.
4. Mọi quy tắc tạo liên kết, kiểm tra và xuất Word do hàm/VBA xử lý. Giao diện chính chỉ tập trung vào dữ liệu cần nhập.
5. Thông tin tài sản được trình bày theo **phiếu dọc cho từng tài sản**, không phải một bảng ngang dài. Điều này cho phép bổ sung nhiều trường mà vẫn dễ đọc.
6. Mã kỹ thuật không hiển thị và không yêu cầu người dùng nhập trên sheet nhập chính.

## 4. Sheet nhập chính

### 4.1. Phần đầu sheet

Giữ ở mức tối giản:

- Mã/tên hồ sơ để nhận biết hồ sơ đang mở.
- Danh sách chọn `Loại hồ sơ` (ví dụ: Thừa kế, Chuyển nhượng, Ủy quyền).
- Lệnh `Lưu`.
- Lệnh `Xuất Word`.

Không đưa thanh tiến độ, tình trạng, hay ghi chú nghiệp vụ vào phần đầu sheet.

### 4.2. Vùng `Thông tin người`

Một dòng là một người trong phạm vi hồ sơ. Các cột hiển thị và cho phép nhập là:

| Thứ tự hiển thị | Trường nhập |
| --- | --- |
| 1 | Họ và tên |
| 2 | Ngày sinh |
| 3 | Ngày mất |
| 4 | Số giấy tờ |
| 5 | Ngày cấp |
| 6 | Địa chỉ |

Có lệnh thêm dòng/người. Người dùng nhập liên tục trong một danh sách; không tách danh sách theo vai trò, Bên A/B hoặc nhánh thừa kế.

Khi `Loại hồ sơ = Thừa kế`, danh sách hiển thị thêm cột chỉ-đọc `Mã nhánh` và cụm nút vai trò `[Thêm bố] [Thêm mẹ] [Thêm vợ/chồng] [Thêm con] [Con đã mất – mở nhánh] [Người khác]` áp dụng cho dòng đang chọn; nút sinh mã theo slot và tạo dòng người mới tự động (xem [đặc tả thừa kế](inheritance-workflow.md)).

### 4.3. Vùng `Thông tin tài sản`

Mỗi tài sản là một phiếu dọc độc lập, được đánh số theo thứ tự hiển thị (`Tài sản 1`, `Tài sản 2`, ...). Các trường MVP đang hiển thị là:

- Loại tài sản (danh sách chọn).
- Số phát hành / số GCN.
- Số thửa.
- Tờ bản đồ.
- Địa chỉ.
- Diện tích (m²).
- Nguồn gốc.

Có lệnh thêm tài sản. Khi cần thêm trường tài sản theo template hoặc biểu mẫu, bổ sung vào phiếu dọc; không chuyển vùng này thành bảng ngang nhiều cột.

## 5. Mã kỹ thuật và dữ liệu ẩn

`NguoiID` và `TaiSanID` là khóa kỹ thuật do VBA tự sinh, ví dụ `P0001` và `TS001`.

Mục đích của chúng là liên kết ổn định giữa các bảng, kể cả khi:

- Người dùng đổi họ tên hoặc hai người có cùng tên.
- Danh sách bị lọc, sắp xếp hoặc chèn thêm dòng.
- Một người hoặc tài sản được tham chiếu trong quan hệ thừa kế, sở hữu, phân chia hay ánh xạ xuất Word.

Các mã này được lưu trong bảng dữ liệu nhưng **ẩn khỏi giao diện nhập chính**. Người dùng không tự gõ, sửa hoặc cần nhớ mã. VBA phải tự sinh mã duy nhất khi thêm dòng mới.

Ngoại lệ: với hồ sơ thừa kế, `MaNhanh` mang ý nghĩa điều hướng nên hiển thị chỉ-đọc trên vùng Người; người dùng vẫn không tự gõ hay sửa.

`STT` và nhãn `Tài sản 1`, `Tài sản 2` chỉ phục vụ quan sát/thứ tự nhập; chúng không phải khóa liên kết và không được dùng thay cho ID kỹ thuật.

## 6. Sheet phụ và trách nhiệm của VBA

Sheet phụ lưu các dữ liệu không cần nhập thường xuyên, bao gồm tối thiểu:

- Thông tin hồ sơ và cấu hình loại hồ sơ.
- Khối di sản, chủ chính/đồng chủ và quan hệ sở hữu.
- Quan hệ, nhánh và trạng thái thừa kế.
- Phân chia di sản.
- Quy tắc giấy tờ, trường tự do, cấu hình template và dữ liệu xuất trung gian.

VBA/hàm chịu trách nhiệm:

- Sinh `NguoiID`, `TaiSanID` và các ID liên kết khác.
- Tạo hoặc cập nhật dữ liệu phái sinh từ dữ liệu nhập chính.
- Áp dụng quy tắc riêng theo `LoaiHoSo`; với hồ sơ thừa kế, cột `Ben` không hiện trên giao diện chính và luôn được khóa/trống.
- Kiểm tra tính hợp lệ trước khi xuất.
- Tạo `ExportMap`, sao chép template và xuất Word trong Word instance riêng theo các đặc tả hiện có.

## 7. Ngoài phạm vi giao diện chính

Các nội dung sau vẫn là dữ liệu hợp lệ của hệ thống, nhưng không nằm ở màn hình nhập chính: vai trò pháp lý, Bên A/B, chủ/đồng chủ, người để lại di sản, mã nhánh, tỷ lệ/phần được chia, cảnh báo kiểm tra và placeholder Word.

Chúng chỉ xuất hiện khi cần trong sheet phụ hoặc do VBA xử lý trong lúc lưu/xuất.
