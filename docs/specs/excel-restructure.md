# Đặc tả tái cấu trúc workbook và xuất Word

## 1. Mục tiêu

Workbook phục vụ nhiều loại hồ sơ: thừa kế, phân chia/chuyển nhượng hai bên, ủy quyền và biểu mẫu khác. Người dùng luôn nhập theo **một danh sách người thống nhất**; không còn ghép dữ liệu thủ công thành một hàng ngang dài để xuất Word.

Nguyên tắc chính:

- Một người trong một hồ sơ là một dòng dữ liệu.
- Thêm dòng không làm đẩy, lệch hoặc làm hỏng dữ liệu bên còn lại.
- File Word xuất ra là bản sao độc lập của template, không dùng Mail Merge hoặc liên kết Word--Excel.
- Các bảng xuất trung gian chỉ do VBA tạo/refresh; người dùng không nhập vào đó.

## 2. Sheet và bảng dữ liệu

### `Nguoi` -- `tblNguoi`: giao diện nhập chung

Đây là màn hình duy nhất người dùng cần nhập người cho mọi loại hồ sơ. Mỗi dòng là một người **trong ngữ cảnh một hồ sơ**; không cần cố gắng gộp cùng một cá nhân giữa các hồ sơ khác nhau ở giai đoạn đầu.

| Cột | Cách tạo | Ý nghĩa |
| --- | --- | --- |
| `HoSoID` | Chọn/tự điền | Hồ sơ hiện hành |
| `NguoiID` | Hệ thống | ID cố định, ví dụ `P0001` |
| `Ben` | Thao tác nhanh | Chỉ dùng cho hồ sơ có hai bên; bắt buộc trống và khóa ở hồ sơ thừa kế |
| `STTBen` | Hệ thống | Thứ tự của người trong riêng Bên A/B khi xuất |
| `VaiTro` | Danh sách chọn | Bên chuyển nhượng, bên nhận, người ủy quyền, người được ủy quyền, người thừa kế... |
| `HoTen`, `NgaySinh`, `NgayChet` | Nhập tay | Thông tin nhân thân |
| `SoGiayTo`, `NgayCap`, `DiaChi` | Nhập tay | Giấy tờ và địa chỉ |
| `GhiChu` | Nhập tay | Ngoại lệ |

Các trường thừa kế chỉ bật khi `LoaiHoSo = Thừa kế`: `MaNhanh`, `TrangThaiTN` và `QuanHe` (nhãn chữ cho câu văn Word). Chi tiết quy tắc mã nhánh nằm tại [đặc tả thừa kế](inheritance-workflow.md). Ở loại hồ sơ này, `Ben` bị khóa vì vai trò và nhánh thừa kế mới là dữ liệu phân loại; những hồ sơ khác để trống các cột nhánh.

### `TaiSan` -- `tblTaiSan`

Mỗi tài sản là một dòng, có `HoSoID` để gắn vào hồ sơ. Bảng không giới hạn số lượng tài sản. Các trường gồm loại tài sản, số phát hành, số vào sổ, số thửa, số tờ bản đồ, địa chỉ, diện tích, hình thức/mục đích sử dụng, thời hạn, nguồn gốc, ngày cấp, cơ quan cấp và ghi chú.

### `Phu`

Sheet này chứa dữ liệu dùng chung và cấu hình:

- `tblHoSo(HoSoID, LoaiHoSo, TenHoSo, CoChiaHaiBen, ...)`.
- `tblNguoiUyQuyen` cho thông tin ủy quyền chuyên biệt.
- `tblQuyTacGiayTo` cho loại giấy tờ, nơi cấp và nhãn địa chỉ suy ra theo ngày cấp.
- `tblTruongTuDo(Placeholder, GiaTri)` cho trường riêng của từng template.

## 3. UX cột `Bên` cho hồ sơ hai bên

Mô hình này **ổn và nên chọn làm giao diện chuẩn**. Người dùng chỉ nhập liên tục một danh sách người, rồi gán mỗi người vào Bên A hoặc B.

Không tạo một control/ActiveX “nút trượt” riêng cho từng dòng: các control đó dễ sai vị trí khi lọc, sắp xếp, chèn dòng hoặc mở trên máy khác. Thay vào đó, dùng chính ô trong cột `Ben` như một nút trượt trực quan:

- Dòng mới của hồ sơ hai bên mặc định là `A`.
- Double-click ô `Ben` sẽ đổi `A` thành `B`; double-click lần nữa đổi lại `A`.
- Ô `A` hiển thị nền xanh, ô `B` hiển thị nền cam; biểu tượng trong ô có thể là `A ◀` và `▶ B`.
- Với hồ sơ không chia hai bên, `Ben` để trống và bị khóa khỏi thao tác toggle.
- Vẫn có danh sách chọn `A/B/trống` làm phương án dự phòng khi macro bị tắt.

Double-click an toàn hơn click đơn vì click một lần còn cần để chọn ô, sửa dữ liệu, dùng phím mũi tên hoặc lọc bảng. Nếu cần thao tác một click thật sự, có thể bổ sung sau bằng vùng nút ở cạnh bảng, nhưng không đặt một đối tượng trên từng hàng.

Ví dụ nhập liệu:

| Họ tên | Bên | Vai trò |
| --- | --- | --- |
| Nguyễn Văn An | A | Bên chuyển nhượng |
| Trần Thị Bình | A | Đồng sở hữu |
| Lê Văn Cường | B | Bên nhận chuyển nhượng |
| Phạm Thị Dung | B | Đồng sở hữu |

Khi thêm người thứ 3 của Bên A hoặc thứ 10 của Bên B, chỉ thêm một dòng mới vào `tblNguoi`. Không có bảng Bên A/B hiển thị cạnh nhau nên không có rủi ro đẩy hàng.

## 4. Bảng xuất ẩn theo bên

Hai bảng ẩn là hợp lý **nếu template cũ cần các vị trí cố định**, nhưng chúng không phải nguồn dữ liệu.

Sheet kỹ thuật `XuatAn` được đặt `VeryHidden`, gồm:

- `tblXuatBenA`: các dòng `tblNguoi` có `Ben = A`.
- `tblXuatBenB`: các dòng `tblNguoi` có `Ben = B`.

Khi bấm xuất, VBA:

1. Đọc `tblNguoi` theo `HoSoID` đang chọn.
2. Tách dữ liệu theo cột `Ben`.
3. Đánh lại `STTBen` liên tục từ 1 cho mỗi bên.
4. Refresh `tblXuatBenA` và `tblXuatBenB`.
5. Tạo `ExportMap` để thay placeholder trong template Word.

Về kỹ thuật, VBA có thể tạo `ExportMap` trực tiếp mà không cần sheet ẩn. Giữ `XuatAn` giúp kiểm tra, tương thích macro cũ và dễ truy lỗi; người dùng bình thường không thấy hoặc không sửa sheet này.

## 5. Thừa kế là biến thể chuyên biệt

Hồ sơ này dùng cùng `tblNguoi`, thêm ba cột hệ thống: `MaNhanh` (khóa cấu trúc cây, ví dụ `1`, `2`, `1.4.4`), `TrangThaiTN` (Hưởng, Từ chối, Đã chết, Không hưởng) và `QuanHe` (nhãn chữ cho câu văn Word). Số cuối của mã cố định vai trò: `.1` bố đẻ, `.2` mẹ đẻ, `.3` vợ/chồng, từ `.4` là con; cụm nút vai trò trên sheet nhập sinh mã theo slot này. Không cần bảng quan hệ thừa kế riêng -- nhóm người trong văn bản được truy vấn trực tiếp bằng tiền tố mã kèm lọc trạng thái.

Hồ sơ có thể có một hoặc nhiều chủ sở hữu; trường hợp phổ biến là một chủ hoặc hai vợ chồng. Người dùng luôn nhập chủ đất chính đầu tiên. Hệ thống lưu ID của người đó làm `ChuSoHuuChinhID`, thay vì dựa vào vị trí dòng hiện tại. Một người có thể đồng thời là chủ/đồng chủ và là người nhận di sản, nên hai quan hệ này được lưu ở bảng liên kết riêng, không ghi đè `VaiTro` trong `tblNguoi`.

Chi tiết mô hình, luồng nhập và quy tắc xuất nằm tại [đặc tả thừa kế](inheritance-workflow.md).

## 6. Xuất Word

VBA lấy dữ liệu từ `tblNguoi`, `tblTaiSan`, `Phu` và, với hồ sơ hai bên, hai bảng xuất ẩn để tạo `ExportMap` trong bộ nhớ.

Ví dụ placeholder:

- Cũ: `[Tên bên A 1]`, `[Tên bên B 1]`.
- Mới: `{{ben.a.1.ho_ten}}`, `{{ben.b.1.ho_ten}}`.
- Thừa kế: `{{nguoi.1.1.ho_ten}}`.
- Tài sản: `{{tai_san.1.so_phat_hanh}}`.

Macro tạo một Word instance riêng, thay placeholder trong phần thân, header, footer và textbox, lưu bản mới vào `output/`, rồi chỉ đóng instance do macro tạo. Các file Word người dùng đang mở không bị đóng hoặc hỏi lưu.

## 7. Kiểm tra trước khi xuất

- Hồ sơ có chia hai bên phải có ít nhất một người ở mỗi bên.
- `Ben` chỉ nhận `A`, `B` hoặc trống; người thuộc hồ sơ thừa kế phải để trống. Sự kiện toggle bị hủy trong hồ sơ thừa kế.
- Không trùng `NguoiID`; `STTBen` được VBA tạo lại, không cho nhập tay.
- Với thừa kế: không tạo vòng nhánh, không trùng mã nhánh.
- Cảnh báo placeholder không được thay và dữ liệu bắt buộc bị thiếu.
- Ngày trống xuất chuỗi rỗng, không xuất `0` hoặc `1900`.

## 8. Phạm vi MVP

1. Một `tblNguoi` chung có cột `Ben`.
2. Toggle A/B bằng double-click và màu sắc trong ô; có dropdown dự phòng.
3. `XuatAn` rất ẩn với hai bảng export được VBA refresh khi xuất.
4. Ánh xạ placeholder theo Bên A/B, thừa kế và tài sản.
5. Giữ xuất Word độc lập trong instance Word riêng.
