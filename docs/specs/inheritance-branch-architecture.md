# Đặc tả hồ sơ thừa kế theo Hàng TK

> Trạng thái: **đã chốt và đã triển khai trong workbook MVP v0.2.1**.  
> Ngày cập nhật: 29/08/2026.  
> Phiên bản cấu trúc dữ liệu: `2.0.0`.  
> Phạm vi: nhập danh sách người, xác định nhánh, người nhận đất, người từ chối, xử lý ngày tháng và dữ liệu xuất Word.

## 1. Mục tiêu

Giữ cách nhập quen thuộc: mỗi người nằm trên một dòng và người dùng nhập liên tục, không phải dừng lại để chọn vai trò cha, mẹ, con hay vợ/chồng.

Mỗi hồ sơ cần trả lời bốn câu hỏi:

1. Ai là chủ đất ban đầu?
2. Mỗi người thuộc Hàng TK và nhánh nào?
3. Ai đã chết?
4. Ai nhận đất? Những người còn sống còn lại được xác định là từ chối.

Thông tin dùng để tính phải được bảo vệ khỏi sửa nhầm. Ngày tháng phải chịu được trường hợp chỉ biết năm, chỉ biết tháng/năm hoặc dữ liệu được dán từ nơi khác.

## 2. Quyết định kiến trúc đã chốt

- Dùng một danh sách người phẳng, mỗi người một dòng.
- Thuật ngữ trên giao diện và trong tài liệu là **Hàng TK**, không dùng tên gọi cũ.
- Cấu trúc bắt đầu từ `H0` là hàng chủ đất, sau đó là `H1` đến `H4`.
- `Hàng TK tối đa` mặc định là `2`, hợp lệ từ `0` đến `4`.
- Hai vị trí đầu dành cho chủ đất và cố định ở `H0`.
- Các dòng nhánh bắt đầu từ `H1`. Từ `H2` trở đi, `ParentNguoiID` cho biết dòng thuộc nhánh của ai.
- Mỗi Hàng TK có màu riêng, đủ tương phản để phân biệt bằng mắt; chữ `H0`...`H4` luôn được hiển thị để không phụ thuộc riêng vào màu.
- Cột Hàng TK và cột Nhận đất hoạt động như nút bấm, nhưng bản thân các ô vẫn bị khóa để người dùng không thể gõ đè dữ liệu hệ thống.
- Người nhận đất được chọn trực tiếp trên từng dòng.
- Người từ chối và trạng thái được tính tự động, không nhập tay và không hiển thị cột trạng thái trên giao diện.
- Dữ liệu dự phòng cho nhóm từ chối được tạo ngay từ đầu nhưng chưa có giao diện riêng.

Đây là nguồn quyết định chính để triển khai. Không quay lại mô hình nhập vai trò chi tiết nếu chưa có yêu cầu nghiệp vụ mới.

## 3. Luồng sử dụng chính

1. Mở bản sao workbook của hồ sơ.
2. Giữ `Hàng TK tối đa = 2` hoặc nhập số từ `0` đến `4`.
3. Nhập chủ đất ở dòng 1; nếu hồ sơ có hai chủ đất thì nhập thêm dòng 2. Nếu chỉ có một chủ đất, để trống dòng 2.
4. Nhập liên tục những người còn lại theo thứ tự từng nhánh.
5. Nếu Hàng TK của dòng nhánh chưa đúng, bấm một lần vào ô `Hàng TK` để chuyển `H1 → H2 → ... → Hàng TK tối đa → H1`.
6. Nhập ngày theo một trong ba dạng: `dd/mm/yyyy`, `mm/yyyy` hoặc `yyyy`.
7. Tích `Nhận đất` cho người nhận.
8. Bấm `Xuất Văn bản`.

Không có bước chọn vai trò sau mỗi lần thêm người và không có chế độ “đi vào/đi ra nhánh”.

## 4. Ví dụ cấu trúc

| STT | Người | Hàng TK | Thuộc nhánh của | Kết quả do hệ thống tính |
| ---: | --- | :---: | --- | --- |
| 1 | Ông A | H0 | — | Đã chết |
| 2 | Bà B | H0 | — | Đã chết |
| 3 | Anh C | H1 | Nhóm chủ đất ban đầu | Đã chết |
| 4 | Chị C1 | H2 | Anh C | Nhận đất |
| 5 | Anh C2 | H2 | Anh C | Từ chối |
| 6 | Chị D | H1 | Nhóm chủ đất ban đầu | Nhận đất |

Hai chủ đất ở `H0` được coi là một nhóm gốc chung, không phải hai nhánh riêng. Mọi người `H1` trực thuộc nhóm chủ đất ban đầu và để trống `ParentNguoiID`. Từ `H2` trở đi, hệ thống phải lưu `ParentNguoiID`, vì hai người cùng Hàng TK có thể thuộc các nhánh cha khác nhau.

## 5. Mô hình dữ liệu

### 5.1. Bảng người

| Trường | Hiển thị | Ý nghĩa |
| --- | :---: | --- |
| `NguoiID` | Ẩn | ID bất biến do hệ thống sinh, ví dụ `P0001` |
| `STTNhap` | Có, chỉ đọc | Thứ tự nhập và thứ tự hiển thị ổn định |
| `HoTen` | Có, được nhập | Họ tên dùng để tạo văn bản |
| `NgaySinh`, `NgayChet`, `NgayCap` | Có, được nhập | Text hiển thị đúng nội dung người dùng đã nhập |
| `SoGiayTo`, `DiaChi` | Có, được nhập | Thông tin nhân thân dùng để tạo văn bản |
| `HangTK` | Có, chỉ bấm | Số nguyên từ `0` đến `HangTKToiDa`; chủ đất là `0`, dòng nhánh là `1` trở lên |
| `NhanDat` | Có, chỉ bấm | Lựa chọn người nhận đất |
| `TrangThai` | Ẩn | Kết quả do engine tính: đã chết, nhận đất hoặc từ chối |
| `ParentNguoiID` | Ẩn | Để trống ở H0–H1; từ H2 trở đi là người ở Hàng TK liền trước mà nhánh trực thuộc |
| `LaChuDat` | Ẩn | Cờ chủ đất; hai vị trí đầu được gán tự động trong MVP |
| `NhomTuChoiID` | Ẩn | Khóa nhóm văn bản từ chối, để trống khi chưa phân nhóm |
| `NgaySinhGoc`, `NgayChetGoc`, `NgayCapGoc` | Ẩn | Chuỗi người dùng đã nhập, được giữ ở dạng text để không bị Excel tự đổi |
| `NgaySinhTinh`, `NgayChetTinh`, `NgayCapTinh` | Ẩn | Ngày thật của Excel dùng để so sánh và tính toán |

Không lưu trường nhập tay `TuChoi`. Đây là kết quả được engine tính để tránh các trạng thái mâu thuẫn.

### 5.2. Cấu hình hồ sơ

| Trường | Giá trị |
| --- | --- |
| `HangTKToiDa` | Mặc định `2`, hợp lệ từ `0` đến `4` |
| `BangMauHangTK` | Mã bảng màu đang chọn |
| `PhienBanCauTruc` | `2.0.0` |
| `NguoiIDTiepTheo` | Số dùng để sinh ID người tiếp theo |
| `SucChuaNguoi` | Sức chứa hiện tại của bảng nhập |

## 6. Vùng được nhập và bảo vệ ô

Workbook phải phân biệt rõ vùng người dùng được gõ và vùng chỉ được bấm:

- Ô cấu hình được nhập: `B4`.
- Vùng dữ liệu được nhập: `B9:G38`, gồm họ tên, ngày sinh, ngày chết, số giấy tờ, ngày cấp và địa chỉ.
- Vùng chỉ được bấm: `H9:I38`, gồm Hàng TK và Nhận đất. Người dùng không được gõ đè vào đây.
- Vùng hệ thống: cột `A` và các cột `J:T`; luôn khóa. Các cột `J:T` phải ẩn khỏi giao diện.
- Các sheet `CauHinh`, `KiemTra`, `XuatAn` được bảo vệ; `CauHinh` và `KiemTra` ẩn trong sử dụng bình thường, `XuatAn` ở trạng thái rất ẩn.
- Mật khẩu bảo vệ chỉ nhằm chống sửa nhầm, không được coi là biện pháp bảo mật dữ liệu.
- VBA được phép cập nhật các ô khóa bằng chế độ `UserInterfaceOnly`, và phải khôi phục chế độ này mỗi lần mở workbook.

Trên giao diện phải có nhãn rõ:

- `VÙNG NHẬP DỮ LIỆU` trên các cột người dùng được gõ.
- `BẤM ĐỂ CHỌN` trên các cột hành động.

### 6.1. Tốc độ nhập liệu

- Khi một ô thông thường thay đổi, VBA chỉ xử lý dòng vừa thay đổi và các dữ liệu thật sự liên quan.
- Không quét lại toàn bộ ngày, trạng thái và nhánh sau mỗi ô họ tên, số giấy tờ hoặc địa chỉ.
- Chỉ làm mới toàn bộ bảng khi mở workbook, trước khi kiểm tra/xuất Word hoặc khi thao tác Hàng TK làm thay đổi cả nhánh.
- Việc nhập rồi chuyển sang ô tiếp theo không được tạo độ trễ dễ nhận thấy trên máy làm việc thông thường.

## 7. Quy ước ngày tháng

### 7.1. Dạng nhập hợp lệ

Chỉ chấp nhận thứ tự ngày Việt Nam, không hiểu theo dạng tháng/ngày của Mỹ:

| Người dùng nhập | Excel và Word hiển thị/xuất | Giá trị ẩn dùng để tính |
| --- | --- | --- |
| `15/06/1950` | `15/06/1950` | ngày 15/06/1950 |
| `06/1950` | `06/1950` | ngày 01/06/1950 |
| `1950` | `1950` | ngày 01/01/1950 |
| để trống | để trống | không có ngày |

Ví dụ: sinh `1950`, chết `1999` phải hiển thị và xuất Word đúng là `1950`, `1999`. Chỉ engine hiểu ngầm là sinh `01/01/1950`, chết `01/01/1999` để tính tuổi lúc chết là `49`.

### 7.2. Nguyên tắc xử lý

- Các ô ngày nhìn thấy trên sheet nhập phải có định dạng text (`@`) để Excel không tự đổi năm hoặc đảo ngày/tháng.
- Giá trị nhìn thấy trong Excel và đưa vào Word phải giữ đúng nội dung người dùng đã nhập; không tự thêm `01/01` hoặc `01` vào phần hiển thị.
- Việc bổ sung ngày 01 và tháng 01 chỉ diễn ra trong trường `...Tinh` bị ẩn để phục vụ tính toán.
- Một chuỗi mơ hồ như `02/03/2020` luôn được hiểu là ngày 02 tháng 03, không được đảo thành ngày 03 tháng 02.
- Ô trống phải giữ nguyên trống ở cả dữ liệu gốc, dữ liệu hiển thị và dữ liệu tính. Không được sinh ngày giả như `01/01/1990`.
- Ngày không tồn tại, ví dụ `31/02/1990`, là lỗi chặn xuất.
- Ngày chết trước ngày sinh là lỗi chặn xuất.
- Dữ liệu được dán từ Excel có thể là text, số ngày của Excel hoặc ngày thật. Nếu còn là text thì phải giữ nguyên text; engine đồng thời chuyển một bản ẩn thành ngày thật để tính.
- Với dữ liệu người dùng gõ dạng text, trường `...Goc` phải giữ đúng chuỗi đã nhập, ví dụ `1950` hoặc `06/1950`.
- Nếu dữ liệu dán tới chỉ còn số ngày nội bộ của Excel, định dạng chữ ban đầu không thể khôi phục chắc chắn; khi đó trường gốc lưu ngày chuẩn `dd/mm/yyyy` để tránh đoán sai.
- Các phép tính tuổi, so sánh ngày và xác định đã chết chỉ dùng trường `...Tinh`, không tính trực tiếp từ text đang hiển thị.
- Dữ liệu xuất trung gian phải có cả text dùng để hiển thị/Word và ngày thật dùng để tính. Word mặc định lấy text, không lấy ngày đã bổ sung trong cột ẩn.

## 8. Thanh công cụ phía trên

Thanh công cụ cố định phía trên vùng nhập:

```text
Hàng TK tối đa  [ 2 ]                         [ Xuất Văn bản ]
```

Quy tắc:

- `Hàng TK tối đa` mặc định `2`, tối thiểu `0`, tối đa `4`.
- `Xuất Văn bản` là hành động chính và đặt bên phải, đúng trong vùng `H3:I4`.
- Nếu giảm mức tối đa thấp hơn Hàng TK đang có dữ liệu, không tự sửa dữ liệu; chặn và chỉ rõ các dòng cần xử lý.
- Nếu tăng mức tối đa, bảng màu và vòng chuyển Hàng TK cập nhật ngay.

## 9. Chủ đất và gán nhánh

- Hai vị trí đầu là chủ đất mặc định và cố định ở `H0`.
- Hồ sơ một chủ đất: nhập vị trí 1, để vị trí 2 trống.
- Hồ sơ hai chủ đất: nhập cả hai vị trí.
- Các dòng có dữ liệu ở hai vị trí này được gán `HangTK = 0` và `LaChuDat = True`.
- Dòng nhánh đầu tiên mặc định là `H1`.
- Dòng mới trong phần nhánh mặc định kế thừa Hàng TK của dòng ngay phía trên.
- Với dòng nhánh, bấm ô Hàng TK để chuyển trong phạm vi `H1..HangTKToiDa`; ở mức cuối thì vòng về `H1`, không vòng về `H0`.
- Sau khi xử lý, vùng chọn chuyển sang ô Họ tên cùng dòng để lần bấm sau luôn được nhận.

`H0` là nhóm chủ đất. `H1` là các nhánh trực tiếp của cả nhóm chủ đất nên `ParentNguoiID` để trống.

Khi một dòng được gán `Hk` với `k >= 2`, hệ thống tìm ngược lên dòng gần nhất có Hàng TK `H(k-1)` và lưu ID của dòng đó vào `ParentNguoiID`. Sau lần gán đầu tiên, `ParentNguoiID` là nguồn sự thật; thứ tự dòng và màu chỉ phục vụ nhập liệu.

Không chấp nhận:

- `HangTK < 0` hoặc `HangTK > HangTKToiDa`.
- Dòng nhánh có `HangTK = 0`.
- Dòng `Hk` với `k >= 2` không có dòng cha `H(k-1)` hợp lệ phía trên.
- Dòng H0 hoặc H1 có `ParentNguoiID`.
- Dòng từ H2 trở đi thiếu cha, tự trỏ hoặc trỏ đến người không tồn tại.
- Hàng TK của người cha không thấp hơn đúng một cấp so với người con.

## 10. Màu Hàng TK

Màu phải đủ khác nhau giữa các hàng, đồng thời chữ H0...H4 vẫn là dấu hiệu chính:

| Hàng TK | Nền nhạt | Màu đậm ở ô Hàng TK |
| :---: | :---: | :---: |
| H0 | `#D5E4FF` | `#2F5DA8` |
| H1 | `#CDEFD8` | `#237A50` |
| H2 | `#FFE1A0` | `#B26A00` |
| H3 | `#E2D0F6` | `#6747A8` |
| H4 | `#F6C8D2` | `#A33D54` |

- Không dùng đỏ lỗi làm màu chính của Hàng TK.
- Màu chỉ giúp nhìn nhanh, không phải dữ liệu và không được xuất sang Word.

## 11. Người nhận đất và trạng thái ẩn

- Cột `Nhận đất` hiển thị trực tiếp trên mỗi dòng và có thể chọn nhiều người.
- Người đã chết không được nhận đất. Nếu một dòng vừa có ngày chết vừa được chọn nhận đất thì đó là lỗi chặn xuất.
- Không suy luận người nhận theo Hàng TK hoặc quan hệ.
- Cột `TrangThai` phải ẩn khỏi giao diện vì đây là kết quả kỹ thuật, không phải nội dung người dùng cần thao tác.
- Engine vẫn phải tính trạng thái để chia nhóm xuất Word.

Quy tắc người từ chối:

```text
DaChet = NgayChetTinh có giá trị
TuChoi = (không DaChet) AND (không NhanDat)
```

Do đó:

- Người đã chết thuộc nhóm `Đã chết`, tuyệt đối không thuộc nhóm từ chối.
- Người còn sống và được chọn nhận đất thuộc nhóm `Nhận đất`.
- Người còn sống và không nhận đất thuộc nhóm `Từ chối`.
- Không có checkbox `Từ chối` và không nhập tay trạng thái.

## 12. Chuẩn bị cho nhóm từ chối

Phiên bản hiện tại chưa có giao diện chia nhóm từ chối. Mỗi dòng vẫn có trường ẩn `NhomTuChoiID` và hàm xuất trả về một tập nhóm:

```text
LayCacNhomTuChoi() -> Collection<NhomTuChoi>

NhomTuChoi:
  NhomID
  ParentNguoiID
  DanhSachNguoi
```

Quy tắc MVP:

- Nếu mọi `NhomTuChoiID` đều trống, đưa người từ chối vào nhóm `TC_DEFAULT`.
- API xuất Word nhận một collection có một phần tử, không nhận một danh sách phẳng.
- Tính năng chia nhóm sau này không được làm đổi công thức `TuChoi` hoặc hợp đồng xuất.

## 13. Hợp đồng dữ liệu xuất Word

Lớp lập dữ liệu xuất tạo ít nhất các nhóm sau:

| Nhóm | Cách lấy |
| --- | --- |
| `ChuDat` | Dòng có `LaChuDat = True` và có dữ liệu |
| `NguoiDaChet` | Dòng có `NgayChetTinh` |
| `NguoiNhanDat` | Dòng còn sống và `NhanDat = True` |
| `CacNhomTuChoi` | Collection tạo theo mục 12 |
| `CayNhanh` | Nhóm gốc H0, các dòng H1 trực thuộc nhóm gốc, và `NguoiID`, `HangTK`, `ParentNguoiID` từ H2 trở đi |

Dữ liệu xuất trung gian có các cột ngày hiển thị, ngày gốc, ngày dùng để tính và `TuoiLucChet`. Template Word không được đọc màu ô, vị trí dòng 1–2 hoặc suy luận nhánh từ STT.

## 14. Kiểm tra trước khi xuất

### Lỗi chặn xuất

- `HangTKToiDa` nằm ngoài `0..4`.
- Có `HangTK` lớn hơn mức tối đa.
- Có Hàng TK bị nhảy cấp.
- H0/H1 có `ParentNguoiID`, hoặc dòng từ H2 trở đi thiếu cha, tự trỏ, trỏ sai người hay không ở Hàng TK liền trước.
- Có chu kỳ trong chuỗi `ParentNguoiID`.
- Ngày sai định dạng, ngày không tồn tại hoặc ngày chết trước ngày sinh.
- Người đã chết vẫn có `NhanDat = True`.
- Cùng một `NguoiID` xuất hiện ở nhiều dòng.
- Dữ liệu vượt sức chứa của template được chọn.
- Còn placeholder Word chưa được thay.

### Cảnh báo

- Dòng chủ đất thứ hai để trống: hiểu là hồ sơ một chủ đất, không coi là lỗi.
- Có hai dòng nghi cùng một người do trùng họ tên và ngày sinh.
- Nhánh có người đã chết nhưng chưa có dòng ở Hàng TK kế tiếp.
- Người từ chối chưa được chia nhóm: dùng `TC_DEFAULT` ở MVP.

Mọi thông báo phải ghi rõ STT và họ tên; khi có thể, cho phép bấm để nhảy tới dòng lỗi.

## 15. Ngoài phạm vi phiên bản này

- Không chọn hoặc mô hình hóa vai trò cha, mẹ, con, vợ/chồng.
- Không tự kết luận hàng thừa kế, thế vị hoặc phần di sản.
- Không có giao diện chọn 3+ đồng chủ đất.
- Không có giao diện chia nhóm người từ chối.
- Không vẽ sơ đồ gia đình.
- Không dùng double-click.

## 16. Các hướng cũ không sử dụng

- **Mã đường dẫn `1`, `1.4`, `1.4.4`:** bỏ vì con chung không thuộc riêng nhánh 1 hay 2, đổi cấu trúc phải đánh lại mã cả cây và dễ nhân đôi người.
- **Quan hệ chi tiết `ChaID/MeID/VoChongID`:** không dùng cho MVP vì làm gián đoạn cách nhập hiện tại và mô hình hóa nhiều hơn nhu cầu thực tế.
- **Điểm được giữ:** mỗi người có một `NguoiID`, chỉ nhập một lần; dùng `ParentNguoiID` để nhánh không phụ thuộc màu hay vị trí dòng.

Chỉ xem xét lại các hướng trên khi có yêu cầu pháp lý hoặc xuất văn bản mà mô hình `HangTK + ParentNguoiID` không biểu diễn được.

## 17. Điều kiện nghiệm thu v0.2.1

- Workbook được tạo ở file mới, không ghi đè mẫu nguồn.
- Chỉ `B4` và `B9:G38` cho phép gõ; các vùng khác được khóa đúng mục 6.
- Cột trạng thái và các cột kỹ thuật được ẩn.
- Hàng TK dùng H0–H4, điều khiển tối đa dùng 0–4 và mặc định 2.
- Màu H0–H4 đúng bảng màu ở mục 10 và đủ khác nhau khi nhìn trên màn hình.
- Ngày `yyyy`, `mm/yyyy`, `dd/mm/yyyy`, ô trống và dữ liệu dán được xử lý đúng mục 7.
- Ví dụ sinh `1950`, chết `1999` phải hiển thị và xuất đúng `1950`, `1999`, trong khi engine vẫn tính tuổi `49`.
- Nhập dữ liệu thông thường chỉ xử lý dòng liên quan, không làm mới toàn bộ bảng sau mỗi ô.
- Người đã chết không bị đưa vào nhóm từ chối.
- File Excel và file Word thử vượt kiểm tra cấu trúc, không còn placeholder chưa thay.
