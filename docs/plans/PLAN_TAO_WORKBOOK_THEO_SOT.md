# Kế hoạch tạo workbook Excel theo SOT thừa kế

> SOT là “nguồn quyết định chính”: tài liệu được ưu tiên cao nhất khi các tài liệu khác nói khác nhau.
>
> SOT dùng cho kế hoạch này: `docs/specs/inheritance-branch-architecture.md`, chốt ngày 27/08/2026.
>
> `docs/specs/noi-dung-ngoai-sot.md` chỉ là danh sách chờ quyết định. Không tự đưa các nội dung trong đó vào bản MVP.

## 1. Mục tiêu

Tạo một workbook Excel có macro, phục vụ một hồ sơ thừa kế, với luồng làm việc ngắn:

1. Nhập một hoặc hai chủ đất ở hai dòng đầu.
2. Nhập liên tục những người còn lại.
3. Bấm ô `Tầng` để đổi `T2 → T3 → ... → Tn → T2`.
4. Nhập ngày chết nếu người đó đã chết.
5. Tích `Nhận đất` cho người nhận.
6. Kiểm tra dữ liệu.
7. Chuẩn bị đúng các nhóm dữ liệu để xuất Word.

Bản đầu không mô hình hóa vai trò bố, mẹ, con hoặc vợ/chồng; không dùng mã nhánh dạng `1.4.4`; không tự kết luận hàng thừa kế hoặc phần di sản.

## 2. Nguyên tắc an toàn

- Không sửa trực tiếp file nguồn `templates/excel/Dữ liệu thừa kế (2).xlsb`.
- Trước mỗi lần thử nghiệm, tạo một bản sao có ngày giờ trong tên file.
- Bản phát triển đặt trong `output/workbook-dev/`; thư mục `output/` không đưa vào Git.
- Khi đọc file `.xlsb`, mở Excel ở chế độ ẩn, chỉ đọc và tắt macro.
- Không đụng vào cửa sổ Excel hoặc Word người dùng đang mở.
- Nếu workbook đang mở và có dữ liệu chưa lưu, dừng lại và yêu cầu người dùng lưu/đóng trước.
- Khi thử xuất Word, luôn tạo file Word mới; không ghi đè mẫu Word và không đóng các cửa sổ Word có sẵn.
- Ghi lại mã băm SHA-256 của file nguồn. Có thể hiểu mã băm như “dấu vân tay” của file, dùng để chứng minh bản gốc không bị thay đổi.

## 3. Giả định triển khai cần xác nhận trước khi viết VBA

Các mục này chưa được SOT chốt hoàn toàn. Kế hoạch tạm dùng lựa chọn được đề xuất để công việc không bị mơ hồ:

| Mục | Đề xuất cho MVP | Lý do |
| --- | --- | --- |
| Định dạng file phát triển | `.xlsm` | Có thể chứa VBA, dễ kiểm tra hơn `.xlsb`; vẫn giữ nguyên `.xlsb` nguồn |
| Cách quản lý hồ sơ | Một workbook là một hồ sơ | Phù hợp luồng “mở bản sao workbook của hồ sơ” trong SOT |
| Số tầng | Mặc định 3, cho phép 1–5 | Đúng SOT |
| Màu tầng | Chọn một dải màu từ `prototypes/demo-nhanh-con.html` trước khi hoàn thiện giao diện | SOT yêu cầu người dùng chọn một dải màu |
| Checkbox `Nhận đất` | Ưu tiên checkbox nằm trong ô; nếu phiên bản Excel không hỗ trợ thì dùng ô TRUE/FALSE bấm một lần | Tránh Form Control/ActiveX bị lệch khi thêm hoặc lọc dòng |
| Trường nhân thân hiển thị | Tạm đề xuất: Họ tên, Ngày sinh, Ngày chết, Số giấy tờ, Ngày cấp, Địa chỉ | SOT chỉ ghi “Họ tên và thông tin nhân thân”, nên cần người dùng xác nhận danh sách cuối |
| Xuất Word | Trước hết tạo lớp dữ liệu xuất đúng SOT; nối từng mẫu Word sau khi rà placeholder | SOT đã chốt nhóm dữ liệu nhưng chưa chốt ánh xạ từng mẫu Word |

Nếu người dùng đổi một lựa chọn trên, chỉ cập nhật phần liên quan; không thay đổi mô hình lõi `Tang + ParentNguoiID`.

## 4. Cấu trúc workbook đề xuất

### 4.1. Sheet `NhapLieu`

Đây là sheet người dùng làm việc hằng ngày.

- Thanh trên cùng:
  - `Số tầng`: ô số nguyên từ 1 đến 5, mặc định 3.
  - `Kiểm tra dữ liệu`.
  - `Xuất Văn bản`.
- Bảng người `tblNguoi`:
  - Hai dòng đầu dành cho chủ đất, cố định ở T1.
  - Các dòng tiếp theo mặc định là T2 hoặc kế thừa tầng của dòng ngay trên.
  - Toàn bộ dòng đổi màu theo tầng.
  - Ô `Tầng` có hình thức giống nút và chạy bằng một lần bấm.

Các cột lõi:

| Cột | Người dùng thấy | Cách tạo |
| --- | :---: | --- |
| `NguoiID` | Không | VBA sinh, ví dụ `P0001`; không đổi sau khi đã sinh |
| `STTNhap` | Có | VBA duy trì thứ tự nhập ổn định |
| `HoTen` | Có | Người dùng nhập |
| Các trường nhân thân đã duyệt | Có | Người dùng nhập |
| `NgayChet` | Có | Có ngày nghĩa là người đã chết |
| `Tang` | Có | Lưu số 1–5, hiển thị `T1`–`T5` |
| `ParentNguoiID` | Không | Trống ở T1–T2; từ T3 trở đi trỏ tới người ở tầng ngay trên |
| `LaChuDat` | Không trong MVP | TRUE cho dòng chủ đất có dữ liệu |
| `NhanDat` | Có | TRUE/FALSE, có thể chọn nhiều người |
| `NhomTuChoiID` | Không trong MVP | Mặc định trống; lớp xuất dùng `TC_DEFAULT` |
| `TrangThaiHienThi` | Có hoặc chỉ dùng kiểm tra | Công thức/VBA cho ra `Đã chết`, `Nhận đất` hoặc `Từ chối`; không lưu cột `TuChoi` |

Quy tắc trạng thái:

```text
Đã chết  = NgayChet có giá trị
Nhận đất = chưa chết và NhanDat = TRUE
Từ chối  = chưa chết và NhanDat = FALSE
```

### 4.2. Sheet `CauHinh`

Sheet kỹ thuật, ẩn với người dùng bình thường.

| Tên | Ý nghĩa |
| --- | --- |
| `SoTang` | Mặc định 3, chỉ nhận 1–5 |
| `BangMauTang` | Mã dải màu đã chọn |
| `PhienBanCauTruc` | Phiên bản cấu trúc, bắt đầu từ `1.0.0` |
| `NguoiIDTiepTheo` | Số dùng để sinh ID tiếp theo, không tái sử dụng ID đã xóa |

Sheet này cũng chứa bảng màu T1–T5 và các danh mục dùng cho kiểm tra dữ liệu.

### 4.3. Sheet `KiemTra`

Chứa kết quả kiểm tra theo dạng bảng:

| Mức | Mã lỗi | STT | Họ tên | Nội dung | Ô cần sửa |
| --- | --- | ---: | --- | --- | --- |

- `Lỗi`: chặn xuất.
- `Cảnh báo`: cho phép tiếp tục sau khi người dùng xem.
- Nếu có thể, bấm vào dòng lỗi để nhảy tới ô cần sửa.
- Sheet chỉ hiện khi có lỗi/cảnh báo hoặc khi người dùng bấm `Kiểm tra dữ liệu`.

### 4.4. Sheet `XuatAn`

Đặt ở trạng thái `VeryHidden`, nghĩa là không thể hiện lại bằng thao tác Excel thông thường.

Sheet này không phải nguồn dữ liệu. VBA tạo lại mỗi lần xuất, gồm tối thiểu:

- `tblXuatChuDat`.
- `tblXuatNguoiDaChet`.
- `tblXuatNguoiNhanDat`.
- `tblXuatNhomTuChoi`.
- `tblXuatCayNhanh`.

Mục đích là giúp kiểm tra dữ liệu đã chuẩn bị trước khi đưa sang Word. Mã xuất Word vẫn phải đọc dữ liệu thật bằng `NguoiID`, `Tang`, `ParentNguoiID`, không đọc màu hoặc suy đoán theo vị trí dòng.

## 5. Các giai đoạn thực hiện

### Giai đoạn 0 — Chốt đầu vào và tạo vùng làm việc an toàn

Việc làm:

- Xác nhận bốn lựa chọn: định dạng `.xlsm`, các cột nhân thân, dải màu, mẫu Word đầu tiên dùng để thử.
- Kiểm tra file nguồn có đang mở hoặc có thay đổi chưa lưu hay không.
- Ghi kích thước, ngày sửa và SHA-256 của file nguồn.
- Tạo bản sao làm việc trong `output/workbook-dev/`.
- Ghi tên bản sao và dấu vân tay nguồn vào nhật ký triển khai.

Hoàn tất khi:

- File nguồn không bị sửa.
- Có bản sao mở được, macro bị tắt khi kiểm tra ban đầu.
- Các lựa chọn MVP được ghi rõ.

### Giai đoạn 1 — Khảo sát workbook cũ

Việc làm:

- Dùng OfficeCLI cho phần Office Open XML mà công cụ đọc được.
- Với `.xlsb`, dùng Excel cài trên máy ở chế độ ẩn, chỉ đọc, tắt macro để lấy:
  - Danh sách sheet và trạng thái ẩn.
  - Table, Name, công thức, data validation, conditional formatting.
  - Module VBA, sự kiện workbook/sheet và thư viện tham chiếu.
  - Các vùng dữ liệu cũ cần chuyển.
- Xuất mã VBA cũ ra file văn bản để đọc và so sánh; không chạy macro cũ.
- Lập bảng “giữ / viết lại / bỏ” cho từng sheet và module.

Hoàn tất khi:

- Biết rõ dữ liệu nào cần giữ.
- Biết lỗi cũ nào không được mang sang bản mới.
- Không có bước triển khai nào còn phụ thuộc vào việc đoán cấu trúc workbook cũ.

### Giai đoạn 2 — Dựng khung workbook mới

Việc làm:

- Tạo các sheet `NhapLieu`, `CauHinh`, `KiemTra`, `XuatAn`.
- Tạo `tblNguoi` với đúng cột lõi.
- Tạo Name cho `SoTang`, `BangMauTang`, `PhienBanCauTruc`.
- Thiết lập kiểu dữ liệu:
  - Ngày lưu là ngày Excel thật, không lưu thành chuỗi.
  - `Tang` là số nguyên.
  - ID là chuỗi.
  - `NhanDat`, `LaChuDat` là TRUE/FALSE.
- Khóa các cột kỹ thuật và bảo vệ cấu trúc cần thiết, nhưng vẫn cho người dùng thêm dòng và nhập dữ liệu.
- Đóng băng hàng tiêu đề để cuộn danh sách dài vẫn nhìn thấy tên cột.

Hoàn tất khi:

- Thêm, sửa, xóa một người không làm hỏng bảng.
- ID, ngày và TRUE/FALSE giữ đúng kiểu dữ liệu.
- Sheet kỹ thuật không cản trở việc nhập liệu.

### Giai đoạn 3 — Làm giao diện tầng và nhánh

Việc làm:

- Áp màu T1–T5 bằng conditional formatting. Đây là tô màu tự động theo giá trị `Tang`, không phải tô tay.
- Viết sự kiện một-lần-bấm cho ô `Tầng`:
  - Dòng chủ đất không đổi khỏi T1.
  - Dòng nhánh chuyển từ T2 đến `SoTang`, sau đó vòng về T2.
  - Sau khi đổi tầng, con trỏ chuyển sang ô `HoTen` cùng dòng.
- Khi gán Tk với k > 2, tìm ngược dòng gần nhất ở T(k-1) và ghi `ParentNguoiID`.
- Khi đổi tầng của người đã có hậu duệ, nhận diện khối hậu duệ liên tiếp, hỏi xác nhận rồi dịch chuyển cả khối.
- Nếu một dòng trong khối vượt giới hạn 1–`SoTang`, hủy toàn bộ thao tác.
- Khi giảm `SoTang`, chặn nếu đang có người ở tầng cao hơn và chỉ rõ các dòng đó.

Hoàn tất khi:

- Màu chỉ giúp nhìn; xóa màu không làm mất cấu trúc nhánh.
- Sau khi sắp xếp hoặc lọc, quan hệ đã lưu vẫn dựa vào `ParentNguoiID`.
- Không thể tạo nhánh thiếu cha hoặc nhảy tầng bằng thao tác bình thường.

### Giai đoạn 4 — Trạng thái người và quy tắc chủ đất

Việc làm:

- Hai dòng đầu có dữ liệu tự nhận `Tang = 1`, `LaChuDat = TRUE`.
- Dòng chủ đất thứ hai trống là trường hợp hợp lệ.
- Hiển thị checkbox `NhanDat` hoặc cách bấm tương đương trong ô.
- Tính trạng thái phái sinh:
  - Có ngày chết → `Đã chết`.
  - Còn sống và nhận đất → `Nhận đất`.
  - Còn sống và không nhận đất → `Từ chối`.
- Nếu nhập ngày chết cho người đang nhận đất, cảnh báo; khi người dùng đồng ý thì bỏ `NhanDat`.
- Không tạo cột nhập tay `TuChoi`.

Hoàn tất khi:

- Mỗi người chỉ có một trạng thái rõ ràng.
- Người đã chết không thể đồng thời nhận đất.
- Mã xử lý lọc chủ đất theo `LaChuDat`, không viết cứng `STT <= 2`.

### Giai đoạn 5 — Bộ kiểm tra trước khi xuất

Viết một hàm kiểm tra chung, dùng cho cả nút `Kiểm tra dữ liệu` và `Xuất Văn bản`.

Lỗi phải chặn xuất:

1. `SoTang` ngoài 1–5.
2. Có `Tang > SoTang` hoặc tầng nhỏ hơn 1.
3. Có nhảy tầng, ví dụ T3 chưa có T2 hợp lệ phía trên.
4. T1/T2 có `ParentNguoiID`.
5. T3 trở lên thiếu cha, tự trỏ, trỏ người không tồn tại hoặc cha không ở tầng ngay trên.
6. Có vòng tròn cha-con; ví dụ A trỏ B nhưng B lại trỏ A.
7. Người đã chết vẫn có `NhanDat = TRUE`.
8. Trùng `NguoiID`.
9. Số người hoặc dữ liệu vượt sức chứa mẫu Word đang chọn.
10. Sau khi xuất vẫn còn placeholder chưa được thay.

Cảnh báo:

- Dòng chủ đất thứ hai trống.
- Hai người có cùng họ tên và ngày sinh.
- Nhánh có người chết nhưng chưa có dòng ở tầng kế tiếp.
- Người từ chối chưa chia nhóm; MVP dùng `TC_DEFAULT`.

Hoàn tất khi:

- Mỗi thông báo ghi rõ STT, họ tên và ô cần sửa.
- Không có trường hợp lỗi chặn mà nút xuất vẫn chạy tiếp.
- Danh sách lỗi được xóa và tạo lại sạch sau mỗi lần kiểm tra.

### Giai đoạn 6 — Tạo lớp dữ liệu xuất

Việc làm:

- Từ `tblNguoi`, tạo năm nhóm đúng SOT:
  - `ChuDat`.
  - `NguoiDaChet`.
  - `NguoiNhanDat`.
  - `CacNhomTuChoi`.
  - `CayNhanh`.
- Khi chưa chia nhóm từ chối, tạo collection có một nhóm `TC_DEFAULT`; không truyền một chuỗi danh sách phẳng.
- Ghi bản xem trước của các nhóm vào `XuatAn` để kiểm tra.
- Viết hàm nối danh sách người dùng chung, xử lý đúng trường hợp 0, 1, 2 và nhiều người.
- Không đọc màu, STT hoặc vị trí dòng 1–2 để suy ra nhánh/chủ đất.

Hoàn tất khi:

- Cùng một dữ liệu đầu vào luôn tạo cùng một kết quả xuất.
- Có thể kiểm tra từng người đã đi vào đúng nhóm nào.
- Các nhóm xuất không phụ thuộc giao diện đang lọc hay sắp xếp.

### Giai đoạn 7 — Nối với mẫu Word đầu tiên

Chỉ bắt đầu sau khi chọn mẫu Word thử và rà xong placeholder.

Việc làm:

- Lập bảng ánh xạ giữa placeholder và năm nhóm dữ liệu xuất.
- Kiểm tra sức chứa thật của mẫu: tối đa bao nhiêu người, tài sản và nhóm từ chối.
- Nút xuất phải tạo một phiên Word riêng bằng `CreateObject`.
- Thay placeholder trong thân văn bản, bảng, header, footer và textbox.
- Lưu thành file Word mới trong `output/`; không ghi đè template.
- Chỉ đóng tài liệu và phiên Word do macro vừa tạo.
- Quét lại file kết quả; còn placeholder thì báo lỗi.

Hoàn tất khi:

- Xuất được một DOCX mới với dữ liệu giả lập.
- Một tài liệu Word khác đang mở vẫn nguyên trạng.
- Không còn placeholder chưa thay.
- Dữ liệu vượt sức chứa mẫu bị chặn trước khi xuất, không bị mất âm thầm.

### Giai đoạn 8 — Kiểm thử và đóng gói

Chạy ít nhất các tình huống sau bằng dữ liệu giả:

| Mã | Tình huống | Kết quả cần đạt |
| --- | --- | --- |
| TC01 | Một chủ đất, hai tầng | Dòng 2 trống chỉ cảnh báo, không chặn |
| TC02 | Hai chủ đất, ba tầng | Hai chủ đất cùng thuộc nhóm gốc |
| TC03 | T3 thuộc hai nhánh T2 khác nhau | `ParentNguoiID` đúng cho từng nhánh |
| TC04 | Đổi T3 thành T2 khi có hậu duệ | Cả khối dịch đúng hoặc bị chặn an toàn |
| TC05 | Giảm `SoTang` dưới tầng đang dùng | Bị chặn và chỉ rõ dòng |
| TC06 | Người nhận đất sau đó nhập ngày chết | Cảnh báo và bỏ nhận đất sau xác nhận |
| TC07 | T3 thiếu T2 | Chặn xuất |
| TC08 | Cha-con tạo vòng tròn | Chặn xuất |
| TC09 | Hai `NguoiID` trùng nhau | Chặn xuất |
| TC10 | Người sống không nhận đất | Tự vào nhóm từ chối |
| TC11 | Chưa chia nhóm từ chối | Dùng `TC_DEFAULT` |
| TC12 | Sắp xếp/lọc bảng | Quan hệ không đổi vì dùng ID |
| TC13 | Vượt sức chứa mẫu Word | Chặn trước khi tạo file |
| TC14 | Word khác đang mở | Không đóng hoặc sửa Word của người dùng |
| TC15 | Còn placeholder | Báo lỗi, không coi là xuất thành công |

Kiểm tra thêm:

- Không có lỗi công thức `#REF!`, `#NAME?`, `#VALUE!` hoặc vòng lặp công thức.
- Không có macro tự lưu workbook khi người dùng đóng nếu chưa được đồng ý.
- Không có liên kết ngoài hoặc add-in bắt buộc không được ghi rõ.
- Giao diện không bị cắt chữ ở mức zoom thông thường.
- Các cột kỹ thuật bị ẩn/khóa nhưng vẫn có thể kiểm tra khi mở chế độ phát triển.
- Mở lại file trên một phiên Excel mới vẫn hoạt động.

Hoàn tất khi:

- Toàn bộ TC01–TC15 đạt.
- Có bản template sạch và một bản dữ liệu giả đã kiểm thử.
- File nguồn `.xlsb` giữ nguyên dấu vân tay ban đầu.
- Mã VBA được xuất ra dạng text để có thể review và đưa vào Git.

## 6. Cách chia mã VBA đề xuất

Không viết toàn bộ vào một module lớn. Chia theo trách nhiệm để dễ sửa:

| Thành phần | Trách nhiệm |
| --- | --- |
| `ThisWorkbook` | Khởi tạo an toàn; không tự lưu ngoài ý muốn |
| Sheet `NhapLieu` | Nhận sự kiện bấm ô tầng và thay đổi ngày chết/nhận đất |
| `modNguoi` | Sinh `NguoiID`, thêm/xóa người, cập nhật chủ đất |
| `modTangNhanh` | Đổi tầng, tìm cha, dịch khối hậu duệ |
| `modTrangThai` | Tính Đã chết/Nhận đất/Từ chối |
| `modValidation` | Tạo danh sách lỗi và cảnh báo |
| `modExportData` | Tạo năm nhóm dữ liệu xuất |
| `modWordExport` | Mở Word riêng, thay placeholder, lưu và đóng đúng phiên |
| `modCommon` | Hàm dùng chung, xử lý lỗi và khôi phục trạng thái Excel |

Mỗi hàm xử lý sự kiện phải luôn khôi phục `EnableEvents`, `ScreenUpdating` và `Calculation` kể cả khi có lỗi. Nếu không, Excel có thể rơi vào trạng thái “bấm mà không chạy” cho đến khi mở lại file.

## 7. Thứ tự bàn giao

1. `M1 — Khung dữ liệu`: sheet, bảng, ID, kiểu dữ liệu.
2. `M2 — Giao diện tầng`: bấm một lần, màu, cha-con.
3. `M3 — Trạng thái và validation`: nhận đất, từ chối, lỗi/cảnh báo.
4. `M4 — Dữ liệu xuất`: năm nhóm dữ liệu và `XuatAn`.
5. `M5 — Xuất Word thử`: nối một template, kiểm tra an toàn Word.
6. `M6 — Bản phát hành thử`: template sạch, dữ liệu giả, mã VBA text và báo cáo test.

Sau mỗi mốc:

- Lưu thành file mới có số phiên bản.
- Chạy lại các test liên quan.
- Không chồng thêm chức năng của mốc sau khi mốc hiện tại chưa ổn định.

## 8. Những nội dung cố ý chưa làm

Các mục sau chỉ được đưa vào khi người dùng chốt riêng:

- Nhiều loại hồ sơ trong cùng workbook.
- Hồ sơ hai bên và cột `Ben`.
- Khối di sản, đồng sở hữu và tỷ lệ phân chia.
- Ba chủ đất trở lên trên giao diện.
- Chia nhiều nhóm người từ chối trên giao diện.
- Sơ đồ gia đình.
- Vai trò bố/mẹ/con/vợ/chồng.
- Mã nhánh `1`, `1.4`, `1.4.4`.
- Tự kết luận hàng thừa kế, thế vị hoặc phần di sản.
- Chuyển dữ liệu thật từ workbook cũ trước khi bản dùng dữ liệu giả vượt qua kiểm thử.

## 9. Điều kiện coi là hoàn thành

Workbook chỉ được coi là hoàn thành khi đồng thời đạt các điều kiện:

- Đúng mô hình `Tang + ParentNguoiID` của SOT.
- Người dùng nhập liên tục, không phải chọn vai trò gia đình.
- Hai dòng chủ đất, đổi tầng, nhận đất và trạng thái từ chối hoạt động đúng.
- Tất cả lỗi chặn và cảnh báo của SOT đã được kiểm thử.
- Năm nhóm dữ liệu xuất được tạo đúng và kiểm tra được.
- Xuất được một file Word mới mà không ảnh hưởng các cửa sổ Word đang mở.
- Không ghi đè workbook/Word mẫu nguồn.
- Không mất dữ liệu âm thầm khi vượt sức chứa template.
- Có bản template sạch, bản kiểm thử dữ liệu giả, mã VBA dạng text và báo cáo kết quả test.

