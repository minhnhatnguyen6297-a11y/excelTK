# Đặc tả hồ sơ thừa kế — SOT tổng

> Trạng thái: **§1–§9 đã triển khai trong workbook MVP v0.2.1. Các thay đổi bố cục, trường mở rộng, placeholder và cách kiểm tra tại §2–§14 đã chốt lại ngày 01/09/2026, chưa triển khai (đích v0.3.0).**
> Ngày cập nhật: 01/09/2026.
> Phiên bản cấu trúc dữ liệu: hiện hành `2.0.0`; đích của đặc tả này là `2.2.0`.
> Phạm vi: toàn bộ luồng một hồ sơ thừa kế — nhập người, nhập tài sản, nhập thông tin phụ, kiểm tra, xuất Word.

Đây là nguồn quyết định chính (SOT) duy nhất cho hồ sơ thừa kế. Tài liệu được sắp theo đúng thứ tự thao tác của người dùng, không theo lớp kỹ thuật.

Nội dung chờ quyết định, chưa thuộc SOT: [`noi-dung-ngoai-sot.md`](noi-dung-ngoai-sot.md). Không tự đưa nội dung từ file đó vào triển khai.

## Mục lục

| Mục | Nội dung | Trạng thái |
| --- | --- | --- |
| §1 | Mục tiêu và mô hình một-file-một-hồ-sơ | Đã triển khai |
| §2 | Quyết định kiến trúc đã chốt | §2.1 đã triển khai; §2.2–2.3 cập nhật cho v0.3.0 |
| §3 | Luồng sử dụng đầy đủ | §3.1–3.2 đã triển khai; §3.3–3.5 đích v0.3.0 |
| §4 | Bố cục sheet `NhapLieu` | Vùng người đã triển khai; tài sản/phụ đích v0.3.0 |
| §5 | Bước 1 — Nhập người | Trường chuẩn đã triển khai; vùng mở rộng đích v0.3.0 |
| §6 | Quy ước ngày tháng và Text | Xử lý ngày đã triển khai; chuẩn Text toàn bộ đích v0.3.0 |
| §7 | Hàng TK, nhánh và màu | Đã triển khai |
| §8 | Người nhận đất, từ chối, trạng thái | Đã triển khai |
| §9 | Bảo vệ ô và tốc độ nhập | Vùng chuẩn đã triển khai; bảo vệ vùng mở rộng đích v0.3.0 |
| §10 | Bước 2 — Nhập tài sản theo thẻ dọc | Đích v0.3.0 |
| §11 | Bước 3 — Thông tin phụ của hồ sơ | Đích v0.3.0 |
| §12 | Bước 4 — Kiểm tra trước khi xuất | Cơ chế chẩn đoán không chặn xuất là đích v0.3.0 |
| §13 | Bước 5 — Xuất Word | Template kiểm thử đã chạy; mẫu thật đích v0.3.0 |
| §14 | Danh mục tra cứu | Đích v0.3.0 |
| §15 | Ngoài phạm vi | — |
| §16 | Các hướng cũ không sử dụng | — |
| §17 | Điều kiện nghiệm thu | — |

---

## 1. Mục tiêu và mô hình một-file-một-hồ-sơ

Một workbook `.xlsm` là một hồ sơ. Mở hồ sơ mới bằng cách copy file mẫu sạch vào thư mục của hồ sơ đó rồi nhập liệu. Không có cơ chế chọn hồ sơ trong file, không có `HoSoID`.

Giữ cách nhập quen thuộc: mỗi người nằm trên một dòng và người dùng nhập liên tục, không phải dừng lại để chọn vai trò cha, mẹ, con hay vợ/chồng.

Mỗi hồ sơ cần trả lời sáu câu hỏi:

1. Ai là chủ đất ban đầu?
2. Mỗi người thuộc Hàng TK và nhánh nào?
3. Ai đã chết?
4. Ai nhận đất? Những người còn sống còn lại được xác định là từ chối.
5. Di sản gồm những thửa đất nào?
6. Hồ sơ niêm yết ở đâu, số công chứng nào, ai là người ủy quyền?

Thông tin dùng để tính phải được bảo vệ khỏi sửa nhầm. Ngày tháng phải chịu được trường hợp chỉ biết năm, chỉ biết tháng/năm hoặc dữ liệu được dán từ nơi khác.

## 2. Quyết định kiến trúc đã chốt

### 2.1. Người và nhánh

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

### 2.2. Tài sản

- **Mỗi tài sản là một thẻ dọc độc lập**, nhãn ở bên trái và giá trị ở bên phải. Không dùng bảng ngang một-dòng-một-tài-sản.
- **Ba thẻ tài sản đặt cạnh nhau**, cùng bắt đầu ở hàng 8 hoặc 9, phía bên phải bảng Người trên cùng sheet `NhapLieu`. Vị trí chuẩn là `AA:AB`, `AC:AD`, `AE:AF` tương ứng Tài sản 1, 2, 3.
- Bảng Người dùng vùng `A:I`; vùng `J:Z` là 17 cột dự phòng cho trường Người mở rộng. Các cột dự phòng chưa dùng được ẩn; khi có tiêu đề thì hiện cột đó.
- Cách bố trí này giữ được dạng thẻ dọc nhưng không tạo một vùng dài 60 hàng. Người dùng chỉ cuộn ngang khi cần xem hoặc nhập tài sản, không phải cuộn xuống dưới bảng Người.
- Số thẻ cố định là **3**. Thẻ không dùng thì để trống.
- Người dùng nhập trực tiếp vào thẻ. VBA dò nhãn trường, đọc ô giá trị bên cạnh và dựng bảng ẩn `tblTaiSan` trên `XuatAn` để phục vụ tính toán/xuất; người dùng không nhập vào bảng ẩn đó.
- Các cột kỹ thuật của Người và Tài sản không còn đặt ở `J:W` trên `NhapLieu`; chúng phải chuyển sang bảng hệ thống trên sheet rất ẩn `XuatAn` để không chiếm vùng mở rộng.

### 2.3. Trường nhập và trường mở rộng

- Mọi ô nhập nhìn thấy của Người, Tài sản và Hồ sơ mặc định có định dạng Text (`@`). Chuỗi `05`, `00123` hoặc `1/2` phải được giữ nguyên, không để Excel tự đổi thành số hay ngày.
- Trường chuẩn do workbook cung cấp sẵn. Trường mở rộng do người dùng khai báo trực tiếp bằng cách gõ tiêu đề/nhãn vào ô dành sẵn, không dùng form và không cần sửa VBA.
- Không tạo sheet `Trường mở rộng`, không tạo `tblTruongTuDo` và không bắt người dùng đăng ký key-value ở nơi khác; tiêu đề/nhãn ngay trên `NhapLieu` là nguồn khai báo duy nhất.
- Trường mở rộng có thể là dữ liệu nhập tay hoặc công thức Excel. Công thức là trách nhiệm của người dùng; VBA chỉ đọc kết quả đang có trong ô, không tự suy luận hay viết lại công thức.
- Nút `Đồng bộ trường` chỉ sao chép cấu trúc cần thiết: áp dụng công thức xuống các dòng Người cùng nhóm, hoặc sao chép nhãn/công thức từ thẻ Tài sản 1 sang thẻ 2 và 3. Nút này không sao chép giá trị nhập tay.
- VBA luôn dò trường theo tên tiêu đề/nhãn đã chuẩn hóa, không dựa vào số thứ tự cột hoặc độ lệch hàng cố định.

### 2.4. Thông tin phụ

- Thông tin Hồ sơ nằm trong khối `B40:C51` trên `NhapLieu`: cột `B` là nhãn, cột `C` là giá trị. Khối này chỉ dài 12 hàng và không làm ba thẻ Tài sản kéo dài xuống dưới.
- MVP có sẵn các trường: niêm yết, số công chứng, người ủy quyền, người ủy quyền 2.
- Khối có các hàng trống dự phòng để người dùng tự thêm nhãn, dữ liệu hoặc công thức theo cùng hợp đồng tại §11.

### 2.5. Không quay lại

Không quay lại mô hình nhập vai trò chi tiết nếu chưa có yêu cầu nghiệp vụ mới. Không mô hình hóa khối di sản, tỷ lệ phân chia hay đồng chủ thứ ba trở lên trong phiên bản này.

## 3. Luồng sử dụng đầy đủ

### 3.1. Chuẩn bị

1. Copy file mẫu sạch vào thư mục hồ sơ, đổi tên theo hồ sơ.
2. Mở file, bật macro.
3. Giữ `Hàng TK tối đa = 2` hoặc nhập số từ `0` đến `4`.

### 3.2. Bước 1 — Nhập người

4. Nhập chủ đất ở dòng 1; nếu hồ sơ có hai chủ đất thì nhập thêm dòng 2. Nếu chỉ có một chủ đất, để trống dòng 2.
5. Nhập liên tục những người còn lại theo thứ tự từng nhánh.
6. Nếu Hàng TK của dòng nhánh chưa đúng, bấm một lần vào ô `Hàng TK` để chuyển `H1 → H2 → ... → Hàng TK tối đa → H1`.
7. Nhập ngày theo một trong ba dạng: `dd/mm/yyyy`, `mm/yyyy` hoặc `yyyy`.
8. Tích `Nhận đất` cho người nhận.

Không có bước chọn vai trò sau mỗi lần thêm người và không có chế độ “đi vào/đi ra nhánh”.

### 3.3. Bước 2 — Nhập tài sản

9. Cuộn ngang sang phải tới khối `THÔNG TIN TÀI SẢN` bắt đầu từ cột `AA`.
10. Điền thẻ `Tài sản 1`. Loại sổ, hình thức sử dụng, loại đất, cơ quan cấp chọn từ danh sách.
11. Nếu hồ sơ có thửa thứ hai hoặc thứ ba, điền thẻ `Tài sản 2`, `Tài sản 3` ở ngay bên cạnh. Thẻ không dùng để trống.

### 3.4. Bước 3 — Thông tin phụ

12. Điền khối `THÔNG TIN HỒ SƠ`: xã niêm yết, số công chứng, người ủy quyền.
13. Bấm `Chọn nơi lấy mẫu` để chọn thư mục chứa các mẫu Word `.docx`.
14. Bấm `Chọn văn bản` và chọn một mẫu `.docx` trong thư mục đã chọn. Ô `Mẫu Word` chỉ hiển thị tên file để kiểm tra, không nhận nhập trực tiếp.
15. Bấm `Chọn nơi xuất` để chọn thư mục nhận file Word cuối cùng.

### 3.5. Bước 4–5 — Kiểm tra và xuất

16. Có thể bấm `Kiểm tra dữ liệu` để xem danh sách gợi ý sửa. Đây là bước hỗ trợ, không phải cổng chặn xuất.
17. Bấm `Xuất Văn bản`. File Word mới được tạo trong thư mục đã chọn ở bước 15; lỗi dữ liệu trong ô được thể hiện rõ trong Word để người dùng tự tìm và sửa. Workbook không tự mở file và không tự mở thư mục.

## 4. Bố cục sheet `NhapLieu`

Một sheet, các khối đặt cạnh nhau để tránh vùng nhập dài. Bảng Người ở trái; ba thẻ Tài sản ở phải.

```text
hàng 1–4     THANH CÔNG CỤ
hàng 6–7     nhãn vùng Người và nhãn THÔNG TIN TÀI SẢN
hàng 8       tiêu đề bảng Người; tiêu đề TÀI SẢN 1, 2, 3
hàng 9–38    30 dòng Người; đồng thời là các hàng trường của ba thẻ Tài sản
hàng 40–51   THÔNG TIN HỒ SƠ tại B:C

cột A:I      bảng Người chuẩn
cột J:Z      17 cột dự phòng trường Người mở rộng
cột AA:AB    nhãn | giá trị của Tài sản 1
cột AC:AD    nhãn | giá trị của Tài sản 2
cột AE:AF    nhãn | giá trị của Tài sản 3
```

- Không đặt ba thẻ Tài sản nối tiếp nhau thành vùng 60 hàng.
- Không dùng cột `J:W` làm nơi chứa dữ liệu kỹ thuật. Dữ liệu kỹ thuật chuyển sang `XuatAn`.
- Các cột `J:Z` chưa có tiêu đề được ẩn để giao diện gọn; `Đồng bộ trường` hiện cột đã được khai báo.
- Mỗi thẻ Tài sản có cùng tập nhãn theo cùng thứ tự. Thẻ 1 là bản mẫu khi thêm trường mới; xem §10.6.
- Nếu cần đúng 20 cột dự phòng cho Người thì phải dời ba thẻ bắt đầu từ `AD`; đây là thay đổi bố cục, không được tự làm trong code. Bản v0.3.0 chốt dùng 17 cột `J:Z` và bắt đầu thẻ tại `AA`.

### 4.1. Thanh công cụ

```text
Hàng TK tối đa: 2        | Kiểm tra dữ liệu | Xuất Văn bản
Nơi lấy mẫu: <thư mục>    | Chọn nơi lấy mẫu
Mẫu Word: 1. PCDS .docx   | Chọn văn bản
Nơi xuất: <thư mục>       | Chọn nơi xuất
```

- `Hàng TK tối đa` mặc định `2`, tối thiểu `0`, tối đa `4`.
- `Xuất Văn bản` là hành động chính và đặt ở vùng dễ thấy bên trái, không đè lên cột trường mở rộng.
- Nếu giảm mức tối đa thấp hơn Hàng TK đang có dữ liệu, không tự sửa dữ liệu; chặn và chỉ rõ các dòng cần xử lý.
- Nếu tăng mức tối đa, bảng màu và vòng chuyển Hàng TK cập nhật ngay.
- Giao diện giữ cách trình bày trực tiếp trên sheet như bản cũ; không tạo UserForm hoặc wizard. Chỉ giữ các chức năng `Chọn nơi lấy mẫu`, `Chọn văn bản`, `Chọn nơi xuất`, `Kiểm tra dữ liệu`, `Đồng bộ trường` và `Xuất Văn bản`; bỏ nút hoặc tùy chọn Word khác.
- `Chọn nơi lấy mẫu` dùng hộp chọn thư mục. Hủy hộp chọn thì giữ nguyên cấu hình. Chọn thành công cập nhật `ThuMucMauWord` và làm mới nguồn mẫu `.docx`.
- Ô `Mẫu Word` tại `C5` bị khóa và chỉ hiển thị tên mẫu đang chọn; `MauWordDangChon` giữ đường dẫn đầy đủ. Nút `Chọn văn bản` là đường chọn mẫu duy nhất: mở hộp chọn file bắt đầu tại `ThuMucMauWord` và chỉ nhận `.docx` nằm trực tiếp trong thư mục này.
- Hủy hộp `Chọn văn bản` giữ nguyên `MauWordDangChon`, nội dung `C5`, `SucChuaNguoiMau` và `SucChuaTaiSanMau`.
- Khi đổi nơi lấy mẫu, nếu thư mục mới có file cùng tên với mẫu đang chọn thì cập nhật `MauWordDangChon` sang đường dẫn đầy đủ mới và làm mới sức chứa. Nếu không có file cùng tên thì xóa `MauWordDangChon`, xóa `C5`, đặt hai sức chứa mẫu về `0` và yêu cầu người dùng chọn lại; không tự chọn file đầu tiên.
- Đổi mẫu làm cập nhật thông tin sức chứa dùng cho chẩn đoán (§12.3), nhưng không tự chặn xuất.
- `Chọn nơi xuất` dùng hộp chọn thư mục và ghi `ThuMucXuat`; hủy hộp chọn thì giữ nguyên cấu hình. `Xuất Văn bản` dùng đúng thư mục này và không tự đổi về thư mục mặc định.
- Đường dẫn dài được rút gọn khi hiển thị trên `NhapLieu`, nhưng giá trị đầy đủ luôn nằm trong `CauHinh` và là nguồn duy nhất cho VBA.

### 4.2. Nhãn bắt buộc trên giao diện

- `VÙNG NHẬP DỮ LIỆU` trên các cột người dùng được gõ.
- `BẤM ĐỂ CHỌN` trên các cột hành động.
- `THÔNG TIN TÀI SẢN` trên ba thẻ cạnh nhau, kèm dòng phụ `Thẻ không dùng thì để trống`.
- `THÔNG TIN HỒ SƠ` trên khối hồ sơ.

---

## 5. Bước 1 — Nhập người

### 5.1. Ví dụ cấu trúc

| STT | Người | Hàng TK | Thuộc nhánh của | Kết quả do hệ thống tính |
| ---: | --- | :---: | --- | --- |
| 1 | Ông A | H0 | — | Đã chết |
| 2 | Bà B | H0 | — | Đã chết |
| 3 | Anh C | H1 | Nhóm chủ đất ban đầu | Đã chết |
| 4 | Chị C1 | H2 | Anh C | Nhận đất |
| 5 | Anh C2 | H2 | Anh C | Từ chối |
| 6 | Chị D | H1 | Nhóm chủ đất ban đầu | Nhận đất |

Hai chủ đất ở `H0` được coi là một nhóm gốc chung, không phải hai nhánh riêng. Mọi người `H1` trực thuộc nhóm chủ đất ban đầu và để trống `ParentNguoiID`. Từ `H2` trở đi, hệ thống phải lưu `ParentNguoiID`, vì hai người cùng Hàng TK có thể thuộc các nhánh cha khác nhau.

### 5.2. Bảng người `tblNguoi`

| Trường | Hiển thị | Ý nghĩa |
| --- | :---: | --- |
| `NguoiID` | Ẩn trên `XuatAn` | ID bất biến do hệ thống sinh, ví dụ `P0001` |
| `STTNhap` | Có, chỉ đọc | Thứ tự nhập và thứ tự hiển thị ổn định |
| `HoTen` | Có, được nhập | Họ tên dùng để tạo văn bản |
| `NgaySinh`, `NgayChet`, `NgayCap` | Có, được nhập | Text hiển thị đúng nội dung người dùng đã nhập |
| `SoGiayTo`, `DiaChi` | Có, được nhập | Thông tin nhân thân dùng để tạo văn bản |
| `HangTK` | Có, chỉ bấm | Số nguyên từ `0` đến `HangTKToiDa`; chủ đất là `0`, dòng nhánh là `1` trở lên |
| `NhanDat` | Có, chỉ bấm | Lựa chọn người nhận đất |
| `TrangThai` | Ẩn trên `XuatAn` | Kết quả do engine tính: đã chết, nhận đất hoặc từ chối |
| `ParentNguoiID` | Ẩn trên `XuatAn` | Để trống ở H0–H1; từ H2 trở đi là người ở Hàng TK liền trước mà nhánh trực thuộc |
| `LaChuDat` | Ẩn trên `XuatAn` | Cờ chủ đất; hai vị trí đầu được gán tự động trong MVP |
| `NhomTuChoiID` | Ẩn trên `XuatAn` | Khóa nhóm văn bản từ chối, để trống khi chưa phân nhóm |
| `NgaySinhGoc`, `NgayChetGoc`, `NgayCapGoc` | Ẩn trên `XuatAn` | Chuỗi người dùng đã nhập, được giữ ở dạng text để không bị Excel tự đổi |
| `NgaySinhTinh`, `NgayChetTinh`, `NgayCapTinh` | Ẩn trên `XuatAn` | Ngày thật của Excel dùng để so sánh và tính toán |
| `LoaiCC` | Ẩn trên `XuatAn` | Loại giấy tờ suy từ `NgayCapTinh`, xem §14.2 |
| `NoiCapCC` | Ẩn trên `XuatAn` | Nơi cấp tra theo `LoaiCC`, xem §14.2 |
| `NhanDiaChi` | Ẩn trên `XuatAn` | Nhãn địa chỉ suy theo `LoaiCC`, xem §14.2 |

Các trường kỹ thuật liên kết với dòng hiển thị bằng `NguoiID` hoặc `STTNhap`, không bằng vị trí cột. VBA không được dùng số cột cố định như `Cells(row, 10)` để tìm trường. Khi cần cột của `ListObject`, phải tìm `ListColumns` theo đúng tên tiêu đề.

Ba trường `LoaiCC`, `NoiCapCC`, `NhanDiaChi` do VBA tính, người dùng không gõ.

Không lưu trường nhập tay `TuChoi`. Đây là kết quả được engine tính để tránh các trạng thái mâu thuẫn.

### 5.3. Trường Người mở rộng

- Vùng dự phòng là `J8:Z38`: hàng 8 chứa tiêu đề, hàng 9–38 chứa giá trị của từng người.
- Người dùng gõ tiêu đề vào cột trống tiếp theo rồi nhập dữ liệu hoặc công thức. Không được chèn cột vào giữa `A:I`.
- Ô nhập tay có định dạng Text (`@`). Nếu một cột là trường tính, ô chứa công thức phải dùng `General` để Excel thực thi công thức; kết quả đưa sang Word vẫn luôn chuyển thành text. Đây là ngoại lệ kỹ thuật duy nhất cho yêu cầu định dạng Text.
- `Đồng bộ trường` tìm công thức đầu tiên trong cột và điền cùng mẫu công thức xuống 30 dòng bằng cơ chế tương đương `FillDown`. Dữ liệu nhập tay không được sao chép xuống dòng khác.
- VBA quét các tiêu đề không rỗng từ trái sang phải, chuẩn hóa tiêu đề theo §13.2 và đọc kết quả hiện tại của từng ô. VBA không tính thay công thức của người dùng.
- Không tự nới `ListObject` qua vùng Tài sản. Giới hạn cứng của vùng Người mở rộng là cột `Z` trong v0.3.0.

### 5.4. Cấu hình hồ sơ

| Trường | Giá trị |
| --- | --- |
| `HangTKToiDa` | Mặc định `2`, hợp lệ từ `0` đến `4` |
| `BangMauHangTK` | Mã bảng màu đang chọn |
| `PhienBanCauTruc` | `2.2.0` |
| `NguoiIDTiepTheo` | Số dùng để sinh ID người tiếp theo |
| `TaiSanIDTiepTheo` | Số dùng để sinh ID tài sản tiếp theo |
| `SucChuaNguoi` | Sức chứa hiện tại của bảng nhập |
| `ThuMucMauWord` | Thư mục người dùng chọn làm nơi lấy mẫu `.docx` |
| `MauWordDangChon` | Đường dẫn mẫu Word đang chọn |
| `ThuMucXuat` | Thư mục người dùng chọn để nhận file Word cuối cùng |
| `SucChuaNguoiMau` | Số người tối đa mẫu đang chọn chứa được |
| `SucChuaTaiSanMau` | Số tài sản tối đa mẫu đang chọn chứa được |

## 6. Quy ước ngày tháng

Áp dụng cho mọi ô ngày trong workbook: ngày sinh, ngày chết, ngày cấp giấy tờ của người và ngày cấp sổ của tài sản.

Quy tắc nền cho **mọi ô nhập nhìn thấy**, không chỉ ô ngày:

- `NumberFormat = "@"` trước khi người dùng nhập hoặc dán dữ liệu.
- VBA lấy nội dung gốc bằng `Value2`, chuyển thành chuỗi có kiểm soát và không dùng `.Text` làm nguồn chính, vì `.Text` phụ thuộc độ rộng cột và định dạng đang hiển thị.
- Mã có số 0 đầu như `05`, `00123`, serial và số vào sổ phải được giữ đúng từng ký tự từ lúc nhập đến lúc xuất Word.
- Chuỗi giống ngày như `1/2` vẫn là text; Excel không được tự đổi thành ngày.
- Ô công thức là ngoại lệ: phải có định dạng `General` để công thức chạy. VBA đọc kết quả công thức rồi chuyển kết quả đó thành text khi lập dữ liệu xuất.

### 6.1. Dạng nhập hợp lệ

Chỉ chấp nhận thứ tự ngày Việt Nam, không hiểu theo dạng tháng/ngày của Mỹ:

| Người dùng nhập | Excel và Word hiển thị/xuất | Giá trị ẩn dùng để tính |
| --- | --- | --- |
| `15/06/1950` | `15/06/1950` | ngày 15/06/1950 |
| `06/1950` | `06/1950` | ngày 01/06/1950 |
| `1950` | `1950` | ngày 01/01/1950 |
| để trống | để trống | không có ngày |

Ví dụ: sinh `1950`, chết `1999` phải hiển thị và xuất Word đúng là `1950`, `1999`. Chỉ engine hiểu ngầm là sinh `01/01/1950`, chết `01/01/1999` để tính tuổi lúc chết là `49`.

### 6.2. Nguyên tắc xử lý

- Các ô ngày nhìn thấy trên sheet nhập phải có định dạng text (`@`) để Excel không tự đổi năm hoặc đảo ngày/tháng.
- Giá trị nhìn thấy trong Excel và đưa vào Word phải giữ đúng nội dung người dùng đã nhập; không tự thêm `01/01` hoặc `01` vào phần hiển thị.
- Việc bổ sung ngày 01 và tháng 01 chỉ diễn ra trong trường `...Tinh` bị ẩn để phục vụ tính toán.
- Một chuỗi mơ hồ như `02/03/2020` luôn được hiểu là ngày 02 tháng 03, không được đảo thành ngày 03 tháng 02.
- Ô trống phải giữ nguyên trống ở cả dữ liệu gốc, dữ liệu hiển thị và dữ liệu tính. Không được sinh ngày giả như `01/01/1990`.
- Ngày không tồn tại, ví dụ `31/02/1990`, tạo kết quả tính lỗi nhưng không chặn thao tác xuất; xem §12.
- Ngày chết trước ngày sinh được ghi vào danh sách chẩn đoán nhưng không chặn thao tác xuất.
- Dữ liệu được dán từ Excel có thể là text, số ngày của Excel hoặc ngày thật. Nếu còn là text thì phải giữ nguyên text; engine đồng thời chuyển một bản ẩn thành ngày thật để tính.
- Với dữ liệu người dùng gõ dạng text, trường `...Goc` phải giữ đúng chuỗi đã nhập, ví dụ `1950` hoặc `06/1950`.
- Nếu dữ liệu dán tới chỉ còn số ngày nội bộ của Excel, định dạng chữ ban đầu không thể khôi phục chắc chắn; khi đó trường gốc lưu ngày chuẩn `dd/mm/yyyy` để tránh đoán sai.
- Các phép tính tuổi, so sánh ngày và xác định đã chết chỉ dùng trường `...Tinh`, không tính trực tiếp từ text đang hiển thị.
- Dữ liệu xuất trung gian phải có cả text dùng để hiển thị/Word và ngày thật dùng để tính. Word luôn lấy text đã chuẩn hóa một lần ở lớp xuất, không lấy ngày đã bổ sung trong cột ẩn.

Mọi ô ngày dùng cùng một bộ xử lý. Không viết lại bộ phân tích ngày riêng cho tài sản.

## 7. Hàng TK, nhánh và màu

### 7.1. Chủ đất và gán nhánh

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

### 7.2. Màu Hàng TK

| Hàng TK | Nền nhạt | Màu đậm ở ô Hàng TK |
| :---: | :---: | :---: |
| H0 | `#D5E4FF` | `#2F5DA8` |
| H1 | `#CDEFD8` | `#237A50` |
| H2 | `#FFE1A0` | `#B26A00` |
| H3 | `#E2D0F6` | `#6747A8` |
| H4 | `#F6C8D2` | `#A33D54` |

- Không dùng đỏ lỗi làm màu chính của Hàng TK.
- Màu chỉ giúp nhìn nhanh, không phải dữ liệu và không được xuất sang Word.

## 8. Người nhận đất, từ chối và trạng thái

- Cột `Nhận đất` hiển thị trực tiếp trên mỗi dòng và có thể chọn nhiều người.
- Người đã chết không nên nhận đất. Nếu một dòng vừa có ngày chết vừa được chọn nhận đất thì ghi chẩn đoán rõ, nhưng vẫn cho phép xuất để người dùng nhìn kết quả và tự sửa.
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

### 8.1. Chuẩn bị cho nhóm từ chối

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

## 9. Bảo vệ ô và tốc độ nhập

### 9.1. Vùng được nhập

| Vùng | Quyền |
| --- | --- |
| `B4` | Nhập — Hàng TK tối đa |
| `C5` | Chỉ đọc — hiển thị tên Mẫu Word đang chọn |
| Nút `Chọn nơi lấy mẫu` | Chỉ bấm — chọn thư mục mẫu Word `.docx` |
| Nút `Chọn văn bản` | Chỉ bấm — chọn một mẫu `.docx` trong thư mục mẫu |
| Nút `Chọn nơi xuất` | Chỉ bấm — chọn thư mục nhận Word cuối |
| `B9:G38` | Nhập — họ tên, ngày sinh, ngày chết, số giấy tờ, ngày cấp, địa chỉ |
| `H9:I38` | Chỉ bấm — Hàng TK, Nhận đất |
| `J8:Z38` | Nhập tiêu đề, dữ liệu hoặc công thức của trường Người mở rộng |
| Ô nhãn/giá trị dành sẵn của thẻ tài sản | Nhập hoặc chọn, xem §10.3 và §10.6 |
| Ô giá trị của khối hồ sơ | Nhập hoặc chọn, xem §11 |
| Cột `A` và các ô giao diện không nêu trên | Hệ thống hoặc nhãn cố định, khóa |

- Các sheet `CauHinh`, `KiemTra`, `XuatAn` được bảo vệ; `CauHinh` và `KiemTra` ẩn trong sử dụng bình thường, `XuatAn` ở trạng thái rất ẩn. Toàn bộ dữ liệu kỹ thuật cũ ở `J:W` phải chuyển sang `XuatAn`.
- Sheet `DanhMuc` ẩn; người dùng chỉ gặp nội dung của nó qua danh sách chọn.
- Mật khẩu bảo vệ chỉ nhằm chống sửa nhầm, không được coi là biện pháp bảo mật dữ liệu.
- VBA được phép cập nhật các ô khóa bằng chế độ `UserInterfaceOnly`, và phải khôi phục chế độ này mỗi lần mở workbook.

### 9.2. Tốc độ nhập liệu

- Khi một ô thông thường thay đổi, VBA chỉ xử lý dòng vừa thay đổi và các dữ liệu thật sự liên quan.
- Không quét lại toàn bộ ngày, trạng thái và nhánh sau mỗi ô họ tên, số giấy tờ hoặc địa chỉ.
- Sửa một ô của thẻ tài sản chỉ xử lý thẻ đó, không quét lại bảng người.
- Sửa một trường mở rộng chỉ xử lý cột Người hoặc thẻ Tài sản liên quan. Chỉ nút `Đồng bộ trường` mới quét toàn vùng mở rộng để sao chép cấu trúc.
- Chỉ làm mới toàn bộ khi mở workbook, khi bấm `Đồng bộ trường`, trước khi kiểm tra/xuất Word hoặc khi thao tác Hàng TK làm thay đổi cả nhánh.
- Việc nhập rồi chuyển sang ô tiếp theo không được tạo độ trễ dễ nhận thấy trên máy làm việc thông thường.

---

## 10. Bước 2 — Nhập tài sản theo thẻ dọc

### 10.1. Hình dạng thẻ

Mỗi tài sản là một thẻ dọc hai cột: nhãn ở trái, giá trị ở phải. Ba thẻ có cùng hàng và đặt cạnh nhau:

```text
AA        AB            AC        AD            AE        AF
TÀI SẢN 1               TÀI SẢN 2               TÀI SẢN 3
Loại sổ  | giá trị      Loại sổ  | giá trị      Loại sổ  | giá trị
Serial   | giá trị      Serial   | giá trị      Serial   | giá trị
...                     ...                     ...
Ghi chú  | giá trị      Ghi chú  | giá trị      Ghi chú  | giá trị
trường tự thêm          trường tự thêm          trường tự thêm
```

Các ô giá trị dài bật `WrapText`. Không gộp ô giữa hai thẻ. Thẻ 2 và 3 để trống khi hồ sơ chỉ có một thửa.

### 10.2. Trường của một thẻ

| Trường | Nhập kiểu trên giao diện | Placeholder Word tài sản 1 |
| --- | --- | --- |
| `LoaiSo` | Chọn từ danh mục, lưu text | `{{loaiso1}}` |
| `Serial` | Text | `{{serial1}}` |
| `SoVaoSo` | Text | `{{sovaoso1}}` |
| `SoThua` | Text | `{{sothua1}}` |
| `SoTo` | Text | `{{soto1}}` |
| `DiaChiDat` | Text dài | `{{diachidat1}}` |
| `DienTich` | Text; VBA đổi sang số khi cần tính | `{{dientich1}}` |
| `HinhThucSuDung` | Chọn từ danh mục, lưu text | `{{hinhthucsudung1}}` |
| `LoaiDat` | Chọn từ danh mục, lưu text | `{{loaidat1}}` |
| `ThoiHan` | Text | `{{thoihan1}}` |
| `ONT` | Text; VBA đổi sang số khi cần tính | `{{ont1}}` |
| `CLN` | Text; VBA đổi sang số khi cần tính | `{{cln1}}` |
| `NTS` | Text; VBA đổi sang số khi cần tính | `{{nts1}}` |
| `LUC` | Text; VBA đổi sang số khi cần tính | `{{luc1}}` |
| `NguonGoc` | Text dài | `{{nguongoc1}}` |
| `NgayCapSo` | Text ngày theo §6 | `{{ngaycapso1}}` |
| `CoQuanCapSo` | Chọn từ danh mục, lưu text | `{{coquancapso1}}` |
| `GhiChu` | Text dài | `{{ghichu1}}` |

Tài sản 2 và 3 đổi số cuối thành `2` và `3`, ví dụ `{{serial2}}`, `{{dientich3}}`. Số cuối luôn là số thẻ trên giao diện; không dồn lại khi một thẻ phía trước để trống.

### 10.3. Dữ liệu ẩn của thẻ

Mỗi thẻ có một dòng tương ứng trong bảng hệ thống `tblTaiSan` trên sheet rất ẩn `XuatAn`:

| Trường ẩn | Ý nghĩa |
| --- | --- |
| `TaiSanID` | ID bất biến do hệ thống sinh, dạng `TS001`; sinh khi thẻ bắt đầu có dữ liệu |
| `NgayCapSoGoc` | Chuỗi người dùng đã nhập |
| `NgayCapSoTinh` | Ngày thật của Excel dùng để so sánh |
| `CoDuLieu` | Thẻ có dữ liệu hay không, xem §10.4 |

`TaiSanID` không hiển thị và người dùng không cần nhớ. Dòng kỹ thuật liên kết bằng số thẻ cố định 1–3 và `TaiSanID`, không liên kết bằng địa chỉ ô cứng.

Nhãn `Tài sản 1`, `Tài sản 2`, `Tài sản 3` là số thứ tự hiển thị đồng thời là số cuối của placeholder tĩnh; không phải khóa liên kết.

### 10.4. Thẻ có dữ liệu

Một thẻ được coi là **có dữ liệu** khi bất kỳ ô giá trị chuẩn hoặc mở rộng nào có kết quả khác rỗng. Chỉ có nhãn trường mà chưa có giá trị thì chưa tính là có dữ liệu.

Thẻ không có dữ liệu không sinh `TaiSanID`; dòng kỹ thuật tương ứng vẫn tồn tại nhưng toàn bộ giá trị xuất của thẻ là rỗng.

Thẻ chỉ có một trường lẻ vẫn được xuất đúng trường đó; `KiemTra` có thể nêu đây là dữ liệu nhập dở nhưng không chặn xuất.

### 10.5. Bảng xuất ẩn `tblTaiSan`

VBA đọc ba thẻ rồi dựng `tblTaiSan` trên sheet `XuatAn`:

- Ba dòng cố định tương ứng thẻ 1, 2, 3; cột `SoTheTaiSan` không đổi.
- Cột `TaiSanID` giữ nguyên ID của thẻ có dữ liệu; thẻ trống có ID rỗng.
- Các cột trường được tạo hoặc tìm theo tên nhãn đã chuẩn hóa, không theo số cột cố định.
- Bảng chỉ do VBA tạo/refresh trước khi kiểm tra và trước khi xuất. Người dùng không nhập vào bảng này.

VBA đọc giá trị bằng `Value2`. Ô nhập là Text nên giữ được số 0 đầu; ô công thức được đọc kết quả. Trước khi đưa vào Word, mọi giá trị được đổi thành text đúng một lần theo §13.4.

### 10.6. Thêm trường, đồng bộ và xóa thẻ

- Ba thẻ luôn tồn tại; v0.3.0 không có lệnh thêm thẻ thứ tư.
- Sau 18 trường chuẩn, mỗi thẻ dành các hàng còn lại đến hàng 38 làm vùng trường mở rộng. Không chèn thêm hàng giữa các trường chuẩn.
- Người dùng gõ nhãn mới ở hàng trống tiếp theo của **thẻ 1**, rồi nhập giá trị mẫu hoặc công thức nếu cần.
- Khi bấm `Đồng bộ trường`, VBA sao chép nhãn mới sang đúng hàng của thẻ 2 và 3. Nếu ô giá trị của thẻ 1 là công thức, sao chép công thức theo tham chiếu tương đối sang hai thẻ; nếu là dữ liệu nhập tay, giữ ô giá trị thẻ 2 và 3 trống.
- Nếu nhãn ba thẻ khác nhau, thẻ 1 là nguồn chuẩn và thao tác đồng bộ phải hỏi xác nhận trước khi ghi đè **nhãn hoặc công thức**, nhưng không được ghi đè dữ liệu nhập tay.
- Xóa một tài sản là xóa trắng các ô giá trị của thẻ đó. VBA xóa dữ liệu kỹ thuật tương ứng khi thẻ trở về trạng thái không có dữ liệu.

## 11. Bước 3 — Thông tin phụ của hồ sơ

Khối Hồ sơ dùng từng cặp ô `nhãn | giá trị`, cùng nguyên tắc với thẻ Tài sản nhưng mỗi trường chỉ có một giá trị cho cả hồ sơ.

| Trường | Nhập kiểu | Placeholder Word |
| --- | --- | --- |
| `NiemYet` | Text | `{{niemyet}}` |
| `SoCongChung` | Text | `{{socongchung}}` |
| `NguoiUyQuyen` | Chọn từ danh mục, lưu text | `{{nguoiuyquyen}}` |
| `NguoiUyQuyen2` | Chọn từ danh mục, lưu text | `{{nguoiuyquyen2}}` |

Quy tắc:

- `NiemYet` là xã hoặc địa bàn niêm yết văn bản.
- `SoCongChung` là số công chứng của văn bản. Để trống thì xuất chuỗi rỗng, không xuất `0`.
- `NguoiUyQuyen` và `NguoiUyQuyen2` chọn từ danh mục người ủy quyền (§14.3). Hồ sơ chỉ có một người ủy quyền thì để trống trường thứ hai.
- Khối có ít nhất 8 hàng trống dự phòng. Người dùng gõ nhãn mới vào hàng trống tiếp theo rồi nhập dữ liệu hoặc công thức ở ô giá trị.
- Trường Hồ sơ là giá trị đơn nên placeholder không có số cuối, trừ khi chính nhãn trường đã có số phân biệt như `NguoiUyQuyen2`.
- VBA quét nhãn không rỗng và đọc ô giá trị bên cạnh. Thêm trường Hồ sơ không cần sửa VBA.

---

## 12. Bước 4 — Kiểm tra trước khi xuất

`Kiểm tra dữ liệu` là công cụ chẩn đoán tùy chọn. Nó giúp người dùng nhìn nhanh các điểm đáng nghi, nhưng **không phải điều kiện để bấm Xuất Văn bản** và không tự chạy như một cổng chặn trước khi xuất.

### 12.1. Nội dung chẩn đoán

Danh sách ngắn gọn có thể gồm:

- Hàng TK ngoài giới hạn, nhảy cấp, thiếu liên kết cha hoặc có chu kỳ.
- Người đã chết vẫn được chọn nhận đất; không có người nhận đất; nghi trùng người.
- Ngày sai định dạng, ngày không tồn tại hoặc ngày chết trước ngày sinh.
- Tài sản thiếu trường thường cần, diện tích không đổi được sang số, số âm, hoặc tổng loại đất lệch diện tích.
- Mẫu Word có ít slot hơn số người hoặc số tài sản đang có.
- Tiêu đề/nhãn sau chuẩn hóa bị trùng nhau.

Mỗi dòng chẩn đoán nên ghi STT người hoặc số thẻ tài sản và vị trí ô để dễ sửa. Không dựng nhiều hộp thoại cảnh báo nối tiếp nhau.

### 12.2. Quy tắc không chặn xuất

- Giá trị ô lỗi Excel như `#VALUE!`, `#REF!`, `#N/A` được đưa nguyên ký hiệu lỗi sang Word.
- Trường đã biết nhưng ô để trống được thay bằng chuỗi rỗng.
- Placeholder đúng cú pháp nhưng không tìm thấy trường nguồn được thay bằng `#KHONG_CO_TRUONG:<ten>#`.
- Hai tiêu đề/nhãn khác nhau nhưng cùng chuẩn hóa thành một tên được thay bằng `#TRUNG_TEN:<ten>#` cho placeholder bị ảnh hưởng.
- Dữ liệu đáng nghi vẫn được xuất. Sau khi xuất có thể hiện **một** thông báo ngắn về tổng số dấu lỗi đã ghi vào Word, không hiện từng cảnh báo riêng.

Chỉ dừng xuất khi không thể tạo kết quả kỹ thuật, ví dụ: không mở được mẫu Word, đường dẫn không tồn tại, không tạo được phiên Word, không ghi được file hoặc tài liệu Word bị hỏng. Đây là lỗi hạ tầng, không phải lỗi nội dung hồ sơ.

### 12.3. Sức chứa mẫu tĩnh

Placeholder tĩnh có số cuối chỉ dùng được đến số slot đã đặt trong mẫu Word. Ví dụ mẫu có `{{ten1}}` đến `{{ten10}}` thì người thứ 11 không có vị trí để xuất.

- `KiemTra` nêu rõ sức chứa mẫu và số dữ liệu hiện có, nhưng không chặn xuất.
- Slot có trong mẫu nhưng không có dữ liệu được thay bằng chuỗi rỗng.
- Dữ liệu vượt slot không được tự ghép vào một slot khác và không được làm đổi số cuối.
- Placeholder danh sách/khối lặp là một loại khác, không dùng giới hạn slot này và chưa thuộc phạm vi v0.3.0; xem §13.3 và §15.

## 13. Bước 5 — Xuất Word

### 13.1. Nhóm dữ liệu xuất

Lớp lập dữ liệu xuất tạo ít nhất các nhóm sau:

| Nhóm | Cách lấy |
| --- | --- |
| `ChuDat` | Dòng có `LaChuDat = True` và có dữ liệu |
| `NguoiDaChet` | Dòng có `NgayChetTinh` |
| `NguoiNhanDat` | Dòng còn sống và `NhanDat = True` |
| `CacNhomTuChoi` | Collection tạo theo §8.1 |
| `CayNhanh` | Nhóm gốc H0, các dòng H1 trực thuộc nhóm gốc, và `NguoiID`, `HangTK`, `ParentNguoiID` từ H2 trở đi |
| `TaiSan` | Ba dòng `tblTaiSan` theo `SoTheTaiSan = 1..3` ở §10.5 |
| `HoSo` | Các trường ở §11 |

Dữ liệu xuất trung gian có các cột ngày hiển thị, ngày gốc, ngày dùng để tính và `TuoiLucChet`. Template Word không được đọc màu ô, vị trí dòng 1–2 hoặc suy luận nhánh từ STT.

### 13.2. Placeholder tĩnh

Toàn hệ thống chỉ dùng một cặp dấu `{{...}}`. Không còn cú pháp ngoặc vuông và không duy trì parser tương thích ngược.

Tên placeholder tĩnh có các quy tắc:

- Viết thường, không dấu, viết liền; chỉ gồm chữ `a-z` và số `0-9`.
- Không có dấu cách, dấu chấm, gạch dưới hoặc ký tự tiếng Việt.
- Trường lặp theo Người hoặc Tài sản có số thứ tự gắn ngay cuối tên: `{{tentruong1}}`, `{{tentruong2}}`.
- Trường chỉ có một giá trị cho cả Hồ sơ không cần số cuối: `{{niemyet}}`, `{{socongchung}}`.

Tên công khai của các trường Người chuẩn:

| Trường nguồn | Gốc placeholder | Ví dụ slot 1 |
| --- | --- | --- |
| `HoTen` | `ten` | `{{ten1}}` |
| `NgaySinh` | `namsinh` | `{{namsinh1}}` |
| `NgayChet` | `namchet` | `{{namchet1}}` |
| `SoGiayTo` | `cccd` | `{{cccd1}}` |
| `NgayCap` | `ngaycap` | `{{ngaycap1}}` |
| `DiaChi` | `diachi` | `{{diachi1}}` |
| `LoaiCC` | `loaicc` | `{{loaicc1}}` |
| `NoiCapCC` | `noicapcc` | `{{noicapcc1}}` |
| `NhanDiaChi` | `thuongtru` | `{{thuongtru1}}` |

Mẫu tĩnh hiện có dùng slot người `n = 1..10`. Quy tắc gán slot giữ nguyên:

```text
slot 1, 2   = chủ đất theo STTNhap; thiếu chủ đất thứ hai thì slot 2 rỗng
slot 3..n   = những người còn lại có dữ liệu, không phải chủ đất, theo STTNhap
```

Với trường mở rộng, VBA lấy gốc placeholder bằng cách chuẩn hóa chính tiêu đề/nhãn người dùng đã gõ: đổi `đ` thành `d`, bỏ dấu tiếng Việt, bỏ khoảng trắng và ký tự không phải chữ/số, rồi chuyển về chữ thường. Ví dụ `Số điện thoại` thành `sodienthoai`, từ đó sinh `{{sodienthoai4}}` cho người slot 4.

Vì placeholder tĩnh không có namespace, tên sau chuẩn hóa phải duy nhất trong toàn bộ dữ liệu xuất. Nếu hai nguồn cùng sinh một token, dùng dấu lỗi trùng tên ở §12.2; không âm thầm chọn nguồn đầu tiên.

### 13.3. Placeholder động để issue riêng

Placeholder danh sách, khối lặp hoặc bảng là loại có logic bên trong. Ví dụ nghiệp vụ gồm danh sách người nhận đất, nhóm từ chối và cây nhánh. Loại này **chưa được triển khai hoặc chốt cú pháp trong v0.3.0**.

- Không đưa token có dấu chấm hoặc gạch dưới vào parser placeholder tĩnh.
- Không cố thay placeholder động bằng chuỗi rỗng. Khi gặp token động chưa hỗ trợ, ghi dấu rõ `#CHUA_HO_TRO_DONG:<ten>#` trong Word.
- Việc định nghĩa nguồn dữ liệu, mẫu câu một phần tử, dấu nối và cách xử lý danh sách rỗng phải được chốt trong một issue/SOT riêng trước khi triển khai.
- Placeholder động không bị giới hạn bởi số slot tĩnh tại §12.3 sau khi được triển khai.

### 13.4. Quy tắc thay thế

- VBA lập một bảng tra `token -> text` bằng cách quét tiêu đề/nhãn, không bằng số cột hardcode.
- Giá trị trống xuất chuỗi rỗng. Số `0` do người dùng thật sự nhập phải xuất `0`; chỉ ngày kỹ thuật rỗng mới không được biến thành `00/01/1900`.
- Ngày xuất theo text người dùng đã nhập (§6), không dùng cột `...Tinh`.
- Giá trị nhập lấy bằng `Value2`; kết quả cuối được đổi sang chuỗi một lần trong lớp xuất. Không dùng `.Text` làm nguồn dữ liệu vì cột hẹp có thể trả về `####`.
- Ô công thức hợp lệ xuất kết quả đang tính. Ô công thức lỗi xuất nguyên mã lỗi Excel như `#VALUE!` hoặc `#REF!`.
- Nội dung có xuống dòng phải đổi sang ký tự newline của Word.
- Chuỗi dài hơn 254 ký tự phải dùng `TypeText` thay vì `Replacement.Text`, do giới hạn của Find/Replace.
- Parser chỉ nhận token tĩnh bắt đầu bằng chữ thường, theo sau chỉ là chữ thường hoặc số, tất cả nằm trong `{{...}}`; token sai cú pháp được đổi thành `#PLACEHOLDER_SAI:<noi_dung>#`, không đoán tên.
- Không để một placeholder bị tách thành nhiều run trong Word, vì Find/Replace sẽ không nhận diện được.
- Thêm trường mới chỉ cần: gõ tiêu đề/nhãn, nhập dữ liệu hoặc công thức, bấm `Đồng bộ trường` nếu cần nhân công thức/cấu trúc, rồi đặt placeholder đã chuẩn hóa vào Word. Không sửa VBA.
- Trước khi phát hành v0.3.0, mọi mẫu trong `templates/word/`, parser VBA, kiểm tra placeholder và script kiểm kê phải chuyển hoàn toàn sang `{{...}}`; không để hai cú pháp chạy song song.

### 13.5. An toàn phiên Office

- Dùng `CreateObject` để tạo phiên Word **riêng**. Không dùng `GetObject`.
  Lỗi cũ cần tránh: macro dùng `GetObject` bám vào phiên Word đang mở của người dùng rồi gọi `Quit`, làm đóng luôn tài liệu khác của người dùng.
- Chỉ mở và đóng tài liệu do phiên đó tạo; lưu file rồi thoát đúng phiên đó.
- Thay placeholder trong body, bảng, header, footer và textbox bằng Word Range. Không dùng `Selection.Find` vì không phủ hết header/footer/textbox.
- File Word xuất ra là bản sao độc lập của mẫu. Không dùng Mail Merge, không tạo liên kết Word–Excel.
- Sau khi xuất chỉ hiển thị kết quả. Không tự mở file, không tự mở thư mục.
- Không tự lưu workbook khi đóng.
- Tên file đầu ra có kiểm soát ký tự đặc biệt và chống trùng bằng hậu tố `_1`, `_2`.

### 13.6. Cổng chuyển đổi sang cú pháp duy nhất

Trước khi phát hành v0.3.0 phải hoàn thành đồng thời:

- `modExportData.bas` tạo duy nhất bảng tra token tĩnh theo §13.2 và nhận cả trường chuẩn lẫn trường mở rộng.
- `modWordExport.bas` là nơi duy nhất quét và thay token `{{...}}` trong các Word Range.
- `modXuatWord.bas` chỉ điều phối chọn mẫu, gọi lập dữ liệu, xuất file và báo tổng số dấu lỗi; không giữ một parser thứ hai.
- `scripts/inspect-word-placeholders.ps1` kiểm kê token `{{...}}`, báo token sai quy tắc, token trùng và mọi placeholder ngoặc vuông còn sót.
- Toàn bộ file `.docx` trong `templates/word/` được kiểm kê và chuyển token tĩnh sang cú pháp mới. Mẫu `.doc` cũ phải được chuyển sang `.docx` hoặc ghi rõ chưa được hỗ trợ; không âm thầm bỏ qua.
- Cổng phát hành phải thất bại nếu code, metadata tra cứu hoặc mẫu Word còn placeholder ngoặc vuông. Không giữ chế độ parser kép bằng cờ cấu hình.

## 14. Danh mục tra cứu

Sheet `DanhMuc` ẩn, chứa các bảng nguồn phục vụ ô chọn và tra cứu. Người dùng làm hồ sơ không nhập trực tiếp tại đây.

### 14.1. Các danh mục

| Danh mục | Dùng cho |
| --- | --- |
| Loại sổ | `LoaiSo` của thẻ tài sản |
| Loại đất | `LoaiDat`; gồm đất ở đô thị, đất ở nông thôn, cây lâu năm, nuôi trồng thủy sản, trồng lúa |
| Hình thức sử dụng | `HinhThucSuDung`; sử dụng riêng, sử dụng chung |
| Cơ quan cấp sổ | `CoQuanCapSo` |
| Loại giấy tờ → nơi cấp → nhãn địa chỉ | §14.2 |
| Người ủy quyền | `NguoiUyQuyen`, `NguoiUyQuyen2` |
| Tờ bản đồ 2025 | Tra cứu thủ công số tờ sau điều chỉnh |
| Sáp nhập thôn | Tra cứu thủ công thôn/xóm cũ và mới |

### 14.2. Quy tắc loại giấy tờ

Suy từ `NgayCapTinh` của từng người:

```text
NgayCapTinh trống            -> LoaiCC rỗng
NgayCapTinh <  01/07/2024    -> "Căn cước công dân"
NgayCapTinh >= 01/07/2024    -> "Căn cước"
```

`NoiCapCC` tra theo `LoaiCC` trong danh mục. `NhanDiaChi` suy theo `LoaiCC`: `Thường trú tại` hoặc `Cư trú tại`.

Ba giá trị này do VBA ghi vào bảng kỹ thuật Người trên `XuatAn` khi `NgayCap` của dòng đó thay đổi; không quét lại toàn bảng.

### 14.3. Hai bảng tra cứu chưa tự động

`Tờ bản đồ 2025` (điều chỉnh số tờ sau sáp nhập) và `Sáp nhập thôn` (thôn/xóm huyện Ý Yên) là hai `ListObject` riêng trên `DanhMuc`. Trong v0.3.0 chúng chỉ để người bảo trì tra cứu, chưa tự điền vào vùng nhập.

### 14.4. Phân loại dữ liệu kỹ thuật

| Loại dữ liệu | Nơi đặt | Ai được sửa |
| --- | --- | --- |
| Danh sách chọn và bảng tra cứu dùng chung | `DanhMuc`, mỗi nguồn là một `ListObject` có tên ổn định | Người bảo trì workbook; không sửa trong lúc làm hồ sơ |
| Cấu hình phiên bản, đường dẫn mẫu, sức chứa | `CauHinh` | VBA hoặc người bảo trì |
| Ngày tính, ID, trạng thái, bảng xuất trung gian | `XuatAn` rất ẩn | Chỉ VBA |
| Tiêu đề/nhãn và giá trị trường mở rộng của hồ sơ hiện tại | `NhapLieu` | Người dùng làm hồ sơ |

- Danh mục và bảng tra cứu **không** thuộc cơ chế trường mở rộng. Chúng chỉ trở thành dữ liệu xuất khi một giá trị đã được chọn hoặc công thức tra cứu đã đưa kết quả vào một ô trường trên `NhapLieu`.
- Tên `ListObject` là hợp đồng kỹ thuật; VBA tìm bảng theo tên, không theo địa chỉ cột/hàng.
- Việc cập nhật danh mục phải thực hiện trên file mẫu nguồn rồi phát hành phiên bản mới. Không sửa riêng từng file hồ sơ trừ khi cần xử lý một trường hợp khẩn cấp và người dùng hiểu file đó sẽ khác mẫu chuẩn.

---

## 15. Ngoài phạm vi phiên bản này

- Không chọn hoặc mô hình hóa vai trò cha, mẹ, con, vợ/chồng.
- Không tự kết luận hàng thừa kế, thế vị hoặc phần di sản.
- Không có giao diện chọn 3+ đồng chủ đất.
- Không có giao diện chia nhóm người từ chối.
- Không có khối di sản, quan hệ sở hữu hay bảng phân chia theo tỷ lệ phần trăm.
- Không có hồ sơ hai bên A/B, hồ sơ ủy quyền hay loại hồ sơ khác trong file này. Hồ sơ hai bên là workbook riêng, có spec riêng.
- Không có cột `Ben` trong workbook thừa kế.
- Không vẽ sơ đồ gia đình.
- Không dùng double-click.
- Không triển khai placeholder danh sách, block lặp hoặc bảng động trong v0.3.0; tách thành issue và lần chốt SOT riêng.

## 16. Các hướng cũ không sử dụng

- **Mã đường dẫn `1`, `1.4`, `1.4.4`:** bỏ vì con chung không thuộc riêng nhánh 1 hay 2, đổi cấu trúc phải đánh lại mã cả cây và dễ nhân đôi người.
- **Quan hệ chi tiết `ChaID/MeID/VoChongID`:** không dùng cho MVP vì làm gián đoạn cách nhập hiện tại và mô hình hóa nhiều hơn nhu cầu thực tế.
- **Bảng tài sản ngang một-dòng-một-tài-sản:** không dùng. Thay bằng ba thẻ dọc đặt cạnh nhau ở §10; mỗi thẻ vẫn đọc từ trên xuống.
- **Sheet tài sản riêng:** bỏ để người dùng nhập một mạch trên cùng sheet `NhapLieu`.
- **Điểm được giữ:** mỗi người có một `NguoiID`, chỉ nhập một lần; dùng `ParentNguoiID` để nhánh không phụ thuộc màu hay vị trí dòng; mỗi tài sản có một `TaiSanID` ổn định.

Chỉ xem xét lại các hướng trên khi có yêu cầu pháp lý hoặc xuất văn bản mà mô hình `HangTK + ParentNguoiID` không biểu diễn được.

## 17. Điều kiện nghiệm thu

### 17.1. Đã đạt ở v0.2.1

- Workbook được tạo ở file mới, không ghi đè mẫu nguồn.
- Ở workbook v0.2.1, chỉ `B4` và `B9:G38` cho phép gõ; đây là hiện trạng trước khi có vùng mở rộng v0.3.0.
- Cột trạng thái và các cột kỹ thuật được ẩn.
- Hàng TK dùng H0–H4, điều khiển tối đa dùng 0–4 và mặc định 2.
- Màu H0–H4 đúng bảng màu ở §7.2 và đủ khác nhau khi nhìn trên màn hình.
- Ngày `yyyy`, `mm/yyyy`, `dd/mm/yyyy`, ô trống và dữ liệu dán được xử lý đúng §6.
- Ví dụ sinh `1950`, chết `1999` phải hiển thị và xuất đúng `1950`, `1999`, trong khi engine vẫn tính tuổi `49`.
- Nhập dữ liệu thông thường chỉ xử lý dòng liên quan, không làm mới toàn bộ bảng sau mỗi ô.
- Người đã chết không bị đưa vào nhóm từ chối.
- File Excel và file Word thử vượt kiểm tra cấu trúc, không còn placeholder chưa thay.

### 17.2. Cần đạt ở v0.3.0

- `PhienBanCauTruc` là `2.2.0`.
- Bảng Người ở `A:I`, vùng mở rộng ở `J:Z`; ba thẻ Tài sản dọc ở `AA:AF`, đặt cạnh nhau và cùng bắt đầu ở hàng 8/9.
- Không còn dữ liệu kỹ thuật ở `J:W` trên `NhapLieu`; ID, ngày tính và trạng thái nằm trên `XuatAn` rất ẩn.
- Tất cả ô nhập tay của Người, Tài sản và Hồ sơ có định dạng Text. Nhập `05`, `00123`, `1/2` rồi lưu/mở lại vẫn giữ nguyên từng ký tự và xuất Word giống hệt.
- Ngày `2015`, `05/2015`, `12/05/2015` hiển thị và xuất đúng nguyên văn; VBA vẫn tạo được ngày tính tương ứng khi chuỗi hợp lệ.
- Cột trường Người mở rộng nhận tiêu đề mới; công thức mẫu được `Đồng bộ trường` điền xuống các dòng, còn dữ liệu nhập tay không bị nhân bản.
- Thêm nhãn ở thẻ Tài sản 1 rồi đồng bộ: thẻ 2 và 3 có cùng nhãn/công thức, không ghi đè giá trị nhập tay.
- VBA tìm trường theo tên tiêu đề/nhãn, không theo offset cột. Chèn thêm trường trong vùng dự phòng không làm sai Hàng TK, Nhận đất hoặc dữ liệu kỹ thuật.
- Thẻ có dữ liệu sinh `TaiSanID` dạng `TS001`; xóa trắng thẻ thì ID được xóa theo.
- `LoaiSo`, `HinhThucSuDung`, `LoaiDat`, `CoQuanCapSo`, `NguoiUyQuyen` chọn được từ danh mục.
- Chỉ thẻ 1 và thẻ 3 có dữ liệu thì thẻ 3 vẫn dùng token có số cuối `3`, không bị dồn thành `2`.
- Bỏ trống tài sản, diện tích lệch hoặc dữ liệu vượt sức chứa mẫu chỉ hiện trong chẩn đoán; người dùng vẫn xuất được Word.
- Người có `NgayCap` trước 01/07/2024 xuất `Căn cước công dân`; từ 01/07/2024 xuất `Căn cước`.
- Parser và tất cả mẫu Word chỉ dùng `{{...}}`; không còn parser tương thích cú pháp cũ.
- `{{ten1}}`, `{{namsinh1}}`, `{{cccd1}}`, `{{loaiso1}}`, `{{sothua1}}`, `{{dientich1}}` lấy đúng giá trị text tương ứng.
- Trường mở rộng `Số điện thoại` tạo token `{{sodienthoai1}}` mà không sửa VBA.
- Ô công thức lỗi xuất nguyên mã lỗi; placeholder không có nguồn và tên bị trùng xuất dấu lỗi rõ theo §12.2, không chặn tạo file.
- Chỉ lỗi mở mẫu, tạo Word hoặc ghi file mới dừng xuất.
- Xuất trong khi người dùng đang mở một tài liệu Word khác: tài liệu đó không bị đóng và không bị thay đổi.
- Xuất xong không tự mở file, không tự mở thư mục, không tự lưu workbook.
- Giao diện Word chỉ có ba hành động chọn: `Chọn nơi lấy mẫu`, `Chọn văn bản`, `Chọn nơi xuất`; `C5` chỉ hiển thị và bị khóa, không có UserForm/wizard hoặc chức năng Word thừa.
- Ba hộp chọn đều giữ nguyên cấu hình khi người dùng bấm hủy; chọn văn bản chỉ nhận `.docx` trong `ThuMucMauWord`.
- Đổi thư mục mẫu cập nhật đúng đường dẫn của mẫu cùng tên hoặc xóa lựa chọn/sức chứa theo §4.1; không tự chọn mẫu khác.
- Chọn thư mục xuất rồi xuất thử tạo file Word đúng tại `ThuMucXuat`; có ảnh QA thanh công cụ và bằng chứng build/workbook, Word thật.
