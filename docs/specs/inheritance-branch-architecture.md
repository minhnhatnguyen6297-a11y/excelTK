# Đặc tả hồ sơ thừa kế — SOT tổng

> Trạng thái: **§1–§9 đã chốt và đã triển khai trong workbook MVP v0.2.1. §10–§14 đã chốt về đặc tả, chưa triển khai (đích v0.3.0).**
> Ngày cập nhật: 31/08/2026.
> Phiên bản cấu trúc dữ liệu: hiện hành `2.0.0`; đích của đặc tả này là `2.1.0`.
> Phạm vi: toàn bộ luồng một hồ sơ thừa kế — nhập người, nhập tài sản, nhập thông tin phụ, kiểm tra, xuất Word.

Đây là nguồn quyết định chính (SOT) duy nhất cho hồ sơ thừa kế. Tài liệu được sắp theo đúng thứ tự thao tác của người dùng, không theo lớp kỹ thuật.

Nội dung chờ quyết định, chưa thuộc SOT: [`noi-dung-ngoai-sot.md`](noi-dung-ngoai-sot.md). Không tự đưa nội dung từ file đó vào triển khai.

## Mục lục

| Mục | Nội dung | Trạng thái |
| --- | --- | --- |
| §1 | Mục tiêu và mô hình một-file-một-hồ-sơ | Đã triển khai |
| §2 | Quyết định kiến trúc đã chốt | Đã triển khai |
| §3 | Luồng sử dụng đầy đủ | §3.1–3.2 đã triển khai; §3.3–3.5 đích v0.3.0 |
| §4 | Bố cục sheet `NhapLieu` | Vùng người đã triển khai; tài sản/phụ đích v0.3.0 |
| §5 | Bước 1 — Nhập người | Đã triển khai |
| §6 | Quy ước ngày tháng | Đã triển khai |
| §7 | Hàng TK, nhánh và màu | Đã triển khai |
| §8 | Người nhận đất, từ chối, trạng thái | Đã triển khai |
| §9 | Bảo vệ ô và tốc độ nhập | Đã triển khai |
| §10 | Bước 2 — Nhập tài sản theo phiếu dọc | Đích v0.3.0 |
| §11 | Bước 3 — Thông tin phụ của hồ sơ | Đích v0.3.0 |
| §12 | Bước 4 — Kiểm tra trước khi xuất | Phần người đã triển khai; tài sản/sức chứa đích v0.3.0 |
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

- **Mỗi tài sản là một phiếu dọc độc lập**, nhãn ở bên trái và giá trị ở bên phải. Không dùng bảng ngang một-dòng-một-tài-sản.
  Lý do: một thửa đất có 18 trường; xếp ngang tạo bảng rất rộng, phải cuộn ngang để đọc và dễ nhập lệch cột.
- **Phiếu tài sản nằm ngay dưới vùng người, trên cùng sheet `NhapLieu`.** Không tách sang sheet riêng, để người dùng nhập một mạch từ trên xuống dưới trong một màn hình làm việc.
- Số phiếu cố định là **3**, khớp quy ước hậu tố tài sản của mẫu Word (`không hậu tố` / ` 2` / ` 3`). Phiếu không dùng thì để trống.
- Người dùng nhập trực tiếp vào phiếu. VBA đọc phiếu và dựng bảng ẩn `tblTaiSan` trên `XuatAn` để phục vụ xuất; người dùng không nhập vào bảng ẩn đó.

### 2.3. Thông tin phụ

- Thông tin phụ của hồ sơ nằm ở khối cuối sheet `NhapLieu`, sau vùng tài sản.
- MVP chỉ gồm những trường mẫu `1. PCDS .docx` thật sự dùng: niêm yết, số công chứng, người ủy quyền, người ủy quyền 2.
- Trường workbook cũ có mà mẫu chưa dùng (giá chuyển nhượng, số điện thoại) **không** đưa vào MVP; thêm khi có mẫu cần.

### 2.4. Không quay lại

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

9. Cuộn xuống khối `THÔNG TIN TÀI SẢN`.
10. Điền phiếu `Tài sản 1`. Loại sổ, hình thức sử dụng, loại đất, cơ quan cấp chọn từ danh sách.
11. Nếu hồ sơ có thửa thứ hai hoặc thứ ba, điền tiếp phiếu `Tài sản 2`, `Tài sản 3`. Phiếu không dùng để trống.

### 3.4. Bước 3 — Thông tin phụ

12. Cuộn xuống khối `THÔNG TIN HỒ SƠ`, điền xã niêm yết, số công chứng, người ủy quyền.
13. Chọn mẫu Word ở ô `Mẫu Word`.

### 3.5. Bước 4–5 — Kiểm tra và xuất

14. Bấm `Kiểm tra dữ liệu`. Nếu sheet `KiemTra` báo lỗi chặn xuất, sửa theo STT và họ tên được chỉ ra.
15. Bấm `Xuất Văn bản`. File Word mới được tạo trong thư mục kết quả; workbook không tự mở file và không tự mở thư mục.

## 4. Bố cục sheet `NhapLieu`

Một sheet, ba khối, đọc từ trên xuống theo đúng thứ tự thao tác.

```text
hàng 1–4    THANH CÔNG CỤ
hàng 6–7    nhãn vùng người
hàng 8      tiêu đề bảng người
hàng 9–38   30 dòng người
hàng 40     nhãn THÔNG TIN TÀI SẢN
hàng 41–60  phiếu Tài sản 1
hàng 61–80  phiếu Tài sản 2
hàng 81–100 phiếu Tài sản 3
hàng 102    nhãn THÔNG TIN HỒ SƠ
hàng 103–108 các trường hồ sơ
```

### 4.1. Thanh công cụ

```text
Hàng TK tối đa  [ 2 ]        [ Kiểm tra dữ liệu ]   [ Xuất Văn bản ]
Mẫu Word        [ 1. PCDS .docx                                   ]
```

- `Hàng TK tối đa` mặc định `2`, tối thiểu `0`, tối đa `4`.
- `Xuất Văn bản` là hành động chính và đặt bên phải, đúng trong vùng `H3:I4`.
- Nếu giảm mức tối đa thấp hơn Hàng TK đang có dữ liệu, không tự sửa dữ liệu; chặn và chỉ rõ các dòng cần xử lý.
- Nếu tăng mức tối đa, bảng màu và vòng chuyển Hàng TK cập nhật ngay.
- Ô `Mẫu Word` giữ đường dẫn mẫu đang chọn. Đổi mẫu làm cập nhật sức chứa người và sức chứa tài sản dùng cho kiểm tra (§12.3).

### 4.2. Nhãn bắt buộc trên giao diện

- `VÙNG NHẬP DỮ LIỆU` trên các cột người dùng được gõ.
- `BẤM ĐỂ CHỌN` trên các cột hành động.
- `THÔNG TIN TÀI SẢN` trên khối tài sản, kèm dòng phụ `Phiếu không dùng thì để trống`.
- `THÔNG TIN HỒ SƠ` trên khối cuối.

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
| `LoaiCC` | Ẩn | Loại giấy tờ suy từ `NgayCapTinh`, xem §14.2 |
| `NoiCapCC` | Ẩn | Nơi cấp tra theo `LoaiCC`, xem §14.2 |
| `NhanDiaChi` | Ẩn | Nhãn địa chỉ suy theo `LoaiCC`, xem §14.2 |

Ba cột `LoaiCC`, `NoiCapCC`, `NhanDiaChi` là phần thêm của `2.1.0`. Chúng do VBA ghi, người dùng không gõ.

Không lưu trường nhập tay `TuChoi`. Đây là kết quả được engine tính để tránh các trạng thái mâu thuẫn.

### 5.3. Cấu hình hồ sơ

| Trường | Giá trị |
| --- | --- |
| `HangTKToiDa` | Mặc định `2`, hợp lệ từ `0` đến `4` |
| `BangMauHangTK` | Mã bảng màu đang chọn |
| `PhienBanCauTruc` | `2.1.0` |
| `NguoiIDTiepTheo` | Số dùng để sinh ID người tiếp theo |
| `TaiSanIDTiepTheo` | Số dùng để sinh ID tài sản tiếp theo |
| `SucChuaNguoi` | Sức chứa hiện tại của bảng nhập |
| `MauWordDangChon` | Đường dẫn mẫu Word đang chọn |
| `SucChuaNguoiMau` | Số người tối đa mẫu đang chọn chứa được |
| `SucChuaTaiSanMau` | Số tài sản tối đa mẫu đang chọn chứa được |

## 6. Quy ước ngày tháng

Áp dụng cho mọi ô ngày trong workbook: ngày sinh, ngày chết, ngày cấp giấy tờ của người và ngày cấp sổ của tài sản.

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
- Ngày không tồn tại, ví dụ `31/02/1990`, là lỗi chặn xuất.
- Ngày chết trước ngày sinh là lỗi chặn xuất.
- Dữ liệu được dán từ Excel có thể là text, số ngày của Excel hoặc ngày thật. Nếu còn là text thì phải giữ nguyên text; engine đồng thời chuyển một bản ẩn thành ngày thật để tính.
- Với dữ liệu người dùng gõ dạng text, trường `...Goc` phải giữ đúng chuỗi đã nhập, ví dụ `1950` hoặc `06/1950`.
- Nếu dữ liệu dán tới chỉ còn số ngày nội bộ của Excel, định dạng chữ ban đầu không thể khôi phục chắc chắn; khi đó trường gốc lưu ngày chuẩn `dd/mm/yyyy` để tránh đoán sai.
- Các phép tính tuổi, so sánh ngày và xác định đã chết chỉ dùng trường `...Tinh`, không tính trực tiếp từ text đang hiển thị.
- Dữ liệu xuất trung gian phải có cả text dùng để hiển thị/Word và ngày thật dùng để tính. Word mặc định lấy text, không lấy ngày đã bổ sung trong cột ẩn.

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
| `C5` | Chọn — Mẫu Word |
| `B9:G38` | Nhập — họ tên, ngày sinh, ngày chết, số giấy tờ, ngày cấp, địa chỉ |
| `H9:I38` | Chỉ bấm — Hàng TK, Nhận đất |
| Ô giá trị của phiếu tài sản | Nhập hoặc chọn, xem §10.3 |
| Ô giá trị của khối hồ sơ | Nhập hoặc chọn, xem §11 |
| Cột `A`, cột `J:W` | Hệ thống, luôn khóa và ẩn |

- Các sheet `CauHinh`, `KiemTra`, `XuatAn` được bảo vệ; `CauHinh` và `KiemTra` ẩn trong sử dụng bình thường, `XuatAn` ở trạng thái rất ẩn.
- Sheet `DanhMuc` ẩn; người dùng chỉ gặp nội dung của nó qua danh sách chọn.
- Mật khẩu bảo vệ chỉ nhằm chống sửa nhầm, không được coi là biện pháp bảo mật dữ liệu.
- VBA được phép cập nhật các ô khóa bằng chế độ `UserInterfaceOnly`, và phải khôi phục chế độ này mỗi lần mở workbook.

### 9.2. Tốc độ nhập liệu

- Khi một ô thông thường thay đổi, VBA chỉ xử lý dòng vừa thay đổi và các dữ liệu thật sự liên quan.
- Không quét lại toàn bộ ngày, trạng thái và nhánh sau mỗi ô họ tên, số giấy tờ hoặc địa chỉ.
- Sửa một ô của phiếu tài sản chỉ xử lý phiếu đó, không quét lại bảng người.
- Chỉ làm mới toàn bộ khi mở workbook, trước khi kiểm tra/xuất Word hoặc khi thao tác Hàng TK làm thay đổi cả nhánh.
- Việc nhập rồi chuyển sang ô tiếp theo không được tạo độ trễ dễ nhận thấy trên máy làm việc thông thường.

---

## 10. Bước 2 — Nhập tài sản theo phiếu dọc

### 10.1. Hình dạng phiếu

Mỗi tài sản là một phiếu dọc: nhãn ở cột `B`, giá trị ở cột `C` (gộp `C:G` cho các trường dài như địa chỉ và nguồn gốc). Phiếu đọc từ trên xuống, không cuộn ngang.

```text
────────────────────────────────────────────────────
 TÀI SẢN 1
────────────────────────────────────────────────────
 Loại sổ             │ Giấy chứng nhận quyền sử dụng đất
 Số phát hành/Serial │ CS 123456
 Số vào sổ           │ CS 00123
 Số thửa             │ 45
 Số tờ bản đồ        │ 12
 Địa chỉ đất         │ Thôn ..., xã ..., tỉnh ...
 Diện tích (m²)      │ 250
 Hình thức sử dụng   │ Sử dụng riêng
 Loại đất            │ Đất ở nông thôn
 Thời hạn sử dụng    │ Lâu dài
 ONT (m²)            │ 120
 CLN (m²)            │ 130
 NTS (m²)            │
 LUC (m²)            │
 Nguồn gốc           │ Nhà nước công nhận quyền sử dụng đất
 Ngày cấp sổ         │ 12/05/2015
 Cơ quan cấp sổ      │ UBND huyện ...
 Ghi chú             │
────────────────────────────────────────────────────
```

Ba phiếu có hình dạng giống nhau, nhãn `TÀI SẢN 1`, `TÀI SẢN 2`, `TÀI SẢN 3`. Phiếu 2 và 3 để trống khi hồ sơ chỉ có một thửa.

### 10.2. Trường của một phiếu

| Trường | Nhập kiểu | Placeholder Word tài sản 1 |
| --- | --- | --- |
| `LoaiSo` | Chọn từ danh mục | `[Loại sổ]` |
| `Serial` | Text | `[Serial]` |
| `SoVaoSo` | Text | `[Số vào sổ]` |
| `SoThua` | Text | `[Số thửa]` |
| `SoTo` | Text | `[Số tờ]` |
| `DiaChiDat` | Text dài | `[Địa chỉ đất]` |
| `DienTich` | Số | `[Diện tích]` |
| `HinhThucSuDung` | Chọn từ danh mục | `[Hình thức sử dụng]` |
| `LoaiDat` | Chọn từ danh mục | `[Loại đất]` |
| `ThoiHan` | Text | `[Thời hạn 1]` |
| `ONT` | Số | `[ONT]` |
| `CLN` | Số | `[CLN]` |
| `NTS` | Số | `[NTS]` |
| `LUC` | Số | `[LUC]` |
| `NguonGoc` | Text dài | `[Nguồn gốc]` |
| `NgayCapSo` | Ngày theo §6 | `[Ngày cấp sổ]` |
| `CoQuanCapSo` | Chọn từ danh mục | `[Cơ quan cấp sổ]` |
| `GhiChu` | Text dài | không xuất |

Hậu tố Word: tài sản 1 không hậu tố, tài sản 2 hậu tố ` 2`, tài sản 3 hậu tố ` 3`. Ví dụ `[Serial 2]`, `[Diện tích 3]`.

### 10.3. Dữ liệu ẩn của phiếu

Mỗi phiếu có ô ẩn trong vùng cột `J:W` cùng hàng với phiếu:

| Trường ẩn | Ý nghĩa |
| --- | --- |
| `TaiSanID` | ID bất biến do hệ thống sinh, dạng `TS001`; sinh khi phiếu bắt đầu có dữ liệu |
| `NgayCapSoGoc` | Chuỗi người dùng đã nhập |
| `NgayCapSoTinh` | Ngày thật của Excel dùng để so sánh |
| `CoDuLieu` | Phiếu có dữ liệu hay không, xem §10.4 |

`TaiSanID` không hiển thị và người dùng không cần nhớ. Lý do cần ID ổn định: phiếu có thể bị xóa và nhập lại, và bảng xuất ẩn phải tham chiếu đúng phiếu.

Nhãn `Tài sản 1`, `Tài sản 2`, `Tài sản 3` chỉ là số thứ tự hiển thị và số thứ tự hậu tố Word; không phải khóa liên kết.

### 10.4. Phiếu có dữ liệu

Một phiếu được coi là **có dữ liệu** khi bất kỳ trường nào trong bốn trường sau khác rỗng: `LoaiSo`, `Serial`, `SoThua`, `DiaChiDat`.

Phiếu không có dữ liệu bị bỏ qua hoàn toàn: không sinh `TaiSanID`, không đưa vào bảng xuất, không bị kiểm tra trường bắt buộc.

Phiếu chỉ có `GhiChu` hoặc chỉ có `DienTich` không được coi là có dữ liệu — đó là dấu hiệu nhập dở, và §12.2 sinh cảnh báo cho trường hợp này.

### 10.5. Bảng xuất ẩn `tblTaiSan`

VBA đọc ba phiếu rồi dựng `tblTaiSan` trên sheet `XuatAn`:

- Một dòng cho mỗi phiếu có dữ liệu, theo thứ tự phiếu 1 → 2 → 3.
- Cột `STTTaiSan` được đánh lại liên tục từ 1, không bỏ số. Nếu chỉ phiếu 1 và phiếu 3 có dữ liệu thì phiếu 3 nhận `STTTaiSan = 2` và **xuất bằng hậu tố ` 2`**.
- Cột `TaiSanID` giữ nguyên ID của phiếu.
- Bảng chỉ do VBA tạo/refresh trước khi kiểm tra và trước khi xuất. Người dùng không nhập vào bảng này.

Hệ quả cần nhớ: hậu tố Word đi theo `STTTaiSan` liên tục, không theo số phiếu trên giao diện. Cách này để mẫu chỉ chứa một tài sản vẫn xuất được khi người dùng vô tình bỏ trống phiếu 1.

### 10.6. Thêm và xóa phiếu

- Không có lệnh thêm phiếu trong MVP. Ba phiếu luôn tồn tại, dùng bằng cách điền vào.
- Xóa một tài sản là xóa trắng các ô giá trị của phiếu đó. VBA phát hiện phiếu trở về trạng thái không có dữ liệu và xóa `TaiSanID`, `NgayCapSoGoc`, `NgayCapSoTinh` của phiếu.
- Nếu về sau cần nhiều hơn ba tài sản, thêm phiếu bằng cách chèn khối 20 hàng theo đúng bố cục §4, đồng thời mẫu Word phải có hậu tố tương ứng. Không chuyển vùng này thành bảng ngang.

## 11. Bước 3 — Thông tin phụ của hồ sơ

Khối cuối sheet `NhapLieu`, nhãn ở cột `B`, giá trị ở cột `C`.

| Trường | Nhập kiểu | Placeholder Word |
| --- | --- | --- |
| `NiemYet` | Text | `[Niêm Yết]` |
| `SoCongChung` | Text | `[Số công chứng]` |
| `NguoiUyQuyen` | Chọn từ danh mục | `[Người ủy quyền]` |
| `NguoiUyQuyen2` | Chọn từ danh mục | `[Người ủy quyền2]` |

Quy tắc:

- `NiemYet` là xã hoặc địa bàn niêm yết văn bản. Bắt buộc khi mẫu đang chọn có `[Niêm Yết]`.
- `SoCongChung` là số công chứng của văn bản. Để trống thì xuất chuỗi rỗng, không xuất `0`.
- `NguoiUyQuyen` và `NguoiUyQuyen2` chọn từ danh mục người ủy quyền (§14.3). Hồ sơ chỉ có một người ủy quyền thì để trống trường thứ hai.
- Không thêm `GiaChuyenNhuong` và `SoDienThoai` vào MVP. Hai trường này có trong workbook cũ nhưng mẫu `1. PCDS .docx` không dùng; thêm khi có mẫu thật cần đến, cùng lúc bổ sung placeholder.

---

## 12. Bước 4 — Kiểm tra trước khi xuất

Kiểm tra chạy khi bấm `Kiểm tra dữ liệu` và tự chạy lại trước khi xuất. Kết quả ghi ra sheet `KiemTra`. Mọi thông báo phải ghi rõ STT và họ tên hoặc số phiếu tài sản; khi có thể, cho phép bấm để nhảy tới ô lỗi.

### 12.1. Lỗi chặn xuất

Người và nhánh:

- `HangTKToiDa` nằm ngoài `0..4`.
- Có `HangTK` lớn hơn mức tối đa.
- Có Hàng TK bị nhảy cấp.
- H0/H1 có `ParentNguoiID`, hoặc dòng từ H2 trở đi thiếu cha, tự trỏ, trỏ sai người hay không ở Hàng TK liền trước.
- Có chu kỳ trong chuỗi `ParentNguoiID`.
- Người đã chết vẫn có `NhanDat = True`.
- Cùng một `NguoiID` xuất hiện ở nhiều dòng.
- Không có người nào được chọn `NhanDat`.

Ngày tháng, áp dụng cho cả ngày của người và ngày cấp sổ:

- Ngày sai định dạng hoặc ngày không tồn tại.
- Ngày chết trước ngày sinh.

Tài sản:

- Không có phiếu tài sản nào có dữ liệu.
- Phiếu có dữ liệu nhưng thiếu trường bắt buộc: `LoaiSo`, `SoThua`, `DiaChiDat`, `DienTich`, `LoaiDat`.
- `DienTich`, `ONT`, `CLN`, `NTS`, `LUC` không phải số hoặc là số âm.
- Tổng `ONT + CLN + NTS + LUC` lớn hơn `DienTich` của cùng phiếu.

Thông tin phụ và mẫu:

- Chưa chọn mẫu Word, hoặc đường dẫn mẫu không tồn tại.
- Mẫu đang chọn có `[Niêm Yết]` nhưng `NiemYet` để trống.
- Dữ liệu vượt sức chứa của mẫu đang chọn, xem §12.3.
- Sau khi thay, còn placeholder của mẫu chưa được thay.

### 12.2. Cảnh báo

- Dòng chủ đất thứ hai để trống: hiểu là hồ sơ một chủ đất, không coi là lỗi.
- Có hai dòng nghi cùng một người do trùng họ tên và ngày sinh.
- Nhánh có người đã chết nhưng chưa có dòng ở Hàng TK kế tiếp.
- Người từ chối chưa được chia nhóm: dùng `TC_DEFAULT` ở MVP.
- Phiếu tài sản có vài trường lẻ nhưng không đạt điều kiện có dữ liệu ở §10.4 — nghi nhập dở.
- Phiếu 1 để trống trong khi phiếu 2 hoặc 3 có dữ liệu — hậu tố Word sẽ dồn lên, xem §10.5.
- Tổng `ONT + CLN + NTS + LUC` nhỏ hơn `DienTich`: phần diện tích còn lại chưa được phân loại.
- `NguoiUyQuyen` để trống trong khi mẫu có `[Người ủy quyền]`.

### 12.3. Sức chứa mẫu

Mẫu Word là template tĩnh: khi một slot không có dữ liệu, macro chỉ thay bằng rỗng và **không tự xóa** đoạn văn, dòng danh sách hay dòng bảng tương ứng. Ngược lại, dữ liệu vượt số slot của mẫu sẽ bị bỏ âm thầm nếu không kiểm tra.

Vì vậy kiểm tra phải so hai chiều:

```text
số người có dữ liệu   <= SucChuaNguoiMau
số phiếu có dữ liệu   <= SucChuaTaiSanMau
```

- Vượt là lỗi chặn xuất, thông báo nêu rõ sức chứa của mẫu và số lượng hồ sơ đang có.
- `1. PCDS .docx`: `SucChuaNguoiMau = 10`, `SucChuaTaiSanMau = 1`.
- Hai giá trị này gắn với mẫu đang chọn, cập nhật khi đổi ô `Mẫu Word`.
- Thiếu dữ liệu so với sức chứa **không** phải lỗi; các slot dư xuất chuỗi rỗng.

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
| `TaiSan` | `tblTaiSan` theo §10.5, sắp theo `STTTaiSan` |
| `HoSo` | Các trường ở §11 |

Dữ liệu xuất trung gian có các cột ngày hiển thị, ngày gốc, ngày dùng để tính và `TuoiLucChet`. Template Word không được đọc màu ô, vị trí dòng 1–2 hoặc suy luận nhánh từ STT.

### 13.2. Ánh xạ slot người của mẫu cũ

Mẫu hiện có dùng placeholder đánh số `n = 1..10`. Quy tắc gán slot:

```text
slot 1, 2   = chủ đất theo STTNhap; thiếu chủ đất thứ hai thì slot 2 rỗng
slot 3..n   = những người còn lại có dữ liệu, không phải chủ đất, theo STTNhap
```

Mỗi slot cấp các placeholder:

```text
[Tên n] [Năm sinh n] [CCCD n] [Ngày cấp n] [Địa chỉ n]
[Loại CC n] [Nơi cấp CC n] [Thường trú n]
```

Slot 1 và 2 có thêm `[Năm chết]` và `[Năm chết 2]` — đúng tên cũ, không phải `[Năm chết 1]`.

Câu chữ của mẫu cũ giả định người số 3 là người nhận và người 4 trở đi tặng cho. Giả định đó **không còn đúng** vì người nhận được chọn bằng `NhanDat` trên từng dòng. Đoạn thỏa thuận trong mẫu phải viết lại theo nhóm, không theo số slot:

```text
Người nhận quyền sử dụng đất:  {{nguoi_nhan_dat.danh_sach}}
Những người từ chối nhận di sản: {{nhom_tu_choi.TC_DEFAULT.danh_sach}}
```

Bảng nhân thân trong mẫu vẫn dùng slot đánh số như trên.

### 13.3. Chiến lược placeholder kép

Giữ đồng thời hai cú pháp:

- Placeholder cũ dạng ngoặc vuông: `[Tên 4]`, `[CCCD 4]`, `[Serial 2]`. Cần để các mẫu Word hiện có chạy được ngay.
- Namespace mới dạng `{{ }}`: `{{nguoi_nhan_dat.danh_sach}}`, `{{nhom_tu_choi.TC_DEFAULT.danh_sach}}`, `{{cay_nhanh.danh_sach}}`, `{{ho_so.niem_yet}}`, `{{tai_san.2.serial}}`.

Quy tắc parse namespace mới: đoạn cuối là tên trường, các đoạn số ở giữa là định danh đối tượng.

Block lặp động trong Word (danh sách người không giới hạn) để giai đoạn sau. Đến khi đó, số slot của mẫu vẫn là giới hạn cứng và §12.3 vẫn là lớp chặn.

### 13.4. Quy tắc thay thế

- Giá trị trống xuất chuỗi rỗng. Không xuất `0`, không xuất `00/01/1900`.
- Ngày xuất theo text người dùng đã nhập (§6), không dùng cột `...Tinh`.
- Nội dung có xuống dòng phải đổi sang ký tự newline của Word.
- Chuỗi dài hơn 254 ký tự phải dùng `TypeText` thay vì `Replacement.Text`, do giới hạn của Find/Replace.
- Placeholder phải giữ đúng cú pháp; tên trong placeholder phải trùng chính xác tên trường nguồn.
- Không để một placeholder bị tách thành nhiều run trong Word, vì Find/Replace sẽ không nhận diện được.
- Thêm field mới phải làm đủ ba việc: thêm trường nguồn, tạo dữ liệu tương ứng, thêm placeholder cùng tên vào Word.

### 13.5. An toàn phiên Office

- Dùng `CreateObject` để tạo phiên Word **riêng**. Không dùng `GetObject`.
  Lỗi cũ cần tránh: macro dùng `GetObject` bám vào phiên Word đang mở của người dùng rồi gọi `Quit`, làm đóng luôn tài liệu khác của người dùng.
- Chỉ mở và đóng tài liệu do phiên đó tạo; lưu file rồi thoát đúng phiên đó.
- Thay placeholder trong body, bảng, header, footer và textbox bằng Word Range. Không dùng `Selection.Find` vì không phủ hết header/footer/textbox.
- File Word xuất ra là bản sao độc lập của mẫu. Không dùng Mail Merge, không tạo liên kết Word–Excel.
- Sau khi xuất chỉ hiển thị kết quả. Không tự mở file, không tự mở thư mục.
- Không tự lưu workbook khi đóng.
- Tên file đầu ra có kiểm soát ký tự đặc biệt và chống trùng bằng hậu tố `_1`, `_2`.

## 14. Danh mục tra cứu

Sheet `DanhMuc` ẩn, chứa các danh sách phục vụ ô chọn và tra cứu. Người dùng không nhập trực tiếp trong quá trình làm hồ sơ.

### 14.1. Các danh mục

| Danh mục | Dùng cho |
| --- | --- |
| Loại sổ | `LoaiSo` của phiếu tài sản |
| Loại đất | `LoaiDat`; gồm đất ở đô thị, đất ở nông thôn, cây lâu năm, nuôi trồng thủy sản, trồng lúa |
| Hình thức sử dụng | `HinhThucSuDung`; sử dụng riêng, sử dụng chung |
| Cơ quan cấp sổ | `CoQuanCapSo` |
| Loại giấy tờ → nơi cấp → nhãn địa chỉ | §14.2 |
| Người ủy quyền | `NguoiUyQuyen`, `NguoiUyQuyen2` |

### 14.2. Quy tắc loại giấy tờ

Suy từ `NgayCapTinh` của từng người:

```text
NgayCapTinh trống            -> LoaiCC rỗng
NgayCapTinh <  01/07/2024    -> "Căn cước công dân"
NgayCapTinh >= 01/07/2024    -> "Căn cước"
```

`NoiCapCC` tra theo `LoaiCC` trong danh mục. `NhanDiaChi` suy theo `LoaiCC`: `Thường trú tại` hoặc `Cư trú tại`.

Ba giá trị này do VBA ghi vào cột ẩn khi `NgayCap` của dòng đó thay đổi; không quét lại toàn bảng.

### 14.3. Hai bảng tra cứu chưa tự động

`Tờ bản đồ 2025` (điều chỉnh số tờ sau sáp nhập) và `Sáp nhập thôn` (thôn/xóm huyện Ý Yên) là tra cứu thủ công. Chưa nối vào luồng nhập trong phiên bản này.

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
- Không dùng block lặp động trong Word.

## 16. Các hướng cũ không sử dụng

- **Mã đường dẫn `1`, `1.4`, `1.4.4`:** bỏ vì con chung không thuộc riêng nhánh 1 hay 2, đổi cấu trúc phải đánh lại mã cả cây và dễ nhân đôi người.
- **Quan hệ chi tiết `ChaID/MeID/VoChongID`:** không dùng cho MVP vì làm gián đoạn cách nhập hiện tại và mô hình hóa nhiều hơn nhu cầu thực tế.
- **Bảng tài sản ngang một-dòng-một-tài-sản:** bỏ vì 18 trường tạo bảng quá rộng, phải cuộn ngang và dễ nhập lệch cột. Thay bằng phiếu dọc ở §10.
- **Sheet tài sản riêng:** bỏ để người dùng nhập một mạch trên cùng sheet `NhapLieu`.
- **Điểm được giữ:** mỗi người có một `NguoiID`, chỉ nhập một lần; dùng `ParentNguoiID` để nhánh không phụ thuộc màu hay vị trí dòng; mỗi tài sản có một `TaiSanID` ổn định.

Chỉ xem xét lại các hướng trên khi có yêu cầu pháp lý hoặc xuất văn bản mà mô hình `HangTK + ParentNguoiID` không biểu diễn được.

## 17. Điều kiện nghiệm thu

### 17.1. Đã đạt ở v0.2.1

- Workbook được tạo ở file mới, không ghi đè mẫu nguồn.
- Chỉ `B4` và `B9:G38` cho phép gõ; các vùng khác được khóa đúng §9.1.
- Cột trạng thái và các cột kỹ thuật được ẩn.
- Hàng TK dùng H0–H4, điều khiển tối đa dùng 0–4 và mặc định 2.
- Màu H0–H4 đúng bảng màu ở §7.2 và đủ khác nhau khi nhìn trên màn hình.
- Ngày `yyyy`, `mm/yyyy`, `dd/mm/yyyy`, ô trống và dữ liệu dán được xử lý đúng §6.
- Ví dụ sinh `1950`, chết `1999` phải hiển thị và xuất đúng `1950`, `1999`, trong khi engine vẫn tính tuổi `49`.
- Nhập dữ liệu thông thường chỉ xử lý dòng liên quan, không làm mới toàn bộ bảng sau mỗi ô.
- Người đã chết không bị đưa vào nhóm từ chối.
- File Excel và file Word thử vượt kiểm tra cấu trúc, không còn placeholder chưa thay.

### 17.2. Cần đạt ở v0.3.0

- `PhienBanCauTruc` là `2.1.0`.
- Ba phiếu tài sản dọc nằm dưới vùng người trên cùng sheet `NhapLieu`, đúng bố cục §4.
- Phiếu tài sản chỉ cho nhập vào ô giá trị; nhãn và ô ẩn bị khóa.
- Phiếu có dữ liệu sinh `TaiSanID` dạng `TS001`; xóa trắng phiếu thì ID được xóa theo.
- `LoaiSo`, `HinhThucSuDung`, `LoaiDat`, `CoQuanCapSo`, `NguoiUyQuyen` chọn được từ danh mục.
- `NgayCapSo` nhập `2015`, `05/2015`, `12/05/2015` đều hiển thị và xuất đúng nguyên văn.
- Chỉ phiếu 1 và phiếu 3 có dữ liệu thì phiếu 3 xuất bằng hậu tố ` 2`, đúng §10.5.
- Bỏ trống toàn bộ ba phiếu thì kiểm tra chặn xuất.
- Tổng `ONT + CLN + NTS + LUC` vượt `DienTich` thì chặn xuất.
- 11 người với mẫu `1. PCDS .docx` thì chặn xuất và thông báo nêu rõ sức chứa 10.
- Hai phiếu tài sản có dữ liệu với mẫu `1. PCDS .docx` thì chặn xuất và thông báo nêu rõ sức chứa 1.
- Người có `NgayCap` trước 01/07/2024 xuất `Căn cước công dân`; từ 01/07/2024 xuất `Căn cước`.
- Xuất `1. PCDS .docx` với dữ liệu giả: không còn `[Tên n]`, `[Serial]`, `[Niêm Yết]` hay `{{...}}` trong file kết quả.
- Đoạn thỏa thuận nêu đúng người được chọn `NhanDat`, không phụ thuộc người đó ở slot số mấy.
- Xuất trong khi người dùng đang mở một tài liệu Word khác: tài liệu đó không bị đóng và không bị thay đổi.
- Xuất xong không tự mở file, không tự mở thư mục, không tự lưu workbook.
