# Nguồn dự án: Excel VBA Mail Merge – Văn bản thỏa thuận phân chia di sản thừa kế

## 1. Mục đích của tài liệu này

Tài liệu này là bản mô tả kỹ thuật gộp từ hai tệp nguồn:

- `Dữ liệu thừa kế.xlsb`: workbook nhập liệu và VBA xuất Word/PDF.
- `1. PCDS .docx`: mẫu **Văn bản thỏa thuận phân chia di sản thừa kế**.

Dùng tài liệu này làm ngữ cảnh khi tiếp tục phát triển hệ thống. Không cần đọc lại toàn bộ hai tệp gốc, trừ khi cần kiểm tra định dạng Word/Excel cụ thể.

---

## 2. Kiến trúc hiện tại

```text
Người dùng nhập dữ liệu trên sheet HĐCN
        ↓
Công thức đưa dữ liệu về hàng nguồn tblDanhSach
        ↓
VBA mở UserForm frmMergeData
        ↓
VBA đọc hàng B6 trong tblDanhSach
        ↓
VBA mở từng mẫu Word được chọn
        ↓
Thay [Tên trường] trong Word bằng giá trị của cột cùng tên
        ↓
Lưu .docx hoặc .pdf
```

Cơ chế hiện tại **không dùng Mail Merge chuẩn của Word**. Đây là cơ chế Find/Replace qua Word COM Automation, nhưng có tác dụng tương tự mail merge.

---

## 3. Workbook `Dữ liệu thừa kế.xlsb`

### 3.1. Các sheet

| Sheet | Vai trò | Ghi chú |
|---|---|---|
| `HĐCN` | Sheet làm việc chính | Nhập người, tài sản, dữ liệu hồ sơ; chứa bảng nguồn `tblDanhSach`. |
| `Tables` | Dữ liệu phụ trợ | Danh mục loại căn cước, loại sổ, loại đất, hình thức sử dụng, người ủy quyền, cơ quan cấp. |
| `Tờ bản đồ 2025` | Tra cứu thủ công | Danh sách điều chỉnh số tờ bản đồ sau sáp nhập. Chưa được VBA gọi tự động. |
| `Sáp nhập thôn` | Tra cứu thủ công | Danh sách sáp nhập thôn/xóm tại huyện Ý Yên. Chưa được VBA gọi tự động. |

### 3.2. Bảng nguồn Mail Merge

- Tên bảng Excel: `tblDanhSach`
- Phạm vi hiện tại: `A5:HY7`
- Có **233 cột**.
- Hàng 5: tên trường xuất Word.
- Hàng 6: hàng dữ liệu chính được công thức tổng hợp từ khu vực nhập liệu.
- Hàng 7: là hàng dữ liệu trống trong table, hiện không được macro mặc định dùng.
- Form VBA đang cố định vùng đọc là `B6`; vì vậy thực tế macro chỉ xuất từ **hàng 6**.

### 3.3. Các khu vực trên sheet `HĐCN`

| Khu vực | Phạm vi thực tế | Vai trò |
|---|---|---|
| Danh sách placeholder tham chiếu | Hàng 1 | Lưu tên placeholder Word theo dạng `[Tên 1]`, `[CCCD 1]`... |
| Nguồn xuất Word | Hàng 5–7 | `tblDanhSach`; hàng 5 là header, hàng 6 là dữ liệu merge. |
| Tiêu đề khu vực người | Hàng 9 | Header của vùng nhập người. |
| Nhập người | Hàng 10–29 | 20 dòng người, STT ở cột B. |
| Nhập dữ liệu chung/tài sản | Hàng 31–53 | Cột C/D/E tương ứng tài sản 1/2/3; một số dòng là dữ liệu chung hồ sơ. |

> Người dùng mô tả vùng người là dòng 9–29; thực tế dòng 9 là tiêu đề, dữ liệu bắt đầu từ dòng 10.

---

## 4. Dữ liệu người trên sheet `HĐCN`

### 4.1. Các cột nhập trực tiếp

| Cột | Trường | Ghi chú |
|---|---|---|
| `B` | STT | 1 đến 20. |
| `C` | Tên | Họ và tên. Công thức nguồn dùng `PROPER()` khi đưa ra bảng merge. |
| `D` | Năm sinh | Chấp nhận năm hoặc ngày tháng năm; công thức phụ chuẩn hóa sang ngày tại cột Q. |
| `E` | Năm chết | Chấp nhận năm hoặc ngày tháng năm; công thức phụ chuẩn hóa sang ngày tại cột R. |
| `F` | CCCD | Số giấy tờ. |
| `G` | Ngày cấp | Dùng để suy ra loại giấy tờ. |
| `H` | Địa chỉ | Địa chỉ cư trú/thường trú hiển thị trong văn bản. |

### 4.2. Các cột công thức/phụ trợ

| Cột | Mục đích | Logic hiện có |
|---|---|---|
| `I` | Loại CC | Nếu ngày cấp trước 01/07/2024: `Căn cước công dân`; từ ngày này trở đi: `Căn cước`. |
| `J` | Nơi cấp CC | Tra từ sheet `Tables` theo loại giấy tờ. |
| `K` | Nhãn địa chỉ | `Thường trú tại` hoặc `Cư trú tại` theo loại giấy tờ. |
| `L` | Tình trạng chết so với người 1 | So sánh ngày chết với người số 1. |
| `M` | Tình trạng chết so với người 2 | So sánh ngày chết với người số 2. |
| `N` | Chênh lệch tuổi với người 2 | `(Ngày sinh người đó - ngày sinh người 2) / 365`. |
| `O` | Chênh lệch tuổi với người 1 | `(Ngày sinh người đó - ngày sinh người 1) / 365`. |
| `P` | Tên file phụ | Lấy từ cuối cùng của họ tên, chuyển chữ thường. Ví dụ dùng để tạo tên file. |
| `Q` | Ngày sinh chuẩn hóa | Nếu chỉ nhập năm thì đổi ngầm thành `01/01/năm`. |
| `R` | Ngày chết chuẩn hóa | Nếu chỉ nhập năm thì đổi ngầm thành `01/01/năm`. |

### 4.3. Mapping người → bảng nguồn hiện tại

| Người nhập | Dòng nhập | Trường merge đang được tạo |
|---|---:|---|
| Người 1 | 10 | `Tên 1`, `Năm sinh 1`, `Năm chết`, `CCCD 1`, `Ngày cấp 1`, `Địa chỉ 1`, `Loại CC 1`, `Nơi cấp CC 1`, `Thường trú 1`. |
| Người 2 | 11 | `Tên 2`, `Năm sinh 2`, `Năm chết 2`, `CCCD 2`, `Ngày cấp 2`, `Địa chỉ 2`, `Loại CC 2`, `Nơi cấp CC 2`, `Thường trú 2`. |
| Người 3–10 | 12–19 | Mỗi người có: tên, năm sinh, CCCD, ngày cấp, địa chỉ, loại CC, nơi cấp CC, nhãn thường trú/cư trú. |
| Người 11–20 | 20–29 | Có header trong `tblDanhSach`, nhưng **hàng 6 hiện không có công thức lấy dữ liệu từ các dòng này**. |

### 4.4. Hạn chế quan trọng

- Mẫu Word hiện chỉ dùng người **1 đến 10**.
- Công thức hàng 6 hiện chỉ map đúng người **1 đến 10**.
- Người 11–20 không thể xuất ra mẫu Word hiện tại, dù vùng nhập liệu có 20 dòng.
- Cấu trúc field cho người 11–20 không đồng nhất; người 20 còn thiếu `Nơi cấp CC 20` và `Thường trú 20` trong header nguồn.

---

## 5. Dữ liệu tài sản và hồ sơ

### 5.1. Dữ liệu chung

| Ô | Ý nghĩa | Field Word liên quan |
|---|---|---|
| `C31` | Xã/địa bàn niêm yết | `[Niêm Yết]` |
| `C32` | Chọn người ủy quyền | `[Người ủy quyền]`, `[Người ủy quyền2]` qua VLOOKUP sheet `Tables` |
| `C51` | Số công chứng | `[Số công chứng]` |
| `C52` | Giá chuyển nhượng | Chưa thấy dùng trong mẫu Word này. |
| `C53` | Số điện thoại | Chưa thấy dùng trong mẫu Word này. |

### 5.2. Tài sản 1, 2, 3

- Cột `C`: tài sản 1.
- Cột `D`: tài sản 2.
- Cột `E`: tài sản 3.

| Dòng | Nội dung | Field tài sản 1 |
|---:|---|---|
| 33 | Loại sổ | `[Loại sổ]` |
| 34 | Serial | `[Serial]` |
| 35 | Số vào sổ | `[Số vào sổ]` |
| 36 | Số thửa | `[Số thửa]` |
| 37 | Số tờ bản đồ | `[Số tờ]` |
| 38 | Địa chỉ đất | `[Địa chỉ đất]` |
| 39 | Diện tích | `[Diện tích]` |
| 40 | Hình thức sử dụng | `[Hình thức sử dụng]` |
| 41 | Mục đích sử dụng / loại đất | `[Loại đất]` |
| 42 | Thời hạn sử dụng | `[Thời hạn 1]` |
| 43 | ONT | `[ONT]` |
| 44 | CLN | `[CLN]` |
| 45 | NTS | `[NTS]` |
| 46 | LUC | `[LUC]` |
| 47 | Nguồn gốc | `[Nguồn gốc]` |
| 48 | Ngày cấp | `[Ngày cấp sổ]` |
| 49 | Cơ quan cấp | `[Cơ quan cấp sổ]` |

### 5.3. Field tài sản trong bảng nguồn

- Tài sản 1: field không có hậu tố, ví dụ `Serial`, `Số thửa`, `Diện tích`.
- Tài sản 2: field có hậu tố `2`, ví dụ `Serial 2`, `Số thửa 2`, `Diện tích 2`.
- Tài sản 3: field có hậu tố `3`, ví dụ `Serial 3`, `Số thửa 3`, `Diện tích 3`.

Mẫu Word hiện tại chỉ dùng **tài sản 1**.

---

## 6. Sheet `Tables`

Các danh mục đang được công thức/VLOOKUP sử dụng:

| Vùng | Nội dung |
|---|---|
| `D3:F4` | Mapping loại giấy tờ → cơ quan cấp → nhãn địa chỉ. |
| `D6:D8` | Loại Giấy chứng nhận. |
| `D11:D15` | Loại đất: đất ở đô thị, đất ở nông thôn, đất trồng cây lâu năm, nuôi trồng thủy sản, đất trồng lúa. |
| `D20:D21` | Hình thức sử dụng: sử dụng riêng/chung. |
| `D23:F30` | Người ủy quyền và hai đoạn thông tin xuất văn bản. |
| `D30:D33` | Cơ quan cấp Giấy chứng nhận. |

Các sheet tra cứu còn lại chưa được VBA/công thức tự động gọi:

- `Tờ bản đồ 2025`: danh sách đổi số tờ bản đồ.
- `Sáp nhập thôn`: danh sách thôn/xóm cũ → mới.
- Hàm VBA `FindLongestMatch(Target, LookupRange)` đã có nhưng chưa thấy được gọi ở macro hiện tại.

---

## 7. Mẫu Word `1. PCDS .docx`

### 7.1. Mục đích mẫu

Mẫu là **Văn bản thỏa thuận phân chia di sản thừa kế**, có thêm phần chứng thực và danh sách hàng thừa kế.

### 7.2. Logic pháp lý/nhân sự cố định của mẫu hiện tại

Mẫu mặc định giả định:

1. Người 1 và người 2 là vợ chồng.
2. Người 1 và người 2 đều đã chết.
3. Người 3 đến người 10 là con của người 1 và người 2.
4. Di sản là quyền sử dụng đất của hai người để lại.
5. Người 3 là người nhận phần di sản/quyền sử dụng đất.
6. Những người còn lại có nội dung tặng cho hoặc đồng ý phân chia theo câu chữ cố định.
7. Danh sách hàng thừa kế phần cuối hiển thị người 2 đến người 10.

Đây là mẫu **tĩnh**. Khi một người không có dữ liệu, macro chỉ thay placeholder bằng rỗng; nó không tự xóa cả đoạn văn, dòng danh sách hoặc dòng bảng tương ứng.

### 7.3. Placeholder Word đang dùng

Mẫu có **99 placeholder khác nhau**, tổng cộng **300 lần xuất hiện**.

#### Nhóm người (1–10)

Mỗi trường có dạng `[Trường n]`, `n` từ 1 đến 10:

```text
[Tên n]
[Năm sinh n]
[CCCD n]
[Ngày cấp n]
[Địa chỉ n]
[Loại CC n]
[Nơi cấp CC n]
[Thường trú n]
```

Riêng người 1 và người 2 còn có:

```text
[Năm chết]
[Năm chết 2]
```

#### Nhóm tài sản/hồ sơ

```text
[Niêm Yết]
[Loại sổ]
[Địa chỉ đất]
[Serial]
[Số vào sổ]
[Số thửa]
[Số tờ]
[Diện tích]
[Hình thức sử dụng]
[Loại đất]
[Thời hạn 1]
[Nguồn gốc]
[Ngày cấp sổ]
[Cơ quan cấp sổ]
[ONT]
[CLN]
[NTS]
```

### 7.4. Vị trí logic của placeholder trong Word

| Phần văn bản | Placeholder chính |
|---|---|
| Mở đầu và danh sách thừa kế | Người 1–10, `[Niêm Yết]`. |
| Thông tin khai tử | Người 1, 2; năm chết, số/ ngày cấp dùng chung field CCCD/ngày cấp hiện tại. |
| Mô tả di sản | Tài sản 1. |
| Thỏa thuận phân chia | Chủ yếu người 2–10 và người 3 là người nhận. |
| Chứng thực | Người 2–10, `[Niêm Yết]`. |
| Danh sách hàng thừa kế cuối mẫu | Người 2–10, năm sinh và địa chỉ. |

### 7.5. Lưu ý về placeholder

- Placeholder phải giữ nguyên đúng cú pháp ngoặc vuông, ví dụ `[Tên 3]`.
- Tên trong ngoặc vuông phải trùng chính xác header của `tblDanhSach`.
- Không nên tách một placeholder thành nhiều đoạn định dạng/runs trong Word nếu có thể tránh được; việc Find/Replace có thể không nhận diện khi Word chia nhỏ text theo run.
- Mọi field mới cần thêm theo quy trình: thêm header vào `tblDanhSach` → tạo công thức cho hàng dữ liệu → thêm placeholder cùng tên vào Word.

---

## 8. VBA hiện có

### 8.1. Modules và chức năng

| Module/Form | Chức năng |
|---|---|
| `modFunctions` | Macro mở form, hàm hỗ trợ folder/file, tên range, tạo thư mục, tối ưu Excel. |
| `frmMergeData` | Form chính: chọn folder Word, folder lưu, mẫu, định dạng Word/PDF và thực hiện thay placeholder. |
| `modMsgboxTV` | Hiển thị MsgBox Unicode bằng `MessageBoxW`. |
| `Module1` | Hàm `FindLongestMatch` cho tra cứu chuỗi dài nhất. Hiện chưa được gọi. |
| `Sheet1` | Macro `PasteValuesOnly`. |
| `ThisWorkbook` | Tự lưu workbook khi đóng. |

### 8.2. Macro mở form

```vba
Sub TronDuLieu()
    frmMergeData.Show
End Sub
```

### 8.3. Logic khởi tạo form

`frmMergeData.UserForm_Initialize` đang làm 4 việc quan trọng:

```vba
Set rngSelect = Range("$B$6")
Me.txtChonDuLieu.Value = "$B$6"
Call layDuLieuTron
'Đổ toàn bộ header tblDanhSach vào cboTenCot
Me.cboTenCot.ListIndex = 0
```

Hệ quả:

- Chỉ đọc một record tại hàng 6.
- Cột đầu tiên `Tên file` là lựa chọn mặc định để đặt hậu tố file xuất.
- Người dùng không cần quét vùng dữ liệu khi chạy, vì form đã ép vùng `B6`.

### 8.4. Hàm đọc dữ liệu nguồn

`layDuLieuTron`:

1. Xác định table chứa ô được chọn.
2. Đọc toàn bộ tên cột của `tblDanhSach` vào `arrFields`.
3. Đọc dữ liệu theo `.Text` vào `arrData` để giữ định dạng hiển thị ngày/tháng/số.
4. Lưu NumberFormat và Font từng cột để hỗ trợ thay thế trong Word.

Ý nghĩa: với header `Tên 3`, macro tìm placeholder `[Tên 3]` trong Word và thay bằng giá trị của cột `Tên 3` ở hàng 6.

### 8.5. Logic xuất Word/PDF

`SearchReplaceText` làm theo vòng lặp:

```text
Mỗi mẫu Word được chọn
    Mở template
    Mỗi hàng dữ liệu được chọn
        Tạo document mới từ template
        Mỗi field trong tblDanhSach
            Tìm [Tên field] trong Word
            Thay bằng giá trị cùng field
        Lưu .docx hoặc .pdf
```

Có các xử lý đặc biệt:

- Header bắt đầu bằng `Bảng`: chèn Excel Table từ sheet `Tables`.
- Nội dung chứa xuống dòng: đổi sang newline của Word.
- Chuỗi dài hơn 254 ký tự: dùng `TypeText` thay vì `Replacement.Text`.
- Giá trị là đường dẫn `.jpg`, `.png`, `.bmp`: chèn ảnh vào Word.
- Tên file đầu ra: `TênMẫu_GiáTrịCộtĐượcChọn.docx` hoặc `.pdf`.
- Có xử lý trùng tên file bằng hậu tố `_1`, `_2`...

### 8.6. Kiểu lưu đường dẫn

Form lưu folder mẫu và folder xuất vào Name của workbook:

```vba
DuongDanFolderVB
DuongDanFolderLuu
```

---

## 9. Các lỗi/rủi ro kỹ thuật đã thấy

| Mức độ | Nội dung | Hướng xử lý khi nâng cấp |
|---|---|---|
| Cao | Hệ thống nhập 20 người nhưng thực tế chỉ merge được người 1–10. | Làm lại mapping động hoặc mở rộng công thức hàng nguồn và Word theo danh sách biến thiên. |
| Cao | Mẫu Word có nội dung/đánh số cố định cho 10 người. Bỏ trống người sẽ để lại dòng văn bản thừa. | Sinh block người thừa kế bằng code thay vì 10 block cứng. |
| Cao | Quan hệ đang được giả định cố định: 1 chồng, 2 vợ, 3–10 con, người 3 nhận tài sản. | Tạo trường quan hệ/role và sinh câu văn theo role. |
| Trung bình | Property 2: field `Loại đất 2` đang lấy từ `D40` (hình thức sử dụng), khả năng đúng phải là `D41`. | Sửa công thức `GY6` thành tham chiếu dòng 41. |
| Trung bình | `Diện tích chữ`, `Ngày chữ`, `Tháng chữ` đang hiển thị `#NAME?` trong bản kiểm tra đã chuyển đổi. | Kiểm tra công thức/hàm đọc số bằng chữ trong Excel gốc; đóng gói hàm VBA hoặc thay bằng hàm mới. |
| Trung bình | `Diện tích chữ 2` và `Diện tích chữ 3` chưa có công thức. | Bổ sung công thức/hàm đọc số cho cả 3 tài sản. |
| Trung bình | `Tables!E17:F18` có `#REF!`. | Sửa tham chiếu cho danh mục người từ chối nếu còn sử dụng. |
| Trung bình | Macro dựa vào `Selection.Find` trong Word. | Khi nâng cấp nên thay bằng xử lý Range/StoryRange, có thể kiểm tra cả header, footer, textbox. |
| Thấp | Tên file hiện lấy từ từ cuối của tên người 1. | Nên dùng số hồ sơ, tên hồ sơ hoặc mã hồ sơ có kiểm soát ký tự. |
| Thấp | `Workbook_BeforeClose` tự lưu file khi đóng. | Cân nhắc bỏ/tạo tùy chọn để tránh ghi đè ngoài ý muốn. |

---

## 10. Hướng phát triển nên ưu tiên

### Giai đoạn 1: Giữ tương thích Excel/Word hiện tại

1. Sửa toàn bộ công thức lỗi và mapping tài sản 2/3.
2. Chuẩn hóa header/placeholder, không có field thừa hoặc không map.
3. Cho phép chọn record nguồn thay vì ép `B6`, nếu muốn xuất nhiều hồ sơ.
4. Thêm kiểm tra trước khi xuất: placeholder nào trong Word chưa có dữ liệu thì báo danh sách.
5. Dùng mã hồ sơ làm tên file đầu ra.

### Giai đoạn 2: Chuyển từ placeholder theo số thứ tự sang dữ liệu theo vai trò

Không nên phát triển tiếp theo hướng `[Tên 1]`, `[Tên 2]`, `[Tên 3]` cho mọi tình huống. Nên dùng mô hình:

```text
Hồ sơ
├── Người
│   ├── họ tên
│   ├── giới tính
│   ├── ngày sinh
│   ├── ngày chết
│   ├── số giấy tờ
│   ├── ngày cấp
│   ├── địa chỉ
│   └── quan hệ/role
└── Tài sản
    ├── serial
    ├── số vào sổ
    ├── số thửa
    ├── số tờ
    ├── địa chỉ đất
    ├── diện tích và loại đất
    └── thông tin cấp giấy
```

Các role cần hỗ trợ tối thiểu cho hồ sơ thừa kế:

```text
nguoi_de_lai_di_san
vo_chong
con_de
con_nuoi
cha_me
chau_the_vi
nguoi_tu_choi_nhan_di_san
nguoi_nhan_di_san
nguoi_dai_dien
```

Sau đó template nên có block động, ví dụ:

```text
{{#nguoi_thua_ke}}
{{stt}}. {{xung_ho}}: {{ho_ten}}; sinh ngày {{ngay_sinh}};
{{loai_giay_to}} số {{so_giay_to}} do {{noi_cap}} cấp ngày {{ngay_cap}};
{{nhan_dia_chi}}: {{dia_chi}}.
Là {{quan_he_mo_ta}}.
{{/nguoi_thua_ke}}
```

Điều này giải quyết đồng thời các vấn đề: hơn 10 người, cháu thừa kế thế vị, nhiều quan hệ, tự đánh số, tự xóa dòng trống và thay đổi người nhận di sản.

### Giai đoạn 3: Chuyển sang ứng dụng nội bộ

Đích kiến trúc phù hợp:

```text
Ảnh/nhập tay → chuẩn hóa dữ liệu → cơ sở dữ liệu → chọn mẫu → sinh Word/PDF
```

Excel VBA hiện tại nên được xem là **bản mô tả nghiệp vụ và nguồn dữ liệu khởi đầu**, không nên là lõi dài hạn của hệ thống.

---

## 11. Quy tắc khi sửa mẫu Word hoặc Excel

1. Không đổi tên một header `tblDanhSach` nếu chưa kiểm tra placeholder tương ứng trong Word.
2. Không đổi placeholder Word nếu chưa tạo field cùng tên ở hàng nguồn Excel.
3. Nếu thêm người thứ 11 trở đi theo cách cũ, phải làm đủ 3 việc:
   - thêm field trong table;
   - tạo công thức hàng 6;
   - thêm block nội dung/placeholder vào Word.
4. Với văn bản thừa kế, ưu tiên sinh danh sách động thay vì copy thêm block thủ công.
5. Không dùng dữ liệu mẫu có thật để kiểm thử phiên bản phát hành; nên dùng dữ liệu giả lập.

---

## 12. Tóm tắt ngắn cho AI tiếp nhận dự án

```text
- File Excel .xlsb dùng VBA Word COM Automation, không phải Word Mail Merge chuẩn.
- Nút/macro chính: TronDuLieu → frmMergeData.
- Data source: tblDanhSach, header row 5, record mặc định row 6.
- Word placeholder có dạng [Tên field], khớp chính xác header Excel.
- Excel có khu nhập 20 người nhưng row nguồn và mẫu Word mới hỗ trợ thực tế 10 người.
- Mẫu PCDS là mẫu thừa kế tĩnh: người 1/2 là vợ chồng đã chết; người 3–10 là con; người 3 nhận di sản.
- Có tối đa 3 tài sản tại cột C/D/E của khu hàng 33–49; Word hiện chỉ dùng tài sản 1.
- Các vấn đề cần chú ý: field người 11–20 chưa map, static Word blocks, #NAME? ở đọc số/chữ ngày, mapping Loại đất 2 sai, Tables có #REF!.
- Hướng phát triển đúng: dữ liệu theo role/quan hệ và sinh block Word động.
```
