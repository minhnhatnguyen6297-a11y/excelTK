# Nội dung ngoài SOT — chờ quyết định

> Đây KHÔNG phải spec đã chốt. Đây là danh mục nội dung nằm ngoài SOT, được gom lại để người dùng quyết định cái nào đưa vào spec chính thức tiếp theo.
>
> SOT duy nhất hiện tại là [`inheritance-branch-architecture.md`](inheritance-branch-architecture.md) (chốt 27/08/2026). Không sửa file đó.
>
> File này được tạo ngày 29/08/2026 khi hợp nhất spec, tổng hợp từ 3 file spec đã bị xóa (`inheritance-workflow.md`, `excel-restructure.md`, `approved-minimal-entry-ui.md`) cùng hai file session (`SESSION_MA_NHANH_VAI_TRO_2026-08-26.md`, `SESSION_EXCEL_RESTRUCTURE_2026-08-24.md`) và một file tham chiếu (`TONGQUANDUAN.md`).
>
> Mọi nội dung dựa trên mô hình mã nhánh `1`/`1.4`/`1.4.4`, slot vai trò `.1/.2/.3/.4+`, cột `TrangThaiTN`/`QuanHe`, quan hệ `ChaID/MeID/VoChongID` đã bị SOT khai tử có chủ đích (xem SOT §16) và KHÔNG được đưa vào file này.

## Cập nhật 31/08/2026 — các mục đã được đưa vào SOT

SOT đã được cấu trúc lại thành tài liệu tổng chia mục theo flow. Những mục dưới đây **đã được quyết định** và chuyển vào SOT; phần mô tả cũ trong file này chỉ còn giá trị lịch sử:

| Mục ở file này | Quyết định | Vị trí trong SOT |
| --- | --- | --- |
| §2.2 Vùng thông tin tài sản | Nhận, dạng phiếu dọc, 3 phiếu, đặt dưới vùng người trên `NhapLieu` | SOT §2.2, §10 |
| §2.3 Mã kỹ thuật ẩn | Nhận: `TaiSanID` dạng `TS001`, ẩn | SOT §10.3 |
| §4.2 Xuất Word an toàn | Nhận toàn bộ | SOT §13.5 |
| §4.3 Placeholder kép | Nhận | SOT §13.3 |
| §4.4 Quy tắc placeholder | Nhận | SOT §13.4 |
| §5 Loại giấy tờ, nhãn địa chỉ, hậu tố tài sản | Nhận | SOT §14.2, §10.2 |
| §6 Sức chứa template | Nhận, thành lỗi chặn xuất hai chiều | SOT §12.3 |
| §9 câu 1, 3 | Chốt: một workbook là một hồ sơ, copy file mẫu; không có `HoSoID` | SOT §1 |
| §9 câu 4, 5, 8, 9 | Chốt: hồ sơ hai bên là workbook riêng; workbook thừa kế không có cột `Ben` | SOT §15 |
| §9 câu 7 | Chốt: hiển thị lỗi theo STT và họ tên, bấm để nhảy tới ô lỗi | SOT §12 |

Vẫn **chưa** vào SOT: §1 (khối di sản, sở hữu, phân chia theo tỷ lệ), §2.4 sheet phụ mở rộng, §3 hồ sơ nhiều loại trong cùng file, §4.5 chèn bảng/ảnh/`TypeText` mở rộng, §7 danh sách lỗi workbook cũ, §9 câu 2, 6, 10, 11.

## 1. Di sản, sở hữu và phân chia tài sản

SOT chỉ có checkbox `NhanDat` trên từng dòng người; SOT không có khối di sản, không có tỷ lệ phân chia, không có mô hình đồng chủ thứ 3 trở lên. Toàn bộ mục này SOT im lặng.

### 1.1. `tblDiSan`

| Trường | Ý nghĩa |
| --- | --- |
| `DiSanID` | ID hệ thống, ví dụ `DS001` |
| `HoSoID` | Hồ sơ sở hữu di sản |
| `TenDiSan` | Tên gợi nhớ, ví dụ "Di sản đất tại ..." |
| `ChuSoHuuChinhID` | ID của chủ đất |
| `CheDoSoHuu` | Riêng, chung vợ chồng hoặc chung khác |
| `GhiChu` | Ngoại lệ |

MVP tạo một `DiSanID` mặc định cho hồ sơ. Nếu sau này một hồ sơ có nhiều khối di sản độc lập, người dùng tạo thêm dòng thay vì tạo workbook mới.

Điểm cần quyết: `ChuSoHuuChinhID` trước đây định nghĩa theo "mã nhánh 1" (đã bỏ). Nếu giữ khối này, cần dẫn xuất `ChuSoHuuChinhID` từ `LaChuDat` của SOT thay vì mã nhánh.

### 1.2. `tblSoHuuDiSan`

Một dòng là quan hệ một người sở hữu một di sản.

| Trường | Ý nghĩa |
| --- | --- |
| `SoHuuID` | ID hệ thống |
| `DiSanID` | Khối di sản |
| `NguoiID` | Chủ hoặc đồng chủ |
| `ThuTuChu` | Thứ tự hiển thị |
| `LaChuChinh` | Chỉ một người có giá trị Có trong mỗi `DiSanID` |
| `VaiTroSoHuu` | Chủ chính, vợ/chồng đồng chủ, đồng chủ khác |
| `TinhTrang` | Còn sống, đã chết hoặc chưa xác định |

### 1.3. Liên kết tài sản

`tblTaiSan` gắn `DiSanID`. Nếu cần sở hữu khác nhau theo từng tài sản, thêm `TaiSanID` vào `tblSoHuuDiSan`; để trống nghĩa là quan hệ sở hữu áp dụng cho cả khối di sản.

### 1.4. `tblPhanChiaDiSan`

Một dòng là quan hệ tài sản/phần di sản × người nhận.

| Trường | Ý nghĩa |
| --- | --- |
| `PhanChiaID` | ID hệ thống |
| `DiSanID`, `TaiSanID` | Khối/tài sản được phân chia |
| `NguoiID` | Người nhận |
| `HinhThucNhan` | Toàn bộ, tỷ lệ, phần mô tả hoặc chưa xác định |
| `TyLe` | Phần trăm nếu áp dụng |
| `NoiDungPhanNhan` | Mô tả phần nhận nếu không dùng tỷ lệ |
| `GhiChu` | Ngoại lệ |

Một `NguoiID` có thể đồng thời xuất hiện trong `tblSoHuuDiSan` và `tblPhanChiaDiSan`.

### 1.5. Validation riêng của khối này

- Mỗi `DiSanID` có đúng một chủ chính; tham chiếu `NguoiID` phải có thật; không xóa người đang được phân chia tham chiếu.
- Nếu nhập tỷ lệ thì tổng tỷ lệ mỗi tài sản không vượt 100%.
- `HinhThucNhan = tỷ lệ` thì `TyLe` bắt buộc, ngược lại `TyLe` phải trống.
- Cảnh báo khi tài sản chưa có phân chia.

## 2. Giao diện nhập liệu tối giản

SOT tập trung vào cấu trúc dữ liệu nhánh thừa kế và toolbar `Số tầng` + `Xuất Văn bản`; SOT không nói về bố cục sheet nhập chính, vùng tài sản, hay quy ước ẩn mã kỹ thuật. Phần dưới đây là phần còn dùng được của `approved-minimal-entry-ui.md` sau khi đã bỏ hết cột/nút liên quan mã nhánh.

- Sheet nhập chính chỉ có hai vùng nhập liệu: `Thông tin người` và `Thông tin tài sản`.
- Không hiển thị trên sheet nhập chính: tiến độ hồ sơ, sơ đồ nhánh thừa kế, bảng phân chia, cảnh báo chi tiết, chú thích giải thích nghiệp vụ, hay các bảng kỹ thuật.
- Phần đầu sheet giữ tối giản: mã/tên hồ sơ để nhận biết hồ sơ đang mở, danh sách chọn `Loại hồ sơ`, lệnh `Lưu`, lệnh `Xuất Word`. Không đưa thanh tiến độ, tình trạng hay ghi chú nghiệp vụ vào phần đầu sheet.

### 2.1. Vùng `Thông tin người`

Một dòng là một người. Cột hiển thị và cho phép nhập, theo thứ tự:

| Thứ tự | Trường nhập |
| --- | --- |
| 1 | Họ và tên |
| 2 | Ngày sinh |
| 3 | Ngày mất |
| 4 | Số giấy tờ |
| 5 | Ngày cấp |
| 6 | Địa chỉ |

Có lệnh thêm dòng/người. Nhập liên tục trong một danh sách; không tách danh sách theo vai trò, Bên A/B hoặc nhánh thừa kế.

### 2.2. Vùng `Thông tin tài sản`

Mỗi tài sản là một phiếu dọc độc lập, đánh số theo thứ tự hiển thị (`Tài sản 1`, `Tài sản 2`...), không phải một bảng ngang dài. Trường MVP: Loại tài sản (danh sách chọn), Số phát hành/số GCN, Số thửa, Tờ bản đồ, Địa chỉ, Diện tích (m²), Nguồn gốc.

Có lệnh thêm tài sản. Khi cần thêm trường theo template hoặc biểu mẫu, bổ sung vào phiếu dọc; không chuyển vùng này thành bảng ngang nhiều cột.

### 2.3. Mã kỹ thuật và dữ liệu ẩn

`NguoiID` (`P0001`) và `TaiSanID` (`TS001`) do VBA tự sinh, ẩn khỏi giao diện nhập chính. Người dùng không tự gõ, sửa hoặc cần nhớ mã. Lý do cần ID ổn định: người dùng đổi họ tên hoặc hai người trùng tên, danh sách bị lọc/sắp xếp/chèn dòng, một người hoặc tài sản được tham chiếu trong quan hệ sở hữu/phân chia/ánh xạ xuất Word.

`STT` và nhãn `Tài sản 1`, `Tài sản 2` chỉ phục vụ quan sát/thứ tự nhập; không phải khóa liên kết.

### 2.4. Sheet phụ

Sheet phụ lưu dữ liệu không cần nhập thường xuyên: thông tin hồ sơ và cấu hình loại hồ sơ, khối di sản/chủ chính/đồng chủ và quan hệ sở hữu, phân chia di sản, quy tắc giấy tờ, trường tự do, cấu hình template và dữ liệu xuất trung gian.

### 2.5. Ghi chú demo

`prototypes/ho-so-thua-ke-demo.html` theo mô hình mã nhánh cũ nên đã lỗi thời. Demo hiện hành theo mô hình SOT là `prototypes/demo-nhanh-con.html`.

## 3. Hồ sơ nhiều loại và hồ sơ hai bên

SOT chỉ mô tả nhánh thừa kế (`LoaiHoSo = Thừa kế`); SOT không đề cập hồ sơ chuyển nhượng hai bên, ủy quyền hay biểu mẫu khác. Ghi lại nguyên trạng để người dùng quyết có đưa vào MVP hay không.

- Workbook dự kiến phục vụ nhiều loại hồ sơ: thừa kế, phân chia/chuyển nhượng hai bên, ủy quyền, biểu mẫu khác.
- `tblNguoi` dùng chung cho mọi loại hồ sơ, gồm các cột: `HoSoID`, `NguoiID`, `Ben`, `STTBen`, `VaiTro`, `HoTen`, `NgaySinh`, `NgayChet`, `SoGiayTo`, `NgayCap`, `DiaChi`, `GhiChu`.
- Cột `Ben` chỉ dùng cho hồ sơ hai bên; bắt buộc trống và bị khóa ở hồ sơ thừa kế.

### 3.1. UX toggle `Ben`

Dùng chính ô trong cột `Ben` như một nút trượt, không tạo control/ActiveX riêng cho từng dòng (dễ lệch khi lọc, sắp xếp, chèn dòng, hoặc mở trên máy khác). Dòng mới của hồ sơ hai bên mặc định là `A`. Ô `A` hiển thị nền xanh, ô `B` hiển thị nền cam. Vẫn có dropdown `A/B/trống` làm phương án dự phòng khi macro bị tắt.

**Điểm cần chốt:** tài liệu cũ dùng double-click để đổi `A` ↔ `B` (lý do: click đơn còn cần để chọn ô, sửa dữ liệu, dùng phím mũi tên, lọc bảng). Nhưng SOT §8.1 cấm double-click cho nút đổi tầng, yêu cầu một-lần-bấm. Hai cơ chế nằm ở hai cột khác nhau nên không mâu thuẫn trực tiếp, nhưng tạo hai quy ước thao tác khác nhau trong cùng workbook. Cần chốt một quy ước thống nhất trước khi triển khai.

### 3.2. Bảng xuất ẩn theo bên

Sheet kỹ thuật `XuatAn` đặt `VeryHidden`, gồm `tblXuatBenA` và `tblXuatBenB` — các dòng `tblNguoi` tách theo `Ben`. Không phải nguồn dữ liệu, chỉ do VBA tạo/refresh khi xuất: đọc `tblNguoi` theo `HoSoID`, tách theo `Ben`, đánh lại `STTBen` liên tục từ 1 cho mỗi bên, refresh hai bảng, rồi tạo `ExportMap`.

### 3.3. Validation hồ sơ hai bên

- Hồ sơ có chia hai bên phải có ít nhất một người ở mỗi bên.
- `Ben` chỉ nhận `A`, `B` hoặc trống; người thuộc hồ sơ thừa kế phải để trống, sự kiện toggle bị hủy trong hồ sơ thừa kế.
- `STTBen` do VBA tạo lại, không cho nhập tay.

### 3.4. Bảng phụ liên quan

`tblHoSo(HoSoID, LoaiHoSo, TenHoSo, CoChiaHaiBen, ...)`, `tblNguoiUyQuyen` (thông tin ủy quyền chuyên biệt), `tblQuyTacGiayTo` (loại giấy tờ, nơi cấp, nhãn địa chỉ suy theo ngày cấp), `tblTruongTuDo(Placeholder, GiaTri)` cho trường riêng của từng template.

## 4. Cấu trúc workbook và hạ tầng xuất Word

SOT không nói gì về số lượng/tên sheet, định dạng file, hay cách macro tạo phiên Word. Toàn bộ hạ tầng kỹ thuật này SOT im lặng và vẫn cần cho bất kỳ mô hình dữ liệu nào (kể cả mô hình SOT).

### 4.1. Cấu trúc sheet dự kiến

`Tổng quan` (dashboard: thông tin hồ sơ, đường dẫn mẫu/kết quả, bảng chọn mẫu, định dạng xuất, trạng thái kiểm tra, các nút thao tác), `Người`, `Tài sản`, `Danh mục`, `Trường mở rộng` (bảng `Mã trường – Giá trị`, cho phép thêm placeholder không cần thêm cột hoặc sửa VBA), `Xuất Word` (chọn mẫu, thư mục, Word/PDF).

File mới là `.xlsm`. Tạo hai bản: template sạch và bản đã chuyển dữ liệu; giữ `.xlsb` gốc bên ngoài. Không mang sheet cũ và `tblDanhSach` vào file mới.

### 4.2. Xuất Word an toàn

- Dùng `CreateObject` để tạo phiên Word **riêng**, không dùng `GetObject`. Bệnh cũ: macro dùng `GetObject` bám vào phiên Word đang mở của người dùng rồi gọi `WordApp.Quit`, làm đóng luôn cả tài liệu khác của người dùng.
- Chỉ mở/đóng tài liệu do phiên đó tạo, lưu file rồi thoát đúng phiên đó.
- Thay placeholder trong body, bảng, header, footer, textbox bằng Word Range — không dùng `Selection.Find` (không phủ hết header/footer/textbox).
- Giá trị trống xuất chuỗi rỗng, không xuất `0` hay `00/01/1900`.
- Bỏ phụ thuộc add-in `codedocso.xlam` đang gây lỗi `#NAME?`.
- Sau khi xuất chỉ hiển thị kết quả, không tự mở file hoặc thư mục.
- Bỏ UserForm cũ, đưa điều khiển lên dashboard; loại bỏ control ép người dùng chọn ô `B6`.

### 4.3. Chiến lược placeholder kép

Giữ placeholder cũ dạng `[Tên 4]`, `[CCCD 4]`, `[Serial 2]` để tương thích các mẫu Word hiện có, đồng thời hỗ trợ namespace mới: `{{ho_so.niem_yet}}`, `{{ho_so.so_cong_chung}}`, `{{tai_san.2.serial}}`, `{{tai_san.2.dien_tich}}`.

Quy tắc parse: đoạn cuối là tên trường, các đoạn số ở giữa là định danh đối tượng. Block lặp động trong Word (ví dụ lặp cho danh sách người không giới hạn) để giai đoạn sau.

### 4.4. Quy tắc placeholder bắt buộc

- Giữ đúng cú pháp ngoặc vuông hoặc `{{ }}`; tên trong placeholder phải trùng chính xác header bảng nguồn.
- Không để một placeholder bị tách thành nhiều run trong Word — Find/Replace sẽ không nhận diện được.
- Thêm field mới phải làm đủ ba việc: thêm header nguồn, tạo công thức/dữ liệu tương ứng, thêm placeholder cùng tên vào Word.

### 4.5. Xử lý đặc biệt cần giữ trong logic thay thế

- Header bắt đầu bằng "Bảng" thì chèn Excel Table từ sheet `Tables` vào vị trí placeholder.
- Nội dung có xuống dòng thì đổi sang ký tự newline của Word.
- Chuỗi dài hơn 254 ký tự phải dùng `TypeText` thay vì `Replacement.Text` (giới hạn của Find/Replace).
- Giá trị là đường dẫn `.jpg/.png/.bmp` thì chèn ảnh thay vì chèn text.
- Tên file đầu ra dạng `TênMẫu_GiáTrịCộtĐượcChọn.docx|.pdf`; chống trùng tên bằng hậu tố `_1`, `_2`.
- Form lưu đường dẫn folder mẫu/folder xuất vào Names của workbook: `DuongDanFolderVB`, `DuongDanFolderLuu`.

### 4.6. Nguyên tắc chung

Không dùng Mail Merge hay liên kết Word–Excel; file Word xuất ra là bản sao độc lập của template; các bảng xuất trung gian chỉ do VBA tạo/refresh, người dùng không nhập vào đó.

## 5. Quy tắc nghiệp vụ đang chạy trong workbook cũ

SOT không đề cập cách suy ra loại giấy tờ, nhãn địa chỉ, hay chuẩn hóa ngày — đây là logic nghiệp vụ cụ thể của hồ sơ đất đai, độc lập với mô hình cây thừa kế.

- Loại giấy tờ theo ngày cấp: trước 01/07/2024 → "Căn cước công dân"; từ 01/07/2024 → "Căn cước". Nơi cấp tra theo sheet `Tables`.
- Nhãn địa chỉ suy theo loại giấy tờ: "Thường trú tại" hoặc "Cư trú tại".
- Chuẩn hóa ngày: chỉ nhập năm thì đổi ngầm thành `01/01/năm`.
- Tên file phụ hiện lấy từ từ cuối của họ tên, chữ thường — nên đổi sang mã hồ sơ có kiểm soát ký tự để tránh trùng/lỗi ký tự đặc biệt.
- Hậu tố tài sản: tài sản 1 không hậu tố, tài sản 2 hậu tố `2`, tài sản 3 hậu tố `3`.
- Danh mục trên sheet `Tables`: loại giấy tờ → cơ quan cấp → nhãn địa chỉ; loại GCN; loại đất (đất ở đô thị, đất ở nông thôn, cây lâu năm, nuôi trồng thủy sản, trồng lúa); hình thức sử dụng riêng/chung; người ủy quyền; cơ quan cấp GCN.
- Hai bảng tra cứu chưa được VBA gọi tự động: `Tờ bản đồ 2025` (điều chỉnh số tờ sau sáp nhập) và `Sáp nhập thôn` (thôn/xóm huyện Ý Yên); hàm `FindLongestMatch(Target, LookupRange)` đã viết nhưng chưa được gọi ở đâu.
- Trường có trong workbook nhưng chưa dùng trong mẫu Word: giá chuyển nhượng, số điện thoại.

## 6. Sức chứa template và rủi ro mất dữ liệu

SOT có mục validation riêng (§14) nhưng chỉ nói về nhánh/tầng; SOT không nói cụ thể về giới hạn số người/tài sản của từng mẫu Word thật. Đây là rủi ro kỹ thuật cụ thể của mẫu hiện có, cần biết trước khi thiết kế lại.

- Mẫu `1. PCDS .docx`: 99 placeholder khác nhau, 300 lần xuất hiện. Là văn bản thỏa thuận phân chia di sản thừa kế, có thêm phần chứng thực và danh sách hàng thừa kế.
- Mẫu là template **tĩnh**: khi người không có dữ liệu, macro chỉ thay bằng rỗng, **không tự xóa** đoạn văn, dòng danh sách hay dòng bảng tương ứng.
- Nhóm placeholder người `n = 1..10`: `[Tên n] [Năm sinh n] [CCCD n] [Ngày cấp n] [Địa chỉ n] [Loại CC n] [Nơi cấp CC n] [Thường trú n]`. Người 1 và 2 có thêm `[Năm chết]`, `[Năm chết 2]`.
- Nhóm placeholder tài sản/hồ sơ: `[Niêm Yết] [Loại sổ] [Địa chỉ đất] [Serial] [Số vào sổ] [Số thửa] [Số tờ] [Diện tích] [Hình thức sử dụng] [Loại đất] [Thời hạn 1] [Nguồn gốc] [Ngày cấp sổ] [Cơ quan cấp sổ] [ONT] [CLN] [NTS]`.
- Workbook cũ: sheet `HĐCN` hàng 10–29 nhập 20 người (STT ở cột B), hàng 31–53 dữ liệu chung/tài sản ở cột C/D/E (tài sản 1/2/3). `tblDanhSach` là `HĐCN!A5:HZ7`, 233–257 cột đã dùng; hàng 5 là tên trường, hàng 6 là dữ liệu merge; form ép đọc vùng `B6`.
- **Rủi ro cần xử lý:** nhập liệu hiện không giới hạn số người, nhưng mẫu Word chỉ chứa placeholder cho tối đa 10 người và 1 tài sản. Người 11–20 có header nhưng hàng 6 không có công thức nên không thể xuất. Người 20 còn thiếu `Nơi cấp CC 20` và `Thường trú 20` trong header nguồn. Nếu thiết kế mới không kiểm tra hai chiều (số liệu nhập vs. sức chứa mẫu đang chọn), rủi ro mất dữ liệu âm thầm sẽ lặp lại.
- Kế hoạch kiểm thử đề xuất: chuyển dữ liệu người/tài sản/hồ sơ từ workbook đã lưu; thử thêm người thứ 21–25 bằng cách thêm dòng; kiểm tra placeholder cũ, tài sản, trường mở rộng; xuất DOCX và PDF **trong khi có Word khác đang mở** rồi xác nhận tài liệu khác không bị tác động; không dùng dữ liệu thật để kiểm thử bản phát hành, dùng dữ liệu giả lập.

## 7. Lỗi kỹ thuật đã ghi nhận, chưa sửa

SOT chưa triển khai VBA/workbook (SOT dòng 3: "chưa triển khai VBA/workbook"), nên các lỗi cụ thể trong workbook cũ dưới đây chưa nằm trong phạm vi SOT, nhưng cần biết trước khi tái cấu trúc để không mang lỗi cũ sang bản mới.

| Lỗi | Vị trí | Hướng xử lý |
| --- | --- | --- |
| `Loại đất 2` lấy sai nguồn | Công thức `GY6` đang trỏ `D40` (hình thức sử dụng), đúng phải là `D41` | Sửa công thức `GY6` trỏ dòng 41 |
| `#NAME?` khi đọc số bằng chữ | `Diện tích chữ`, `Ngày chữ`, `Tháng chữ` ở bản đã chuyển đổi | Kiểm tra add-in/hàm đọc số bằng chữ còn khả dụng ở máy đích |
| Thiếu công thức | `Diện tích chữ 2`, `Diện tích chữ 3` | Bổ sung công thức tương ứng |
| `#REF!` | `Tables!E17:F18` (từng là danh mục người từ chối) | Sửa tham chiếu; lưu ý SOT coi `TuChoi` là phái sinh (§11) nên danh mục này chỉ còn giá trị câu chữ, không còn là dữ liệu nguồn |
| Không phủ hết vùng thay thế | Macro dùng `Selection.Find` | Chuyển sang Range/StoryRange để phủ header, footer, textbox |
| Tự lưu ngoài ý muốn | `ThisWorkbook.Workbook_BeforeClose` tự lưu workbook khi đóng | Cân nhắc bỏ hoặc biến thành tùy chọn |
| Hàm chưa dùng | `Module1.FindLongestMatch` đã viết nhưng chưa được gọi | Gọi trong luồng tra cứu tờ bản đồ/sáp nhập thôn, hoặc bỏ nếu không cần |

## 8. Ràng buộc môi trường và công cụ

SOT là đặc tả nghiệp vụ/dữ liệu, không đề cập ràng buộc vận hành của công cụ agent hay môi trường Office. Phần này cần biết để không làm hỏng phiên Office của người dùng khi triển khai bất kỳ spec nào.

- OfficeCLI chỉ hỗ trợ đọc/sửa định dạng gốc `.docx/.xlsx/.pptx`. File `.doc` và `.xlsb` phải mở qua Word/Excel cài sẵn ở chế độ ẩn, chỉ-đọc, **tắt macro**, không lưu.
- Không đụng vào phiên Word/Excel đang mở của người dùng; phải đóng sạch tiến trình ẩn do agent tạo ra.
- Cần bật một lần: `File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model`.
- Nếu workbook đang mở và có thay đổi chưa lưu, phải yêu cầu người dùng lưu/đóng trước khi triển khai.
- Không commit `OfficeCLI` vào repo dự án (kích thước lớn, ~275 MB); loại trừ `~$*`, `*.tmp`, file khóa/tạm.
- Repo là public, chỉ chứa dữ liệu mẫu, không phải dữ liệu cá nhân thật.
- Lưu ý: các tài liệu lịch sử (session 24/08, 26/08) ghi workspace là `D:\TK`; workspace hiện tại của repo là `D:\excelTK`.

## 9. Câu hỏi còn mở, SOT chưa trả lời

1. Chưa có cơ chế tạo hồ sơ mới hoặc chọn hồ sơ đang làm việc trong workbook.
2. Nút `Lưu` chưa được định nghĩa hành vi cụ thể. Xuất Word có tự động Lưu trước không?
3. Mô hình "một workbook = một hồ sơ" đã đúng chưa — copy file mẫu vào thư mục riêng, xóa dữ liệu cũ, mở lên nhập?
4. Vùng `Thông tin người` chỉ có 6 cột (mục 2.1) nên hồ sơ hai bên không có chỗ gán `A`/`B`. Cột `Ben` đang được mô tả theo ba cách khác nhau trong các tài liệu cũ (ẩn hẳn / hiện nhưng xám khóa / hiện và toggle được). Chốt theo cách nào?
5. Đổi `Loại hồ sơ` giữa chừng (ví dụ từ hai bên sang thừa kế) tạo ngõ cụt: dữ liệu `Ben` đã nhập nhưng cột bị khóa lại. Xử lý bằng cách nào — cảnh báo, chặn đổi loại, hay tự xóa dữ liệu `Ben`?
6. `tblPhanChiaDiSan` (mục 1.4) yêu cầu nhập `TaiSanID/NguoiID/DiSanID`, trái nguyên tắc "người dùng không tự gõ mã kỹ thuật" (mục 2.3). Có nên đổi sang dropdown chọn tên người + nhãn tài sản, để VBA tự quy ra ID khi lưu?
7. Cách hiển thị lỗi validation chưa được định nghĩa: hiển thị danh sách lỗi theo tên người, có cho phép bấm để nhảy tới ô lỗi hay không?
8. Hồ sơ hai bên (chuyển nhượng) và hồ sơ ủy quyền có nằm trong phạm vi MVP hiện tại hay không? Câu trả lời ảnh hưởng trực tiếp đến việc có cần giữ cột `Ben`, sheet `XuatAn`, và `tblNguoiUyQuyen` hay không.
9. Quy ước thao tác chuột chưa thống nhất trong toàn workbook: SOT cấm double-click cho nút đổi tầng (yêu cầu một-lần-bấm), nhưng cột `Ben` cũ dùng double-click để toggle. Thống nhất thành một quy ước chung, hay chấp nhận hai cột có hai quy ước khác nhau với lý do khác nhau?
10. Hướng dài hạn đã được ghi nhận trong tài liệu khảo sát: Excel VBA nên được coi là "bản mô tả nghiệp vụ và nguồn dữ liệu khởi đầu", không phải lõi dài hạn của hệ thống; đích kiến trúc là ảnh/nhập tay → chuẩn hóa dữ liệu → cơ sở dữ liệu → chọn mẫu → sinh Word/PDF. Có tiếp tục giữ định hướng này làm kim chỉ nam, hay xác định Excel/VBA là giải pháp lâu dài?
11. Mẫu Word hiện tại (`1. PCDS .docx`) và các mẫu khác trong `templates/word/` chưa được kiểm tra khả năng tương thích với mô hình `Tang + ParentNguoiID` của SOT — cần rà soát placeholder nào còn dùng được, placeholder nào phải viết lại theo `CayNhanh`/`CacNhomTuChoi` (SOT §13).
