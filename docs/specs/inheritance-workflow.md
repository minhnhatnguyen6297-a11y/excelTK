# Đặc tả chi tiết: Hồ sơ thừa kế

## 1. Mục tiêu và quy ước

Mục tiêu là nhập một danh sách người duy nhất, nhưng cho phép:

- Một hoặc nhiều chủ sở hữu, thường là một người hoặc hai vợ chồng.
- Một tài sản có nhiều chủ sở hữu.
- Một tài sản có nhiều người nhận di sản.
- Một người vừa là chủ/đồng chủ, vừa là người nhận di sản.
- Một người thừa kế đã chết có thể tạo một nhánh người thừa kế tiếp theo.

`Bên A/B` không xuất hiện trong hồ sơ thừa kế. Quan hệ pháp lý được thể hiện bằng vai trò, sở hữu tài sản, nhánh thừa kế và phân chia di sản.

Trong hồ sơ thừa kế, hai người đầu tiên của `tblNguoi` là **vợ chồng chủ đất**: mã nhánh `1` là chồng/chủ đất chính, mã nhánh `2` là vợ. Cặp gốc là **một đơn vị thừa kế duy nhất** (di sản khối chung vợ chồng được chia một lần cho nhóm thừa kế chung). Khi lưu hồ sơ, VBA ghi `NguoiID` của mã `1` vào `ChuSoHuuChinhID`.

Mỗi hồ sơ dùng một bản sao của file mẫu trong thư mục riêng; tạo hồ sơ mới bằng cách copy file mẫu và xóa dữ liệu cũ.

## 2. Mô hình dữ liệu

### `tblNguoi` trên sheet `Nguoi`

Nguồn thông tin nhân thân duy nhất. Mỗi người có `HoSoID`, `NguoiID`, `STTNhap`, họ tên, ngày sinh, ngày chết, giấy tờ, địa chỉ và ghi chú. Không dùng các cột này để lưu nhiều vai trò hoặc tỷ lệ tài sản.

Với hồ sơ thừa kế, `tblNguoi` thêm ba cột hệ thống:

| Trường | Ý nghĩa |
| --- | --- |
| `MaNhanh` | Khóa cấu trúc cây: `1`, `2`, `1.4`, `1.4.4`... Sinh tự động theo quy tắc §4 |
| `TrangThaiTN` | Hưởng, Từ chối, Đã chết, Không hưởng, Chưa xác định |
| `QuanHe` | Nhãn chữ phục vụ câu chữ Word ("con nuôi", "cha đỡ đầu"); không ảnh hưởng cấu trúc |

Cây thừa kế được mã hóa ngay trong `MaNhanh`; **không cần bảng quan hệ thừa kế riêng**. Nhóm người trong văn bản được truy vấn bằng tiền tố mã kết hợp lọc `TrangThaiTN` (§5). Người dùng không tự gõ hoặc sửa `MaNhanh`.

### `tblDiSan` trên sheet `Phu`

Mỗi dòng là một khối di sản trong hồ sơ.

| Trường | Ý nghĩa |
| --- | --- |
| `DiSanID` | ID hệ thống, ví dụ `DS001` |
| `HoSoID` | Hồ sơ sở hữu di sản |
| `TenDiSan` | Tên gợi nhớ, ví dụ `Di sản đất tại ...` |
| `ChuSoHuuChinhID` | ID của chủ đất (mã nhánh `1`) |
| `CheDoSoHuu` | Riêng, chung vợ chồng hoặc chung khác |
| `GhiChu` | Ngoại lệ |

MVP tạo một `DiSanID` mặc định cho hồ sơ. Nếu sau này một hồ sơ có nhiều khối di sản độc lập, người dùng tạo thêm dòng thay vì tạo workbook mới.

### `tblSoHuuDiSan` trên sheet `Phu`

Một dòng là quan hệ **một người sở hữu một di sản**.

| Trường | Ý nghĩa |
| --- | --- |
| `SoHuuID` | ID hệ thống |
| `DiSanID` | Khối di sản |
| `NguoiID` | Chủ hoặc đồng chủ |
| `ThuTuChu` | Thứ tự hiển thị |
| `LaChuChinh` | Chỉ một người có giá trị Có trong mỗi `DiSanID` |
| `VaiTroSoHuu` | Chủ chính, vợ/chồng đồng chủ, đồng chủ khác |
| `TinhTrang` | Còn sống, đã chết hoặc chưa xác định |

Khi tạo hồ sơ, VBA tự tạo dòng chủ chính từ người mã `1`. Nút `Thêm đồng chủ` chỉ thêm một quan hệ mới, không tạo bản sao người.

### `tblTaiSan` trên sheet `TaiSan`

Mỗi tài sản có `TaiSanID` và `DiSanID`. Nhiều tài sản có thể cùng thuộc một khối di sản. Nếu cần sở hữu khác nhau theo từng tài sản, thêm `TaiSanID` vào `tblSoHuuDiSan`; để trống nghĩa là quan hệ sở hữu áp dụng cho cả khối di sản.

### `tblPhanChiaDiSan` trên sheet `Phu`

Một dòng là quan hệ **tài sản/phần di sản × người nhận**.

| Trường | Ý nghĩa |
| --- | --- |
| `PhanChiaID` | ID hệ thống |
| `DiSanID`, `TaiSanID` | Khối/tài sản được phân chia |
| `NguoiID` | Người nhận |
| `HinhThucNhan` | Toàn bộ, tỷ lệ, phần mô tả hoặc chưa xác định |
| `TyLe` | Phần trăm nếu áp dụng |
| `NoiDungPhanNhan` | Mô tả phần nhận nếu không dùng tỷ lệ |
| `GhiChu` | Ngoại lệ |

`NguoiID` có thể đồng thời xuất hiện trong `tblSoHuuDiSan` và `tblPhanChiaDiSan`.

## 3. Luồng nhập liệu

1. Copy file mẫu vào thư mục hồ sơ, xóa dữ liệu cũ, mở workbook, chọn `LoaiHoSo = Thừa kế`.
2. Nhập hai vợ chồng chủ đất vào hai dòng đầu; mã `1`, `2` tự gán. Điền ngày mất cho người đã mất.
3. Nhập tài sản theo phiếu dọc.
4. Chọn dòng một người rồi bấm nút vai trò để mở nhánh: `Thêm bố`, `Thêm mẹ`, `Thêm vợ/chồng`, `Thêm con`, `Con đã mất – mở nhánh`, `Người khác`. VBA sinh mã theo §4, tạo dòng người mới và đưa con trỏ tới ô `HoTen`.
5. Người từ chối nhận di sản: đổi `TrangThaiTN` của dòng đó thành `Từ chối`.
6. Phân chia: chọn tên người nhận (dropdown) và nhãn tài sản; VBA quy ra ID khi lưu vào `tblPhanChiaDiSan`.

## 4. Quy tắc mã nhánh

Số cuối của mỗi đoạn mã cố định vai trò:

| Số cuối | Vai trò | Ràng buộc |
| --- | --- | --- |
| `.1` | Bố đẻ của người mở nhánh | Tối đa một mỗi nhánh |
| `.2` | Mẹ đẻ | Tối đa một mỗi nhánh |
| `.3` | Vợ/chồng | Cấm ở độ sâu 1 (vợ của `1` là người `2`); từ sâu ≥ 2, mặc định `Không hưởng` |
| `.4` trở đi | Con (con chung, nuôi, ngoài giá thú — phân biệt bằng `QuanHe`) | Sinh tiếp từ `.4` trong nhánh |

Quy tắc đi kèm:

- Con của người đã chết bỏ qua slot `1/2/3` (tổ tiên và vợ/chồng đã tồn tại trên cây): con của `1.4` bắt đầu từ `1.4.4`.
- Con chung của cặp gốc đặt dưới **nhánh `1`**; nhóm thừa kế gia đình gốc là `HopNhanh("GOC")` = hợp nhánh `1` ∪ nhánh `2`.
- `MaNhanh` bất biến: xóa hoặc sửa người không đánh lại mã khác; giữ lỗ số để placeholder template ổn định.
- Sắp xếp hiển thị theo giá trị số từng đoạn (`1.9` trước `1.10`), không sắp xếp theo chuỗi.
- Case ít gặp (ông bà thuộc hàng thừa kế thứ hai, anh chị em...): nút `Người khác` sinh mã tiếp số lớn nhất, người dùng ghi `QuanHe` bằng tay.
- `MaNhanh` không phải nhận định pháp lý về quan hệ; nhận định nằm ở `QuanHe` và `TrangThaiTN`.

## 5. Hàm nhóm và xuất Word

| Hàm | Chữ ký | Hành vi |
| --- | --- | --- |
| `MaTiepTheo` | `(maGoc, vaiTro)` | Sinh mã slot tất định; slot đã có thì báo lỗi thân thiện |
| `DoiTienTo` | `(oldPrefix, newPrefix)` | Đổi tiền tố cả cây con khi chuyển nhánh; kiểm tra xung đột trước khi ghi |
| `NhomTheoMa` | `(prefix)` | Mảng dòng có mã bằng hoặc bắt đầu `prefix.` |
| `HopNhanh` | `(prefix hoặc "GOC")` | Duyệt cây: nút `Đã chết` thay bằng tập con của nó; trả danh sách người hưởng thực tế |
| `LocTrangThai` | `(danhSach, trangThai)` | Lọc theo `TrangThaiTN` |
| `NoiDanhSach` | `(danhSach, kieu)` | Ghép tên kiểu văn xuôi ("A, B và C") hoặc kiểu bảng |

UDF mỏng cho sheet ánh xạ template: `=NguoiThuaKe("1.4")` trả chuỗi danh sách tên người hưởng của nhánh.

Placeholder:

- Cá nhân: `{{nguoi.<ma>.<truong>}}`, ví dụ `{{nguoi.1.4.ho_ten}}`.
- Nhóm: `{{nhom.<ma>.danh_sach}}`, `{{nhom.goc.danh_sach}}`, `{{nhom.<ma>.so_luong}}` — chạy chuỗi `HopNhanh` → `LocTrangThai("Hưởng")` → `NoiDanhSach`.

Template cũ vẫn dùng placeholder theo số thứ tự; trước xuất, VBA tạo thứ tự Word theo cấu hình template bằng cách sắp các nhánh theo số từng đoạn của `MaNhanh`; thứ tự này không phải `STTNhap`.

## 6. Kiểm tra trước khi xuất

- `Ben` trống ở toàn bộ người của hồ sơ thừa kế.
- Mỗi `DiSanID` có đúng một chủ chính và tham chiếu `NguoiID` có thật; không xóa người đang được phân chia tham chiếu.
- `MaNhanh` duy nhất, đúng khuôn dạng số(`.số`)*; tiền tố cha tồn tại.
- Slot `.1/.2/.3` không lặp trong cùng nhánh; `.3` cấm ở độ sâu 1.
- Mọi nút trung gian của thế vị phải `Đã chết` (`1.4` còn sống thì không được tồn tại `1.4.x`).
- Ít nhất một người gốc (`1`/`2`) đã chết.
- Phân chia: tài sản/người nhận có thật; nếu nhập tỷ lệ thì tổng tỷ lệ mỗi tài sản không vượt 100%; `HinhThucNhan = tỷ lệ` thì `TyLe` bắt buộc, ngược lại `TyLe` phải trống.
- Số người hưởng không vượt sức chứa template (mẫu cũ: 10 người, 1 tài sản).
- Cảnh báo khi tài sản chưa có phân chia hoặc template còn placeholder không được thay.

## 7. Phạm vi MVP thừa kế

1. Hai người gốc vợ chồng (mã `1`, `2`) là một đơn vị thừa kế gộp.
2. Mã nhánh vai trò cố định lưu trong `tblNguoi`; không có bảng quan hệ thừa kế.
3. Truy vấn nhóm bằng tiền tố mã + trạng thái; placeholder nhóm cho văn bản.
4. Phân chia qua `tblPhanChiaDiSan` với dropdown tên người/nhãn tài sản.
5. Khóa tuyệt đối cột `Ben`.
6. Giữ ánh xạ placeholder template cũ; placeholder nhóm động là bước chuyển tiếp tiếp theo.
