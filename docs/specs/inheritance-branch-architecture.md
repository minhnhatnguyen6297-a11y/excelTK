# Đặc tả kiến trúc nhánh thừa kế theo tầng

> Trạng thái: **đã chốt kiến trúc**, chưa triển khai VBA/workbook.
> Ngày chốt: 27/08/2026.
> Phạm vi: nhập danh sách người, xác định nhánh, người nhận đất, người từ chối và dữ liệu xuất Word.

## 1. Mục tiêu

Giữ nguyên cách làm quen thuộc của người dùng: nhập người liên tục trên một bảng, không phải dừng lại để chọn vai trò cha, mẹ, con hay vợ/chồng.

Mỗi hồ sơ chỉ cần trả lời bốn câu hỏi:

1. Ai là chủ đất ban đầu?
2. Mỗi người thuộc tầng và nhánh nào?
3. Ai đã chết?
4. Ai nhận đất? Những người còn lại được xác định là từ chối.

Chi tiết quan hệ cha, mẹ, con, vợ/chồng chưa phải yêu cầu của phiên bản này.

## 2. Quyết định kiến trúc đã chốt

- Dùng một danh sách người phẳng, mỗi người một dòng.
- Cấu trúc bắt đầu từ `T1` là tầng chủ đất, sau đó là `T2` đến `Tn`.
- Số tầng mặc định là `3`; người dùng được nhập lại nhưng chỉ hỗ trợ từ `1` đến `5`.
- Mỗi tầng có một màu nhạt riêng; toàn bộ dòng đổi màu khi đổi tầng.
- Mỗi dòng có một ô dạng nút, bấm một lần để chuyển tầng. Không dùng double-click.
- `Tang` cho biết độ sâu; từ `T3` trở đi, `ParentNguoiID` cho biết dòng thuộc nhánh của ai. Màu không phải dữ liệu.
- Hai vị trí đầu được dành cho chủ đất. Việc tích chọn chủ đất được thiết kế trong dữ liệu nhưng ẩn hoàn toàn khỏi giao diện MVP.
- Người nhận đất được tích chọn trực tiếp trên từng dòng.
- Người từ chối được tính tự động, không nhập tay.
- Dữ liệu dự phòng cho nhóm từ chối được tạo ngay từ đầu nhưng chưa có UI.

Đây là quyết định của người dùng và là cơ sở để triển khai. Không quay lại mô hình nhập vai trò chi tiết nếu chưa có một yêu cầu nghiệp vụ mới đủ mạnh.

## 3. Luồng sử dụng chính

1. Mở bản sao workbook của hồ sơ.
2. Trên thanh công cụ, giữ `Số tầng = 3` hoặc nhập số từ `1` đến `5`.
3. Nhập chủ đất ở dòng 1; nếu hồ sơ có hai chủ đất thì nhập thêm dòng 2. Nếu chỉ có một chủ đất, để trống dòng 2.
4. Nhập liên tục những người còn lại theo thứ tự từng nhánh.
5. Nếu tầng của dòng nhánh chưa đúng, bấm một lần vào ô `Tầng` để chuyển `T2 → T3 → ... → Tn → T2`.
6. Điền ngày chết nếu người đó đã chết.
7. Tích `Nhận đất` cho người nhận.
8. Bấm `Xuất Văn bản`.

Không có bước chọn vai trò sau mỗi lần thêm người và không có chế độ “đi vào/đi ra nhánh”.

## 4. Ví dụ cấu trúc

| STT | Người | Tầng | Thuộc nhánh của | Kết quả |
| ---: | --- | :---: | --- | --- |
| 1 | Ông A | T1 | — | Đã chết |
| 2 | Bà B | T1 | — | Đã chết |
| 3 | Anh C | T2 | Nhóm chủ đất ban đầu | Đã chết |
| 4 | Chị C1 | T3 | Anh C | Nhận đất |
| 5 | Anh C2 | T3 | Anh C | Từ chối |
| 6 | Chị D | T2 | Nhóm chủ đất ban đầu | Nhận đất |

Hai chủ đất ở `T1` được coi là **một nhóm gốc chung**, không phải hai nhánh riêng. Vì vậy, mọi người `T2` trực thuộc nhóm chủ đất ban đầu và để trống `ParentNguoiID`; không gán tùy ý họ cho chủ đất số 1 hay số 2. Từ `T3` trở đi, hệ thống phải lưu `ParentNguoiID`, vì hai người cùng tầng có thể thuộc các nhánh cha khác nhau.

## 5. Mô hình dữ liệu

### 5.1. Bảng người

| Trường | Hiển thị | Ý nghĩa |
| --- | :---: | --- |
| `NguoiID` | Ẩn | ID bất biến do hệ thống sinh, ví dụ `P0001` |
| `STTNhap` | Có | Thứ tự nhập và thứ tự hiển thị ổn định |
| `HoTen` và thông tin nhân thân | Có | Dữ liệu dùng để tạo văn bản |
| `NgayChet` | Có | Có giá trị nghĩa là người đã chết |
| `Tang` | Có | Số nguyên từ `1` đến `SoTang` |
| `ParentNguoiID` | Ẩn | Để trống ở T1–T2; từ T3 trở đi là người ở tầng liền trên mà nhánh này trực thuộc |
| `LaChuDat` | Ẩn trong MVP | Cờ chủ đất; thiết kế sẵn cho UI tích chọn trong tương lai |
| `NhanDat` | Có | Checkbox người nhận đất |
| `NhomTuChoiID` | Ẩn trong MVP | Khóa nhóm văn bản từ chối, để trống khi chưa phân nhóm |

Không lưu trường `TuChoi`. Đây là kết quả phái sinh để tránh ba trạng thái mâu thuẫn nhau.

### 5.2. Cấu hình hồ sơ

| Trường | Giá trị |
| --- | --- |
| `SoTang` | Mặc định `3`, hợp lệ từ `1` đến `5` |
| `BangMauTang` | Mã bảng màu đang chọn |
| `PhienBanCauTruc` | Phiên bản schema để có thể nâng cấp an toàn |

## 6. Thanh công cụ phía trên

Thanh công cụ dùng chung nằm cố định phía trên vùng nhập người:

```text
Số tầng  [ 3 ▾ ]                              [ Xuất Văn bản ]
```

Quy tắc:

- `Số tầng` là ô nhập số nhỏ, mặc định `3`, tối thiểu `1`, tối đa `5`.
- `Xuất Văn bản` là hành động chính và đặt ở phía phải.
- Chưa thêm nút khác. Chỉ thêm khi một hành động thực sự dùng chung và đủ thường xuyên.
- Nếu giảm số tầng thấp hơn tầng đang có dữ liệu, không tự sửa dữ liệu; chặn và chỉ rõ các dòng cần xử lý.
- Nếu tăng số tầng, bảng màu và vòng chuyển tầng cập nhật ngay.

## 7. Chủ đất

### 7.1. Hành vi MVP

- Vị trí 1 và 2 là hai ô chủ đất mặc định.
- Hồ sơ một chủ đất: nhập người ở vị trí 1, để vị trí 2 trống.
- Hồ sơ hai chủ đất: nhập cả vị trí 1 và 2.
- Các dòng có dữ liệu ở hai vị trí này được gán `Tang = 1` và `LaChuDat = True`.
- Hai vị trí chủ đất được giữ cố định ở `T1`; nút tầng chỉ đổi được từ dòng nhánh đầu tiên trở đi.
- Người dùng không nhìn thấy checkbox `LaChuDat` và không phải thao tác chọn chủ đất.

### 7.2. Chuẩn bị cho tương lai

Mã xử lý và xuất Word phải lọc theo `LaChuDat`, không được viết cứng điều kiện `STT <= 2`.

Trong code phải có ghi chú tương đương:

```text
MVP: UI chọn chủ đất đang bị ẩn; mặc định dùng hai vị trí đầu.
Future: hiển thị checkbox LaChuDat để hỗ trợ 3+ đồng chủ đất.
Không thay schema hoặc hợp đồng xuất Word khi mở tính năng này.
```

## 8. Nút đổi tầng một lần bấm

### 8.1. Hình thức

Không tạo Shape, Form Control hoặc ActiveX riêng cho từng dòng vì các điều khiển này dễ lệch khi thêm, lọc hoặc sắp xếp dòng.

Dùng chính ô trong cột `Tầng` như một nút:

```text
[ T1 › ]   [ T2 › ]   [ T3 › ]
```

Ô có nền đậm hơn màu của tầng, viền rõ và con trỏ/ghi chú cho biết “Bấm để đổi tầng”. Một lần chọn ô phải chạy ngay; không dùng double-click.

### 8.2. Hành vi

- Hai vị trí chủ đất cố định ở `T1`; bấm vào đó không đổi tầng.
- Với dòng nhánh, bấm vào ô tầng để chuyển sang tầng kế tiếp trong phạm vi `T2..SoTang`.
- Khi dòng nhánh đang ở tầng cuối: vòng về `T2`, không vòng về `T1`.
- Sau khi xử lý, chuyển vùng chọn sang ô `Họ tên` cùng dòng để lần bấm sau luôn được nhận.
- Cả dòng đổi màu bằng conditional formatting dựa trên `Tang`.
- Dòng mới mặc định kế thừa tầng của dòng ngay phía trên; riêng dòng đầu của phần nhánh mặc định là `T2`.

### 8.3. Gán nhánh

`T1` là nhóm chủ đất. `T2` là các nhánh trực tiếp của cả nhóm chủ đất, nên `ParentNguoiID` để trống.

Khi một dòng được gán `Tk` với `k > 2`, hệ thống tìm ngược lên dòng gần nhất có tầng `T(k-1)` và lưu ID của dòng đó vào `ParentNguoiID`.

Sau lần gán đầu tiên, `ParentNguoiID` là nguồn sự thật. Thứ tự dòng và màu chỉ phục vụ nhập liệu, không được dùng để suy đoán lại nhánh mỗi lần xuất Word.

Không chấp nhận:

- `Tang < 1` hoặc `Tang > SoTang`.
- Một dòng `Tk` với `k > 2` không tìm thấy dòng cha `T(k-1)` ở phía trên.
- Dòng T1 hoặc T2 có `ParentNguoiID`.
- Dòng từ T3 trở đi thiếu `ParentNguoiID`, tự trỏ hoặc trỏ tới người không tồn tại.
- Tầng của người cha không thấp hơn đúng một cấp so với tầng của người con.

Khi đổi tầng của một dòng đã có hậu duệ, hệ thống xác định khối hậu duệ liên tiếp và dịch chuyển cả khối cùng số tầng. Chỉ hỏi xác nhận ở trường hợp hiếm này; chặn nếu bất kỳ dòng nào vượt phạm vi `1..SoTang`.

## 9. Màu tầng

- Chỉ hỗ trợ năm tầng `T1` đến `T5`.
- Mỗi dòng luôn hiện chữ `T1`...`T5`; không phụ thuộc riêng vào khả năng phân biệt màu.
- Dùng màu nền nhạt cho toàn dòng và một vạch màu đậm ở mép trái để cấu trúc dễ quét bằng mắt.
- Không dùng đỏ làm màu tầng; đỏ dành cho lỗi.
- Màu tầng không được xuất sang văn bản Word.

Demo HTML cung cấp nhiều dải màu để lựa chọn. Sau khi người dùng chọn một dải, chỉ giữ một bảng màu trong workbook thật.

## 10. Người nhận đất

- Cột `Nhận đất` là checkbox hiển thị trực tiếp trên mỗi dòng.
- Có thể tích nhiều người.
- `NhanDat = True` tạo nhóm dữ liệu riêng khi xuất Word.
- Người đã chết không được nhận đất trong mô hình MVP. Nếu người dùng nhập ngày chết cho người đang được tích nhận đất, hệ thống phải cảnh báo và bỏ tích sau khi xác nhận.
- Không suy luận người nhận theo tầng hoặc quan hệ.

## 11. Người từ chối

Đây là quyết định nghiệp vụ đã chốt của người dùng:

```text
Từ chối = Tổng người - Người đã chết - Người nhận đất
```

Tương đương theo từng dòng:

```text
DaChet = NgayChet có giá trị
TuChoi = (không DaChet) AND (không NhanDat)
```

Hệ quả:

- Không có checkbox `Từ chối`.
- Người dùng chỉ cần nhập ngày chết và tích người nhận.
- Trạng thái đọc được trên dòng là một trong ba giá trị: `Đã chết`, `Nhận đất`, `Từ chối`.
- Chủ đất đã chết cũng thuộc nhóm `Đã chết`, không thuộc nhóm từ chối.

## 12. Chuẩn bị cho nhóm từ chối

Phiên bản hiện tại chưa có UI chia nhóm từ chối. Tuy nhiên không được thiết kế việc xuất Word như chỉ có một chuỗi danh sách từ chối duy nhất.

Mỗi dòng có trường ẩn `NhomTuChoiID`. Hàm xuất luôn trả về một tập nhóm:

```text
LayCacNhomTuChoi() -> Collection<NhomTuChoi>

NhomTuChoi:
  NhomID
  ParentNguoiID
  DanhSachNguoi
```

Quy tắc MVP:

- Nếu mọi `NhomTuChoiID` đều trống, đưa người từ chối vào nhóm `TC_DEFAULT`.
- API xuất Word vẫn nhận một collection có một phần tử, không nhận một danh sách phẳng.
- Khi phát triển tính năng, có thể gán nhóm theo `ParentNguoiID` — ví dụ những người thuộc nhánh của B cùng từ chối trong một văn bản — hoặc cho người dùng chỉnh `NhomTuChoiID`.
- Việc thêm UI phân nhóm sau này không được làm đổi schema bảng người, công thức `TuChoi` hoặc hợp đồng của hàm xuất.

## 13. Hợp đồng dữ liệu xuất Word

Lớp lập dữ liệu xuất tạo ít nhất các nhóm sau:

| Nhóm | Cách lấy |
| --- | --- |
| `ChuDat` | Dòng có `LaChuDat = True` và có dữ liệu |
| `NguoiDaChet` | Dòng có `NgayChet` |
| `NguoiNhanDat` | Dòng còn sống và `NhanDat = True` |
| `CacNhomTuChoi` | Collection tạo theo mục 12 |
| `CayNhanh` | Nhóm gốc gồm các `LaChuDat`, các dòng T2 trực thuộc nhóm gốc, và dữ liệu `NguoiID`, `Tang`, `ParentNguoiID` từ T3 trở đi |

Template Word không được đọc màu ô, vị trí dòng 1–2 hoặc suy luận nhánh từ STT.

## 14. Validation

### Lỗi chặn xuất

- `SoTang` nằm ngoài `1..5`.
- Có `Tang` lớn hơn `SoTang`.
- Có tầng bị nhảy cấp, ví dụ `T3` mà phía trên chưa có `T2` hợp lệ.
- T1/T2 có `ParentNguoiID`, hoặc dòng từ T3 trở đi thiếu, tự trỏ, trỏ sai người hay không ở tầng liền trước.
- Có chu kỳ trong chuỗi `ParentNguoiID`.
- Người đã chết vẫn có `NhanDat = True`.
- Cùng một `NguoiID` xuất hiện ở nhiều dòng.
- Số người hoặc dữ liệu vượt sức chứa của template được chọn.
- Còn placeholder Word chưa được thay.

### Cảnh báo

- Dòng chủ đất thứ hai để trống: hiểu là hồ sơ một chủ đất, không coi là lỗi.
- Có hai dòng nghi cùng một người do trùng họ tên và ngày sinh.
- Nhánh có người đã chết nhưng chưa có dòng tầng kế tiếp; chỉ cảnh báo để người dùng kiểm tra, không tự tạo người.
- Người từ chối chưa được chia nhóm: dùng `TC_DEFAULT` ở MVP.

Mọi thông báo phải ghi rõ STT và họ tên; khi có thể, cho phép bấm để nhảy tới dòng lỗi.

## 15. Ngoài phạm vi phiên bản này

- Không chọn hoặc mô hình hóa vai trò cha, mẹ, con, vợ/chồng.
- Không tự kết luận hàng thừa kế, thế vị hoặc phần di sản.
- Không có UI chọn 3+ đồng chủ đất.
- Không có UI chia nhóm người từ chối.
- Không vẽ sơ đồ gia đình.
- Không dùng double-click.
- Không sửa workbook, VBA hoặc template Word ở bước đặc tả này.

## 16. Các hướng cũ — lưu rất ngắn để không quay lại

- **Mã đường dẫn `1`, `1.4`, `1.4.4`:** bỏ vì con chung không thuộc riêng nhánh 1 hay 2, đổi cấu trúc phải đánh lại mã cả cây và dễ nhân đôi người.
- **Quan hệ chi tiết `ChaID/MeID/VoChongID` rồi truy vấn theo người chết:** không dùng cho MVP vì buộc luồng “thêm người → chọn vai trò → thêm người”, làm gián đoạn cách nhập hiện tại và mô hình hóa nhiều hơn nhu cầu thực tế.
- **Điểm được giữ:** mỗi người có một `NguoiID`, chỉ nhập một lần; dùng `ParentNguoiID` để nhánh không phụ thuộc màu hay vị trí dòng.

Chỉ xem xét lại hai hướng trên khi phát sinh yêu cầu pháp lý hoặc xuất văn bản mà mô hình `Tang + ParentNguoiID` không biểu diễn được.

## 17. Điều kiện hoàn tất bước kiến trúc

Bước này hoàn tất khi:

- Spec này được dùng làm nguồn quyết định chính cho nhánh thừa kế.
- Demo thể hiện đúng thanh công cụ, nút đổi tầng một lần bấm, người nhận đất và trạng thái từ chối phái sinh.
- Người dùng chọn một dải màu trong demo.

Sau đó mới đồng bộ các spec liên quan và triển khai trên bản sao workbook; không ghi đè template nguồn.
