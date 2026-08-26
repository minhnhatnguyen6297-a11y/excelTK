# Session: Mã nhánh vai trò cho hồ sơ thừa kế

- Ngày: 2026-08-26
- Múi giờ: Asia/Saigon
- Workspace: `D:\TK`
- Repo dự án: `https://github.com/minhnhatnguyen6297-a11y/excelTK.git`

> Bản ghi hội thoại người dùng – trợ lý của phiên làm việc. Log công cụ nội bộ, prompt hệ thống và output trung gian không đưa vào.

---

## Người dùng

Đang viết spec cho dự án: đọc ở góc độ UX người dùng, chia đánh giá làm 2 phần (hồ sơ thừa kế / hồ sơ khác), đi theo flow thật từ khi nhập liệu đến khi bấm Xuất Word (dừng ở nút xuất). Là agent parent, gọi codex và opencode (model claude opus 5) để lấy các quan điểm khác nhau rồi tổng hợp. Lưu ý: làm tối giản, không overengineering — đây là tool excel, không phải app hoàn chỉnh.

## Trợ lý

Đã đọc 3 spec (`approved-minimal-entry-ui.md`, `inheritance-workflow.md`, `excel-restructure.md`) và chạy review song song:

- Codex (`gpt-5.6-sol`): tập tính chạy được của flow từng bước.
- OpenCode (`claude-opus-5` như yêu cầu; có một lượt chạy nhầm sang `sonnet-4.6` do bỏ sót cờ model, kết quả hội tụ nên giữ làm tham chiếu).

**Kết luận chung:** gốc rễ mọi vấn đề là mâu thuẫn giữa `approved-minimal-entry-ui.md` (màn hình chính chỉ 2 vùng nhập + Lưu/Xuất Word) với `inheritance-workflow.md` §3 và `excel-restructure.md` §3/§5 (flow dựa trên các nút không tồn tại trong giao diện đã chốt). Flow hiện tại chạy không hết ở cả 2 loại hồ sơ.

**Phát hiện chính (hội tụ cả 3 góc nhìn):**

1. [Chặn] Không có cơ chế tạo/chọn hồ sơ.
2. [Chặn] Cột `Ben` bị mô tả 3 kiểu khác nhau (ẩn / xám khóa / toggle); vùng Người chỉ có 6 cột nên hồ sơ hai bên không có chỗ gán A/B; `VaiTro` cũng vậy.
3. [Chặn] Các nút thừa kế (`Thiết lập di sản`, `Thêm đồng chủ`, `Thêm người thừa kế`, `Cập nhật mã nhánh`) không tồn tại trong UI đã chốt.
4. [Chặn nghiêm trọng nhất] `tblPhanChiaDiSan` bắt buộc nhập `TaiSanID/NguoiID/DiSanID` — những mã mà spec cấm người dùng gõ và ẩn khỏi giao diện.
5. [Chặn] 13 nhóm kiểm tra trước xuất nhưng không định nghĩa lỗi hiển thị thế nào, có dẫn tới ô sai hay không.
6. [Chặn — riêng opus-5 tìm ra] Mất dữ liệu âm thầm do sức chứa template: mẫu cũ chỉ dùng 10 người / 1 tài sản trong khi nhập không giới hạn; validation chỉ có chiều ngược lại.
7. [Khó chịu] Chủ chính = "dòng đầu tiên" vỡ khi lọc/sắp xếp; `Cập nhật mã nhánh` thủ công dễ quên; không phản hồi sau thao tác quan hệ; `Lưu` chưa được định nghĩa; đổi loại hồ sơ giữa chừng tạo ngõ cụt (còn A/B nhưng cột bị khóa).

**Top 5 cần chốt trước khi triển khai:** chỗ bấm nghiệp vụ thừa kế; một hành vi duy nhất cho `Ben` (+`VaiTro`); mọi bảng quan hệ nhập bằng tên không ID; hình thức báo lỗi trước xuất (danh sách lỗi theo tên người, click nhảy tới ô); mô hình 1 workbook = 1 hồ sơ + tự động hóa (Xuất Word tự Lưu, sinh ID cho mọi dòng mới).

---

## Người dùng

Định hướng mới:

- Không đi sâu vào engine chia thừa kế — đó là lĩnh vực khác, không dành cho Excel.
- Không thêm nhiều chức năng: mỗi hồ sơ dùng 1 bản copy file excel mẫu trong folder riêng. User mở file, xóa dữ liệu cũ, chọn luôn thừa kế hoặc hồ sơ khác -> nhập người, chọn vai trò -> nhập tài sản -> xuất Word. Chỉ đơn giản vậy.
- Ý tưởng mã nhánh: 2 người đầu tiên mặc định là 2 vợ chồng chủ đất (mã 1 - chồng, mã 2 - vợ). 95% hồ sơ là 1 trong 2 hoặc cả 2 chết -> mở nhánh: mẹ đẻ, bố đẻ, con là mã 1.n và 2.n. Con chung áp mã nào? Con đã chết thì tiếp tục nhánh 1.1.n. Engine không vẽ sơ đồ, chỉ suy cây phả hệ từ mã: 1.1 = người thừa kế của 1.
- Có ý định cố định số nhánh ghép cứng vai trò: 1.n hoặc 1.n.m — n/m = 1 bố, = 2 mẹ, = 3 vợ/chồng, từ 4 trở đi là con (đúng ~90% case). Engine nhìn mã là hiểu cây.
- Mục đích: khi văn bản cần "những người thừa kế của ông X là:", viết hàm lấy người thừa kế của 1.n — script tự hiểu mã nhánh, import thông tin nhóm mà không phải chỉ định tay.

Cần thảo luận 2 nhóm việc: (1) cách xuất chính xác nhóm người trong văn bản thừa kế — các nhóm nào, đặc điểm chung, nhận diện nhóm qua mã nhánh, chuyển nhánh trên giao diện thế nào, thiết kế hệ thống mã chuẩn; (2) các hàm cần thiết kế để xuất nhóm người.

## Trợ lý

Đã thiết kế đầy đủ theo hướng này:

**Hệ thống mã nhánh (tách 2 lớp):** mã nhánh = cấu trúc cây; cột `QuanHe` = nhãn chữ cho câu văn Word, không ảnh hưởng cấu trúc. Slot cố định trong mỗi nhánh: `.1` bố đẻ, `.2` mẹ đẻ, `.3` vợ/chồng, `.4+` con.

Giải 3 câu hỏi:

- **Con chung:** chỉ một mã, quy ước đặt dưới nhánh 1 (đảo giá rẻ được). Pháp lý không mất mát: Điều 651 vợ chồng cùng hàng, khối chung vợ chồng chia một lần cho nhóm chung; nhóm gia đình gốc = `HopNhanh("GOC")` = hợp nhánh 1 ∪ 2.
- **Xung đột `.3` với người gốc:** cấm `.3` ở độ sâu 1 (vợ của 1 là người 2 sẵn); từ sâu ≥ 2, `.3` mặc định Không hưởng (vợ người thừa kế chết không phải thế vị).
- **Con đã chết mở nhánh tiếp:** bỏ qua slot 1/2/3 (tổ tiên/vợ chồng đã có trên cây), con bắt đầu từ `.4`: con của `1.4` là `1.4.4`, `1.4.5`...

Quy tắc vàng: mã bất biến (xóa/sửa giữ lỗ số để placeholder sống sót); sắp xếp theo số từng đoạn (`1.9` trước `1.10`). Case lạ 10% (ông bà hàng thứ hai...): nút `Người khác`, ghi `QuanHe` tay.

**Nhóm xuất trong văn bản** = (prefix cấu trúc) × (lọc trạng thái): người để lại di sản (gốc có ngày mất); danh sách hàng thừa kế; đồng thừa kế trực hệ; thế vị (prefix + mọi trung gian đã chết); người từ chối; người tham chiếu; từng cá nhân. Không cần bảng quan hệ, không cần sơ đồ.

**UI chuyển nhánh:** thêm cột chỉ-đọc `Mã nhánh` vào vùng Người (chỉ hồ sơ thừa kế) + cụm nút `[Thêm bố] [Thêm mẹ] [Thêm vợ/chồng] [Thêm con] [Con đã mất – mở nhánh] [Người khác]`. Chuyển nhánh = click dòng rồi bấm nút; sinh mã tất định kèm MsgBox xác nhận. Sửa sai nhánh bằng `DoiTienTo` đổi tiền tố cả cây con.

**Các hàm:** `MaTiepTheo(maGoc, vaiTro)`, `DoiTienTo(old, new)`, `NhomTheoMa(prefix)`, `HopNhanh(prefix|GOC)` — thuật toán lõi rút gọn nhánh chết ~20 dòng, `LocTrangThai`, `NoiDanhSach` ("A, B và C"); UDF `=NguoiThuaKe("1.4")`; placeholder `{{nguoi.<ma>.<truong>}}`, `{{nhom.<ma>.danh_sach}}`, `{{nhom.goc.danh_sach}}`, `{{nhom.<ma>.so_luong}}`.

**Lợi ích phụ:** xóa sổ `tblQuanHeThuaKe` — cây nằm ngay trong `tblNguoi`.

---

## Người dùng

Tạo file lưu toàn bộ session này vào folder dự án, commit tất cả thay đổi và push lên.

## Trợ lý

Ghi bản lưu session này vào `docs/session/SESSION_MA_NHANH_VAI_TRO_2026-08-26.md`, commit toàn bộ (bao gồm việc tái cấu trúc thư mục templates/docs/prototypes từ trước đó chưa commit, cập nhật 3 spec theo hướng mã nhánh vai trò) và push lên `origin/main`.
