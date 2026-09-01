# Kho workbook Excel theo phiên bản

Mỗi phiên bản MVP có một thư mục riêng:

```text
templates/excel/
├── v0.1.0/
├── v0.2.0/
└── v0.2.1/
```

Trong mỗi thư mục:

- `Ho_so_thua_ke_MVP_vX.Y.Z.xlsm`: file thô/sạch để copy thành hồ sơ mới.
- `Ho_so_thua_ke_MVP_vX.Y.Z-TEST.xlsm`: file có dữ liệu kiểm thử gần thực tế.
- `qa-clean-vX.Y.Z.xlsx`: bản không macro để kiểm tra cấu trúc file sạch.
- `qa-test-vX.Y.Z.xlsx`: bản không macro để kiểm tra dữ liệu mẫu.
- `README.md`: ghi nguồn gốc và phạm vi của phiên bản.

Repo là public nên file TEST chỉ dùng **dữ liệu giả lập**, không lưu dữ liệu cá nhân thật của khách hàng.

File `Dữ liệu thừa kế (2).xlsb` là workbook cũ để đối chiếu, không phải một bản phát hành MVP. File này đang có thay đổi riêng của người dùng nên giữ nguyên tại chỗ, không đưa vào commit và không dùng làm nguồn phát triển.

Các file build thử, log, ảnh tạm và bản backup lặp lại nằm trong `output/` và không thuộc kho phát hành.
