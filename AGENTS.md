# Hướng dẫn dự án

- `templates/word/`: mẫu Word có placeholder, chỉ sửa khi có yêu cầu rõ ràng.
- `templates/excel/`: workbook mẫu nguồn; không ghi đè lên bản gốc khi thử nghiệm.
- `docs/specs/`: đặc tả đã thống nhất. `docs/reference/` và `docs/session/` là tài liệu khảo sát/lịch sử.
- `output/`: file Word được sinh ra, không đưa vào Git.
- `OfficeCLI/` là mã nguồn công cụ được clone; agent hãy sử dụng công cụ này để đọc, sửa, tạo các file word, excel, pdf; giữ tách biệt với dữ liệu mẫu.

Khi thay đổi VBA hoặc template, luôn thử xuất ra file Word mới và không đóng các cửa sổ Word có sẵn của người dùng.
