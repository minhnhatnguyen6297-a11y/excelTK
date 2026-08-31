# Bản chốt Excel thừa kế v0.2.1

Ngày chốt: 29/08/2026 21:28. Schema: `2.0.0`. SOT: `docs/specs/inheritance-branch-architecture.md`.

## File dùng

| File | Dùng khi |
| --- | --- |
| `Ho_so_thua_ke_MVP_v0.2.1.xlsm` | **Mẫu sạch.** Copy file này ra folder hồ sơ rồi nhập liệu. Có VBA. |
| `Ho_so_thua_ke_MVP_v0.2.1-TEST.xlsm` | Cùng khung, đã điền dữ liệu giả để kiểm tra Hàng TK / nhận đất / xuất thử. |
| `qa-clean-v0.2.1.xlsx` | Xem cấu trúc sheet, không VBA. |
| `qa-test-v0.2.1.xlsx` | Xem dữ liệu giả, không VBA. |
| `preview-qa-test-v0.2.1.png` | Ảnh chụp bản QA. |

## Đã có / chưa có

Đã có: nhập người theo Hàng TK, ngày `yyyy` / `mm/yyyy` / `dd/mm/yyyy`, nhận đất, kiểm tra, năm nhóm xuất, xuất Word thử bằng template kiểm thử.

Chưa có: sheet tài sản, nối mẫu Word thật (`1. PCDS .docx` và các mẫu khác), workbook hồ sơ chia bên A/B.

Không dùng `templates/excel/Dữ liệu thừa kế (2).xlsb` làm bản phát triển. File đó chỉ là workbook cũ để đối chiếu.
