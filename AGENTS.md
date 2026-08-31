# AGENTS.md - excelTK

Chỉ chứa thẩm quyền, định tuyến nguồn, điều kiện dừng và cổng review. Đặc tả
hành vi nằm trong `docs/specs/`; kế hoạch và việc cần làm nằm trên Linear.

## Thẩm quyền

- Giữ nguyên thay đổi của người dùng và trạng thái bên ngoài. Không sửa, không
  revert, không commit, không publish công việc không liên quan.
- Chỉ thay đổi hành vi trong phạm vi nhiệm vụ người dùng đã cho phép rõ ràng.
- Quy tắc nghiệp vụ, schema workbook, hợp đồng placeholder Word, và hành vi dùng
  chung giữa nhiều module VBA cần phạm vi cho phép bao trùm mọi module bị ảnh hưởng.
- Xung đột giữa nhiệm vụ / issue Linear / workbook thực tế là bằng chứng để báo
  cáo, không phải quyền để tự sửa bất kỳ bên nào. Dừng và hỏi bên nào thắng.
- Tuyên bố hoàn thành phải kèm bằng chứng kiểm chứng mới chạy.
- Không tự viết lại issue hoặc mô tả kế hoạch để hợp thức hóa việc đã làm.

## Ranh giới nhiệm vụ

Trước khi sửa file, xác lập:

```text
LINEAR ISSUE:
SPEC THAM CHIẾU:
GOAL:
HÀNH VI ĐƯỢC PHÉP:
FILE/MODULE DỰ KIẾN:
BIÊN DÙNG CHUNG ĐÃ BIẾT:
BẰNG CHỨNG NGHIỆM THU:
SCOPE: LOCKED
```

Nếu phát sinh nhu cầu chạm vào module khác, sheet khác, mẫu Word khác hoặc quy
tắc nghiệp vụ khác, dừng lại với `SCOPE BREAK REQUEST` nêu rõ phụ thuộc, hành vi
bị ảnh hưởng, và vì sao không thể tiếp tục an toàn. Chỉ người dùng mở rộng phạm vi.

## Nguồn quyết định

Thứ tự ưu tiên:

`AGENTS.md` -> `docs/specs/inheritance-branch-architecture.md` (SOT tổng, đã chốt)
-> issue Linear đang chạy -> mã nguồn VBA và workbook chốt trong
`templates/excel/chot-v*` -> `docs/reference/` và `docs/session/` (khảo sát/lịch
sử, không phải thẩm quyền).

Đặc tả **ở lại trong repo**:

- `docs/specs/inheritance-branch-architecture.md`: SOT tổng, mục §1–§17. Chỉ
  người dùng chốt và sửa. Đây là nguồn quyết định hành vi. §1–§9 đã triển khai
  trong v0.2.1; §10–§14 đã chốt đặc tả nhưng chưa triển khai (đích v0.3.0).
- `docs/specs/noi-dung-ngoai-sot.md`: nội dung chờ quyết định. Không triển khai
  bất kỳ mục nào trong file này khi chưa được đưa vào SOT.

Kế hoạch triển khai **không** lưu trong repo. `docs/plans/` đã bị xóa và chuyển
sang Linear; không tạo lại thư mục này và không viết file plan markdown mới.

Hành vi nghiệp vụ mới hoặc thay đổi cần mục SOT tương ứng đã chốt. Mục còn ở
trạng thái đích tương lai, mơ hồ, hoặc xung đột với workbook thực tế thì chặn
triển khai đến khi người dùng quyết. Chỉ người dùng chốt đặc tả nghiệp vụ.

## Linear

Linear chứa việc cần làm, không chứa đặc tả. Quy ước:

| Loại nội dung | Nơi đặt |
| --- | --- |
| Việc cần giải quyết | Issue trong project `excelTK`, một issue một việc |
| Nhóm việc theo pha | Parent issue + sub-issue |
| Tiến độ, quyết định, bằng chứng kiểm chứng | Comment trên issue |
| Đặc tả hành vi | `docs/specs/`, không phải Linear |

Workspace: `minhnotary`. Team: `Minh` (prefix `MIN`).
Project: `excelTK` (`https://linear.app/minhnotary/project/exceltk-ddd609aa414d`).

Backlog hiện tại:

- `MIN-5` — v0.3.0: tài sản phiếu dọc, thông tin phụ, xuất Word thật. Sub-issue
  `MIN-7`..`MIN-19`. `MIN-7` (inventory placeholder Word) chặn `MIN-16`, `MIN-17`,
  `MIN-18`, `MIN-20`.
- `MIN-6` — workbook hồ sơ hai bên A/B, file riêng. Sub-issue `MIN-20`, `MIN-21`.
  Làm sau khi `MIN-5` chốt.

Nguyên tắc làm việc:

- Mỗi issue phải nêu rõ mục SOT mà nó triển khai (ví dụ `SOT: §10.3`). Issue không
  trỏ được về SOT thì không triển khai, báo người dùng.
- Đầu nhiệm vụ: đọc issue và mục SOT nó tham chiếu. Không suy luận đặc tả từ code.
- Trong nhiệm vụ: ghi tiến độ và bằng chứng kiểm chứng bằng comment trên issue.
  Không tạo file markdown tổng kết trong repo.
- Không tự tạo issue mới cho việc ngoài phạm vi; báo cáo và để người dùng quyết.
- Không sửa `docs/specs/` để khớp với issue. Xung đột giữa issue và SOT là điều
  kiện dừng.

### Kết nối Linear MCP theo từng agent

Credential Linear MCP lưu riêng cho từng CLI, **không dùng chung**. Đổi agent thì
phải cấu hình và chạy lại OAuth một lần cho agent đó. Endpoint chung:
`https://mcp.linear.app/mcp` (read-write) hoặc `https://mcp.linear.app/mcp/readonly`.

| Agent | Cấu hình | Đăng nhập |
| --- | --- | --- |
| Codex | `~/.codex/config.toml` -> `[mcp_servers.linear] url = "https://mcp.linear.app/mcp"` | `codex mcp login linear` |
| OpenCode | `~/.config/opencode/opencode.json` -> `mcp.servers.linear` kiểu `remote` | Khởi động lại OpenCode, chạy OAuth qua giao diện quản lý MCP |
| Claude Code | `claude mcp add --transport http linear-server https://mcp.linear.app/mcp` | `/mcp` trong session |
| Gemini CLI | `~/.gemini/settings.json` -> `mcpServers.linear`, `httpUrl` là endpoint trên | OAuth mở browser lần đầu gọi tool |
| Cursor | Cài từ MCP directory hoặc `mcp.json` với `url` là endpoint trên | OAuth trong Cursor Settings > MCP |
| Amp | Thêm MCP server `linear` với `npx -y mcp-remote https://mcp.linear.app/mcp` | OAuth mở browser lần đầu |

Client không hỗ trợ remote MCP thì dùng cầu nối `npx -y mcp-remote <endpoint>`.
Nhiều workspace Linear thì tách `MCP_REMOTE_CONFIG_DIR` cho từng workspace.

Nếu Linear MCP chưa kết nối, dừng và báo người dùng. Không tự chuyển sang lưu
plan/spec vào file trong repo để đi tiếp.

## Bố cục repo

- `templates/word/`: mẫu Word có placeholder. Chỉ sửa khi có yêu cầu rõ ràng.
- `templates/excel/`: workbook mẫu nguồn và bản chốt `chot-v*`. Không ghi đè bản
  gốc khi thử nghiệm. `Dữ liệu thừa kế (2).xlsb` đang dirty so với Git: không
  commit, không dùng làm nguồn phát triển.
- `src/vba/`: mã nguồn VBA, là nguồn duy nhất của code. Không sửa VBA trực tiếp
  trong workbook rồi bỏ qua file `.bas` / `.cls`.
- `scripts/`: `build-inheritance-workbook.ps1` dựng workbook từ `src/vba`;
  `create-qa-sheet-view.ps1` xuất ảnh QA.
- `output/`: sản phẩm sinh ra, đã gitignore. Mọi workbook và Word thử nghiệm phải
  ở đây.
- `prototypes/`: HTML demo, không phải nguồn hành vi.
- `docs/specs/`: SOT và nội dung chờ quyết định. Thẩm quyền hành vi.
- `docs/reference/`, `docs/session/`: khảo sát và lịch sử. Tham khảo, không thẩm quyền.
- `OfficeCLI/`: công cụ clone để đọc/sửa/tạo file Word, Excel, PDF. Giữ tách biệt
  với dữ liệu mẫu; dùng công cụ này thay vì tự viết parser.

## Ràng buộc kỹ thuật

- Không đóng cửa sổ Word hoặc Excel người dùng đang mở. Excel do agent tạo:
  `Visible=false`, tắt macro khi chỉ đọc, `Quit` đúng phiên mình mở.
- Xuất Word bằng `CreateObject`, không `GetObject`. Thay placeholder qua Range,
  không `Selection.Find`.
- Đổi VBA hoặc template thì luôn xuất thử file Word mới vào `output/word-test/`.
- Workbook đang mở và có dữ liệu chưa lưu: dừng, yêu cầu người dùng lưu/đóng trước.
- Literal tiếng Việt trong VBA đi qua `UnicodeText(...)` như hàm
  `Convert-VbaUnicodeLiterals` trong script dựng workbook.
- Bảo vệ sheet: `UserInterfaceOnly:=True`, khôi phục ở `Workbook_Open`.
  Không `ThisWorkbook.Save` trong `BeforeClose`.
- Ngày của người: ô hiển thị format `@`, engine đọc cột `...Tinh`.
- Bản chốt chỉ copy vào `templates/excel/chot-v*` sau khi kiểm chứng xanh.

## Cổng review

Review mỗi task hoàn thành trước khi làm việc phụ thuộc hoặc commit. Người review
phải có context mới; người triển khai không tự duyệt việc của mình. Review yêu cầu
gốc, mục SOT liên quan, diff base-to-head, module VBA bị ảnh hưởng, và output
kiểm chứng thật, không phải bản tự tổng kết.

```text
SCOPE: PASS/FAIL — thiếu, thừa, hoặc hiểu sai hành vi
SPEC: PASS/FAIL — khớp hành vi chuẩn trong mục SOT
SHARED IMPACT: PASS/FAIL — module và placeholder bị ảnh hưởng đã kiểm
TEST EVIDENCE: SUFFICIENT/INSUFFICIENT — xuất workbook/Word thật cho từng module
VERDICT: APPROVE/BLOCK
```

Hành vi chưa được phê duyệt, ảnh hưởng dùng chung chưa xử lý, hoặc xung đột
spec/thực tế đều chặn task tiếp theo và chặn commit kể cả khi test xanh.

## Hoàn tất

Ưu tiên thay đổi nhỏ nhất giữ nguyên hành vi và tận dụng code có sẵn. Chạy lại
`scripts/build-inheritance-workbook.ps1` và xuất Word thử cho thay đổi không tầm
thường. Không mô tả kiểm tra cục bộ là đã chạy toàn bộ. Báo cáo:

```text
LINEAR ISSUE:
SPEC THAM CHIẾU:
FILE ĐÃ SỬA:
SCOPE VERDICT:
HÀNH VI DÙNG CHUNG: YES/NO; MODULE ẢNH HƯỞNG:
KIỂM CHỨNG ĐÃ CHẠY:
RỦI RO CÒN LẠI:
```

Chi tiết quy tắc nghiệp vụ thuộc `docs/specs/inheritance-branch-architecture.md`,
không thuộc file này.
