# Tài sản, Word thật, hồ sơ A/B — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Từ bản chốt thừa kế v0.2.1, thêm sheet tài sản + dữ liệu hồ sơ, nối mẫu Word thật, rồi tạo workbook riêng cho hồ sơ chia hai bên A/B.

**Architecture:** Giữ workbook thừa kế một-hồ-sơ-một-file, mô hình `HangTK + ParentNguoiID` không đổi. Tài sản là bảng `tblTaiSan` trên sheet riêng, ID ổn định `TS001`. Xuất Word vẫn `CreateObject`, thay placeholder bằng Range, không đụng Word đang mở. Hồ sơ A/B là **workbook khác**, không nhét cột `Ben` vào file thừa kế.

**Tech Stack:** Excel VBA (`.xlsm`), OfficeCLI, PowerShell COM, Word COM `CreateObject`.

**Spec:**
- SOT thừa kế: `docs/specs/inheritance-branch-architecture.md` (đã chốt, schema `2.0.0`)
- Hàng đợi ngoài SOT: `docs/specs/noi-dung-ngoai-sot.md`
- Workbook cũ / placeholder Word: `docs/reference/TONGQUANDUAN.md`
- Bản chốt Excel: `templates/excel/chot-v0.2.1/Ho_so_thua_ke_MVP_v0.2.1.xlsm`
- Script dựng file: `scripts/build-inheritance-workbook.ps1`

## Global Constraints

- Không ghi đè `templates/excel/Dữ liệu thừa kế (2).xlsb`. File này đang dirty trên disk so với Git; không commit, không dùng làm nguồn phát triển.
- Không đóng cửa sổ Word/Excel người dùng đang mở. Excel ẩn: `Visible=false`, tắt macro khi chỉ đọc, `Quit` đúng phiên agent tạo.
- Mọi workbook thử nằm trong `output/` (đã gitignore). Bản chốt copy vào `templates/excel/chot-v*` sau khi test xanh.
- VBA Unicode: literal tiếng Việt phải qua `UnicodeText("....")` như `scripts/build-inheritance-workbook.ps1` hàm `Convert-VbaUnicodeLiterals`.
- Bảo vệ sheet: mật khẩu `HoSoTK_MVP_2026`, `UserInterfaceOnly:=True`, khôi phục mỗi lần `Workbook_Open`.
- Ngày người: ô hiển thị format `@`, engine dùng cột `...Tinh`. Ngày cấp sổ tài sản cùng quy ước `yyyy` / `mm/yyyy` / `dd/mm/yyyy`.
- Không mô hình hóa vai trò bố/mẹ/con/vợ chồng. Không `tblDiSan` / `tblSoHuuDiSan` / `tblPhanChiaDiSan` / tỷ lệ phần trăm ở pha này.
- Không `GetObject` Word. Không `Selection.Find`. Không `ThisWorkbook.Save` trong `BeforeClose`.
- Sức chứa mẫu Word phải chặn xuất, không nuốt dữ liệu thừa.

## Quyết định đã khóa (không hỏi lại lúc code)

1. **Sheet tài sản = một dòng một tài sản** (`tblTaiSan`), không phiếu dọc kiểu cột C/D/E của file cũ. Thêm tài sản = thêm dòng. Lý do: ListObject + VBA ID, lọc, xuất hậu tố ` / 2 / 3` ổn định hơn phiếu dọc.
2. **Dữ liệu hồ sơ** (niêm yết, số công chứng, người ủy quyền) nằm sheet `HoSo`, không nhét lên `NhapLieu`.
3. **Danh mục** copy từ sheet `Tables` của workbook cũ sang sheet `DanhMuc` workbook mới; tra cứu loại giấy tờ / nơi cấp / nhãn địa chỉ / loại sổ / loại đất / hình thức sử dụng / cơ quan cấp.
4. **Mẫu Word thừa kế đầu tiên:** `templates/word/1. PCDS .docx`. Các mẫu thừa kế khác (`1. Khai nhận...`, `1.1. Từ chối...`, `VP_Phân chia...`) chỉ sau khi PCDS xanh.
5. **Ánh xạ người PCDS:** slot `[Tên 1]` / `[Tên 2]` = `ChuDat` theo `STTNhap`. Slot `[Tên 3]` trở đi = người còn lại theo `STTNhap` (không phải chủ đất). Câu chữ “người 3 nhận đất” của mẫu cũ **không còn đúng** với checkbox `NhanDat`; Task 6 viết lại đoạn thỏa thuận dùng `{{nguoi_nhan_dat.danh_sach}}` và `{{nhom_tu_choi.TC_DEFAULT.danh_sach}}`, giữ bảng nhân thân đánh số.
6. **Tài sản Word:** tài sản 1 không hậu tố (`[Serial]`); tài sản 2/3 hậu tố ` 2` / ` 3`. PCDS hiện chỉ có slot tài sản 1 — nếu `tblTaiSan` có >1 dòng có dữ liệu thì **chặn xuất** với mẫu đó, trừ khi mẫu đã có đủ hậu tố.
7. **Workbook A/B** là file `.xlsm` riêng, script `scripts/build-ab-workbook.ps1`, không dùng `HangTK`. Cột `Ben` bấm một lần `A → B → A`. Hồ sơ thừa kế không có cột `Ben`.

## File map

| File | Trách nhiệm |
| --- | --- |
| `src/vba/modCommon.bas` | Hằng số sheet/cột, `PeopleTable`, thêm `AssetsTable`, `NextAssetId` |
| `src/vba/modTaiSan.bas` | Sinh `TaiSanID`, thêm/xóa dòng tài sản, refresh STT |
| `src/vba/modGiayTo.bas` | Loại CC / nơi cấp / nhãn địa chỉ từ `NgayCapTinh` + `DanhMuc` |
| `src/vba/modHoSo.bas` | Đọc ô hồ sơ (niêm yết, số công chứng, ủy quyền) |
| `src/vba/modExportData.bas` | Thêm nhóm `TaiSan`, `HoSo`; map slot Word `[Tên n]` |
| `src/vba/modValidation.bas` | Sức chứa mẫu, tài sản bắt buộc, ngày cấp sổ |
| `src/vba/modWordExport.bas` | Dictionary placeholder cũ + mới; `CreateObject` Word |
| `src/vba/modNguoi.bas` | Không đổi mô hình nhánh; gọi `RefreshIdCardFields` |
| `scripts/build-inheritance-workbook.ps1` | Dựng v0.3: thêm sheet `TaiSan`, `HoSo`, `DanhMuc` |
| `scripts/inspect-word-placeholders.ps1` | Liệt kê placeholder mọi mẫu Word |
| `scripts/build-ab-workbook.ps1` | Dựng workbook A/B |
| `src/vba-ab/*` | VBA riêng workbook A/B (không trộn `HangTK`) |
| `templates/word/1. PCDS .docx` | Chỉ sửa khi Task 6 yêu cầu rõ; luôn xuất file mới khi thử |
| `templates/excel/chot-v0.3.0/` | Snapshot sau pha tài sản+PCDS |
| `templates/excel/chot-ab-v0.1.0/` | Snapshot workbook A/B |
| `docs/specs/inheritance-branch-architecture.md` | Chỉ bổ sung § tài sản / hợp đồng Word nếu user đã chốt trong plan này |
| `docs/specs/ab-transfer-workbook.md` | Spec ngắn workbook A/B, tạo ở Task 10 |

Pha 1 (Task 1–7) ra workbook thừa kế v0.3 dùng được với PCDS. Pha 2 (Task 8–9) ra workbook A/B độc lập. Task 10 nối thêm mẫu thừa kế còn lại, chỉ sau khi PCDS đã chốt. Có thể dừng sau pha 1.

---

### Task 1: Inventory placeholder Word

**Files:**
- Create: `scripts/inspect-word-placeholders.ps1`
- Create: `docs/reference/word-placeholders-2026-08-31.md` (output của script, commit)
- Test: chạy script, đối chiếu `docs/reference/TONGQUANDUAN.md` §7.3

**Interfaces:**
- Consumes: `templates/word/*.docx` (bỏ `.doc` nếu OfficeCLI không đọc; với `.doc` mở Word ẩn chỉ-đọc)
- Produces: bảng `{file, placeholder, count}` ; danh sách mẫu thừa kế vs chuyển nhượng

- [ ] **Step 1: Viết script OfficeCLI**

```powershell
# scripts/inspect-word-placeholders.ps1
$ErrorActionPreference = "Stop"
$cli = Join-Path $env:LOCALAPPDATA "OfficeCLI\officecli.exe"
$root = Split-Path -Parent $PSScriptRoot
$wordDir = Join-Path $root "templates\word"
$out = Join-Path $root "docs\reference\word-placeholders-2026-08-31.md"
$files = Get-ChildItem -LiteralPath $wordDir -File | Where-Object { $_.Extension -in ".docx", ".doc" }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("# Placeholder mẫu Word — quét $(Get-Date -Format 'yyyy-MM-dd')")
[void]$sb.AppendLine("")
foreach ($f in $files) {
    $text = & $cli view $f.FullName text
    $matches = [regex]::Matches($text, '\[([^\[\]]+)\]|\{\{([^}]+)\}\}')
    $groups = $matches | ForEach-Object {
        if ($_.Groups[1].Success) { "[" + $_.Groups[1].Value + "]" } else { "{{" + $_.Groups[2].Value + "}}" }
    } | Group-Object | Sort-Object Name
    [void]$sb.AppendLine("## $($f.Name)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("- Số placeholder khác nhau: $($groups.Count)")
    [void]$sb.AppendLine("- Tổng lần xuất hiện: $($matches.Count)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("| Placeholder | Số lần |")
    [void]$sb.AppendLine("| --- | ---: |")
    foreach ($g in $groups) {
        [void]$sb.AppendLine("| `$($g.Name)` | $($g.Count) |")
    }
    [void]$sb.AppendLine("")
}
Set-Content -LiteralPath $out -Value $sb.ToString() -Encoding UTF8
Write-Output $out
```

Trong vòng lặp, escape markdown của tên placeholder (đừng bọc trong backtick bằng string `'| `$($g.Name)` |'` nếu PowerShell ăn nhầm). Dùng:

```powershell
[void]$sb.AppendLine("| ``$($g.Name)`` | $($g.Count) |")
```

- [ ] **Step 2: Chạy script**

```powershell
powershell -NoProfile -File scripts/inspect-word-placeholders.ps1
```

Expected: file `docs/reference/word-placeholders-2026-08-31.md` tồn tại; `1. PCDS .docx` có `[Tên 1]`…`[Tên 10]`, `[Serial]`, `[Niêm Yết]`.

- [ ] **Step 3: Phân loại mẫu trong cùng file markdown**

Thêm mục cuối:

```markdown
## Phân loại dùng cho kế hoạch

- Thừa kế / PCDS: `1. PCDS .docx`, `1. Khai nhận di sản Vũ Dương.docx`, `1.1. Từ chối di sản_ Ý Yên.docx`, `VP_Phân chia di sản.docx`, `VP_Từ chối di sản.docx`
- Chuyển nhượng / A/B: `2. DK18, thuế  1.docx`, `2. DK18, thuế  2.docx`, `2. DK18, thuế  3.docx`
- Ủy quyền: `3. UQ.docx`
- Khác: để sau
```

Chỉnh danh sách cho khớp file thật vừa quét. Ghi rõ mẫu nào chỉ có tài sản 1, mẫu nào có hậu tố 2/3.

- [ ] **Step 4: Commit**

```bash
git add scripts/inspect-word-placeholders.ps1 docs/reference/word-placeholders-2026-08-31.md
git commit -m "docs: inventory Word placeholders for remaining templates"
```

---

### Task 2: Hằng số, ID tài sản, sheet `TaiSan` trong script dựng

**Files:**
- Modify: `src/vba/modCommon.bas`
- Create: `src/vba/modTaiSan.bas`
- Modify: `scripts/build-inheritance-workbook.ps1` (param `$Version = "0.3.0"`, thêm sheet)

**Interfaces:**
- Consumes: `PeopleTable()`, `PROTECTION_PASSWORD`, `UnicodeText`
- Produces:
  - `Public Const SHEET_ASSETS As String = "TaiSan"`
  - `Public Const SHEET_CASE As String = "HoSo"`
  - `Public Const SHEET_CATALOG As String = "DanhMuc"`
  - `Public Const TABLE_ASSETS As String = "tblTaiSan"`
  - `Public Function AssetsTable() As ListObject`
  - `Public Function AssetCell(ByVal rowIndex As Long, ByVal headerName As String) As Range`
  - `Public Function AssetHasData(ByVal rowIndex As Long) As Boolean` — TRUE nếu `LoaiSo` hoặc `Serial` hoặc `SoThua` hoặc `DiaChiDat` khác rỗng
  - `Public Function NextAssetId() As String` — `TS` + 3 chữ số, counter name `TaiSanIDTiepTheo`
  - `Public Sub RefreshAllAssets()`

Cột `tblTaiSan` (header đúng tên này):

```text
STTTaiSan, TaiSanID, LoaiSo, Serial, SoVaoSo, SoThua, SoTo, DiaChiDat,
DienTich, HinhThucSuDung, LoaiDat, ThoiHan, ONT, CLN, NTS, LUC,
NguonGoc, NgayCapSo, NgayCapSoGoc, NgayCapSoTinh, CoQuanCapSo, GhiChu
```

- [ ] **Step 1: Thêm hằng số và helper vào `modCommon.bas` ngay dưới `TABLE_PEOPLE`**

```vb
Public Const SHEET_ASSETS As String = "TaiSan"
Public Const SHEET_CASE As String = "HoSo"
Public Const SHEET_CATALOG As String = "DanhMuc"
Public Const TABLE_ASSETS As String = "tblTaiSan"

Public Function AssetsTable() As ListObject
    Set AssetsTable = ThisWorkbook.Worksheets(SHEET_ASSETS).ListObjects(TABLE_ASSETS)
End Function

Public Function AssetColumnIndex(ByVal headerName As String) As Long
    AssetColumnIndex = AssetsTable.ListColumns(headerName).Index
End Function

Public Function AssetCell(ByVal rowIndex As Long, ByVal headerName As String) As Range
    Set AssetCell = AssetsTable.DataBodyRange.Cells(rowIndex, AssetColumnIndex(headerName))
End Function

Public Function AssetHasData(ByVal rowIndex As Long) As Boolean
    AssetHasData = (Len(Trim$(CStr(AssetCell(rowIndex, "LoaiSo").Value2))) > 0) _
                Or (Len(Trim$(CStr(AssetCell(rowIndex, "Serial").Value2))) > 0) _
                Or (Len(Trim$(CStr(AssetCell(rowIndex, "SoThua").Value2))) > 0) _
                Or (Len(Trim$(CStr(AssetCell(rowIndex, "DiaChiDat").Value2))) > 0)
End Function

Public Function NextAssetId() As String
    Dim counterCell As Range
    Dim nextNumber As Long
    Set counterCell = ThisWorkbook.Names("TaiSanIDTiepTheo").RefersToRange
    nextNumber = SafeLong(counterCell.Value2, 1)
    If nextNumber < 1 Then nextNumber = 1
    NextAssetId = "TS" & Format$(nextNumber, "000")
    counterCell.Value2 = nextNumber + 1
End Function
```

- [ ] **Step 2: Tạo `modTaiSan.bas`**

```vb
Attribute VB_Name = "modTaiSan"
Option Explicit

Public Sub RefreshAllAssets()
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim maxStt As Long
    Dim currentStt As Long
    Dim assetId As String

    Set lo = AssetsTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If AssetHasData(rowIndex) Then
            currentStt = SafeLong(AssetCell(rowIndex, "STTTaiSan").Value2, 0)
            If currentStt > maxStt Then maxStt = currentStt
        End If
    Next rowIndex

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If AssetHasData(rowIndex) Then
            assetId = Trim$(CStr(AssetCell(rowIndex, "TaiSanID").Value2))
            If Len(assetId) = 0 Then AssetCell(rowIndex, "TaiSanID").Value2 = NextAssetId()
            currentStt = SafeLong(AssetCell(rowIndex, "STTTaiSan").Value2, 0)
            If currentStt <= 0 Then
                maxStt = maxStt + 1
                AssetCell(rowIndex, "STTTaiSan").Value2 = maxStt
            End If
            NormalizeAssetDate rowIndex
        Else
            AssetCell(rowIndex, "STTTaiSan").ClearContents
            AssetCell(rowIndex, "TaiSanID").ClearContents
        End If
    Next rowIndex
End Sub
```

`NormalizeAssetDate` copy logic `modNgayThang.bas` cho đúng một cột `NgayCapSo` / `NgayCapSoGoc` / `NgayCapSoTinh`. Tách hàm dùng chung trong `modNgayThang.bas`:

```vb
Public Sub NormalizeDateTriplet(ByVal displayCell As Range, _
                                ByVal rawCell As Range, _
                                ByVal calcCell As Range)
```

Rồi `NormalizeChangedDateCells` và `NormalizeAssetDate` cùng gọi hàm này. Không copy-paste parser.

- [ ] **Step 3: Sửa `scripts/build-inheritance-workbook.ps1`**

- `$Version = "0.3.0"`
- Sau khi tạo `XuatAn`, thêm:

```powershell
Invoke-OfficeCli @("add", $baseWorkbook, "/", "--type", "sheet", "--prop", "name=TaiSan")
Invoke-OfficeCli @("add", $baseWorkbook, "/", "--type", "sheet", "--prop", "name=HoSo")
Invoke-OfficeCli @("add", $baseWorkbook, "/", "--type", "sheet", "--prop", "name=DanhMuc")
```

- Header `TaiSan` hàng 8, table `tblTaiSan` phạm vi 10 dòng dữ liệu (`A8:V18` — đếm đúng số cột).
- `CauHinh!B10` = `TaiSanIDTiepTheo` mặc định 1; Name `TaiSanIDTiepTheo`.
- Import `modTaiSan`.
- `ThisWorkbook_Open` gọi thêm `RefreshAllAssets`.
- Khóa sheet: vùng nhập tài sản `C9:V18` trừ cột `STTTaiSan`/`TaiSanID` (cột A–B khóa).
- Format `NgayCapSo` là `@`.

- [ ] **Step 4: Chạy dựng file**

```powershell
powershell -NoProfile -File scripts/build-inheritance-workbook.ps1
```

Expected: `output/workbook-dev/Ho_so_thua_ke_MVP_v0.3.0.xlsm` mở được; có sheet `TaiSan`; SHA-256 của `templates/excel/Dữ liệu thừa kế (2).xlsb` **không đổi so với HEAD** (nếu file working tree đang dirty, `git restore -- "templates/excel/Dữ liệu thừa kế (2).xlsb"` trước khi chạy, hoặc so với hash đã ghi trong log lần build v0.2.1).

Assert trong script sau `RefreshAllAssets` trên bản TEST:

```powershell
# điền 1 tài sản giả rồi
$id = [string]$testAssets.Range("B9").Value2
if ($id -notmatch '^TS\d{3}$') { throw "TaiSanID không sinh đúng: $id" }
```

- [ ] **Step 5: Commit**

```bash
git add src/vba/modCommon.bas src/vba/modTaiSan.bas src/vba/modNgayThang.bas src/vba/ThisWorkbook.cls scripts/build-inheritance-workbook.ps1
git commit -m "feat(excel): add tblTaiSan and asset IDs to inheritance workbook"
```

---

### Task 3: Sheet `HoSo` + `DanhMuc`

**Files:**
- Create: `src/vba/modHoSo.bas`
- Create: `src/vba/modGiayTo.bas`
- Modify: `scripts/build-inheritance-workbook.ps1`
- Modify: `src/vba/modNguoi.bas` — sau `UpdatePersonStatus` gọi làm mới loại giấy tờ **chỉ dòng đổi ngày cấp**

**Interfaces:**
- Produces:
  - `Public Function CaseValue(ByVal fieldName As String) As String`
  - `Public Function IdCardType(ByVal issuedCalc As Variant) As String` — trước 01/07/2024 → `Căn cước công dân`, từ ngày đó → `Căn cước`, trống ngày cấp → chuỗi rỗng
  - `Public Function IdCardIssuer(ByVal cardType As String) As String`
  - `Public Function AddressLabel(ByVal cardType As String) As String` — `Thường trú tại` / `Cư trú tại`

Ô `HoSo` (cột A nhãn, cột B giá trị, Name workbook):

| Name | Ô | Placeholder Word |
| --- | --- | --- |
| `NiemYet` | `HoSo!B3` | `[Niêm Yết]` |
| `SoCongChung` | `HoSo!B4` | `[Số công chứng]` |
| `NguoiUyQuyen` | `HoSo!B5` | `[Người ủy quyền]` |
| `NguoiUyQuyen2` | `HoSo!B6` | `[Người ủy quyền2]` |
| `MauWordDangChon` | `HoSo!B7` | (đường dẫn file) |
| `SucChuaNguoiMau` | `HoSo!B8` | số, PCDS = 10 |
| `SucChuaTaiSanMau` | `HoSo!B9` | số, PCDS = 1 |

`DanhMuc` layout tối thiểu (copy giá trị từ workbook cũ bằng Excel ẩn, macro tắt, không lưu `.xlsb`):

```text
A1 LoaiCC | B1 NoiCap | C1 NhanDiaChi
A2 Căn cước công dân | ... | Thường trú tại
A3 Căn cước | ... | Cư trú tại

E1 LoaiSo
E2.. danh sách GCN

G1 LoaiDat
G2.. 

I1 HinhThucSuDung
I2 Sử dụng riêng
I3 Sử dụng chung

K1 CoQuanCapSo
```

Đọc workbook cũ:

```powershell
$excel.AutomationSecurity = 3
$src = $excel.Workbooks.Open($sourceWorkbook, 0, $true)
# copy Values only từ sheet Tables sang DanhMuc
$src.Close($false)
```

- [ ] **Step 1: `modGiayTo.bas`**

```vb
Public Function IdCardType(ByVal issuedCalc As Variant) As String
    If Not IsDate(issuedCalc) Then Exit Function
    If CDate(issuedCalc) < DateSerial(2024, 7, 1) Then
        IdCardType = UnicodeText("...") ' Căn cước công dân
    Else
        IdCardType = UnicodeText("...") ' Căn cước
    End If
End Function
```

Viết literal tiếng Việt bình thường trong `.bas`; script convert lúc import.

Tra cứu nơi cấp: `DanhMuc!A2:C10` vòng `For` so khớp cột A.

- [ ] **Step 2: Cột ẩn người**

Thêm vào `tblNguoi` (kéo table tới cột W): `LoaiCC`, `NoiCapCC`, `NhanDiaChi`. VBA ghi, user không gõ. Ẩn cùng `J:W`.

Trong `HandlePeopleChange`, nếu `Intersect` với `NgayCap` thì chỉ update 3 cột này cho dòng đó — không `RefreshAllPeople`.

- [ ] **Step 3: Test trong build script**

Điền `NgayCap` người 1 = `2021` → `LoaiCC` chứa `Căn cước công dân`.  
Điền `15/07/2024` → `Căn cước`.

- [ ] **Step 4: Commit**

```bash
git add src/vba/modHoSo.bas src/vba/modGiayTo.bas src/vba/modNguoi.bas src/vba/modCommon.bas scripts/build-inheritance-workbook.ps1
git commit -m "feat(excel): add HoSo fields and ID-card lookup catalogs"
```

---

### Task 4: Validation sức chứa + tài sản

**Files:**
- Modify: `src/vba/modValidation.bas`

**Interfaces:**
- Consumes: `AssetHasData`, `CaseValue`/`SucChuaNguoiMau`/`SucChuaTaiSanMau`
- Produces: mã lỗi mới `ASSET_REQUIRED`, `TEMPLATE_PEOPLE_CAP`, `TEMPLATE_ASSET_CAP`, `ASSET_DATE_INVALID`

- [ ] **Step 1: Thêm đếm và chặn trong `ValidateWorkbook` sau vòng người**

```vb
Dim assetCount As Long
Dim peopleCap As Long
Dim assetCap As Long

peopleCap = SafeLong(ThisWorkbook.Names("SucChuaNguoiMau").RefersToRange.Value2, 10)
assetCap = SafeLong(ThisWorkbook.Names("SucChuaTaiSanMau").RefersToRange.Value2, 1)

If personCount > peopleCap Then
    AddIssue "Lỗi", "TEMPLATE_PEOPLE_CAP", 0, vbNullString, _
             "Mẫu Word chỉ chứa " & CStr(peopleCap) & " người, hồ sơ đang có " & CStr(personCount) & ".", _
             Nothing
End If

assetCount = 0
If Not AssetsTable.DataBodyRange Is Nothing Then
    For rowIndex = 1 To AssetsTable.DataBodyRange.Rows.Count
        If AssetHasData(rowIndex) Then
            assetCount = assetCount + 1
            ' ngày cấp sổ invalid: tái sử dụng parser, nếu display có text mà Tinh trống → lỗi
        End If
    Next rowIndex
End If

If assetCount = 0 Then
    AddIssue "Lỗi", "ASSET_REQUIRED", 0, vbNullString, _
             "Chưa có tài sản nào. Nhập ít nhất Serial, số thửa hoặc địa chỉ đất.", _
             ThisWorkbook.Worksheets(SHEET_ASSETS).Range("C9")
End If
If assetCount > assetCap Then
    AddIssue "Lỗi", "TEMPLATE_ASSET_CAP", 0, vbNullString, _
             "Mẫu Word chỉ chứa " & CStr(assetCap) & " tài sản, hồ sơ đang có " & CStr(assetCount) & ".", _
             Nothing
End If
```

- [ ] **Step 2: Test build**

Bản TEST v0.3: 6 người + 1 tài sản → `ValidateWorkbook=True`.  
Tạm ghi `SucChuaNguoiMau=5` → `False` mã `TEMPLATE_PEOPLE_CAP`. Khôi phục 10.  
Xóa tài sản → `ASSET_REQUIRED`. Điền lại 1 tài sản.

- [ ] **Step 3: Commit**

```bash
git add src/vba/modValidation.bas
git commit -m "feat(excel): block export when people or assets exceed template capacity"
```

---

### Task 5: Lớp xuất `TaiSan` + `HoSo` + slot `[Tên n]`

**Files:**
- Modify: `src/vba/modExportData.bas`
- Modify: `src/vba/modWordExport.bas`

**Interfaces:**
- Produces:
  - `Public Function BuildPlaceholderMap() As Object` — `Scripting.Dictionary` key = placeholder đúng chữ, value = text thay
  - `Public Function WordSlotPeople() As Collection` — item là Dictionary một người theo thứ tự slot 1..n
  - `BuildExportData` ghi thêm section `TaiSan` và `HoSo` trên `XuatAn`

Quy tắc slot:

```text
slots 1..2  = chủ đất theo STTNhap (thiếu thì slot trống)
slots 3..N  = mọi người còn lại có dữ liệu, không phải chủ đất, theo STTNhap
```

Mỗi người trong map:

```text
[Tên n], [Năm sinh n], [CCCD n], [Ngày cấp n], [Địa chỉ n],
[Loại CC n], [Nơi cấp CC n], [Thường trú n]
```

n=1 và n=2 thêm `[Năm chết]` / `[Năm chết 2]` (đúng tên cũ, không phải `[Năm chết 1]`).

Tài sản `k=1` không hậu tố; `k>=2` hậu tố `" " & k`:

```text
[Loại sổ], [Serial], [Số vào sổ], [Số thửa], [Số tờ], [Địa chỉ đất],
[Diện tích], [Hình thức sử dụng], [Loại đất], [Thời hạn 1],
[ONT], [CLN], [NTS], [LUC], [Nguồn gốc], [Ngày cấp sổ], [Cơ quan cấp sổ]
```

Hồ sơ: `[Niêm Yết]`, `[Số công chứng]`, `[Người ủy quyền]`, `[Người ủy quyền2]`.

Giữ placeholder mới:

```text
{{chu_dat.danh_sach}}
{{nguoi_da_chet.danh_sach}}
{{nguoi_nhan_dat.danh_sach}}
{{nhom_tu_choi.TC_DEFAULT.danh_sach}}
{{cay_nhanh.danh_sach}}
```

`Năm sinh` Word: dùng `YearDisplay` nếu người dùng nhập `yyyy`/`mm/yyyy` thì **xuất đúng text đã nhập**, không ép `01/01`. Cùng quy tắc SOT §7.

- [ ] **Step 1: `WordSlotPeople`**

```vb
Public Function WordSlotPeople() As Collection
    Dim result As Collection
    Dim owners As Collection
    Dim others As Collection
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim person As Object

    Set result = New Collection
    Set owners = New Collection
    Set others = New Collection
    Set lo = PeopleTable()

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If PersonHasData(rowIndex) Then
            If SafeBool(PersonCell(rowIndex, COL_OWNER).Value2) Then
                owners.Add rowIndex
            Else
                others.Add rowIndex
            End If
        End If
    Next rowIndex

    AppendSlotRows result, owners
    AppendSlotRows result, others
    Set WordSlotPeople = result
End Function
```

- [ ] **Step 2: `BuildPlaceholderMap` duyệt slot, tài sản, hồ sơ, nhóm SOT**

Giá trị trống → `""`, không `"0"`, không `00/01/1900`.

- [ ] **Step 3: `RunWordExport`**

```vb
Dim placeholders As Object
Dim key As Variant
Set placeholders = BuildPlaceholderMap()
For Each key In placeholders.Keys
    ReplaceEverywhere wordDoc, CStr(key), CStr(placeholders(key))
Next key
```

Sau thay: nếu còn `[` paired `]` **và** token nằm trong map nhưng chưa thay thì lỗi. Token `[` lẻ trong văn bản pháp lý (không phải placeholder đã biết) **không** được coi là lỗi. Cách kiểm:

```vb
For Each key In placeholders.Keys
    If DocumentContainsToken(wordDoc, CStr(key)) Then
        errorText = "Còn placeholder chưa thay: " & CStr(key)
        GoTo ExportFailed
    End If
Next key
```

Bỏ check thô `DocumentContainsToken(..., "{{")` một khi đã duyệt key. Vẫn fail nếu `{{` còn lại (placeholder mới quên map).

- [ ] **Step 4: Test**

Build TEST: xuất ra `output/word-test/`. Dùng OfficeCLI:

```powershell
officecli view output/word-test/Ho_so_thua_ke_*.docx text
```

Expected: có `Ông Nguyễn Văn A`, không còn `[Tên 1]`, `[Serial]`; năm chết người 1 là `1999` không phải `01/01/1999`.

Mở sẵn một file Word của user (nếu có). Macro không được `Quit` phiên đó — test bằng đếm `WinWord` trước/sau: số cửa sổ user không giảm. Agent chỉ `Quit` instance `Visible=False` vừa tạo.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modExportData.bas src/vba/modWordExport.bas
git commit -m "feat(excel): map legacy Word placeholders from HangTK export groups"
```

---

### Task 6: Làm lại đoạn thỏa thuận trong `1. PCDS .docx`

**Files:**
- Modify: `templates/word/1. PCDS .docx` — **chỉ sau khi backup copy** `templates/word/archive/1. PCDS.backup-v0.docx` (git add archive)
- Không đụng các file Word khác

Mẫu cũ giả định người 3 nhận đất, người 4–10 tặng cho. Checkbox `NhanDat` phá giả định đó.

- [ ] **Step 1: Copy backup**

```powershell
New-Item -ItemType Directory -Force -Path templates/word/archive | Out-Null
Copy-Item -LiteralPath "templates/word/1. PCDS .docx" -Destination "templates/word/archive/1. PCDS.backup-v0.docx"
```

- [ ] **Step 2: Đọc layout bằng OfficeCLI**

```powershell
officecli view "templates/word/1. PCDS .docx" text --max-lines 200
officecli view "templates/word/1. PCDS .docx" outline
```

Tìm đoạn thỏa thuận (thường quanh `[Tên 3]`, tặng cho). Thay **chỉ các câu gán vai trò nhận/từ chối** bằng:

```text
Người nhận quyền sử dụng đất: {{nguoi_nhan_dat.danh_sach}}.
Những người còn lại từ chối nhận di sản: {{nhom_tu_choi.TC_DEFAULT.danh_sach}}.
```

Giữ bảng/danh sách `[Tên n]` `[CCCD n]` … cho nhân thân. Giữ block tài sản `[Serial]` `[Số thửa]` …

Nếu placeholder bị tách run: `officecli set` replace cả cụm, hoặc gộp run bằng raw-set. Verify:

```powershell
officecli query "templates/word/1. PCDS .docx" "run" 
# không có run chỉ chứa "[Tên" mà thiếu " 1]"
```

Cách chắc: `officecli set file / --find "[Tên 1]" --replace "[Tên 1]"` không đổi chữ nhưng có thể không gộp run. Nếu `DocumentContainsToken` VBA không thấy, gộp XML.

- [ ] **Step 3: Xuất thử từ TEST workbook, trỏ `DuongDanMauWord` sang PCDS**

Trong build script, thêm switch hoặc param:

```powershell
param(..., [string]$WordTemplate = "")
if ($WordTemplate -eq "") { $WordTemplate = $testTemplate } # mặc định biên bản kiểm thử
```

Chạy một lần:

```powershell
powershell -NoProfile -File scripts/build-inheritance-workbook.ps1 -WordTemplate "D:\excelTK\templates\word\1. PCDS .docx"
```

Expected: DOCX mới trong `output/word-test/`; không còn `{{nguoi_nhan_dat.danh_sach}}`; có tên `Chị Nguyễn Thị C1` và `Chị Nguyễn Thị D` ở câu người nhận; `Anh Nguyễn Văn C2` ở câu từ chối; Word khác đang mở không bị đóng.

- [ ] **Step 4: Commit**

```bash
git add "templates/word/1. PCDS .docx" "templates/word/archive/1. PCDS.backup-v0.docx" scripts/build-inheritance-workbook.ps1
git commit -m "fix(word): PCDS nhận/từ chối theo NhanDat, giữ slot nhân thân"
```

---

### Task 7: Snapshot chốt v0.3.0

**Files:**
- Create: `templates/excel/chot-v0.3.0/` (xlsm sạch + TEST + qa xlsx + README)
- Modify: `docs/specs/inheritance-branch-architecture.md` — thêm mục 5.3 `tblTaiSan`, 5.4 `HoSo`, 13 bổ sung nhóm xuất `TaiSan`/`HoSo`, sức chứa mẫu

- [ ] **Step 1: Chạy full build không `-WordTemplate` rồi một lần với PCDS**

Cả hai phải xanh. Copy 4 file giống v0.2.1 sang `templates/excel/chot-v0.3.0/`. README ghi version `0.3.0`, schema vẫn `2.0.0` hoặc bump `2.1.0` nếu thêm cột người `LoaiCC` — **bump `PhienBanCauTruc` lên `2.1.0`** vì thêm cột.

- [ ] **Step 2: Cập nhật SOT các mục đã khóa ở đầu plan** (tài sản một dòng, sức chứa, placeholder cũ). Không đưa `tblDiSan` vào SOT.

- [ ] **Step 3: Commit**

```bash
git add templates/excel/chot-v0.3.0 docs/specs/inheritance-branch-architecture.md
git commit -m "feat(excel): chốt thừa kế MVP v0.3.0 có tài sản và PCDS"
```

Pha 1 xong. Có thể dừng để user dùng file thừa kế.

---

### Task 8: Spec workbook A/B

**Files:**
- Create: `docs/specs/ab-transfer-workbook.md`

Nội dung bắt buộc, không TBD:

```markdown
# Workbook hồ sơ hai bên A/B

Một file = một hồ sơ chuyển nhượng / phân chia hai bên.
Không có HangTK, ParentNguoiID, NhanDat, LaChuDat.

## Người — tblNguoiAB
NguoiID, STTNhap, Ben (A|B), STTBen, HoTen, NgaySinh, NgayChet,
SoGiayTo, NgayCap, DiaChi, LoaiCC, NoiCapCC, NhanDiaChi
+ bộ Goc/Tinh như thừa kế.

Ben: ô khóa, bấm một lần A↔B. Dòng mới mặc định A.
STTBen do VBA đánh lại mỗi bên từ 1.

## Tài sản — tblTaiSan (cùng cột v0.3)

## HoSo
NiemYet, SoCongChung, GiaChuyenNhuong, SoDienThoai, MauWordDangChon,
SucChuaNguoiBen (mỗi bên), SucChuaTaiSanMau

## XuatAn VeryHidden
tblXuatBenA, tblXuatBenB, section TaiSan, HoSo.
STTBen liên tục. Placeholder [Tên n] của ĐK18: n là STTBen trong đúng bên mà mẫu đòi hỏi — ghi rõ sau khi Task 1 phân loại: mẫu `2. DK18, thuế  1.docx` map bên nào.

## Validation chặn
- Thiếu người bên A hoặc bên B
- Ben không thuộc A/B
- Vượt sức chứa mẫu
- Ngày sai như SOT thừa kế

## An toàn Word
CreateObject, Range.Find, file mới, không Quit phiên user.
```

Sau Task 1, điền chính xác map placeholder ĐK18 vào spec này (số người mỗi bên, hậu tố tài sản).

- [ ] **Step 1: Viết spec từ inventory Task 1**
- [ ] **Step 2: Commit**

```bash
git add docs/specs/ab-transfer-workbook.md
git commit -m "docs(specs): define A/B transfer workbook"
```

---

### Task 9: Script + VBA workbook A/B

**Files:**
- Create: `scripts/build-ab-workbook.ps1` (clone cấu trúc build thừa kế, **không** import `modHangTKNhanh`)
- Create: `src/vba-ab/modCommon.bas`, `modNguoi.bas`, `modBen.bas`, `modTaiSan.bas`, `modNgayThang.bas`, `modGiayTo.bas`, `modValidation.bas`, `modExportData.bas`, `modWordExport.bas`, `NhapLieu.cls`, `ThisWorkbook.cls`

**Interfaces:**
- `Public Sub HandleBenClick(ByVal targetCell As Range)` — nếu `A` thì `B`, nếu `B` thì `A`, `RebuildSttBen`
- `Public Sub RebuildSttBen()` — mỗi bên đếm 1..n theo thứ tự dòng
- Cột `Ben` format text, màu: A `#D5E4FF`, B `#FFE1A0` (conditional formatting)

- [ ] **Step 1: `modBen.bas`**

```vb
Public Sub HandleBenClick(ByVal targetCell As Range)
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim current As String
    Set lo = PeopleTable()
    If Intersect(targetCell, lo.ListColumns("Ben").DataBodyRange) Is Nothing Then Exit Sub
    rowIndex = targetCell.Row - lo.DataBodyRange.Row + 1
    If Not PersonHasData(rowIndex) Then
        MsgBox UnicodeText("..."), vbInformation ' Hãy nhập họ tên trước khi chọn bên
        GoTo MoveSel
    End If
    current = UCase$(Trim$(CStr(PersonCell(rowIndex, "Ben").Value2)))
    If current = "B" Then
        PersonCell(rowIndex, "Ben").Value2 = "A"
    Else
        PersonCell(rowIndex, "Ben").Value2 = "B"
    End If
    RebuildSttBen
MoveSel:
    PersonCell(rowIndex, "HoTen").Select
End Sub
```

`NhapLieu.cls` `SelectionChange`: nếu ô `Ben` thì `HandleBenClick`.

- [ ] **Step 2: Validation**

```vb
Dim countA As Long, countB As Long
' ...
If countA = 0 Or countB = 0 Then
    AddIssue "Lỗi", "BOTH_SIDES_REQUIRED", ...
End If
```

- [ ] **Step 3: Fake data TEST**

3 người A, 2 người B, 1 tài sản. Export `2. DK18, thuế  1.docx` vào `output/word-test-ab/`.

Assert: `STTBen` bên A = 1,2,3; bên B = 1,2. Placeholder bên A không lấy nhầm người B.

- [ ] **Step 4: Snapshot `templates/excel/chot-ab-v0.1.0/` + README**

- [ ] **Step 5: Commit**

```bash
git add scripts/build-ab-workbook.ps1 src/vba-ab templates/excel/chot-ab-v0.1.0
git commit -m "feat(excel): add A/B transfer workbook v0.1.0"
```

---

### Task 10: Nối thêm mẫu thừa kế còn lại (sau PCDS)

Chỉ làm khi v0.3 PCDS đã chốt.

Với mỗi file: `1. Khai nhận di sản Vũ Dương.docx`, `1.1. Từ chối di sản_ Ý Yên.docx`, `VP_Phân chia di sản.docx`, `VP_Từ chối di sản.docx`:

1. Đọc bảng placeholder Task 1.
2. Nếu mẫu dùng cùng `[Tên n]` + tài sản 1 → trỏ `MauWordDangChon`, set `SucChua*`, chạy `RunWordExport` silent trên TEST.
3. Nếu mẫu có placeholder lạ: thêm vào `BuildPlaceholderMap` hoặc ghi `docs/specs/noi-dung-ngoai-sot.md` mục “placeholder chưa map”, **không đoán**.
4. Backup trước khi sửa Word.

Commit từng mẫu một:

```bash
git commit -m "feat(word): connect <tên mẫu> to inheritance export map"
```

---

## Thứ tự cấm đảo

1. Không sửa PCDS trước khi `BuildPlaceholderMap` chạy trên template kiểm thử (biên bản `{{...}}` v0.2.1).
2. Không viết workbook A/B vào `src/vba/` thừa kế.
3. Không commit `.xlsb` nguồn.
4. Không thêm `tblPhanChiaDiSan` / tỷ lệ % / 3+ chủ đất.

## Self-review

- SOT §5–14: người/nhánh/ngày/xuất nhóm — đã có v0.2.1; plan chỉ thêm tài sản, hồ sơ, sức chứa, placeholder cũ.
- `noi-dung-ngoai-sot.md` §1 di sản/tỷ lệ: cố ý bỏ.
- §2.2 phiếu dọc: cố ý đổi thành một dòng/tài sản (khóa ở đầu plan).
- §3 hồ sơ hai bên: Task 8–9, file riêng.
- §4.2 Word an toàn: Task 5, giữ `CreateObject`.
- §5 loại giấy tờ: Task 3.
- §6 sức chứa: Task 4.
- Placeholder scan plan: không còn TBD; map ĐK18 chi tiết chờ output Task 1 rồi ghi vào spec Task 8 — đó là bước tuần tự, không phải chỗ trống trong Task 1–7.
)
