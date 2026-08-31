param(
    [string]$WorkspaceRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$Version = "0.2.1"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$officeCli = Join-Path $env:LOCALAPPDATA "OfficeCLI\officecli.exe"
if (-not (Test-Path -LiteralPath $officeCli)) {
    throw "Không tìm thấy OfficeCLI tại $officeCli"
}

$workspace = (Resolve-Path -LiteralPath $WorkspaceRoot).Path
$sourceWorkbook = Join-Path $workspace "templates\excel\Dữ liệu thừa kế (2).xlsb"
$placeholderReport = Join-Path $workspace "docs\reference\word-placeholders-2026-08-31.md"
$vbaRoot = Join-Path $workspace "src\vba"
$outputDir = Join-Path $workspace "output\workbook-dev"
$wordOutputDir = Join-Path $workspace "output\word-test"
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$baseWorkbook = Join-Path $outputDir "base-inheritance-$timestamp.xlsx"
$targetWorkbook = Join-Path $outputDir "Ho_so_thua_ke_MVP_v$Version.xlsm"
$testWorkbook = Join-Path $outputDir "Ho_so_thua_ke_MVP_v$Version-TEST.xlsm"
$qaCleanWorkbook = Join-Path $outputDir "qa-clean-v$Version.xlsx"
$qaTestWorkbook = Join-Path $outputDir "qa-test-v$Version.xlsx"
$testTemplate = Join-Path $outputDir "template-test-sot-$timestamp.docx"

New-Item -ItemType Directory -Force -Path $outputDir, $wordOutputDir | Out-Null

function Move-ToBackup {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        $directory = Split-Path -Parent $Path
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $extension = [System.IO.Path]::GetExtension($Path)
        $backup = Join-Path $directory "$baseName.backup-$timestamp$extension"
        Move-Item -LiteralPath $Path -Destination $backup
        Write-Output "BACKUP=$backup"
    }
}

function Invoke-OfficeCli {
    param([string[]]$Arguments)
    & $officeCli @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "OfficeCLI thất bại: $($Arguments -join ' ')"
    }
}

function Invoke-OfficeCliBatch {
    param([string]$FilePath, [array]$Operations)
    $json = $Operations | ConvertTo-Json -Depth 8 -Compress
    $json | & $officeCli batch $FilePath --json
    if ($LASTEXITCODE -ne 0) {
        throw "OfficeCLI batch thất bại cho $FilePath"
    }
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Color-Ref {
    param([string]$Hex)
    $clean = $Hex.TrimStart('#')
    $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Convert-VbaUnicodeLiterals {
    param([string]$Code)

    # VBE stores source code using the machine's legacy code page. This machine
    # uses Windows-1252, which cannot store Vietnamese characters. Convert only
    # affected string literals to calls that rebuild the exact Unicode text at
    # runtime, while keeping the imported VBA source itself ASCII-only.
    return [regex]::Replace($Code, '"(?:[^\"]|"\")*"', {
        param($match)

        $inner = $match.Value.Substring(1, $match.Value.Length - 2).Replace('""', '"')
        if (-not [regex]::IsMatch($inner, '[^\x00-\x7F]')) {
            return $match.Value
        }

        $hexCodes = ($inner.ToCharArray() | ForEach-Object {
            '{0:X4}' -f [int][char]$_
        }) -join ' '

        return 'UnicodeText("' + $hexCodes + '")'
    })
}

function Read-VbaCode {
    param([string]$Path)
    $code = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $code = [regex]::Replace($code, '(?m)^Attribute VB_Name = ".*"\r?\n', '')
    return Convert-VbaUnicodeLiterals $code
}

function Add-StandardModule {
    param([object]$Project, [string]$Name, [string]$SourcePath)

    $existing = $null
    try { $existing = $Project.VBComponents.Item($Name) } catch { $existing = $null }
    if ($null -ne $existing) {
        $Project.VBComponents.Remove($existing)
        Release-ComObject $existing
    }

    $component = $Project.VBComponents.Add(1)
    $component.Name = $Name
    $component.CodeModule.AddFromString((Read-VbaCode $SourcePath))
    Release-ComObject $component
}

function Set-DocumentModuleCode {
    param([object]$Component, [string]$SourcePath)
    $module = $Component.CodeModule
    if ($module.CountOfLines -gt 0) {
        $module.DeleteLines(1, $module.CountOfLines)
    }
    $module.AddFromString((Read-VbaCode $SourcePath))
    Release-ComObject $module
}

function Add-FormulaFormat {
    param(
        [object]$Range,
        [string]$Formula,
        [string]$Fill,
        [string]$FontColor = "000000"
    )
    $condition = $Range.FormatConditions.Add(2, $null, $Formula)
    $condition.Interior.Color = Color-Ref $Fill
    $condition.Font.Color = Color-Ref $FontColor
    Release-ComObject $condition
}

function Create-QaCopy {
    param([object]$Excel, [string]$SourcePath, [string]$DestinationPath)

    Move-ToBackup $DestinationPath
    $Excel.AutomationSecurity = 3
    $book = $Excel.Workbooks.Open($SourcePath, 0, $true)
    try {
        foreach ($sheet in $book.Worksheets) {
            $sheet.Visible = -1
            Release-ComObject $sheet
        }
        $book.SaveAs($DestinationPath, 51)
    }
    finally {
        $book.Close($false)
        Release-ComObject $book
    }
}

function Get-WordCapacityRows {
    param([string]$ReportPath)

    $lines = Get-Content -LiteralPath $ReportPath -Encoding UTF8
    $rows = @()
    for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
        if ($lines[$lineIndex] -notmatch '^\| (.+?) \| [^|]+ \| [^|]+ \| (\d+) \| (\d+) \| (.+?) \|$') {
            continue
        }

        $templateName = $matches[1]
        $assetText = $matches[4]
        $sectionStart = -1
        for ($candidate = 0; $candidate -lt $lines.Count; $candidate++) {
            if ($lines[$candidate] -eq "## $templateName") {
                $sectionStart = $candidate
                break
            }
        }

        $peopleCapacity = 0
        if ($sectionStart -ge 0) {
            for ($detailIndex = $sectionStart + 1; $detailIndex -lt $lines.Count; $detailIndex++) {
                if ($lines[$detailIndex] -match '^## ') { break }
                $nameMatches = [regex]::Matches($lines[$detailIndex], '\[Tên (\d+)\]')
                foreach ($nameMatch in $nameMatches) {
                    $peopleCapacity = [Math]::Max($peopleCapacity, [int]$nameMatch.Groups[1].Value)
                }
            }
        }

        $assetCapacity = 0
        foreach ($assetMatch in [regex]::Matches($assetText, '(\d+)')) {
            $assetCapacity = [Math]::Max($assetCapacity, [int]$assetMatch.Groups[1].Value)
        }
        $rows += [pscustomobject]@{
            TemplateName = $templateName
            SucChuaNguoi = $peopleCapacity
            SucChuaTaiSan = $assetCapacity
        }
    }
    return $rows
}

if (-not (Test-Path -LiteralPath $sourceWorkbook)) {
    throw "Không tìm thấy workbook nguồn: $sourceWorkbook"
}
if (-not (Test-Path -LiteralPath $placeholderReport)) {
    throw "Không tìm thấy báo cáo placeholder để dựng sức chứa mẫu: $placeholderReport"
}

$sourceHashBefore = (Get-FileHash -Algorithm SHA256 -LiteralPath $sourceWorkbook).Hash

Move-ToBackup $targetWorkbook
Move-ToBackup $testWorkbook

Invoke-OfficeCli @("create", $baseWorkbook)
Invoke-OfficeCli @("set", $baseWorkbook, "/Sheet1", "--prop", "name=NhapLieu")
Invoke-OfficeCli @("add", $baseWorkbook, "/", "--type", "sheet", "--prop", "name=CauHinh")
Invoke-OfficeCli @("add", $baseWorkbook, "/", "--type", "sheet", "--prop", "name=KiemTra")
Invoke-OfficeCli @("add", $baseWorkbook, "/", "--type", "sheet", "--prop", "name=XuatAn")

$baseOperations = @(
    @{ command = "set"; path = "/NhapLieu/A1"; props = @{ value = "HỒ SƠ THỪA KẾ — NHẬP THEO HÀNG THỪA KẾ" } },
    @{ command = "set"; path = "/NhapLieu/A2"; props = @{ value = "Ngày hiển thị và xuất đúng nội dung đã nhập: dd/mm/yyyy, mm/yyyy hoặc yyyy; ngày bổ sung để tính được giữ ẩn." } },
    @{ command = "set"; path = "/NhapLieu/A4"; props = @{ value = "Hàng TK tối đa" } },
    @{ command = "set"; path = "/NhapLieu/B4"; props = @{ value = 2 } },
    @{ command = "set"; path = "/NhapLieu/D4"; props = @{ value = "Hai dòng đầu là chủ đất cố định ở Hàng TK 0" } },
    @{ command = "set"; path = "/NhapLieu/A8"; props = @{ value = "STTNhap" } },
    @{ command = "set"; path = "/NhapLieu/B8"; props = @{ value = "HoTen" } },
    @{ command = "set"; path = "/NhapLieu/C8"; props = @{ value = "NgaySinh" } },
    @{ command = "set"; path = "/NhapLieu/D8"; props = @{ value = "NgayChet" } },
    @{ command = "set"; path = "/NhapLieu/E8"; props = @{ value = "SoGiayTo" } },
    @{ command = "set"; path = "/NhapLieu/F8"; props = @{ value = "NgayCap" } },
    @{ command = "set"; path = "/NhapLieu/G8"; props = @{ value = "DiaChi" } },
    @{ command = "set"; path = "/NhapLieu/H8"; props = @{ value = "HangTK" } },
    @{ command = "set"; path = "/NhapLieu/I8"; props = @{ value = "NhanDat" } },
    @{ command = "set"; path = "/NhapLieu/J8"; props = @{ value = "TrangThai" } },
    @{ command = "set"; path = "/NhapLieu/K8"; props = @{ value = "NguoiID" } },
    @{ command = "set"; path = "/NhapLieu/L8"; props = @{ value = "ParentNguoiID" } },
    @{ command = "set"; path = "/NhapLieu/M8"; props = @{ value = "LaChuDat" } },
    @{ command = "set"; path = "/NhapLieu/N8"; props = @{ value = "NhomTuChoiID" } },
    @{ command = "set"; path = "/NhapLieu/O8"; props = @{ value = "NgaySinhGoc" } },
    @{ command = "set"; path = "/NhapLieu/P8"; props = @{ value = "NgaySinhTinh" } },
    @{ command = "set"; path = "/NhapLieu/Q8"; props = @{ value = "NgayChetGoc" } },
    @{ command = "set"; path = "/NhapLieu/R8"; props = @{ value = "NgayChetTinh" } },
    @{ command = "set"; path = "/NhapLieu/S8"; props = @{ value = "NgayCapGoc" } },
    @{ command = "set"; path = "/NhapLieu/T8"; props = @{ value = "NgayCapTinh" } },
    @{ command = "set"; path = "/CauHinh/A1"; props = @{ value = "CẤU HÌNH HỒ SƠ" } },
    @{ command = "set"; path = "/CauHinh/A3"; props = @{ value = "HangTKToiDa" } },
    @{ command = "set"; path = "/CauHinh/A4"; props = @{ value = "BangMauHangTK" } },
    @{ command = "set"; path = "/CauHinh/A5"; props = @{ value = "PhienBanCauTruc" } },
    @{ command = "set"; path = "/CauHinh/A6"; props = @{ value = "NguoiIDTiepTheo" } },
    @{ command = "set"; path = "/CauHinh/A7"; props = @{ value = "SucChuaNguoi" } },
    @{ command = "set"; path = "/CauHinh/A8"; props = @{ value = "DuongDanMauWord" } },
    @{ command = "set"; path = "/CauHinh/A9"; props = @{ value = "ThuMucXuat" } },
    @{ command = "set"; path = "/CauHinh/A10"; props = @{ value = "TaiSanIDTiepTheo" } },
    @{ command = "set"; path = "/KiemTra/A1"; props = @{ value = "KẾT QUẢ KIỂM TRA DỮ LIỆU" } },
    @{ command = "set"; path = "/KiemTra/A2"; props = @{ value = "Số lỗi chặn" } },
    @{ command = "set"; path = "/KiemTra/A3"; props = @{ value = "Số cảnh báo" } },
    @{ command = "set"; path = "/XuatAn/A1"; props = @{ value = "DỮ LIỆU XUẤT — KHÔNG NHẬP TAY" } }
)
Invoke-OfficeCliBatch -FilePath $baseWorkbook -Operations $baseOperations
Invoke-OfficeCli @("add", $baseWorkbook, "/NhapLieu", "--type", "table", "--prop", "ref=A8:T38", "--prop", "name=tblNguoi", "--prop", "displayName=tblNguoi", "--prop", "style=light1", "--prop", "showRowStripes=false")
Invoke-OfficeCli @("add", $baseWorkbook, "/", "--type", "sheet", "--prop", "name=DanhMuc")
Invoke-OfficeCli @("close", $baseWorkbook)

Invoke-OfficeCli @("create", $testTemplate)
Invoke-OfficeCli @("set", $testTemplate, "/", "--prop", "docDefaults.font=Arial", "--prop", "docDefaults.fontSize=11pt")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=BIÊN BẢN KIỂM THỬ DỮ LIỆU THỪA KẾ", "--prop", "style=Heading1")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Chủ đất: {{chu_dat.danh_sach}}")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Người đã chết: {{nguoi_da_chet.danh_sach}}")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Người nhận đất: {{nguoi_nhan_dat.danh_sach}}")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Nhóm từ chối mặc định: {{nhom_tu_choi.TC_DEFAULT.danh_sach}}")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Cây nhánh: {{cay_nhanh.danh_sach}}")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Ngày sinh người đầu tiên: {{kiem_thu.ngay_sinh_nguoi_1}}")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Ngày chết người đầu tiên: {{kiem_thu.ngay_chet_nguoi_1}}")
Invoke-OfficeCli @("add", $testTemplate, "/body", "--type", "paragraph", "--prop", "text=Tuổi lúc chết: {{kiem_thu.tuoi_luc_chet_nguoi_1}}")
Invoke-OfficeCli @("validate", $testTemplate)

$excel = $null
$workbook = $null
$testBook = $null
$sourceBook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Open($baseWorkbook)
    $inputSheet = $workbook.Worksheets.Item("NhapLieu")
    $configSheet = $workbook.Worksheets.Item("CauHinh")
    $checkSheet = $workbook.Worksheets.Item("KiemTra")
    $exportSheet = $workbook.Worksheets.Item("XuatAn")
    $catalogSheet = $workbook.Worksheets.Item("DanhMuc")
    $peopleTable = $inputSheet.ListObjects.Item("tblNguoi")

    $sourceBook = $excel.Workbooks.Open($sourceWorkbook, 0, $true)
    $sourceTables = $sourceBook.Worksheets.Item("Tables")
    $catalogSheet.Range("A4:A6").Value2 = $sourceTables.Range("D6:D8").Value2
    $catalogSheet.Range("B4:B8").Value2 = $sourceTables.Range("D11:D15").Value2
    $catalogSheet.Range("C4:C5").Value2 = $sourceTables.Range("D20:D21").Value2
    $catalogSheet.Range("D4:D5").Value2 = $sourceTables.Range("E3:E4").Value2
    $catalogSheet.Range("E4:E5").Value2 = $sourceTables.Range("D3:D4").Value2
    $catalogSheet.Range("F4:F5").Value2 = $sourceTables.Range("E3:E4").Value2
    $catalogSheet.Range("G4:G5").Value2 = $sourceTables.Range("F3:F4").Value2
    $catalogSheet.Range("H4:H11").Value2 = $sourceTables.Range("D23:D30").Value2
    $sourceBook.Close($false)
    Release-ComObject $sourceTables
    Release-ComObject $sourceBook
    $sourceBook = $null

    $catalogSheet.Range("A1:L1").Merge()
    $catalogSheet.Range("A1").Value2 = "DANH MỤC TRA CỨU"
    $catalogSheet.Range("A1:L1").Font.Name = "Arial"
    $catalogSheet.Range("A1:L1").Font.Bold = $true
    $catalogSheet.Range("A1:L1").Font.Color = Color-Ref "FFFFFF"
    $catalogSheet.Range("A1:L1").Interior.Color = Color-Ref "304F78"
    $catalogHeaders = @("Loại sổ", "Loại đất", "Hình thức sử dụng", "Cơ quan cấp sổ", "Loại giấy tờ", "Nơi cấp CC", "Nhãn địa chỉ", "Người ủy quyền")
    for ($headerIndex = 0; $headerIndex -lt $catalogHeaders.Count; $headerIndex++) {
        $catalogSheet.Cells.Item(3, $headerIndex + 1).Value2 = [string]$catalogHeaders[$headerIndex]
    }
    $catalogSheet.Range("A3:H3").Font.Bold = $true
    $catalogSheet.Range("A3:H3").Interior.Color = Color-Ref "D8EBFF"

    $capacityRows = Get-WordCapacityRows -ReportPath $placeholderReport
    $catalogSheet.Range("J3").Value2 = "TenMau"
    $catalogSheet.Range("K3").Value2 = "SucChuaNguoi"
    $catalogSheet.Range("L3").Value2 = "SucChuaTaiSan"
    for ($capacityIndex = 0; $capacityIndex -lt $capacityRows.Count; $capacityIndex++) {
        $capacityRow = $capacityIndex + 4
        $catalogSheet.Cells.Item($capacityRow, 10).Value2 = [string]$capacityRows[$capacityIndex].TemplateName
        $catalogSheet.Cells.Item($capacityRow, 11).Value2 = [int]$capacityRows[$capacityIndex].SucChuaNguoi
        $catalogSheet.Cells.Item($capacityRow, 12).Value2 = [int]$capacityRows[$capacityIndex].SucChuaTaiSan
    }
    $catalogSheet.Range("N3").Value2 = "Tờ bản đồ 2025"
    $catalogSheet.Range("O3").Value2 = "Sáp nhập thôn"
    $catalogSheet.Range("N3:O3").Font.Bold = $true
    $catalogSheet.Range("N3:O3").Interior.Color = Color-Ref "FFF1B8"

    $catalogNamedRanges = @(
        @{ Name = "DanhMuc_LoaiSo"; Formula = "=DanhMuc!`$A`$4:`$A`$6" },
        @{ Name = "DanhMuc_LoaiDat"; Formula = "=DanhMuc!`$B`$4:`$B`$8" },
        @{ Name = "DanhMuc_HinhThucSuDung"; Formula = "=DanhMuc!`$C`$4:`$C`$5" },
        @{ Name = "DanhMuc_CoQuanCapSo"; Formula = "=DanhMuc!`$D`$4:`$D`$5" },
        @{ Name = "DanhMuc_LoaiGiayTo"; Formula = "=DanhMuc!`$E`$4:`$E`$5" },
        @{ Name = "DanhMuc_NoiCapCC"; Formula = "=DanhMuc!`$F`$4:`$F`$5" },
        @{ Name = "DanhMuc_NhanDiaChi"; Formula = "=DanhMuc!`$G`$4:`$G`$5" },
        @{ Name = "DanhMuc_NguoiUyQuyen"; Formula = "=DanhMuc!`$H`$4:`$H`$11" },
        @{ Name = "DanhMuc_SucChuaMau"; Formula = ("=DanhMuc!`$J`$4:`$L`${0}" -f ($capacityRows.Count + 3)) }
    )
    foreach ($namedRange in $catalogNamedRanges) {
        $workbook.Names.Add($namedRange.Name, $namedRange.Formula) | Out-Null
    }

    foreach ($cardStartRow in @(41, 61, 81)) {
        $inputSheet.Cells.Item($cardStartRow + 1, 3).Validation.Delete()
        $inputSheet.Cells.Item($cardStartRow + 1, 3).Validation.Add(3, 1, 1, "=DanhMuc_LoaiSo")
        $inputSheet.Cells.Item($cardStartRow + 8, 3).Validation.Delete()
        $inputSheet.Cells.Item($cardStartRow + 8, 3).Validation.Add(3, 1, 1, "=DanhMuc_HinhThucSuDung")
        $inputSheet.Cells.Item($cardStartRow + 9, 3).Validation.Delete()
        $inputSheet.Cells.Item($cardStartRow + 9, 3).Validation.Add(3, 1, 1, "=DanhMuc_LoaiDat")
        $inputSheet.Cells.Item($cardStartRow + 17, 3).Validation.Delete()
        $inputSheet.Cells.Item($cardStartRow + 17, 3).Validation.Add(3, 1, 1, "=DanhMuc_CoQuanCapSo")
    }

    try {
        $workbook.BuiltinDocumentProperties.Item("Title").Value = "Hồ sơ thừa kế — MVP theo Hàng TK"
        $workbook.BuiltinDocumentProperties.Item("Comments").Value = "SOT: inheritance-branch-architecture.md; schema 2.0.0"
    }
    catch {
        # Một số bản Office không cho ghi BuiltinDocumentProperties qua COM.
        # Đây chỉ là mô tả phụ, không ảnh hưởng workbook.
    }

    $inputSheet.Range("A1:T1").Merge()
    $inputSheet.Range("A2:T2").Merge()
    $inputSheet.Range("A1").Value2 = "HỒ SƠ THỪA KẾ — NHẬP THEO HÀNG THỪA KẾ"
    $inputSheet.Range("A2").Value2 = "Ngày hiển thị và xuất đúng nội dung đã nhập: dd/mm/yyyy, mm/yyyy hoặc yyyy; ngày bổ sung để tính được giữ ẩn."
    $inputSheet.Range("A1:T1").Font.Name = "Arial"
    $inputSheet.Range("A1:T1").Font.Size = 18
    $inputSheet.Range("A1:T1").Font.Bold = $true
    $inputSheet.Range("A1:T1").Font.Color = Color-Ref "FFFFFF"
    $inputSheet.Range("A1:T1").Interior.Color = Color-Ref "1B5E3B"
    $inputSheet.Range("A1:T1").HorizontalAlignment = -4108
    $inputSheet.Range("A1:T1").VerticalAlignment = -4108
    $inputSheet.Range("A2:T2").Font.Name = "Arial"
    $inputSheet.Range("A2:T2").Font.Size = 10
    $inputSheet.Range("A2:T2").Font.Color = Color-Ref "334E68"
    $inputSheet.Range("A2:T2").Interior.Color = Color-Ref "EAF4EE"
    $inputSheet.Range("A2:T2").WrapText = $true
    $inputSheet.Rows.Item(1).RowHeight = 30
    $inputSheet.Rows.Item(2).RowHeight = 34

    $inputSheet.Range("A4").Value2 = "Hàng TK tối đa"
    $inputSheet.Range("B4").Value2 = 2
    $inputSheet.Range("D4:G4").Merge()
    $inputSheet.Range("D4").Value2 = "Hai dòng đầu là chủ đất cố định ở Hàng TK 0"
    $inputSheet.Range("A4:I4").Font.Name = "Arial"
    $inputSheet.Range("A4:I4").VerticalAlignment = -4108
    $inputSheet.Range("A4").Font.Bold = $true
    $inputSheet.Range("B4").Font.Bold = $true
    $inputSheet.Range("B4").HorizontalAlignment = -4108
    $inputSheet.Range("B4").Interior.Color = Color-Ref "FFF1B8"
    $inputSheet.Range("B4").Borders.Weight = 2
    $inputSheet.Range("B4").Validation.Delete()
    $inputSheet.Range("B4").Validation.Add(1, 1, 1, "0", "4")
    $inputSheet.Range("B4").Validation.ErrorTitle = "Hàng TK không hợp lệ"
    $inputSheet.Range("B4").Validation.ErrorMessage = "Chỉ nhập số nguyên từ 0 đến 4."
    $inputSheet.Range("B4").Validation.ShowError = $true

    $inputSheet.Range("B6:G6").Merge()
    $inputSheet.Range("H6:I6").Merge()
    $inputSheet.Range("B6").Value2 = "VÙNG NHẬP DỮ LIỆU"
    $inputSheet.Range("H6").Value2 = "BẤM ĐỂ CHỌN"
    $inputSheet.Range("B6:G6").Interior.Color = Color-Ref "D8EBFF"
    $inputSheet.Range("H6:I6").Interior.Color = Color-Ref "FFE3A3"
    $inputSheet.Range("B6:I6").Font.Name = "Arial"
    $inputSheet.Range("B6:I6").Font.Bold = $true
    $inputSheet.Range("B6:I6").HorizontalAlignment = -4108

    $dataRange = $peopleTable.DataBodyRange
    for ($row = 1; $row -le $dataRange.Rows.Count; $row++) {
        $dataRange.Cells.Item($row, 8).Value2 = if ($row -le 2) { 0 } else { 1 }
        $dataRange.Cells.Item($row, 9).Value2 = 0
        $dataRange.Cells.Item($row, 10).ClearContents()
        $dataRange.Cells.Item($row, 13).Value2 = $false
    }

    $peopleTable.TableStyle = "TableStyleLight9"
    $peopleTable.ShowTableStyleRowStripes = $false
    $peopleTable.HeaderRowRange.Font.Name = "Arial"
    $peopleTable.HeaderRowRange.Font.Bold = $true
    $peopleTable.HeaderRowRange.Font.Color = Color-Ref "FFFFFF"
    $peopleTable.HeaderRowRange.Interior.Color = Color-Ref "17653D"
    $peopleTable.HeaderRowRange.HorizontalAlignment = -4108
    $peopleTable.HeaderRowRange.VerticalAlignment = -4108
    $peopleTable.HeaderRowRange.WrapText = $true
    $inputSheet.Rows.Item(8).RowHeight = 30
    $inputSheet.Range("A8").Interior.Color = Color-Ref "5B677A"
    $inputSheet.Range("B8:G8").Interior.Color = Color-Ref "2F75B5"
    $inputSheet.Range("H8:I8").Interior.Color = Color-Ref "B26A00"
    $dataRange.Font.Name = "Arial"
    $dataRange.Font.Size = 10
    $dataRange.VerticalAlignment = -4108
    $dataRange.Rows.RowHeight = 26

    $inputSheet.Columns.Item("A").ColumnWidth = 9
    $inputSheet.Columns.Item("B").ColumnWidth = 24
    $inputSheet.Columns.Item("C").ColumnWidth = 13
    $inputSheet.Columns.Item("D").ColumnWidth = 13
    $inputSheet.Columns.Item("E").ColumnWidth = 18
    $inputSheet.Columns.Item("F").ColumnWidth = 13
    $inputSheet.Columns.Item("G").ColumnWidth = 34
    $inputSheet.Columns.Item("H").ColumnWidth = 10
    $inputSheet.Columns.Item("I").ColumnWidth = 12
    $inputSheet.Columns.Item("J").ColumnWidth = 15
    $inputSheet.Columns.Item("K").ColumnWidth = 14
    $inputSheet.Columns.Item("L").ColumnWidth = 18
    $inputSheet.Columns.Item("M").ColumnWidth = 12
    $inputSheet.Columns.Item("N").ColumnWidth = 18
    $inputSheet.Columns.Item("O").ColumnWidth = 16
    $inputSheet.Columns.Item("P").ColumnWidth = 16
    $inputSheet.Columns.Item("Q").ColumnWidth = 16
    $inputSheet.Columns.Item("R").ColumnWidth = 16
    $inputSheet.Columns.Item("S").ColumnWidth = 16
    $inputSheet.Columns.Item("T").ColumnWidth = 16
    $inputSheet.Range("J:W").EntireColumn.Hidden = $true

    $inputSheet.Range("C9:D38").NumberFormat = "@"
    $inputSheet.Range("F9:F38").NumberFormat = "@"
    $inputSheet.Range("E9:E38").NumberFormat = "@"
    $inputSheet.Range("O9:O38").NumberFormat = "@"
    $inputSheet.Range("Q9:Q38").NumberFormat = "@"
    $inputSheet.Range("S9:S38").NumberFormat = "@"
    $inputSheet.Range("P9:P38").NumberFormat = "dd/mm/yyyy"
    $inputSheet.Range("R9:R38").NumberFormat = "dd/mm/yyyy"
    $inputSheet.Range("T9:T38").NumberFormat = "dd/mm/yyyy"
    $inputSheet.Range("H9:H38").NumberFormat = '"H"0" ›"'
    $inputSheet.Range("I9:I38").NumberFormat = '[=1]"☑";[=0]"☐";"☐"'
    $inputSheet.Range("I9:I38").Font.Name = "Segoe UI Symbol"
    $inputSheet.Range("I9:I38").Font.Size = 14
    $inputSheet.Range("H9:I38").HorizontalAlignment = -4108
    $inputSheet.Range("A9:A38").HorizontalAlignment = -4108
    $inputSheet.Range("C9:F38").HorizontalAlignment = -4108
    $inputSheet.Range("G9:G38").WrapText = $true

    $inputSheet.Range("H9:H38").Validation.Delete()
    $inputSheet.Range("H9:H38").Validation.Add(1, 1, 1, "0", "4")
    $inputSheet.Range("H9:H38").Validation.ErrorTitle = "Hàng TK không hợp lệ"
    $inputSheet.Range("H9:H38").Validation.ErrorMessage = "Chỉ dùng Hàng TK 0 đến 4."
    $inputSheet.Range("H9:H38").Validation.ShowError = $true
    $inputSheet.Range("I9:I38").Validation.Delete()
    $inputSheet.Range("I9:I38").Validation.Add(1, 1, 1, "0", "1")
    $inputSheet.Range("I9:I38").Validation.ErrorTitle = "Giá trị không hợp lệ"
    $inputSheet.Range("I9:I38").Validation.ErrorMessage = "Nhận đất chỉ nhận 0 hoặc 1."
    $inputSheet.Range("I9:I38").Validation.ShowError = $true

    $rowColors = @("D5E4FF", "CDEFD8", "FFE1A0", "E2D0F6", "F6C8D2")
    $darkColors = @("2F5DA8", "237A50", "B26A00", "6747A8", "A33D54")
    $displayRows = $inputSheet.Range("A9:I38")
    $displayRows.FormatConditions.Delete()
    for ($level = 0; $level -le 4; $level++) {
        Add-FormulaFormat -Range $displayRows -Formula "=`$H9=$level" -Fill $rowColors[$level]
    }
    $stripeRange = $inputSheet.Range("A9:A38")
    $stripeRange.FormatConditions.Delete()
    for ($level = 0; $level -le 4; $level++) {
        Add-FormulaFormat -Range $stripeRange -Formula "=`$H9=$level" -Fill $darkColors[$level] -FontColor "FFFFFF"
    }

    $receiveRange = $inputSheet.Range("I9:I38")
    Add-FormulaFormat -Range $receiveRange -Formula '=$I9=1' -Fill "217346" -FontColor "FFFFFF"

    $assetFields = @(
        "Loại sổ", "Số phát hành/Serial", "Số vào sổ", "Số thửa", "Số tờ bản đồ",
        "Địa chỉ đất", "Diện tích (m²)", "Hình thức sử dụng", "Loại đất",
        "Thời hạn sử dụng", "ONT (m²)", "CLN (m²)", "NTS (m²)", "LUC (m²)",
        "Nguồn gốc", "Ngày cấp sổ", "Cơ quan cấp sổ", "Ghi chú"
    )
    $assetCardStartRows = @(41, 61, 81)
    $assetLongFieldOffsets = @(6, 15, 18)

    $inputSheet.Range("B40").Value2 = "THÔNG TIN TÀI SẢN"
    $inputSheet.Range("B40").Font.Name = "Arial"
    $inputSheet.Range("B40").Font.Bold = $true
    $inputSheet.Range("B40").Font.Color = Color-Ref "FFFFFF"
    $inputSheet.Range("B40").Interior.Color = Color-Ref "304F78"
    $inputSheet.Range("C40:G40").Merge()
    $inputSheet.Range("C40").Value2 = "Phiếu không dùng thì để trống"
    $inputSheet.Range("C40:G40").Font.Name = "Arial"
    $inputSheet.Range("C40:G40").Font.Italic = $true
    $inputSheet.Range("C40:G40").Interior.Color = Color-Ref "EAF0F7"
    $inputSheet.Range("J40").Value2 = "TaiSanID"
    $inputSheet.Range("K40").Value2 = "NgayCapSoGoc"
    $inputSheet.Range("L40").Value2 = "NgayCapSoTinh"
    $inputSheet.Range("M40").Value2 = "CoDuLieu"

    for ($assetIndex = 0; $assetIndex -lt $assetCardStartRows.Count; $assetIndex++) {
        $cardStartRow = $assetCardStartRows[$assetIndex]
        $cardHeader = $inputSheet.Range("B${cardStartRow}:G${cardStartRow}")
        $cardHeader.Merge()
        $cardHeader.Value2 = "TÀI SẢN $($assetIndex + 1)"
        $cardHeader.Font.Name = "Arial"
        $cardHeader.Font.Bold = $true
        $cardHeader.Font.Color = Color-Ref "FFFFFF"
        $cardHeader.Interior.Color = Color-Ref "17653D"
        $cardHeader.HorizontalAlignment = -4108

        for ($fieldIndex = 0; $fieldIndex -lt $assetFields.Count; $fieldIndex++) {
            $fieldRow = $cardStartRow + $fieldIndex + 1
            $labelCell = $inputSheet.Cells.Item($fieldRow, 2)
            $labelCell.Value2 = [string]$assetFields[$fieldIndex]
            $labelCell.Font.Name = "Arial"
            $labelCell.Font.Bold = $true
            $labelCell.Interior.Color = Color-Ref "F4F6F9"

            if ($assetLongFieldOffsets -contains ($fieldIndex + 1)) {
                $valueRange = $inputSheet.Range("C${fieldRow}:G${fieldRow}")
                $valueRange.Merge()
            }
            else {
                $valueRange = $inputSheet.Cells.Item($fieldRow, 3)
            }
            $valueRange.Font.Name = "Arial"
            $valueRange.Interior.Color = Color-Ref "FFFFFF"
            $valueRange.Borders.Color = Color-Ref "7AA7D8"
            $valueRange.Borders.Weight = 1
            Release-ComObject $labelCell
            Release-ComObject $valueRange
        }

        $inputSheet.Range("B${cardStartRow}:G$($cardStartRow + 18)").Borders.Color = Color-Ref "7AA7D8"
        $inputSheet.Range("B${cardStartRow}:G$($cardStartRow + 18)").Borders.Weight = 1
        $inputSheet.Rows.Item($cardStartRow).RowHeight = 22
        Release-ComObject $cardHeader
    }
    $inputSheet.Range("C57,C77,C97").NumberFormat = "@"
    $inputSheet.Columns.Item("B").ColumnWidth = 25
    $inputSheet.Columns.Item("C").ColumnWidth = 24
    $inputSheet.Columns.Item("D:G").ColumnWidth = 12

    foreach ($sheetToLock in @($inputSheet, $configSheet, $checkSheet, $exportSheet, $catalogSheet)) {
        $sheetToLock.Cells.Locked = $true
    }
    $inputSheet.Range("B4").Locked = $false
    $inputSheet.Range("B9:G38").Locked = $false
    foreach ($cardStartRow in $assetCardStartRows) {
        for ($fieldIndex = 0; $fieldIndex -lt $assetFields.Count; $fieldIndex++) {
            $fieldRow = $cardStartRow + $fieldIndex + 1
            if ($assetLongFieldOffsets -contains ($fieldIndex + 1)) {
                $inputSheet.Range("C${fieldRow}:G${fieldRow}").Locked = $false
            }
            else {
                $inputSheet.Cells.Item($fieldRow, 3).Locked = $false
            }
        }
    }
    $inputSheet.Range("B9:G38").Borders.Color = Color-Ref "7AA7D8"
    $inputSheet.Range("B9:G38").Borders.Weight = 1

    $inputSheet.Activate()
    $inputSheet.Range("A9").Select()
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.DisplayGridlines = $false
    $excel.ActiveWindow.Zoom = 90

    $buttonAnchor = $inputSheet.Range("H3:I4")
    $exportButton = $inputSheet.Shapes.AddShape(5, $buttonAnchor.Left, $buttonAnchor.Top, $buttonAnchor.Width, $buttonAnchor.Height)
    $exportButton.Name = "btnXuatVanBan"
    $exportButton.TextFrame2.TextRange.Text = "Xuất Văn bản"
    $exportButton.TextFrame2.TextRange.Font.Name = "Arial"
    $exportButton.TextFrame2.TextRange.Font.Size = 12
    $exportButton.TextFrame2.TextRange.Font.Bold = $true
    $exportButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color-Ref "FFFFFF"
    $exportButton.Fill.ForeColor.RGB = Color-Ref "217346"
    $exportButton.Line.ForeColor.RGB = Color-Ref "155F38"
    $exportButton.OnAction = "XuatVanBan"

    $inputSheet.PageSetup.Orientation = 2
    $inputSheet.PageSetup.Zoom = $false
    $inputSheet.PageSetup.FitToPagesWide = 1
    $inputSheet.PageSetup.FitToPagesTall = $false
    $inputSheet.PageSetup.PrintArea = '$A$1:$I$100'
    $inputSheet.Tab.Color = Color-Ref "217346"

    $configSheet.Range("A1:B1").Merge()
    $configSheet.Range("A1").Value2 = "CẤU HÌNH HỒ SƠ"
    $configSheet.Range("A1:B1").Font.Name = "Arial"
    $configSheet.Range("A1:B1").Font.Size = 16
    $configSheet.Range("A1:B1").Font.Bold = $true
    $configSheet.Range("A1:B1").Font.Color = Color-Ref "FFFFFF"
    $configSheet.Range("A1:B1").Interior.Color = Color-Ref "304F78"
    $configSheet.Range("A3:A10").Font.Name = "Arial"
    $configSheet.Range("A3:A10").Font.Bold = $true
    $configSheet.Range("B3").Formula = "=HangTKToiDa"
    $configSheet.Range("B4").Value2 = "PALETTE_B_TUONG_PHAN"
    $configSheet.Range("B5").Value2 = "2.0.0"
    $configSheet.Range("B6").Value2 = 1
    $configSheet.Range("B7").Value2 = 100
    $configSheet.Range("B8").NumberFormat = "@"
    $configSheet.Range("B8").Value = [string]$testTemplate
    $configSheet.Range("B9").NumberFormat = "@"
    $configSheet.Range("B9").Value = [string]$wordOutputDir
    $configSheet.Range("B10").Value2 = 1
    $configSheet.Range("B3:B10").Font.Name = "Arial"
    $configSheet.Range("B3:B10").Interior.Color = Color-Ref "F4F6F9"
    $configSheet.Columns.Item("A").ColumnWidth = 24
    $configSheet.Columns.Item("B").ColumnWidth = 75
    $configSheet.Range("A3:B10").Borders.Color = Color-Ref "D8DEE8"

    $workbook.Names.Add("HangTKToiDa", '=NhapLieu!$B$4') | Out-Null
    $workbook.Names.Add("BangMauHangTK", '=CauHinh!$B$4') | Out-Null
    $workbook.Names.Add("PhienBanCauTruc", '=CauHinh!$B$5') | Out-Null
    $workbook.Names.Add("NguoiIDTiepTheo", '=CauHinh!$B$6') | Out-Null
    $workbook.Names.Add("SucChuaNguoi", '=CauHinh!$B$7') | Out-Null
    $workbook.Names.Add("DuongDanMauWord", '=CauHinh!$B$8') | Out-Null
    $workbook.Names.Add("ThuMucXuat", '=CauHinh!$B$9') | Out-Null
    $workbook.Names.Add("TaiSanIDTiepTheo", '=CauHinh!$B$10') | Out-Null

    $checkSheet.Range("A1:F1").Merge()
    $checkSheet.Range("A1").Value2 = "KẾT QUẢ KIỂM TRA DỮ LIỆU"
    $checkSheet.Range("A1:F1").Font.Name = "Arial"
    $checkSheet.Range("A1:F1").Font.Size = 16
    $checkSheet.Range("A1:F1").Font.Bold = $true
    $checkSheet.Range("A1:F1").Font.Color = Color-Ref "FFFFFF"
    $checkSheet.Range("A1:F1").Interior.Color = Color-Ref "304F78"
    $checkSheet.Range("A2").Value2 = "Số lỗi chặn"
    $checkSheet.Range("A3").Value2 = "Số cảnh báo"
    $checkSheet.Range("B2:B3").Value2 = 0
    $checkHeaders = @("Mức", "Mã lỗi", "STT", "Họ tên", "Nội dung", "Ô cần sửa")
    for ($headerIndex = 0; $headerIndex -lt $checkHeaders.Count; $headerIndex++) {
        $checkSheet.Cells.Item(4, $headerIndex + 1).Value = [string]$checkHeaders[$headerIndex]
    }
    $checkSheet.Range("A4:F4").Font.Bold = $true
    $checkSheet.Range("A4:F4").Font.Color = Color-Ref "FFFFFF"
    $checkSheet.Range("A4:F4").Interior.Color = Color-Ref "738096"
    $checkSheet.Columns.Item("A").ColumnWidth = 12
    $checkSheet.Columns.Item("B").ColumnWidth = 24
    $checkSheet.Columns.Item("C").ColumnWidth = 9
    $checkSheet.Columns.Item("D").ColumnWidth = 24
    $checkSheet.Columns.Item("E").ColumnWidth = 70
    $checkSheet.Columns.Item("F").ColumnWidth = 20
    $checkSheet.Columns.Item("E").WrapText = $true
    $checkSheet.Tab.Color = Color-Ref "C85C70"

    $exportSheet.Range("A1:V1").Merge()
    $exportSheet.Range("A1").Value2 = "DỮ LIỆU XUẤT — KHÔNG NHẬP TAY"
    $exportSheet.Range("A1:V1").Font.Name = "Arial"
    $exportSheet.Range("A1:V1").Font.Bold = $true
    $exportSheet.Range("A1:V1").Font.Color = Color-Ref "FFFFFF"
    $exportSheet.Range("A1:V1").Interior.Color = Color-Ref "304F78"
    $exportSheet.Tab.Color = Color-Ref "7B61B3"

    $vbProject = $workbook.VBProject
    Add-StandardModule $vbProject "modCommon" (Join-Path $vbaRoot "modCommon.bas")
    Add-StandardModule $vbProject "modNgayThang" (Join-Path $vbaRoot "modNgayThang.bas")
    Add-StandardModule $vbProject "modNguoi" (Join-Path $vbaRoot "modNguoi.bas")
    Add-StandardModule $vbProject "modTaiSan" (Join-Path $vbaRoot "modTaiSan.bas")
    Add-StandardModule $vbProject "modHangTKNhanh" (Join-Path $vbaRoot "modHangTKNhanh.bas")
    Add-StandardModule $vbProject "modValidation" (Join-Path $vbaRoot "modValidation.bas")
    Add-StandardModule $vbProject "modExportData" (Join-Path $vbaRoot "modExportData.bas")
    Add-StandardModule $vbProject "modWordExport" (Join-Path $vbaRoot "modWordExport.bas")
    Set-DocumentModuleCode $vbProject.VBComponents.Item("ThisWorkbook") (Join-Path $vbaRoot "ThisWorkbook.cls")
    Set-DocumentModuleCode $vbProject.VBComponents.Item($inputSheet.CodeName) (Join-Path $vbaRoot "NhapLieu.cls")

    foreach ($sheetToProtect in @($inputSheet, $configSheet, $checkSheet, $exportSheet, $catalogSheet)) {
        $sheetToProtect.Protect("HoSoTK_MVP_2026", $true, $true, $true, $true, $false, $false, $false, $false, $false, $false, $false, $false, $false, $true, $false)
        $sheetToProtect.EnableSelection = -4142
    }

    $configSheet.Visible = 0
    $checkSheet.Visible = 0
    $catalogSheet.Visible = 0
    $exportSheet.Visible = 2
    $inputSheet.Activate()
    $inputSheet.Range("B9").Select()

    $workbook.SaveAs($targetWorkbook, 52)
    $workbook.Close($true)
    Release-ComObject $workbook
    $workbook = $null

    Copy-Item -LiteralPath $targetWorkbook -Destination $testWorkbook
    $excel.AutomationSecurity = 1
    $excel.EnableEvents = $false
    $testBook = $excel.Workbooks.Open($testWorkbook)
    $testInput = $testBook.Worksheets.Item("NhapLieu")
    $testTable = $testInput.ListObjects.Item("tblNguoi")
    $testInput.Unprotect("HoSoTK_MVP_2026")

    $fakePeople = @(
        @("Ông Nguyễn Văn A", "1950", "1999", "001050000001", "2021", "Thành phố Hà Nội", 0, 0),
        @("Bà Trần Thị B", "02/1952", "11/06/2021", "001052000002", "02/2021", "Thành phố Hà Nội", 0, 0),
        @("Anh Nguyễn Văn C", "03/03/1975", "07/2022", "001075000003", [datetime]"2022-03-03", "Tỉnh Ninh Bình", 1, 0),
        @("Chị Nguyễn Thị C1", "2000", $null, "001200000004", "04/2023", "Tỉnh Ninh Bình", 2, 1),
        @("Anh Nguyễn Văn C2", "05/2002", $null, "001202000005", "05/05/2023", "Tỉnh Ninh Bình", 2, 0),
        @("Chị Nguyễn Thị D", [datetime]"1978-06-06", $null, "001078000006", "2022", "Tỉnh Nam Định", 1, 1)
    )

    for ($i = 0; $i -lt $fakePeople.Count; $i++) {
        $rowRange = $testTable.DataBodyRange.Rows.Item($i + 1)
        $rowRange.Cells.Item(1, 2).Value = [string]$fakePeople[$i][0]
        if ($fakePeople[$i][1] -is [datetime]) {
            $rowRange.Cells.Item(1, 3).Value2 = [double]$fakePeople[$i][1].ToOADate()
        }
        else {
            $rowRange.Cells.Item(1, 3).Value = [string]$fakePeople[$i][1]
        }
        if ($null -ne $fakePeople[$i][2]) {
            if ($fakePeople[$i][2] -is [datetime]) {
                $rowRange.Cells.Item(1, 4).Value2 = [double]$fakePeople[$i][2].ToOADate()
            }
            else {
                $rowRange.Cells.Item(1, 4).Value = [string]$fakePeople[$i][2]
            }
        }
        $rowRange.Cells.Item(1, 5).NumberFormat = "@"
        $rowRange.Cells.Item(1, 5).Value = [string]$fakePeople[$i][3]
        if ($fakePeople[$i][4] -is [datetime]) {
            $rowRange.Cells.Item(1, 6).Value2 = [double]$fakePeople[$i][4].ToOADate()
        }
        else {
            $rowRange.Cells.Item(1, 6).Value = [string]$fakePeople[$i][4]
        }
        $rowRange.Cells.Item(1, 7).Value = [string]$fakePeople[$i][5]
        $rowRange.Cells.Item(1, 8).Value = [int]$fakePeople[$i][6]
        $rowRange.Cells.Item(1, 9).Value = [int]$fakePeople[$i][7]
        Release-ComObject $rowRange
    }

    $macroPrefix = "'$($testBook.Name)'!"
    [void]$excel.Run($macroPrefix + "ApplyWorkbookProtection")
    [void]$excel.Run($macroPrefix + "RefreshAllPeople")

    if ([string]$testInput.Range("C9").Value2 -ne "1950") {
        throw "Năm sinh 1950 không được giữ nguyên trên giao diện."
    }
    if ([string]$testInput.Range("D9").Value2 -ne "1999") {
        throw "Năm chết 1999 không được giữ nguyên trên giao diện."
    }
    if ([string]$testInput.Range("O9").Value2 -ne "1950") { throw "Dữ liệu gốc năm sinh không đúng." }
    if ([string]$testInput.Range("Q9").Value2 -ne "1999") { throw "Dữ liệu gốc năm chết không đúng." }
    if ([string]$testInput.Range("C10").Value2 -ne "02/1952") { throw "Tháng/năm không được giữ nguyên trên giao diện." }
    $ageAtDeath = [int]$excel.Run($macroPrefix + "AgeAtDeath", 1)
    if ($ageAtDeath -ne 49) { throw "Tính tuổi từ ngày chuẩn hóa không đúng." }

    $testInput.Range("C63").Value2 = "SERIAL-002"
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("C63"))
    if ([string]$testInput.Range("J61").Value2 -ne "TS001") { throw "Phiếu tài sản 2 chưa sinh TS001." }
    if (-not [bool]$testInput.Range("M61").Value2) { throw "Phiếu tài sản 2 chưa được đánh dấu có dữ liệu." }
    if ([string]$testInput.Range("J41").Value2 -ne "" -or [string]$testInput.Range("J81").Value2 -ne "") {
        throw "Phiếu tài sản trống lại có TaiSanID."
    }

    $testInput.Range("C42").Value2 = "GCN"
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("C42"))
    $assetDateExpectations = @{
        "2015" = [datetime]"2015-01-01"
        "05/2015" = [datetime]"2015-05-01"
        "12/05/2015" = [datetime]"2015-05-12"
    }
    foreach ($dateText in $assetDateExpectations.Keys) {
        $testInput.Range("C57").Value2 = [string]$dateText
        [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("C57"))
        if ([string]$testInput.Range("C57").Value2 -ne [string]$dateText -or
            [string]$testInput.Range("K41").Value2 -ne [string]$dateText -or
            [math]::Abs(([double]$testInput.Range("L41").Value2) - $assetDateExpectations[$dateText].ToOADate()) -gt 0.0001) {
            throw "Ngày cấp sổ '$dateText' không được giữ nguyên và chuẩn hóa đúng."
        }
    }

    $testInput.Range("C63").ClearContents()
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("C63"))
    if ([string]$testInput.Range("J61").Value2 -ne "" -or [string]$testInput.Range("K61").Value2 -ne "" -or
        [string]$testInput.Range("L61").Value2 -ne "" -or [bool]$testInput.Range("M61").Value2) {
        throw "Xóa trắng phiếu tài sản chưa xóa dữ liệu hệ thống."
    }
    if (-not [bool]$testInput.Range("J:W").EntireColumn.Hidden) { throw "Vùng dữ liệu tài sản J:W chưa được ẩn." }
    if ([string]$testBook.Names.Item("TaiSanIDTiepTheo").RefersToRange.Address($false, $false, 1, $true) -notmatch "B10") {
        throw "Bộ đếm TaiSanIDTiepTheo chưa trỏ đến CauHinh!B10."
    }

    $originalAddress = [string]$testInput.Range("G14").Value2
    $editTimer = [System.Diagnostics.Stopwatch]::StartNew()
    for ($editIndex = 1; $editIndex -le 20; $editIndex++) {
        $testInput.Range("G14").Value2 = "Kiểm tra tốc độ $editIndex"
        [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("G14"))
    }
    $editTimer.Stop()
    $testInput.Range("G14").Value = [string]$originalAddress
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("G14"))
    if ($editTimer.ElapsedMilliseconds -gt 3000) {
        throw "Xử lý 20 lần nhập ô còn chậm: $($editTimer.ElapsedMilliseconds) ms."
    }
    Write-Host "INPUT_EDIT_20X_MS=$($editTimer.ElapsedMilliseconds)"

    if (-not [bool]$testInput.ProtectContents) { throw "Sheet nhập liệu chưa được bảo vệ." }
    if (-not [bool]$testInput.Range("A9").Locked) { throw "Ô hệ thống A9 chưa được khóa." }
    if ([bool]$testInput.Range("B9").Locked) { throw "Ô nhập liệu B9 đang bị khóa nhầm." }
    if (-not [bool]$testInput.Range("H9").Locked) { throw "Ô hành động H9 chưa được khóa nhập tay." }

    $testInput.Range("B38").Value2 = "QA vùng nhập"
    $testInput.Range("B38").ClearContents()

    [void]$excel.Run($macroPrefix + "HandleLevelClick", $testInput.Range("H14"))
    if ([int]$testInput.Range("H14").Value2 -ne 2) { throw "Nút Hàng TK không hoạt động khi sheet được bảo vệ." }
    [void]$excel.Run($macroPrefix + "HandleLevelClick", $testInput.Range("H14"))
    if ([int]$testInput.Range("H14").Value2 -ne 1) { throw "Vòng Hàng TK 1-2 không đúng." }

    [void]$excel.Run($macroPrefix + "HandleReceiveClick", $testInput.Range("I14"))
    [void]$excel.Run($macroPrefix + "HandleReceiveClick", $testInput.Range("I14"))
    if ([int]$testInput.Range("I14").Value2 -ne 1) { throw "Nút Nhận đất không hoạt động khi sheet được bảo vệ." }

    $testInput.Range("C14").Value2 = "31/02/1990"
    [void]$excel.Run($macroPrefix + "NormalizeAllDates", $false)
    $invalidWasBlocked = -not [bool]$excel.Run($macroPrefix + "ValidateWorkbook", $false)
    if (-not $invalidWasBlocked) { throw "Ngày 31/02/1990 chưa bị chặn." }
    $testInput.Range("C14").Value2 = "1978"
    [void]$excel.Run($macroPrefix + "RefreshAllPeople")

    $isValid = [bool]$excel.Run($macroPrefix + "ValidateWorkbook", $false)
    if (-not $isValid) { throw "Bộ dữ liệu giả không vượt qua validation." }
    [void]$excel.Run($macroPrefix + "BuildExportData")
    $exportSucceeded = [bool]$excel.Run($macroPrefix + "RunWordExport", $true)
    if (-not $exportSucceeded) { throw "Xuất Word thử thất bại." }
    $testBook.Save()
    $testBook.Close($true)
    Release-ComObject $testBook
    $testBook = $null

    Create-QaCopy -Excel $excel -SourcePath $targetWorkbook -DestinationPath $qaCleanWorkbook
    Create-QaCopy -Excel $excel -SourcePath $testWorkbook -DestinationPath $qaTestWorkbook
}
finally {
    if ($null -ne $testBook) {
        try { $testBook.Close($false) } catch {}
        Release-ComObject $testBook
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    if ($null -ne $sourceBook) {
        try { $sourceBook.Close($false) } catch {}
        Release-ComObject $sourceBook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$sourceHashAfter = (Get-FileHash -Algorithm SHA256 -LiteralPath $sourceWorkbook).Hash
if ($sourceHashBefore -ne $sourceHashAfter) {
    throw "Workbook nguồn đã thay đổi ngoài ý muốn."
}

Write-Output "SOURCE_SHA256=$sourceHashAfter"
Write-Output "TARGET=$targetWorkbook"
Write-Output "TEST=$testWorkbook"
Write-Output "QA_CLEAN=$qaCleanWorkbook"
Write-Output "QA_TEST=$qaTestWorkbook"
Write-Output "WORD_TEMPLATE=$testTemplate"
Write-Output "WORD_OUTPUT_DIR=$wordOutputDir"
