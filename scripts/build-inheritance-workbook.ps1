param(
    [string]$WorkspaceRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$Version = "0.3.0"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$officeCli = Join-Path $env:LOCALAPPDATA "OfficeCLI\officecli.exe"
if (-not (Test-Path -LiteralPath $officeCli)) {
    throw "Không tìm thấy OfficeCLI tại $officeCli"
}

$workspace = (Resolve-Path -LiteralPath $WorkspaceRoot).Path
$placeholderReport = Join-Path $workspace "docs\reference\word-placeholders-2026-08-31.md"
$catalogDataPath = Join-Path $workspace "src\data\danh-muc.json"
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

function Write-ColumnValues {
    param([object]$Sheet, [int]$Column, [object[]]$Values)
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $Sheet.Cells.Item($i + 4, $Column).Value2 = [string]$Values[$i]
    }
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
        $hasNiemYet = $false
        $hasUyQuyen = $false
        if ($sectionStart -ge 0) {
            for ($detailIndex = $sectionStart + 1; $detailIndex -lt $lines.Count; $detailIndex++) {
                if ($lines[$detailIndex] -match '^## ') { break }
                $nameMatches = [regex]::Matches($lines[$detailIndex], '\[Tên (\d+)\]')
                if ($lines[$detailIndex] -match '\[(Niêm Yết|Niem Yet)\]') { $hasNiemYet = $true }
                if ($lines[$detailIndex] -match '\[(Người ủy quyền|Nguoi uy quyen)\]') { $hasUyQuyen = $true }
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
            CoNiemYet = $hasNiemYet
            CoNguoiUyQuyen = $hasUyQuyen
        }
    }
    return $rows
}

if (-not (Test-Path -LiteralPath $placeholderReport)) {
    throw "Không tìm thấy báo cáo placeholder để dựng sức chứa mẫu: $placeholderReport"
}
if (-not (Test-Path -LiteralPath $catalogDataPath)) {
    throw "Không tìm thấy dữ liệu danh mục đã chốt: $catalogDataPath"
}
$catalogData = (Get-Content -Raw -Encoding UTF8 -LiteralPath $catalogDataPath | ConvertFrom-Json)

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
    @{ command = "set"; path = "/CauHinh/A1"; props = @{ value = "CẤU HÌNH HỒ SƠ" } },
    @{ command = "set"; path = "/CauHinh/A3"; props = @{ value = "HangTKToiDa" } },
    @{ command = "set"; path = "/CauHinh/A4"; props = @{ value = "BangMauHangTK" } },
    @{ command = "set"; path = "/CauHinh/A5"; props = @{ value = "PhienBanCauTruc" } },
    @{ command = "set"; path = "/CauHinh/A6"; props = @{ value = "NguoiIDTiepTheo" } },
    @{ command = "set"; path = "/CauHinh/A7"; props = @{ value = "SucChuaNguoiMau" } },
    @{ command = "set"; path = "/CauHinh/A8"; props = @{ value = "MauWordDangChon" } },
    @{ command = "set"; path = "/CauHinh/A9"; props = @{ value = "ThuMucXuat" } },
    @{ command = "set"; path = "/CauHinh/A10"; props = @{ value = "TaiSanIDTiepTheo" } },
    @{ command = "set"; path = "/CauHinh/A11"; props = @{ value = "ThuMucMauWord" } },
    @{ command = "set"; path = "/CauHinh/A12"; props = @{ value = "SucChuaTaiSanMau" } },
    @{ command = "set"; path = "/KiemTra/A1"; props = @{ value = "KẾT QUẢ KIỂM TRA DỮ LIỆU" } },
    @{ command = "set"; path = "/KiemTra/A2"; props = @{ value = "Số lỗi chặn" } },
    @{ command = "set"; path = "/KiemTra/A3"; props = @{ value = "Số cảnh báo" } },
    @{ command = "set"; path = "/XuatAn/A1"; props = @{ value = "DỮ LIỆU XUẤT — KHÔNG NHẬP TAY" } }
)
Invoke-OfficeCliBatch -FilePath $baseWorkbook -Operations $baseOperations
Invoke-OfficeCli @("add", $baseWorkbook, "/NhapLieu", "--type", "table", "--prop", "ref=A8:I38", "--prop", "name=tblNguoi", "--prop", "displayName=tblNguoi", "--prop", "style=light1", "--prop", "showRowStripes=false")
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
    try {
        $catalogSheet = $workbook.Worksheets.Item("DanhMuc")
    }
    catch {
        $catalogSheet = $workbook.Worksheets.Add()
        $catalogSheet.Name = "DanhMuc"
    }
    $peopleTable = $inputSheet.ListObjects.Item("tblNguoi")
    $peopleTable.Resize($inputSheet.Range("A8:I38"))
    $peopleTechHeaders = @(
        "STTNhap", "NguoiID", "ParentNguoiID", "LaChuDat", "NhomTuChoiID",
        "TrangThai", "NgaySinhGoc", "NgaySinhTinh", "NgayChetGoc", "NgayChetTinh",
        "NgayCapGoc", "NgayCapTinh", "LoaiCC", "NoiCapCC", "NhanDiaChi"
    )
    for ($techHeaderIndex = 0; $techHeaderIndex -lt $peopleTechHeaders.Count; $techHeaderIndex++) {
        $exportSheet.Cells.Item(1, 49 + $techHeaderIndex).Value2 = $peopleTechHeaders[$techHeaderIndex]
    }
    $peopleTechTable = $exportSheet.ListObjects.Add(1, $exportSheet.Range("AW1:BK31"), $null, 1)
    $peopleTechTable.Name = "tblNguoiKyThuat"
    $peopleTechTable.TableStyle = "TableStyleLight1"
    for ($techRow = 1; $techRow -le 30; $techRow++) {
        $peopleTechTable.DataBodyRange.Cells.Item($techRow, 1).Value2 = [double]$techRow
    }
    foreach ($rawHeader in @("NgaySinhGoc", "NgayChetGoc", "NgayCapGoc")) {
        $peopleTechTable.ListColumns.Item($rawHeader).DataBodyRange.NumberFormat = "@"
    }
    foreach ($calcHeader in @("NgaySinhTinh", "NgayChetTinh", "NgayCapTinh")) {
        $peopleTechTable.ListColumns.Item($calcHeader).DataBodyRange.NumberFormat = "dd/mm/yyyy"
    }
    $assetTechHeaders = @("TaiSanID", "NgayCapSoGoc", "NgayCapSoTinh", "CoDuLieu")
    for ($assetTechHeaderIndex = 0; $assetTechHeaderIndex -lt $assetTechHeaders.Count; $assetTechHeaderIndex++) {
        $exportSheet.Cells.Item(1, 65 + $assetTechHeaderIndex).Value2 = $assetTechHeaders[$assetTechHeaderIndex]
    }
    $exportSheet.Range("BN41:BN81").NumberFormat = "@"
    $exportSheet.Range("BO41:BO81").NumberFormat = "dd/mm/yyyy"

    Write-ColumnValues $catalogSheet 1 @($catalogData.LoaiSo)
    Write-ColumnValues $catalogSheet 2 @($catalogData.LoaiDat)
    Write-ColumnValues $catalogSheet 3 @($catalogData.HinhThucSuDung)
    Write-ColumnValues $catalogSheet 4 @($catalogData.CoQuanCapSo)
    Write-ColumnValues $catalogSheet 5 @($catalogData.LoaiGiayTo)
    Write-ColumnValues $catalogSheet 6 @($catalogData.NoiCapCC)
    Write-ColumnValues $catalogSheet 7 @($catalogData.NhanDiaChi)
    Write-ColumnValues $catalogSheet 8 @($catalogData.NguoiUyQuyen)

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
    $catalogSheet.Range("M3").Value2 = "CoNiemYet"
    $catalogSheet.Range("N3").Value2 = "CoNguoiUyQuyen"
    for ($capacityIndex = 0; $capacityIndex -lt $capacityRows.Count; $capacityIndex++) {
        $capacityRow = $capacityIndex + 4
        $catalogSheet.Cells.Item($capacityRow, 10).Value2 = [string]$capacityRows[$capacityIndex].TemplateName
        $catalogSheet.Cells.Item($capacityRow, 11).Value2 = [int]$capacityRows[$capacityIndex].SucChuaNguoi
        $catalogSheet.Cells.Item($capacityRow, 12).Value2 = [int]$capacityRows[$capacityIndex].SucChuaTaiSan
        $catalogSheet.Cells.Item($capacityRow, 13).Value2 = [bool]$capacityRows[$capacityIndex].CoNiemYet
        $catalogSheet.Cells.Item($capacityRow, 14).Value2 = [bool]$capacityRows[$capacityIndex].CoNguoiUyQuyen
    }
    $catalogSheet.Range("P3").Value2 = "Tờ bản đồ 2025"
    $catalogSheet.Range("Q3").Value2 = "Sáp nhập thôn"
    $catalogSheet.Range("P3:Q3").Font.Bold = $true
    $catalogSheet.Range("P3:Q3").Interior.Color = Color-Ref "FFF1B8"

    $catalogNamedRanges = @(
        @{ Name = "DanhMuc_LoaiSo"; Formula = ("=DanhMuc!`$A`$4:`$A`${0}" -f ($catalogData.LoaiSo.Count + 3)) },
        @{ Name = "DanhMuc_LoaiDat"; Formula = ("=DanhMuc!`$B`$4:`$B`${0}" -f ($catalogData.LoaiDat.Count + 3)) },
        @{ Name = "DanhMuc_HinhThucSuDung"; Formula = ("=DanhMuc!`$C`$4:`$C`${0}" -f ($catalogData.HinhThucSuDung.Count + 3)) },
        @{ Name = "DanhMuc_CoQuanCapSo"; Formula = ("=DanhMuc!`$D`$4:`$D`${0}" -f ($catalogData.CoQuanCapSo.Count + 3)) },
        @{ Name = "DanhMuc_LoaiGiayTo"; Formula = ("=DanhMuc!`$E`$4:`$E`${0}" -f ($catalogData.LoaiGiayTo.Count + 3)) },
        @{ Name = "DanhMuc_NoiCapCC"; Formula = ("=DanhMuc!`$F`$4:`$F`${0}" -f ($catalogData.NoiCapCC.Count + 3)) },
        @{ Name = "DanhMuc_NhanDiaChi"; Formula = ("=DanhMuc!`$G`$4:`$G`${0}" -f ($catalogData.NhanDiaChi.Count + 3)) },
        @{ Name = "DanhMuc_NguoiUyQuyen"; Formula = ("=DanhMuc!`$H`$4:`$H`${0}" -f ($catalogData.NguoiUyQuyen.Count + 3)) },
        @{ Name = "DanhMuc_SucChuaMau"; Formula = ("=DanhMuc!`$J`$4:`$L`${0}" -f ($capacityRows.Count + 3)) }
    )
    $loaiCCBeforeIndex = [array]::IndexOf([string[]]$catalogData.LoaiGiayTo, [string]$catalogData.LoaiCCTruocMoc) + 4
    $loaiCCFromIndex = [array]::IndexOf([string[]]$catalogData.LoaiGiayTo, [string]$catalogData.LoaiCCTuMoc) + 4
    $catalogNamedRanges += @{ Name = "DanhMuc_LoaiCCTruocMoc"; Formula = ("=DanhMuc!`$E`${0}" -f $loaiCCBeforeIndex) }
    $catalogNamedRanges += @{ Name = "DanhMuc_LoaiCCTuMoc"; Formula = ("=DanhMuc!`$E`${0}" -f $loaiCCFromIndex) }
    foreach ($namedRange in $catalogNamedRanges) {
        $workbook.Names.Add($namedRange.Name, $namedRange.Formula) | Out-Null
    }

    foreach ($assetValueColumn in @(28, 30, 32)) {
        $inputSheet.Cells.Item(9, $assetValueColumn).Validation.Delete()
        $inputSheet.Cells.Item(9, $assetValueColumn).Validation.Add(3, 1, 1, "=DanhMuc_LoaiSo")
        $inputSheet.Cells.Item(16, $assetValueColumn).Validation.Delete()
        $inputSheet.Cells.Item(16, $assetValueColumn).Validation.Add(3, 1, 1, "=DanhMuc_HinhThucSuDung")
        $inputSheet.Cells.Item(17, $assetValueColumn).Validation.Delete()
        $inputSheet.Cells.Item(17, $assetValueColumn).Validation.Add(3, 1, 1, "=DanhMuc_LoaiDat")
        $inputSheet.Cells.Item(25, $assetValueColumn).Validation.Delete()
        $inputSheet.Cells.Item(25, $assetValueColumn).Validation.Add(3, 1, 1, "=DanhMuc_CoQuanCapSo")
    }

    try {
        $workbook.BuiltinDocumentProperties.Item("Title").Value = "Hồ sơ thừa kế — MVP theo Hàng TK"
        $workbook.BuiltinDocumentProperties.Item("Comments").Value = "SOT: inheritance-branch-architecture.md; schema 2.2.0"
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

    $inputSheet.Range("A3").Value2 = "Nơi lấy mẫu"
    $inputSheet.Range("B3:E3").Merge()
    $inputSheet.Range("B3").NumberFormat = "@"
    $inputSheet.Range("B3").Value2 = [string]$outputDir
    $inputSheet.Range("A5").Value2 = "Mẫu Word"
    $inputSheet.Range("C5:E5").Merge()
    $inputSheet.Range("C5").NumberFormat = "@"
    $inputSheet.Range("C5").Value2 = [System.IO.Path]::GetFileName($testTemplate)
    $inputSheet.Range("A7").Value2 = "Nơi xuất"
    $inputSheet.Range("B7:G7").Merge()
    $inputSheet.Range("B7").NumberFormat = "@"
    $inputSheet.Range("B7").Value2 = [string]$wordOutputDir
    $inputSheet.Range("A3:A7").Font.Name = "Arial"
    $inputSheet.Range("A3:A7").Font.Bold = $true
    $inputSheet.Range("B3:E3,C5:E5,B7:G7").Interior.Color = Color-Ref "F4F6F9"
    $inputSheet.Range("B3:E3,C5:E5,B7:G7").Borders.Color = Color-Ref "D8DEE8"
    $inputSheet.Range("B3:E3,C5:E5,B7:G7").Borders.Weight = 1

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
        if ($row -le 2) { $dataRange.Cells.Item($row, 8).Value2 = 0 } else { $dataRange.Cells.Item($row, 8).Value2 = 1 }
        $dataRange.Cells.Item($row, 9).Value2 = 0
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
    $inputSheet.Range("J:Z").EntireColumn.Hidden = $true

    $inputSheet.Range("B9:B38").NumberFormat = "@"
    $inputSheet.Range("C9:D38").NumberFormat = "@"
    $inputSheet.Range("F9:F38").NumberFormat = "@"
    $inputSheet.Range("E9:E38").NumberFormat = "@"
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
    $assetCardStartColumns = @(27, 29, 31)
    $assetLongFieldOffsets = @(6, 15, 18)
    $profileLayoutCapRow = 70
    for ($assetIndex = 0; $assetIndex -lt $assetCardStartColumns.Count; $assetIndex++) {
        $cardStartColumn = [int]$assetCardStartColumns[$assetIndex]
        $inputSheet.Cells.Item(6, $cardStartColumn).Value2 = "THÔNG TIN TÀI SẢN"
        $inputSheet.Cells.Item(7, $cardStartColumn).Value2 = "Thẻ không dùng thì để trống"
        $inputSheet.Range($inputSheet.Cells.Item(6, $cardStartColumn), $inputSheet.Cells.Item(7, $cardStartColumn + 1)).Font.Name = "Arial"
        $inputSheet.Range($inputSheet.Cells.Item(6, $cardStartColumn), $inputSheet.Cells.Item(7, $cardStartColumn + 1)).WrapText = $true
        $inputSheet.Range($inputSheet.Cells.Item(6, $cardStartColumn), $inputSheet.Cells.Item(7, $cardStartColumn + 1)).Interior.Color = Color-Ref "EAF0F7"
        $cardHeader = $inputSheet.Range($inputSheet.Cells.Item(8, $cardStartColumn), $inputSheet.Cells.Item(8, $cardStartColumn + 1))
        $inputSheet.Cells.Item(8, $cardStartColumn).Value2 = "TÀI SẢN $($assetIndex + 1)"
        $cardHeader.Font.Name = "Arial"
        $cardHeader.Font.Bold = $true
        $cardHeader.Font.Color = Color-Ref "FFFFFF"
        $cardHeader.Interior.Color = Color-Ref "17653D"
        $cardHeader.HorizontalAlignment = -4108

        for ($fieldIndex = 0; $fieldIndex -lt $assetFields.Count; $fieldIndex++) {
            $fieldRow = $fieldIndex + 9
            $labelCell = $inputSheet.Cells.Item($fieldRow, $cardStartColumn)
            $labelCell.Value2 = [string]$assetFields[$fieldIndex]
            $labelCell.Font.Name = "Arial"
            $labelCell.Font.Bold = $true
            $labelCell.Interior.Color = Color-Ref "F4F6F9"

            $valueRange = $inputSheet.Cells.Item($fieldRow, $cardStartColumn + 1)
            $valueRange.Font.Name = "Arial"
            $valueRange.Interior.Color = Color-Ref "FFFFFF"
            $valueRange.Borders.Color = Color-Ref "7AA7D8"
            $valueRange.Borders.Weight = 1
            Release-ComObject $labelCell
            Release-ComObject $valueRange
        }

        $inputSheet.Range($inputSheet.Cells.Item(8, $cardStartColumn), $inputSheet.Cells.Item(26, $cardStartColumn + 1)).Borders.Color = Color-Ref "7AA7D8"
        $inputSheet.Range($inputSheet.Cells.Item(8, $cardStartColumn), $inputSheet.Cells.Item(26, $cardStartColumn + 1)).Borders.Weight = 1
        $inputSheet.Rows.Item(8).RowHeight = 22
        Release-ComObject $cardHeader
    }
    $inputSheet.Range("AB9:AB38,AD9:AD38,AF9:AF38").NumberFormat = "@"
    $inputSheet.Range("AB14,AB23,AB26,AD14,AD23,AD26,AF14,AF23,AF26").WrapText = $true
    $inputSheet.Columns.Item("AA").ColumnWidth = 18
    $inputSheet.Columns.Item("AB").ColumnWidth = 22
    $inputSheet.Columns.Item("AC").ColumnWidth = 18
    $inputSheet.Columns.Item("AD").ColumnWidth = 22
    $inputSheet.Columns.Item("AE").ColumnWidth = 18
    $inputSheet.Columns.Item("AF").ColumnWidth = 22
    $inputSheet.Range("J8:Z38").NumberFormat = "@"
    for ($extensionColumn = 10; $extensionColumn -le 26; $extensionColumn++) {
        $inputSheet.Columns.Item($extensionColumn).ColumnWidth = 18
    }

    foreach ($sheetToLock in @($inputSheet, $configSheet, $checkSheet, $exportSheet, $catalogSheet)) {
        $sheetToLock.Cells.Locked = $true
    }
    $inputSheet.Range("B4").Locked = $false
    $inputSheet.Range("B9:G38").Locked = $false
    $inputSheet.Range("J8:Z38").Locked = $false
    # The profile is open-ended at runtime. Keep the first 30 input rows
    # formatted for the normal layout; ProfileLastDataRow scans further rows
    # when a user adds them, while the page layout remains capped at row 70.
    $inputSheet.Range("B45:B1048576").Locked = $false
    $inputSheet.Range("C41:C1048576").Locked = $false
    foreach ($cardStartColumn in $assetCardStartColumns) {
        for ($fieldIndex = 0; $fieldIndex -lt $assetFields.Count; $fieldIndex++) {
            $fieldRow = $fieldIndex + 9
            $inputSheet.Cells.Item($fieldRow, ([int]$cardStartColumn + 1)).Locked = $false
        }
    }
    $inputSheet.Range("B40").Value2 = "THÔNG TIN HỒ SƠ"
    $inputSheet.Range("B40:C40").Font.Name = "Arial"
    $inputSheet.Range("B40:C40").Font.Bold = $true
    $inputSheet.Range("B40:C40").Font.Color = Color-Ref "FFFFFF"
    $inputSheet.Range("B40:C40").Interior.Color = Color-Ref "304F78"
    $inputSheet.Range("B40:C40").Borders.Color = Color-Ref "7AA7D8"
    $inputSheet.Range("B40:C40").Borders.Weight = 1
    $profileFields = @(
        @{ Row = 41; Name = "NiemYet"; Label = "Niêm yết" },
        @{ Row = 42; Name = "SoCongChung"; Label = "Số công chứng" },
        @{ Row = 43; Name = "NguoiUyQuyen"; Label = "Người ủy quyền" },
        @{ Row = 44; Name = "NguoiUyQuyen2"; Label = "Người ủy quyền 2" }
    )
    foreach ($profileField in $profileFields) {
        $inputSheet.Cells.Item([int]$profileField.Row, 2).Value2 = [string]$profileField.Label
        $inputSheet.Cells.Item($profileField.Row, 2).Font.Name = "Arial"
        $inputSheet.Cells.Item($profileField.Row, 2).Font.Bold = $true
        $inputSheet.Cells.Item($profileField.Row, 2).Interior.Color = Color-Ref "F4F6F9"
        $inputSheet.Cells.Item($profileField.Row, 3).Font.Name = "Arial"
        $inputSheet.Cells.Item($profileField.Row, 3).NumberFormat = "@"
        $inputSheet.Cells.Item($profileField.Row, 3).Interior.Color = Color-Ref "FFFFFF"
        $inputSheet.Range("B$($profileField.Row):C$($profileField.Row)").Borders.Color = Color-Ref "7AA7D8"
        $inputSheet.Range("B$($profileField.Row):C$($profileField.Row)").Borders.Weight = 1
    }
    foreach ($profileRow in @(43, 44)) {
        $inputSheet.Cells.Item($profileRow, 3).Validation.Delete()
        $inputSheet.Cells.Item($profileRow, 3).Validation.Add(3, 1, 1, "=DanhMuc_NguoiUyQuyen")
    }
    for ($profileRow = 45; $profileRow -le $profileLayoutCapRow; $profileRow++) {
        $inputSheet.Cells.Item($profileRow, 2).Font.Name = "Arial"
        $inputSheet.Cells.Item($profileRow, 2).Interior.Color = Color-Ref "F4F6F9"
        $inputSheet.Cells.Item($profileRow, 3).Font.Name = "Arial"
        $inputSheet.Cells.Item($profileRow, 3).NumberFormat = "@"
        $inputSheet.Cells.Item($profileRow, 3).Interior.Color = Color-Ref "FFFFFF"
        $inputSheet.Range("B$profileRow:C$profileRow").Borders.Color = Color-Ref "7AA7D8"
        $inputSheet.Range("B$profileRow:C$profileRow").Borders.Weight = 1
    }
    $inputSheet.Range("C41:C1048576").NumberFormat = "@"
    $inputSheet.Range("B40:C$profileLayoutCapRow").Borders.Color = Color-Ref "7AA7D8"
    $inputSheet.Range("B40:C$profileLayoutCapRow").Borders.Weight = 1
    $inputSheet.Columns.Item("B").ColumnWidth = 25
    $inputSheet.Columns.Item("C").ColumnWidth = 24
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

    $pickFolderAnchor = $inputSheet.Range("F3:G3")
    $pickFolderButton = $inputSheet.Shapes.AddShape(5, $pickFolderAnchor.Left, $pickFolderAnchor.Top, $pickFolderAnchor.Width, $pickFolderAnchor.Height)
    $pickFolderButton.Name = "btnChonNoiLayMau"
    $pickFolderButton.TextFrame2.TextRange.Text = "Chọn nơi lấy mẫu"
    $pickFolderButton.TextFrame2.TextRange.Font.Name = "Arial"
    $pickFolderButton.TextFrame2.TextRange.Font.Size = 10
    $pickFolderButton.TextFrame2.TextRange.Font.Bold = $true
    $pickFolderButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color-Ref "FFFFFF"
    $pickFolderButton.Fill.ForeColor.RGB = Color-Ref "2F75B5"
    $pickFolderButton.Line.ForeColor.RGB = Color-Ref "205493"
    $pickFolderButton.OnAction = "ChonNoiLayMau"

    $pickFileAnchor = $inputSheet.Range("F5:G5")
    $pickFileButton = $inputSheet.Shapes.AddShape(5, $pickFileAnchor.Left, $pickFileAnchor.Top, $pickFileAnchor.Width, $pickFileAnchor.Height)
    $pickFileButton.Name = "btnChonVanBan"
    $pickFileButton.TextFrame2.TextRange.Text = "Chọn văn bản"
    $pickFileButton.TextFrame2.TextRange.Font.Name = "Arial"
    $pickFileButton.TextFrame2.TextRange.Font.Size = 10
    $pickFileButton.TextFrame2.TextRange.Font.Bold = $true
    $pickFileButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color-Ref "FFFFFF"
    $pickFileButton.Fill.ForeColor.RGB = Color-Ref "2F75B5"
    $pickFileButton.Line.ForeColor.RGB = Color-Ref "205493"
    $pickFileButton.OnAction = "ChonVanBan"

    $syncAnchor = $inputSheet.Range("H5:I5")
    $syncButton = $inputSheet.Shapes.AddShape(5, $syncAnchor.Left, $syncAnchor.Top, $syncAnchor.Width, $syncAnchor.Height)
    $syncButton.Name = "btnDongBoTruong"
    $syncButton.TextFrame2.TextRange.Text = "Đồng bộ trường"
    $syncButton.TextFrame2.TextRange.Font.Name = "Arial"
    $syncButton.TextFrame2.TextRange.Font.Size = 10
    $syncButton.TextFrame2.TextRange.Font.Bold = $true
    $syncButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color-Ref "FFFFFF"
    $syncButton.Fill.ForeColor.RGB = Color-Ref "2F75B5"
    $syncButton.Line.ForeColor.RGB = Color-Ref "205493"
    $syncButton.OnAction = "DongBoTruong"

    $pickOutputAnchor = $inputSheet.Range("H7:I7")
    $pickOutputButton = $inputSheet.Shapes.AddShape(5, $pickOutputAnchor.Left, $pickOutputAnchor.Top, $pickOutputAnchor.Width, $pickOutputAnchor.Height)
    $pickOutputButton.Name = "btnChonNoiXuat"
    $pickOutputButton.TextFrame2.TextRange.Text = "Chọn nơi xuất"
    $pickOutputButton.TextFrame2.TextRange.Font.Name = "Arial"
    $pickOutputButton.TextFrame2.TextRange.Font.Size = 10
    $pickOutputButton.TextFrame2.TextRange.Font.Bold = $true
    $pickOutputButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color-Ref "FFFFFF"
    $pickOutputButton.Fill.ForeColor.RGB = Color-Ref "2F75B5"
    $pickOutputButton.Line.ForeColor.RGB = Color-Ref "205493"
    $pickOutputButton.OnAction = "ChonNoiXuat"

    $inputSheet.PageSetup.Orientation = 2
    $inputSheet.PageSetup.Zoom = $false
    $inputSheet.PageSetup.FitToPagesWide = 1
    $inputSheet.PageSetup.FitToPagesTall = $false
    $inputSheet.PageSetup.PrintArea = "`$A`$1:`$AF`$$profileLayoutCapRow"
    $inputSheet.Tab.Color = Color-Ref "217346"

    $configSheet.Range("A1:B1").Merge()
    $configSheet.Range("A1").Value2 = "CẤU HÌNH HỒ SƠ"
    $configSheet.Range("A1:B1").Font.Name = "Arial"
    $configSheet.Range("A1:B1").Font.Size = 16
    $configSheet.Range("A1:B1").Font.Bold = $true
    $configSheet.Range("A1:B1").Font.Color = Color-Ref "FFFFFF"
    $configSheet.Range("A1:B1").Interior.Color = Color-Ref "304F78"
    $configSheet.Range("A3:A12").Font.Name = "Arial"
    $configSheet.Range("A3:A12").Font.Bold = $true
    $configSheet.Range("B3").Formula = "=HangTKToiDa"
    $configSheet.Range("B4").Value2 = "PALETTE_B_TUONG_PHAN"
    $configSheet.Range("B5").Value2 = "2.2.0"
    $configSheet.Range("B6").Value2 = 1
    $configSheet.Range("B7").Value2 = 100
    $configSheet.Range("B8").NumberFormat = "@"
    $configSheet.Range("B8").Value = [string]$testTemplate
    $configSheet.Range("B9").NumberFormat = "@"
    $configSheet.Range("B9").Value = [string]$wordOutputDir
    $configSheet.Range("B10").Value2 = 1
    $configSheet.Range("B11").NumberFormat = "@"
    $configSheet.Range("B11").Value = [string]$outputDir
    $configSheet.Range("B12").Value2 = 3
    $configSheet.Range("B3:B12").Font.Name = "Arial"
    $configSheet.Range("B3:B12").Interior.Color = Color-Ref "F4F6F9"
    $configSheet.Columns.Item("A").ColumnWidth = 24
    $configSheet.Columns.Item("B").ColumnWidth = 75
    $configSheet.Range("A3:B12").Borders.Color = Color-Ref "D8DEE8"

    $workbook.Names.Add("HangTKToiDa", '=NhapLieu!$B$4') | Out-Null
    $workbook.Names.Add("BangMauHangTK", '=CauHinh!$B$4') | Out-Null
    $workbook.Names.Add("PhienBanCauTruc", '=CauHinh!$B$5') | Out-Null
    $workbook.Names.Add("NguoiIDTiepTheo", '=CauHinh!$B$6') | Out-Null
    $workbook.Names.Add("DuongDanMauWord", '=CauHinh!$B$8') | Out-Null
    $workbook.Names.Add("MauWordDangChon", '=CauHinh!$B$8') | Out-Null
    $workbook.Names.Add("ThuMucXuat", '=CauHinh!$B$9') | Out-Null
    $workbook.Names.Add("TaiSanIDTiepTheo", '=CauHinh!$B$10') | Out-Null
    $workbook.Names.Add("ThuMucMauWord", '=CauHinh!$B$11') | Out-Null
    $workbook.Names.Add("SucChuaNguoiMau", '=CauHinh!$B$7') | Out-Null
    $workbook.Names.Add("SucChuaTaiSanMau", '=CauHinh!$B$12') | Out-Null

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
    Add-StandardModule $vbProject "modGiayTo" (Join-Path $vbaRoot "modGiayTo.bas")
    Add-StandardModule $vbProject "modHangTKNhanh" (Join-Path $vbaRoot "modHangTKNhanh.bas")
    Add-StandardModule $vbProject "modValidation" (Join-Path $vbaRoot "modValidation.bas")
    Add-StandardModule $vbProject "modExportData" (Join-Path $vbaRoot "modExportData.bas")
    Add-StandardModule $vbProject "modXuatWord" (Join-Path $vbaRoot "modXuatWord.bas")
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
    $testConfig = $testBook.Worksheets.Item("CauHinh")
    $testExport = $testBook.Worksheets.Item("XuatAn")
    $testTable = $testInput.ListObjects.Item("tblNguoi")
    $expectedPeopleTechHeaders = @(
        "STTNhap", "NguoiID", "ParentNguoiID", "LaChuDat", "NhomTuChoiID",
        "TrangThai", "NgaySinhGoc", "NgaySinhTinh", "NgayChetGoc", "NgayChetTinh",
        "NgayCapGoc", "NgayCapTinh", "LoaiCC", "NoiCapCC", "NhanDiaChi"
    )
    $schemaFailures = [System.Collections.Generic.List[string]]::new()
    $peopleTableAddress = [string]$testTable.Range.Address($false, $false)
    if ($peopleTableAddress -ne "A8:I38") {
        $schemaFailures.Add("tblNguoi đang ở $peopleTableAddress, cần A8:I38.")
    }
    if ([double]$excel.WorksheetFunction.CountA($testInput.Range("J8:Z38")) -ne 0) {
        $schemaFailures.Add("NhapLieu!J8:Z38 vẫn còn tiêu đề hoặc dữ liệu kỹ thuật Người.")
    }

    $testPeopleTech = $null
    try { $testPeopleTech = $testExport.ListObjects.Item("tblNguoiKyThuat") } catch { $testPeopleTech = $null }
    if ($null -eq $testPeopleTech) {
        $schemaFailures.Add("XuatAn chưa có tblNguoiKyThuat.")
    }
    else {
        $actualPeopleTechHeaders = @($testPeopleTech.ListColumns | ForEach-Object { [string]$_.Name })
        if (($actualPeopleTechHeaders -join "|") -ne ($expectedPeopleTechHeaders -join "|")) {
            $schemaFailures.Add("tblNguoiKyThuat sai tên hoặc thứ tự cột: $($actualPeopleTechHeaders -join ', ').")
        }
        if ($null -eq $testPeopleTech.DataBodyRange -or $testPeopleTech.DataBodyRange.Rows.Count -ne 30) {
            $actualRowCount = if ($null -eq $testPeopleTech.DataBodyRange) { 0 } else { $testPeopleTech.DataBodyRange.Rows.Count }
            $schemaFailures.Add("tblNguoiKyThuat có $actualRowCount dòng dữ liệu, cần đúng 30.")
        }
        else {
            $sttColumn = $testPeopleTech.ListColumns.Item("STTNhap").Index
            for ($techRow = 1; $techRow -le 30; $techRow++) {
                if ([int]$testPeopleTech.DataBodyRange.Cells.Item($techRow, $sttColumn).Value2 -ne $techRow) {
                    $schemaFailures.Add("STTNhap kỹ thuật ở dòng $techRow không bằng $techRow.")
                    break
                }
            }
        }
    }
    if ([string]$testInput.Range("AA8").Value2 -ne "TÀI SẢN 1" -or
        [string]$testInput.Range("AC8").Value2 -ne "TÀI SẢN 2" -or
        [string]$testInput.Range("AE8").Value2 -ne "TÀI SẢN 3" -or
        [string]$testInput.Range("AA9").Value2 -ne "Loại sổ" -or
        [string]$testInput.Range("AB9").Value2 -ne "") {
        $schemaFailures.Add("Ba thẻ Tài sản phải ở AA:AF, cùng bắt đầu từ hàng 8/9.")
    }
    if ([bool]$testInput.Range("AA8:AF26").MergeCells -or
        [bool]$testInput.Range("AB9").Locked -or
        -not [bool]$testInput.Range("AA9").Locked -or
        [string]$testInput.Range("AB9").NumberFormat -ne "@" -or
        -not [bool]$testInput.Range("AB14").WrapText) {
        $schemaFailures.Add("Thẻ Tài sản phải không gộp ô, chỉ mở khóa ô giá trị, giữ Text và bọc dòng cho địa chỉ.")
    }
    foreach ($assetColumn in @("AA", "AC", "AE")) {
        if ([string]$testInput.Range("${assetColumn}6").Value2 -ne "THÔNG TIN TÀI SẢN" -or
            [string]$testInput.Range("${assetColumn}7").Value2 -ne "Thẻ không dùng thì để trống") {
            $schemaFailures.Add("Thiếu nhãn hướng dẫn của thẻ Tài sản tại ${assetColumn}6:${assetColumn}7.")
            break
        }
    }
    if ([string]$testInput.Range("B40").Value2 -ne "THÔNG TIN HỒ SƠ" -or
        [string]$testInput.Range("B41").Value2 -ne "Niêm yết" -or
        [string]$testInput.Range("B42").Value2 -ne "Số công chứng" -or
        [string]$testInput.Range("B43").Value2 -ne "Người ủy quyền" -or
        [string]$testInput.Range("B44").Value2 -ne "Người ủy quyền 2" -or
        [bool]$testInput.Range("B45").Locked -or [bool]$testInput.Range("C45").Locked -or
        [bool]$testInput.Range("B70").Locked -or [bool]$testInput.Range("C70").Locked -or
        [string]$testInput.Range("C41:C70").NumberFormat -ne "@") {
        $schemaFailures.Add("Khối Hồ sơ phải bắt đầu từ B40, mở rộng đến B40:C70 khi cần và có 30 hàng nhập dự phòng dạng Text.")
    }
    if ([string]$testConfig.Range("B5").Value2 -ne "2.2.0") {
        $schemaFailures.Add("PhienBanCauTruc phải là 2.2.0.")
    }
    if ([string]$testInput.Range("A5").Value2 -ne "Mẫu Word" -or
        [string]$testInput.Range("C5").Value2 -ne [System.IO.Path]::GetFileName($testTemplate) -or
        -not [bool]$testInput.Range("C5").Locked -or
        [string]$testInput.Range("B3").Value2 -ne [string]$outputDir -or
        [string]$testInput.Range("B7").Value2 -ne [string]$wordOutputDir) {
        $schemaFailures.Add("Thanh công cụ Word phải hiển thị đúng mẫu, nơi lấy mẫu và nơi xuất; C5 chỉ đọc.")
    }
    if ([string]$testInput.Range("B9:B38").NumberFormat -ne "@" -or
        [string]$testInput.Range("G9:G38").NumberFormat -ne "@") {
        $schemaFailures.Add("Họ tên và địa chỉ Người phải có định dạng Text.")
    }
    if (-not [bool]$testInput.Range("L:L").EntireColumn.Hidden -or
        [bool]$testInput.Range("J8").Locked -or [bool]$testInput.Range("J9").Locked -or
        [string]$testInput.Range("J9:Z38").NumberFormat -ne "@") {
        $schemaFailures.Add("Vùng Người mở rộng J8:Z38 phải ẩn khi chưa có tiêu đề, mở khóa và giữ Text.")
    }
    foreach ($requiredName in @("MauWordDangChon", "DuongDanMauWord", "ThuMucMauWord", "ThuMucXuat", "SucChuaNguoiMau", "SucChuaTaiSanMau")) {
        try { $null = $testBook.Names.Item($requiredName) } catch { $schemaFailures.Add("Thiếu named range cấu hình $requiredName.") }
    }
    $expectedConfigNames = @{
        "MauWordDangChon" = "B8"
        "DuongDanMauWord" = "B8"
        "ThuMucMauWord" = "B11"
        "ThuMucXuat" = "B9"
        "SucChuaNguoiMau" = "B7"
        "SucChuaTaiSanMau" = "B12"
    }
    foreach ($configName in $expectedConfigNames.Keys) {
        try {
            $configRef = $testBook.Names.Item($configName).RefersToRange
            $configAddress = [string]$configRef.Address($false, $false, 1, $false)
            if ($configRef.Worksheet.Name -ne "CauHinh" -or $configAddress -ne [string]$expectedConfigNames[$configName]) {
                $schemaFailures.Add("Named range $configName không trỏ đúng CauHinh!$($expectedConfigNames[$configName]).")
            }
            Release-ComObject $configRef
        }
        catch {
            # Missing names are reported by the existence check above.
        }
    }
    if ($schemaFailures.Count -gt 0) {
        throw "Workbook schema assertions failed: $($schemaFailures -join ' ')"
    }
    $testInput.Unprotect("HoSoTK_MVP_2026")

    $fakePeople = @(
        @("Ông Nguyễn Văn A", "1950", "1999", "001050000001", "2021", "Thành phố Hà Nội", 0, 0),
        @("Bà Trần Thị B", "02/1952", "11/06/2021", "001052000002", "02/2021", "Thành phố Hà Nội", 0, 0),
        @("Anh Nguyễn Văn C", "03/03/1975", "07/2022", "001075000003", [datetime]"2022-03-03", "Tỉnh Ninh Bình", 1, 0),
        @("Chị Nguyễn Thị C1", "2000", $null, "001200000004", "04/2023", "Tỉnh Ninh Bình", 2, 1),
        @("Anh Nguyễn Văn C2", "05/2002", $null, "001202000005", "05/05/2023", "Tỉnh Ninh Bình", 2, 0),
        @("Chị Nguyễn Thị D", [datetime]"1978-06-06", $null, "001078000006", "2022", "Tỉnh Nam Định", 1, 1),
        @("Người kiểm thử năm", "2015", $null, "001015000007", "30/06/2024", "Tỉnh Hà Nam", 1, 0),
        @("Người kiểm thử tháng", "05/2015", $null, "001015000008", "01/07/2024", "Tỉnh Hà Nam", 1, 0),
        @("Người kiểm thử ngày", "12/05/2015", $null, "001015000009", "02/07/2024", "Tỉnh Hà Nam", 1, 0)
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
    if ([string]$excel.Run($macroPrefix + "NormalizeKey", "Số điện thoại") -ne "sodienthoai") {
        throw "NormalizeKey chưa bỏ dấu tiếng Việt và ký tự phân cách đúng."
    }
    $longPathForDisplay = "C:\một\thư\mục\rất\dài\để\kiểm\tra\đường\dẫn\mẫu.docx"
    $shortPathForDisplay = [string]$excel.Run($macroPrefix + "ShortPathText", $longPathForDisplay, 20)
    if ($shortPathForDisplay.Length -gt 20 -or $shortPathForDisplay.Substring(0, 3) -ne "C:\" -or
        $shortPathForDisplay.IndexOf("...") -lt 0) {
        throw "ShortPathText chưa rút gọn đường dẫn hiển thị đúng giới hạn."
    }
    $testInput.Range("J8").Value2 = "Số điện thoại"
    $testInput.Range("K8").Value2 = "Mã hồ sơ"
    $testInput.Range("J9").Value2 = "00123"
    $testInput.Range("J10").Formula = '=$B10'
    $testInput.Range("J11").Value2 = "NHẬP TAY"
    $testInput.Range("K9").Value2 = "ABC"
    [void]$excel.Run($macroPrefix + "HandlePersonExtensionChange", $testInput.Range("J8:K11"))
    if ([bool]$testInput.Range("J:J").EntireColumn.Hidden -or [bool]$testInput.Range("K:K").EntireColumn.Hidden -or
        -not [bool]$testInput.Range("L:L").EntireColumn.Hidden) {
        throw "Cột Người mở rộng không hiện/ẩn đúng theo tiêu đề."
    }

    # Exercise the actual worksheet event with one mixed multi-area change,
    # not only the public handlers called directly below.
    $excel.EnableEvents = $true
    $mixedChange = $excel.Union($testInput.Range("J13"), $testInput.Range("AB9"), $testInput.Range("B38"))
    $mixedChange.Value2 = "MIXED_EVENT"
    $excel.EnableEvents = $false
    if ([string]$testInput.Range("J13").Value2 -ne "MIXED_EVENT" -or
        [string]$testExport.Range("BM41").Value2 -eq "" -or
        -not [bool]$testExport.Range("BP41").Value2) {
        throw "Worksheet_Change chưa xử lý đủ thay đổi hỗn hợp Người/Tài sản."
    }
    Release-ComObject $mixedChange
    $testInput.Range("J13,AB9,B38").ClearContents()
    [void]$excel.Run($macroPrefix + "HandlePersonExtensionChange", $testInput.Range("J13"))
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("AB9"))
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("B38"))
    $excel.EnableEvents = $true
    $testInput.Range("J13").Value2 = "=1+1"
    $excel.EnableEvents = $false
    if (-not [bool]$testInput.Range("J13").HasFormula -or [double]$testInput.Range("J13").Value2 -ne 2) {
        throw "Nhập công thức vào ô Người mở rộng dạng Text chưa được chuyển thành công thức."
    }
    $testInput.Range("J13").ClearContents()
    [void]$excel.Run($macroPrefix + "HandlePersonExtensionChange", $testInput.Range("J13"))

    # Z is the hard boundary; no extension column beyond it may be touched.
    $testInput.Range("Z8").Value2 = "Trường cuối"
    [void]$excel.Run($macroPrefix + "HandlePersonExtensionChange", $testInput.Range("Z8"))
    if ([bool]$testInput.Range("Y:Y").EntireColumn.Hidden -eq $false -or
        [bool]$testInput.Range("Z:Z").EntireColumn.Hidden) {
        throw "Biên vùng Người mở rộng chưa dừng đúng tại cột Z."
    }
    $testInput.Range("Z8").ClearContents()
    [void]$excel.Run($macroPrefix + "HandlePersonExtensionChange", $testInput.Range("Z8"))
    $testInput.Range("J14").Formula = "=1/0"
    [void]$excel.Run($macroPrefix + "HandlePersonExtensionChange", $testInput.Range("J14"))
    if ([string]$excel.Run($macroPrefix + "ValueToExportText", $testInput.Range("J14").Value2) -ne "#DIV/0!") {
        throw "Lỗi công thức Excel chưa được giữ nguyên dạng #DIV/0!."
    }
    $testInput.Range("J14").ClearContents()
    $testInput.Range("L8").Value2 = "Số-điện thoại"
    [void]$excel.Run($macroPrefix + "ValidateWorkbook", $false)
    $duplicateLabelFound = $false
    for ($checkRow = 5; $checkRow -le 1000; $checkRow++) {
        if ([string]$testCheck.Cells.Item($checkRow, 2).Value2 -eq "LABEL_DUPLICATE") {
            $duplicateLabelFound = $true
            break
        }
    }
    if (-not $duplicateLabelFound) { throw "Chẩn đoán chưa ghi nhận nhãn Người mở rộng trùng sau chuẩn hóa." }
    $testInput.Range("L8").ClearContents()
    [void]$excel.Run($macroPrefix + "HandlePersonExtensionChange", $testInput.Range("L8"))
    [void]$excel.Run($macroPrefix + "DongBoTruong")
    if (-not [bool]$testInput.Range("J12").HasFormula -or
        [string]$testInput.Range("J11").Value2 -ne "NHẬP TAY" -or
        [string]$testInput.Range("J9").Value2 -ne "00123" -or
        [string]$testInput.Range("J10").NumberFormat -ne "General") {
        throw "Đồng bộ trường Người phải nhân công thức nhưng giữ nguyên dữ liệu nhập tay."
    }
    [void]$excel.Run($macroPrefix + "ApplyWorkbookProtection")
    # EnableEvents is intentionally false while seeding QA data, so explicitly
    # exercise the same per-change handler that a user edit in NgayCap would call.
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("F9:F17"))
    [void]$excel.Run($macroPrefix + "RefreshAllPeople")

    if (-not [bool]$testExport.ProtectContents) { throw "XuatAn chưa được bảo vệ." }

    $techIdColumn = $testPeopleTech.ListColumns.Item("NguoiID").Index
    $techParentColumn = $testPeopleTech.ListColumns.Item("ParentNguoiID").Index
    $techOwnerColumn = $testPeopleTech.ListColumns.Item("LaChuDat").Index
    $techStatusColumn = $testPeopleTech.ListColumns.Item("TrangThai").Index
    $techBirthRawColumn = $testPeopleTech.ListColumns.Item("NgaySinhGoc").Index
    $techBirthCalcColumn = $testPeopleTech.ListColumns.Item("NgaySinhTinh").Index
    $techLoaiCcColumn = $testPeopleTech.ListColumns.Item("LoaiCC").Index
    $techNoiCapCcColumn = $testPeopleTech.ListColumns.Item("NoiCapCC").Index
    $techNhanDiaChiColumn = $testPeopleTech.ListColumns.Item("NhanDiaChi").Index

    if (-not [bool]$testPeopleTech.DataBodyRange.Cells.Item(1, $techOwnerColumn).Value2 -or
        -not [bool]$testPeopleTech.DataBodyRange.Cells.Item(2, $techOwnerColumn).Value2) {
        throw "Cờ chủ đất của hai dòng đầu không còn đúng sau khi chuyển bảng kỹ thuật."
    }
    $row3Id = [string]$testPeopleTech.DataBodyRange.Cells.Item(3, $techIdColumn).Value2
    $row4ParentId = [string]$testPeopleTech.DataBodyRange.Cells.Item(4, $techParentColumn).Value2
    if ([string]::IsNullOrWhiteSpace($row3Id) -or $row4ParentId -ne $row3Id) {
        throw "ParentNguoiID của dòng H2 không liên kết tới đúng dòng H1."
    }
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(1, $techStatusColumn).Value2 -ne "Đã chết" -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(4, $techStatusColumn).Value2 -ne "Nhận đất" -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(5, $techStatusColumn).Value2 -ne "Từ chối") {
        throw "Trạng thái Người thay đổi sau khi chuyển bảng kỹ thuật."
    }

    $personDateExpectations = @(
        @{ Row = 7; Text = "2015"; Date = [datetime]"2015-01-01" },
        @{ Row = 8; Text = "05/2015"; Date = [datetime]"2015-05-01" },
        @{ Row = 9; Text = "12/05/2015"; Date = [datetime]"2015-05-12" }
    )
    foreach ($dateExpectation in $personDateExpectations) {
        $personRow = [int]$dateExpectation.Row
        $sheetRow = $personRow + 8
        if ([string]$testInput.Cells.Item($sheetRow, 3).Value2 -ne [string]$dateExpectation.Text -or
            [string]$testPeopleTech.DataBodyRange.Cells.Item($personRow, $techBirthRawColumn).Value2 -ne [string]$dateExpectation.Text -or
            [math]::Abs(([double]$testPeopleTech.DataBodyRange.Cells.Item($personRow, $techBirthCalcColumn).Value2) - $dateExpectation.Date.ToOADate()) -gt 0.0001) {
            throw "Ngày Người '$($dateExpectation.Text)' không được giữ nguyên và chuẩn hóa đúng trong tblNguoiKyThuat."
        }
    }

    $loaiCcBefore = [string]$testBook.Names.Item("DanhMuc_LoaiCCTruocMoc").RefersToRange.Value2
    $loaiCcFrom = [string]$testBook.Names.Item("DanhMuc_LoaiCCTuMoc").RefersToRange.Value2
    $loaiGiayToRange = $testBook.Names.Item("DanhMuc_LoaiGiayTo").RefersToRange
    $loaiCcBeforeIndex = 0
    $loaiCcFromIndex = 0
    for ($catalogIndex = 1; $catalogIndex -le $loaiGiayToRange.Rows.Count; $catalogIndex++) {
        if ([string]$loaiGiayToRange.Cells.Item($catalogIndex, 1).Value2 -eq $loaiCcBefore) { $loaiCcBeforeIndex = $catalogIndex }
        if ([string]$loaiGiayToRange.Cells.Item($catalogIndex, 1).Value2 -eq $loaiCcFrom) { $loaiCcFromIndex = $catalogIndex }
    }
    if ($loaiCcBeforeIndex -eq 0 -or $loaiCcFromIndex -eq 0) { throw "Danh mục LoaiCC thiếu giá trị tại mốc ngày." }
    $expectedNoiCapBefore = [string]$testBook.Names.Item("DanhMuc_NoiCapCC").RefersToRange.Cells.Item($loaiCcBeforeIndex, 1).Value2
    $expectedNhanDiaChiBefore = [string]$testBook.Names.Item("DanhMuc_NhanDiaChi").RefersToRange.Cells.Item($loaiCcBeforeIndex, 1).Value2
    $expectedNoiCapFrom = [string]$testBook.Names.Item("DanhMuc_NoiCapCC").RefersToRange.Cells.Item($loaiCcFromIndex, 1).Value2
    $expectedNhanDiaChiFrom = [string]$testBook.Names.Item("DanhMuc_NhanDiaChi").RefersToRange.Cells.Item($loaiCcFromIndex, 1).Value2
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(1, $techLoaiCcColumn).Value2 -ne $loaiCcBefore -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techLoaiCcColumn).Value2 -ne $loaiCcBefore -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techLoaiCcColumn).Value2 -ne $loaiCcFrom) {
        throw "Phân loại giấy tờ trước/từ 01/07/2024 không còn đúng."
    }
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techNoiCapCcColumn).Value2 -ne $expectedNoiCapBefore -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techNhanDiaChiColumn).Value2 -ne $expectedNhanDiaChiBefore -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techNoiCapCcColumn).Value2 -ne $expectedNoiCapFrom -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techNhanDiaChiColumn).Value2 -ne $expectedNhanDiaChiFrom) {
        throw "Nơi cấp hoặc nhãn địa chỉ không được tra đúng theo LoaiCC."
    }
    $row8LoaiCcBeforeClear = [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techLoaiCcColumn).Value2
    $testInput.Range("F15").ClearContents()
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("F15"))
    if (-not [string]::IsNullOrWhiteSpace([string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techLoaiCcColumn).Value2) -or
        -not [string]::IsNullOrWhiteSpace([string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techNoiCapCcColumn).Value2) -or
        -not [string]::IsNullOrWhiteSpace([string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techNhanDiaChiColumn).Value2)) {
        throw "NgayCapTinh trống chưa xóa LoaiCC, NoiCapCC và NhanDiaChi."
    }
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techLoaiCcColumn).Value2 -ne $row8LoaiCcBeforeClear) {
        throw "Sửa NgayCap của một Người đã thay đổi LoaiCC của dòng khác."
    }
    $testInput.Range("F15").Value2 = [double]([datetime]"2024-06-30").ToOADate()
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("F15"))

    if ([string]$testInput.Range("C9").Value2 -ne "1950") {
        throw "Năm sinh 1950 không được giữ nguyên trên giao diện."
    }
    if ([string]$testInput.Range("D9").Value2 -ne "1999") {
        throw "Năm chết 1999 không được giữ nguyên trên giao diện."
    }
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(1, $techBirthRawColumn).Value2 -ne "1950") { throw "Dữ liệu gốc năm sinh không đúng." }
    $techDeathRawColumn = $testPeopleTech.ListColumns.Item("NgayChetGoc").Index
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(1, $techDeathRawColumn).Value2 -ne "1999") { throw "Dữ liệu gốc năm chết không đúng." }
    if ([string]$testInput.Range("C10").Value2 -ne "02/1952") { throw "Tháng/năm không được giữ nguyên trên giao diện." }
    $ageAtDeath = [int]$excel.Run($macroPrefix + "AgeAtDeath", 1)
    if ($ageAtDeath -ne 49) { throw "Tính tuổi từ ngày chuẩn hóa không đúng." }

    $row8IdBefore = [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techIdColumn).Value2
    $row8BirthBefore = [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techBirthRawColumn).Value2
    $testInput.Range("C15").Value2 = "2016"
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("C15"))
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techBirthRawColumn).Value2 -ne "2016" -or
        [math]::Abs(([double]$testPeopleTech.DataBodyRange.Cells.Item(7, $techBirthCalcColumn).Value2) - ([datetime]"2016-01-01").ToOADate()) -gt 0.0001) {
        throw "Sửa dòng Người 7 không cập nhật đúng dòng kỹ thuật 7."
    }
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techIdColumn).Value2 -ne $row8IdBefore -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techBirthRawColumn).Value2 -ne $row8BirthBefore) {
        throw "Sửa dòng Người 7 đã làm thay đổi nhầm dòng kỹ thuật 8."
    }
    $testInput.Range("C15").Value2 = "2015"
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("C15"))

    $row8IdBeforeDelete = [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techIdColumn).Value2
    $testInput.Range("B15:G15").ClearContents()
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("B15:G15"))
    if (-not [string]::IsNullOrWhiteSpace([string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techIdColumn).Value2) -or
        -not [string]::IsNullOrWhiteSpace([string]$testPeopleTech.DataBodyRange.Cells.Item(7, $techBirthRawColumn).Value2)) {
        throw "Xóa Người ở dòng 7 chưa xóa đúng dữ liệu kỹ thuật dòng 7."
    }
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techIdColumn).Value2 -ne $row8IdBeforeDelete) {
        throw "Xóa Người ở dòng 7 đã làm lệch dữ liệu kỹ thuật dòng 8."
    }
    $testExport.Unprotect("HoSoTK_MVP_2026")
    $testPeopleTech.DataBodyRange.Cells.Item(7, $techIdColumn).Value2 = "GIU_LAI_KIEM_THU"
    # Buộc STT không trùng bất kỳ dòng kỹ thuật nào và giữ dòng kỹ thuật 7 bận.
    # Nếu VBA dùng vị trí dòng, dữ liệu sẽ ghi sai vào dòng 7 thay vì cấp dòng trống khác.
    $testInput.Unprotect("HoSoTK_MVP_2026")
    $inputSttColumn = $testTable.ListColumns.Item("STTNhap").Index
    $techSttColumn = $testPeopleTech.ListColumns.Item("STTNhap").Index
    $testTable.DataBodyRange.Cells.Item(7, $inputSttColumn).Value2 = 31
    $testInput.Range("B15").Value2 = "Người kiểm thử năm"
    $testInput.Range("C15").Value2 = "2015"
    $testInput.Range("E15").NumberFormat = "@"
    $testInput.Range("E15").Value2 = "001015000007"
    $testInput.Range("F15").Value2 = "30/06/2024"
    $testInput.Range("G15").Value2 = "Tỉnh Hà Nam"
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("B15:G15"))
    $readdedInputStt = [string]$testTable.DataBodyRange.Cells.Item(7, $inputSttColumn).Value2
    $readdedTechRow = 0
    $readdedTechMatchCount = 0
    for ($techRow = 1; $techRow -le $testPeopleTech.DataBodyRange.Rows.Count; $techRow++) {
        if ([string]$testPeopleTech.DataBodyRange.Cells.Item($techRow, $techSttColumn).Value2 -eq $readdedInputStt) {
            $readdedTechMatchCount++
            $readdedTechRow = $techRow
        }
    }
    if ($readdedTechMatchCount -ne 1 -or $readdedTechRow -eq 0 -or
        [string]::IsNullOrWhiteSpace([string]$testPeopleTech.DataBodyRange.Cells.Item($readdedTechRow, $techIdColumn).Value2) -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item($readdedTechRow, $techBirthRawColumn).Value2 -ne "2015") {
        throw "Thêm lại Người ở dòng 7 chưa tạo/cập nhật đúng dữ liệu kỹ thuật theo STTNhap."
    }
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item($readdedTechRow, $techSttColumn).Value2 -ne $readdedInputStt) {
        throw "STTNhap kỹ thuật chưa liên kết với STTNhap của Người sau khi xóa/thêm lại."
    }
    if ($readdedInputStt -ne "31") {
        throw "Người thêm lại không giữ STTNhap 31 trong ca kiểm tra liên kết."
    }
    if ($readdedTechRow -eq 7) {
        throw "Dữ liệu kỹ thuật Người vẫn bị liên kết theo vị trí dòng thay vì STTNhap."
    }
    $testInput.Unprotect("HoSoTK_MVP_2026")
    $testInput.Range("B15:G15").ClearContents()
    [void]$excel.Run($macroPrefix + "HandlePeopleChange", $testInput.Range("B15:G15"))
    $testPeopleTech.DataBodyRange.Cells.Item(7, $techIdColumn).ClearContents()
    [void]$excel.Run($macroPrefix + "ApplyWorkbookProtection")
    if ([string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techIdColumn).Value2 -ne $row8IdBeforeDelete) {
        throw "Thêm lại Người ở dòng 7 đã làm lệch dữ liệu kỹ thuật dòng 8."
    }

    $testInput.Range("AD10").Value2 = "SERIAL-002"
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("AD10"))
    if ([string]$testExport.Range("BM61").Value2 -ne "TS001") { throw "Phiếu tài sản 2 chưa sinh TS001." }
    if (-not [bool]$testExport.Range("BP61").Value2) { throw "Phiếu tài sản 2 chưa được đánh dấu có dữ liệu." }
    if ([string]$testExport.Range("BM41").Value2 -ne "" -or [string]$testExport.Range("BM81").Value2 -ne "") {
        throw "Phiếu tài sản trống lại có TaiSanID."
    }

    $testInput.Range("AB9").Value2 = "GCN"
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("AB9"))
    $assetDateExpectations = @{
        "2015" = [datetime]"2015-01-01"
        "05/2015" = [datetime]"2015-05-01"
        "12/05/2015" = [datetime]"2015-05-12"
    }
    foreach ($dateText in $assetDateExpectations.Keys) {
        $testInput.Range("AB24").Value2 = [string]$dateText
        [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("AB24"))
        if ([string]$testInput.Range("AB24").Value2 -ne [string]$dateText -or
            [string]$testExport.Range("BN41").Value2 -ne [string]$dateText -or
            [math]::Abs(([double]$testExport.Range("BO41").Value2) - $assetDateExpectations[$dateText].ToOADate()) -gt 0.0001) {
            throw "Ngày cấp sổ '$dateText' không được giữ nguyên và chuẩn hóa đúng."
        }
    }

    $testInput.Range("AD10").ClearContents()
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("AD10"))
    if ([string]$testExport.Range("BM61").Value2 -ne "" -or [string]$testExport.Range("BN61").Value2 -ne "" -or
        [string]$testExport.Range("BO61").Value2 -ne "" -or [bool]$testExport.Range("BP61").Value2) {
        throw "Xóa trắng phiếu tài sản chưa xóa dữ liệu hệ thống."
    }
    if (-not [bool]$testInput.Range("L:Z").EntireColumn.Hidden -or
        [bool]$testInput.Range("J:J").EntireColumn.Hidden -or
        [bool]$testInput.Range("K:K").EntireColumn.Hidden) {
        throw "Vùng Người mở rộng không giữ đúng trạng thái hiện/ẩn sau khi xử lý tài sản."
    }
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
    $invalidExportSucceeded = [bool]$excel.Run($macroPrefix + "RunWordExport", $true)
    if (-not $invalidExportSucceeded) { throw "Dữ liệu nội dung không hợp lệ đã chặn nhầm xuất Word." }
    $testInput.Range("C14").Value2 = "1978"
    [void]$excel.Run($macroPrefix + "RefreshAllPeople")
    $testInput.Range("AB9").ClearContents()
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("AB9"))

    # Restore one complete, valid asset after the empty-card diagnostic case so
    # the final validation test measures the intended export path.
    $testInput.Range("AB9").Value2 = "GCN"
    $testInput.Range("AB12").Value2 = "12"
    $testInput.Range("AB14").Value2 = "Số 1, Hà Nội"
    $testInput.Range("AB15").Value2 = 100
    $testInput.Range("AB17").Value2 = "ONT"
    [void]$excel.Run($macroPrefix + "HandleTaiSanChange", $testInput.Range("AB9:AB17"))

    # Cấu hình tối thiểu cho khối thông tin phụ khi mẫu thử có placeholder bắt buộc.
    $testInput.Range("C41").Value2 = "Có"
    $testInput.Range("C43").Value2 = [string]$catalogData.NguoiUyQuyen[0]
    $testInput.Range("B45").Value2 = "Trường mở rộng"
    $testInput.Range("C45").Value2 = "Giá trị mở rộng"
    $testConfig.Unprotect("HoSoTK_MVP_2026")
    $testCatalog = $testBook.Worksheets.Item("DanhMuc")
    $testCatalog.Unprotect("HoSoTK_MVP_2026")
    $testCapacityRow = $testCatalog.Cells($testCatalog.Rows.Count, 10).End(-4162).Row + 1
    $testCatalog.Cells.Item($testCapacityRow, 10).Value2 = [System.IO.Path]::GetFileName($testTemplate)
    $testCatalog.Cells.Item($testCapacityRow, 11).Value2 = 100
    $testCatalog.Cells.Item($testCapacityRow, 12).Value2 = 3

    $isValid = [bool]$excel.Run($macroPrefix + "ValidateWorkbook", $false)
    if (-not $isValid) { throw "Bộ dữ liệu giả không vượt qua validation." }
    [void]$excel.Run($macroPrefix + "ApplyWorkbookProtection")
    $techRow8IdBeforeExport = [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techIdColumn).Value2
    [void]$excel.Run($macroPrefix + "BuildExportData")
    if ([string]$testPeopleTech.ListColumns.Item(1).Name -ne "STTNhap" -or
        [string]$testPeopleTech.DataBodyRange.Cells.Item(8, $techIdColumn).Value2 -ne $techRow8IdBeforeExport) {
        throw "BuildExportData đã làm mất hoặc thay đổi tblNguoiKyThuat."
    }
    $hoSoStartRow = 0
    for ($exportRow = 1; $exportRow -le $testExport.UsedRange.Rows.Count; $exportRow++) {
        if ([string]$testExport.Cells.Item($exportRow, 1).Value2 -eq "HoSo") { $hoSoStartRow = $exportRow; break }
    }
    if ($hoSoStartRow -eq 0) { throw "BuildExportData chưa tạo phần HoSo." }
    $extensionColumn = 0
    for ($exportColumn = 1; $exportColumn -le 20; $exportColumn++) {
        if ([string]$testExport.Cells.Item($hoSoStartRow + 1, $exportColumn).Value2 -eq "Trường mở rộng") {
            $extensionColumn = $exportColumn
            break
        }
    }
    if ($extensionColumn -eq 0 -or [string]$testExport.Cells.Item($hoSoStartRow + 2, $extensionColumn).Value2 -ne "Giá trị mở rộng") {
        throw "BuildExportData chưa quét và xuất trường Hồ sơ mở rộng theo nhãn."
    }
    $soCongChungColumn = 0
    for ($exportColumn = 1; $exportColumn -le 20; $exportColumn++) {
        if ([string]$testExport.Cells.Item($hoSoStartRow + 1, $exportColumn).Value2 -eq "SoCongChung") {
            $soCongChungColumn = $exportColumn
            break
        }
    }
    if ($soCongChungColumn -eq 0 -or -not [string]::IsNullOrEmpty([string]$testExport.Cells.Item($hoSoStartRow + 2, $soCongChungColumn).Value2)) {
        throw "SoCongChung trống phải xuất chuỗi rỗng, không được xuất 0."
    }
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
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "TARGET=$targetWorkbook"
Write-Output "TEST=$testWorkbook"
Write-Output "QA_CLEAN=$qaCleanWorkbook"
Write-Output "QA_TEST=$qaTestWorkbook"
Write-Output "WORD_TEMPLATE=$testTemplate"
Write-Output "WORD_OUTPUT_DIR=$wordOutputDir"
