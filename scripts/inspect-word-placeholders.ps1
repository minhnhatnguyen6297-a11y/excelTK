[CmdletBinding()]
param(
    [string]$TemplatesPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "templates\\word"),
    [string]$ReportPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "docs\\reference\\word-placeholders-2026-08-31.md")
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$templatesRoot = (Resolve-Path -LiteralPath $TemplatesPath).Path
$reportDirectory = Split-Path -Parent $ReportPath
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
}

function Release-ComObject {
    param([object]$Object)

    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-TemplateCategory {
    param([string]$FileName)

    $name = $FileName.ToUpperInvariant()
    if ($name -match "PCDS|DI SẢN|KHAI NHẬN") { return "Thừa kế/PCDS" }
    if ($name -match "DK18|ĐK18") { return "Chuyển nhượng A/B" }
    if ($name -match "UQ|ỦY QUYỀN") { return "Ủy quyền" }
    return "Chưa phân loại"
}

function Get-WordDocumentText {
    param(
        [object]$Word,
        [string]$SourcePath
    )

    $document = $null
    try {
        $document = $Word.Documents.Open($SourcePath, $false, $true, $false)

        $segments = [System.Collections.Generic.List[string]]::new()
        foreach ($initialStory in $document.StoryRanges) {
            $story = $initialStory
            while ($null -ne $story) {
                $segments.Add([string]$story.Text)
                $nextStory = $story.NextStoryRange
                Release-ComObject $story
                $story = $nextStory
            }
        }
        return ($segments -join [Environment]::NewLine)
    }
    finally {
        if ($null -ne $document) {
            try { $document.Close($false) } catch {}
            Release-ComObject $document
        }
    }
}

function Get-Placeholders {
    param([string]$Text)

    $matches = [regex]::Matches($Text, '\[[^\]\r\n]+\]|\{\{[^}\r\n]+\}\}')
    return @($matches | ForEach-Object {
        $value = $_.Value.Trim()
        # Numbered form section labels such as [01] and [06.1] are printed
        # document text, not placeholders. Do not report them as legacy tokens.
        if ($value -notmatch '^\[[0-9]+(?:\.[0-9]+)*\]$') { $value }
    })
}

function Get-NumberedPlaceholderSlots {
    param(
        [string[]]$Placeholders,
        [string[]]$Roots,
        [int]$MaxSlot = 0
    )

    $slots = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($placeholder in $Placeholders) {
        $match = [regex]::Match($placeholder, '^\{\{([a-z][a-z0-9]*?)(\d*)\}\}$')
        if (-not $match.Success -or $Roots -notcontains $match.Groups[1].Value) {
            continue
        }

        $slotText = $match.Groups[2].Value
        $slot = if ([string]::IsNullOrWhiteSpace($slotText)) { 1 } else { [int]$slotText }
        if ($MaxSlot -gt 0 -and $slot -gt $MaxSlot) {
            continue
        }
        [void]$slots.Add($slot)
    }
    return @($slots | Sort-Object)
}

function Get-AssetSlots {
    param([string[]]$Placeholders)

    $assetRoots = @(
        "loaiso", "serial", "sovaoso", "sothua", "soto", "diachidat",
        "dientich", "hinhthucsudung", "loaidat", "thoihan", "ont", "cln",
        "nts", "luc", "nguongoc", "ngaycapso", "coquancapso", "ghichu"
    )
    return @(Get-NumberedPlaceholderSlots -Placeholders $Placeholders -Roots $assetRoots -MaxSlot 3)
}

function ConvertTo-MarkdownCell {
    param([string]$Value)

    return $Value.Replace("|", "\\|").Replace("`r", " ").Replace("`n", " ")
}

$templates = @(Get-ChildItem -LiteralPath $templatesRoot -File | Where-Object { $_.Extension -in ".doc", ".docx" } | Sort-Object Name)
if ($templates.Count -eq 0) {
    throw "Không tìm thấy mẫu .docx hoặc .doc trong $templatesRoot"
}

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3

    $results = foreach ($template in $templates) {
        $text = Get-WordDocumentText -Word $word -SourcePath $template.FullName

        $placeholders = @(Get-Placeholders -Text $text)
        $counts = [ordered]@{}
        foreach ($placeholder in $placeholders) {
            if ($counts.Contains($placeholder)) { $counts[$placeholder]++ }
            else { $counts[$placeholder] = 1 }
        }
        $assetSlots = @(Get-AssetSlots -Placeholders $placeholders)
        $assetCapacity = if ($assetSlots.Count -eq 0) { "Không thấy placeholder tài sản" }
        else { "Tài sản " + ($assetSlots -join ", ") }

        [pscustomobject]@{
            Name = $template.Name
            Extension = $template.Extension
            Category = Get-TemplateCategory -FileName $template.Name
            Placeholders = $counts
            DistinctCount = $counts.Count
            TotalCount = $placeholders.Count
            AssetCapacity = $assetCapacity
        }
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# Kiểm kê placeholder mẫu Word")
    $lines.Add("")
    $lines.Add("- Ngày tạo: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss K")")
    $lines.Add("- Phạm vi: tất cả file .docx và .doc trực tiếp trong thư mục templates/word/.")
    $lines.Add("- Cách đọc: Word COM StoryRanges trong phiên riêng, chỉ đọc cho cả .docx và .doc.")
    $lines.Add("- Quy tắc phân loại: chỉ dựa vào tên file; mẫu không đủ dấu hiệu được ghi Chưa phân loại.")
    $lines.Add("")
    $lines.Add("## Tổng quan")
    $lines.Add("")
    $lines.Add("| Mẫu | Nhóm | Định dạng | Placeholder khác nhau | Tổng lần xuất hiện | Sức chứa tài sản |")
    $lines.Add("| --- | --- | --- | ---: | ---: | --- |")
    foreach ($result in $results) {
        $lines.Add("| $(ConvertTo-MarkdownCell $result.Name) | $($result.Category) | $($result.Extension) | $($result.DistinctCount) | $($result.TotalCount) | $($result.AssetCapacity) |")
    }

    foreach ($result in $results) {
        $lines.Add("")
        $lines.Add("## $($result.Name)")
        $lines.Add("")
        $lines.Add("- Nhóm: $($result.Category)")
        $lines.Add("- Sức chứa tài sản theo placeholder: $($result.AssetCapacity)")
        $lines.Add("- Đọc trực tiếp từ mẫu gốc: Có")
        $lines.Add("")
        $lines.Add("| Placeholder | Số lần |")
        $lines.Add("| --- | ---: |")
        foreach ($placeholder in @($result.Placeholders.Keys | Sort-Object)) {
            $lines.Add("| $(ConvertTo-MarkdownCell $placeholder) | $($result.Placeholders[$placeholder]) |")
        }
    }

Set-Content -LiteralPath $ReportPath -Value $lines -Encoding UTF8
Write-Output "REPORT=$ReportPath"
Write-Output "TEMPLATES=$($results.Count)"
}
finally {
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

