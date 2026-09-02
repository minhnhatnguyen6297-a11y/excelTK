[CmdletBinding()]
param(
    [string]$TemplatesPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "templates\\word"),
    [string]$ReportPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "output\\word-placeholder-audit.md"),
    [switch]$FailOnViolation
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

function Get-TokenInventory {
    param([string]$Text)

    $tokens = [System.Collections.Generic.List[object]]::new()
    $matches = [regex]::Matches($Text, '\[[^\]\r\n]+\]|\{\{[^\r\n]*?\}\}')

    foreach ($match in $matches) {
        $token = $match.Value.Trim()
        $kind = if ($token.StartsWith("[")) {
            "Ngoặc vuông cũ"
        }
        elseif ($token -match '^\{\{[a-z][a-z0-9]*\}\}$') {
            "Tĩnh hợp lệ"
        }
        else {
            "Tĩnh sai quy tắc"
        }

        $tokens.Add([pscustomobject]@{
            Token = $token
            Kind = $kind
        })
    }

    return @($tokens)
}

function ConvertTo-MarkdownCell {
    param([string]$Value)

    return $Value.Replace("|", "\\|").Replace("`r", " ").Replace("`n", " ")
}

$templates = @(Get-ChildItem -LiteralPath $templatesRoot -File -Recurse |
    Where-Object { $_.Extension -in ".doc", ".docx" } |
    Sort-Object FullName)
if ($templates.Count -eq 0) {
    throw "Không tìm thấy mẫu .docx hoặc .doc trong $templatesRoot"
}

$word = $null
$violationCount = 0
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3

    $results = @(foreach ($template in $templates) {
        $text = Get-WordDocumentText -Word $word -SourcePath $template.FullName
        $inventory = @(Get-TokenInventory -Text $text)
        $counts = [ordered]@{}
        foreach ($item in $inventory) {
            if ($counts.Contains($item.Token)) { $counts[$item.Token]++ }
            else { $counts[$item.Token] = 1 }
        }

        $legacyTokens = @($inventory | Where-Object { $_.Kind -eq "Ngoặc vuông cũ" } | ForEach-Object { $_.Token } | Sort-Object -Unique)
        $invalidTokens = @($inventory | Where-Object { $_.Kind -eq "Tĩnh sai quy tắc" } | ForEach-Object { $_.Token } | Sort-Object -Unique)
        $duplicateTokens = @($counts.GetEnumerator() |
            Where-Object { $_.Value -gt 1 } |
            ForEach-Object { $_.Key } |
            Sort-Object)
        $violations = [System.Collections.Generic.List[string]]::new()

        if ($template.Extension -eq ".doc") {
            $violations.Add("Mẫu .doc phải chuyển sang .docx hoặc được ghi rõ chưa hỗ trợ.")
        }
        foreach ($token in $legacyTokens) {
            $violations.Add("Còn placeholder ngoặc vuông: $token")
        }
        foreach ($token in $invalidTokens) {
            $violations.Add("Placeholder {{...}} sai quy tắc: $token")
        }

        $violationCount += $violations.Count
        [pscustomobject]@{
            Name = $template.Name
            RelativePath = $template.FullName.Substring($templatesRoot.Length).TrimStart('\', '/')
            Extension = $template.Extension
            Category = Get-TemplateCategory -FileName $template.Name
            Placeholders = $counts
            ValidTokens = @($inventory | Where-Object { $_.Kind -eq "Tĩnh hợp lệ" } | ForEach-Object { $_.Token } | Sort-Object -Unique)
            InvalidTokens = $invalidTokens
            LegacyTokens = $legacyTokens
            DuplicateTokens = $duplicateTokens
            Violations = @($violations)
        }
    })

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# Kiểm kê placeholder mẫu Word")
    $lines.Add("")
    $lines.Add("- Ngày tạo: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')")
    $lines.Add("- Phạm vi: tất cả file .docx và .doc trong templates/word/, gồm cả thư mục con.")
    $lines.Add("- Cách đọc: Word COM StoryRanges trong phiên riêng, chỉ đọc.")
    $lines.Add("- Token tĩnh hợp lệ: {{ten1}} — bắt đầu bằng chữ thường, sau đó chỉ có chữ thường hoặc số.")
    $lines.Add("- Chế độ nghiêm: $([bool]$FailOnViolation)")
    $lines.Add("")
    $lines.Add("## Tổng quan")
    $lines.Add("")
    $lines.Add("| Mẫu | Nhóm | Định dạng | Token hợp lệ | Token sai | Ngoặc vuông | Token lặp | Vi phạm |")
    $lines.Add("| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |")
    foreach ($result in $results) {
        $lines.Add("| $(ConvertTo-MarkdownCell $result.RelativePath) | $($result.Category) | $($result.Extension) | $($result.ValidTokens.Count) | $($result.InvalidTokens.Count) | $($result.LegacyTokens.Count) | $($result.DuplicateTokens.Count) | $($result.Violations.Count) |")
    }

    foreach ($result in $results) {
        $validTokenText = if ($result.ValidTokens.Count -eq 0) { "Không có" } else { $result.ValidTokens -join ", " }
        $invalidTokenText = if ($result.InvalidTokens.Count -eq 0) { "Không có" } else { $result.InvalidTokens -join ", " }
        $legacyTokenText = if ($result.LegacyTokens.Count -eq 0) { "Không có" } else { $result.LegacyTokens -join ", " }
        $duplicateTokenText = if ($result.DuplicateTokens.Count -eq 0) { "Không có" } else { $result.DuplicateTokens -join ", " }

        $lines.Add("")
        $lines.Add("## $($result.RelativePath)")
        $lines.Add("")
        $lines.Add("- Nhóm: $($result.Category)")
        $lines.Add("- Token hợp lệ: $validTokenText")
        $lines.Add("- Token sai: $invalidTokenText")
        $lines.Add("- Ngoặc vuông cũ: $legacyTokenText")
        $lines.Add("- Token xuất hiện nhiều lần: $duplicateTokenText")
        $lines.Add("")
        $lines.Add("### Vi phạm")
        if ($result.Violations.Count -eq 0) {
            $lines.Add("- Không có")
        }
        else {
            foreach ($violation in $result.Violations) {
                $lines.Add("- $violation")
            }
        }

        $lines.Add("")
        $lines.Add("| Token | Số lần |")
        $lines.Add("| --- | ---: |")
        foreach ($token in @($result.Placeholders.Keys | Sort-Object)) {
            $lines.Add("| $(ConvertTo-MarkdownCell $token) | $($result.Placeholders[$token]) |")
        }
    }

    Set-Content -LiteralPath $ReportPath -Value $lines -Encoding UTF8
    Write-Output "REPORT=$ReportPath"
    Write-Output "TEMPLATES=$($results.Count)"
    Write-Output "VIOLATIONS=$violationCount"
}
finally {
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($FailOnViolation -and $violationCount -gt 0) {
    exit 1
}
