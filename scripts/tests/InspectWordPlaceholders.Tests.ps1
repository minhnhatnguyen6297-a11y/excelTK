$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) "inspect-word-placeholders.ps1"

function New-WordFixture {
    param(
        [string]$Path,
        [string]$Text,
        [int]$FileFormat = 16
    )

    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 3
        $document = $word.Documents.Add()
        $document.Content.Text = $Text
        $document.SaveAs2($Path, $FileFormat)
    }
    finally {
        if ($null -ne $document) { $document.Close($false) }
        if ($null -ne $word) { $word.Quit() }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Describe "inspect-word-placeholders" {
    It "rejects a malformed static token containing an early closing brace" {
        $templatesPath = Join-Path $TestDrive "malformed-templates"
        $reportPath = Join-Path $TestDrive "malformed-report.md"
        New-Item -ItemType Directory -Path $templatesPath | Out-Null
        New-WordFixture -Path (Join-Path $templatesPath "MauSai.docx") -Text "{{ten}1}}"

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

        $LASTEXITCODE | Should Be 1
        (Test-Path -LiteralPath $reportPath) | Should Be $true
        (Get-Content -Raw -LiteralPath $reportPath) | Should Match "Placeholder \{\{\.\.\.\}\} sai quy tắc: \{\{ten\}1\}\}"
    }

    It "writes violations for square brackets and invalid static tokens in strict mode" {
        $templatesPath = Join-Path $TestDrive "templates"
        $reportPath = Join-Path $TestDrive "report.md"
        New-Item -ItemType Directory -Path $templatesPath | Out-Null
        New-WordFixture -Path (Join-Path $templatesPath "Mau.docx") -Text "[Ten 1] {{ten1}} {{ten_2}}"

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

        $LASTEXITCODE | Should Be 1
        (Test-Path -LiteralPath $reportPath) | Should Be $true
        (Get-Content -Raw -LiteralPath $reportPath) | Should Match "\[Ten 1\]"
        (Get-Content -Raw -LiteralPath $reportPath) | Should Match "\{\{ten_2\}\}"
    }

    It "accepts a single template that only has valid static tokens in strict mode" {
        $templatesPath = Join-Path $TestDrive "valid-templates"
        $reportPath = Join-Path $TestDrive "valid-report.md"
        New-Item -ItemType Directory -Path $templatesPath | Out-Null
        New-WordFixture -Path (Join-Path $templatesPath "MauHopLe.docx") -Text "{{ten1}} {{cccd2}}"

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

        $LASTEXITCODE | Should Be 0
        (Get-Content -Raw -LiteralPath $reportPath) | Should Match "Token hợp lệ: \{\{cccd2\}\}, \{\{ten1\}\}"
        (Get-Content -Raw -LiteralPath $reportPath) | Should Match "- Không có"
    }

    It "audits docx templates in nested directories" {
        $templatesPath = Join-Path $TestDrive "recursive-templates"
        $nestedPath = Join-Path $templatesPath "nhom-con"
        $reportPath = Join-Path $TestDrive "recursive-report.md"
        New-Item -ItemType Directory -Path $nestedPath -Force | Out-Null
        New-WordFixture -Path (Join-Path $nestedPath "MauCon.docx") -Text "{{ten1}}"

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

        $LASTEXITCODE | Should Be 0
        (Get-Content -Raw -LiteralPath $reportPath) | Should Match ([regex]::Escape((Join-Path "nhom-con" "MauCon.docx")))
    }

    It "reports a legacy doc template as a strict-mode violation" {
        $templatesPath = Join-Path $TestDrive "doc-templates"
        $reportPath = Join-Path $TestDrive "doc-report.md"
        New-Item -ItemType Directory -Path $templatesPath | Out-Null
        New-WordFixture -Path (Join-Path $templatesPath "MauCu.doc") -Text "{{ten1}}" -FileFormat 0

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

        $LASTEXITCODE | Should Be 1
        (Get-Content -Raw -LiteralPath $reportPath) | Should Match "Mẫu \.doc phải chuyển sang \.docx hoặc được ghi rõ chưa hỗ trợ\."
    }

    It "reports a repeated token and its occurrence count" {
        $templatesPath = Join-Path $TestDrive "duplicate-templates"
        $reportPath = Join-Path $TestDrive "duplicate-report.md"
        New-Item -ItemType Directory -Path $templatesPath | Out-Null
        New-WordFixture -Path (Join-Path $templatesPath "MauLap.docx") -Text "{{ten1}} {{ten1}}"

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

        $LASTEXITCODE | Should Be 0
        $report = Get-Content -Raw -LiteralPath $reportPath
        $report | Should Match "Token xuất hiện nhiều lần: \{\{ten1\}\}"
        $report | Should Match "\| \{\{ten1\}\} \| 2 \|"
    }

    It "does not modify a template while auditing it" {
        $templatesPath = Join-Path $TestDrive "unchanged-templates"
        $templatePath = Join-Path $templatesPath "MauKhongDoi.docx"
        $reportPath = Join-Path $TestDrive "unchanged-report.md"
        New-Item -ItemType Directory -Path $templatesPath | Out-Null
        New-WordFixture -Path $templatePath -Text "{{ten1}}"
        $hashBefore = (Get-FileHash -Algorithm SHA256 -LiteralPath $templatePath).Hash

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

        $LASTEXITCODE | Should Be 0
        (Get-FileHash -Algorithm SHA256 -LiteralPath $templatePath).Hash | Should Be $hashBefore
    }

    It "leaves an already-running Word session and its document open" {
        $templatesPath = Join-Path $TestDrive "session-templates"
        $reportPath = Join-Path $TestDrive "session-report.md"
        New-Item -ItemType Directory -Path $templatesPath | Out-Null
        New-WordFixture -Path (Join-Path $templatesPath "MauPhienRieng.docx") -Text "{{ten1}}"

        $userWord = $null
        $userDocument = $null
        try {
            $userWord = New-Object -ComObject Word.Application
            $userWord.Visible = $false
            $userWord.DisplayAlerts = 0
            $userDocument = $userWord.Documents.Add()
            $userDocument.Content.Text = "PHIEN WORD NGUOI DUNG"

            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -TemplatesPath $templatesPath -ReportPath $reportPath -FailOnViolation

            $LASTEXITCODE | Should Be 0
            $userWord.Documents.Count | Should Be 1
            $userDocument.Content.Text | Should Match "PHIEN WORD NGUOI DUNG"
        }
        finally {
            if ($null -ne $userDocument) { $userDocument.Close($false) }
            if ($null -ne $userWord) { $userWord.Quit() }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
