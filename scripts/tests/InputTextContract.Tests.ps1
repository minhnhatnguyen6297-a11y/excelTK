$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) "build-inheritance-workbook.ps1"

Describe "visible input Text contract" {
    BeforeAll {
        $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
    }

    It "formats every standard Person text input range as Text" {
        $scriptText | Should Match '\$inputSheet\.Range\("B9:B38"\)\.NumberFormat = "@"'
        $scriptText | Should Match '\$inputSheet\.Range\("C9:D38"\)\.NumberFormat = "@"'
        $scriptText | Should Match '\$inputSheet\.Range\("E9:E38"\)\.NumberFormat = "@"'
        $scriptText | Should Match '\$inputSheet\.Range\("F9:F38"\)\.NumberFormat = "@"'
        $scriptText | Should Match '\$inputSheet\.Range\("G9:G38"\)\.NumberFormat = "@"'
    }

    It "keeps visible Asset values and the open-ended Profile values as Text" {
        $scriptText | Should Match '\$inputSheet\.Range\("AB9:AB38,AD9:AD38,AF9:AF38"\)\.NumberFormat = "@"'
        $scriptText | Should Match '\$inputSheet\.Range\("B45:B1048576"\)\.NumberFormat = "@"'
        $scriptText | Should Match '\$inputSheet\.Range\("C41:C1048576"\)\.NumberFormat = "@"'
    }
}
