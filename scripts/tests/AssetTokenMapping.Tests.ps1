$sourcePath = Join-Path (Split-Path -Parent $PSScriptRoot) "..\src\vba\modExportData.bas"

Describe "asset placeholder label mapping contract" {
    BeforeAll {
        $sourceText = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
        $helperMatch = [regex]::Match(
            $sourceText,
            '(?ms)^Private Function AssetTokenBaseForLabel\(.*?^End Function'
        )
        $helperText = $helperMatch.Value
    }

    It "maps the descriptive UI labels to the canonical SOT token roots" {
        $helperText | Should Match 'Case "sophathanhserial".*AssetTokenBaseForLabel = "serial"'
        $helperText | Should Match 'Case "sotobando".*AssetTokenBaseForLabel = "soto"'
        $helperText | Should Match 'Case "dientichm".*AssetTokenBaseForLabel = "dientich"'
        $helperText | Should Match 'Case "thoihansudung".*AssetTokenBaseForLabel = "thoihan"'
        $helperText | Should Match 'Case "ontm".*AssetTokenBaseForLabel = "ont"'
        $helperText | Should Match 'Case "clnm".*AssetTokenBaseForLabel = "cln"'
        $helperText | Should Match 'Case "ntsm".*AssetTokenBaseForLabel = "nts"'
        $helperText | Should Match 'Case "lucm".*AssetTokenBaseForLabel = "luc"'
    }
}
