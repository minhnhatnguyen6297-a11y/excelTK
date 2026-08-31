param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [Parameter(Mandatory = $true)]
    [string]$SheetName,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$fullOutputPath = [System.IO.Path]::GetFullPath($OutputPath)

if (Test-Path -LiteralPath $fullOutputPath) {
    throw "QA file already exists: $fullOutputPath"
}

$excel = $null
$sourceBook = $null
$viewBook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3

    $sourceBook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $SourcePath).Path, 0, $true)
    $sourceBook.Worksheets.Item($SheetName).Copy()
    $viewBook = $excel.ActiveWorkbook
    $viewBook.SaveAs($fullOutputPath, 51)
}
finally {
    if ($null -ne $viewBook) { $viewBook.Close($false) }
    if ($null -ne $sourceBook) { $sourceBook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
