[CmdletBinding()]
param(
    [string]$TemplatesPath = "",
    [switch]$ListOnly,
    [switch]$RestoreNumberedLabels
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
if ([string]::IsNullOrWhiteSpace($TemplatesPath)) {
    $TemplatesPath = Join-Path (Split-Path -Parent $PSScriptRoot) "templates\word"
}

function Get-XmlEntries {
    param([string]$Root)
    Get-ChildItem -LiteralPath $Root -Recurse -File -Filter *.xml |
        Where-Object {
            $_.FullName -match '\\word\\(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$'
        }
}

function Save-XmlUtf8NoBom {
    param([System.Xml.XmlDocument]$Document, [string]$Path)
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Document.OuterXml, $encoding)
}

function Normalize-LegacyKey {
    param([string]$Value)
    $text = $Value.Trim().ToLowerInvariant()
    # FormD removes Vietnamese tone marks; đ/Đ need an explicit replacement.
    $text = $text.Replace(([char]0x0111).ToString(), 'd')
    $text = $text.Replace(([char]0x0110).ToString(), 'd')
    $text = $text.Normalize([System.Text.NormalizationForm]::FormD)
    $builder = [System.Text.StringBuilder]::new()
    foreach ($char in $text.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if (($char -ge 'a' -and $char -le 'z') -or ($char -ge '0' -and $char -le '9')) {
            [void]$builder.Append($char)
        } elseif ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            # Other punctuation and symbols are intentionally dropped.
        }
    }
    $builder.ToString()
}

function Convert-LegacyToken {
    param([string]$Token)
    $inner = $Token.Substring(1, $Token.Length - 2).Trim()
    if ($inner -match '^(?<name>.+?)\s+(?<slot>\d+)$') {
        $base = Normalize-LegacyKey $Matches.name
        $slot = [int]$Matches.slot
    }
    else {
        $base = Normalize-LegacyKey $inner
        $slot = 0
    }
    if ([string]::IsNullOrWhiteSpace($base)) { return "#PLACEHOLDER_SAI:$inner#" }

    # Fields in the old samples that represent a person are slot-based.  The
    # asset fields are also slot-based; a missing suffix means card 1.
    $personBases = @('ten','namsinh','namchet','cccd','ngaycap','diachi','loaicc','noicapcc','thuongtru')
    $assetBases = @('loaiso','serial','sovaoso','sothua','soto','diachidat','dientich',
                    'hinhthucsudung','loaidat','thoihan','ont','cln','nts','luc',
                    'nguongoc','ngaycapso','coquancapso','ghichu')
    if ($personBases -contains $base -or $assetBases -contains $base) {
        if ($slot -eq 0) { $slot = 1 }
        return "{{$base$slot}}"
    }

    # Profile fields are single values.  A number in the original label is
    # retained because it is part of the field name (for example Nguoi uy quyen 2).
    $profileBases = @('niemyet','socongchung','nguoiuyquyen','nguoiuyquyen2')
    if ($slot -gt 0 -and ($base -eq 'nguoiuyquyen' -or $base -eq 'nguoiuyquyen2')) {
        $base = "$base$slot"
        $slot = 0
    }
    if ($profileBases -contains $base) { return "{{$base}}" }

    # Keep an unknown legacy label as a valid-looking token when possible.
    # The VBA parser will then emit #KHONG_CO_TRUONG instead of silently
    # deleting content.  A name that starts with a number remains invalid and
    # is intentionally left visible by the parser.
    if ($slot -gt 0) { return "{{$base$slot}}" }
    return "{{$base}}"
}

function Convert-ParagraphTokens {
    param(
        [System.Xml.XmlElement]$Paragraph,
        [switch]$ListOnly,
        [switch]$RestoreNumberedLabels
    )
    $textNodes = @($Paragraph.SelectNodes('.//*[local-name()="t"]'))
    if ($textNodes.Count -eq 0) { return @() }
    $parts = foreach ($node in $textNodes) { [string]$node.InnerText }
    $joined = $parts -join ''
    $pattern = if ($RestoreNumberedLabels) {
        '\{\{[0-9]+(?:\.[0-9]+)*\}\}'
    } else {
        '\[[^\]\r\n]+\]'
    }
    $matches = @([regex]::Matches($joined, $pattern))
    if ($matches.Count -eq 0) { return @() }

    $converted = [System.Collections.Generic.List[object]]::new()
    foreach ($match in $matches) {
        $old = $match.Value
        if ($RestoreNumberedLabels) {
            $new = '[' + $old.Substring(2, $old.Length - 4) + ']'
        } else {
            $new = Convert-LegacyToken $old
        }
        $converted.Add([pscustomobject]@{ Old = $old; New = $new })
    }
    if ($ListOnly) { return $converted }

    # Work backwards so offsets remain valid while text is moved between runs.
    $states = [System.Collections.Generic.List[object]]::new()
    $offset = 0
    foreach ($node in $textNodes) {
        $value = [string]$node.InnerText
        $states.Add([pscustomobject]@{ Node = $node; Start = $offset; End = $offset + $value.Length; Value = $value })
        $offset += $value.Length
    }
    for ($matchIndex = $matches.Count - 1; $matchIndex -ge 0; $matchIndex--) {
        $match = $matches[$matchIndex]
        $replacement = $converted[$matchIndex].New
        $start = $match.Index
        $end = $match.Index + $match.Length
        $first = $null
        foreach ($state in $states) {
            if ($state.End -gt $start -and $state.Start -lt $end) {
                if ($null -eq $first) { $first = $state }
                $localStart = [Math]::Max(0, $start - $state.Start)
                $localEnd = [Math]::Min($state.Value.Length, $end - $state.Start)
                $before = $state.Value.Substring(0, $localStart)
                $after = $state.Value.Substring($localEnd)
                if ($state -eq $first) { $state.Value = $before + $replacement + $after }
                else { $state.Value = $before + $after }
            }
        }
        if ($null -eq $first) { throw "Không tìm thấy node cho placeholder $($match.Value)" }
        foreach ($state in $states) { $state.Node.InnerText = $state.Value }
    }
    return $converted
}

function Process-Docx {
    param(
        [System.IO.FileInfo]$File,
        [switch]$ListOnly,
        [switch]$RestoreNumberedLabels
    )
    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ("word-convert-" + [guid]::NewGuid().ToString('N'))
    [System.IO.Directory]::CreateDirectory($temp) | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($File.FullName, $temp)
        $all = [System.Collections.Generic.List[object]]::new()
        foreach ($xmlFile in Get-XmlEntries $temp) {
            $xml = [System.Xml.XmlDocument]::new()
            $xml.PreserveWhitespace = $true
            $xml.Load($xmlFile.FullName)
            foreach ($paragraph in @($xml.SelectNodes('//*[local-name()="p"]'))) {
                foreach ($item in @(Convert-ParagraphTokens $paragraph -ListOnly:$ListOnly -RestoreNumberedLabels:$RestoreNumberedLabels)) {
                    $all.Add([pscustomobject]@{ File = $File.Name; Old = $item.Old; New = $item.New })
                }
            }
            if (-not $ListOnly) { Save-XmlUtf8NoBom -Document $xml -Path $xmlFile.FullName }
        }
        if (-not $ListOnly) {
            $output = $File.FullName + '.converted'
            [System.IO.Compression.ZipFile]::CreateFromDirectory($temp, $output, [System.IO.Compression.CompressionLevel]::Optimal, $false)
            Move-Item -LiteralPath $output -Destination $File.FullName -Force
        }
        return $all
    }
    finally {
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force }
    }
}

$root = (Resolve-Path -LiteralPath $TemplatesPath).Path
$files = @(Get-ChildItem -LiteralPath $root -File -Filter *.docx | Sort-Object Name)
if ($files.Count -eq 0) { throw "Không tìm thấy mẫu .docx trong $root" }
$changes = [System.Collections.Generic.List[object]]::new()
foreach ($file in $files) {
    foreach ($item in @(Process-Docx -File $file -ListOnly:$ListOnly -RestoreNumberedLabels:$RestoreNumberedLabels)) { $changes.Add($item) }
}
$changes | Group-Object File, Old, New | Sort-Object Name | ForEach-Object {
    "{0}`t{1}`t{2}`tcount={3}" -f $_.Group[0].File, $_.Group[0].Old, $_.Group[0].New, $_.Count
}
"FILES=$($files.Count)"
"CONVERSIONS=$($changes.Count)"
