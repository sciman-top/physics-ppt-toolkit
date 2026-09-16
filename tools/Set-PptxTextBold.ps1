<#
.SYNOPSIS
  Set bold on every text run of a copied PPTX to improve projection visibility.

.DESCRIPTION
  Copies the source PPTX and writes b="1" onto every a:rPr (body runs, math
  runs, math ctrlPr runs) and every a:endParaRPr in all slides. Colours, fonts,
  sizes, italic and positions are untouched, and the source PPTX is never
  modified.

.PARAMETER InputPath
  Source .pptx file.

.PARAMETER OutputPath
  New PPTX path to save.

.PARAMETER ReportPath
  CSV report path. Defaults next to OutputPath.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$ReportPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$script:NsA = 'http://schemas.openxmlformats.org/drawingml/2006/main'

function Convert-XmlDocumentToString {
    param([System.Xml.XmlDocument]$Document)
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.OmitXmlDeclaration = $true
    $builder = New-Object System.Text.StringBuilder
    $writer = [System.Xml.XmlWriter]::Create($builder, $settings)
    try {
        $Document.Save($writer)
    } finally {
        if ($null -ne $writer) { $writer.Dispose() }
    }
    return $builder.ToString()
}

function Write-ZipEntryText {
    param($Zip, [string]$EntryName, [string]$Text)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { throw "Zip entry not found: $EntryName" }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $bytes = $utf8NoBom.GetBytes($Text)
    $entry.Delete() | Out-Null
    $newEntry = $Zip.CreateEntry($EntryName)
    $stream = $newEntry.Open()
    try {
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path (Split-Path -Parent $OutputPath) 'text-bold-report.csv'
}
$ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
$reportDir = Split-Path -Parent $ReportPath
if (-not (Test-Path -LiteralPath $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }

foreach ($path in @($InputPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required input not found: $path" }
}
$outDir = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
if ([string]::Equals($InputPath, $OutputPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputPath must differ from InputPath; the source PPTX is never modified in place."
}

Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$reportRows = New-Object System.Collections.Generic.List[object]
$zip = $null
try {
    $zip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
    $slideEntries = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } |
        Sort-Object { [int]([regex]::Match($_.FullName, 'slide(\d+)\.xml$').Groups[1].Value) })
    foreach ($entry in $slideEntries) {
        $entryName = $entry.FullName
        $slideNo = [int]([regex]::Match($entryName, 'slide(\d+)\.xml$').Groups[1].Value)
        $doc = New-Object System.Xml.XmlDocument
        $doc.PreserveWhitespace = $true
        $doc.LoadXml((Read-ZipEntryText -Zip $zip -EntryName $entryName))
        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace('a', $script:NsA)

        $bolded = 0
        $alreadyBold = 0
        foreach ($rPr in @($doc.SelectNodes('//a:rPr', $ns)) + @($doc.SelectNodes('//a:endParaRPr', $ns))) {
            $current = [string]$rPr.GetAttribute('b')
            if ($current -eq '1') { $alreadyBold++; continue }
            $rPr.SetAttribute('b', '1')
            $bolded++
        }
        if ($bolded -gt 0) {
            Write-ZipEntryText -Zip $zip -EntryName $entryName -Text (Convert-XmlDocumentToString -Document $doc)
        }
        $reportRows.Add([pscustomobject]@{
            Slide = $slideNo
            BoldSet = $bolded
            AlreadyBold = $alreadyBold
        }) | Out-Null
    }
} finally {
    if ($null -ne $zip) { $zip.Dispose() }
}

$totalSet = @($reportRows | Measure-Object -Property BoldSet -Sum).Sum
$totalAlready = @($reportRows | Measure-Object -Property AlreadyBold -Sum).Sum
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($ReportPath, (@($reportRows | ConvertTo-Csv -NoTypeInformation | ForEach-Object { [string]$_ })), $utf8Bom)
Write-Host ("Bold visibility done: set b=1 on {0} runs ({1} already bold); report: {2}" -f $totalSet, $totalAlready, $ReportPath)
