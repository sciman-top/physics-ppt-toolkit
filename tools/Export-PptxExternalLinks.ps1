<#
.SYNOPSIS
  Export external PPTX relationship links for classroom delivery checks.

.DESCRIPTION
  Scans .pptx packages for TargetMode="External" relationships and writes
  CSV/JSON reports. The script is read-only: it never modifies presentations.

.PARAMETER InputPath
  Path to a .pptx file or directory containing PPTX files.

.PARAMETER OutputDir
  Directory where reports are written.

.PARAMETER Recurse
  Search subdirectories when InputPath is a directory.

.PARAMETER FilePattern
  File name pattern used when InputPath is a directory.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [switch]$Recurse,

    [string]$FilePattern = '*.pptx'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PptxFiles {
    param([string]$Path, [string]$Pattern, [switch]$Recurse)
    if (-not (Test-Path -LiteralPath $Path)) { throw "InputPath not found: $Path" }
    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) {
        $opt = @{ LiteralPath = $item.FullName; Filter = $Pattern; File = $true }
        if ($Recurse) { $opt.Recurse = $true }
        return @(Get-ChildItem @opt | Where-Object { $_.Name -notlike '~$*' })
    }
    if ($item.Extension -ne '.pptx') { throw "Only .pptx files are supported: $($item.FullName)" }
    return @($item)
}

function Get-RelationshipTypeName {
    param([string]$Type)
    if ([string]::IsNullOrWhiteSpace($Type)) { return '' }
    return ($Type -replace '^.*/', '')
}

function Get-SlideNumberFromRelationshipPath {
    param([string]$EntryName)
    if ($EntryName -match '^ppt/slides/_rels/slide(\d+)\.xml\.rels$') {
        return [int]$Matches[1]
    }
    return 0
}

function Get-TargetInfo {
    param(
        [string]$Target,
        [string]$PresentationDir
    )

    $decoded = try { [System.Uri]::UnescapeDataString($Target) } catch { $Target }
    if ($decoded -match '^(https?|ftp)://') {
        return [pscustomobject]@{ Kind = 'Web'; LocalPath = ''; Exists = $null }
    }

    if ($decoded -match '^file:///') {
        try {
            $uri = [System.Uri]$decoded
            $local = $uri.LocalPath
            return [pscustomobject]@{ Kind = 'LocalFile'; LocalPath = $local; Exists = (Test-Path -LiteralPath $local) }
        } catch {
            return [pscustomobject]@{ Kind = 'LocalFile'; LocalPath = $decoded; Exists = $false }
        }
    }

    if ($decoded -match '^[A-Za-z]:\\|^\\\\') {
        return [pscustomobject]@{ Kind = 'LocalFile'; LocalPath = $decoded; Exists = (Test-Path -LiteralPath $decoded) }
    }

    $candidate = Join-Path $PresentationDir $decoded
    return [pscustomobject]@{ Kind = 'RelativeFile'; LocalPath = $candidate; Exists = (Test-Path -LiteralPath $candidate) }
}

function Get-LinkRisk {
    param(
        [string]$TypeName,
        [string]$TargetKind,
        [object]$Exists
    )

    if ($TargetKind -in @('LocalFile', 'RelativeFile') -and $Exists -eq $false) { return 'HighMissingLocalFile' }
    if ($TypeName -eq 'image' -and $TargetKind -in @('LocalFile', 'RelativeFile')) { return 'HighExternalImage' }
    if ($TypeName -eq 'hyperlink' -and $TargetKind -in @('LocalFile', 'RelativeFile')) { return 'MediumExternalMediaOrFile' }
    if ($TargetKind -eq 'Web') { return 'LowWebLink' }
    return 'Info'
}

function Read-EntryText {
    param($Entry)
    $stream = $Entry.Open()
    $reader = $null
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        return $reader.ReadToEnd()
    } finally {
        if ($null -ne $reader) { $reader.Dispose() } else { $stream.Dispose() }
    }
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$rows = New-Object System.Collections.Generic.List[object]
$files = @(Get-PptxFiles -Path $InputPath -Pattern $FilePattern -Recurse:$Recurse)
if ($files.Count -eq 0) { throw "No .pptx files found in $InputPath" }

foreach ($file in $files) {
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
        foreach ($entry in @($zip.Entries | Where-Object { $_.FullName -like '*.rels' })) {
            $xmlText = Read-EntryText -Entry $entry
            [xml]$xml = $xmlText
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
            foreach ($rel in @($xml.SelectNodes('//r:Relationship[@TargetMode="External"]', $ns))) {
                $typeName = Get-RelationshipTypeName -Type $rel.Type
                $targetInfo = Get-TargetInfo -Target $rel.Target -PresentationDir $file.DirectoryName
                $risk = Get-LinkRisk -TypeName $typeName -TargetKind $targetInfo.Kind -Exists $targetInfo.Exists
                $rows.Add([pscustomobject]@{
                    File = $file.Name
                    Slide = Get-SlideNumberFromRelationshipPath -EntryName $entry.FullName
                    Entry = $entry.FullName
                    RelationshipId = $rel.Id
                    Type = $typeName
                    Target = $rel.Target
                    TargetKind = $targetInfo.Kind
                    LocalPath = $targetInfo.LocalPath
                    TargetExists = $targetInfo.Exists
                    Risk = $risk
                }) | Out-Null
            }
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

$csvPath = Join-Path $OutputDir 'pptx-external-links.csv'
$jsonPath = Join-Path $OutputDir 'pptx-external-links.json'
$rows | Sort-Object File, Slide, Entry, RelationshipId | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

$summary = [pscustomobject]@{
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    input = $InputPath
    files = $files.Count
    linkCount = $rows.Count
    highRiskCount = @($rows | Where-Object { $_.Risk -like 'High*' }).Count
    missingLocalFileCount = @($rows | Where-Object { $_.Risk -eq 'HighMissingLocalFile' }).Count
    byRisk = @($rows | Group-Object Risk | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ risk = $_.Name; count = $_.Count }
    })
    links = @($rows | Sort-Object File, Slide, Entry, RelationshipId)
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

Write-Host "External link report written:"
Write-Host "  CSV: $csvPath"
Write-Host "  JSON: $jsonPath"
