<#
.SYNOPSIS
  Record whether MathType/OLE objects expose a trustworthy structured source.

.DESCRIPTION
  This is a read-only Open XML package probe. It records OLE relationship and
  embedding evidence but deliberately does not parse MTEF or infer TeX from
  binary data. A missing structured MathML/TeX export is a successful probe
  result with OriginalKept, not a conversion failure.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CarrierInventoryJson,
    [Parameter(Mandatory = $true)][string]$ReferencePptxPath,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-FileSha256Local { param([string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-TextSha256Local {
    param([string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))).Replace('-', '').ToLowerInvariant()) }
    finally { $sha.Dispose() }
}
function Get-EntrySha256 {
    param([IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try { $hash = [Security.Cryptography.SHA256]::Create(); try { return ([BitConverter]::ToString($hash.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() } finally { $hash.Dispose() } }
    finally { $stream.Dispose() }
}
function Get-OleBinding {
    param([IO.Compression.ZipArchive]$Zip, $Source)
    $slidePart = [string]$Source.packagePart
    $slideText = Read-ZipEntryText -Zip $Zip -EntryName $slidePart
    $relsPart = (Split-Path -Parent $slidePart).Replace('\', '/') + '/_rels/' + (Split-Path -Leaf $slidePart) + '.rels'
    $relsText = Read-ZipEntryText -Zip $Zip -EntryName $relsPart
    $relTargets = @{}
    foreach ($match in [regex]::Matches($relsText, 'Id="([^"]+)"[^>]*Target="([^"]+)"')) {
        $target = $match.Groups[2].Value
        if ($target -like '../*') { $target = 'ppt/' + $target.Substring(3) }
        $relTargets[$match.Groups[1].Value] = $target
    }
    foreach ($frame in [regex]::Matches($slideText, '<p:graphicFrame>.*?</p:graphicFrame>', 'Singleline')) {
        $id = [regex]::Match($frame.Value, '<p:cNvPr id="(\d+)"')
        $rid = [regex]::Match($frame.Value, 'r:id="([^"]+)"')
        if ($id.Success -and $rid.Success -and [int]$id.Groups[1].Value -eq [int]$Source.shapeId) {
            $part = [string]$relTargets[$rid.Groups[1].Value]
            return [ordered]@{ embeddingPart = $part; embeddingSha256 = Get-EntrySha256 -Zip $Zip -EntryName $part }
        }
    }
    return [ordered]@{ embeddingPart = ''; embeddingSha256 = $null }
}

$CarrierInventoryJson = [IO.Path]::GetFullPath($CarrierInventoryJson)
$ReferencePptxPath = [IO.Path]::GetFullPath($ReferencePptxPath)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
foreach ($path in @($CarrierInventoryJson, $ReferencePptxPath)) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required input not found: $path" } }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$inventory = Get-Content -LiteralPath $CarrierInventoryJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$inventory.schemaVersion -ne 1) { throw "Unsupported carrier inventory schemaVersion: $($inventory.schemaVersion)" }
$sourceHash = Get-FileSha256Local -Path $ReferencePptxPath
if ($sourceHash -ne ([string]$inventory.input.sha256).ToLowerInvariant()) { throw 'Reference PPTX hash does not match the carrier inventory.' }

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream = [IO.File]::Open($ReferencePptxPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
$zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read)
try {
    $records = New-Object System.Collections.Generic.List[object]
    foreach ($record in @($inventory.records | Where-Object { [string]$_.source.carrier -eq 'MathTypeOle' })) {
        $binding = Get-OleBinding -Zip $zip -Source $record.source
        $cjk = Get-OleEquationCjkFingerprint -PptxPath $ReferencePptxPath -SlideNumber ([int]$record.source.slide) -ShapeId ([int]$record.source.shapeId)
        $records.Add([ordered]@{
            recordId = [string]$record.recordId
            source = [ordered]@{ carrier = 'MathTypeOle'; slide = [int]$record.source.slide; shapeId = [int]$record.source.shapeId; sourceSha256 = [string]$record.source.sourceSha256; oleProgId = [string]$record.source.oleProgId }
            package = [ordered]@{ slidePart = [string]$record.source.packagePart; embeddingPart = $binding.embeddingPart; embeddingSha256 = $binding.embeddingSha256 }
            probe = [ordered]@{ structuredSourceStatus = 'Unavailable'; extractedFormat = 'None'; cjkFingerprint = $cjk; mtefParsed = $false; reason = 'Open XML exposes an opaque OLE embedding only. No reliable MathML/TeX export is present, and this probe never parses or reconstructs MTEF.' }
            decision = [ordered]@{ status = 'OriginalKept'; targetCarrier = 'Original'; writeBackAllowed = $false; reason = 'No trustworthy structured MathType source was exposed by the package.' }
        }) | Out-Null
    }
} finally {
    $zip.Dispose()
    $stream.Dispose()
}
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ readOnly = $true; writeBackAllowed = $false; mtefParsingAllowed = $false; fallback = 'OriginalKept' }
    input = [ordered]@{ inventoryPath = $CarrierInventoryJson; inventorySha256 = Get-FileSha256Local -Path $CarrierInventoryJson; pptxPath = $ReferencePptxPath; pptxSha256 = $sourceHash }
    counts = [ordered]@{ total = $records.Count; structuredSourceAvailable = 0; originalKept = $records.Count }
    evidenceSetSha256 = Get-TextSha256Local -Text ((@($records | ForEach-Object { "$($_.recordId)|$($_.source.sourceSha256)|$($_.package.embeddingPart)|$($_.package.embeddingSha256)|$($_.probe.structuredSourceStatus)" }) | Sort-Object) -join "`n")
    records = @($records.ToArray())
}
$manifest | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath (Join-Path $OutputDir 'formula-ole-structured-source-probe.json') -Encoding UTF8
Write-Utf8BomCsv -InputObject @($records | ForEach-Object { [pscustomobject]@{ RecordId = $_.recordId; Slide = $_.source.slide; ShapeId = $_.source.shapeId; ProgId = $_.source.oleProgId; EmbeddingPart = $_.package.embeddingPart; StructuredSourceStatus = $_.probe.structuredSourceStatus; CjkFingerprint = $_.probe.cjkFingerprint; DecisionStatus = $_.decision.status; WriteBackAllowed = $_.decision.writeBackAllowed } }) -Path (Join-Path $OutputDir 'formula-ole-structured-source-probe.csv')
Write-Host "MathType structured-source probe complete: $OutputDir"
Write-Host "OLE records: $($records.Count); structured sources: 0; OriginalKept: $($records.Count)"
