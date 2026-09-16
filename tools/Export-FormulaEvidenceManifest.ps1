<#
.SYNOPSIS
  Create a relative-path, hash-bound formula evidence manifest.

.DESCRIPTION
  Converts a carrier inventory into formula-evidence-manifest.json.  This is
  a read-only evidence operation: it never writes a PPTX and it never changes
  a decision into Converted.  Paths are relative to the manifest directory;
  PPTX media are represented as verified package-entry references with their
  own hashes.

.PARAMETER CarrierInventoryJson
  formula-carrier-inventory.json produced by Export-FormulaCarrierInventory.

.PARAMETER OutputDir
  Directory that becomes the manifest path base.

.PARAMETER Mode
  Evidence policy mode recorded in the manifest.  The tool itself remains
  write-back disabled for every mode.

.PARAMETER RecognitionPath
  Optional OCR/model result file to bind into the manifest by hash.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CarrierInventoryJson,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [ValidateSet('ReportOnly', 'CandidateOnly', 'ReviewRequired', 'ClosedWorldUnattended', 'ExplicitMigration')]
    [string]$Mode = 'CandidateOnly',

    [string]$RecognitionPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$script:Sha256Pattern = '^[A-Fa-f0-9]{64}$'

function Get-Sha256TextLocal {
    param([Parameter(Mandatory = $true)][string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($Text)
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Get-RelativeManifestPath {
    param(
        [Parameter(Mandatory = $true)][string]$BaseDirectory,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )
    $base = [System.IO.Path]::GetFullPath($BaseDirectory)
    $target = [System.IO.Path]::GetFullPath($TargetPath)
    $relative = [System.IO.Path]::GetRelativePath($base, $target)
    if ([string]::IsNullOrWhiteSpace($relative)) { $relative = '.' }
    # .NET returns an absolute path when the two paths are on different
    # volumes.  Absolute paths would make a manifest machine-specific, so
    # fail closed and ask the caller to keep the evidence root on the source
    # volume (or provide a copied evidence file under the output root).
    if ([System.IO.Path]::IsPathFullyQualified($relative)) {
        throw "Cannot create a manifest-relative path across volumes: base=$base target=$target"
    }
    return $relative.Replace('\', '/')
}

function Get-FileEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$BaseDirectory,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Kind
    )
    $full = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "Evidence file not found: $full" }
    $item = Get-Item -LiteralPath $full
    return [ordered]@{
        path = Get-RelativeManifestPath -BaseDirectory $BaseDirectory -TargetPath $full
        sha256 = (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash.ToLowerInvariant()
        bytes = [int64]$item.Length
        kind = $Kind
    }
}

function Get-PackageEntryEvidence {
    param(
        [Parameter(Mandatory = $true)][System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory = $true)][string]$ArchiveRelativePath,
        [Parameter(Mandatory = $true)][string]$EntryName
    )
    $normalized = $EntryName.Replace('\', '/').TrimStart('/')
    $entry = $Zip.GetEntry($normalized)
    if ($null -eq $entry) { throw "PPTX package entry not found: $normalized" }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = $entry.Open()
        $digest = $sha.ComputeHash($stream)
        $hash = ([System.BitConverter]::ToString($digest)).Replace('-', '').ToLowerInvariant()
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
        $sha.Dispose()
    }
    return [ordered]@{
        path = "$ArchiveRelativePath#$normalized"
        sha256 = $hash
        bytes = [int64]$entry.Length
        kind = 'PptxPackageEntry'
        sourceArchive = $ArchiveRelativePath
    }
}

function Get-SourceSha256 {
    param($Source)
    $value = [string]$Source.sourceSha256
    if ($value -notmatch $script:Sha256Pattern) { throw "Inventory record has invalid sourceSha256: $($Source.recordId)" }
    return $value.ToLowerInvariant()
}

function Get-InventoryPath {
    param([object]$Value)
    $path = [string]$Value
    if ([string]::IsNullOrWhiteSpace($path)) { throw 'Inventory contains an empty path.' }
    return [System.IO.Path]::GetFullPath($path)
}

$inventoryPath = [System.IO.Path]::GetFullPath($CarrierInventoryJson)
$outputDir = [System.IO.Path]::GetFullPath($OutputDir)
$recognitionPath = if ([string]::IsNullOrWhiteSpace($RecognitionPath)) { '' } else { [System.IO.Path]::GetFullPath($RecognitionPath) }
if (-not (Test-Path -LiteralPath $inventoryPath -PathType Leaf)) { throw "Carrier inventory not found: $inventoryPath" }
if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
if ($recognitionPath -and -not (Test-Path -LiteralPath $recognitionPath -PathType Leaf)) { throw "Recognition evidence not found: $recognitionPath" }

$inventory = Get-Content -LiteralPath $inventoryPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$inventory.schemaVersion -ne 1) { throw "Unsupported carrier inventory schemaVersion: $($inventory.schemaVersion)" }
$inputPath = Get-InventoryPath -Value $inventory.input.path
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "Inventory source PPTX not found: $inputPath" }
$actualInputHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
$declaredInputHash = ([string]$inventory.input.sha256).ToLowerInvariant()
if ($declaredInputHash -notmatch $script:Sha256Pattern -or $actualInputHash -ne $declaredInputHash) {
    throw "Inventory source hash mismatch: declared=$declaredInputHash actual=$actualInputHash"
}

$inputRelative = Get-RelativeManifestPath -BaseDirectory $outputDir -TargetPath $inputPath
$inventoryEvidence = Get-FileEvidence -BaseDirectory $outputDir -Path $inventoryPath -Kind 'CarrierInventory'
$inputItem = Get-Item -LiteralPath $inputPath
$inputEvidence = [ordered]@{
    path = $inputRelative
    sha256 = $actualInputHash
    bytes = [int64]$inputItem.Length
    kind = 'PptxSource'
}

$zip = $null
$records = New-Object System.Collections.Generic.List[object]
$carrierCounts = @{}
$decisionCounts = @{}
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($inputPath)
    foreach ($inventoryRecord in @($inventory.records)) {
        $source = $inventoryRecord.source
        $carrier = [string]$source.carrier
        $decision = [string]$inventoryRecord.decision.status
        if ([string]::IsNullOrWhiteSpace($carrier)) { throw "Inventory record has empty carrier: $($inventoryRecord.recordId)" }
        if (-not $carrierCounts.ContainsKey($carrier)) { $carrierCounts[$carrier] = 0 }
        if (-not $decisionCounts.ContainsKey($decision)) { $decisionCounts[$decision] = 0 }
        $carrierCounts[$carrier]++
        $decisionCounts[$decision]++

        $packagePart = [string]$source.packagePart
        if ([string]::IsNullOrWhiteSpace($packagePart)) { $packagePart = 'unknown' }
        $sourceHash = Get-SourceSha256 -Source $source
        $recordFiles = New-Object System.Collections.Generic.List[object]
        $recordFiles.Add($inputEvidence) | Out-Null
        $packageEntries = New-Object System.Collections.Generic.List[object]
        foreach ($mediaPath in @($source.mediaPaths)) {
            if ([string]::IsNullOrWhiteSpace([string]$mediaPath)) { continue }
            $packageEntries.Add((Get-PackageEntryEvidence -Zip $zip -ArchiveRelativePath $inputRelative -EntryName ([string]$mediaPath))) | Out-Null
        }
        $recordPaths = New-Object System.Collections.Generic.List[string]
        $recordPaths.Add($inputRelative) | Out-Null
        $recordPaths.Add($inventoryEvidence.path) | Out-Null
        foreach ($packageEntry in $packageEntries) { $recordPaths.Add([string]$packageEntry.path) | Out-Null }
        if ($recognitionPath) { $recordPaths.Add((Get-RelativeManifestPath -BaseDirectory $outputDir -TargetPath $recognitionPath)) | Out-Null }

        $recordSource = [ordered]@{
            path = $inputRelative
            sha256 = $actualInputHash
            packagePart = $packagePart
            sourceSha256 = $sourceHash
            carrier = $carrier
            slide = [int]$source.slide
            shapeId = if ($null -eq $source.shapeId) { $null } else { [int]$source.shapeId }
            shapeName = [string]$source.shapeName
            bbox = $source.bbox
        }
        $recordEvidence = [ordered]@{
            paths = @($recordPaths.ToArray() | Select-Object -Unique)
            files = @($recordFiles.ToArray())
            packageEntries = @($packageEntries.ToArray())
            rollbackPath = $inputRelative
        }
        $records.Add([ordered]@{
            recordId = [string]$inventoryRecord.recordId
            source = $recordSource
            detection = $inventoryRecord.detection
            canonical = $inventoryRecord.canonical
            decision = $inventoryRecord.decision
            evidence = $recordEvidence
        }) | Out-Null
    }
} finally {
    if ($null -ne $zip) { $zip.Dispose() }
}

$recognitionEvidence = $null
if ($recognitionPath) { $recognitionEvidence = Get-FileEvidence -BaseDirectory $outputDir -Path $recognitionPath -Kind 'RecognitionResult' }
$evidenceHashInput = @($inputEvidence.sha256, $inventoryEvidence.sha256)
if ($null -ne $recognitionEvidence) { $evidenceHashInput += $recognitionEvidence.sha256 }
foreach ($record in $records) {
    $evidenceHashInput += [string]$record.recordId
    $evidenceHashInput += [string]$record.source.sourceSha256
    foreach ($entry in @($record.evidence.packageEntries)) { $evidenceHashInput += [string]$entry.sha256 }
}
$evidenceSetSha256 = Get-Sha256TextLocal -Text (($evidenceHashInput -join "`n"))

$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{
        mode = $Mode
        writeBackAllowed = $false
        canonicalSources = @('Whitelist', 'GoldSet')
        pathBase = 'manifest-directory'
        hashAlgorithm = 'SHA-256'
        note = 'This manifest is evidence-only. It never authorizes PPTX write-back.'
    }
    input = $inputEvidence
    inventory = $inventoryEvidence
    recognition = $recognitionEvidence
    counts = [ordered]@{
        recordCount = $records.Count
        byCarrier = [ordered]@{}
        byDecision = [ordered]@{}
    }
    evidenceSetSha256 = $evidenceSetSha256
    records = @($records.ToArray())
}
foreach ($key in @($carrierCounts.Keys | Sort-Object)) { $manifest.counts.byCarrier[$key] = [int]$carrierCounts[$key] }
foreach ($key in @($decisionCounts.Keys | Sort-Object)) { $manifest.counts.byDecision[$key] = [int]$decisionCounts[$key] }

$manifestPath = Join-Path $outputDir 'formula-evidence-manifest.json'
Write-Utf8BomText -Path $manifestPath -Text ($manifest | ConvertTo-Json -Depth 32)
Write-Output ("Formula evidence manifest done: {0}`nRecords: {1}; evidenceSetSha256: {2}" -f $manifestPath, $records.Count, $evidenceSetSha256)
