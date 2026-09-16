<#
.SYNOPSIS
  Convert an approved MathType/OLE mapping into review-only FormulaIR artifacts.

.DESCRIPTION
  Reuses Export-FormulaOmmlCandidates.ps1 as the single canonical parser and
  renderer. The mapping and its passed manifest are the only source of truth;
  this tool never edits a PPTX and never changes a decision to Converted.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FormulaOleMappingCsv,

    [Parameter(Mandatory = $true)]
    [string]$MappingManifestJson,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [ValidateRange(1, 1000)]
    [int]$MaxItems = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$FormulaOleMappingCsv = [System.IO.Path]::GetFullPath($FormulaOleMappingCsv)
$MappingManifestJson = [System.IO.Path]::GetFullPath($MappingManifestJson)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
foreach ($path in @($FormulaOleMappingCsv, $MappingManifestJson)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required input not found: $path" }
}
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$mapping = @(Import-Csv -LiteralPath $FormulaOleMappingCsv -Encoding UTF8 | Select-Object -First $MaxItems)
$mappingManifest = Get-Content -LiteralPath $MappingManifestJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$mappingManifest.status -ne 'Passed') {
    throw "OLE mapping manifest is not Passed: $([string]$mappingManifest.status)"
}
if ($null -eq $mappingManifest.input -or [string]::IsNullOrWhiteSpace([string]$mappingManifest.input.path)) {
    throw 'OLE mapping manifest has no input.path.'
}
$goldRows = @($mappingManifest.rows)
if ($goldRows.Count -eq 0) { throw 'OLE mapping manifest contains no approved rows.' }

$goldByKey = @{}
foreach ($gold in $goldRows) {
    $key = "$( [int]$gold.Slide)|$( [string]$gold.ShapeIds)|$( [string]$gold.OleIndex)"
    $goldByKey[$key] = $gold
}

$reviewRows = New-Object System.Collections.Generic.List[object]
foreach ($row in $mapping) {
    foreach ($required in @('Slide', 'OleIndex', 'WhitelistName', 'SizePt')) {
        if ($null -eq $row.PSObject.Properties[$required] -or [string]::IsNullOrWhiteSpace([string]$row.$required)) {
            throw "OLE mapping row is missing $required."
        }
    }
    $key = "$( [int]$row.Slide)|$( [string]$row.ShapeIds)|$( [string]$row.OleIndex)"
    if (-not $goldByKey.ContainsKey($key)) { throw "OLE mapping row has no approved GoldSet row: $key" }
    $gold = $goldByKey[$key]
    $unicodeMath = [string]$gold.TargetUnicodeMath
    $tex = [string]$gold.TargetTex
    if ([string]::IsNullOrWhiteSpace($unicodeMath) -or [string]::IsNullOrWhiteSpace($tex)) {
        throw "Approved OLE mapping row has empty canonical source: $key"
    }
    $candidate = 'name={0}; targetUnicodeMath={1}; targetTex={2}; note=approved OLE GoldSet mapping' -f ([string]$row.WhitelistName), $unicodeMath, $tex
    $reviewRows.Add([pscustomobject]@{
        File = [System.IO.Path]::GetFileName([string]$mappingManifest.input.path)
        FilePath = [string]$mappingManifest.input.path
        FileRelativePath = ''
        Slide = [int]$row.Slide
        Shape = [string]$row.ShapeIds
        FormulaText = [string]$gold.SourceFormulaText
        CandidateClass = 'ApprovedMathTypeOleGoldSet'
        WhitelistCandidate = $candidate
        ConversionStatus = 'Approved'
        StyleStatus = 'Reviewed'
        SuggestedAction = 'ReviewWhitelistConversion'
    }) | Out-Null
}

$reviewCsv = Join-Path $OutputDir 'formula-ole-ir-review.csv'
Write-Utf8BomCsv -InputObject $reviewRows.ToArray() -Path $reviewCsv

$candidateTool = Join-Path $PSScriptRoot 'Export-FormulaOmmlCandidates.ps1'
& pwsh -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $candidateTool -FormulaReviewCsv $reviewCsv -OutputDir $OutputDir -MaxItems $MaxItems
if ($LASTEXITCODE -ne 0) { throw "FormulaIR candidate exporter failed with exit code $LASTEXITCODE." }

$candidateManifestPath = Join-Path $OutputDir 'formula-omml-candidates-manifest.json'
$candidateManifest = Get-Content -LiteralPath $candidateManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$summary = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    input = [ordered]@{ mappingCsv = $FormulaOleMappingCsv; mappingManifest = $MappingManifestJson; sourcePptx = [string]$mappingManifest.input.path }
    policy = [ordered]@{ status = 'CandidateOnly'; writeBackAllowed = $false; sourceOfTruth = 'Approved OLE GoldSet mapping' }
    counts = [ordered]@{ mappingRows = $mapping.Count; reviewRows = $reviewRows.Count; formulaIrResolved = [int]$candidateManifest.formulaIrResolvedCount; generated = [int]$candidateManifest.generatedCount; failed = [int]$candidateManifest.failedCount }
    artifacts = [ordered]@{ reviewCsv = $reviewCsv; candidatesCsv = [string]$candidateManifest.csv; candidatesJson = [string]$candidateManifest.json; formulaIrDir = [string]$candidateManifest.formulaIrDir }
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'formula-ole-ir-manifest.json') -Encoding UTF8
Write-Host "FormulaIR OLE mapping export done: $OutputDir"
Write-Host "Resolved: $($summary.counts.formulaIrResolved) / Rows: $($summary.counts.reviewRows)"
