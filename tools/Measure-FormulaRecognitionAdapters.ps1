<#
.SYNOPSIS
  Produce a hash-bound FormulaIR recognition-adapter evaluation report.

.DESCRIPTION
  Executes only configured adapter transport boundaries against human-reviewed
  GoldSet evidence. A runner result is never canonical content and can never
  authorize a PPTX write-back. Missing adapters remain Unavailable and the
  release decision remains CandidateOnly.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$GoldSetManifestJson,

    [Parameter(Mandatory = $true)]
    [string]$AdapterConfigJson,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [ValidateRange(1, 1000)]
    [int]$MaxSamples = 1000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Resolve-ManifestPath {
    param([string]$ManifestDirectory, [string]$Path)
    if ([IO.Path]::IsPathRooted($Path)) { throw "GoldSet evidence path must be manifest-relative: $Path" }
    $resolved = [IO.Path]::GetFullPath((Join-Path $ManifestDirectory ($Path -replace '/', '\\')))
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) { throw "GoldSet evidence path is missing: $resolved" }
    return $resolved
}
function Get-Sha256FileLocal { param([string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Write-JsonUtf8 { param([string]$Path, $Value) $Value | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $Path -Encoding UTF8 }

$GoldSetManifestJson = [IO.Path]::GetFullPath($GoldSetManifestJson)
$AdapterConfigJson = [IO.Path]::GetFullPath($AdapterConfigJson)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
foreach ($input in @($GoldSetManifestJson, $AdapterConfigJson)) {
    if (-not (Test-Path -LiteralPath $input -PathType Leaf)) { throw "Required input not found: $input" }
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$goldSet = Get-Content -LiteralPath $GoldSetManifestJson -Raw -Encoding UTF8 | ConvertFrom-Json
$config = Get-Content -LiteralPath $AdapterConfigJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$goldSet.schemaVersion -ne 1 -or [int]$config.schemaVersion -ne 1) { throw 'Unsupported GoldSet or adapter config schema version.' }
if ([bool]$goldSet.policy.writeBackAllowed -or [bool]$config.writeBackAllowed) { throw 'Recognition evaluation inputs must keep writeBackAllowed=false.' }

$adapterTool = Join-Path $PSScriptRoot 'Run-FormulaRecognitionAdapter.ps1'
$goldSetDirectory = Split-Path -Parent $GoldSetManifestJson
$records = @($goldSet.records | Select-Object -First $MaxSamples)
$resultRoot = Join-Path $OutputDir 'results'
New-Item -ItemType Directory -Path $resultRoot -Force | Out-Null
$runs = New-Object System.Collections.Generic.List[object]
foreach ($adapter in @($config.adapters)) {
    $adapterId = [string]$adapter.id
    if ([string]::IsNullOrWhiteSpace($adapterId)) { throw 'Adapter config contains an empty id.' }
    $adapterDirectory = Join-Path $resultRoot $adapterId
    New-Item -ItemType Directory -Path $adapterDirectory -Force | Out-Null
    foreach ($record in $records) {
        $imagePath = Resolve-ManifestPath -ManifestDirectory $goldSetDirectory -Path ([string]$record.evidence.path)
        $resultPath = Join-Path $adapterDirectory ([string]$record.goldSetId + '.json')
        $expectedInputHash = Get-Sha256FileLocal -Path $imagePath
        $useExisting = $false
        if (Test-Path -LiteralPath $resultPath -PathType Leaf) {
            try {
                $existing = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
                $useExisting = ([string]$existing.adapter -eq $adapterId -and [string]$existing.input.sha256 -eq $expectedInputHash -and -not [bool]$existing.writeBackAllowed)
            } catch { $useExisting = $false }
        }
        if (-not $useExisting) {
            & $adapterTool -Adapter $adapterId -ImagePath $imagePath -OutputPath $resultPath -TimeoutSeconds 10
            if (-not $?) { throw "Adapter transport failed for $adapterId/$($record.goldSetId)." }
        }
        $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([bool]$result.writeBackAllowed) { throw "Adapter result illegally enables write-back: $adapterId/$($record.goldSetId)" }
        if ([string]$result.input.sha256 -ne $expectedInputHash) { throw "Adapter result input hash mismatch: $adapterId/$($record.goldSetId)" }
        $runs.Add([pscustomobject]@{
            Adapter = $adapterId
            GoldSetId = [string]$record.goldSetId
            GroundTruthClass = [string]$record.groundTruth.class
            Status = [string]$result.status
            CandidateCount = @($result.candidates).Count
            ResultPath = $resultPath
            ResultSha256 = Get-Sha256FileLocal -Path $resultPath
        }) | Out-Null
    }
}

$adapterSummary = @($config.adapters | ForEach-Object {
    $adapterId = [string]$_.id
    $adapterRuns = @($runs | Where-Object { $_.Adapter -eq $adapterId })
    [ordered]@{
        adapter = $adapterId
        configuredStatus = [string]$_.status
        samples = $adapterRuns.Count
        passed = @($adapterRuns | Where-Object { $_.Status -eq 'Passed' }).Count
        unavailable = @($adapterRuns | Where-Object { $_.Status -eq 'Unavailable' }).Count
        failed = @($adapterRuns | Where-Object { $_.Status -in @('Failed', 'InvalidOutput', 'Timeout') }).Count
        exactFormulaIrMatch = $null
        falseAcceptCount = $null
        chineseSubscriptErrorCount = $null
        unitErrorCount = $null
        decision = if (@($adapterRuns | Where-Object { $_.Status -eq 'Passed' }).Count -eq 0) { 'UnavailableNoBenchmark' } else { 'BenchmarkRequired' }
    }
})
$releaseDecision = if (@($adapterSummary | Where-Object { $_.decision -eq 'BenchmarkRequired' }).Count -gt 0) { 'CandidateOnlyPendingGoldSetMetrics' } else { 'CandidateOnlyNoAvailableAdapter' }
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    inputs = [ordered]@{
        goldSetManifest = $GoldSetManifestJson
        goldSetManifestSha256 = Get-Sha256FileLocal -Path $GoldSetManifestJson
        adapterConfig = $AdapterConfigJson
        adapterConfigSha256 = Get-Sha256FileLocal -Path $AdapterConfigJson
    }
    policy = [ordered]@{ writeBackAllowed = $false; releaseDecision = $releaseDecision; selectedAdapter = 'None'; reason = 'No independently benchmarked model may become a content source or write-back authority.' }
    sampleCounts = [ordered]@{ total = $records.Count; formula = @($records | Where-Object { $_.groundTruth.class -eq 'Formula' }).Count; nonFormula = @($records | Where-Object { $_.groundTruth.class -eq 'NonFormula' }).Count; mixedNonIsolatable = @($records | Where-Object { $_.groundTruth.class -eq 'MixedNonIsolatable' }).Count }
    adapters = $adapterSummary
    runCount = $runs.Count
    runsCsv = Join-Path $OutputDir 'formula-recognition-evaluation-runs.csv'
}
Write-JsonUtf8 -Path (Join-Path $OutputDir 'formula-recognition-evaluation-manifest.json') -Value $manifest
Write-Utf8BomCsv -InputObject $runs.ToArray() -Path $manifest.runsCsv
Write-Host "Formula recognition evaluation complete: $OutputDir"
Write-Host "Release decision: $releaseDecision; runs: $($runs.Count)"
