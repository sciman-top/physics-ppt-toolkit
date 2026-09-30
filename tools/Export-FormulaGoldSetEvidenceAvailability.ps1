<#
.SYNOPSIS
  Verify and locally rebind GoldSet evidence by filename and SHA-256.

.DESCRIPTION
  Historical GoldSet manifests may retain paths from an archived run. This
  read-only tool searches a supplied evidence root for a file with the same
  basename and expected hash, then writes a derived availability receipt. It
  never changes the historical manifest, source PPTX, or write-back policy.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$GoldSetManifestPath,
    [Parameter(Mandatory = $true)][string]$EvidenceRoot,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-Sha256Local {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

$GoldSetManifestPath = [IO.Path]::GetFullPath($GoldSetManifestPath)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
foreach ($path in @($GoldSetManifestPath, $EvidenceRoot)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required evidence input not found: $path" }
}
if (-not (Test-Path -LiteralPath $GoldSetManifestPath -PathType Leaf)) { throw "GoldSet manifest is not a file: $GoldSetManifestPath" }
if (-not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) { throw "EvidenceRoot is not a directory: $EvidenceRoot" }

$manifestHash = Get-Sha256Local -Path $GoldSetManifestPath
$manifest = Get-Content -LiteralPath $GoldSetManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$records = New-Object System.Collections.Generic.List[object]
$available = 0
$missing = 0
$ambiguous = 0

foreach ($record in @($manifest.records)) {
    $expected = ([string]$record.evidence.sha256).ToLowerInvariant()
    $leaf = Split-Path -Leaf ([string]$record.evidence.path)
    if ([string]::IsNullOrWhiteSpace($leaf)) { throw "GoldSet record has no evidence filename: $($record.goldSetId)" }
    $candidates = @(Get-ChildItem -LiteralPath $EvidenceRoot -Recurse -File -Filter $leaf -ErrorAction SilentlyContinue |
        Where-Object { (Get-Sha256Local -Path $_.FullName) -eq $expected })
    $status = 'Missing'
    $resolved = $null
    if ($candidates.Count -eq 1) { $status = 'Available'; $resolved = $candidates[0].FullName; $available++ }
    elseif ($candidates.Count -gt 1) { $status = 'Ambiguous'; $ambiguous++ }
    else { $missing++ }
    $records.Add([ordered]@{
        goldSetId = [string]$record.goldSetId
        groundTruthClass = [string]$record.groundTruth.class
        expectedSha256 = $expected
        originalPath = [string]$record.evidence.path
        status = $status
        resolvedPath = $resolved
        candidateCount = $candidates.Count
    }) | Out-Null
}

$receipt = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ readOnly = $true; writeBackAllowed = $false }
    sourceManifest = [ordered]@{ path = $GoldSetManifestPath; sha256 = $manifestHash }
    counts = [ordered]@{ total = $records.Count; available = $available; missing = $missing; ambiguous = $ambiguous }
    records = @($records.ToArray())
}
$parent = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
$receipt | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Host "Formula GoldSet evidence availability written: $OutputPath"
