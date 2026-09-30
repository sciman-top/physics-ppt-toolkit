<#
.SYNOPSIS
  Resolve formula candidates against whitelist and exact source context.

.DESCRIPTION
  A candidate is a review artifact only. Context matching requires the exact
  carrier, slide, shape identifier, and source hash from a reviewed mapping.
  Ambiguous, absent, or incomplete context deterministically becomes
  ManualRequired and never authorizes a PPTX modification.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CandidateCsv,

    [Parameter(Mandatory = $true)]
    [string]$ContextMapCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$StyleConfigPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'config\physics-ppt-style.config.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Normalize-FormulaSourceText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace '\s+', '') -replace '[∙·•*]', '×' -replace '−', '-')
}
function Get-Sha256TextLocal {
    param([string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))).Replace('-', '').ToLowerInvariant()) }
    finally { $sha.Dispose() }
}

$CandidateCsv = [IO.Path]::GetFullPath($CandidateCsv)
$ContextMapCsv = [IO.Path]::GetFullPath($ContextMapCsv)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
$StyleConfigPath = [IO.Path]::GetFullPath($StyleConfigPath)
foreach ($path in @($CandidateCsv, $ContextMapCsv, $StyleConfigPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required input not found: $path" }
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$config = Get-Content -LiteralPath $StyleConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$whitelist = @($config.formulaWhitelist)
if ($whitelist.Count -eq 0) { throw 'Formula whitelist is empty.' }

$contextRows = @(Import-Csv -LiteralPath $ContextMapCsv -Encoding UTF8)
$contextByKey = @{}
foreach ($row in $contextRows) {
    foreach ($field in @('Carrier', 'Slide', 'ShapeId', 'SourceSha256', 'FormulaName', 'TargetUnicodeMath', 'TargetTex')) {
        if ([string]::IsNullOrWhiteSpace([string]$row.$field)) { throw "Context map row is missing $field." }
    }
    if ([string]$row.SourceSha256 -notmatch '^[A-Fa-f0-9]{64}$') { throw "Context map SourceSha256 is invalid: $($row.SourceSha256)" }
    $key = ('{0}|{1}|{2}|{3}' -f [string]$row.Carrier, [string]$row.Slide, [string]$row.ShapeId, ([string]$row.SourceSha256).ToLowerInvariant())
    if (-not $contextByKey.ContainsKey($key)) { $contextByKey[$key] = New-Object System.Collections.Generic.List[object] }
    $contextByKey[$key].Add($row) | Out-Null
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($candidate in @(Import-Csv -LiteralPath $CandidateCsv -Encoding UTF8)) {
    $candidateId = [string]$candidate.CandidateId
    if ([string]::IsNullOrWhiteSpace($candidateId)) { throw 'Candidate row is missing CandidateId.' }
    $rawText = Normalize-FormulaSourceText -Text ([string]$candidate.SourceText)
    $carrier = [string]$candidate.Carrier
    $slide = [string]$candidate.Slide
    $shapeId = [string]$candidate.ShapeId
    $sourceSha = ([string]$candidate.SourceSha256).ToLowerInvariant()
    $contextComplete = -not [string]::IsNullOrWhiteSpace($carrier) -and -not [string]::IsNullOrWhiteSpace($slide) -and -not [string]::IsNullOrWhiteSpace($shapeId) -and $sourceSha -match '^[a-f0-9]{64}$'
    $contextKey = if ($contextComplete) { "$carrier|$slide|$shapeId|$sourceSha" } else { '' }
    [object[]]$contextMatches = @()
    if ($contextComplete -and $contextByKey.ContainsKey($contextKey)) {
        [object[]]$contextMatches = @($contextByKey[$contextKey].ToArray())
    }
    [object[]]$whitelistMatches = @($whitelist | Where-Object { $rawText -match ([string]$_.sourcePattern) })
    $status = 'ManualRequired'
    $canonicalSource = 'None'
    $formulaName = ''
    $unicodeMath = ''
    $tex = ''
    $reason = ''
    if (-not $contextComplete) {
        $reason = 'Source context is incomplete; OCR/text alone cannot choose a canonical formula.'
    } elseif ($contextMatches.Count -gt 1) {
        $reason = 'More than one reviewed canonical mapping matches the exact source context.'
    } elseif ($contextMatches.Count -eq 1) {
        $match = $contextMatches[0]
        if ($whitelistMatches.Count -gt 0 -and @($whitelistMatches | Where-Object { [string]$_.targetUnicodeMath -eq [string]$match.TargetUnicodeMath -and [string]$_.targetTex -eq [string]$match.TargetTex }).Count -eq 0) {
            $reason = 'Reviewed context canonical content conflicts with current whitelist; manual review is required.'
        } else {
            $status = 'CandidateOnly'
            $canonicalSource = 'GoldSet'
            $formulaName = [string]$match.FormulaName
            $unicodeMath = [string]$match.TargetUnicodeMath
            $tex = [string]$match.TargetTex
            $reason = 'Exact reviewed source context matched one canonical GoldSet record.'
        }
    } elseif ($whitelistMatches.Count -eq 1) {
        $match = $whitelistMatches[0]
        $status = 'CandidateOnly'
        $canonicalSource = 'Whitelist'
        $formulaName = [string]$match.name
        $unicodeMath = [string]$match.targetUnicodeMath
        $tex = [string]$match.targetTex
        $reason = 'One current whitelist rule matched; no reviewed source-context mapping was required for this text candidate.'
    } elseif ($whitelistMatches.Count -eq 0) {
        $reason = 'No whitelist or reviewed source-context mapping matched the candidate.'
    } else {
        $reason = 'More than one whitelist rule matches the candidate text.'
    }
    $results.Add([pscustomobject]@{
        CandidateId = $candidateId
        SourceText = $rawText
        Carrier = $carrier
        Slide = $slide
        ShapeId = $shapeId
        SourceSha256 = $sourceSha
        ContextKey = $contextKey
        ContextMatchCount = $contextMatches.Count
        WhitelistMatchCount = $whitelistMatches.Count
        DecisionStatus = $status
        TargetCarrier = 'None'
        WriteBackAllowed = $false
        CanonicalSource = $canonicalSource
        FormulaName = $formulaName
        TargetUnicodeMath = $unicodeMath
        TargetTex = $tex
        Reason = $reason
    }) | Out-Null
}
$csvPath = Join-Path $OutputDir 'formula-context-resolution.csv'
$jsonPath = Join-Path $OutputDir 'formula-context-resolution.json'
Write-Utf8BomCsv -InputObject $results.ToArray() -Path $csvPath
$evidenceRows = @($results.ToArray() | ForEach-Object { "$($_.CandidateId)|$($_.SourceSha256)|$($_.DecisionStatus)|$($_.CanonicalSource)|$($_.TargetUnicodeMath)" } | Sort-Object)
$evidenceSetSha256 = Get-Sha256TextLocal -Text ($evidenceRows -join "`n")
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ mode = 'CandidateOnly'; writeBackAllowed = $false; sourceContextRequiredForGoldSet = $true }
    inputs = [ordered]@{ candidateCsv = $CandidateCsv; candidateCsvSha256 = (Get-FileHash -LiteralPath $CandidateCsv -Algorithm SHA256).Hash.ToLowerInvariant(); contextMapCsv = $ContextMapCsv; contextMapCsvSha256 = (Get-FileHash -LiteralPath $ContextMapCsv -Algorithm SHA256).Hash.ToLowerInvariant(); styleConfig = $StyleConfigPath }
    counts = [ordered]@{ total = $results.Count; candidateOnly = @($results | Where-Object { $_.DecisionStatus -eq 'CandidateOnly' }).Count; manualRequired = @($results | Where-Object { $_.DecisionStatus -eq 'ManualRequired' }).Count }
    evidenceSetSha256 = $evidenceSetSha256
    results = @($results.ToArray())
}
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
Write-Host "Formula context resolution complete: $OutputDir"
Write-Host "CandidateOnly: $($manifest.counts.candidateOnly); ManualRequired: $($manifest.counts.manualRequired)"
