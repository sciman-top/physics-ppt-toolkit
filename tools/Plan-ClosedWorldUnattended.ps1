<#
.SYNOPSIS
  Build a fail-closed ClosedWorldUnattended formula plan without editing a PPTX.

.DESCRIPTION
  The plan consumes a hash-bound carrier inventory, reviewed GoldSet evidence,
  context resolution, and recognition evaluation. It can never authorize a
  write. Missing gates, unavailable recognition, image carriers, and any
  ambiguity become a preserved or manual outcome. A circuit breaker remains
  closed until a separate explicit write-back pilot proves all required gates.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CarrierInventoryJson,
    [Parameter(Mandatory = $true)][string]$GoldSetManifestJson,
    [Parameter(Mandatory = $true)][string]$ContextResolutionJson,
    [Parameter(Mandatory = $true)][string]$RecognitionEvaluationJson,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [string]$ReferencePptxPath = ''
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
function Get-RecordKey { param($Source) return ('{0}|{1}|{2}|{3}' -f [string]$Source.carrier, [string]$Source.slide, [string]$Source.shapeId, ([string]$Source.sourceSha256).ToLowerInvariant()) }

foreach ($name in @('CarrierInventoryJson', 'GoldSetManifestJson', 'ContextResolutionJson', 'RecognitionEvaluationJson')) {
    $value = [IO.Path]::GetFullPath((Get-Variable -Name $name -ValueOnly))
    Set-Variable -Name $name -Value $value
    if (-not (Test-Path -LiteralPath $value -PathType Leaf)) { throw "Required input not found: $value" }
}
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$inventory = Get-Content -LiteralPath $CarrierInventoryJson -Raw -Encoding UTF8 | ConvertFrom-Json
$goldSet = Get-Content -LiteralPath $GoldSetManifestJson -Raw -Encoding UTF8 | ConvertFrom-Json
$context = Get-Content -LiteralPath $ContextResolutionJson -Raw -Encoding UTF8 | ConvertFrom-Json
$recognition = Get-Content -LiteralPath $RecognitionEvaluationJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$inventory.schemaVersion -ne 1 -or [int]$goldSet.schemaVersion -ne 1 -or [int]$context.schemaVersion -ne 1 -or [int]$recognition.schemaVersion -ne 1) { throw 'Unsupported input schema version.' }
if ([bool]$goldSet.policy.writeBackAllowed -or [bool]$context.policy.writeBackAllowed -or [bool]$recognition.policy.writeBackAllowed) { throw 'Closed-world planning inputs must keep writeBackAllowed=false.' }
$sourceHash = ([string]$inventory.input.sha256).ToLowerInvariant()
if ($sourceHash -notmatch '^[a-f0-9]{64}$') { throw 'Carrier inventory input hash is invalid.' }

# Every evidence layer must describe the same source package.  Without this
# check, a stale GoldSet could be accepted while the planner merely falls
# back to ManualRequired, hiding an evidence-binding defect from callers.
$goldSourceHash = ([string]$goldSet.input.sourcePptxSha256).ToLowerInvariant()
if ($goldSourceHash -notmatch '^[a-f0-9]{64}$' -or $goldSourceHash -ne $sourceHash) {
    throw "GoldSet source hash does not match carrier inventory: gold=$goldSourceHash inventory=$sourceHash"
}
$recognitionGoldSetHash = ([string]$recognition.inputs.goldSetManifestSha256).ToLowerInvariant()
$actualGoldSetManifestHash = Get-FileSha256Local -Path $GoldSetManifestJson
if ($recognitionGoldSetHash -notmatch '^[a-f0-9]{64}$' -or $recognitionGoldSetHash -ne $actualGoldSetManifestHash) {
    throw "Recognition evaluation is not bound to the supplied GoldSet manifest: evaluation=$recognitionGoldSetHash manifest=$actualGoldSetManifestHash"
}

$inventoryRecordKeys = @{}
foreach ($inventoryRecord in @($inventory.records)) {
    $inventoryRecordKeys[(Get-RecordKey -Source $inventoryRecord.source)] = $true
}
foreach ($contextResult in @($context.results)) {
    $contextKey = ('{0}|{1}|{2}|{3}' -f [string]$contextResult.Carrier, [string]$contextResult.Slide, [string]$contextResult.ShapeId, ([string]$contextResult.SourceSha256).ToLowerInvariant())
    if (-not $inventoryRecordKeys.ContainsKey($contextKey)) {
        throw "Context resolution contains a source record absent from the supplied inventory: $contextKey"
    }
}
if (-not [string]::IsNullOrWhiteSpace($ReferencePptxPath)) {
    $ReferencePptxPath = [IO.Path]::GetFullPath($ReferencePptxPath)
    if (-not (Test-Path -LiteralPath $ReferencePptxPath -PathType Leaf)) { throw "Reference PPTX not found: $ReferencePptxPath" }
    if ((Get-FileSha256Local -Path $ReferencePptxPath) -ne $sourceHash) { throw 'Reference PPTX hash does not match the carrier inventory.' }
}

$contextByKey = @{}
foreach ($row in @($context.results)) {
    $key = ('{0}|{1}|{2}|{3}' -f [string]$row.Carrier, [string]$row.Slide, [string]$row.ShapeId, ([string]$row.SourceSha256).ToLowerInvariant())
    if (-not $contextByKey.ContainsKey($key)) { $contextByKey[$key] = @() }
    $contextByKey[$key] += $row
}
$modelUnavailable = [string]$recognition.policy.releaseDecision -like 'CandidateOnly*'
$falseAcceptanceAdapters = New-Object System.Collections.Generic.List[string]
foreach ($adapterEvaluation in @($recognition.adapters)) {
    $falseAcceptCount = 0
    $falseAcceptText = [string]$adapterEvaluation.falseAcceptCount
    if (-not [string]::IsNullOrWhiteSpace($falseAcceptText) -and [int]::TryParse($falseAcceptText, [ref]$falseAcceptCount) -and $falseAcceptCount -gt 0) {
        $falseAcceptanceAdapters.Add([string]$adapterEvaluation.adapter) | Out-Null
        continue
    }
    if ([string]$adapterEvaluation.decision -match '(?i)false.?accept|unexplained.?accept') {
        $falseAcceptanceAdapters.Add([string]$adapterEvaluation.adapter) | Out-Null
    }
}
$falseAcceptanceDetected = $falseAcceptanceAdapters.Count -gt 0
$results = New-Object System.Collections.Generic.List[object]
foreach ($record in @($inventory.records)) {
    $source = $record.source
    $carrier = [string]$source.carrier
    $decision = 'ManualRequired'
    $target = 'None'
    $reason = 'No closed-world canonical source and gate set is available for this carrier.'
    $recordContextKey = Get-RecordKey -Source $source
    $contextRows = New-Object System.Collections.Generic.List[object]
    if ($contextByKey.ContainsKey($recordContextKey)) {
        foreach ($contextRow in @($contextByKey[$recordContextKey])) {
            if ($null -ne $contextRow) { $contextRows.Add($contextRow) | Out-Null }
        }
    }
    switch ($carrier) {
        'OfficeMath' { $decision = 'NativeKept'; $target = 'OfficeMath'; $reason = 'Existing OfficeMath is preserved and not rebuilt.' }
        'MathTypeOle' {
            $decision = 'OriginalKept'; $target = 'Original'
            if ($contextRows.Count -eq 1 -and [string]$contextRows[0].DecisionStatus -eq 'CandidateOnly') {
                $reason = 'Exact GoldSet context is available, but MathType migration requires a separate ExplicitMigration batch and host/visual gates.'
            } else { $reason = 'MathType OLE remains original because no exact reviewed context is available for an explicit migration batch.' }
        }
        'FormulaImage' { $decision = 'ManualRequired'; $reason = 'Formula image has no approved isolated-image canonical FormulaIR; recognition is not a content source.' }
        'MixedImage' { $decision = 'ManualRequired'; $reason = 'Mixed image cannot be safely isolated into a formula carrier.' }
        'TextFormula' { $decision = 'ManualRequired'; $reason = 'Text formula is outside this no-write plan until exact canonical/context and writer gates are supplied.' }
        'GroupFormula' { $decision = 'Skipped'; $reason = 'Grouped carrier is protected by the no-ungroup invariant.' }
        'Unknown' { $decision = 'ManualRequired'; $reason = 'Unknown carrier cannot enter unattended processing.' }
        default { throw "Unsupported inventory carrier: $carrier" }
    }
    $results.Add([ordered]@{
        recordId = [string]$record.recordId
        source = [ordered]@{ carrier = $carrier; slide = [int]$source.slide; shapeId = $source.shapeId; sourceSha256 = [string]$source.sourceSha256 }
        context = [ordered]@{ matchCount = $contextRows.Count; candidateOnly = @($contextRows | Where-Object { [string]$_.DecisionStatus -eq 'CandidateOnly' }).Count }
        decision = [ordered]@{ mode = 'ClosedWorldUnattended'; status = $decision; targetCarrier = $target; writeBackAllowed = $false; reason = $reason }
    }) | Out-Null
}
$statusNames = @('NativeKept', 'Converted', 'FallbackSvg', 'OriginalKept', 'ManualRequired', 'Skipped', 'Failed')
$counts = [ordered]@{}
foreach ($name in $statusNames) { $counts[$name] = @($results | Where-Object { [string]$_.decision.status -eq $name }).Count }
$circuitBreaker = [ordered]@{
    state = 'Closed'
    triggered = $true
    reason = if ($falseAcceptanceDetected) { "Unexplained false acceptance detected in recognition evaluation: $($falseAcceptanceAdapters -join ', ')." } elseif ($modelUnavailable) { 'No independently benchmarked specialized recognition adapter is available, and this evidence set has no approved isolated FormulaImage positive case.' } else { 'A write-back pilot and independent host/visual gates are still required.' }
    recoveryCondition = 'An explicit pilot must provide approved isolated-image canonical FormulaIR, a non-candidate recognition evaluation, structural validation, PowerPoint host export, and visual acceptance without unexplained false acceptance.'
}
$inputEvidence = New-Object System.Collections.Generic.List[object]
foreach ($inputPath in @($CarrierInventoryJson, $GoldSetManifestJson, $ContextResolutionJson, $RecognitionEvaluationJson)) {
    $inputEvidence.Add([ordered]@{ path = $inputPath; sha256 = Get-FileSha256Local -Path $inputPath }) | Out-Null
}
$evidenceRows = @($inputEvidence | ForEach-Object { "$($_.path)|$($_.sha256)" }) + @($results | ForEach-Object { "$($_.recordId)|$($_.decision.status)|$($_.decision.targetCarrier)|$($_.source.sourceSha256)" })
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ requestedMode = 'ClosedWorldUnattended'; writeBackAllowed = $false; planOnly = $true; sourcePptxSha256 = $sourceHash }
    referencePptx = if ($ReferencePptxPath) { [ordered]@{ path = $ReferencePptxPath; sha256 = Get-FileSha256Local -Path $ReferencePptxPath } } else { $null }
    inputs = @($inputEvidence.ToArray())
    circuitBreaker = $circuitBreaker
    counts = $counts
    evidenceSetSha256 = Get-TextSha256Local -Text (($evidenceRows | Sort-Object) -join "`n")
    results = @($results.ToArray())
}
$csvRows = @($results | ForEach-Object { [pscustomobject]@{ RecordId = $_.recordId; Carrier = $_.source.carrier; Slide = $_.source.slide; ShapeId = $_.source.shapeId; DecisionStatus = $_.decision.status; TargetCarrier = $_.decision.targetCarrier; WriteBackAllowed = $_.decision.writeBackAllowed; ContextMatchCount = $_.context.matchCount; Reason = $_.decision.reason } })
Write-Utf8BomCsv -InputObject $csvRows -Path (Join-Path $OutputDir 'closed-world-unattended-plan.csv')
$manifest | ConvertTo-Json -Depth 24 | Set-Content -LiteralPath (Join-Path $OutputDir 'closed-world-unattended-plan.json') -Encoding UTF8
Write-Host "Closed-world unattended plan complete: $OutputDir"
Write-Host "Circuit breaker: $($circuitBreaker.state); Converted: $($counts.Converted); ManualRequired: $($counts.ManualRequired)"
