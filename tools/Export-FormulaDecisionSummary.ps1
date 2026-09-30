<#
.SYNOPSIS
  Aggregate every formula carrier into a hash-bound, no-write decision summary.

.DESCRIPTION
  Joins a plan-only closed-world manifest to the full carrier inventory and
  optional evidence receipts. It rejects missing or duplicate decisions and
  always reports every terminal state, including zero counts.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CarrierInventoryJson,
    [Parameter(Mandatory = $true)][string]$ClosedWorldPlanJson,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [string[]]$EvidencePaths = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-FileSha256Local { param([string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-TextSha256Local { param([string]$Text) $sha = [Security.Cryptography.SHA256]::Create(); try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))).Replace('-', '').ToLowerInvariant()) } finally { $sha.Dispose() } }

$CarrierInventoryJson = [IO.Path]::GetFullPath($CarrierInventoryJson)
$ClosedWorldPlanJson = [IO.Path]::GetFullPath($ClosedWorldPlanJson)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
foreach ($path in @($CarrierInventoryJson, $ClosedWorldPlanJson)) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required input not found: $path" } }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$inventory = Get-Content -LiteralPath $CarrierInventoryJson -Raw -Encoding UTF8 | ConvertFrom-Json
$plan = Get-Content -LiteralPath $ClosedWorldPlanJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$inventory.schemaVersion -ne 1 -or [int]$plan.schemaVersion -ne 1) { throw 'Unsupported inventory or plan schema version.' }
if ([bool]$plan.policy.writeBackAllowed -or -not [bool]$plan.policy.planOnly) { throw 'Decision summary accepts only a plan-only, no-write closed-world manifest.' }
$planByRecord = @{}
foreach ($row in @($plan.results)) {
    $id = [string]$row.recordId
    if ([string]::IsNullOrWhiteSpace($id) -or $planByRecord.ContainsKey($id)) { throw "Plan has an empty or duplicate record decision: $id" }
    if ([bool]$row.decision.writeBackAllowed) { throw "Plan illegally enables write-back: $id" }
    $planByRecord[$id] = $row
}
$terminalStates = @('NativeKept', 'Converted', 'FallbackSvg', 'OriginalKept', 'ManualRequired', 'Skipped', 'Failed')
$rows = New-Object System.Collections.Generic.List[object]
foreach ($record in @($inventory.records)) {
    $id = [string]$record.recordId
    if (-not $planByRecord.ContainsKey($id)) { throw "Plan is missing a decision for inventory record: $id" }
    $decision = $planByRecord[$id].decision
    if ([string]$decision.status -notin $terminalStates) { throw "Plan has an invalid terminal status for ${id}: $($decision.status)" }
    $rows.Add([ordered]@{ recordId = $id; carrier = [string]$record.source.carrier; slide = [int]$record.source.slide; shapeId = $record.source.shapeId; sourceSha256 = [string]$record.source.sourceSha256; decision = $decision }) | Out-Null
}
if ($planByRecord.Count -ne $rows.Count) { throw 'Plan contains decisions that do not belong to the carrier inventory.' }
$evidence = New-Object System.Collections.Generic.List[object]
foreach ($path in @($CarrierInventoryJson, $ClosedWorldPlanJson) + @($EvidencePaths)) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    $full = [IO.Path]::GetFullPath($path)
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "Evidence path not found: $full" }
    $evidence.Add([ordered]@{ path = $full; sha256 = Get-FileSha256Local -Path $full }) | Out-Null
}
$counts = [ordered]@{}
foreach ($state in $terminalStates) { $counts[$state] = @($rows | Where-Object { [string]$_.decision.status -eq $state }).Count }
$evidenceLines = @($evidence | ForEach-Object { "$($_.path)|$($_.sha256)" }) + @($rows | ForEach-Object { "$($_.recordId)|$($_.sourceSha256)|$($_.decision.status)|$($_.decision.targetCarrier)" })
$summary = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ mode = 'ClosedWorldUnattended'; planOnly = $true; writeBackAllowed = $false; decisionCoverageRequired = $true }
    input = [ordered]@{ sourcePptxSha256 = [string]$inventory.input.sha256; carrierInventory = $CarrierInventoryJson; closedWorldPlan = $ClosedWorldPlanJson }
    counts = $counts
    evidence = @($evidence.ToArray())
    evidenceSetSha256 = Get-TextSha256Local -Text (($evidenceLines | Sort-Object) -join "`n")
    circuitBreaker = $plan.circuitBreaker
    records = @($rows.ToArray())
}
$csv = @($rows | ForEach-Object { [pscustomobject]@{ RecordId = $_.recordId; Carrier = $_.carrier; Slide = $_.slide; ShapeId = $_.shapeId; DecisionStatus = $_.decision.status; TargetCarrier = $_.decision.targetCarrier; WriteBackAllowed = $_.decision.writeBackAllowed; Reason = $_.decision.reason } })
Write-Utf8BomCsv -InputObject $csv -Path (Join-Path $OutputDir 'formula-decision-summary.csv')
$summary | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $OutputDir 'formula-decision-summary.json') -Encoding UTF8
Write-Host "Formula decision summary complete: $OutputDir"
Write-Host "Records: $($rows.Count); Converted: $($counts.Converted); OriginalKept: $($counts.OriginalKept); ManualRequired: $($counts.ManualRequired)"
