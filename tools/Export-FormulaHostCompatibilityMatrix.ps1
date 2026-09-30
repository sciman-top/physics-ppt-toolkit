<#
.SYNOPSIS
  Build a read-only host compatibility matrix for multiple real PPTX samples.

.DESCRIPTION
  Runs Export-FormulaHostCompatibilityReceipt.ps1 once per input deck and
  aggregates the resulting PowerPoint, WPS, and projector statuses. Every
  receipt is written outside the source deck, input hashes are retained, and a
  missing host never counts as a pass. The matrix is evidence only; it never
  enables formula or PPTX write-back.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string[]]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-Sha256Local {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

$OutputDir = [IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$receiptTool = Join-Path $PSScriptRoot 'Export-FormulaHostCompatibilityReceipt.ps1'
if (-not (Test-Path -LiteralPath $receiptTool -PathType Leaf)) { throw "Host receipt tool not found: $receiptTool" }

$resolvedInputs = New-Object System.Collections.Generic.List[string]
foreach ($path in @($InputPath)) {
    $full = [IO.Path]::GetFullPath($path)
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "Input PPTX not found: $full" }
    if ([IO.Path]::GetExtension($full) -notin @('.pptx', '.pptm')) { throw "Input is not a PowerPoint package: $full" }
    if ($resolvedInputs -contains $full) { throw "Duplicate input path: $full" }
    $resolvedInputs.Add($full) | Out-Null
}
if ($resolvedInputs.Count -eq 0) { throw 'At least one input deck is required.' }

$decks = New-Object System.Collections.Generic.List[object]
$allTargets = New-Object System.Collections.Generic.List[object]
foreach ($full in $resolvedInputs) {
    $stem = Convert-ToSafePathSegment -Name ([IO.Path]::GetFileNameWithoutExtension($full))
    $deckDir = Join-Path $OutputDir $stem
    New-Item -ItemType Directory -Path $deckDir -Force | Out-Null
    & pwsh -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $receiptTool -InputPath $full -OutputDir $deckDir | Out-Null
    $receiptExit = $LASTEXITCODE
    if ($receiptExit -ne 0) { throw "Host receipt generation failed (exit=$receiptExit): $full" }
    $receiptPath = Join-Path $deckDir 'formula-host-compatibility-receipt.json'
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw "Host receipt missing: $receiptPath" }
    $receipt = Get-Content -LiteralPath $receiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $inputHash = Get-Sha256Local -Path $full
    if ([string]$receipt.input.sha256 -ne $inputHash) { throw "Receipt input hash mismatch: $full" }
    if ([bool]$receipt.policy.writeBackAllowed -or [bool]$receipt.policy.inputModified) { throw "Unsafe host receipt policy: $receiptPath" }
    $targets = @($receipt.targets)
    foreach ($target in $targets) { $allTargets.Add($target) | Out-Null }
    $decks.Add([ordered]@{
        path = $full
        sha256 = $inputHash
        receiptPath = $receiptPath
        targets = $targets
    }) | Out-Null
}

$hostSummary = New-Object System.Collections.Generic.List[object]
foreach ($group in @($allTargets | Group-Object -Property host)) {
    $rows = @($group.Group)
    $hostSummary.Add([ordered]@{
        host = [string]$group.Name
        total = $rows.Count
        passed = @($rows | Where-Object { [string]$_.status -eq 'Passed' }).Count
        failed = @($rows | Where-Object { [string]$_.status -eq 'Failed' }).Count
        unavailable = @($rows | Where-Object { [string]$_.status -eq 'Unavailable' }).Count
        notEvaluated = @($rows | Where-Object { [string]$_.status -eq 'NotEvaluated' }).Count
    }) | Out-Null
}

$matrix = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{
        readOnly = $true
        writeBackAllowed = $false
        migrationExpansionAllowed = $false
    }
    decks = @($decks.ToArray())
    hostSummary = @($hostSummary.ToArray())
    limitations = @(
        'WPS is unavailable unless every matrix row has an actual WPS open/export receipt.',
        'Physical projection and extended-display behavior are not evaluated by local automation.',
        'A PowerPoint pass proves host open/export only; it does not prove formula semantic equivalence or classroom projection stability.'
    )
}
$matrixPath = Join-Path $OutputDir 'formula-host-compatibility-matrix.json'
$matrix | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $matrixPath -Encoding UTF8
Write-Host "Formula host compatibility matrix written: $matrixPath"
