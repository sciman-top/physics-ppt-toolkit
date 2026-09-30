<#
.SYNOPSIS
  Verify fallback SVG and original-preservation downgrade behavior.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('physics-formula-fallback-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $input = Join-Path $testRoot 'candidates.csv'
    Write-Utf8BomCsv -InputObject @(
        [pscustomobject]@{ CandidateId = 'unsupported-with-tex'; Status = 'Failed'; FormulaIrStatus = 'Unsupported'; TargetTex = 'x^2' }
        [pscustomobject]@{ CandidateId = 'unsupported-no-tex'; Status = 'Failed'; FormulaIrStatus = 'Unsupported'; TargetTex = '' }
        [pscustomobject]@{ CandidateId = 'resolved'; Status = 'Generated'; FormulaIrStatus = 'Resolved'; TargetTex = 'Q_{\text{放}}=qm' }
    ) -Path $input
    & (Join-Path $PSScriptRoot 'Plan-FormulaFallbackSvg.ps1') -FormulaCandidateCsv $input -OutputDir (Join-Path $testRoot 'output')
    if (-not $?) { throw 'Formula fallback planner did not complete.' }
    $manifest = Get-Content -LiteralPath (Join-Path $testRoot 'output\formula-fallback-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $rows = @{}; foreach ($row in @($manifest.rows)) { $rows[[string]$row.CandidateId] = $row }
    if ([string]$rows['unsupported-with-tex'].DecisionStatus -ne 'FallbackSvg' -or -not (Test-Path -LiteralPath ([string]$rows['unsupported-with-tex'].SvgPath))) { throw 'Unsupported formula with canonical TeX did not produce FallbackSvg evidence.' }
    if ([string]$rows['unsupported-no-tex'].DecisionStatus -ne 'OriginalKept') { throw 'Unsupported formula without TeX did not preserve the original.' }
    if ([string]$rows['resolved'].DecisionStatus -ne 'CandidateOnly') { throw 'Resolved candidate unexpectedly selected SVG fallback.' }
    if (@($manifest.rows | Where-Object { $_.WriteBackAllowed }).Count -ne 0) { throw 'Fallback planner enabled write-back.' }
    Write-Host 'Formula fallback SVG tests passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
