<#
.SYNOPSIS
  Verify Formula Context Resolver collision and downgrade behavior.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('physics-formula-context-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $hashA = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
    $hashB = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
    $candidates = @(
        [pscustomobject]@{ CandidateId = 'exact-goldset'; SourceText = 'Q放'; Carrier = 'MathTypeOle'; Slide = '6'; ShapeId = '10'; SourceSha256 = $hashA }
        [pscustomobject]@{ CandidateId = 'incomplete-context'; SourceText = 'Q放'; Carrier = 'MathTypeOle'; Slide = '6'; ShapeId = ''; SourceSha256 = $hashA }
        [pscustomobject]@{ CandidateId = 'ambiguous-context'; SourceText = 'Q放'; Carrier = 'MathTypeOle'; Slide = '6'; ShapeId = '11'; SourceSha256 = $hashB }
        [pscustomobject]@{ CandidateId = 'no-match'; SourceText = '未知量=42'; Carrier = 'TextFormula'; Slide = '1'; ShapeId = '1'; SourceSha256 = $hashB }
    )
    $maps = @(
        [pscustomobject]@{ Carrier = 'MathTypeOle'; Slide = '6'; ShapeId = '10'; SourceSha256 = $hashA; FormulaName = '放热量符号'; TargetUnicodeMath = 'Q_放'; TargetTex = 'Q_{\text{放}}' }
        [pscustomobject]@{ Carrier = 'MathTypeOle'; Slide = '6'; ShapeId = '11'; SourceSha256 = $hashB; FormulaName = '冲突一'; TargetUnicodeMath = 'Q_放'; TargetTex = 'Q_{\text{放}}' }
        [pscustomobject]@{ Carrier = 'MathTypeOle'; Slide = '6'; ShapeId = '11'; SourceSha256 = $hashB; FormulaName = '冲突二'; TargetUnicodeMath = 'Q_吸'; TargetTex = 'Q_{\text{吸}}' }
    )
    $candidateCsv = Join-Path $testRoot 'candidates.csv'
    $mapCsv = Join-Path $testRoot 'context.csv'
    Write-Utf8BomCsv -InputObject $candidates -Path $candidateCsv
    Write-Utf8BomCsv -InputObject $maps -Path $mapCsv
    & (Join-Path $PSScriptRoot 'Resolve-FormulaCanonicalContext.ps1') -CandidateCsv $candidateCsv -ContextMapCsv $mapCsv -OutputDir (Join-Path $testRoot 'output')
    if (-not $?) { throw 'Context resolver did not complete.' }
    $result = Get-Content -LiteralPath (Join-Path $testRoot 'output\formula-context-resolution.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $rows = @{}; foreach ($row in @($result.results)) { $rows[[string]$row.CandidateId] = $row }
    if ([string]$rows['exact-goldset'].DecisionStatus -ne 'CandidateOnly' -or [string]$rows['exact-goldset'].CanonicalSource -ne 'GoldSet') { throw 'Exact GoldSet context was not resolved as CandidateOnly.' }
    foreach ($id in @('incomplete-context', 'ambiguous-context', 'no-match')) {
        if ([string]$rows[$id].DecisionStatus -ne 'ManualRequired') { throw "Unsafe candidate '$id' did not downgrade to ManualRequired." }
        if ([bool]$rows[$id].WriteBackAllowed) { throw "Unsafe candidate '$id' enabled write-back." }
    }
    if ([int]$result.counts.candidateOnly -ne 1 -or [int]$result.counts.manualRequired -ne 3) { throw 'Context resolver count mismatch.' }
    Write-Host 'Formula context resolver tests passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
