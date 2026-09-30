<#
.SYNOPSIS
  Plan a non-destructive SVG fallback for unsupported OMML candidates.

.DESCRIPTION
  This tool renders canonical TeX into SVG evidence only. It never applies an
  SVG to a PPTX. Unsupported candidates with missing/failed rendering remain
  OriginalKept, ensuring no half-converted formula reaches a delivery copy.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FormulaCandidateCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$NodeExe = 'node'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-Sha256FileLocal { param([string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }

$FormulaCandidateCsv = [IO.Path]::GetFullPath($FormulaCandidateCsv)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $FormulaCandidateCsv -PathType Leaf)) { throw "FormulaCandidateCsv not found: $FormulaCandidateCsv" }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$svgDir = Join-Path $OutputDir 'svg'
New-Item -ItemType Directory -Path $svgDir -Force | Out-Null
$renderer = Join-Path $PSScriptRoot 'Render-FormulaSvg.mjs'
if (-not (Test-Path -LiteralPath $renderer -PathType Leaf)) { throw "SVG renderer not found: $renderer" }

$results = New-Object System.Collections.Generic.List[object]
$rows = @(Import-Csv -LiteralPath $FormulaCandidateCsv -Encoding UTF8)
for ($index = 0; $index -lt $rows.Count; $index++) {
    $row = $rows[$index]
    $candidateId = if ($null -ne $row.PSObject.Properties['CandidateId'] -and -not [string]::IsNullOrWhiteSpace([string]$row.CandidateId)) { [string]$row.CandidateId } else { 'candidate-{0:000}' -f ($index + 1) }
    $candidateStatus = if ($null -ne $row.PSObject.Properties['Status']) { [string]$row.Status } else { '' }
    $irStatus = if ($null -ne $row.PSObject.Properties['FormulaIrStatus']) { [string]$row.FormulaIrStatus } else { '' }
    $tex = if ($null -ne $row.PSObject.Properties['TargetTex']) { [string]$row.TargetTex } else { '' }
    $needsFallback = $candidateStatus -ne 'Generated' -or $irStatus -in @('Unsupported', 'Unresolved', 'Failed')
    $decision = 'CandidateOnly'
    $svgPath = ''
    $reason = 'OMML candidate is structurally resolved; this planner does not select a fallback.'
    if ($needsFallback) {
        $decision = 'OriginalKept'
        $reason = 'Unsupported OMML candidate has no canonical TeX fallback.'
        if (-not [string]::IsNullOrWhiteSpace($tex)) {
            $safeName = Convert-ToSafeFormulaPathSegment -Name $candidateId
            $texPath = Join-Path $svgDir ($safeName + '.tex')
            $svgPath = Join-Path $svgDir ($safeName + '.svg')
            [IO.File]::WriteAllText($texPath, $tex, (New-Object Text.UTF8Encoding($false)))
            $renderOutput = & $NodeExe $renderer '--tex-file' $texPath '--out' $svgPath 2>&1
            if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $svgPath -PathType Leaf) -and (Get-Item -LiteralPath $svgPath).Length -gt 0) {
                $decision = 'FallbackSvg'
                $reason = 'OMML candidate is unsupported; canonical TeX rendered to SVG review evidence. Applying it remains a separate explicit PPTX operation.'
            } else {
                $svgPath = ''
                $reason = 'OMML candidate is unsupported and SVG rendering failed; original carrier must remain.'
            }
        }
    }
    $results.Add([pscustomobject]@{
        CandidateId = $candidateId
        InputStatus = $candidateStatus
        FormulaIrStatus = $irStatus
        DecisionStatus = $decision
        TargetCarrier = if ($decision -eq 'FallbackSvg') { 'Svg' } elseif ($decision -eq 'OriginalKept') { 'Original' } else { 'None' }
        WriteBackAllowed = $false
        TargetTex = $tex
        SvgPath = $svgPath
        SvgSha256 = if ($svgPath) { Get-Sha256FileLocal -Path $svgPath } else { '' }
        Reason = $reason
    }) | Out-Null
}
$csvPath = Join-Path $OutputDir 'formula-fallback-plan.csv'
$jsonPath = Join-Path $OutputDir 'formula-fallback-plan.json'
Write-Utf8BomCsv -InputObject $results.ToArray() -Path $csvPath
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ mode = 'CandidateOnly'; writeBackAllowed = $false; fallbackRequiresCanonicalTex = $true; originalPreservedOnFailure = $true }
    input = [ordered]@{ candidateCsv = $FormulaCandidateCsv; candidateCsvSha256 = Get-Sha256FileLocal -Path $FormulaCandidateCsv }
    counts = [ordered]@{ total = $results.Count; fallbackSvg = @($results | Where-Object { $_.DecisionStatus -eq 'FallbackSvg' }).Count; originalKept = @($results | Where-Object { $_.DecisionStatus -eq 'OriginalKept' }).Count; candidateOnly = @($results | Where-Object { $_.DecisionStatus -eq 'CandidateOnly' }).Count }
    rows = @($results.ToArray())
}
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
Write-Host "Formula fallback plan complete: $OutputDir"
Write-Host "FallbackSvg: $($manifest.counts.fallbackSvg); OriginalKept: $($manifest.counts.originalKept)"
