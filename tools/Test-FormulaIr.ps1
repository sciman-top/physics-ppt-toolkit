<#
.SYNOPSIS
  Run deterministic FormulaIR parser, canonical rendering, and negative-path tests.

.DESCRIPTION
  Uses the production candidate exporter against a temporary CSV. The test
  never writes a PPTX and removes only its own temporary directory.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('physics-formula-ir-' + [Guid]::NewGuid().ToString('N'))
$outputDir = Join-Path $testRoot 'output'
$csvPath = Join-Path $testRoot 'formula-review.csv'
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $rows = @(
        [pscustomobject]@{ File = 'fixture.pptx'; FilePath = 'fixture.pptx'; FileRelativePath = ''; Slide = 1; Shape = 'Text 1'; FormulaText = 'Q放=qm'; CandidateClass = 'Test'; WhitelistCandidate = 'name=test-equality; targetUnicodeMath=Q_放=qm; targetTex=Q_{\text{放}}=qm'; ConversionStatus = 'Test'; StyleStatus = ''; SuggestedAction = 'ReviewWhitelistConversion' }
        [pscustomobject]@{ File = 'fixture.pptx'; FilePath = 'fixture.pptx'; FileRelativePath = ''; Slide = 2; Shape = 'Text 2'; FormulaText = 'η=Q吸/Q放'; CandidateClass = 'Test'; WhitelistCandidate = 'name=test-fraction; targetUnicodeMath=η=Q_吸/Q_放; targetTex=\eta=\frac{Q_{\text{吸}}}{Q_{\text{放}}}'; ConversionStatus = 'Test'; StyleStatus = ''; SuggestedAction = 'ReviewWhitelistConversion' }
        [pscustomobject]@{ File = 'fixture.pptx'; FilePath = 'fixture.pptx'; FileRelativePath = ''; Slide = 3; Shape = 'Text 3'; FormulaText = '(Q吸=cmΔt)'; CandidateClass = 'Test'; WhitelistCandidate = 'name=test-group; targetUnicodeMath=(Q_吸=cmΔt); targetTex=(Q_{\text{吸}}=cm\Delta t)'; ConversionStatus = 'Test'; StyleStatus = ''; SuggestedAction = 'ReviewWhitelistConversion' }
        [pscustomobject]@{ File = 'fixture.pptx'; FilePath = 'fixture.pptx'; FileRelativePath = ''; Slide = 4; Shape = 'Text 4'; FormulaText = 'x2'; CandidateClass = 'Test'; WhitelistCandidate = 'name=test-superscript; targetUnicodeMath=x^2; targetTex=x^2'; ConversionStatus = 'Test'; StyleStatus = ''; SuggestedAction = 'ReviewWhitelistConversion' }
        [pscustomobject]@{ File = 'fixture.pptx'; FilePath = 'fixture.pptx'; FileRelativePath = ''; Slide = 5; Shape = 'Text 5'; FormulaText = '=cmΔt'; CandidateClass = 'Test'; WhitelistCandidate = 'name=test-leading-equality; targetUnicodeMath==cmΔt; targetTex==cm\Delta t'; ConversionStatus = 'Test'; StyleStatus = ''; SuggestedAction = 'ReviewWhitelistConversion' }
    )
    Write-Utf8BomCsv -InputObject $rows -Path $csvPath
    & pwsh -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Export-FormulaOmmlCandidates.ps1') -FormulaReviewCsv $csvPath -OutputDir $outputDir -MaxItems 100
    if ($LASTEXITCODE -ne 0) { throw "FormulaIR exporter exited with $LASTEXITCODE." }

    $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'formula-omml-candidates-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int]$manifest.candidateCount -ne 5 -or [int]$manifest.generatedCount -ne 5 -or [int]$manifest.formulaIrResolvedCount -ne 5) {
        throw "Unexpected FormulaIR test counts: candidates=$($manifest.candidateCount), generated=$($manifest.generatedCount), resolved=$($manifest.formulaIrResolvedCount)."
    }
    $records = @(Get-ChildItem -LiteralPath (Join-Path $outputDir 'formula-ir') -Filter '*.json' -File)
    if ($records.Count -ne 5) { throw "Expected 5 FormulaIR records, found $($records.Count)." }
    $all = @(Get-Content -LiteralPath (Join-Path $outputDir 'formula-omml-candidates.json') -Raw -Encoding UTF8 | ConvertFrom-Json)
    $bySlide = @{}
    foreach ($row in $all) { $bySlide[[int]$row.Slide] = $row }
    foreach ($slide in 1..5) {
        if ([string]$bySlide[$slide].FormulaIrStatus -ne 'Resolved') { throw "Slide $slide was not resolved." }
        $ir = Get-Content -LiteralPath ([string]$bySlide[$slide].FormulaIrFragment) -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]$ir.decision.status -ne 'CandidateOnly' -or [bool]$ir.decision.writeBackAllowed) { throw "Slide $slide violated CandidateOnly/writeBackAllowed=false." }
        if ([string]$ir.canonical.unicodeMath -eq '' -or [string]$ir.canonical.tex -eq '' -or $ir.canonical.tokens.Count -eq 0) { throw "Slide $slide has incomplete canonical output." }
    }
    $fraction = Get-Content -LiteralPath ([string]$bySlide[2].FormulaIrFragment) -Raw -Encoding UTF8 | ConvertFrom-Json
    $fractionJson = $fraction.canonical.tokens | ConvertTo-Json -Depth 20 -Compress
    if ([string]$fraction.canonical.tokens.type -ne 'Equality' -or $fractionJson -notmatch 'Fraction') { throw 'Fraction token was not preserved.' }
    $group = Get-Content -LiteralPath ([string]$bySlide[3].FormulaIrFragment) -Raw -Encoding UTF8 | ConvertFrom-Json
    $groupJson = $group.canonical.tokens | ConvertTo-Json -Depth 20 -Compress
    if ($groupJson -notmatch 'Delimiter') { throw 'Parenthesis delimiter was not preserved.' }
    $product = Get-Content -LiteralPath ([string]$bySlide[1].FormulaIrFragment) -Raw -Encoding UTF8 | ConvertFrom-Json
    $productJson = $product.canonical.tokens | ConvertTo-Json -Depth 20 -Compress
    if ($productJson -notmatch 'implicitMultiplication') { throw 'Implicit multiplication was not represented.' }
    Write-Host 'FormulaIR tests passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
