<#
.SYNOPSIS
  Build a human-reviewed, hash-bound formula recognition gold set, or verify
  the evidence availability of an existing GoldSet manifest.

.DESCRIPTION
  BuildGoldSet parameter set: combines approved OLE visual evidence (positive
  formula examples) with an explicit image-candidate adjudication CSV. It is an
  evaluation artifact only: it never writes a PPTX and never authorizes
  write-back.

  VerifyGoldSetEvidence parameter set (formerly the standalone
  Export-FormulaGoldSetEvidenceAvailability.ps1): historical GoldSet manifests
  may retain paths from an archived run. This read-only mode searches a
  supplied evidence root for a file with the same basename and expected hash,
  then writes a derived availability receipt. It never changes the historical
  manifest, source PPTX, or write-back policy.
#>
[CmdletBinding(DefaultParameterSetName = 'BuildGoldSet')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'BuildGoldSet')]
    [string]$FormulaImageCandidatesCsv,

    [Parameter(Mandatory = $true, ParameterSetName = 'BuildGoldSet')]
    [string]$ImageAdjudicationCsv,

    [Parameter(Mandatory = $true, ParameterSetName = 'BuildGoldSet')]
    [string]$OleMappingManifestJson,

    [Parameter(Mandatory = $true, ParameterSetName = 'BuildGoldSet')]
    [string]$OutputDir,

    [Parameter(Mandatory = $true, ParameterSetName = 'VerifyGoldSetEvidence')]
    [string]$GoldSetManifestPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'VerifyGoldSetEvidence')]
    [string]$EvidenceRoot,

    [Parameter(Mandatory = $true, ParameterSetName = 'VerifyGoldSetEvidence')]
    [string]$EvidenceReceiptPath,

    [ValidateRange(1, 1000)]
    [int]$MaxItems = 1000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-RelativeOrFail {
    param([string]$BaseDirectory, [string]$Path)
    # Shared helper: Path.GetRelativePath does not exist on the 5.1 fallback host.
    return (Get-ManifestRelativePathOrThrow -BaseDirectory $BaseDirectory -TargetPath $Path)
}

if ($PSCmdlet.ParameterSetName -eq 'VerifyGoldSetEvidence') {
    $GoldSetManifestPath = [IO.Path]::GetFullPath($GoldSetManifestPath)
    $EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
    $EvidenceReceiptPath = [IO.Path]::GetFullPath($EvidenceReceiptPath)
    foreach ($path in @($GoldSetManifestPath, $EvidenceRoot)) {
        if (-not (Test-Path -LiteralPath $path)) { throw "Required evidence input not found: $path" }
    }
    if (-not (Test-Path -LiteralPath $GoldSetManifestPath -PathType Leaf)) { throw "GoldSet manifest is not a file: $GoldSetManifestPath" }
    if (-not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) { throw "EvidenceRoot is not a directory: $EvidenceRoot" }

    $manifestHash = Get-FileSha256Hex -Path $GoldSetManifestPath
    $evidenceManifest = Get-Content -LiteralPath $GoldSetManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $records = New-Object System.Collections.Generic.List[object]
    $available = 0
    $missing = 0
    $ambiguous = 0

    foreach ($record in @($evidenceManifest.records)) {
        $expected = ([string]$record.evidence.sha256).ToLowerInvariant()
        $leaf = Split-Path -Leaf ([string]$record.evidence.path)
        if ([string]::IsNullOrWhiteSpace($leaf)) { throw "GoldSet record has no evidence filename: $($record.goldSetId)" }
        $candidates = @(Get-ChildItem -LiteralPath $EvidenceRoot -Recurse -File -Filter $leaf -ErrorAction SilentlyContinue |
            Where-Object { (Get-FileSha256Hex -Path $_.FullName) -eq $expected })
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
    $receiptParent = Split-Path -Parent $EvidenceReceiptPath
    if (-not [string]::IsNullOrWhiteSpace($receiptParent)) { New-Item -ItemType Directory -Path $receiptParent -Force | Out-Null }
    $receipt | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $EvidenceReceiptPath -Encoding UTF8
    Write-Host "Formula GoldSet evidence availability written: $EvidenceReceiptPath"
    return
}

$FormulaImageCandidatesCsv = [IO.Path]::GetFullPath($FormulaImageCandidatesCsv)
$ImageAdjudicationCsv = [IO.Path]::GetFullPath($ImageAdjudicationCsv)
$OleMappingManifestJson = [IO.Path]::GetFullPath($OleMappingManifestJson)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
foreach ($path in @($FormulaImageCandidatesCsv, $ImageAdjudicationCsv, $OleMappingManifestJson)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required input not found: $path" }
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$imageRows = @(Import-Csv -LiteralPath $FormulaImageCandidatesCsv -Encoding UTF8)
$adjudicationRows = @(Import-Csv -LiteralPath $ImageAdjudicationCsv -Encoding UTF8)
$adjudication = @{}
foreach ($row in $adjudicationRows) {
    if ([string]::IsNullOrWhiteSpace([string]$row.MediaPath)) { throw 'Image adjudication row has empty MediaPath.' }
    if ($adjudication.ContainsKey([string]$row.MediaPath)) { throw "Duplicate image adjudication: $($row.MediaPath)" }
    $adjudication[[string]$row.MediaPath] = $row
}
$mappingManifest = Get-Content -LiteralPath $OleMappingManifestJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$mappingManifest.status -ne 'Passed') { throw "OLE mapping manifest is not Passed: $($mappingManifest.status)" }
$manifestInputPath = [IO.Path]::GetFullPath([string]$mappingManifest.input.path)
$expectedInputSha = ([string]$mappingManifest.input.sha256).ToLowerInvariant()
$sourcePptx = $manifestInputPath
if (-not (Test-Path -LiteralPath $sourcePptx -PathType Leaf) -or ((Get-FileSha256Hex -Path $sourcePptx).ToLowerInvariant() -ne $expectedInputSha)) {
    # Reports may outlive a user-visible source file deletion. Re-resolve only
    # by the manifest basename and exact SHA-256; a same-name file is never
    # accepted on name alone.
    $searchRoot = Split-Path -Parent $PSScriptRoot
    $matches = @(Get-ChildItem -LiteralPath (Join-Path $searchRoot 'reports') -Recurse -File -Filter ([IO.Path]::GetFileName($manifestInputPath)) -ErrorAction SilentlyContinue | Where-Object { (Get-FileSha256Hex -Path $_.FullName).ToLowerInvariant() -eq $expectedInputSha } | Sort-Object FullName)
    if ($matches.Count -eq 0) { throw "GoldSet source PPTX is missing or hash-mismatched, and no exact-hash report backup was found: $manifestInputPath" }
    $sourcePptx = $matches[0].FullName
}

$records = New-Object System.Collections.Generic.List[object]
foreach ($gold in @($mappingManifest.rows)) {
    if ($records.Count -ge $MaxItems) { break }
    $evidencePath = [IO.Path]::GetFullPath([string]$gold.EvidencePath)
    if (-not (Test-Path -LiteralPath $evidencePath -PathType Leaf)) { throw "OLE evidence not found: $evidencePath" }
    $records.Add([ordered]@{
        goldSetId = "ole-$([int]$gold.Slide)-$([string]$gold.OleIndex)"
        source = [ordered]@{ carrier = 'MathTypeOle'; slide = [int]$gold.Slide; sourceId = "ole-$([string]$gold.OleIndex)"; sourceSha256 = (Get-TextSha256Hex -Text "ole|$([string]$gold.InventoryRecordIds)|$([int]$gold.Slide)|$([string]$gold.OleIndex)|$([string]$mappingManifest.input.sha256)") }
        groundTruth = [ordered]@{ class = 'Formula'; difficulty = if ([string]$gold.TargetTex -match 'frac') { 'Hard' } elseif ([string]$gold.TargetUnicodeMath -match '[_^]') { 'Medium' } else { 'Easy' }; reviewStatus = 'HumanReviewed'; reviewBasis = 'Approved OLE GoldSet row with visual evidence'; highRiskTokens = @('ChineseSubscript') }
        canonical = [ordered]@{ status = 'Resolved'; formulaName = [string]$gold.WhitelistName; unicodeMath = [string]$gold.TargetUnicodeMath; tex = [string]$gold.TargetTex; tokens = @() }
        evidence = [ordered]@{ path = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $evidencePath); sha256 = (Get-FileSha256Hex -Path $evidencePath); kind = 'OleVisualEvidence'; rollbackPath = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $sourcePptx) }
    }) | Out-Null
}

foreach ($candidate in $imageRows) {
    if ($records.Count -ge $MaxItems) { break }
    $mediaPath = [string]$candidate.MediaPath
    if (-not $adjudication.ContainsKey($mediaPath)) { throw "Image candidate has no human adjudication: $mediaPath" }
    $review = $adjudication[$mediaPath]
    if ([string]$review.ReviewStatus -ne 'HumanReviewed') { throw "Image adjudication is not HumanReviewed: $mediaPath" }
    if ([string]$review.GroundTruthClass -notin @('Formula', 'NonFormula', 'MixedNonIsolatable')) { throw "Invalid image GroundTruthClass: $($review.GroundTruthClass)" }
    $evidencePath = [IO.Path]::GetFullPath([string]$candidate.FormulaExtractedPath)
    if (-not (Test-Path -LiteralPath $evidencePath -PathType Leaf)) { throw "Image crop evidence not found: $evidencePath" }
    $canonicalStatus = if ([string]$review.GroundTruthClass -eq 'Formula') { 'Resolved' } else { 'NotApplicable' }
    $records.Add([ordered]@{
        goldSetId = "image-$([IO.Path]::GetFileNameWithoutExtension($mediaPath))"
        source = [ordered]@{ carrier = if ([string]$candidate.MediaRole -eq 'MixedImage') { 'MixedImage' } else { 'FormulaImage' }; slide = [int](([string]$candidate.UsedOnSlides -split ',')[0]); sourceId = $mediaPath; sourceSha256 = (Get-FileSha256Hex -Path $evidencePath) }
        groundTruth = [ordered]@{ class = [string]$review.GroundTruthClass; difficulty = if ([string]$review.Difficulty) { [string]$review.Difficulty } else { 'NotApplicable' }; reviewStatus = 'HumanReviewed'; reviewBasis = [string]$review.ReviewBasis; highRiskTokens = @() }
        canonical = [ordered]@{ status = $canonicalStatus; formulaName = [string]$review.FormulaName; unicodeMath = [string]$review.UnicodeMath; tex = [string]$review.TeX; tokens = @() }
        evidence = [ordered]@{ path = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $evidencePath); sha256 = (Get-FileSha256Hex -Path $evidencePath); kind = 'FormulaCrop'; rollbackPath = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $sourcePptx) }
    }) | Out-Null
}

$hashInput = @($records | ForEach-Object { "$($_.goldSetId)|$($_.source.sourceSha256)|$($_.evidence.path)|$($_.evidence.sha256)|$($_.groundTruth.class)|$($_.canonical.unicodeMath)" } | Sort-Object) -join "`n"
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    input = [ordered]@{ sourcePptx = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $sourcePptx); sourcePptxSha256 = (Get-FileSha256Hex -Path $sourcePptx); imageCandidatesCsv = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $FormulaImageCandidatesCsv); imageAdjudicationCsv = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $ImageAdjudicationCsv); oleMappingManifest = (Get-RelativeOrFail -BaseDirectory $OutputDir -Path $OleMappingManifestJson) }
    policy = [ordered]@{ writeBackAllowed = $false; reviewRequirement = 'HumanReviewed'; canonicalSourceRule = 'Only approved OLE GoldSet or explicit human adjudication may populate canonical fields.' }
    counts = [ordered]@{ total = $records.Count; formula = @($records | Where-Object { $_.groundTruth.class -eq 'Formula' }).Count; nonFormula = @($records | Where-Object { $_.groundTruth.class -eq 'NonFormula' }).Count; mixedNonIsolatable = @($records | Where-Object { $_.groundTruth.class -eq 'MixedNonIsolatable' }).Count }
    evidenceSetSha256 = Get-TextSha256Hex -Text $hashInput
    records = @($records.ToArray())
}
$jsonPath = Join-Path $OutputDir 'formula-goldset-manifest.json'
$csvPath = Join-Path $OutputDir 'formula-goldset.csv'
$manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
Write-Utf8BomCsv -InputObject @($records | ForEach-Object { [pscustomobject]@{ GoldSetId = $_.goldSetId; Carrier = $_.source.carrier; Slide = $_.source.slide; SourceId = $_.source.sourceId; SourceSha256 = $_.source.sourceSha256; GroundTruthClass = $_.groundTruth.class; Difficulty = $_.groundTruth.difficulty; ReviewStatus = $_.groundTruth.reviewStatus; CanonicalStatus = $_.canonical.status; UnicodeMath = $_.canonical.unicodeMath; TeX = $_.canonical.tex; EvidencePath = $_.evidence.path; EvidenceSha256 = $_.evidence.sha256; EvidenceKind = $_.evidence.kind } }) -Path $csvPath
Write-Host "Formula GoldSet exported: $OutputDir"
Write-Host "Total: $($manifest.counts.total); Formula: $($manifest.counts.formula); NonFormula: $($manifest.counts.nonFormula); MixedNonIsolatable: $($manifest.counts.mixedNonIsolatable)"
