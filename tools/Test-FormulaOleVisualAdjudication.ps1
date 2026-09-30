<##
.SYNOPSIS
  Verify current OLE visual adjudication coverage and hash-drift rejection.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('physics-formula-ole-adjudication-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $source = Join-Path $PSScriptRoot '..\examples\fixtures\minimal-physics-sample.pptx'
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Fixture PPTX not found: $source" }
    $sourceCopy = Join-Path $testRoot 'fixture.pptx'
    Copy-Item -LiteralPath $source -Destination $sourceCopy -Force
    $sourceSha = (Get-FileHash -LiteralPath $sourceCopy -Algorithm SHA256).Hash.ToLowerInvariant()

    Add-Type -AssemblyName System.Drawing | Out-Null
    $pageDir = Join-Path $testRoot 'crops'; New-Item -ItemType Directory -Path $pageDir -Force | Out-Null
    $pagePath = Join-Path $pageDir 'slide-001.png'
    $bitmap = [Drawing.Bitmap]::new(12, 12)
    try { $bitmap.SetPixel(4, 4, [Drawing.Color]::Black); $bitmap.Save($pagePath, [Drawing.Imaging.ImageFormat]::Png) } finally { $bitmap.Dispose() }
    $pageSha = (Get-FileHash -LiteralPath $pagePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $cropPath = Join-Path $pageDir 's1-sh1-MathTypeOle_slide-001.png'
    Copy-Item -LiteralPath $pagePath -Destination $cropPath -Force
    $cropSha = (Get-FileHash -LiteralPath $cropPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $recordSourceSha = ('1' * 64)
    $mediaSha = ('2' * 64)
    $record = [ordered]@{
        schemaVersion = 1
        recordId = 's1-sh1-MathTypeOle'
        source = [ordered]@{ filePath = $sourceCopy; fileSha256 = $sourceSha; sourceSha256 = $recordSourceSha; carrier = 'MathTypeOle'; slide = 1; shapeId = 1; shapeName = 'Object 1'; bbox = [ordered]@{ left = 1; top = 1; width = 2; height = 2 }; mediaPaths = @('ppt/media/image1.wmf'); mediaSha256 = @($mediaSha); oleProgId = 'Equation.DSMT4'; fallbackCarrier = ''; originalText = '' }
        detection = [ordered]@{ method = 'OpenXml'; status = 'Observed'; confidence = 1; rawCandidates = @() }
        canonical = [ordered]@{ status = 'Unresolved'; source = [ordered]@{ kind = 'None'; id = ''; sha256 = $null }; unicodeMath = ''; tex = ''; mathml = ''; tokens = @() }
        decision = [ordered]@{ mode = 'CandidateOnly'; status = 'OriginalKept'; targetCarrier = 'Original'; reason = 'fixture' }
        evidence = [ordered]@{ paths = @($sourceCopy); rollbackPath = $sourceCopy; validatorPaths = @(); sourcePreviewSha256 = $null; targetPreviewSha256 = $null }
    }
    $inventoryPath = Join-Path $testRoot 'inventory.json'
    $inventory = [ordered]@{ schemaVersion = 1; generatedAt = 'fixture'; input = [ordered]@{ path = $sourceCopy; sha256 = $sourceSha }; records = @($record) }
    Write-Utf8BomText -Text ($inventory | ConvertTo-Json -Depth 20) -Path $inventoryPath
    $inventorySha = (Get-FileHash -LiteralPath $inventoryPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $cropCsv = Join-Path $testRoot 'ole-crops.csv'
    Write-Utf8BomCsv -InputObject @([pscustomobject]@{ RecordId = 's1-sh1-MathTypeOle'; CropPng = $cropPath; CropSha256 = $cropSha; PageSha256 = $pageSha }) -Path $cropCsv
    $cropManifestPath = Join-Path $testRoot 'formula-ole-crops-manifest.json'
    $cropManifest = [ordered]@{ schemaVersion = 1; generatedAt = 'fixture'; input = [ordered]@{ path = $sourceCopy; sha256 = $sourceSha }; inventory = [ordered]@{ path = $inventoryPath; sha256 = $inventorySha }; pages = [ordered]@{ dir = $pageDir; imageWidthPx = 12; slideSizeEmu = '12x12' }; carrier = 'MathTypeOle'; recordCount = 1; croppedCount = 1; missingCount = 0; missing = @(); evidenceSetSha256 = ('3' * 64); csv = $cropCsv; cropDir = $pageDir; writeBackAllowed = $false; note = 'fixture' }
    Write-Utf8BomText -Text ($cropManifest | ConvertTo-Json -Depth 20) -Path $cropManifestPath
    $adjudicationPath = Join-Path $testRoot 'adjudication.csv'
    Write-Utf8BomCsv -InputObject @([pscustomobject]@{ RecordId = 's1-sh1-MathTypeOle'; SourceSha256 = $recordSourceSha; CropPng = $cropPath; CropSha256 = $cropSha; PageSha256 = $pageSha; Status = 'CandidateOnly'; TargetCarrier = 'OfficeMath'; UnicodeMath = 'Q_放'; TeX = 'Q_{\text{放}}'; Basis = 'fixture'; ContextAssessment = 'isolated'; Reviewer = 'fixture'; Confidence = '0.9' }) -Path $adjudicationPath
    $tool = Join-Path $PSScriptRoot 'Export-FormulaOleVisualAdjudication.ps1'
    $goodOutput = Join-Path $testRoot 'good-output'
    & $tool -CarrierInventoryJson $inventoryPath -CropManifestJson $cropManifestPath -AdjudicationCsv $adjudicationPath -OutputDir $goodOutput | Out-Null
    $goodManifest = Get-Content -LiteralPath (Join-Path $goodOutput 'formula-ole-visual-adjudication-proposal.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int]$goodManifest.counts.total -ne 1 -or [int]$goodManifest.counts.candidateOnly -ne 1 -or [bool]$goodManifest.policy.writeBackAllowed) { throw 'Valid OLE adjudication fixture did not produce a safe candidate-only manifest.' }

    $badAdjudication = Join-Path $testRoot 'drifted-adjudication.csv'
    $badRow = [pscustomobject]@{ RecordId = 's1-sh1-MathTypeOle'; SourceSha256 = $recordSourceSha; CropPng = $cropPath; CropSha256 = ('f' * 64); PageSha256 = $pageSha; Status = 'CandidateOnly'; TargetCarrier = 'OfficeMath'; UnicodeMath = 'Q_放'; TeX = 'Q_{\text{放}}'; Basis = 'fixture'; ContextAssessment = 'isolated'; Reviewer = 'fixture'; Confidence = '0.9' }
    Write-Utf8BomCsv -InputObject @($badRow) -Path $badAdjudication
    $badOutput = Join-Path $testRoot 'bad-output'
    $failedAsExpected = $false
    try { & $tool -CarrierInventoryJson $inventoryPath -CropManifestJson $cropManifestPath -AdjudicationCsv $badAdjudication -OutputDir $badOutput | Out-Null } catch { $failedAsExpected = $true }
    if (-not $failedAsExpected) { throw 'Drifted crop hash was not rejected.' }
    Write-Host 'Formula OLE visual adjudication tests passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
