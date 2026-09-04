<#
.SYNOPSIS
  Fixture tests for invariant comparison and host-AI visual-review gates.

.DESCRIPTION
  Uses JSON-only temporary fixtures. It never starts PowerPoint and never creates or edits a PPTX.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolRoot = $PSScriptRoot
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('physics-ppt-policy-' + [Guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
    $beforePath = Join-Path $testRoot 'before.json'
    $afterPath = Join-Path $testRoot 'after.json'
    $blockedPath = Join-Path $testRoot 'blocked.json'
    $comparePath = Join-Path $testRoot 'compare.json'
    $manifestPath = Join-Path $testRoot 'review-manifest.json'
    $packetPath = Join-Path $testRoot 'ai-visual-review-request.json'
    $resultPath = Join-Path $testRoot 'ai-visual-review-result.json'
    $baselineAuditDir = Join-Path $testRoot 'baseline-audit'
    $normalizedAuditDir = Join-Path $testRoot 'normalized-audit'

    $baseSnapshot = [ordered]@{
        schemaVersion = 1; sourceSha256 = ('a' * 64); slideWidth = 960; slideHeight = 540; slideCount = 1
        slides = @([ordered]@{
            index = 1; slideId = 256; hidden = $false
            transition = [ordered]@{ advanceOnClick = -1; advanceOnTime = 0; advanceTime = 0; entryEffect = 0; speed = 0 }
            shapes = @([ordered]@{ id = 2; name = '正文'; type = 17; left = 10; top = 20; width = 200; height = 50; rotation = 0; zOrder = 1; autoSize = 0; wordWrap = -1; text = '速度 v=s/t' })
            animations = @()
        })
        package = [ordered]@{ relationships = @([ordered]@{ part = 'ppt/_rels/presentation.xml.rels'; id = 'rId1'; type = 'theme'; target = 'theme/theme1.xml'; targetMode = '' }); media = @([ordered]@{ part = 'ppt/media/image1.png'; length = 1; sha256 = ('b' * 64) }) }
    }
    $before = $baseSnapshot | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($beforePath, $before, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $afterObject = $before | ConvertFrom-Json
    $afterObject.slides[0].transition.advanceOnClick = 0
    [System.IO.File]::WriteAllText($afterPath, ($afterObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $afterPath -OutputPath $comparePath | Out-Null
    $comparison = Get-Content -LiteralPath $comparePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $comparison.passed -or $comparison.blockerCount -ne 0 -or $comparison.allowedChangeCount -ne 1) { throw 'Invariant comparison did not allow only AdvanceOnClick true -> false.' }

    $roundTripObject = $afterObject | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $roundTripObject.slides[0].shapes[0].top = 20.0001
    [System.IO.File]::WriteAllText($blockedPath, ($roundTripObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $roundTripResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $blockedPath
    $roundTripComparison = $roundTripResult | ConvertFrom-Json
    if (-not $roundTripComparison.passed -or $roundTripComparison.blockerCount -ne 0 -or $roundTripComparison.allowedChangeCount -ne 1) { throw 'Invariant comparison did not tolerate PowerPoint geometry round-trip noise.' }

    $blockedObject = $afterObject | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $blockedObject.slides[0].shapes[0].left = 11
    [System.IO.File]::WriteAllText($blockedPath, ($blockedObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $blockedResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $blockedPath
    $blockedComparison = $blockedResult | ConvertFrom-Json
    if ($blockedComparison.passed -or $blockedComparison.blockerCount -lt 1) { throw 'Invariant comparison did not block a geometry change.' }

    $mediaChangedObject = $afterObject | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $mediaChangedObject.package.media[0].sha256 = ('c' * 64)
    [System.IO.File]::WriteAllText($blockedPath, ($mediaChangedObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $mediaChangedResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $blockedPath
    $mediaChangedComparison = $mediaChangedResult | ConvertFrom-Json
    if ($mediaChangedComparison.passed -or $mediaChangedComparison.blockerCount -lt 1) { throw 'Invariant comparison did not block a media change.' }

    $input = 'C:\fixture\lesson.pptx'
    $sourceImage = 'C:\fixture\source\slide-1.png'
    $normalizedImage = 'C:\fixture\normalized\slide-1.png'
    [System.IO.File]::WriteAllText($manifestPath, ([ordered]@{ files = @([ordered]@{ input = $input }) } | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    [System.IO.File]::WriteAllText($packetPath, ([ordered]@{ protocolVersion = 1; files = @([ordered]@{ input = $input; status = 'Ready'; reason = ''; pages = @([ordered]@{ slide = 1; sourceImage = $sourceImage; normalizedImage = $normalizedImage }) }) } | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $passedResult = [ordered]@{
        protocolVersion = 1; reviewVersion = 'fixture-1'; model = 'fixture-model'; reviewedAt = '2026-01-01T00:00:00Z'; status = 'Passed'
        files = @([ordered]@{ input = $input; status = 'Passed'; pages = @([ordered]@{ slide = 1; sourceImage = $sourceImage; normalizedImage = $normalizedImage; status = 'Passed'; evidence = 'Fixture pair inspected; no introduced visual issue.'; issues = @() }) })
    }
    [System.IO.File]::WriteAllText($resultPath, ($passedResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $gate = & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | ConvertFrom-Json
    if ($gate.status -ne 'Passed' -or $gate.deliveryBlocked) { throw 'AI review gate did not accept a fully paired passed result.' }

    $passedResult.files[0].status = 'Blocked'
    $passedResult.files[0].pages[0].status = 'Blocked'
    $passedResult.files[0].pages[0].issues = @([ordered]@{ type = 'Clipping'; severity = 'Error'; evidence = 'Fixture regression: normalized text is visibly clipped.' })
    $passedResult.status = 'Blocked'
    [System.IO.File]::WriteAllText($resultPath, ($passedResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $blockedGate = & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | ConvertFrom-Json
    if ($blockedGate.status -ne 'Blocked' -or -not $blockedGate.deliveryBlocked) { throw 'AI review gate did not block a known clipping regression.' }

    $passedResult.status = 'Passed'
    [System.IO.File]::WriteAllText($resultPath, ($passedResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $rejected = $false
    try { & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | Out-Null } catch { $rejected = $true }
    if (-not $rejected) { throw 'AI review gate accepted an invalid aggregate status.' }

    # A rendered issue present in the original PPTX must remain a known baseline issue,
    # rather than becoming a false failure for the normalized copy.
    New-Item -ItemType Directory -Path $baselineAuditDir, $normalizedAuditDir -Force | Out-Null
    Add-Type -AssemblyName System.Drawing
    $fixtureImage = Join-Path $testRoot 'slide-001.png'
    $bitmap = New-Object System.Drawing.Bitmap 160, 90
    try { $bitmap.Save($fixtureImage, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $bitmap.Dispose() }
    $auditColumns = @('Timestamp', 'File', 'Slide', 'Shape', 'Issue', 'Severity', 'Details')
    $metricColumns = @('File', 'Slide', 'ImagePath', 'Width', 'Height', 'WhitePercent', 'NonWhitePercent', 'DarkPercent', 'IsVisuallyBlank', 'ContentTouchesEdge', 'ContentBounds', 'IssueCount')
    $baselineAudit = [pscustomobject]@{ Timestamp = '2026-01-01 00:00:00'; File = 'source.pptx'; Slide = 1; Shape = 'TextBox 2'; Issue = 'ShapeOutOfSlideBounds'; Severity = 'Error'; Details = 'fixture baseline overflow' }
    $normalizedAudit = [pscustomobject]@{ Timestamp = '2026-01-01 00:00:00'; File = 'normalized.pptx'; Slide = 1; Shape = 'TextBox 2'; Issue = 'ShapeOutOfSlideBounds'; Severity = 'Error'; Details = 'fixture baseline overflow' }
    $metric = [pscustomobject]@{ File = 'fixture.pptx'; Slide = 1; ImagePath = $fixtureImage; Width = 160; Height = 90; WhitePercent = 100; NonWhitePercent = 0; DarkPercent = 0; IsVisuallyBlank = $false; ContentTouchesEdge = $false; ContentBounds = ''; IssueCount = 1 }
    [System.IO.File]::WriteAllText((Join-Path $baselineAuditDir 'pptx-visual-audit.csv'), (($auditColumns -join ',') + "`r`n" + (($baselineAudit | Select-Object $auditColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n")), (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    [System.IO.File]::WriteAllText((Join-Path $normalizedAuditDir 'pptx-visual-audit.csv'), (($auditColumns -join ',') + "`r`n" + (($normalizedAudit | Select-Object $auditColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n")), (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    $metricCsv = (($metricColumns -join ',') + "`r`n" + (($metric | Select-Object $metricColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n"))
    [System.IO.File]::WriteAllText((Join-Path $baselineAuditDir 'pptx-slide-visual-metrics.csv'), $metricCsv, (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    [System.IO.File]::WriteAllText((Join-Path $normalizedAuditDir 'pptx-slide-visual-metrics.csv'), $metricCsv, (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    & (Join-Path $toolRoot 'Export-PptxVisualConfirmation.ps1') -VisualAuditDir $normalizedAuditDir -BaselineVisualAuditDir $baselineAuditDir | Out-Null
    $baselineConfirmation = Get-Content -LiteralPath (Join-Path $normalizedAuditDir 'pptx-visual-confirmation-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($baselineConfirmation.confirmationStatus -ne 'PassedWithKnownIssues' -or $baselineConfirmation.newErrorCount -ne 0) { throw 'Visual confirmation did not retain an original rendered issue as a known baseline issue.' }
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}

Write-Host 'Physics PPT policy fixture tests passed.'
