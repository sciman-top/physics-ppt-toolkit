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
            index = 1; slideId = 256; hidden = $false; unreadableShapeCount = 0; slideReadStatus = 'Readable'; animationReadStatus = 'Readable'; transitionReadStatus = 'Readable'
            transition = [ordered]@{ advanceOnClick = -1; advanceOnTime = 0; advanceTime = 0; entryEffect = 0; speed = 0 }
            shapes = @([ordered]@{ id = 2; name = '正文'; type = 17; left = 10; top = 20; width = 200; height = 50; rotation = 0; zOrder = 1; autoSize = 0; wordWrap = -1; text = '速度 v=s/t'; readStatus = 'Readable' })
            animations = @()
        })
        package = [ordered]@{ relationships = @([ordered]@{ part = 'ppt/_rels/presentation.xml.rels'; id = 'rId1'; type = 'theme'; target = 'theme/theme1.xml'; targetMode = '' }); media = @([ordered]@{ part = 'ppt/media/image1.png'; length = 1; sha256 = ('b' * 64) }) }
    }
    $before = $baseSnapshot | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($beforePath, $before, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $afterObject = $before | ConvertFrom-Json
    $afterObject.slides[0].transition.advanceOnClick = 0
    [System.IO.File]::WriteAllText($afterPath, ($afterObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $afterPath -OutputPath $comparePath -AllowAdvanceOnClickDisable | Out-Null
    $comparison = Get-Content -LiteralPath $comparePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $comparison.passed -or $comparison.blockerCount -ne 0 -or $comparison.allowedChangeCount -ne 1) { throw 'Invariant comparison did not allow only AdvanceOnClick true -> false.' }

    $unauthorizedAdvanceResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $afterPath
    $unauthorizedAdvanceComparison = $unauthorizedAdvanceResult | ConvertFrom-Json
    if ($unauthorizedAdvanceComparison.passed -or $unauthorizedAdvanceComparison.blockerCount -lt 1 -or $unauthorizedAdvanceComparison.allowedChangeCount -ne 0) { throw 'Invariant comparison allowed AdvanceOnClick disable without explicit authorization.' }

    $roundTripObject = $afterObject | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $roundTripObject.slides[0].shapes[0].top = 20.0001
    [System.IO.File]::WriteAllText($blockedPath, ($roundTripObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $roundTripResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $blockedPath -AllowAdvanceOnClickDisable
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

    # FR-04's core promise: any text content change is a blocker even when
    # geometry, media, and animations are untouched.
    $textChangedObject = $afterObject | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $textChangedObject.slides[0].shapes[0].text = '速度 v=u/t'
    [System.IO.File]::WriteAllText($blockedPath, ($textChangedObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $textChangedResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $beforePath -AfterPath $blockedPath
    $textChangedComparison = $textChangedResult | ConvertFrom-Json
    if ($textChangedComparison.passed -or $textChangedComparison.blockerCount -lt 1) { throw 'Invariant comparison did not block a text content change.' }

    $input = 'C:\fixture\lesson.pptx'
    $sourceImage = Join-Path $testRoot 'source-slide-1.png'
    $normalizedImage = Join-Path $testRoot 'normalized-slide-1.png'
    [System.IO.File]::WriteAllBytes($sourceImage, [byte[]](1, 2, 3))
    [System.IO.File]::Copy($sourceImage, $normalizedImage)
    [System.IO.File]::WriteAllText($manifestPath, ([ordered]@{ files = @([ordered]@{ input = $input }) } | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $manifestHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $imageHash = (Get-FileHash -LiteralPath $sourceImage -Algorithm SHA256).Hash.ToLowerInvariant()
    $packet = [ordered]@{
        protocolVersion = 1
        sourceManifest = [System.IO.Path]::GetFullPath($manifestPath)
        manifestSha256 = $manifestHash
        status = 'Ready'
        files = @([ordered]@{
            input = $input
            status = 'Ready'
            reason = ''
            pages = @([ordered]@{ slide = 1; sourceImage = $sourceImage; normalizedImage = $normalizedImage; sourceSha256 = $imageHash; normalizedSha256 = $imageHash })
        })
    }
    [System.IO.File]::WriteAllText($packetPath, ($packet | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $packetHash = (Get-FileHash -LiteralPath $packetPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $passedResult = [ordered]@{
        protocolVersion = 1; reviewVersion = 'fixture-1'; model = 'fixture-model'; reviewedAt = '2026-01-01T00:00:00Z'; status = 'Passed'
        packetPath = [System.IO.Path]::GetFullPath($packetPath); packetSha256 = $packetHash
        files = @([ordered]@{ input = $input; status = 'Passed'; pages = @([ordered]@{ slide = 1; sourceImage = $sourceImage; normalizedImage = $normalizedImage; sourceSha256 = $imageHash; normalizedSha256 = $imageHash; status = 'Passed'; evidence = 'Fixture pair inspected; no introduced visual issue.'; issues = @() }) })
    }
    [System.IO.File]::WriteAllText($resultPath, ($passedResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $gate = & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | ConvertFrom-Json
    if ($gate.status -ne 'Passed' -or $gate.deliveryBlocked) { throw 'AI review gate did not accept a fully paired passed result.' }

    # Applying a result adds the delivery gate to the manifest, so the packet
    # remains bound to the exact prepared manifest through the recorded hash.
    & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath | Out-Null
    $repeatGate = & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | ConvertFrom-Json
    if ($repeatGate.status -ne 'Passed' -or $repeatGate.deliveryBlocked) { throw 'AI review gate did not accept a repeated import after the manifest gained delivery metadata.' }

    # Windows paths are case-insensitive identities; equivalent casing must not
    # invalidate an otherwise hash-bound prepared packet.
    $passedResult.packetPath = [string]$packetPath.ToUpperInvariant()
    [System.IO.File]::WriteAllText($resultPath, ($passedResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $caseInsensitiveGate = & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | ConvertFrom-Json
    if ($caseInsensitiveGate.status -ne 'Passed' -or $caseInsensitiveGate.deliveryBlocked) { throw 'AI review gate rejected an equivalent Windows packet path with different casing.' }
    $passedResult.packetPath = [System.IO.Path]::GetFullPath($packetPath)

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

    # A result must not downgrade a packet that was prepared as Ready (full page
    # evidence) into ReviewUnavailable; unavailability comes from the packet only.
    $unavailableResult = [ordered]@{
        protocolVersion = 1; reviewVersion = 'fixture-1'; model = 'fixture-model'; reviewedAt = '2026-01-01T00:00:00Z'; status = 'ReviewUnavailable'; packetPath = [System.IO.Path]::GetFullPath($packetPath); packetSha256 = $packetHash
        files = @([ordered]@{ input = $input; status = 'ReviewUnavailable'; pages = @() })
    }
    [System.IO.File]::WriteAllText($resultPath, ($unavailableResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $rejectedUnavailable = $false
    try { & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | Out-Null } catch { $rejectedUnavailable = $true }
    if (-not $rejectedUnavailable) { throw 'AI review gate accepted an unbacked ReviewUnavailable claim against a Ready packet.' }

    # A rendered issue present in the original PPTX must remain a known baseline issue,
    # rather than becoming a false failure for the normalized copy. Shape names may
    # change across PowerPoint saves, so the same-page fallback key must also match.
    New-Item -ItemType Directory -Path $baselineAuditDir, $normalizedAuditDir -Force | Out-Null
    Add-Type -AssemblyName System.Drawing
    $fixtureImage = Join-Path $testRoot 'slide-001.png'
    $bitmap = New-Object System.Drawing.Bitmap 160, 90
    try { $bitmap.Save($fixtureImage, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $bitmap.Dispose() }
    $auditColumns = @('Timestamp', 'File', 'Slide', 'Shape', 'Issue', 'Severity', 'Details')
    $metricColumns = @('File', 'Slide', 'ImagePath', 'Width', 'Height', 'WhitePercent', 'NonWhitePercent', 'DarkPercent', 'IsVisuallyBlank', 'ContentTouchesEdge', 'ContentBounds', 'IssueCount', 'ExportStatus')
    $baselineAudit = [pscustomobject]@{ Timestamp = '2026-01-01 00:00:00'; File = 'source.pptx'; Slide = 1; Shape = 'TextBox 2'; Issue = 'ShapeOutOfSlideBounds'; Severity = 'Error'; Details = 'fixture baseline overflow' }
    $normalizedAudit = [pscustomobject]@{ Timestamp = '2026-01-01 00:00:00'; File = 'normalized.pptx'; Slide = 1; Shape = 'TextBox 9'; Issue = 'ShapeOutOfSlideBounds'; Severity = 'Error'; Details = 'fixture baseline overflow' }
    $metric = [pscustomobject]@{ File = 'fixture.pptx'; Slide = 1; ImagePath = $fixtureImage; Width = 160; Height = 90; WhitePercent = 100; NonWhitePercent = 0; DarkPercent = 0; IsVisuallyBlank = $false; ContentTouchesEdge = $false; ContentBounds = ''; IssueCount = 1; ExportStatus = 'Succeeded' }
    [System.IO.File]::WriteAllText((Join-Path $baselineAuditDir 'pptx-visual-audit.csv'), (($auditColumns -join ',') + "`r`n" + (($baselineAudit | Select-Object $auditColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n")), (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    [System.IO.File]::WriteAllText((Join-Path $normalizedAuditDir 'pptx-visual-audit.csv'), (($auditColumns -join ',') + "`r`n" + (($normalizedAudit | Select-Object $auditColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n")), (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    $metricCsv = (($metricColumns -join ',') + "`r`n" + (($metric | Select-Object $metricColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n"))
    [System.IO.File]::WriteAllText((Join-Path $baselineAuditDir 'pptx-slide-visual-metrics.csv'), $metricCsv, (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    [System.IO.File]::WriteAllText((Join-Path $normalizedAuditDir 'pptx-slide-visual-metrics.csv'), $metricCsv, (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    & (Join-Path $toolRoot 'Export-PptxVisualConfirmation.ps1') -VisualAuditDir $normalizedAuditDir -BaselineVisualAuditDir $baselineAuditDir | Out-Null
    $baselineConfirmation = Get-Content -LiteralPath (Join-Path $normalizedAuditDir 'pptx-visual-confirmation-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($baselineConfirmation.confirmationStatus -ne 'PassedWithKnownIssues' -or $baselineConfirmation.newErrorCount -ne 0) { throw 'Visual confirmation did not retain an original rendered issue as a known baseline issue.' }

    # One baseline issue may only cover one current issue. A duplicated current
    # error must remain new even when both rows use the shape-name fallback key.
    $duplicateAudit = @(
        $normalizedAudit
        [pscustomobject]@{ Timestamp = '2026-01-01 00:00:00'; File = 'normalized.pptx'; Slide = 1; Shape = 'TextBox 10'; Issue = 'ShapeOutOfSlideBounds'; Severity = 'Error'; Details = 'new duplicated overflow' }
    )
    [System.IO.File]::WriteAllText((Join-Path $normalizedAuditDir 'pptx-visual-audit.csv'), (($auditColumns -join ',') + "`r`n" + (($duplicateAudit | Select-Object $auditColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n")), (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    & (Join-Path $toolRoot 'Export-PptxVisualConfirmation.ps1') -VisualAuditDir $normalizedAuditDir -BaselineVisualAuditDir $baselineAuditDir | Out-Null
    $duplicateConfirmation = Get-Content -LiteralPath (Join-Path $normalizedAuditDir 'pptx-visual-confirmation-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($duplicateConfirmation.confirmationStatus -ne 'Failed' -or $duplicateConfirmation.newErrorCount -ne 1) { throw 'Visual confirmation allowed a duplicated fallback issue to consume the same baseline row twice.' }

    # An audit error for a page that has no metrics row means that page was not
    # rendered/covered. It must not disappear from the confirmation gate.
    $missingPageAudit = @(
        $normalizedAudit
        [pscustomobject]@{ Timestamp = '2026-01-01 00:00:00'; File = 'normalized.pptx'; Slide = 2; Shape = '(slide)'; Issue = 'SlidePngExportFailed'; Severity = 'Error'; Details = 'fixture export failure' }
    )
    [System.IO.File]::WriteAllText((Join-Path $normalizedAuditDir 'pptx-visual-audit.csv'), (($auditColumns -join ',') + "`r`n" + (($missingPageAudit | Select-Object $auditColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n")), (New-Object System.Text.UTF8Encoding -ArgumentList $true))
    & (Join-Path $toolRoot 'Export-PptxVisualConfirmation.ps1') -VisualAuditDir $normalizedAuditDir -BaselineVisualAuditDir $baselineAuditDir | Out-Null
    $missingPageConfirmation = Get-Content -LiteralPath (Join-Path $normalizedAuditDir 'pptx-visual-confirmation-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($missingPageConfirmation.confirmationStatus -ne 'Failed' -or $missingPageConfirmation.automationGateStatus -ne 'Failed' -or $missingPageConfirmation.unrenderedAuditIssueSlides -notcontains 2) { throw 'Visual confirmation did not fail when an audit issue had no matching metrics row.' }

    # An empty metrics file is not valid zero-page coverage. It must produce a
    # deterministic failed manifest instead of a strict-mode empty-sum error.
    $emptyMetricsDir = Join-Path $testRoot 'empty-metrics-audit'
    New-Item -ItemType Directory -Path $emptyMetricsDir -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $emptyMetricsDir 'pptx-visual-audit.csv'), ($auditColumns -join ','), (New-Object System.Text.UTF8Encoding($true)))
    [System.IO.File]::WriteAllText((Join-Path $emptyMetricsDir 'pptx-slide-visual-metrics.csv'), ($metricColumns -join ','), (New-Object System.Text.UTF8Encoding($true)))
    & (Join-Path $toolRoot 'Export-PptxVisualConfirmation.ps1') -VisualAuditDir $emptyMetricsDir | Out-Null
    $emptyMetricsConfirmation = Get-Content -LiteralPath (Join-Path $emptyMetricsDir 'pptx-visual-confirmation-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($emptyMetricsConfirmation.confirmationStatus -ne 'Failed' -or $emptyMetricsConfirmation.automationGateStatus -ne 'Failed' -or -not $emptyMetricsConfirmation.noRenderedMetrics) { throw 'Visual confirmation did not fail on an empty metrics file.' }

    # A zero-byte page image must not be accepted merely because its path exists.
    $zeroImageDir = Join-Path $testRoot 'zero-image-audit'
    New-Item -ItemType Directory -Path $zeroImageDir -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $zeroImageDir 'pptx-visual-audit.csv'), ($auditColumns -join ','), (New-Object System.Text.UTF8Encoding($true)))
    $zeroImagePath = Join-Path $zeroImageDir 'slide-001.png'
    [System.IO.File]::WriteAllBytes($zeroImagePath, [byte[]]@())
    $zeroMetric = [pscustomobject]@{ File = 'fixture.pptx'; Slide = 1; ImagePath = $zeroImagePath; Width = 160; Height = 90; WhitePercent = 100; NonWhitePercent = 0; DarkPercent = 0; IsVisuallyBlank = $false; ContentTouchesEdge = $false; ContentBounds = ''; IssueCount = 0; ExportStatus = 'Succeeded' }
    [System.IO.File]::WriteAllText((Join-Path $zeroImageDir 'pptx-slide-visual-metrics.csv'), (($metricColumns -join ',') + "`r`n" + (($zeroMetric | Select-Object $metricColumns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -join "`r`n")), (New-Object System.Text.UTF8Encoding($true)))
    & (Join-Path $toolRoot 'Export-PptxVisualConfirmation.ps1') -VisualAuditDir $zeroImageDir | Out-Null
    $zeroImageConfirmation = Get-Content -LiteralPath (Join-Path $zeroImageDir 'pptx-visual-confirmation-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($zeroImageConfirmation.confirmationStatus -ne 'Failed' -or $zeroImageConfirmation.automationGateStatus -ne 'Failed') { throw 'Visual confirmation accepted a zero-byte page image.' }
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}

Write-Host 'Physics PPT policy fixture tests passed.'
