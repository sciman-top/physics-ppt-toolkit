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
    . (Join-Path $toolRoot 'PhysicsPpt.Common.ps1')
    foreach ($counts in @(@(0,0), @(1,0), @(0,1))) {
        $mockApplication = [pscustomobject]@{ Presentations = @{ Count = $counts[0] }; ProtectedViewWindows = @{ Count = $counts[1] }; quitCalled = $false }
        $mockApplication | Add-Member ScriptMethod Quit { $this.quitCalled = $true }
        Stop-PowerPointApplicationIfIdle $mockApplication
        if ($mockApplication.quitCalled -ne ($counts[0] -eq 0 -and $counts[1] -eq 0)) { throw 'PowerPoint cleanup closed an occupied instance or failed to close an idle instance.' }
    }
    $batchAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $toolRoot 'Invoke-FormulaOleBatch.ps1'), [ref]$null, [ref]$null)
    foreach ($name in @('Get-FileSha256', 'New-StepSignature', 'Test-StepReusable', 'Get-RecordedStep', 'Invoke-BatchStep')) {
        $node = $batchAst.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true) | Select-Object -First 1
        Invoke-Expression $node.Extent.Text
    }
    $script:SignatureToolHash = 'tools'; $script:SignatureConfigHash = 'config'; $MaxItems = 20; $Resume = $false
    $batchSteps = New-Object System.Collections.Generic.List[object]
    $artifact = Join-Path $testRoot 'batch-artifact.txt'
    Invoke-BatchStep -StepName 'probe' -Outputs @($artifact) -Action { [IO.File]::WriteAllText($artifact, 'original') } | Out-Null
    if (-not (Test-StepReusable $batchSteps[0])) { throw 'Unchanged batch artifact was not reusable.' }
    $Resume = $true
    $currentInput = Join-Path $testRoot 'current-input.txt'
    [IO.File]::WriteAllText($currentInput, 'new input')
    Invoke-BatchStep -StepName 'probe' -Inputs @($currentInput) -Outputs @($artifact) -Action { [IO.File]::WriteAllText($artifact, 'current') } | Out-Null
    if ([IO.File]::ReadAllText($artifact) -cne 'current') { throw 'Resume reused a step bound to a different input.' }
    $Resume = $false
    [IO.File]::WriteAllText($artifact, 'tampered')
    if (Test-StepReusable $batchSteps[0]) { throw 'Modified batch artifact was reused.' }
    $directoryArtifact = Join-Path $testRoot 'batch-directory'
    New-Item -ItemType Directory -Path $directoryArtifact | Out-Null
    Invoke-BatchStep -StepName 'directory' -Outputs @($directoryArtifact) -Action { [IO.File]::WriteAllText((Join-Path $directoryArtifact 'fragment.xml'), 'original') } | Out-Null
    if (-not (Test-StepReusable $batchSteps[1])) { throw 'Directory output was not reusable.' }
    [IO.File]::WriteAllText((Join-Path $directoryArtifact 'fragment.xml'), 'tampered')
    if (Test-StepReusable $batchSteps[1]) { throw 'Tampered directory output was reused.' }
    $missingOutputRejected = $false
    try { Invoke-BatchStep -StepName 'missing' -Outputs @((Join-Path $testRoot 'missing.txt')) -Action { } | Out-Null } catch { $missingOutputRejected = $true }
    if (-not $missingOutputRejected -or $batchSteps.Count -ne 2) { throw 'Missing batch output was recorded as Passed.' }
    $snapshotAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $toolRoot 'Export-PptxInvariantSnapshot.ps1'), [ref]$null, [ref]$null)
    Invoke-Expression ($snapshotAst.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-RunFormatFingerprint' }, $true) | Select-Object -First 1).Extent.Text
    function New-FormatProbeShape {
        param([string[]]$Texts, [int[]]$Colors)
        $runs = @()
        for ($i = 0; $i -lt $Texts.Count; $i++) {
            $runs += [pscustomobject]@{ Text = $Texts[$i]; Font = @{ Bold = 0; Fill = @{ ForeColor = @{ RGB = $Colors[$i] } } } }
        }
        $line = [pscustomobject]@{ items = $runs }
        $line | Add-Member ScriptMethod Runs { return $this.items }
        $range = [pscustomobject]@{ line = $line }
        $range | Add-Member ScriptMethod Lines { return $this.line }
        return [pscustomobject]@{ TextFrame2 = @{ TextRange = $range } }
    }
    $readFailed = $false
    $normalEmphasis = Get-RunFormatFingerprint (New-FormatProbeShape @('ab', 'cd') @(255, 0)) ([ref]$readFailed)
    $splitEmphasis = Get-RunFormatFingerprint (New-FormatProbeShape @('a', 'b', 'c', 'd') @(255, 255, 0, 0)) ([ref]$readFailed)
    $swappedEmphasis = Get-RunFormatFingerprint (New-FormatProbeShape @('ab', 'cd') @(0, 255)) ([ref]$readFailed)
    if ($readFailed -or $normalEmphasis -cne $splitEmphasis -or $normalEmphasis -ceq $swappedEmphasis) { throw 'Emphasis fingerprint lost positions or depended on run splitting.' }
    $workflowAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $toolRoot 'Invoke-PhysicsPptWorkflow.ps1'), [ref]$null, [ref]$null)
    $propertyFunction = $workflowAst.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-ObjectPropertyValue' }, $true) | Select-Object -First 1
    Invoke-Expression $propertyFunction.Extent.Text
    foreach ($propertyObject in @(@{ effectiveMode = 'ReportOnly' }, [ordered]@{ effectiveMode = 'ReportOnly' }, [pscustomobject]@{ effectiveMode = 'ReportOnly' })) {
        if ((Get-ObjectPropertyValue $propertyObject 'effectiveMode' 'Unknown') -ne 'ReportOnly') { throw 'Workflow metadata lost a dictionary property before JSON serialization.' }
        if ((Get-ObjectPropertyValue $propertyObject 'absent' 'fallback') -ne 'fallback') { throw 'Workflow metadata did not honor the missing-property default.' }
    }
    $beforePath = Join-Path $testRoot 'before.json'
    $afterPath = Join-Path $testRoot 'after.json'
    $blockedPath = Join-Path $testRoot 'blocked.json'
    $comparePath = Join-Path $testRoot 'compare.json'
    $formulaBeforePath = Join-Path $testRoot 'formula-before.json'
    $formulaAfterPath = Join-Path $testRoot 'formula-after.json'
    $formulaGeometryPath = Join-Path $testRoot 'formula-geometry.json'
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

    # Case changes can change physical quantities; missing slide rows must not
    # disappear behind an unchanged declared slideCount.
    foreach ($mutation in @('TextCase', 'MissingSlide', 'ExternalTarget')) {
        $probeBefore = $before | ConvertFrom-Json
        $probeAfter = $before | ConvertFrom-Json
        switch ($mutation) {
            'TextCase' { $probeAfter.slides[0].shapes[0].text = '速度 V=s/t' }
            'MissingSlide' { $probeAfter.slides = @() }
            'ExternalTarget' {
                $probeBefore.package.relationships[0].targetMode = 'External'
                $probeBefore.package.relationships[0].target = 'https://first.example/image1.png'
                $probeAfter.package.relationships[0].targetMode = 'External'
                $probeAfter.package.relationships[0].target = 'https://second.example/image1.png'
            }
        }
        $probeBefore | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $formulaBeforePath -Encoding UTF8
        $probeAfter | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $blockedPath -Encoding UTF8
        $probeResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $formulaBeforePath -AfterPath $blockedPath | ConvertFrom-Json
        if ($probeResult.passed) { throw "Invariant comparison accepted $mutation." }
    }
    $mediaRenameBefore = $before | ConvertFrom-Json
    $mediaRenameBefore.package.relationships[0].part = 'ppt/slides/_rels/slide1.xml.rels'
    $mediaRenameBefore.package.relationships[0].type = 'image'
    $mediaRenameBefore.package.relationships[0].target = '../media/image1.png'
    $mediaRenameAfter = $mediaRenameBefore | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $mediaRenameAfter.package.media[0].part = 'ppt/media/image2.png'
    $mediaRenameAfter.package.relationships[0].target = '../media/image2.png'
    $mediaRenameBefore | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $formulaBeforePath -Encoding UTF8
    $mediaRenameAfter | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $blockedPath -Encoding UTF8
    $mediaRenameResult = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $formulaBeforePath -AfterPath $blockedPath | ConvertFrom-Json
    if (-not $mediaRenameResult.passed) { throw 'Unchanged media content with renamed package parts was blocked.' }

    # Formula conversion has an explicit, narrow allowance: only a type-7 OLE
    # with empty COM text may become a generated FormulaOmml type-1 shape. The
    # allowance must not hide geometry changes or become a default bypass.
    $formulaBeforeObject = $baseSnapshot | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $formulaBeforeShape = $formulaBeforeObject.slides[0].shapes[0]
    $formulaBeforeShape.name = 'Object 1'
    $formulaBeforeShape.type = 7
    $formulaBeforeShape.autoSize = 0
    $formulaBeforeShape.wordWrap = 0
    $formulaBeforeShape.text = ''
    [System.IO.File]::WriteAllText($formulaBeforePath, ($formulaBeforeObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $formulaAfterObject = $formulaBeforeObject | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $formulaAfterShape = $formulaAfterObject.slides[0].shapes[0]
    $formulaAfterShape.name = 'FormulaOmml 测试'
    $formulaAfterShape.type = 1
    $formulaAfterShape.autoSize = 2
    $formulaAfterShape.wordWrap = -1
    $formulaAfterShape.text = ' Q 吸 '
    [System.IO.File]::WriteAllText($formulaAfterPath, ($formulaAfterObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $formulaAllowed = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $formulaBeforePath -AfterPath $formulaAfterPath -AllowReviewedFormulaConversion | ConvertFrom-Json
    if (-not $formulaAllowed.passed -or $formulaAllowed.blockerCount -ne 0 -or $formulaAllowed.allowedChangeCount -ne 5) { throw 'Reviewed formula conversion allowance did not accept only the expected carrier-property changes.' }
    $formulaDefault = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $formulaBeforePath -AfterPath $formulaAfterPath | ConvertFrom-Json
    if ($formulaDefault.passed -or $formulaDefault.blockerCount -lt 1 -or $formulaDefault.allowedChangeCount -ne 0) { throw 'Formula conversion was allowed without explicit authorization.' }
    $formulaGeometryObject = $formulaAfterObject | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $formulaGeometryObject.slides[0].shapes[0].width = 201
    [System.IO.File]::WriteAllText($formulaGeometryPath, ($formulaGeometryObject | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $formulaGeometry = & (Join-Path $toolRoot 'Compare-PptxInvariantSnapshot.ps1') -BeforePath $formulaBeforePath -AfterPath $formulaGeometryPath -AllowReviewedFormulaConversion | ConvertFrom-Json
    if ($formulaGeometry.passed -or $formulaGeometry.blockerCount -lt 1) { throw 'Reviewed formula conversion allowance hid a geometry change.' }

    $input = 'C:\fixture\lesson.pptx'
    $sourceImage = Join-Path $testRoot 'source-slide-1.png'
    $normalizedImage = Join-Path $testRoot 'normalized-slide-1.png'
    [System.IO.File]::WriteAllBytes($sourceImage, [byte[]](1, 2, 3))
    [System.IO.File]::Copy($sourceImage, $normalizedImage)
    $deliveredArtifact = Join-Path $testRoot 'delivered.pptx'
    [IO.File]::WriteAllBytes($deliveredArtifact, [byte[]](1, 2, 3))
    $deliveredHash = (Get-FileHash -LiteralPath $deliveredArtifact -Algorithm SHA256).Hash.ToLowerInvariant()
    [System.IO.File]::WriteAllText($manifestPath, ([ordered]@{ files = @([ordered]@{ input = $input; normalizedPptx = $deliveredArtifact; normalizedPptxSha256 = $deliveredHash }) } | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
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

    foreach ($bindingField in @('packetPath', 'packetSha256', 'sourceSha256', 'normalizedSha256')) {
        $bindingProbe = $passedResult | ConvertTo-Json -Depth 12 | ConvertFrom-Json
        if ($bindingField -in @('packetPath', 'packetSha256')) {
            $bindingProbe.PSObject.Properties.Remove($bindingField)
        } else {
            $bindingProbe.files[0].pages[0].PSObject.Properties.Remove($bindingField)
        }
        [System.IO.File]::WriteAllText($resultPath, ($bindingProbe | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
        $rejectedBindingField = $false
        try {
            & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | Out-Null
        } catch {
            $rejectedBindingField = $_.Exception.Message -like "*missing required string '$bindingField'*"
        }
        if (-not $rejectedBindingField) { throw "AI review importer did not specifically reject missing binding field '$bindingField'." }
    }
    foreach ($arrayField in @('files', 'pages', 'issues')) {
        $arrayProbe = $passedResult | ConvertTo-Json -Depth 12 | ConvertFrom-Json
        switch ($arrayField) {
            'files' { $arrayProbe.PSObject.Properties.Remove('files') }
            'pages' { $arrayProbe.files[0].PSObject.Properties.Remove('pages') }
            'issues' { $arrayProbe.files[0].pages[0].PSObject.Properties.Remove('issues') }
        }
        [System.IO.File]::WriteAllText($resultPath, ($arrayProbe | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
        $rejectedArrayField = $false
        try {
            & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | Out-Null
        } catch {
            $rejectedArrayField = $_.Exception.Message -like "*missing required array '$arrayField'*"
        }
        if (-not $rejectedArrayField) { throw "AI review importer did not specifically reject missing array '$arrayField'." }
    }
    [System.IO.File]::WriteAllText($resultPath, ($passedResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))

    # Applying a result adds the delivery gate to the manifest, so the packet
    # remains bound to the exact prepared manifest through the recorded hash.
    & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath | Out-Null
    $repeatGate = & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | ConvertFrom-Json
    if ($repeatGate.status -ne 'Passed' -or $repeatGate.deliveryBlocked) { throw 'AI review gate did not accept a repeated import after the manifest gained delivery metadata.' }

    $savedManifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8
    foreach ($priorStatus in @('BlockedDeliveryArtifacts', 'Pending')) {
        $priorManifest = $savedManifest | ConvertFrom-Json
        $priorManifest | Add-Member -NotePropertyName deliveryFailures -NotePropertyValue @('PDF page count mismatch') -Force
        $priorManifest.deliveryStatus = $priorStatus
        $priorManifest.deliveryBlocked = ($priorStatus -eq 'BlockedDeliveryArtifacts')
        $priorManifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
        & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath | Out-Null
        $imported = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $imported.deliveryBlocked -or $imported.deliveryStatus -ne 'BlockedDeliveryArtifacts') { throw 'AI Passed erased a required delivery artifact failure.' }
    }
    [IO.File]::WriteAllText($manifestPath, $savedManifest, (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllBytes($deliveredArtifact, [byte[]](4, 5, 6))
    $staleRejected = $false
    try { & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath -ValidateOnly | Out-Null } catch { $staleRejected = $_.Exception.Message -like '*stale or replaced normalizedPptx*' }
    if (-not $staleRejected) { throw 'AI review accepted a replaced delivery presentation with unchanged page images.' }
    [IO.File]::WriteAllBytes($deliveredArtifact, [byte[]](1, 2, 3))

    $priorManifest = $savedManifest | ConvertFrom-Json
    $priorManifest | Add-Member -NotePropertyName deliveryReviews -NotePropertyValue @('Visual confirmation requires review') -Force
    $priorManifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    & (Join-Path $toolRoot 'Import-PptxAiReviewResult.ps1') -ManifestPath $manifestPath -ResultPath $resultPath -PacketPath $packetPath | Out-Null
    if ((Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json).deliveryStatus -ne 'Pending') { throw 'AI Passed erased an outstanding visual review.' }
    [IO.File]::WriteAllText($manifestPath, $savedManifest, (New-Object Text.UTF8Encoding($false)))

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
    $packetManifest = Join-Path $testRoot 'packet-manifest.json'
    $standalonePacket = Join-Path $testRoot 'standalone-packet.json'
    $sourcePageDir = Join-Path $testRoot 'source-pages'
    $normalizedPageDir = Join-Path $testRoot 'normalized-pages'
    New-Item -ItemType Directory -Path $sourcePageDir, $normalizedPageDir | Out-Null
    Copy-Item -LiteralPath $fixtureImage -Destination (Join-Path $sourcePageDir 'page-1.png')
    Copy-Item -LiteralPath $fixtureImage -Destination (Join-Path $normalizedPageDir 'page-1.png')
    @{ files = @(@{ input = $input; sourcePageImages = $sourcePageDir; pageImages = $normalizedPageDir; validation = @{ expectedSlides = 1 } }) } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $packetManifest -Encoding UTF8
    $testHostPath = (Get-Process -Id $PID).Path
    & $testHostPath -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $toolRoot 'Export-PptxAiReviewPacket.ps1') -ManifestPath $packetManifest -OutputPath $standalonePacket | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Standalone AI packet export failed in a clean PowerShell host.' }
    $standalone = Get-Content -LiteralPath $standalonePacket -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($standalone.status -ne 'Ready' -or $standalone.files[0].pages.Count -ne 1) { throw 'Standalone AI packet did not prepare the complete page pair.' }
    foreach ($collision in @('Packet', 'Snapshot', 'Bold', 'Whitelist', 'Ole')) {
        $rejectedCollision = $false
        try {
            switch ($collision) {
                'Packet' { & (Join-Path $toolRoot 'Export-PptxAiReviewPacket.ps1') -ManifestPath $packetManifest -OutputPath $packetManifest | Out-Null }
                'Snapshot' { & (Join-Path $toolRoot 'Export-PptxInvariantSnapshot.ps1') -InputPath $deliveredArtifact -OutputPath $deliveredArtifact | Out-Null }
                'Bold' { & (Join-Path $toolRoot 'Set-PptxTextBold.ps1') -InputPath $deliveredArtifact -OutputPath (Join-Path $testRoot 'bold.pptx') -ReportPath $deliveredArtifact | Out-Null }
                'Whitelist' { & (Join-Path $toolRoot 'Apply-FormulaOmmlWhitelist.ps1') -InputPath $deliveredArtifact -OutputPath (Join-Path $testRoot 'omml.pptx') -FormulaReviewCsv $resultPath -OmmlCandidateCsv $resultPath -ReportPath $deliveredArtifact | Out-Null }
                'Ole' { & (Join-Path $toolRoot 'Apply-FormulaOmmlForOle.ps1') -InputPath $deliveredArtifact -OutputPath (Join-Path $testRoot 'ole.pptx') -MappingCsv $resultPath -OmmlCandidateCsv $resultPath -ReportPath $deliveredArtifact | Out-Null }
            }
        } catch { $rejectedCollision = $_.Exception.Message -like '*must differ*' }
        if (-not $rejectedCollision) { throw "Source-path collision was not rejected: $collision" }
    }
    if ((Get-FileHash -LiteralPath $deliveredArtifact).Hash.ToLowerInvariant() -ne $deliveredHash) { throw 'Source-path collision probes damaged the input.' }
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
    # A dependency-only edit must invalidate normalization reuse, even with
    # identical source/config/main script and invocation parameters.
    $normalizeAstForSignature = [Management.Automation.Language.Parser]::ParseFile((Join-Path $toolRoot 'Normalize-PhysicsPpt.ps1'), [ref]$null, [ref]$null)
    $signatureFunction = $normalizeAstForSignature.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-NormalizeSignature' }, $true) | Select-Object -First 1
    $signatureProbe = Join-Path $testRoot 'signature-probe.ps1'
    $probeCommon = Join-Path $testRoot 'PhysicsPpt.Common.ps1'
    $probeGeometry = Join-Path $testRoot 'PhysicsPpt.Geometry.ps1'
    Copy-Item -LiteralPath (Join-Path $toolRoot 'PhysicsPpt.Common.ps1') -Destination $probeCommon
    Copy-Item -LiteralPath (Join-Path $toolRoot 'PhysicsPpt.Geometry.ps1') -Destination $probeGeometry
    $signatureProbeCode = @'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')
$script:NormalizeSignatureConstantHashes = $null
$script:NormalizeScriptPath = $PSCommandPath
$ConfigPath = Join-Path $PSScriptRoot 'packet-manifest.json'
$NoPdf = $false; $ReportOnly = $false; $DisableAdvanceOnClick = $false; $FilePattern = '*.ppt*'; $ImageOutputDir = ''
'@ + "`r`n" + $signatureFunction.Extent.Text + "`r`n" + 'Get-NormalizeSignature -File (Get-Item -LiteralPath (Join-Path $PSScriptRoot ''delivered.pptx'')) -SafeName ''probe'' | ConvertTo-Json -Depth 8'
    [IO.File]::WriteAllText($signatureProbe, $signatureProbeCode, (New-Object Text.UTF8Encoding($true)))
    $initialSignature = & $testHostPath -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $signatureProbe | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw 'Initial normalization signature probe failed.' }
    foreach ($dependency in @($probeCommon, $probeGeometry)) {
        Add-Content -LiteralPath $dependency -Value '# dependency revision probe' -Encoding UTF8
        $changedSignature = & $testHostPath -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $signatureProbe | ConvertFrom-Json
        if ($LASTEXITCODE -ne 0 -or $changedSignature.Signature -eq $initialSignature.Signature) { throw 'A module-only edit reused the previous normalization signature.' }
        $initialSignature = $changedSignature
    }

    # Run the actual shape-enumeration/recovery statements with an enumerable
    # whose first three COM yields are dead, while indexed access still works.
    . (Join-Path $toolRoot 'PhysicsPpt.Common.ps1')
    . (Join-Path $toolRoot 'PhysicsPpt.ReviewArtifacts.ps1')
    if (-not (Test-IsFinalFailureIssue 'ShapeEnumerationCorrupted')) { throw 'An unrecovered enumeration gap is not a final delivery failure.' }
    $normalizeAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $toolRoot 'Normalize-PhysicsPpt.ps1'), [ref]$null, [ref]$null)
    $startNode = $normalizeAst.FindAll({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$processedShapeIds' }, $true) | Select-Object -First 1
    $recoveryNode = $normalizeAst.FindAll({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$enumerationBroken' }, $true) | Select-Object -First 1
    if ($null -eq $startNode -or $null -eq $recoveryNode) { throw 'Shape recovery statements could not be located.' }
    $normalizeText = Get-Content -LiteralPath (Join-Path $toolRoot 'Normalize-PhysicsPpt.ps1') -Raw -Encoding UTF8
    $recoveryCode = $normalizeText.Substring($startNode.Extent.StartOffset, $recoveryNode.Extent.EndOffset - $startNode.Extent.StartOffset)
    if (-not ('ToolkitBrokenShapeEnumerable' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections;
public class ToolkitBrokenShapeEnumerable : IEnumerable {
    public object[] Items;
    public bool CountFails;
    public int Count { get { if (CountFails) throw new Exception("Count unavailable"); return Items.Length; } }
    public object Item(int index) { return Items[index - 1]; }
    public IEnumerator GetEnumerator() { return new object[] { null, null, null }.GetEnumerator(); }
}
'@
    }
    & {
        param($Code)
        function Invoke-NormalizeSlideShape { param($Shape, $SlideNumber, $FileName, $ReportOnly, $PictureFillShapeIds, $PreserveSlideStyle, $IsVideo, $IsSectionTitle, $SlideWidth) $script:RecoveredShapeIds.Add([int]$Shape.Id) }
        function Add-ReportRow { param($File, $SlideNumber, $ShapeName, $Issue, $Details, $RiskLevel, $Result) $script:RecoveryIssues.Add($Issue) }
        foreach ($countFails in @($false, $true)) {
            $script:RecoveredShapeIds = New-Object 'Collections.Generic.List[int]'
            $script:RecoveryIssues = New-Object 'Collections.Generic.List[string]'
            $collection = New-Object ToolkitBrokenShapeEnumerable
            $collection.Items = @([pscustomobject]@{ Id = 1 }, [pscustomobject]@{ Id = 2 }, [pscustomobject]@{ Id = 3 })
            $collection.CountFails = $countFails
            $slide = [pscustomobject]@{ Shapes = $collection }
            $file = [pscustomobject]@{ Name = 'probe.pptx' }
            $i = 1; $ReportOnly = [switch]$true; $pictureFillShapeIds = @{}; $preserveSlideStyle = $false; $isVideo = $false; $isSectionTitle = $false; $slideWidth = 960
            Invoke-Expression $Code
            if (-not $countFails -and ($script:RecoveredShapeIds.Count -ne 3 -or $script:RecoveryIssues -notcontains 'ShapeEnumerationRecovered')) { throw 'Dead first COM yields prevented indexed shape recovery.' }
            if ($countFails -and $script:RecoveryIssues -notcontains 'ShapeEnumerationCorrupted') { throw 'An unreadable shape count was falsely reported recovered.' }
        }
    } $recoveryCode

    Write-Host 'Physics PPT policy fixture tests passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
