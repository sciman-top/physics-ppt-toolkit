<#
.SYNOPSIS
  Orchestrate one explicit MathType/OLE → OfficeMath migration batch.

.DESCRIPTION
  Chains the OLE pipeline end to end into a versioned delivery directory
  (reports/<stem>_v<N>): carrier inventory (with optional page-kind annotation)
  -> rendered pages -> OLE crops -> GoldSet mapping -> OMML candidates -> apply
  (new copy, mc:Fallback preserved) -> size fit (PowerPoint-measured overflow
  guard, shrinks FormulaOmml shapes back into the original OLE box) -> Open XML
  validator -> visual audit. Any step failure stops the batch and records the
  completed prefix in the batch manifest. The source PPTX is never modified.

  This is an explicit migration entry point. The default workflow does not call
  it, and config formulaProcessing switches remain untouched.

.PARAMETER InputPath
  Chain-head .pptx copy to convert. The file is copied, never edited.

.PARAMETER GoldSetCsv
  Approved GoldSet CSV (ReviewStatus=Approved rows); see docs/公式GoldSet编制SOP.md.

.PARAMETER PagesDir
  Optional pre-rendered pages directory (slide-NNN.png). When omitted the batch
  renders pages via Export-PptxVisualAudit.

.PARAMETER DeliveryRoot
  Optional delivery directory override. Must stay outside reports/ (reports/ is
  reserved for auto-versioned <deck>_v<N> deliveries).

.PARAMETER MaxItems
  Maximum mapping rows applied. Defaults to 20.

.PARAMETER SkipVisualAudit
  Skip before/after page rendering. Requires -PagesDir for the crop step.

.PARAMETER Resume
  Reuse recorded steps from an existing batch manifest when input hashes and
  outputs still match.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$GoldSetCsv,

    [string]$PagesDir,

    [string]$DeliveryRoot,

    [ValidateRange(1, 100)]
    [int]$MaxItems = 20,

    [switch]$SkipVisualAudit,

    [switch]$Resume
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$inputFullPath = [System.IO.Path]::GetFullPath($InputPath)
$goldSetPath = [System.IO.Path]::GetFullPath($GoldSetCsv)
foreach ($path in @($inputFullPath, $goldSetPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required input not found: $path" }
}
if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Export-FormulaOleCrops.ps1'))) {
    throw 'Export-FormulaOleCrops.ps1 is missing; the batch requires the full OLE pipeline.'
}

function Get-FileSha256 {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return '' }
    # Directories (e.g. the pre-rendered -PagesDir tree) are hashed as a
    # deterministic sorted manifest of relative path + file hash, so any
    # content change still changes the step signature.
    if (Test-Path -LiteralPath $Path -PathType Container) {
        $dirFullPath = [System.IO.Path]::GetFullPath($Path)
        $entries = @(Get-ChildItem -LiteralPath $dirFullPath -Recurse -File | Sort-Object FullName | ForEach-Object {
            $relative = $_.FullName.Substring($dirFullPath.Length).TrimStart('\', '/')
            '{0}:{1}' -f $relative, (Get-FileSha256Hex -Path $_.FullName)
        })
        return (Get-TextSha256Hex -Text ($entries -join "`n"))
    }
    return (Get-FileSha256Hex -Path $Path)
}

function Get-BatchManifestPath {
    param([string]$Root)
    return (Join-Path $Root '00_检查报告\formula-ole-batch-manifest.json')
}

function New-StepSignature {
    # Tool-script, config and batch-parameter hashes are part of the signature:
    # a code/config/parameter change must invalidate recorded steps, otherwise
    # -Resume replays stale artifacts (a tokenizer fix would be ignored).
    param([string]$StepName, [string[]]$InputPaths, [string[]]$Flags)
    $material = @($StepName) + @($Flags) + @($InputPaths | ForEach-Object { Get-FileSha256 -Path $_ }) + @(
        $script:SignatureToolHash, $script:SignatureConfigHash, [string]$MaxItems
    )
    return (Get-TextSha256Hex -Text ($material -join '|'))
}

function Test-StepReusable {
    param($Step)
    if ($null -eq $Step) { return $false }
    if ([string]$Step.status -ne 'Passed') { return $false }
    if ($null -eq $Step.PSObject.Properties['outputHashes']) { return $false }
    if (@($Step.outputHashes).Count -ne @($Step.outputs).Count) { return $false }
    $outputIndex = 0
    foreach ($output in @($Step.outputs)) {
        if ([string]::IsNullOrWhiteSpace([string]$output) -or -not (Test-Path -LiteralPath ([string]$output))) { return $false }
        if ((Get-FileSha256 -Path $output) -cne [string]$Step.outputHashes[$outputIndex]) { return $false }
        $outputIndex++
    }
    return ((New-StepSignature -StepName ([string]$Step.name) -InputPaths @([string[]]$Step.inputs) -Flags @([string[]]$step.flags)) -ceq [string]$Step.signature)
}

# --- signature inputs: child tool scripts + style config (hashed once) ---
$signatureToolFiles = @(
    'Export-FormulaCarrierInventory.ps1', 'Export-PptxVisualAudit.ps1', 'Export-FormulaOleCrops.ps1',
    'Export-FormulaOleMapping.ps1', 'Export-FormulaOmmlCandidates.ps1', 'Apply-FormulaOmmlForOle.ps1',
    'PhysicsPpt.Common.ps1', 'PhysicsPpt.Formula.ps1', 'Invoke-FormulaOleBatch.ps1',
    'FormulaOfficeMathValidator/Program.cs', 'FormulaOfficeMathValidator/FormulaOfficeMathValidator.csproj'
)
$script:SignatureToolHash = $null
$toolHashMaterial = @()
foreach ($toolFile in $signatureToolFiles) { $toolHashMaterial += (Get-FileSha256 -Path (Join-Path $PSScriptRoot $toolFile)) }
$toolHashSha = [System.Security.Cryptography.SHA256]::Create()
try {
    $toolHashBytes = [System.Text.Encoding]::UTF8.GetBytes(($toolHashMaterial -join '|'))
    $script:SignatureToolHash = ([System.BitConverter]::ToString($toolHashSha.ComputeHash($toolHashBytes)) -replace '-', '').ToLowerInvariant()
} finally { $toolHashSha.Dispose() }
$styleConfigPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'config\physics-ppt-style.config.json'
$script:SignatureConfigHash = if (Test-Path -LiteralPath $styleConfigPath) { Get-FileSha256 -Path $styleConfigPath } else { 'no-config' }

# --- rendered size fit -------------------------------------------------------
# OMML display fractions are laid out by the renderer at roughly 2.24x the
# base font size (PowerPoint TextRange.BoundHeight measurement, consistent
# across sizes), while the source MathType bitmaps pack the same glyphs far
# tighter. A GoldSet SizePt that matches the original bitmap glyphs can
# therefore overflow the original OLE box; the applier centers the text in
# that box, so the overflow bleeds symmetrically and clips wherever the box
# touches a slide edge (16.3 slide17 shipped with the numerator cut off at
# the slide top). The batch measures the applied copy in PowerPoint and
# shrinks overflowing shapes back into their box; the decision itself is a
# pure function so the gate can probe it.

$script:NsDrawingMain = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$script:NsPresentationMain = 'http://schemas.openxmlformats.org/presentationml/2006/main'

function Get-OmmlFormulaFitSizePt {
    # Pure decision: half-point stepped corrected base size that brings the
    # PowerPoint-measured formula text bound inside its shape box (+tolerance).
    # $null means no fitted size exists (already fits, below the 8pt mapping
    # floor, or not smaller than current) and the caller must fail the shape
    # loudly instead of writing a clipped formula.
    param(
        [double]$CurrentSizePt,
        [double]$TextBoundHeightPt,
        [double]$BoxHeightPt,
        [double]$TolerancePt = 2.0
    )
    if ($CurrentSizePt -le 0 -or $TextBoundHeightPt -le 0 -or $BoxHeightPt -le 0) { return $null }
    $scale = (($BoxHeightPt + $TolerancePt) / $TextBoundHeightPt) * 0.98
    if ($scale -ge 1.0) { return $null }
    $fitted = [Math]::Floor(($CurrentSizePt * $scale) * 2.0) / 2.0
    if ($fitted -lt 8 -or $fitted -ge $CurrentSizePt) { return $null }
    return $fitted
}

function Measure-FormulaOmmlRenderBounds {
    # PowerPoint-measured laid-out text bounds for every FormulaOmml shape.
    # COM is the only authority: OMML stack height is renderer layout, not
    # derivable from the XML alone. TextBoundHeightPt = -1 marks a shape the
    # COM probe could not measure; callers must treat that as a failure.
    # Activation and enumeration retry on transient COM failures: this runs
    # right after other automation sessions released PowerPoint, where
    # RPC_S_UNKNOWN_IF (0x800706B5) activation races are routine.
    param([string]$PptxPath)
    $fullPath = [System.IO.Path]::GetFullPath($PptxPath)
    $attempts = 0
    while ($true) {
        $attempts++
        $powerPoint = $null
        $presentation = $null
        try {
            $powerPoint = New-Object -ComObject PowerPoint.Application
            $presentation = $powerPoint.Presentations.Open($fullPath, $true, $false, $false)
            $rows = New-Object System.Collections.Generic.List[object]
            foreach ($slide in @($presentation.Slides)) {
                foreach ($shape in @($slide.Shapes)) {
                    if ([string]$shape.Name -notlike 'FormulaOmml*') { continue }
                    $boundHeight = -1.0
                    try { $boundHeight = [double]$shape.TextFrame2.TextRange.BoundHeight } catch { $boundHeight = -1.0 }
                    $null = $rows.Add([pscustomobject]@{
                        Slide = [int]$slide.SlideIndex
                        Name = [string]$shape.Name
                        BoxHeightPt = [double]$shape.Height
                        TextBoundHeightPt = $boundHeight
                    })
                }
            }
            return ,$rows.ToArray()
        } catch {
            $category = Get-ComFailureCategory -ErrorRecord $_
            if ($attempts -ge 4 -or -not (Test-IsRetryablePresentationFailure -Category $category -Message $_.Exception.Message)) { throw }
            Write-Host ("[size-fit] COM attempt {0} failed transiently, retrying: {1}" -f $attempts, (Format-ComFailureDetails -ErrorRecord $_))
            Start-Sleep -Seconds (3 * $attempts)
        } finally {
            if ($null -ne $presentation) { $presentation.Close() }
            if ($null -ne $powerPoint) { $powerPoint.Quit() }
            Release-ComObjectSafe $presentation
            Release-ComObjectSafe $powerPoint
        }
    }
}

function Repair-FormulaOmmlOverflowSizes {
    # One deterministic XML rewrite pass: scale every a:rPr sz inside each
    # overflowing FormulaOmml shape by fitted/base, so subscript runs keep
    # their 0.72 ratio (the applier's 800 floor is re-applied). The measurement
    # comes from Measure-FormulaOmmlRenderBounds; this function is pure zip/XML.
    # Returns $true when at least one shape was rewritten.
    param(
        [object[]]$Measurements,
        [double]$TolerancePt,
        [string]$PptxPath,
        [System.Collections.Generic.List[object]]$ReportRows,
        [int]$Iteration
    )
    $fullPath = [System.IO.Path]::GetFullPath($PptxPath)
    $bySlide = @{}
    foreach ($item in @($Measurements)) {
        if ([double]$item.TextBoundHeightPt -lt 0) { continue }
        if ([double]$item.TextBoundHeightPt -le ([double]$item.BoxHeightPt + $TolerancePt)) { continue }
        $key = [int]$item.Slide
        if (-not $bySlide.ContainsKey($key)) { $bySlide[$key] = New-Object System.Collections.Generic.List[object] }
        $null = $bySlide[$key].Add($item)
    }
    if ($bySlide.Count -eq 0) { return $false }
    $zip = $null
    $openAttempts = 0
    while ($null -eq $zip) {
        $openAttempts++
        if ($openAttempts -gt 5) { throw ("Output pptx still locked after {0} open attempts: {1}" -f ($openAttempts - 1), $fullPath) }
        try {
            $zip = [System.IO.Compression.ZipFile]::Open($fullPath, [System.IO.Compression.ZipArchiveMode]::Update)
        } catch {
            Start-Sleep -Seconds (2 * $openAttempts)
        }
    }
    $rewroteAny = $false
    try {
        foreach ($slideNumber in @($bySlide.Keys | Sort-Object)) {
            $partName = 'ppt/slides/slide{0}.xml' -f $slideNumber
            $doc = New-Object System.Xml.XmlDocument
            $doc.PreserveWhitespace = $false
            $doc.LoadXml((Read-ZipEntryText -Zip $zip -EntryName $partName))
            $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
            $ns.AddNamespace('p', $script:NsPresentationMain)
            $ns.AddNamespace('a', $script:NsDrawingMain)
            $changed = $false
            foreach ($item in $bySlide[$slideNumber]) {
                $target = $null
                foreach ($spNode in @($doc.SelectNodes('//p:sp', $ns))) {
                    $cNvPr = $spNode.SelectSingleNode('./p:nvSpPr/p:cNvPr', $ns)
                    if ($null -ne $cNvPr -and [string]$cNvPr.GetAttribute('name') -eq [string]$item.Name) { $target = $spNode; break }
                }
                if ($null -eq $target) {
                    throw ("FormulaOmml shape not found for size fit: slide={0} name={1}" -f $slideNumber, $item.Name)
                }
                $szNodes = @($target.SelectNodes('.//a:rPr[@sz]', $ns))
                if ($szNodes.Count -eq 0) {
                    throw ("FormulaOmml shape carries no sz runs: slide={0} name={1}" -f $slideNumber, $item.Name)
                }
                $baseSz = 0
                foreach ($szNode in $szNodes) {
                    $value = 0
                    if ([int]::TryParse([string]$szNode.GetAttribute('sz'), [ref]$value) -and $value -gt $baseSz) { $baseSz = $value }
                }
                if ($baseSz -le 0) {
                    throw ("FormulaOmml shape has no positive sz: slide={0} name={1}" -f $slideNumber, $item.Name)
                }
                $basePt = $baseSz / 100.0
                $fittedPt = Get-OmmlFormulaFitSizePt -CurrentSizePt $basePt -TextBoundHeightPt ([double]$item.TextBoundHeightPt) -BoxHeightPt ([double]$item.BoxHeightPt)
                $rowBase = [ordered]@{
                    Slide = $slideNumber
                    Shape = [string]$item.Name
                    Iteration = $Iteration
                    BoxHeightPt = [Math]::Round([double]$item.BoxHeightPt, 2)
                    TextBoundHeightPt = [Math]::Round([double]$item.TextBoundHeightPt, 2)
                    SizePtBefore = $basePt
                    SizePtFitted = ''
                    Status = 'Unfittable'
                }
                if ($null -eq $fittedPt) {
                    $ReportRows.Add([pscustomobject]$rowBase) | Out-Null
                    continue
                }
                $factor = $fittedPt / $basePt
                foreach ($szNode in $szNodes) {
                    $value = 0
                    if (-not [int]::TryParse([string]$szNode.GetAttribute('sz'), [ref]$value)) { continue }
                    $newSz = [int][Math]::Round($value * $factor)
                    if ($newSz -lt 800) { $newSz = 800 }
                    $szNode.SetAttribute('sz', [string]$newSz)
                }
                $changed = $true
                $rewroteAny = $true
                $rowBase.Status = 'Fitted'
                $rowBase.SizePtFitted = $fittedPt
                $ReportRows.Add([pscustomobject]$rowBase) | Out-Null
            }
            if ($changed) {
                Write-ZipEntryText -Zip $zip -EntryName $partName -Text (Convert-XmlDocumentToString -Document $doc)
            }
        }
    } finally {
        $zip.Dispose()
    }
    return $rewroteAny
}

# --- delivery root: explicit (outside reports/) or auto-versioned ---
$reportsRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'reports'
if (-not [string]::IsNullOrWhiteSpace($DeliveryRoot)) {
    $deliveryFullPath = [System.IO.Path]::GetFullPath($DeliveryRoot)
    if (($deliveryFullPath + '\').StartsWith(($reportsRoot + '\'), [System.StringComparison]::OrdinalIgnoreCase) -or
        $deliveryFullPath -eq $reportsRoot) {
        throw ("reports/ is reserved for <deck>_v<N> versioned deliveries; point -DeliveryRoot outside reports/ (got: {0})" -f $deliveryFullPath)
    }
    if (-not (Test-Path -LiteralPath $deliveryFullPath)) { New-Item -ItemType Directory -Path $deliveryFullPath -Force | Out-Null }
} else {
    $stem = Get-CanonicalDeckStem -FileName ([System.IO.Path]::GetFileName($inputFullPath))
    # Pipeline-stage suffixes (.callout/.brand/.normalized) describe the file,
    # not the deck: deliveries must keep the <deck>_v<N> directory convention
    # (a letter directly before _v<N> is a layout-gate violation). The shared
    # helper loops until stable, unlike a single pass over a fixed suffix list.
    if (-not (Test-Path -LiteralPath $reportsRoot)) { New-Item -ItemType Directory -Path $reportsRoot -Force | Out-Null }
    $maxVersion = 0
    $resumableVersion = 0
    foreach ($dir in @(Get-ChildItem -LiteralPath $reportsRoot -Directory -ErrorAction SilentlyContinue)) {
        if ($dir.Name -match ('^' + [regex]::Escape($stem) + '_v(\d+)$')) {
            $maxVersion = [Math]::Max($maxVersion, [int]$Matches[1])
            # -Resume must target the latest existing batch manifest; a fresh
            # max+1 dir would never contain the manifest to resume from.
            if ($Resume -and (Test-Path -LiteralPath (Get-BatchManifestPath -Root $dir.FullName))) {
                $resumableVersion = [Math]::Max($resumableVersion, [int]$Matches[1])
            }
        }
    }
    if ($resumableVersion -gt 0) {
        $deliveryFullPath = Join-Path $reportsRoot ('{0}_v{1}' -f $stem, $resumableVersion)
    } else {
        $deliveryFullPath = Join-Path $reportsRoot ('{0}_v{1}' -f $stem, ($maxVersion + 1))
        New-Item -ItemType Directory -Path $deliveryFullPath | Out-Null
    }
}
$reportDir = Join-Path $deliveryFullPath '00_检查报告'
$deliveryDir = Join-Path $deliveryFullPath '01_交付物'
foreach ($dir in @($reportDir, $deliveryDir)) {
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}
$manifestPath = Get-BatchManifestPath -Root $deliveryFullPath

# --- resume baseline ---
$batchSteps = New-Object System.Collections.Generic.List[object]
if ($Resume -and (Test-Path -LiteralPath $manifestPath)) {
    $previous = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($step in @($previous.steps)) { $batchSteps.Add($step) | Out-Null }
}

function Get-RecordedStep {
    param([string]$StepName)
    foreach ($step in $batchSteps) {
        if ([string]$step.name -eq $StepName) { return $step }
    }
    return $null
}

function Invoke-BatchStep {
    param(
        [string]$StepName,
        [scriptblock]$Action,
        [string[]]$Inputs = @(),
        [string[]]$Flags = @(),
        [string[]]$Outputs = @()
    )
    $existing = $null
    if ($Resume) { $existing = Get-RecordedStep -StepName $StepName }
    $expectedSignature = New-StepSignature -StepName $StepName -InputPaths $Inputs -Flags $Flags
    $samePaths = $null -ne $existing -and
        ((@($existing.inputs) -join "`n") -ceq (@($Inputs) -join "`n")) -and
        ((@($existing.outputs) -join "`n") -ceq (@($Outputs) -join "`n"))
    if ($samePaths -and [string]$existing.signature -ceq $expectedSignature -and (Test-StepReusable -Step $existing)) {
        Write-Output ("[resume] step '{0}' reused (signature unchanged, outputs present)." -f $StepName)
        return
    }
    # Drop any stale recording of this step before rerunning it.
    for ($i = $batchSteps.Count - 1; $i -ge 0; $i--) {
        if ([string]$batchSteps[$i].name -eq $StepName) { $batchSteps.RemoveAt($i) }
    }
    Write-Output ("[step] {0}" -f $StepName)
    $null = & $Action
    foreach ($output in $Outputs) {
        if ([string]::IsNullOrWhiteSpace($output) -or -not (Test-Path -LiteralPath $output)) {
            throw "Step '$StepName' did not produce its required output: $output"
        }
    }
    $batchSteps.Add([pscustomobject]@{
        name = $StepName
        status = 'Passed'
        inputs = @($Inputs)
        flags = @($Flags)
        signature = New-StepSignature -StepName $StepName -InputPaths $Inputs -Flags $Flags
        outputs = @($Outputs)
        outputHashes = @($Outputs | ForEach-Object { Get-FileSha256 -Path $_ })
        at = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }) | Out-Null
}

$inputSha256 = Get-FileSha256 -Path $inputFullPath
$goldSetSha256 = Get-FileSha256 -Path $goldSetPath
$completedManifestWritten = $false
try {
    # --- step 1: carrier inventory ---
    $inventoryDir = Join-Path $reportDir 'inventory'
    $inventoryJson = Join-Path $inventoryDir 'formula-carrier-inventory.json'
    Invoke-BatchStep -StepName 'inventory' -Inputs @($inputFullPath) -Outputs @($inventoryJson) -Action {
        & (Join-Path $PSScriptRoot 'Export-FormulaCarrierInventory.ps1') -InputPath $inputFullPath -OutputDir $inventoryDir | Out-Null
    }

    # --- step 2: rendered pages (before-audit) ---
    $beforeAuditDir = Join-Path $reportDir 'before-audit'
    $resolvedPagesDir = $PagesDir
    if ([string]::IsNullOrWhiteSpace($resolvedPagesDir) -and -not $SkipVisualAudit) {
        Invoke-BatchStep -StepName 'before-audit' -Inputs @($inputFullPath) -Outputs @((Join-Path $beforeAuditDir 'pptx-visual-audit.csv'), $beforeAuditDir) -Action {
            & (Join-Path $PSScriptRoot 'Export-PptxVisualAudit.ps1') -InputPath $inputFullPath -OutputDir $beforeAuditDir | Out-Null
        }
        $resolvedPagesDir = Join-Path $beforeAuditDir 'pages'
    }
    if ([string]::IsNullOrWhiteSpace($resolvedPagesDir)) {
        throw 'Crops require rendered pages: pass -PagesDir or drop -SkipVisualAudit.'
    }
    $resolvedPagesDir = [System.IO.Path]::GetFullPath($resolvedPagesDir)

    # --- step 3: OLE crops ---
    $cropsDir = Join-Path $reportDir 'crops'
    $cropsCsv = Join-Path $cropsDir 'ole-crops.csv'
    Invoke-BatchStep -StepName 'crops' -Inputs @($inputFullPath, $inventoryJson, $resolvedPagesDir) -Outputs @($cropsCsv, $cropsDir) -Action {
        & (Join-Path $PSScriptRoot 'Export-FormulaOleCrops.ps1') -CarrierInventoryJson $inventoryJson -PagesDir $resolvedPagesDir -OutputDir $cropsDir -Carrier Convertable | Out-Null
    }

    # --- step 4: GoldSet mapping (page-kind exclusion lives here) ---
    $mappingCsv = Join-Path $reportDir 'formula-ole-mapping.csv'
    $reviewCsv = Join-Path $reportDir 'formula-ole-review.csv'
    Invoke-BatchStep -StepName 'mapping' -Inputs @($inventoryJson, $goldSetPath) -Outputs @($mappingCsv, $reviewCsv) -Action {
        & (Join-Path $PSScriptRoot 'Export-FormulaOleMapping.ps1') -CarrierInventoryJson $inventoryJson -GoldSetCsv $goldSetPath -OutputDir $reportDir -MaxItems $MaxItems | Out-Null
    }

    # --- step 5: OMML candidates ---
    $fragmentsDir = Join-Path $reportDir 'omml-fragments'
    $candidatesCsv = Join-Path $fragmentsDir 'formula-omml-candidates.csv'
    Invoke-BatchStep -StepName 'candidates' -Inputs @($reviewCsv, $goldSetPath) -Outputs @($candidatesCsv, $fragmentsDir) -Action {
        & (Join-Path $PSScriptRoot 'Export-FormulaOmmlCandidates.ps1') -FormulaReviewCsv $reviewCsv -OutputDir $fragmentsDir | Out-Null
    }

    # --- step 6: apply to a new copy ---
    # The delivery artifact is named after the canonical deck stem (the
    # 2026-09-17 deliverable rule: no stage suffixes), not after the chain
    # head's .callout/.brand file name.
    $outputStem = Get-CanonicalDeckStem -FileName ([System.IO.Path]::GetFileName($inputFullPath))
    $outputName = "$outputStem$([System.IO.Path]::GetExtension($inputFullPath))"
    $outputPptx = Join-Path $deliveryDir $outputName
    $applyReport = Join-Path $reportDir 'formula-ole-apply-report.csv'
    Invoke-BatchStep -StepName 'apply' -Inputs @($inputFullPath, $mappingCsv, $fragmentsDir) -Outputs @($outputPptx, $applyReport) -Action {
        & (Join-Path $PSScriptRoot 'Apply-FormulaOmmlForOle.ps1') -InputPath $inputFullPath -MappingCsv $mappingCsv -OmmlCandidateCsv $candidatesCsv -OutputPath $outputPptx -ReportPath $applyReport -MaxItems $MaxItems | Out-Null
    }

    # --- step 7: rendered size fit (PowerPoint-measured overflow guard) ---
    $sizeFitReport = Join-Path $reportDir 'formula-ole-size-fit-report.csv'
    Invoke-BatchStep -StepName 'size-fit' -Inputs @($outputPptx) -Outputs @($outputPptx, $sizeFitReport) -Action {
        $fitRows = New-Object System.Collections.Generic.List[object]
        $sizeFitTolerancePt = 2.0
        $stillOverflow = @()
        foreach ($iteration in 1..2) {
            $measurements = Measure-FormulaOmmlRenderBounds -PptxPath $outputPptx
            if (@($measurements).Count -eq 0) { break }
            $unmeasured = @($measurements | Where-Object { [double]$_.TextBoundHeightPt -lt 0 })
            if (@($unmeasured).Count -gt 0) {
                throw ("FormulaOmml shapes could not be measured in PowerPoint: {0}" -f (($unmeasured | ForEach-Object { 'slide={0} {1}' -f $_.Slide, $_.Name }) -join '; '))
            }
            foreach ($m in @($measurements | Where-Object { [double]$_.TextBoundHeightPt -le ([double]$_.BoxHeightPt + $sizeFitTolerancePt) })) {
                $fitRows.Add([pscustomobject]@{
                    Slide = [int]$m.Slide
                    Shape = [string]$m.Name
                    Iteration = $iteration
                    BoxHeightPt = [Math]::Round([double]$m.BoxHeightPt, 2)
                    TextBoundHeightPt = [Math]::Round([double]$m.TextBoundHeightPt, 2)
                    SizePtBefore = ''
                    SizePtFitted = ''
                    Status = 'Ok'
                }) | Out-Null
            }
            $rewrote = Repair-FormulaOmmlOverflowSizes -Measurements $measurements -TolerancePt $sizeFitTolerancePt -PptxPath $outputPptx -ReportRows $fitRows -Iteration $iteration
            if (-not $rewrote) { break }
            if ($iteration -eq 2) {
                $finalMeasurements = Measure-FormulaOmmlRenderBounds -PptxPath $outputPptx
                $stillOverflow = @($finalMeasurements | Where-Object { [double]$_.TextBoundHeightPt -gt ([double]$_.BoxHeightPt + $sizeFitTolerancePt) })
            }
        }
        Write-Utf8BomCsv -InputObject $fitRows.ToArray() -Path $sizeFitReport
        if (@($stillOverflow).Count -gt 0) {
            throw ("FormulaOmml shapes still overflow their box after size fit: {0}" -f (($stillOverflow | ForEach-Object { 'slide={0} {1} (bound {2:N1}pt > box {3:N1}pt)' -f $_.Slide, $_.Name, $_.TextBoundHeightPt, $_.BoxHeightPt }) -join '; '))
        }
    }

    # --- step 8: Open XML validator ---
    $validatorJson = Join-Path $reportDir 'validator.json'
    $dotnetPath = $null
    foreach ($candidatePath in @(Get-DotNetCommandCandidates)) {
        $sdkOutput = Invoke-NativeCommandWithEapGuard -FilePath $candidatePath -Arguments @('--list-sdks')
        if ($LASTEXITCODE -eq 0 -and @($sdkOutput | Where-Object { [string]$_ -match '^10\.' }).Count -gt 0) {
            $dotnetPath = $candidatePath
            break
        }
    }
    Invoke-BatchStep -StepName 'validator' -Inputs @($outputPptx) -Outputs @($validatorJson) -Action {
        if ([string]::IsNullOrWhiteSpace($dotnetPath)) { throw 'No .NET 10 SDK found for the Open XML validator.' }
        $validatorOutput = & $dotnetPath run --project (Join-Path $PSScriptRoot 'FormulaOfficeMathValidator\FormulaOfficeMathValidator.csproj') -c Release -- $outputPptx --max-errors 10 --json $validatorJson
        $validatorOutput | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "FormulaOfficeMathValidator failed with exit code $LASTEXITCODE." }
    }

    # --- step 9: same-name final PDF (delivery completeness) ---
    # The converted PPTX is the new final deliverable; its PDF must match it,
    # not the pre-conversion copy the upstream full chain exported.
    $finalPdfPath = Join-Path $deliveryDir "$outputStem.pdf"
    Invoke-BatchStep -StepName 'final-pdf' -Inputs @($outputPptx) -Outputs @($finalPdfPath) -Action {
        # Activation races with the size-fit session that just released
        # PowerPoint (0x800706B5 family), so each attempt activates, opens,
        # and saves as one unit: a failure tears the attempt's objects down
        # before the next retry, and the finally block only handles the last
        # attempt's state.
        $attempts = 0
        while ($true) {
            $attempts++
            $powerPoint = $null
            $presentation = $null
            try {
                $powerPoint = New-PowerPointApplication
                $presentation = Invoke-WithComRetry -MaxRetries 3 -DelayMs 1200 -Action {
                    $powerPoint.Presentations.Open([System.IO.Path]::GetFullPath($outputPptx), $true, $false, $false)
                }.GetNewClosure()
                Invoke-WithComRetry -MaxRetries 3 -DelayMs 1200 -Action {
                    $presentation.SaveAs([System.IO.Path]::GetFullPath($finalPdfPath), 32)  # ppSaveAsPDF
                }.GetNewClosure() | Out-Null
                if (-not (Test-Path -LiteralPath $finalPdfPath) -or (Get-Item -LiteralPath $finalPdfPath).Length -le 0) {
                    throw "PowerPoint did not create a non-empty final PDF: $finalPdfPath"
                }
                break
            } catch {
                if ($null -ne $presentation) {
                    try { $presentation.Close() | Out-Null } catch { }
                    Release-ComObjectSafe $presentation
                    $presentation = $null
                }
                if ($null -ne $powerPoint) {
                    try { Stop-PowerPointApplicationIfIdle -Application $powerPoint | Out-Null } catch { }
                    Release-ComObjectSafe $powerPoint
                    $powerPoint = $null
                }
                if ($attempts -ge 4 -or -not (Test-IsRetryablePresentationFailure -Category (Get-ComFailureCategory -ErrorRecord $_) -Message $_.Exception.Message)) { throw }
                Write-Host "[final-pdf] COM attempt $attempts failed transiently, retrying: $(Format-ComFailureDetails -ErrorRecord $_)"
                Start-Sleep -Seconds (3 * $attempts)
            } finally {
                if ($null -ne $presentation) {
                    try { $presentation.Close() | Out-Null } catch { }
                    Release-ComObjectSafe $presentation
                }
                if ($null -ne $powerPoint) {
                    try { Stop-PowerPointApplicationIfIdle -Application $powerPoint | Out-Null } catch { }
                    Release-ComObjectSafe $powerPoint
                }
            }
        }
    }

    # --- step 10: after visual audit (PowerPoint open proof + pages) ---
    if (-not $SkipVisualAudit) {
        $afterAuditDir = Join-Path $reportDir 'after-audit'
        Invoke-BatchStep -StepName 'after-audit' -Inputs @($outputPptx) -Outputs @((Join-Path $afterAuditDir 'pptx-visual-audit.csv'), $afterAuditDir) -Action {
            & (Join-Path $PSScriptRoot 'Export-PptxVisualAudit.ps1') -InputPath $outputPptx -OutputDir $afterAuditDir | Out-Null
        }
    }

    # --- batch manifest ---
    $applyRows = @()
    if (Test-Path -LiteralPath $applyReport) {
        $applyRows = @(Import-Csv -LiteralPath $applyReport -Encoding UTF8)
    }
    $replacedCount = @($applyRows | Where-Object { $_.Issue -eq 'OleFormulaReplaced' }).Count
    $refusedCount = @($applyRows | Where-Object { $_.Issue -notin @('OleFormulaReplaced', 'SavedAs') }).Count
    $sizeFitRowsForManifest = @()
    if (Test-Path -LiteralPath $sizeFitReport) {
        $sizeFitRowsForManifest = @(Import-Csv -LiteralPath $sizeFitReport -Encoding UTF8)
    }
    $sizeFitAppliedCount = @($sizeFitRowsForManifest | Where-Object { [string]$_.Status -eq 'Fitted' }).Count
    # A batch whose mapping promised rows but replaced nothing is a failure,
    # not a quiet pass: the whole point of the run was the conversion.
    $batchStatus = 'Passed'
    $writeBackAllowed = $true
    $batchNote = 'Explicit migration batch on a new copy. Source PPTX untouched; mc:Fallback keeps original OLE objects.'
    if ($replacedCount -eq 0) {
        $batchStatus = 'Failed'
        $writeBackAllowed = $false
        $batchNote = 'Nothing was replaced; treat this delivery as inert and investigate the mapping/apply reports.'
    } elseif ($refusedCount -gt 0) {
        $batchStatus = 'Degraded'
        $batchNote = 'Some mapping rows were refused; see formula-ole-apply-report.csv for the per-row reasons before accepting this delivery.'
    }
    $manifest = [ordered]@{
        schemaVersion = 1
        generatedAt = Get-Date -Format 'yyyy-MM-dd hh:mm:ss'
        status = $batchStatus
        input = [ordered]@{ path = $inputFullPath; sha256 = $inputSha256 }
        goldSet = [ordered]@{ path = $goldSetPath; sha256 = $goldSetSha256 }
        deliveryRoot = $deliveryFullPath
        outputPptx = $outputPptx
        outputSha256 = Get-FileSha256 -Path $outputPptx
        finalPdf = $finalPdfPath
        finalPdfSha256 = Get-FileSha256 -Path $finalPdfPath
        applySummary = [ordered]@{
            replaced = $replacedCount
            refused = $refusedCount
        }
        sizeFit = [ordered]@{
            report = $sizeFitReport
            appliedCount = $sizeFitAppliedCount
        }
        steps = $batchSteps.ToArray()
        writeBackAllowed = $writeBackAllowed
        note = $batchNote
    }
    Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 10) -Path $manifestPath
    $completedManifestWritten = $true

    # Root summary.md: the same top-level overview the full-chain deliveries
    # carry, so every versioned delivery answers "which file is the final
    # deliverable" the same way regardless of which producer built it.
    $summaryLines = New-Object System.Collections.Generic.List[string]
    $summaryLines.Add('# 公式迁移批次交付总览') | Out-Null
    $summaryLines.Add('') | Out-Null
    $summaryLines.Add("- 生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
    $summaryLines.Add("- 批次状态：$batchStatus") | Out-Null
    $summaryLines.Add("- 输入（链头）：$inputFullPath") | Out-Null
    $summaryLines.Add("- 输出：$outputPptx") | Out-Null
    $summaryLines.Add("- 同名 PDF：$finalPdfPath") | Out-Null
    $summaryLines.Add("- 公式写回：replaced=$($manifest.applySummary.replaced)；refused=$($manifest.applySummary.refused)") | Out-Null
    $summaryLines.Add("- 尺寸守卫：size-fit 收缩 $($manifest.sizeFit.appliedCount) 处（明细见 00_检查报告/formula-ole-size-fit-report.csv）") | Out-Null
    $summaryLines.Add("- GoldSet：$goldSetPath") | Out-Null
    $summaryLines.Add('- 交付物说明：`01_交付物/` 内与源件同名的 PPTX/PDF 为唯一正式交付物；检查报告与逐页审查在 `00_检查报告/`。') | Out-Null
    $summaryLines.Add('- 结构化清单：`00_检查报告/formula-ole-batch-manifest.json`（步骤收据与哈希绑定）。') | Out-Null
    Write-Utf8BomText -Text ($summaryLines -join "`r`n") -Path (Join-Path $deliveryFullPath 'summary.md')

    if ($batchStatus -eq 'Passed') {
        Write-Output ("Batch passed: {0}" -f $deliveryFullPath)
    } else {
        Write-Output ("Batch {0}: {1}" -f $batchStatus, $deliveryFullPath)
    }
    Write-Output ("Output: {0} (replaced={1}; refused={2})" -f $outputPptx, $manifest.applySummary.replaced, $manifest.applySummary.refused)
    if ($batchStatus -eq 'Failed') { throw "Formula migration batch failed: $batchNote" }
} catch {
    # A failure after the Passed manifest was written (e.g. a reporting crash)
    # must not clobber the completed batch record with a Failed status.
    if (-not $completedManifestWritten) {
        $manifest = [ordered]@{
            schemaVersion = 1
            generatedAt = Get-Date -Format 'yyyy-MM-dd hh:mm:ss'
            status = 'Failed'
            error = $_.Exception.Message
            input = [ordered]@{ path = $inputFullPath; sha256 = $inputSha256 }
            goldSet = [ordered]@{ path = $goldSetPath; sha256 = $goldSetSha256 }
            deliveryRoot = $deliveryFullPath
            steps = $batchSteps.ToArray()
            writeBackAllowed = $false
            note = 'Batch failed; completed prefix recorded. Delete the delivery dir to rerun from scratch or use -Resume.'
        }
        Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 10) -Path $manifestPath
    }
    Write-Output ("Batch FAILED at: {0}" -f $_.Exception.Message)
    throw
}
