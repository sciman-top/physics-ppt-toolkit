<#
.SYNOPSIS
  Unify highlight/callout text-box styling (fill, border, corner) across a PPTX.

.DESCRIPTION
  Brings every yellow-family highlight box onto the configured highlight
  palette so conclusions, labels and media tags share one look:

    conclusion boxes  fill -> colors.yellowFill   (#FFF2CC)
    media tag chips   fill -> colors.videoYellow  (#FFD966)
    all of them       border -> colors.yellowBorder (#D6A300),
                      border weight -> STYLE.HIGHLIGHT.BOX_SHAPE.borderWeightPt,
                      corner -> rounded rectangle with
                      STYLE.HIGHLIGHT.BOX_SHAPE.cornerRadiusAdj.

  A box joins the conclusion family when its solid fill is the legacy pale
  yellow (#FFFFCC), the configured yellowFill, or plain white while already
  carrying a legacy orange/red border (white boxes without any border are
  plain text and stay untouched).  The media-tag family covers the amber
  (#FFDB93) video chips.

  Text content, shape position/size, animation and pictures/OLE objects are
  never modified; grouped shapes are skipped by default and only counted.
  The input PPTX is never modified: output goes to a fresh versioned
  delivery folder with a per-shape CSV report and an original backup.

.EXAMPLE
  .\Apply-PptxHighlightBoxStyle.ps1 -PptxPath reports\...\13.3比热容（王耀强）.normalized.brand.pptx
#>

[CmdletBinding()]
param(
    [string]$PptxPath = '',
    [string]$OutputRoot = '',
    [double]$BorderWeightPt = 0,
    [double]$CornerRadiusAdj = -1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'tools\PhysicsPpt.Common.ps1')

if ([string]::IsNullOrWhiteSpace($PptxPath)) {
    $PptxPath = Join-Path $root 'PPTX\13.3比热容（王耀强）.pptx'
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $root 'reports'
}
# PowerPoint COM resolves relative paths against its own process CWD, never
# the caller's — delivery paths must be absolute before any COM use.
$OutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
$PptxPath = [System.IO.Path]::GetFullPath($PptxPath)
if (-not (Test-Path -LiteralPath $PptxPath)) {
    throw "Input PPTX not found: $PptxPath"
}

$configPath = Join-Path $root 'config\physics-ppt-style.config.json'
$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

$boxRule = $null
foreach ($styleRule in @($config.styleRules)) {
    if ([string]$styleRule.id -eq 'STYLE.HIGHLIGHT.BOX_SHAPE') { $boxRule = $styleRule; break }
}
if ($null -eq $boxRule) { throw 'Config styleRules is missing STYLE.HIGHLIGHT.BOX_SHAPE.' }
if (-not $boxRule.enabled) {
    throw 'STYLE.HIGHLIGHT.BOX_SHAPE is disabled by configuration; enable it before unifying highlight boxes.'
}

function Get-BoxRuleNumber {
    param([string]$Name, [double]$Fallback)
    $prop = $boxRule.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Fallback }
    return [double]$prop.Value
}

function Convert-HexToRgbLong {
    param([Parameter(Mandatory = $true)][string]$ColorSpec)
    $digits = $ColorSpec.TrimStart('#')
    if ($digits -notmatch '^[0-9A-Fa-f]{6}$') { throw "Invalid colour spec: $ColorSpec" }
    return ([Convert]::ToInt32($digits.Substring(0, 2), 16) +
        ([Convert]::ToInt32($digits.Substring(2, 2), 16) -shl 8) +
        ([Convert]::ToInt32($digits.Substring(4, 2), 16) -shl 16))
}

function Convert-RgbLongToHex {
    param([int]$Rgb)
    $r = $Rgb -band 0xFF
    $g = ($Rgb -shr 8) -band 0xFF
    $b = ($Rgb -shr 16) -band 0xFF
    return ('#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
}

$script:MsoTrue = -1
$script:MsoFillSolid = 1
$script:MsoGroup = 6
$script:MsoShapeRectangle = 1
$script:MsoShapeRoundedRectangle = 5

$rgbConclusionFill = Convert-HexToRgbLong ([string]$config.colors.yellowFill)
$rgbChipFill = Convert-HexToRgbLong ([string]$config.colors.videoYellow)
$rgbBorder = Convert-HexToRgbLong ([string]$config.colors.yellowBorder)
$rgbLegacyYellow = Convert-HexToRgbLong '#FFFFCC'
$rgbChipAmber = Convert-HexToRgbLong '#FFDB93'
$rgbWhite = 0xFFFFFF
$rgbLegacyBorderOrange = Convert-HexToRgbLong '#FF9900'
$rgbLegacyBorderRed = Convert-HexToRgbLong '#C00000'

if ($BorderWeightPt -le 0) { $BorderWeightPt = Get-BoxRuleNumber -Name 'borderWeightPt' -Fallback 2.0 }
if ($CornerRadiusAdj -lt 0) { $CornerRadiusAdj = Get-BoxRuleNumber -Name 'cornerRadiusAdj' -Fallback 0.1 }
if ($BorderWeightPt -le 0 -or $BorderWeightPt -gt 6) { throw "BorderWeightPt out of range: $BorderWeightPt" }
if ($CornerRadiusAdj -le 0 -or $CornerRadiusAdj -ge 0.5) { throw "CornerRadiusAdj out of range: $CornerRadiusAdj" }

$stem = [System.IO.Path]::GetFileNameWithoutExtension($PptxPath)
# Re-running on an already callout-styled file keeps the canonical output name
# and the same version lineage instead of compounding '.callout.callout', and
# a '.brand'-suffixed input keeps the clean deck stem so delivery directories
# stay in the reports/<deck>_v<N> layout.
if ($stem.EndsWith('.brand.callout')) { $stem = $stem.Substring(0, $stem.Length - '.callout'.Length) }
elseif ($stem.EndsWith('.brand')) { $stem = $stem.Substring(0, $stem.Length - '.brand'.Length) }
elseif ($stem.EndsWith('.callout')) { $stem = $stem.Substring(0, $stem.Length - '.callout'.Length) }
if ($stem.EndsWith('.normalized')) { $stem = $stem.Substring(0, $stem.Length - '.normalized'.Length) }
$existing = @(Get-ChildItem -LiteralPath $OutputRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$stem`_v*" })
$nextVersion = 1
foreach ($dir in $existing) {
    if ($dir.Name -match '_v(\d+)$') { $nextVersion = [Math]::Max($nextVersion, [int]$Matches[1] + 1) }
}
$deliveryRoot = Join-Path $OutputRoot ("$stem`_v$nextVersion")
$deliveryDir = Join-Path $deliveryRoot '01_交付物'
$reportDir = Join-Path $deliveryRoot '00_检查报告'
$backupDir = Join-Path $deliveryRoot '03_原始备份'
foreach ($dir in @($deliveryDir, $reportDir, $backupDir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
Copy-Item -LiteralPath $PptxPath -Destination (Join-Path $backupDir "$stem.pptx") -Force

$workingPptx = Join-Path $deliveryDir "$stem.callout.pptx"
Copy-Item -LiteralPath $PptxPath -Destination $workingPptx -Force

$reportRows = New-Object System.Collections.Generic.List[object]
$script:counters = @{ Changed = 0; AlreadyApplied = 0; GroupsSkipped = 0; Failed = 0 }

function Get-ShapeSnapshot {
    param($Shape)
    $fillRgb = -1
    $lineRgb = -1
    $lineVisible = $false
    $weight = 0.0
    $autoShapeType = -1
    $adj = -1.0
    try { if ($Shape.Fill.Visible -eq $script:MsoTrue) { $fillRgb = [int]$Shape.Fill.ForeColor.RGB } } catch { }
    try { $lineVisible = ($Shape.Line.Visible -eq $script:MsoTrue) } catch { }
    if ($lineVisible) {
        try { $lineRgb = [int]$Shape.Line.ForeColor.RGB } catch { }
        try { $weight = [double]$Shape.Line.Weight } catch { }
    }
    try { $autoShapeType = [int]$Shape.AutoShapeType } catch { }
    try { if ($Shape.Adjustments.Count -ge 1) { $adj = [double]$Shape.Adjustments.Item(1) } } catch { }
    return [pscustomobject]@{
        Fill = $(if ($fillRgb -ge 0) { Convert-RgbLongToHex $fillRgb } else { 'none' })
        LineColor = $(if ($lineRgb -ge 0) { Convert-RgbLongToHex $lineRgb } else { 'none' })
        LineWeightPt = $weight
        AutoShapeType = $autoShapeType
        CornerAdj = $adj
    }
}

function Get-HighlightBoxClass {
    param($Shape)
    try {
        if (-not $Shape.HasTextFrame) { return $null }
        if ($Shape.TextFrame.HasText -ne $script:MsoTrue) { return $null }
        if ($Shape.Fill.Type -ne $script:MsoFillSolid) { return $null }
        if ($Shape.Fill.Visible -ne $script:MsoTrue) { return $null }
        $rgb = [int]$Shape.Fill.ForeColor.RGB
        if ($rgb -eq $rgbLegacyYellow -or $rgb -eq $rgbConclusionFill) { return 'Conclusion' }
        if ($rgb -eq $rgbChipAmber) { return 'MediaTag' }
        if ($rgb -eq $rgbWhite -and $Shape.Line.Visible -eq $script:MsoTrue) {
            $lineRgb = [int]$Shape.Line.ForeColor.RGB
            if ($lineRgb -eq $rgbLegacyBorderOrange -or $lineRgb -eq $rgbLegacyBorderRed) { return 'Conclusion' }
        }
    } catch { return $null }
    return $null
}

function Add-ReportRow {
    param([int]$SlideNumber, [string]$ShapeName, [string]$Class, $Before, $After, [string]$Result, [string]$Details)
    $reportRows.Add([pscustomobject]@{
        File = (Split-Path -Leaf $workingPptx)
        SlideNumber = $SlideNumber
        ShapeName = $ShapeName
        Class = $Class
        FillBefore = $Before.Fill
        FillAfter = $After.Fill
        LineColorBefore = $Before.LineColor
        LineColorAfter = $After.LineColor
        LineWeightBeforePt = $Before.LineWeightPt
        LineWeightAfterPt = $After.LineWeightPt
        CornerAdjBefore = $Before.CornerAdj
        CornerAdjAfter = $After.CornerAdj
        Result = $Result
        Details = $Details
    }) | Out-Null
}

function Set-HighlightBoxStyle {
    param($Shape, [int]$SlideNumber, [string]$ShapeName, [string]$Class)
    $isConclusion = $Class -eq 'Conclusion'
    $targetFill = if ($isConclusion) { $rgbConclusionFill } else { $rgbChipFill }
    $before = Get-ShapeSnapshot -Shape $Shape
    try {
        # Geometry first: it is the assignment most likely to be refused, and
        # failing before any write keeps the shape free of half-applied styles.
        # Assigning AutoShapeType (even the same value) makes PowerPoint rewrite
        # the bounding box, so the original frame is restored inside the same
        # idempotent block; type is only switched for true rectangles.
        Invoke-WithComRetry -Action {
            $left0 = [double]$Shape.Left
            $top0 = [double]$Shape.Top
            $width0 = [double]$Shape.Width
            $height0 = [double]$Shape.Height
            $name0 = [string]$Shape.Name
            if ([int]$Shape.AutoShapeType -eq $script:MsoShapeRectangle) {
                $Shape.AutoShapeType = $script:MsoShapeRoundedRectangle
            }
            if ($Shape.Adjustments.Count -ge 1) { $Shape.Adjustments.Item(1) = $CornerRadiusAdj }
            $Shape.Fill.ForeColor.RGB = $targetFill
            $Shape.Line.Visible = $script:MsoTrue
            $Shape.Line.ForeColor.RGB = $rgbBorder
            $Shape.Line.Weight = $BorderWeightPt
            # Text size guard: runs without an explicit sz inherit the layout
            # master and drift when the deck is re-themed.  Pin conclusion-box
            # text to the 36pt display size; runs already carrying a size are
            # left untouched (per-box overrides stay respected).
            if ($isConclusion) {
                $tr = $Shape.TextFrame.TextRange
                $runCount = $tr.Runs().Count
                for ($runIndex = 1; $runIndex -le $runCount; $runIndex++) {
                    $run = $tr.Runs($runIndex, 1)
                    $runSizeText = [string]$run.Font.Size
                    $runSize = [double]0
                    if (-not [double]::TryParse($runSizeText, [ref]$runSize) -or $runSize -lt 1) {
                        $run.Font.Size = 36
                    }
                }
            }
            if ([Math]::Abs([double]$Shape.Left - $left0) -gt 0.1) { $Shape.Left = $left0 }
            if ([Math]::Abs([double]$Shape.Top - $top0) -gt 0.1) { $Shape.Top = $top0 }
            if ([Math]::Abs([double]$Shape.Width - $width0) -gt 0.1) { $Shape.Width = $width0 }
            if ([Math]::Abs([double]$Shape.Height - $height0) -gt 0.1) { $Shape.Height = $height0 }
            # Assigning AutoShapeType renames the default shape name; downstream
            # tooling matches shapes by name across saves, so keep it stable.
            if ([string]$Shape.Name -ne $name0) { $Shape.Name = $name0 }
        } | Out-Null
    } catch {
        $script:counters.Failed++
        Add-ReportRow -SlideNumber $SlideNumber -ShapeName $ShapeName -Class $Class -Before $before -After $before `
            -Result 'Failed' -Details ([string]$_.Exception.Message)
        return
    }
    $after = Get-ShapeSnapshot -Shape $Shape
    $changed = ($before.Fill -ne $after.Fill) -or ($before.LineColor -ne $after.LineColor) -or
        ([Math]::Abs($before.LineWeightPt - $after.LineWeightPt) -gt 0.01) -or
        ([Math]::Abs($before.CornerAdj - $after.CornerAdj) -gt 0.001)
    if ($changed) {
        $script:counters.Changed++
        Add-ReportRow -SlideNumber $SlideNumber -ShapeName $ShapeName -Class $Class -Before $before -After $after -Result 'Changed' -Details ''
    } else {
        $script:counters.AlreadyApplied++
        Add-ReportRow -SlideNumber $SlideNumber -ShapeName $ShapeName -Class $Class -Before $after -After $after -Result 'AlreadyApplied' -Details ''
    }
}

$application = New-PowerPointApplication
$presentation = $null
try {
    $presentation = Invoke-WithComRetry -Action { $application.Presentations.Open($workingPptx, $false, $false, $false) }
    $slideCount = $presentation.Slides.Count
    for ($i = 1; $i -le $slideCount; $i++) {
        $slide = Invoke-WithComRetry -Action { $presentation.Slides.Item($i) }
        try {
            $shapes = $slide.Shapes
            $shapeCount = $shapes.Count
            for ($j = 1; $j -le $shapeCount; $j++) {
                $shape = Invoke-WithComRetry -Action { $shapes.Item($j) }
                try {
                    $shapeName = ''
                    try { $shapeName = [string]$shape.Name } catch { }
                    if ([int]$shape.Type -eq $script:MsoGroup) {
                        $script:counters.GroupsSkipped++
                        Add-ReportRow -SlideNumber $i -ShapeName $shapeName -Class 'Group' `
                            -Before ([pscustomobject]@{ Fill = 'none'; LineColor = 'none'; LineWeightPt = 0; AutoShapeType = -1; CornerAdj = -1 }) `
                            -After ([pscustomobject]@{ Fill = 'none'; LineColor = 'none'; LineWeightPt = 0; AutoShapeType = -1; CornerAdj = -1 }) `
                            -Result 'GroupSkipped' -Details 'Grouped shapes are skipped by invariant.'
                        continue
                    }
                    $class = Get-HighlightBoxClass -Shape $shape
                    if ($null -eq $class) { continue }
                    Set-HighlightBoxStyle -Shape $shape -SlideNumber $i -ShapeName $shapeName -Class $class
                } finally {
                    Release-ComObjectSafe $shape
                }
            }
        } finally {
            Release-ComObjectSafe $slide
        }
    }
    Invoke-WithComRetry -Action { $presentation.Save() } | Out-Null
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { }
        Release-ComObjectSafe $presentation
    }
    try { $application.Quit() } catch { }
    Release-ComObjectSafe $application
}

$reportCsv = Join-Path $reportDir 'highlight-box-unify-report.csv'
Write-Utf8BomCsv -InputObject $reportRows -Path $reportCsv

Write-Host "Highlight box unify complete: $workingPptx"
Write-Host ("Changed={0} AlreadyApplied={1} GroupsSkipped={2} Failed={3} SlideCount={4}" -f `
    $script:counters.Changed, $script:counters.AlreadyApplied, $script:counters.GroupsSkipped, $script:counters.Failed, $slideCount)
Write-Host "Report: $reportCsv"
