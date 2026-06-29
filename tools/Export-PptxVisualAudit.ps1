<#
.SYNOPSIS
  Export PPTX pages and build a rule-based visual/layout audit package.

.DESCRIPTION
  Opens a PPTX in Microsoft PowerPoint, exports every slide to PNG, optionally
  exports a PDF, scans shapes for layout risks, calculates simple visual page
  metrics, and writes CSV/JSON reports plus a contact sheet. The source PPTX is
  never modified.

.PARAMETER InputPath
  Source .pptx/.pptm file.

.PARAMETER OutputDir
  Directory where page images, reports, and contact sheets are written.

.PARAMETER NoPdf
  Skip PDF export.

.PARAMETER ContactSheet
  Build an all-slide contact sheet with issue tags.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [ValidateRange(640, 4096)]
    [int]$ExportWidth = 1600,

    [ValidateRange(360, 4096)]
    [int]$ExportHeight = 900,

    [ValidateRange(1, 20)]
    [double]$ShapeTolerancePt = 2.0,

    [ValidateRange(6, 72)]
    [double]$MinReadableFontSize = 24.0,

    [ValidateRange(1, 200)]
    [int]$MaxContactSheetItems = 80,

    [switch]$NoPdf,

    [switch]$ContactSheet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:MsoTrue = -1
$script:MsoFalse = 0
$script:MsoGroup = 6
$script:MsoPlaceholder = 14
$script:MsoPicture = 13
$script:MsoMedia = 16

function Convert-ToSafePathSegment {
    param([string]$Name)
    $safe = $Name
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$ch, '_')
    }
    $safe = $safe -replace '\s+', '_'
    $safe = $safe -replace '[^\p{L}\p{Nd}_-]+', '_'
    $safe = $safe.Trim('_')
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'presentation' }
    return $safe
}

function Add-AuditRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$File,
        [int]$Slide,
        [string]$Shape,
        [string]$Issue,
        [string]$Severity,
        [string]$Details
    )

    $Rows.Add([pscustomobject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        File = $File
        Slide = $Slide
        Shape = $Shape
        Issue = $Issue
        Severity = $Severity
        Details = $Details
    }) | Out-Null
}

function Write-Utf8BomCsv {
    param([object[]]$Rows, [string]$Path)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $csvLines = $Rows | ConvertTo-Csv -NoTypeInformation
    [System.IO.File]::WriteAllLines($Path, $csvLines, $utf8Bom)
}

function Get-ShapeText {
    param($Shape)
    try {
        if ($null -ne $Shape.TextFrame2 -and $Shape.TextFrame2.HasText -eq $script:MsoTrue) {
            return [string]$Shape.TextFrame2.TextRange.Text
        }
    } catch { }
    return ''
}

function Get-ShapeName {
    param($Shape)
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$Shape.Name)) { return [string]$Shape.Name }
    } catch { }
    return '(unknown)'
}

function Get-ShapeTypeName {
    param($Shape)
    try {
        switch ([int]$Shape.Type) {
            6 { return 'Group' }
            13 { return 'Picture' }
            14 { return 'Placeholder' }
            16 { return 'Media' }
            17 { return 'TextBox' }
            19 { return 'Table' }
            default { return "Type$([int]$Shape.Type)" }
        }
    } catch { }
    return 'Unknown'
}

function Test-IsFormulaText {
    param([string]$Text)
    $t = ($Text -replace '\s+', '')
    if ([string]::IsNullOrWhiteSpace($t) -or $t.Length -gt 80) { return $false }
    if ($t -match '[=ηΩρ]|[/÷×∙·√]|(W有|W总|W额|G物|G动|R[12]|U[12]|I[12]|P[12])') { return $true }
    return $false
}

function Get-TextFontSize {
    param($Shape)
    try {
        $size = [double]$Shape.TextFrame2.TextRange.Font.Size
        if ($size -gt 0 -and $size -lt 200) { return $size }
    } catch { }
    return $null
}

function Get-TextFontName {
    param($Shape)
    try {
        return [string]$Shape.TextFrame2.TextRange.Font.Name
    } catch { }
    return ''
}

function Get-TextBounds {
    param($Shape)
    try {
        $range = $Shape.TextFrame2.TextRange
        $boundWidth = [double]$range.BoundWidth
        $boundHeight = [double]$range.BoundHeight
        if ($boundWidth -gt 0 -or $boundHeight -gt 0) {
            return [pscustomobject]@{ Width = $boundWidth; Height = $boundHeight }
        }
    } catch { }
    return $null
}

function Export-SlidePng {
    param($Slide, [string]$Path, [int]$Width, [int]$Height)
    $Slide.Export($Path, 'PNG', $Width, $Height) | Out-Null
}

function Get-ImageVisualMetrics {
    param([string]$Path)

    Add-Type -AssemblyName System.Drawing
    $image = $null
    try {
        $image = [System.Drawing.Bitmap]::FromFile($Path)
        $step = [Math]::Max(1, [int][Math]::Floor([Math]::Max($image.Width, $image.Height) / 600))
        $total = 0
        $nearWhite = 0
        $dark = 0
        $nonWhite = 0
        $minX = $image.Width
        $minY = $image.Height
        $maxX = -1
        $maxY = -1
        for ($y = 0; $y -lt $image.Height; $y += $step) {
            for ($x = 0; $x -lt $image.Width; $x += $step) {
                $pixel = $image.GetPixel($x, $y)
                $total++
                $isNearWhite = ($pixel.R -ge 248 -and $pixel.G -ge 248 -and $pixel.B -ge 248)
                if ($isNearWhite) {
                    $nearWhite++
                } else {
                    $nonWhite++
                    if ($x -lt $minX) { $minX = $x }
                    if ($y -lt $minY) { $minY = $y }
                    if ($x -gt $maxX) { $maxX = $x }
                    if ($y -gt $maxY) { $maxY = $y }
                }
                if (($pixel.R + $pixel.G + $pixel.B) -lt 90) { $dark++ }
            }
        }
        $whitePercent = if ($total -gt 0) { [Math]::Round(($nearWhite / [double]$total) * 100, 2) } else { 0 }
        $nonWhitePercent = if ($total -gt 0) { [Math]::Round(($nonWhite / [double]$total) * 100, 2) } else { 0 }
        $darkPercent = if ($total -gt 0) { [Math]::Round(($dark / [double]$total) * 100, 2) } else { 0 }
        $touchMargin = [int][Math]::Round([Math]::Min($image.Width, $image.Height) * 0.015)
        $touchesEdge = $false
        if ($nonWhite -gt 0) {
            $touchesEdge = ($minX -le $touchMargin -or $minY -le $touchMargin -or ($image.Width - $maxX) -le $touchMargin -or ($image.Height - $maxY) -le $touchMargin)
        }
        return [pscustomobject]@{
            Width = $image.Width
            Height = $image.Height
            WhitePercent = $whitePercent
            NonWhitePercent = $nonWhitePercent
            DarkPercent = $darkPercent
            IsVisuallyBlank = ($nonWhitePercent -lt 0.25)
            ContentTouchesEdge = $touchesEdge
            ContentBounds = if ($nonWhite -gt 0) { "$minX,$minY,$maxX,$maxY" } else { '' }
        }
    } finally {
        if ($null -ne $image) { $image.Dispose() }
    }
}

function New-VisualAuditContactSheet {
    param(
        [object[]]$SlideRows,
        [hashtable]$IssueMap,
        [string]$OutputPath,
        [int]$MaxItems
    )

    Add-Type -AssemblyName System.Drawing
    $items = @($SlideRows | Select-Object -First $MaxItems)
    if ($items.Count -eq 0) { return $false }

    $tileWidth = 360
    $tileHeight = 260
    $imageHeight = 190
    [int]$columns = 4
    if ($items.Count -lt 4) { [int]$columns = [Math]::Max(1, [int]$items.Count) }
    $rowsCount = [int][Math]::Ceiling($items.Count / [double]$columns)
    [int]$sheetWidth = [int]$columns * [int]$tileWidth
    [int]$sheetHeight = [int]$rowsCount * [int]$tileHeight

    $bitmap = $null
    $graphics = $null
    $font = $null
    $smallFont = $null
    $pen = $null
    try {
        $bitmap = New-Object System.Drawing.Bitmap -ArgumentList @($sheetWidth, $sheetHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::FromArgb(248, 248, 248))
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $font = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5)
        $smallFont = New-Object System.Drawing.Font('Microsoft YaHei UI', 7.5)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(205, 205, 205), 1)

        for ($i = 0; $i -lt $items.Count; $i++) {
            $row = [int][Math]::Floor($i / $columns)
            $col = $i % $columns
            $x = $col * $tileWidth
            $y = $row * $tileHeight
            $graphics.FillRectangle([System.Drawing.Brushes]::White, $x + 6, $y + 6, $tileWidth - 12, $tileHeight - 12)
            $graphics.DrawRectangle($pen, $x + 6, $y + 6, $tileWidth - 12, $tileHeight - 12)

            $img = $null
            try {
                $img = [System.Drawing.Image]::FromFile([string]$items[$i].ImagePath)
                $maxW = $tileWidth - 24
                $maxH = $imageHeight
                $scale = [Math]::Min($maxW / [double]$img.Width, $maxH / [double]$img.Height)
                $drawW = [Math]::Max(1, [int][Math]::Round($img.Width * $scale))
                $drawH = [Math]::Max(1, [int][Math]::Round($img.Height * $scale))
                $drawX = $x + [int](($tileWidth - $drawW) / 2)
                $drawY = $y + 14 + [int](($imageHeight - $drawH) / 2)
                $graphics.DrawImage($img, $drawX, $drawY, $drawW, $drawH)
            } finally {
                if ($null -ne $img) { $img.Dispose() }
            }

            $slide = [int]$items[$i].Slide
            $issues = if ($IssueMap.ContainsKey($slide)) { [string]$IssueMap[$slide] } else { '' }
            if ($issues.Length -gt 70) { $issues = $issues.Substring(0, 67) + '...' }
            $labelY = $y + $imageHeight + 18
            $graphics.DrawString("Slide $slide  issues=$($items[$i].IssueCount)", $font, [System.Drawing.Brushes]::Black, $x + 12, $labelY)
            $graphics.DrawString("white=$($items[$i].WhitePercent)% dark=$($items[$i].DarkPercent)%", $smallFont, [System.Drawing.Brushes]::DimGray, $x + 12, $labelY + 20)
            $graphics.DrawString($issues, $smallFont, [System.Drawing.Brushes]::DimGray, $x + 12, $labelY + 38)
        }

        $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        return $true
    } finally {
        if ($null -ne $pen) { $pen.Dispose() }
        if ($null -ne $smallFont) { $smallFont.Dispose() }
        if ($null -ne $font) { $font.Dispose() }
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
    }
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $InputPath)) { throw "InputPath not found: $InputPath" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$pageDir = Join-Path $OutputDir 'pages'
if (-not (Test-Path -LiteralPath $pageDir)) { New-Item -ItemType Directory -Path $pageDir -Force | Out-Null }

$fileName = [System.IO.Path]::GetFileName($InputPath)
$safeBase = Convert-ToSafePathSegment -Name ([System.IO.Path]::GetFileNameWithoutExtension($InputPath))
$shapeRows = New-Object System.Collections.Generic.List[object]
$auditRows = New-Object System.Collections.Generic.List[object]
$slideRows = New-Object System.Collections.Generic.List[object]

$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.Visible = $script:MsoTrue
    $pres = $pp.Presentations.Open($InputPath, $script:MsoFalse, $script:MsoFalse, $script:MsoFalse)
    $slideWidth = [double]$pres.PageSetup.SlideWidth
    $slideHeight = [double]$pres.PageSetup.SlideHeight
    $slideCount = [int]$pres.Slides.Count

    if (-not $NoPdf) {
        $pdfPath = Join-Path $OutputDir ($safeBase + '.pdf')
        try {
            $pres.ExportAsFixedFormat($pdfPath, 2) | Out-Null
            Add-AuditRow -Rows $auditRows -File $fileName -Slide 0 -Shape '(presentation)' -Issue 'PdfExported' -Severity 'Info' -Details $pdfPath
        } catch {
            try {
                $pres.SaveAs($pdfPath, 32) | Out-Null
                Add-AuditRow -Rows $auditRows -File $fileName -Slide 0 -Shape '(presentation)' -Issue 'PdfExportedBySaveAsFallback' -Severity 'Info' -Details $pdfPath
            } catch {
                Add-AuditRow -Rows $auditRows -File $fileName -Slide 0 -Shape '(presentation)' -Issue 'PdfExportFailed' -Severity 'Error' -Details $_.Exception.Message
            }
        }
    }

    for ($slideNo = 1; $slideNo -le $slideCount; $slideNo++) {
        $slide = $pres.Slides.Item($slideNo)
        $imagePath = Join-Path $pageDir ('slide-{0:000}.png' -f $slideNo)
        try {
            Export-SlidePng -Slide $slide -Path $imagePath -Width $ExportWidth -Height $ExportHeight
        } catch {
            Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape '(slide)' -Issue 'SlidePngExportFailed' -Severity 'Error' -Details $_.Exception.Message
        }

        foreach ($shape in $slide.Shapes) {
            $shapeName = Get-ShapeName -Shape $shape
            $shapeType = Get-ShapeTypeName -Shape $shape
            try {
                $left = [double]$shape.Left
                $top = [double]$shape.Top
                $width = [double]$shape.Width
                $height = [double]$shape.Height
                $right = $left + $width
                $bottom = $top + $height
                $text = Get-ShapeText -Shape $shape
                $fontSize = Get-TextFontSize -Shape $shape
                $fontName = Get-TextFontName -Shape $shape
                $hasText = -not [string]::IsNullOrWhiteSpace($text)
                $isFormula = if ($hasText) { Test-IsFormulaText -Text $text } else { $false }

                $shapeRows.Add([pscustomobject]@{
                    File = $fileName
                    Slide = $slideNo
                    Shape = $shapeName
                    ShapeType = $shapeType
                    Left = [Math]::Round($left, 2)
                    Top = [Math]::Round($top, 2)
                    Width = [Math]::Round($width, 2)
                    Height = [Math]::Round($height, 2)
                    FontSize = if ($null -eq $fontSize) { '' } else { [Math]::Round($fontSize, 1) }
                    FontName = $fontName
                    IsFormulaText = $isFormula
                    TextLength = if ($null -eq $text) { 0 } else { $text.Length }
                    TextPreview = (($text -replace '\s+', ' ').Trim())
                }) | Out-Null

                if ($left -lt (0 - $ShapeTolerancePt) -or $top -lt (0 - $ShapeTolerancePt) -or $right -gt ($slideWidth + $ShapeTolerancePt) -or $bottom -gt ($slideHeight + $ShapeTolerancePt)) {
                    $overflow = [Math]::Max([Math]::Max(0 - $left, 0 - $top), [Math]::Max($right - $slideWidth, $bottom - $slideHeight))
                    $severity = if ($overflow -gt 12) { 'Error' } else { 'Warning' }
                    Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape $shapeName -Issue 'ShapeOutOfSlideBounds' -Severity $severity -Details ("{0}: overflow={1:N1} pt; left={2:N1}; top={3:N1}; right={4:N1}; bottom={5:N1}; slide={6:N1}x{7:N1}" -f $shapeType, $overflow, $left, $top, $right, $bottom, $slideWidth, $slideHeight)
                }

                if ($hasText -and $null -ne $fontSize -and $fontSize -gt 0 -and $fontSize -lt $MinReadableFontSize) {
                    Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape $shapeName -Issue 'SmallReadableText' -Severity 'Warning' -Details ("fontSize={0:N1} pt; min={1:N1} pt; text={2}" -f $fontSize, $MinReadableFontSize, (($text -replace '\s+', ' ').Trim()))
                }

                if ($hasText) {
                    $bounds = Get-TextBounds -Shape $shape
                    if ($null -ne $bounds -and ($bounds.Width -gt ($width + 6) -or $bounds.Height -gt ($height + 6))) {
                        Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape $shapeName -Issue 'TextMayOverflowShape' -Severity 'Warning' -Details ("textBounds={0:N1}x{1:N1}; shape={2:N1}x{3:N1}; text={4}" -f $bounds.Width, $bounds.Height, $width, $height, (($text -replace '\s+', ' ').Trim()))
                    }
                }

                if ($isFormula) {
                    $severity = if ($fontName -match 'Cambria Math') { 'Info' } else { 'Warning' }
                    Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape $shapeName -Issue 'FormulaTextVisualReview' -Severity $severity -Details ("font={0}; size={1}; text={2}" -f $fontName, $fontSize, (($text -replace '\s+', ' ').Trim()))
                }
            } catch {
                Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape $shapeName -Issue 'ShapeAuditFailed' -Severity 'Warning' -Details $_.Exception.Message
            }
        }

        if (Test-Path -LiteralPath $imagePath) {
            $metrics = Get-ImageVisualMetrics -Path $imagePath
            if ($metrics.IsVisuallyBlank) {
                Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape '(slide)' -Issue 'SlideLooksBlank' -Severity 'Error' -Details ("nonWhite={0}%; bounds={1}" -f $metrics.NonWhitePercent, $metrics.ContentBounds)
            }
            if ($metrics.ContentTouchesEdge) {
                Add-AuditRow -Rows $auditRows -File $fileName -Slide $slideNo -Shape '(slide)' -Issue 'VisualContentTouchesEdge' -Severity 'Warning' -Details ("bounds={0}; image={1}x{2}" -f $metrics.ContentBounds, $metrics.Width, $metrics.Height)
            }
            $slideRows.Add([pscustomobject]@{
                File = $fileName
                Slide = $slideNo
                ImagePath = $imagePath
                Width = $metrics.Width
                Height = $metrics.Height
                WhitePercent = $metrics.WhitePercent
                NonWhitePercent = $metrics.NonWhitePercent
                DarkPercent = $metrics.DarkPercent
                IsVisuallyBlank = $metrics.IsVisuallyBlank
                ContentTouchesEdge = $metrics.ContentTouchesEdge
                ContentBounds = $metrics.ContentBounds
                IssueCount = 0
            }) | Out-Null
        }
    }
} finally {
    if ($null -ne $pres) {
        try { $pres.Close() | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    }
    if ($null -ne $pp) {
        try { $pp.Quit() | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$issueMap = @{}
foreach ($group in @($auditRows | Where-Object { [int]$_.Slide -gt 0 } | Group-Object Slide)) {
    $names = @($group.Group | Where-Object { $_.Severity -ne 'Info' } | Select-Object -ExpandProperty Issue -Unique)
    if ($names.Count -eq 0) { $names = @($group.Group | Select-Object -ExpandProperty Issue -Unique) }
    $issueMap[[int]$group.Name] = ($names -join '; ')
}

for ($i = 0; $i -lt $slideRows.Count; $i++) {
    $slideNo = [int]$slideRows[$i].Slide
    $count = @($auditRows | Where-Object { [int]$_.Slide -eq $slideNo -and $_.Severity -ne 'Info' }).Count
    $slideRows[$i].IssueCount = $count
}

$auditCsv = Join-Path $OutputDir 'pptx-visual-audit.csv'
$shapeCsv = Join-Path $OutputDir 'pptx-shape-audit.csv'
$slideCsv = Join-Path $OutputDir 'pptx-slide-visual-metrics.csv'
$manifestPath = Join-Path $OutputDir 'pptx-visual-audit-manifest.json'
$contactSheetPath = Join-Path $OutputDir 'pptx-visual-audit.contact-sheet.png'

Write-Utf8BomCsv -Rows $auditRows.ToArray() -Path $auditCsv
Write-Utf8BomCsv -Rows $shapeRows.ToArray() -Path $shapeCsv
Write-Utf8BomCsv -Rows $slideRows.ToArray() -Path $slideCsv

$contactSheetCreated = $false
if ($ContactSheet) {
    $contactSheetCreated = New-VisualAuditContactSheet -SlideRows $slideRows.ToArray() -IssueMap $issueMap -OutputPath $contactSheetPath -MaxItems $MaxContactSheetItems
}

$errorCount = @($auditRows | Where-Object { $_.Severity -eq 'Error' }).Count
$warningCount = @($auditRows | Where-Object { $_.Severity -eq 'Warning' }).Count
$manifest = [pscustomobject]@{
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    inputPath = $InputPath
    outputDir = $OutputDir
    pageDir = $pageDir
    slideCount = $slideRows.Count
    errorCount = $errorCount
    warningCount = $warningCount
    auditCsv = $auditCsv
    shapeCsv = $shapeCsv
    slideMetricsCsv = $slideCsv
    contactSheet = if ($contactSheetCreated) { $contactSheetPath } else { '' }
    note = 'Rule-based visual audit. Use exported PNG/PDF for final visual review before replacing originals.'
}
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Host "Visual audit done: $OutputDir"
Write-Host "Slides: $($slideRows.Count); Errors: $errorCount; Warnings: $warningCount"
