<#
.SYNOPSIS
  Apply the sciman brand visual refresh to a physics PPTX lesson.

.DESCRIPTION
  Only the cover, end, resource and link-hub pages carry the brand
  background image (top-right icon baked in).  All other pages stay on
  their original white background and get the brand icon as a uniform
  45%-opacity watermark in the top-right corner, so every page has brand
  coverage without competing with body text.

  Restyles the special pages with a gold/white/light-blue palette, unifies
  divider titles (same font, same size, semantic red/blue colour, centred
  in the text box and centred on the slide) and removes the legacy
  top-right channel icon.  All output goes to a fresh versioned delivery
  folder; the input PPTX is never modified.

  Page roles are detected from slide text, not hard-coded indexes:
    Cover     text contains 知乎主页
    End       text contains END
    Resource  text contains 课件下载地址
    LinkHub   text contains 网盘群
    Divider   a single short title matches the divider pattern and the slide
              has no pictures (第N课时 / 比热容（N）/ 比热容公式 / 拓展)
    Blank     no shapes at all (left untouched)

.EXAMPLE
  .\Apply-PptxBrandVisualRefresh.ps1 -PptxPath PPTX\13.3比热容（王耀强）.pptx
#>

[CmdletBinding()]
param(
    [string]$PptxPath = '',
    [string]$OutputRoot = '',
    [string]$AssetsDir = '',
    [int]$DividerFontSize = 54,
    [int]$IconDiameterPt = 48
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
if ([string]::IsNullOrWhiteSpace($AssetsDir)) {
    $AssetsDir = Join-Path $root 'assets\brand'
}

$bgPath = Join-Path $AssetsDir 'bg-16x9.jpg'
$iconShadowPath = Join-Path $AssetsDir 'sciman-icon-shadow.png'
$iconWatermarkPath = Join-Path $AssetsDir 'sciman-icon-watermark.png'
foreach ($asset in @($bgPath, $iconShadowPath, $iconWatermarkPath)) {
    if (-not (Test-Path -LiteralPath $asset)) {
        throw "Brand asset missing: $asset (run tools\generate_brand_assets.py first)"
    }
}
$PptxPath = [System.IO.Path]::GetFullPath($PptxPath)
if (-not (Test-Path -LiteralPath $PptxPath)) {
    throw "Input PPTX not found: $PptxPath"
}

$stem = [System.IO.Path]::GetFileNameWithoutExtension($PptxPath)
$existing = @(Get-ChildItem -LiteralPath $OutputRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$stem`_v*" })
$nextVersion = 1
foreach ($dir in $existing) {
    if ($dir.Name -match '_v(\d+)$') { $nextVersion = [Math]::Max($nextVersion, [int]$Matches[1] + 1) }
}
$deliveryRoot = Join-Path $OutputRoot ("$stem`_v$nextVersion")
$deliveryDir = Join-Path $deliveryRoot '01_交付物'
$pageImageDir = Join-Path $deliveryDir '页面图片'
$reportDir = Join-Path $deliveryRoot '00_检查报告'
$backupDir = Join-Path $deliveryRoot '03_原始备份'
foreach ($dir in @($deliveryDir, $pageImageDir, $reportDir, $backupDir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
Copy-Item -LiteralPath $PptxPath -Destination (Join-Path $backupDir "$stem.pptx") -Force

$workingPptx = Join-Path $deliveryDir "$stem.brand.pptx"
Copy-Item -LiteralPath $PptxPath -Destination $workingPptx -Force

# --- palette: named roles resolved to hex by Convert-HexToRgbLong ----------
# DividerRed/DividerBlue reuse the sanctioned config palette (emphasisRed /
# sectionTitle) so dividers stay consistent with the normalization toolkit.
$script:Palette = @{
    Title      = '#FFFFFF'
    Subtitle   = '#EAF2FB'
    Body       = '#F2F7FC'
    Secondary  = '#D6E4F0'
    Link       = '#9DC3E6'
    KeepYellow = '#FFD966'
    Gold       = '#FFD966'
    InfoValue  = '#B8CCE4'
    DividerRed   = '#C00000'
    DividerBlue  = '#1F4E79'
}

function Convert-HexToRgbLong {
    param([Parameter(Mandatory = $true)][string]$ColorSpec)
    if ($script:Palette.ContainsKey($ColorSpec)) {
        $ColorSpec = $script:Palette[$ColorSpec]
    }
    $digits = $ColorSpec.TrimStart('#')
    if ($digits -notmatch '^[0-9A-Fa-f]{6}$') {
        throw "Invalid colour spec: $ColorSpec"
    }
    return [int](
        [Convert]::ToInt32($digits.Substring(0, 2), 16) +
        [Convert]::ToInt32($digits.Substring(2, 2), 16) * 256 +
        [Convert]::ToInt32($digits.Substring(4, 2), 16) * 65536
    )
}

function Get-SlidePlainText {
    param($Slide)
    $parts = @()
    foreach ($shape in @($Slide.Shapes)) {
        try {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $parts += [string]$shape.TextFrame.TextRange.Text
            }
        } catch { }
    }
    return ($parts -join "`n")
}

function Test-DividerSlide {
    param($Slide, [string]$SlideText)
    if ($SlideText.Length -gt 12) { return $false }
    $hasDividerText = $false
    foreach ($shape in @($Slide.Shapes)) {
        try {
            if ($shape.Type -eq 13) { return $false }  # msoPicture: content page, never a divider
        } catch { }
        try {
            if ($shape.HasTextFrame -ne -1 -or $shape.TextFrame.HasText -ne -1) { continue }
            $trimmed = ([string]$shape.TextFrame.TextRange.Text).Trim()
            if ($trimmed -match '^(第[0-9０-９]+课时|拓展|比热容公式|比热容[（(][0-9０-９]+[)）])$') {
                $hasDividerText = $true
            }
        } catch { }
    }
    return $hasDividerText
}

function Get-SlideRole {
    param($Slide, [string]$SlideText)
    if ($SlideText -match '知乎主页') { return 'Cover' }
    if ($SlideText -match '\bEND\b') { return 'End' }
    if ($SlideText -match '课件下载地址') { return 'Resource' }
    if ($SlideText -match '网盘群') { return 'LinkHub' }
    if (@($Slide.Shapes).Count -eq 0) { return 'Blank' }
    if (Test-DividerSlide -Slide $Slide -SlideText $SlideText) { return 'Divider' }
    return 'Content'
}

function Remove-LegacyIconShapes {
    param($Slide, [System.Collections.Generic.List[string]]$Actions)
    $slideWidth = $Slide.Parent.PageSetup.SlideWidth
    $removed = 0
    $shapes = @($Slide.Shapes)
    for ($index = $shapes.Count - 1; $index -ge 0; $index--) {
        $shape = $shapes[$index]
        try {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = [string]$shape.TextFrame.TextRange.Text
                if (-not ($text -match 'sciman|逸居')) { continue }
            }
            if ($shape.Left -gt ($slideWidth - 110) -and $shape.Top -lt 95 -and
                $shape.Width -lt 150 -and $shape.Height -lt 150) {
                $shape.Delete()
                $removed++
            }
        } catch { }
    }
    if ($removed -gt 0) {
        $Actions.Add("RemovedLegacyIcon x$removed") | Out-Null
    }
}

function Set-BackgroundImage {
    param($Slide, [string]$ImagePath, [System.Collections.Generic.List[string]]$Actions)
    Invoke-WithComRetry -Action {
        try {
            $Slide.FollowMasterBackground = 0
            $Slide.Background.Fill.UserPicture($ImagePath)
        } catch {
            # Some layouts reject background fills; fall back to a full-bleed
            # picture sent behind every other shape.
            $slideWidth = $Slide.Parent.PageSetup.SlideWidth
            $slideHeight = $Slide.Parent.PageSetup.SlideHeight
            $picture = $Slide.Shapes.AddPicture($ImagePath, 0, -1, 0, 0, $slideWidth, $slideHeight)
            $picture.ZOrder(1)  # msoSendToBack
        }
    }.GetNewClosure() | Out-Null
    $Actions.Add('BackgroundImage') | Out-Null
}

function Set-ParagraphLook {
    param(
        $Paragraph,
        [string]$ColorSpec,
        [int]$Size = 0,
        [string]$FarEastFont = '微软雅黑',
        [string]$LatinFont = '微软雅黑',
        [int]$Bold = -2
    )
    $Paragraph.Font.Color.RGB = Convert-HexToRgbLong $ColorSpec
    if ($Size -gt 0) { $Paragraph.Font.Size = $Size }
    if (-not [string]::IsNullOrWhiteSpace($FarEastFont)) { $Paragraph.Font.NameFarEast = $FarEastFont }
    if (-not [string]::IsNullOrWhiteSpace($LatinFont)) { $Paragraph.Font.Name = $LatinFont }
    if ($Bold -ne -2) { $Paragraph.Font.Bold = $Bold }
}

function Protect-KeptRuns {
    <#  Re-apply the keep-yellow emphasis to the 禁止 characters after a
        paragraph-level recolor.  Character-range targeting keeps the
        highlight even when 禁止 is not an isolated run.  #>
    param($Paragraph, [string]$ExactText, [string]$ColorSpec)
    try {
        $text = [string]$Paragraph.Text
        $index = $text.IndexOf($ExactText)
        if ($index -ge 0) {
            $characters = $Paragraph.Characters($index + 1, $ExactText.Length)
            $characters.Font.Color.RGB = Convert-HexToRgbLong $ColorSpec
            $characters.Font.Bold = -1
        }
    } catch { }
}

function Set-InfoLabelSplit {
    <#  Colour the "label：" prefix of a contact-info paragraph in gold and
        the value in light blue, giving the cover info block a two-tone
        pairing instead of flat white.  #>
    param($Paragraph, [string]$LabelColor, [string]$ValueColor)
    try {
        $text = [string]$Paragraph.Text
        $separator = $text.IndexOf('：')
        if ($separator -lt 0) { $separator = $text.IndexOf(':') }
        if ($separator -ge 0) {
            $label = $Paragraph.Characters(1, $separator + 1)
            $label.Font.Color.RGB = Convert-HexToRgbLong $LabelColor
            $label.Font.Bold = -1
            if ($text.Length -gt $separator + 1) {
                $value = $Paragraph.Characters($separator + 2, $text.Length - $separator - 1)
                $value.Font.Color.RGB = Convert-HexToRgbLong $ValueColor
                $value.Font.Bold = 0
            }
        } else {
            $Paragraph.Font.Color.RGB = Convert-HexToRgbLong $ValueColor
        }
    } catch { }
}

function Update-SpecialSlideText {
    param(
        $Slide,
        [string]$Role,
        [System.Collections.Generic.List[string]]$Actions
    )
    $styled = 0
    $slideWidth = $Slide.Parent.PageSetup.SlideWidth
    foreach ($shape in @($Slide.Shapes)) {
        try {
            if ($shape.HasTextFrame -ne -1 -or $shape.TextFrame.HasText -ne -1) { continue }
            $textRange = $shape.TextFrame.TextRange
            $paragraphCount = $textRange.Paragraphs().Count
            for ($p = 1; $p -le $paragraphCount; $p++) {
                $paragraph = $textRange.Paragraphs($p, 1)
                $trimmed = ([string]$paragraph.Text).Trim()
                if ($trimmed.Length -eq 0) { continue }

                switch ($Role) {
                    'Cover' {
                        if ($trimmed -eq '比热容') {
                            # Calligraphy title keeps its font; gold recolour only.
                            Set-ParagraphLook $paragraph 'Gold' -FarEastFont '' -LatinFont ''
                            $styled++
                        } elseif ($trimmed -eq '广州番禺王耀强') {
                            Set-ParagraphLook $paragraph 'Subtitle' -FarEastFont '' -LatinFont ''
                            $styled++
                        } else {
                            # A leading fullwidth opening quote (U+201C)
                            # renders with a half-character side bearing that
                            # breaks the block's left edge, and the COM Ruler
                            # cannot reliably compensate.  Swap the decorative
                            # quotes to halfwidth (punctuation glyphs only)
                            # BEFORE the label/value split, then re-acquire the
                            # paragraph: assigning Text resets run formatting to
                            # the first character's, which would otherwise wipe
                            # the two-tone split.
                            if ($trimmed.Length -gt 0 -and $trimmed[0] -eq [char]0x201C) {
                                try {
                                    $parText = [string]$paragraph.Text
                                    $fixed = $parText.Replace([char]0x201C, '"').Replace([char]0x201D, '"')
                                    if ($fixed -ne $parText) {
                                        $paragraph.Text = $fixed
                                        $paragraph = $textRange.Paragraphs($p, 1)
                                    }
                                } catch { }
                            }
                            Set-ParagraphLook $paragraph 'Secondary' 22 '微软雅黑' 'Consolas'
                            Set-InfoLabelSplit $paragraph 'Gold' 'InfoValue'
                            # LineRuleAfter must be off, otherwise SpaceAfter is
                            # counted in LINES and pushes the next paragraphs
                            # below the slide edge.  Lines stay LEFT aligned;
                            # the block itself is centred on the slide below.
                            try {
                                $paragraph.ParagraphFormat.LineRuleAfter = 0
                                $paragraph.ParagraphFormat.SpaceAfter = 6
                                $paragraph.ParagraphFormat.Alignment = 1  # ppAlignLeft
                            } catch { }
                            $styled++
                        }
                    }
                    'End' {
                        if ($trimmed -eq 'END') {
                            Set-ParagraphLook $paragraph 'Gold' -Bold -1
                        } elseif ($trimmed -match '^http') {
                            Set-ParagraphLook $paragraph 'Link' 18 -Bold 0
                        } else {
                            Set-ParagraphLook $paragraph 'Body' -Bold 0
                            Protect-KeptRuns $paragraph '禁止' 'KeepYellow'
                        }
                        $styled++
                    }
                    'Resource' {
                        if ($trimmed -match '^课件下载地址') {
                            Set-ParagraphLook $paragraph 'Gold' -Bold -1
                        } elseif ($trimmed -match '^(百度网盘|阿里云盘|夸克网盘)') {
                            Set-ParagraphLook $paragraph 'Title' -Bold -1
                        } elseif ($trimmed -match '^http') {
                            Set-ParagraphLook $paragraph 'Link' 18 -Bold 0
                        } elseif ($trimmed -match '【|加群方法|知乎') {
                            Set-ParagraphLook $paragraph 'Secondary' -Bold 0
                        } else {
                            Set-ParagraphLook $paragraph 'Body' -Bold 0
                        }
                        $styled++
                    }
                    'LinkHub' {
                        if ($trimmed -match '^网盘群') {
                            Set-ParagraphLook $paragraph 'Gold' -Bold -1
                        } elseif ($trimmed -match '^[0-9０-９]{5,}') {
                            # Pure group-number lines go light blue for contrast
                            # against the white instruction lines.
                            Set-ParagraphLook $paragraph 'Link' -Bold 0
                        } else {
                            Set-ParagraphLook $paragraph 'Body' -Bold 0
                        }
                        $styled++
                    }
                }
            }
            if ($Role -eq 'Cover' -and $shape.HasTextFrame -eq -1 -and
                    $shape.TextFrame.HasText -eq -1 -and
                    ([string]$shape.TextFrame.TextRange.Text) -match '知乎主页') {
                # Centre the contact-info BLOCK on the slide while its lines
                # stay left-aligned: shrink the box to hug the longest line,
                # then centre that box.
                $infoFrame = $shape.TextFrame
                $infoFrame.WordWrap = 0
                $infoFrame.AutoSize = 0  # ppAutoSizeNone: manual hug sizing
                $hugHeight = $textRange.BoundHeight + $infoFrame.MarginTop + $infoFrame.MarginBottom
                # BoundWidth counts the trailing paragraph-break advance of the
                # last line (~one Consolas cell at 22pt); trim it so the ink,
                # not the phantom cell, is what gets centred.
                $hugWidth = [Math]::Max(100, $textRange.BoundWidth - 12 + $infoFrame.MarginLeft + $infoFrame.MarginRight)
                $shape.Width = [Math]::Round($hugWidth, 1)
                $shape.Height = [Math]::Round($hugHeight, 1)
                $shape.Left = [Math]::Round(($slideWidth - $shape.Width) / 2, 1)
                $Actions.Add('InfoBlockCentred(left-aligned)') | Out-Null
            }
        } catch {
            $Actions.Add("ShapeSkipped: $($_.Exception.Message)") | Out-Null
        }
    }
    if ($styled -gt 0) {
        $Actions.Add("TextRestyled($Role) x$styled") | Out-Null
    }
}

function Update-DividerSlide {
    <#  Dividers stay on their original white background (only the four
        special pages carry the brand image).  They get a unified font/size
        with double centring plus a sanctioned semantic colour: emphasis red
        for 课时/拓展 section marks, section-title navy for topic titles.  #>
    param(
        $Slide,
        [int]$FontSize,
        [System.Collections.Generic.List[string]]$Actions
    )
    $slideWidth = $Slide.Parent.PageSetup.SlideWidth
    $slideHeight = $Slide.Parent.PageSetup.SlideHeight
    $centered = 0
    foreach ($shape in @($Slide.Shapes)) {
        try {
            if ($shape.HasTextFrame -ne -1 -or $shape.TextFrame.HasText -ne -1) { continue }
            $trimmed = ([string]$shape.TextFrame.TextRange.Text).Trim()
            if ($trimmed.Length -eq 0) { continue }
            $textRange = $shape.TextFrame.TextRange
            if ($trimmed -match '^(第[0-9０-９]+课时|拓展)$') {
                $textRange.Font.Color.RGB = Convert-HexToRgbLong 'DividerRed'
            } else {
                $textRange.Font.Color.RGB = Convert-HexToRgbLong 'DividerBlue'
            }
            $textRange.Font.NameFarEast = '微软雅黑'
            $textRange.Font.Name = '微软雅黑'
            $textRange.Font.Bold = -1
            $textRange.Font.Size = $FontSize
            $shape.TextFrame.AutoSize = 0  # ppAutoSizeNone
            $shape.TextFrame.WordWrap = 0  # msoFalse so the box hugs the text
            $shape.TextFrame.VerticalAnchor = 3  # msoAnchorMiddle
            $textRange.ParagraphFormat.Alignment = 2  # ppAlignCenter
            $boundWidth = $textRange.BoundWidth + 20
            $boundHeight = $textRange.BoundHeight + 12
            # BoundWidth arrives as a COM Single; the +offset makes these
            # Doubles.  Assign [single] so the put marshals as VT_R4 —
            # assigning the raw Double can hit an IDispatch conversion path
            # that throws InvalidCastException Double->Decimal.
            $shape.Width = [single]$boundWidth
            $shape.Height = [single]$boundHeight
            $shape.Left = [single][Math]::Round(($slideWidth - $boundWidth) / 2, 1)
            $shape.Top = [single][Math]::Round(($slideHeight - $boundHeight) / 2, 1)
            $centered++
        } catch {
            $Actions.Add("DividerShapeSkipped: $($_.Exception.Message) << $($_.ScriptStackTrace)") | Out-Null
        }
    }
    if ($centered -gt 0) {
        $Actions.Add("DividerUnified x$centered (微软雅黑 ${FontSize}pt 规范色 双居中)") | Out-Null
    }
}

function Test-IconZoneOccupied {
    param($Slide, [double]$ZoneLeft, [double]$ZoneTop, [double]$ZoneRight, [double]$ZoneBottom)
    $pageSetup = $Slide.Parent.PageSetup
    $slideArea = $pageSetup.SlideWidth * $pageSetup.SlideHeight
    foreach ($shape in @($Slide.Shapes)) {
        try {
            if (($shape.Width * $shape.Height) -gt 0.7 * $slideArea) { continue }  # full-bleed media
            if ($shape.Left -ge $ZoneRight -or ($shape.Left + $shape.Width) -le $ZoneLeft) { continue }
            if ($shape.Top -ge $ZoneBottom -or ($shape.Top + $shape.Height) -le $ZoneTop) { continue }
            return $true
        } catch { }
    }
    return $false
}

function Add-BrandIcon {
    param(
        $Slide,
        [string]$IconPath,
        [int]$Diameter,
        [System.Collections.Generic.List[string]]$Actions,
        [string]$WatermarkPath = '',
        [switch]$Force
    )
    $slideWidth = $Slide.Parent.PageSetup.SlideWidth
    # The icon asset carries a baked 10% shadow padding on each side.
    $pictureSize = [Math]::Round($Diameter * 1280 / 1024)
    $edgePadding = [Math]::Round($pictureSize * 0.1)
    $left = [Math]::Round($slideWidth - 14 - $pictureSize + $edgePadding)
    $top = 14 - $edgePadding
    $resolvedIconPath = $IconPath
    if (-not $Force) {
        if (Test-IconZoneOccupied -Slide $Slide -ZoneLeft $left -ZoneTop $top `
                -ZoneRight $slideWidth -ZoneBottom ($top + $pictureSize)) {
            if ([string]::IsNullOrWhiteSpace($WatermarkPath)) {
                $Actions.Add('IconSkipped(zone occupied)') | Out-Null
                return
            }
            # Pages whose own text reaches the corner still get brand coverage
            # through the 45%-opacity watermark variant at the same position.
            $resolvedIconPath = $WatermarkPath
        }
    }
    Invoke-WithComRetry -Action {
        # Remove same-name icons first so COM retries and re-runs stay
        # idempotent (never two stacked icons).
        $stale = @($Slide.Shapes | Where-Object { $_.Name -eq 'sciman-brand-icon' })
        for ($i = $stale.Count - 1; $i -ge 0; $i--) { $stale[$i].Delete() }
        $picture = $Slide.Shapes.AddPicture($resolvedIconPath, 0, -1, $left, $top, $pictureSize, $pictureSize)
        $picture.Name = 'sciman-brand-icon'
    }.GetNewClosure() | Out-Null
    if ($Force) {
        $Actions.Add('IconAdded(watermark, uniform)') | Out-Null
    } elseif ($resolvedIconPath -eq $WatermarkPath) {
        $Actions.Add('IconAdded(watermark, zone occupied)') | Out-Null
    } else {
        $Actions.Add('IconAdded') | Out-Null
    }
}

$application = New-PowerPointApplication
$presentation = $null
$slideRecords = New-Object System.Collections.Generic.List[object]
try {
    $presentation = Invoke-WithComRetry -Action {
        $application.Presentations.Open($workingPptx, $false, $false, $false)
    }
    $slideCount = $presentation.Slides.Count
    Write-Host "Opened $workingPptx with $slideCount slides."

    for ($index = 1; $index -le $slideCount; $index++) {
        # Whole-slide processing is retried as one unit: PowerPoint COM
        # occasionally rejects calls transiently (RPC_E_CALL_REJECTED) while
        # it is busy.  Every operation inside is idempotent, so a retry can
        # safely re-apply the same slide.
        $record = Invoke-WithComRetry -MaxRetries 5 -DelayMs 800 -Action {
            $slide = $presentation.Slides.Item($index)
            $slideText = Get-SlidePlainText -Slide $slide
            $role = Get-SlideRole -Slide $slide -SlideText $slideText
            $actions = New-Object System.Collections.Generic.List[string]

            switch ($role) {
                'Content' {
                    # By design content pages carry no brand icon at all;
                    # brand coverage lives on the special and divider pages
                    # only. Remove any icon left by an earlier generation so
                    # re-runs stay idempotent.
                    Invoke-WithComRetry -Action {
                        $stale = @($Slide.Shapes | Where-Object { $_.Name -eq 'sciman-brand-icon' })
                        for ($i = $stale.Count - 1; $i -ge 0; $i--) { $stale[$i].Delete() }
                        if ($stale.Count -gt 0) { $actions.Add('IconRemoved(stale)') | Out-Null }
                        else { $actions.Add('NoIcon (content page, by design)') | Out-Null }
                    }.GetNewClosure() | Out-Null
                }
                'Blank' {
                    # Original blank page: left untouched by design.
                    $actions.Add('Unchanged (original blank page)') | Out-Null
                }
                'Divider' {
                    Add-BrandIcon -Slide $slide -IconPath $iconShadowPath -Diameter $IconDiameterPt -Actions $actions -WatermarkPath $iconWatermarkPath
                    Update-DividerSlide -Slide $slide -FontSize $DividerFontSize -Actions $actions
                }
                default {
                    Set-BackgroundImage -Slide $slide -ImagePath $bgPath -Actions $actions
                    Remove-LegacyIconShapes -Slide $slide -Actions $actions
                    Update-SpecialSlideText -Slide $slide -Role $role -Actions $actions
                }
            }

            [pscustomobject]@{
                Slide   = $index
                Role    = $role
                Actions = ($actions -join '; ')
                Status  = 'Ok'
            }
        }.GetNewClosure()
        $slideRecords.Add($record) | Out-Null
        Write-Host ("slide {0,2} {1,-8} {2}" -f $record.Slide, $record.Role, $record.Actions)
    }

    Invoke-WithComRetry -Action { $presentation.Save() } | Out-Null

    foreach ($record in $slideRecords) {
        $pngPath = Join-Path $pageImageDir ('slide-{0:d2}.png' -f $record.Slide)
        Invoke-WithComRetry -Action {
            $presentation.Slides.Item($record.Slide).Export($pngPath, 'PNG', 1536, 864)
        }.GetNewClosure() | Out-Null
    }

    $pdfPath = Join-Path $deliveryDir "$stem.brand.pdf"
    Invoke-WithComRetry -Action { $presentation.SaveAs($pdfPath, 32) } | Out-Null  # ppSaveAsPDF
    Write-Host "Exported PDF: $pdfPath"
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { }
        Release-ComObjectSafe $presentation
    }
    if ($null -ne $application) {
        try { $application.Quit() } catch { }
        Release-ComObjectSafe $application
    }
}

# --- delivery reports -------------------------------------------------------
$inputHash = (Get-FileHash -LiteralPath $PptxPath -Algorithm SHA256).Hash
$outputHash = (Get-FileHash -LiteralPath $workingPptx -Algorithm SHA256).Hash

$csvRows = foreach ($record in $slideRecords) {
    [pscustomobject]@{
        页码 = $record.Slide
        角色 = $record.Role
        动作 = $record.Actions
        状态 = $record.Status
    }
}
Write-Utf8BomCsv -InputObject @($csvRows) -Path (Join-Path $reportDir 'brand-visual-refresh-report.csv')

function Get-AssetFingerprint {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $item = Get-Item -LiteralPath $Path
    return [pscustomobject]@{
        path   = $Path
        bytes  = $item.Length
        sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    }
}

$manifest = [pscustomobject]@{
    generatedAt   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    mode          = 'BrandVisualRefresh'
    source        = $PptxPath
    sourceSha256  = $inputHash
    output        = $workingPptx
    outputSha256  = $outputHash
    pdf           = (Join-Path $deliveryDir "$stem.brand.pdf")
    deliveryRoot  = $deliveryRoot
    deliveryStatus = 'PendingManualVisualReview'
    assets        = [pscustomobject]@{
        background      = Get-AssetFingerprint $bgPath
        iconShadow      = Get-AssetFingerprint $iconShadowPath
        iconWatermark   = Get-AssetFingerprint $iconWatermarkPath
        generatorScript = 'tools/generate_brand_assets.py'
        applyScript     = 'tools/Apply-PptxBrandVisualRefresh.ps1'
    }
    dividerFontSize = $DividerFontSize
    iconDiameterPt  = $IconDiameterPt
    slides        = @($slideRecords | ForEach-Object {
        [pscustomobject]@{ slide = $_.Slide; role = $_.Role; actions = $_.Actions; status = $_.Status }
    })
}
$manifestPath = Join-Path $deliveryRoot 'review-manifest.json'
$manifestJson = $manifest | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($manifestPath), $manifestJson, (New-Object System.Text.UTF8Encoding($false)))

$roleSummary = $slideRecords | Group-Object Role | ForEach-Object { "$($_.Name) x$($_.Count)" }
$summaryLines = @(
    '# 品牌视觉刷新结果总览',
    '',
    "- 生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "- 输入：$PptxPath",
    "- 交付物：$workingPptx",
    "- PDF：$(Join-Path $deliveryDir "$stem.brand.pdf")",
    "- 页面图片：$pageImageDir",
    "- 版本目录：$deliveryRoot（已存在历史版本 $(@($existing).Count) 个，本次 v$nextVersion）",
    "- 页面角色：$($roleSummary -join '，')",
    '',
    '## 本次动作',
    '- 仅封面/END/资源页/网盘群页替换品牌背景图（右上角已合成新公众号图标），其余页面保持白底',
    '- 删除蓝底特殊页上的旧公众号图标，避免与新背景重复',
    '- 封面：标题浅金（保留行楷）、副标题近白；三行联系信息 22pt 左对齐、文本块随最长行收窄后整体水平居中，"标签："浅金加粗 + 内容浅蓝灰，雅黑+Consolas',
    '- END：END 浅金加粗、正文白色、"禁止"金色加粗（按字符定位保留高亮）、长链接天蓝 18pt 单行',
    '- 资源页：页标题浅金加粗、网盘名白色加粗、链接天蓝 18pt 单行、说明文字浅蓝灰',
    '- 网盘群页：标题浅金加粗、群号数字行天蓝、说明行白色（版式与字号保持原样）',
    '- 分隔页（第N课时/比热容（N）/比热容公式/拓展）：不加背景图，白底 + 统一微软雅黑加粗 ${DividerFontSize}pt，课时/拓展用强调红、比热容标题用标题蓝，文字文本框内居中且文本框页面居中，右上角不透明图标',
    '- 正文等其余页面：右上角统一 45% 半透明水印图标（约 ${IconDiameterPt}pt，位置固定），保证每页品牌露出且不与正文抢读',
    '- 原空白页保持原样（不加背景、不加图标）',
    '',
    '## 复核方式',
    '- 逐页 PNG（1536x864）+ PDF 已生成，需人工/视觉模型复核后才能标记交付状态',
    '- review-manifest.json 记录输入/输出 SHA256、素材指纹与每页动作清单',
    ''
)
Write-Utf8BomText -Text ($summaryLines -join "`r`n") -Path (Join-Path $deliveryRoot 'summary.md')

Write-Host "Delivery folder: $deliveryRoot"
Write-Host 'Brand visual refresh completed.'
