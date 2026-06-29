<#
.SYNOPSIS
  Normalize junior-middle-school physics PPT visual style with low-risk PowerPoint COM automation.

.DESCRIPTION
  This script opens .pptx files in Microsoft PowerPoint, applies low-risk style normalization,
  and saves normalized copies to an output folder. It does not change text content,
  animations, picture crops, or video links. It may expand eligible text boxes horizontally
  to prevent forced wrapping after font normalization.

.PARAMETER InputPath
  Path to a .pptx/.pptm file or a directory containing PPT files.

.PARAMETER OutputDir
  Directory where normalized copies and reports are saved.

.PARAMETER Recurse
  Search subdirectories when InputPath is a directory.

.PARAMETER ReportOnly
  Only generate a style issue report; do not modify any file.

.PARAMETER NoBackup
  Skip copying original files to _backup_originals.

.PARAMETER NoPdf
  Skip exporting the normalized PPTX to a same-name PDF.

.PARAMETER UpdateMaster
  Also normalize the slide master text styles.

.EXAMPLE
  .\Normalize-PhysicsPpt.ps1 -InputPath "D:\课件" -OutputDir "D:\输出" -Recurse

.EXAMPLE
  .\Normalize-PhysicsPpt.ps1 -InputPath "D:\课件\物理.pptx" -OutputDir "D:\输出" -ReportOnly

.NOTES
  Requires Windows + Microsoft PowerPoint desktop app.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [switch]$Recurse,
    [switch]$ReportOnly,
    [switch]$NoBackup,
    [switch]$NoPdf,
    [switch]$UpdateMaster,
    [switch]$Force,
    [string]$ImageOutputDir,

    [string]$FilePattern = '*.ppt*',

    [ValidateRange(0, 5)]
    [int]$FileRetryCount = 1,

    [ValidateRange(0, 60000)]
    [int]$FileRetryDelayMs = 2000,

    [ValidateRange(1, 8)]
    [int]$DegreeOfParallelism = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

# --- Office Enum Constants ---
$script:MsoTrue  = -1
$script:MsoFalse = 0
$script:MsoPlaceholder = 14
$script:MsoPicture = 13
$script:MsoTable  = 19
$script:MsoGroup  = 6
$script:MsoMedia  = 16
$script:MsoTextEffect = 15
$script:PpPlaceholderTitle = 1
$script:PpPlaceholderCenterTitle = 3
$script:PpAlignLeft = 1
$script:PpAlignCenter = 2
$script:MsoAnimEffectSplit = 16
$script:MsoAnimationLevelNone = 0
$script:MsoAnimTriggerOnPageClick = 1

function Convert-HexToRgbInt {
    param([Parameter(Mandatory = $true)][string]$Hex)
    $h = $Hex.Trim().TrimStart('#')
    if ($h.Length -ne 6 -or $h -notmatch '^[0-9A-Fa-f]{6}$') {
        throw "Invalid hex color: $Hex"
    }
    $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
    return ($r + ($g * 256) + ($b * 65536))
}

function Invoke-WithComRetry {
    param([scriptblock]$Action, [int]$MaxRetries = 2, [int]$DelayMs = 500)
    $lastError = $null
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            return & $Action
        } catch {
            $lastError = $_
            if ($attempt -lt $MaxRetries) { Start-Sleep -Milliseconds $DelayMs }
        }
    }
    throw $lastError
}

function Get-ComFailureCategory {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $exception = $ErrorRecord.Exception
    $hresult = if ($null -ne $exception) { $exception.HResult } else { 0 }

    switch ($hresult) {
        -2147418111 { return 'PowerPointBusyOrRejectedCall' } # 0x80010001 RPC_E_CALL_REJECTED
        -2147024864 { return 'FileInUseOrSharingViolation' } # 0x80070020
        -2147221164 { return 'PowerPointComNotRegistered' }  # 0x80040154
        -2147287038 { return 'FileNotFoundOrUnavailable' }   # 0x80030002
        default {
            if ($null -ne $exception -and $exception.Message -match 'PowerPoint|COM|RPC|rejected|busy') {
                return 'PowerPointComFailure'
            }
            return 'UnhandledFailure'
        }
    }
}

function Format-ComFailureDetails {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $category = Get-ComFailureCategory -ErrorRecord $ErrorRecord
    $hresult = if ($null -ne $ErrorRecord.Exception) { ('0x{0:X8}' -f ($ErrorRecord.Exception.HResult -band 0xFFFFFFFF)) } else { 'n/a' }
    return "$category [$hresult]: $($ErrorRecord.Exception.Message)"
}

function Test-IsRetryablePresentationFailure {
    param(
        [string]$Category,
        [string]$Message
    )

    if ($Category -in @('PowerPointBusyOrRejectedCall', 'PowerPointComFailure', 'FileInUseOrSharingViolation')) {
        return $true
    }

    if ($Category -eq 'UnhandledFailure' -and $Message -match 'null-valued expression|RPC|COM|PowerPoint|busy|rejected') {
        return $true
    }

    return $false
}

# --- Load style configuration from JSON ---
$ConfigPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'config\physics-ppt-style.config.json'
$script:ConfigJson = $null
if (Test-Path -LiteralPath $ConfigPath) {
    $script:ConfigJson = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
} else {
    Write-Warning "Config file not found: $ConfigPath — using built-in defaults."
}

function Get-ConfigValue {
    param([string]$Category, [string]$Key, $Default)
    if ($null -ne $script:ConfigJson) {
        $categoryProp = $script:ConfigJson.PSObject.Properties[$Category]
        if ($null -ne $categoryProp -and $null -ne $categoryProp.Value) {
            $keyProp = $categoryProp.Value.PSObject.Properties[$Key]
            if ($null -ne $keyProp -and $null -ne $keyProp.Value) {
                return $keyProp.Value
            }
        }
    }
    return $Default
}

$script:VideoSlideKeywords = @(Get-ConfigValue 'rules' 'videoSlideKeywords' @('视频', '播放', '观察视频'))
if ($script:VideoSlideKeywords -is [string]) { $script:VideoSlideKeywords = @($script:VideoSlideKeywords) }
$script:VideoSlideKeywords = @($script:VideoSlideKeywords | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
# Pre-compile keyword regex for Test-IsVideoSlide performance
$script:VideoKeywordPattern = if ($script:VideoSlideKeywords.Count -gt 0) {
    '(' + (($script:VideoSlideKeywords | ForEach-Object { [regex]::Escape([string]$_) }) -join '|') + ')'
} else {
    $null
}

$script:FormulaWhitelist = @()
if ($null -ne $script:ConfigJson) {
    $whitelistProp = $script:ConfigJson.PSObject.Properties['formulaWhitelist']
    if ($null -ne $whitelistProp -and $null -ne $whitelistProp.Value) {
        $script:FormulaWhitelist = @($whitelistProp.Value)
    }
}

$script:Style = [pscustomobject]@{
    FontChinese           = Get-ConfigValue 'fonts' 'chinese'           '微软雅黑'
    FontLatin             = Get-ConfigValue 'fonts' 'latin'             'Arial'
    FontMath              = Get-ConfigValue 'fonts' 'math'              'Cambria Math'
    SizeTitle1            = Get-ConfigValue 'fontSizes' 'title1'        46
    SizeSectionTitle      = Get-ConfigValue 'fontSizes' 'sectionTitle'  56
    SizeBody              = Get-ConfigValue 'fontSizes' 'body'          32
    SizeMinimum           = Get-ConfigValue 'fontSizes' 'minimum'       24
    SizeTableHeader       = Get-ConfigValue 'fontSizes' 'tableHeader'   30
    SizeTableBody         = Get-ConfigValue 'fontSizes' 'tableBody'     28
    SizeFormulaInline     = Get-ConfigValue 'fontSizes' 'formulaInline' 34
    SizeFormulaStandalone = Get-ConfigValue 'fontSizes' 'formulaStandalone' 38
    SizeFormulaCore       = Get-ConfigValue 'fontSizes' 'formulaCore'   42
    ColorWhite            = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'white'         '#FFFFFF')
    ColorBlack            = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'black'         '#000000')
    ColorBody             = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'body'          '#000000')
    ColorSectionTitle     = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'sectionTitle'  '#0000FF')
    ColorExtensionTitle   = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'extensionTitle' '#FF0000')
    ColorFormulaBlue      = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'formulaBlue'   '#0066CC')
    ColorYellowFill       = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'yellowFill'    '#FFF2CC')
    ColorYellowBorder     = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'yellowBorder'  '#D6A300')
}

$script:AllowTextBoxWidthExpansion = [bool](Get-ConfigValue 'rules' 'allowTextBoxWidthExpansion' $true)
$script:FormulaTextStyleDefault = [bool](Get-ConfigValue 'rules' 'formulaTextStyleDefault' $true)
$script:FormulaConversionDefault = [bool](Get-ConfigValue 'rules' 'formulaConversionDefault' $false)

# --- Yellow-ish RGB range for highlight-box detection ---
# Tolerance band: R > 200, G > 200, B < 180 (covers most yellow/cream fills)
function Test-IsYellowishFill {
    param([int]$Rgb)
    $r = $Rgb -band 0xFF
    $g = ($Rgb -shr 8) -band 0xFF
    $b = ($Rgb -shr 16) -band 0xFF
    return ($r -gt 200 -and $g -gt 200 -and $b -lt 180)
}

# --- Report ---
$script:ReportRows = New-Object System.Collections.Generic.List[object]
$script:CurrentFilePath = ''

function Add-ReportRow {
    param(
        [string]$File,
        [string]$FilePath = '',
        [int]$SlideNumber,
        [string]$ShapeName,
        [string]$Issue,
        [string]$Details
    )
    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        $FilePath = $script:CurrentFilePath
    }
    $script:ReportRows.Add([pscustomobject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        File = $File
        FilePath = $FilePath
        Slide = $SlideNumber
        Shape = $ShapeName
        Issue = $Issue
        Details = $Details
    }) | Out-Null
}

function Reset-ReportRowsToCount {
    param([int]$Count)
    while ($script:ReportRows.Count -gt $Count) {
        $script:ReportRows.RemoveAt($script:ReportRows.Count - 1)
    }
}

function Restart-PowerPointApplication {
    param($Current)

    if ($null -ne $Current) {
        try { $Current.Quit() | Out-Null } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Current) | Out-Null } catch { }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if ($FileRetryDelayMs -gt 0) { Start-Sleep -Milliseconds $FileRetryDelayMs }

    $next = New-Object -ComObject PowerPoint.Application
    $next.Visible = $script:MsoTrue
    return $next
}

function Get-PptFiles {
    param([string]$Path, [string]$Pattern, [switch]$Recurse)
    if (-not (Test-Path -LiteralPath $Path)) { throw "InputPath not found: $Path" }
    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) {
        $opt = @{ LiteralPath = $item.FullName; Filter = $Pattern; File = $true }
        if ($Recurse) { $opt.Recurse = $true }
        return Get-ChildItem @opt | Where-Object { $_.Name -notlike '~$*' -and $_.Extension -in '.pptx', '.pptm' }
    }
    if ($item.Extension -notin '.pptx', '.pptm') { throw "Only .pptx/.pptm files are supported: $($item.FullName)" }
    return @($item)
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
        if ($null -ne $Shape -and -not [string]::IsNullOrWhiteSpace([string]$Shape.Name)) {
            return [string]$Shape.Name
        }
    } catch { }
    return '(unknown)'
}

function Test-ShapeHasTable {
    param($Shape)
    try {
        return [bool]$Shape.HasTable
    } catch { }
    return $false
}

function Test-IsTitleShape {
    param($Shape)
    # Primary: PowerPoint placeholder type (most reliable)
    try {
        if ($Shape.Type -eq $script:MsoPlaceholder) {
            $phType = $Shape.PlaceholderFormat.Type
            if ($phType -eq $script:PpPlaceholderTitle -or $phType -eq $script:PpPlaceholderCenterTitle) {
                return $true
            }
            # Subtitle or body placeholder → not a title
            return $false
        }
    } catch { }
    # Fallback heuristic: short text near the top of the slide
    $text = Get-ShapeText $Shape
    if ($text.Length -le 24 -and $Shape.Top -lt 90) { return $true }
    return $false
}

function Test-IsSectionTitleSlide {
    param($Slide)
    $textShapes = 0
    $mainText = ''
    $hasLargeVisual = $false
    foreach ($shape in $Slide.Shapes) {
        try {
            $text = Get-ShapeText $shape
            if ($shape.TextFrame2.HasText -eq $script:MsoTrue -and -not [string]::IsNullOrWhiteSpace($text)) {
                $textShapes++
                $mainText = $text
            }
        } catch { }
        try {
            if ($shape.Type -eq $script:MsoPicture -or $shape.Type -eq $script:MsoMedia -or $shape.Type -eq $script:MsoGroup) {
                if (($shape.Width * $shape.Height) -gt 120000) { $hasLargeVisual = $true }
            }
        } catch { }
    }
    $cleanText = ($mainText -replace '\s+', '').Trim()
    if ($cleanText.Length -eq 0 -or $cleanText.Length -gt 18) { return $false }
    if ($cleanText -match 'https?|www|网盘|QQ群|下载|地址|[，。；：、,;:]') { return $false }
    return ($textShapes -eq 1 -and -not $hasLargeVisual)
}

function Get-SlideTextSummary {
    param($Slide)
    $textBuilder = New-Object System.Text.StringBuilder
    $textShapeCount = 0
    foreach ($shape in $Slide.Shapes) {
        try {
            $text = Get-ShapeText $shape
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $textShapeCount++
                [void]$textBuilder.AppendLine($text)
            }
        } catch { }
    }
    return [pscustomobject]@{
        Text = $textBuilder.ToString()
        TextShapeCount = $textShapeCount
    }
}

function Test-IsExerciseOrQuestionText {
    param([string]$Text)
    $plain = ($Text -replace '\s+', '').Trim()
    if ([string]::IsNullOrWhiteSpace($plain)) { return $false }

    # Exercise pages can look like long appendix text when formulas are OLE objects.
    if ($plain -match '^[0-9０-９一二三四五六七八九十]+[、.．]') { return $true }
    if ($plain -match '求[:：]') { return $true }
    if ($plain -match '[问则].{0,24}(为|是|多少|几|何)[（(]?[A-DＡ-Ｄ]?') { return $true }
    if ($plain -match 'A[.．、].+B[.．、].+C[.．、].+D[.．、]') { return $true }
    return $false
}

function Test-IsExtensionSectionText {
    param([string]$Text)
    $plain = ($Text -replace '\s+', '').Trim()
    if ([string]::IsNullOrWhiteSpace($plain)) { return $false }
    return ($plain -match '^(拓展|扩展|拓展提升|拓展训练|选学|能力提升)$')
}

function Test-IsEmptySlideCandidate {
    param($Slide)

    $summary = Get-SlideTextSummary -Slide $Slide
    if (-not [string]::IsNullOrWhiteSpace($summary.Text)) { return $false }

    try {
        if ($Slide.Shapes.Count -eq 0) { return $true }
    } catch {
        return $false
    }

    foreach ($shape in $Slide.Shapes) {
        try {
            if ($shape.Visible -eq $script:MsoFalse) { continue }
        } catch { }

        try {
            if ($shape.Type -eq $script:MsoPlaceholder) {
                continue
            }
        } catch { }

        try {
            if (($shape.Width * $shape.Height) -gt 100) {
                return $false
            }
        } catch {
            return $false
        }
    }

    return $true
}

function Get-SlideKind {
    param($Slide, [int]$SlideNumber)
    if ($SlideNumber -eq 1) { return 'Cover' }

    $summary = Get-SlideTextSummary -Slide $Slide
    $joined = $summary.Text
    if ([string]::IsNullOrWhiteSpace($joined)) { return 'Normal' }

    if ($joined -match '(?i)\bEND\b') { return 'Ending' }
    if ($joined -match '课件下载|下载地址|网盘|QQ群|Q群|知乎主页|公众号|sciman|zhihu\.com|pan\.baidu|alipan|quark') {
        return 'Resource'
    }
    if (Test-IsSectionTitleSlide -Slide $Slide) {
        if (Test-IsExtensionSectionText -Text $joined) {
            return 'ExtensionSection'
        }
        return 'ContentSection'
    }

    $hasLargeVisual = $false
    foreach ($shape in $Slide.Shapes) {
        try {
            if ($shape.Type -eq $script:MsoPicture -or $shape.Type -eq $script:MsoMedia -or $shape.Type -eq $script:MsoGroup) {
                if (($shape.Width * $shape.Height) -gt 120000) { $hasLargeVisual = $true }
            }
        } catch { }
    }
    $plain = ($joined -replace '\s+', '')
    if (Test-IsExerciseOrQuestionText -Text $joined) {
        return 'Exercise'
    }
    if ($summary.TextShapeCount -le 2 -and -not $hasLargeVisual -and $plain.Length -gt 80) {
        return 'AppendixText'
    }

    return 'Normal'
}

function Test-IsUtilitySlide {
    param($Slide, [int]$SlideNumber)
    return ((Get-SlideKind -Slide $Slide -SlideNumber $SlideNumber) -in @('Cover', 'Ending', 'Resource'))
}

function Get-SpecialSlidePreserveIssue {
    param([string]$SlideKind)
    switch ($SlideKind) {
        'Cover'        { return 'CoverSlideStylePreserved' }
        'Ending'       { return 'EndingSlideStylePreserved' }
        'Resource'     { return 'ResourceSlideStylePreserved' }
        'AppendixText' { return 'AppendixTextSlideStylePreserved' }
        default        { return $null }
    }
}

function Get-SpecialSlidePreserveDetails {
    param([string]$SlideKind)
    switch ($SlideKind) {
        'Cover'        { return 'Cover slide detected; original typography and layout are preserved.' }
        'Ending'       { return 'Ending slide detected; original typography and layout are preserved.' }
        'Resource'     { return 'Resource/download slide detected; original typography and layout are preserved.' }
        'AppendixText' { return 'Appendix explanation slide detected; original text style is preserved.' }
        default        { return '' }
    }
}

function Test-IsLargePictureShape {
    param($Shape)
    try {
        return ($Shape.Type -eq $script:MsoPicture -and ($Shape.Width * $Shape.Height) -gt 120000)
    } catch { }
    return $false
}

function Test-IsVideoSlide {
    param($Slide)
    foreach ($shape in $Slide.Shapes) {
        try {
            if ($shape.Type -eq $script:MsoMedia) { return $true }
        } catch { }
        $text = Get-ShapeText $shape
        if ($null -ne $script:VideoKeywordPattern -and $text -match $script:VideoKeywordPattern) { return $true }
    }
    return $false
}

function Test-IsFormulaCandidateText {
    param([string]$Text)
    $t = ($Text -replace '\s+', '')
    if ([string]::IsNullOrWhiteSpace($t)) { return $false }
    if ($t.Length -gt 80) { return $false }
    if ($t -match '[=ηΩ]') { return $true }
    if ($t -match '([PWUIRFSη]|W有|W总|W额|G物|G动)[=＝].*[/÷]') { return $true }
    if ($t -match '(W有|W总|W额|G物|G动|R[12]|U[12]|I[12]|P[12])') { return $true }
    return $false
}

function Get-NormalizedFormulaText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return (($Text -replace '\s+', '') -replace '＝', '=').Trim()
}

function Get-FormulaCandidateProfile {
    param([string]$Text)

    $normalized = Get-NormalizedFormulaText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return [pscustomobject]@{
            Kind = 'Empty'
            Risk = 'ReviewOnly'
            Normalized = ''
            Length = 0
            Reason = 'empty'
        }
    }

    $lines = @($Text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $hasEquation = $normalized -match '='
    $hasDivision = $normalized -match '[/÷]'
    $hasGreekOrUnit = $normalized -match '[ηΩρλμ]'
    $hasSubscriptToken = $normalized -match '(W有|W总|W额|G物|G动|R[12]|U[12]|I[12]|P[12])'
    $hasSentencePunctuation = $Text -match '[，。；：、？！,;:?!]'

    $withoutKnownChineseTokens = $normalized -replace '(W有|W总|W额|G物|G动)', ''
    $withoutKnownChineseTokens = $withoutKnownChineseTokens -replace '[有总额物动]', ''
    $hasUnexpectedChinese = $withoutKnownChineseTokens -match '[一-龥]'

    $kind = 'ShortMathText'
    if ($hasEquation -and $hasDivision) {
        $kind = 'LinearFractionEquation'
    } elseif ($hasEquation) {
        $kind = 'EquationText'
    } elseif ($hasDivision) {
        $kind = 'FractionLikeText'
    } elseif ($hasSubscriptToken) {
        $kind = 'SubscriptLikeText'
    } elseif ($hasGreekOrUnit) {
        $kind = 'SymbolFormulaText'
    }

    $reasons = New-Object System.Collections.Generic.List[string]
    if ($normalized.Length -gt 48) { $reasons.Add('too-long-for-auto-style') | Out-Null }
    if ($lines.Count -gt 2) { $reasons.Add('multi-line') | Out-Null }
    if ($hasSentencePunctuation) { $reasons.Add('sentence-punctuation') | Out-Null }
    if ($hasUnexpectedChinese) { $reasons.Add('unexpected-chinese-text') | Out-Null }

    $risk = 'ReviewOnly'
    if ($reasons.Count -eq 0 -and ($hasEquation -or $hasDivision -or $hasGreekOrUnit -or $hasSubscriptToken)) {
        $risk = 'LowRiskStandaloneText'
    }

    $reason = if ($reasons.Count -gt 0) { $reasons -join ';' } else { 'low-risk-standalone-text' }
    return [pscustomobject]@{
        Kind = $kind
        Risk = $risk
        Normalized = $normalized
        Length = $normalized.Length
        Reason = $reason
    }
}

function Get-FormulaRuleValue {
    param($Rule, [string]$Name, [string]$Default = '')
    if ($null -eq $Rule) { return $Default }
    $prop = $Rule.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
    return [string]$prop.Value
}

function Get-FormulaWhitelistMatch {
    param($Profile)

    if ($null -eq $Profile -or [string]::IsNullOrWhiteSpace([string]$Profile.Normalized)) {
        return $null
    }

    foreach ($rule in @($script:FormulaWhitelist)) {
        $pattern = Get-FormulaRuleValue -Rule $rule -Name 'sourcePattern'
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        try {
            if ([string]$Profile.Normalized -match $pattern) {
                return $rule
            }
        } catch {
            continue
        }
    }

    return $null
}

function Add-FormulaWhitelistReport {
    param(
        [string]$FileName,
        [int]$SlideNumber,
        [string]$ShapeName,
        $Match
    )

    if ($null -eq $Match) { return }

    $name = Get-FormulaRuleValue -Rule $Match -Name 'name' -Default 'formula'
    $target = Get-FormulaRuleValue -Rule $Match -Name 'targetUnicodeMath'
    $targetTex = Get-FormulaRuleValue -Rule $Match -Name 'targetTex'
    $note = Get-FormulaRuleValue -Rule $Match -Name 'note'
    Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Issue 'FormulaWhitelistCandidate' `
        -Details ("name={0}; targetUnicodeMath={1}; targetTex={2}; note={3}" -f $name, $target, $targetTex, $note)

    if ($script:FormulaConversionDefault) {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Issue 'FormulaConversionPending' `
            -Details 'Whitelist matched; conversion flag is enabled, but OfficeMath conversion is not implemented in this batch.'
    } else {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Issue 'FormulaConversionSkipped' `
            -Details 'Whitelist matched; formulaConversionDefault=false keeps semantic conversion disabled.'
    }
}

function Add-FormulaCandidateReport {
    param(
        [string]$FileName,
        [int]$SlideNumber,
        [string]$ShapeName,
        [string]$Text
    )

    $profile = Get-FormulaCandidateProfile -Text $Text
    Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Issue 'FormulaCandidate' -Details $Text
    Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Issue 'FormulaCandidateClass' `
        -Details ("risk={0}; kind={1}; length={2}; reason={3}" -f $profile.Risk, $profile.Kind, $profile.Length, $profile.Reason)
    $whitelistMatch = Get-FormulaWhitelistMatch -Profile $profile
    Add-FormulaWhitelistReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Match $whitelistMatch
    return $profile
}

function Test-IsLowRiskFormulaProfile {
    param($Profile)
    return ($null -ne $Profile -and $Profile.Risk -eq 'LowRiskStandaloneText')
}

function Resolve-FormulaTargetSize {
    param($Profile)
    if ($null -eq $Profile) { return $script:Style.SizeFormulaInline }
    if ($Profile.Kind -in @('EquationText', 'LinearFractionEquation', 'FractionLikeText')) {
        return $script:Style.SizeFormulaStandalone
    }
    return $script:Style.SizeFormulaInline
}

function Set-FormulaTextStyle {
    param(
        $Shape,
        $Profile,
        [int]$SlideNumber,
        [string]$FileName
    )

    try {
        $textRange = $Shape.TextFrame2.TextRange
        $font = $textRange.Font
        $targetSize = Resolve-FormulaTargetSize -Profile $Profile
        $safeSize = Resolve-SafeFontSize -TextRange $textRange -TargetSize $targetSize -FileName $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name

        $font.Name = $script:Style.FontMath
        $font.NameFarEast = $script:Style.FontChinese
        $font.Size = $safeSize
        $font.Bold = $script:MsoFalse
        $font.Fill.Visible = $script:MsoTrue
        $font.Fill.ForeColor.RGB = $script:Style.ColorFormulaBlue
        $textRange.ParagraphFormat.Alignment = $script:PpAlignCenter

        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'FormulaTextStyleNormalized' `
            -Details ("kind={0}; size={1:N1} pt; font={2}; color=#0066CC; text preserved." -f $Profile.Kind, $safeSize, $script:Style.FontMath)
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'FormulaTextStyleFailed' -Details $_.Exception.Message
    }
}

function Get-TextRangeFontSize {
    param($TextRange)
    try {
        $size = [double]$TextRange.Font.Size
        if ($size -gt 0) { return $size }
    } catch { }
    return $null
}

function Resolve-SafeFontSize {
    param(
        $TextRange,
        [double]$TargetSize,
        [string]$FileName = '',
        [int]$SlideNumber = 0,
        [string]$ShapeName = '',
        [switch]$ForceTargetSize
    )

    $currentSize = Get-TextRangeFontSize $TextRange
    if ($ForceTargetSize) {
        return $TargetSize
    }
    if ($null -ne $currentSize -and $currentSize -lt $TargetSize) {
        if ($FileName -ne '' -and $currentSize -lt $script:Style.SizeMinimum) {
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
                -Issue 'SmallTextPreserved' -Details "$currentSize pt; not increased to avoid layout overflow."
        }
        return $currentSize
    }
    if ($null -ne $currentSize -and $currentSize -le 72) {
        return $currentSize
    }
    return $TargetSize
}

function Set-TextRangeStyle {
    param(
        $TextRange,
        [double]$Size,
        [int]$Color,
        [bool]$Bold,
        [string]$FileName = '',
        [int]$SlideNumber = 0,
        [string]$ShapeName = '',
        [switch]$ForceTargetSize
    )
    try {
        $font = $TextRange.Font
        $safeSize = Resolve-SafeFontSize -TextRange $TextRange -TargetSize $Size -FileName $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -ForceTargetSize:$ForceTargetSize
        $font.Name = $script:Style.FontLatin
        $font.NameFarEast = $script:Style.FontChinese
        $font.Size = $safeSize
    } catch {
        if ($FileName -ne '') {
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
                -Issue 'TextStyleFailed' -Details $_.Exception.Message
        }
    }
}

function Set-SectionTitleTextStyle {
    param($Shape, [int]$SlideNumber, [string]$FileName)
    try {
        $text = Get-ShapeText $Shape
        $isExtensionSection = Test-IsExtensionSectionText -Text $text
        $targetColor = if ($isExtensionSection) { $script:Style.ColorExtensionTitle } else { $script:Style.ColorSectionTitle }
        $issue = if ($isExtensionSection) { 'ExtensionSectionTitleStyleFixed' } else { 'SectionTitleStyleFixed' }
        $details = if ($isExtensionSection) {
            'Extension section title color and bold style normalized.'
        } else {
            'Section title color and bold style normalized.'
        }
        $font = $Shape.TextFrame2.TextRange.Font
        $font.Bold = $script:MsoTrue
        $font.Fill.Visible = $script:MsoTrue
        $font.Fill.ForeColor.RGB = $targetColor
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue $issue -Details $details
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'SectionTitleStyleFailed' -Details $_.Exception.Message
    }
}

function Test-IsTextBoxExpansionCandidate {
    param($Shape, [string]$Text, [bool]$IsSectionTitleSlide)
    if (-not $script:AllowTextBoxWidthExpansion) { return $false }
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    try {
        if ($Shape.TextFrame2.HasText -ne $script:MsoTrue) { return $false }
        $shapeWidth = [double]$Shape.Width
        $shapeHeight = [double]$Shape.Height
        $shapeRotation = [double]$Shape.Rotation
        if ($shapeWidth -le 0 -or $shapeHeight -le 0) { return $false }
        if ([Math]::Abs($shapeRotation) -gt 0.01) { return $false }
        if (-not $IsSectionTitleSlide -and $shapeHeight -gt ($shapeWidth * 1.25) -and $Text -notmatch '\s') {
            return $false
        }
        if ($Shape.Type -eq $script:MsoPlaceholder) { return $IsSectionTitleSlide }
    } catch { return $false }

    $normalized = ($Text -replace '\s+', ' ').Trim()
    if ($IsSectionTitleSlide) { return $true }
    if ($normalized.Length -gt 28) { return $false }
    return ($normalized -match '[<>=＋+\-*/÷×]|[ufvUV]|[一-龥]{1,8}')
}

function Expand-TextBoxWidthIfNeeded {
    param(
        $Shape,
        [string]$Text,
        [double]$SlideWidth,
        [int]$SlideNumber,
        [string]$FileName,
        [bool]$IsSectionTitleSlide
    )

    $step = 'start'
    if (-not (Test-IsTextBoxExpansionCandidate -Shape $Shape -Text $Text -IsSectionTitleSlide $IsSectionTitleSlide)) {
        return
    }

    try {
        $step = 'text-range'
        $textRange = $Shape.TextFrame2.TextRange
        $step = 'font-size'
        $fontSize = Get-TextRangeFontSize $textRange
        if ($null -eq $fontSize) { return }

        $step = 'width-calc'
        $normalized = ($Text -replace '\s+', ' ').Trim()
        $shapeWidth = [double]$Shape.Width
        $requiredWidth = if ($IsSectionTitleSlide) {
            [Math]::Min($SlideWidth * 0.92, [Math]::Max(($SlideWidth * 0.86), ($normalized.Length * $fontSize * 1.05) + 64))
        } else {
            [Math]::Min($SlideWidth * 0.55, [Math]::Max($shapeWidth, ($normalized.Length * $fontSize * 0.72) + 28))
        }

        $maxWidthByPosition = [Math]::Max(0, $SlideWidth - 24)
        $oldWidth = $shapeWidth
        $targetWidth = [Math]::Min($requiredWidth, $maxWidthByPosition)
        if ($targetWidth -le ($oldWidth + 6)) {
            if ($IsSectionTitleSlide) {
                try { $Shape.TextFrame2.WordWrap = $script:MsoFalse } catch { }
                try { $Shape.TextFrame2.TextRange.ParagraphFormat.Alignment = $script:PpAlignCenter } catch { }
            }
            return
        }

        $step = 'position-calc'
        $oldLeft = [double]$Shape.Left
        if ($IsSectionTitleSlide) {
            $centerX = $oldLeft + ($oldWidth / 2)
            $newLeft = $centerX - ($targetWidth / 2)
            if ($newLeft -lt 12) { $newLeft = 12 }
            if (($newLeft + $targetWidth) -gt ($SlideWidth - 12)) {
                $newLeft = $SlideWidth - 12 - $targetWidth
            }
        } else {
            $newLeft = $oldLeft
            if (($newLeft + $targetWidth) -gt ($SlideWidth - 12)) {
                $targetWidth = $SlideWidth - 12 - $newLeft
            }
        }

        $step = 'apply-position'
        $Shape.Left = [single]$newLeft
        $Shape.Width = [single]$targetWidth
        $step = 'word-wrap'
        try { $Shape.TextFrame2.WordWrap = $script:MsoFalse } catch { }
        $step = 'align'
        $Shape.TextFrame2.TextRange.ParagraphFormat.Alignment = $script:PpAlignCenter
        $step = 'report'
        $anchorMode = if ($IsSectionTitleSlide) { 'center preserved' } else { 'left edge preserved' }
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'TextBoxWidthExpanded' -Details ("{0:N1} -> {1:N1} pt; {2}." -f $oldWidth, $targetWidth, $anchorMode)
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'TextBoxWidthExpandFailed' -Details "$step`: $($_.Exception.Message)"
    }
}

function Normalize-TextShape {
    param($Shape, [string]$Text, [int]$SlideNumber, [string]$FileName, [bool]$IsVideoSlide, [bool]$IsSectionTitleSlide, [double]$SlideWidth)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    if (Test-IsFormulaCandidateText $Text) {
        $profile = Add-FormulaCandidateReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Text $Text
        if ($script:FormulaTextStyleDefault -and (Test-IsLowRiskFormulaProfile -Profile $profile)) {
            Set-FormulaTextStyle -Shape $Shape -Profile $profile -SlideNumber $SlideNumber -FileName $FileName
            Expand-TextBoxWidthIfNeeded -Shape $Shape -Text $Text -SlideWidth $SlideWidth -SlideNumber $SlideNumber -FileName $FileName -IsSectionTitleSlide:$false
        } else {
            $reason = if ($script:FormulaTextStyleDefault) { $profile.Reason } else { 'formula text style normalization disabled by config' }
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'FormulaStyleSkipped' -Details $reason
            Expand-TextBoxWidthIfNeeded -Shape $Shape -Text $Text -SlideWidth $SlideWidth -SlideNumber $SlideNumber -FileName $FileName -IsSectionTitleSlide:$false
        }
        return
    }

    try {
        $isTitle = Test-IsTitleShape $Shape
        $size = if ($IsSectionTitleSlide) { $script:Style.SizeSectionTitle } elseif ($isTitle) { $script:Style.SizeTitle1 } else { $script:Style.SizeBody }
        $bold = [bool]$isTitle
        $color = if ($IsVideoSlide) { $script:Style.ColorWhite } else { $script:Style.ColorBody }
        Set-TextRangeStyle -TextRange $Shape.TextFrame2.TextRange -Size $size -Color $color -Bold $bold `
            -FileName $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -ForceTargetSize:$IsSectionTitleSlide
        if ($IsSectionTitleSlide) {
            Set-SectionTitleTextStyle -Shape $Shape -SlideNumber $SlideNumber -FileName $FileName
        }
        Expand-TextBoxWidthIfNeeded -Shape $Shape -Text $Text -SlideWidth $SlideWidth -SlideNumber $SlideNumber -FileName $FileName -IsSectionTitleSlide:$IsSectionTitleSlide
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'TextStyleFailed' -Details $_.Exception.Message
    }
}

function Get-TextColorRgb {
    param($Shape)
    try {
        return [int]$Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    } catch { }
    return $null
}

function Test-IsRedAnswerShape {
    param($Shape, [string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $clean = ($Text -replace '\s+', '').Trim()
    if ($clean.Length -gt 16) { return $false }
    if ($clean -notmatch '^[\-−]?\d+(\.\d+)?[%％]?([A-Za-z一-龥/]+)?$') { return $false }

    $rgb = Get-TextColorRgb -Shape $Shape
    if ($null -eq $rgb) { return $false }
    $r = $rgb -band 0xFF
    $g = ($rgb -shr 8) -band 0xFF
    $b = ($rgb -shr 16) -band 0xFF
    return ($r -ge 150 -and $g -le 90 -and $b -le 90)
}

function Get-TextShapeInfos {
    param($Slide)
    $items = @()
    foreach ($shape in $Slide.Shapes) {
        try {
            if ($shape.Type -eq $script:MsoGroup) { continue }
            if ($shape.TextFrame2.HasText -ne $script:MsoTrue) { continue }
            $text = Get-ShapeText $shape
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $left = [double]$shape.Left
            $top = [double]$shape.Top
            $width = [double]$shape.Width
            $height = [double]$shape.Height
            $isRedAnswer = [bool](Test-IsRedAnswerShape -Shape $shape -Text $text)
            $items += [pscustomobject]@{
                Shape = $shape
                Text = $text
                Left = $left
                Top = $top
                Width = $width
                Height = $height
                CenterY = ($top + ($height / 2))
                IsRedAnswer = $isRedAnswer
            }
        } catch { }
    }
    return @($items)
}

function Set-AnswerSplitAnimation {
    param($Slide, $Shape, [int]$SlideNumber, [string]$FileName)

    try {
        $shapeId = [int]$Shape.Id
        $sequence = $Slide.TimeLine.MainSequence
        $updated = $false

        for ($effectIndex = 1; $effectIndex -le $sequence.Count; $effectIndex++) {
            $effect = $sequence.Item($effectIndex)
            try {
                if ([int]$effect.Shape.Id -eq $shapeId) {
                    $effect.EffectType = $script:MsoAnimEffectSplit
                    $updated = $true
                }
            } catch { }
        }

        if (-not $updated) {
            [void]$sequence.AddEffect($Shape, $script:MsoAnimEffectSplit, $script:MsoAnimationLevelNone, $script:MsoAnimTriggerOnPageClick)
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'AnswerAnimationAdded' -Details 'Split animation added to high-confidence answer text.'
        } else {
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'AnswerAnimationSet' -Details 'Existing answer animation effect changed to Split.'
        }
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'AnswerAnimationFailed' -Details (Format-ComFailureDetails $_)
    }
}

function Align-AnswerTextBoxes {
    param($Slide, [int]$SlideNumber, [string]$FileName)

    $items = @(Get-TextShapeInfos -Slide $Slide)
    if ($items.Count -lt 2) { return }

    foreach ($answer in @($items | Where-Object { $_.IsRedAnswer })) {
        $target = $null
        $bestScore = [double]::MaxValue
        foreach ($candidate in @($items | Where-Object { -not $_.IsRedAnswer })) {
            if ($candidate.Text.Length -lt 8) { continue }
            if ($candidate.Left -ge $answer.Left) { continue }
            if ($answer.Left -gt ($candidate.Left + $candidate.Width + 80)) { continue }

            $centerDelta = [Math]::Abs($candidate.CenterY - $answer.CenterY)
            $maxAllowedDelta = [Math]::Max(36, [Math]::Max($candidate.Height, $answer.Height))
            if ($centerDelta -gt $maxAllowedDelta) { continue }

            $score = $centerDelta + (($answer.Left - $candidate.Left) / 1000)
            if ($score -lt $bestScore) {
                $target = $candidate
                $bestScore = $score
            }
        }

        if ($null -eq $target) { continue }

        try {
            $oldTop = [double]$answer.Shape.Top
            $newTop = [double]($target.CenterY - ($answer.Height / 2))
            $delta = [Math]::Abs($newTop - $oldTop)
            if ($delta -ge 1 -and $delta -le 24) {
                $answer.Shape.Top = [single]$newTop
                Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $answer.Shape) -Issue 'AnswerTextAligned' -Details ("Top {0:N1} -> {1:N1} pt; aligned to {2}." -f $oldTop, $newTop, (Get-ShapeName $target.Shape))
            } else {
                Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $answer.Shape) -Issue 'AnswerTextAlignmentChecked' -Details ("Already aligned or movement too large; delta {0:N1} pt." -f $delta)
            }
            Set-AnswerSplitAnimation -Slide $Slide -Shape $answer.Shape -SlideNumber $SlideNumber -FileName $FileName
        } catch {
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $answer.Shape) -Issue 'AnswerTextAlignFailed' -Details $_.Exception.Message
        }
    }
}

function Normalize-TableShape {
    param($Shape, [int]$SlideNumber, [string]$FileName)
    Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'TableStyleSkipped' -Details 'Table styles are preserved to avoid cell overflow or row-height changes.'
}

function Normalize-HighlightBox {
    param($Shape, [int]$SlideNumber = 0, [string]$FileName = '')
    try {
        if ($Shape.Fill.Visible -eq $script:MsoTrue) {
            $rgb = $Shape.Fill.ForeColor.RGB
            if (Test-IsYellowishFill $rgb) {
                if ($Shape.TextFrame2.HasText -eq $script:MsoTrue) {
                    $Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $script:Style.ColorBody
                    if ($FileName -ne '') {
                        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'HighlightTextColorFixed' -Details 'Yellow highlight text color was set to body color for readability.'
                    }
                } else {
                    if ($FileName -ne '') {
                        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'HighlightBoxPreserved' -Details 'Yellow highlight fill and border are preserved.'
                    }
                }
            }
        }
    } catch { }
}

function Clear-DecorativeEffects {
    param($Shape)
    try { $Shape.Shadow.Visible = $script:MsoFalse } catch { }
    try { $Shape.Shadow.Transparency = 1 } catch { }
    try { $Shape.Shadow.Blur = 0 } catch { }
    try { $Shape.Shadow.OffsetX = 0 } catch { }
    try { $Shape.Shadow.OffsetY = 0 } catch { }
    try { $Shape.Glow.Radius = 0 } catch { }
    try { $Shape.SoftEdge.Radius = 0 } catch { }
    try {
        if ($Shape.TextFrame2.HasText -eq $script:MsoTrue) {
            $Shape.TextFrame2.TextRange.Font.Shadow.Visible = $script:MsoFalse
            $Shape.TextFrame2.TextRange.Font.Shadow.Transparency = 1
            $Shape.TextFrame2.TextRange.Font.Line.Visible = $script:MsoFalse
        }
    } catch { }
}

function Set-SlideBackground {
    param($Slide, [bool]$IsVideoSlide)
    try {
        $Slide.FollowMasterBackground = $script:MsoFalse
        $Slide.Background.Fill.Solid() | Out-Null
        $Slide.Background.Fill.ForeColor.RGB = $(if ($IsVideoSlide) { $script:Style.ColorBlack } else { $script:Style.ColorWhite })
    } catch { }
}

function Disable-SlideAdvanceOnClick {
    param($Slide, [int]$SlideNumber, [string]$FileName)
    try {
        $Slide.SlideShowTransition.AdvanceOnClick = $script:MsoFalse
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'AdvanceOnClickDisabled' -Details 'Slide transition no longer advances on mouse click.'
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'AdvanceOnClickDisableFailed' -Details (Format-ComFailureDetails $_)
    }
}

function Update-SlideMasterStyle {
    param($Presentation, [string]$FileName)
    try {
        $master = $Presentation.SlideMaster
        foreach ($shape in $master.Shapes) {
            $text = Get-ShapeText $shape
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $isTitle = Test-IsTitleShape $shape
            Set-TextRangeStyle -TextRange $shape.TextFrame2.TextRange `
                -Size $(if ($isTitle) { $script:Style.SizeTitle1 } else { $script:Style.SizeBody }) `
                -Color $script:Style.ColorBody -Bold $isTitle `
                -FileName $FileName -SlideNumber 0 -ShapeName 'SlideMaster'
        }
    } catch {
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName 'SlideMaster' -Issue 'MasterUpdateFailed' -Details $_.Exception.Message
    }
}

function Export-PresentationPdf {
    param(
        $Presentation,
        [string]$PdfPath,
        [string]$FileName
    )

    try {
        $pdfFormat = 32 # ppSaveAsPDF
        Invoke-WithComRetry { $Presentation.SaveAs($PdfPath, $pdfFormat) }
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'PdfExported' -Details $PdfPath
    } catch {
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'PdfExportFailed' -Details (Format-ComFailureDetails $_)
    }
}

function Export-PresentationImages {
    param(
        $Presentation,
        [string]$ImageDir,
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($ImageDir)) { return }
    try {
        if (-not (Test-Path -LiteralPath $ImageDir)) { New-Item -ItemType Directory -Path $ImageDir -Force | Out-Null }
        Invoke-WithComRetry { $Presentation.Export($ImageDir, 'PNG') }
        Get-ChildItem -LiteralPath $ImageDir -Filter '*.PNG' -File |
            ForEach-Object {
                if ($_.BaseName -match '(\d+)$') {
                    $pageNo = [int]$Matches[1]
                    $target = Join-Path $ImageDir ('page-{0:000}.png' -f $pageNo)
                    if ($_.FullName -ne $target) {
                        Move-Item -LiteralPath $_.FullName -Destination $target -Force
                    }
                }
            }
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ImagesExported' -Details $ImageDir
        $expectedCount = [int]$Presentation.Slides.Count
        $actualCount = @(Get-ChildItem -LiteralPath $ImageDir -Filter 'page-*.png' -File).Count
        if ($actualCount -ne $expectedCount) {
            Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ImageExportCountMismatch' -Details "Expected $expectedCount PNG files but found $actualCount."
        } else {
            Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ImageExportCountVerified' -Details "$actualCount page images exported."
        }
    } catch {
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ImagesExportFailed' -Details (Format-ComFailureDetails $_)
    }
}

function Normalize-Presentation {
    param($PowerPoint, [System.IO.FileInfo]$File)

    $safeName = Get-RelativePathSafeStem -RootPath $InputPath -TargetPath $File.FullName
    $outFile = Join-Path $OutputDir ($safeName + '.normalized' + $File.Extension)
    $pdfFile = Join-Path $OutputDir ($safeName + '.normalized.pdf')
    $backupDir = Join-Path $OutputDir '_backup_originals'

    # Skip if output already exists and is newer than source (unless -Force)
    if (-not $ReportOnly -and -not $Force -and (Test-Path -LiteralPath $outFile)) {
        $outItem = Get-Item -LiteralPath $outFile
        $pdfIsCurrent = $NoPdf -or ((Test-Path -LiteralPath $pdfFile) -and ((Get-Item -LiteralPath $pdfFile).LastWriteTimeUtc -gt $File.LastWriteTimeUtc))
        if ($outItem.LastWriteTimeUtc -gt $File.LastWriteTimeUtc -and $pdfIsCurrent) {
            Write-Verbose "Skip (up-to-date): $($File.Name)"
            Add-ReportRow -File $File.Name -FilePath $File.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'SkippedUpToDate' -Details $outFile
            return
        }
    }

    if (-not $NoBackup -and -not $ReportOnly) {
        if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
        Copy-Item -LiteralPath $File.FullName -Destination (Join-Path $backupDir $File.Name) -Force
    }

    $pres = $null
    try {
        $script:CurrentFilePath = $File.FullName
        $pres = Invoke-WithComRetry {
            $PowerPoint.Presentations.Open($File.FullName, $script:MsoFalse, $script:MsoFalse, $script:MsoFalse)
        }

        if ($UpdateMaster -and -not $ReportOnly) { Update-SlideMasterStyle -Presentation $pres -FileName $File.Name }
        $slideWidth = [double]$pres.PageSetup.SlideWidth

        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $isVideo = Test-IsVideoSlide $slide
            $slideKind = Get-SlideKind -Slide $slide -SlideNumber $i
            $preserveSlideStyle = ($slideKind -in @('Cover', 'Ending', 'Resource', 'AppendixText'))
            $isSectionTitle = (-not $preserveSlideStyle -and (Test-IsSectionTitleSlide $slide))
            Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName '(slide)' -Issue 'SlideType' -Details $(if ($isVideo) { 'VideoOrMediaCandidate' } else { 'Normal' })
            Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName '(slide)' -Issue 'SlideKind' -Details $slideKind
            $specialIssue = Get-SpecialSlidePreserveIssue -SlideKind $slideKind
            if ($null -ne $specialIssue) {
                Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName '(slide)' -Issue $specialIssue -Details (Get-SpecialSlidePreserveDetails -SlideKind $slideKind)
            }
            if ($isSectionTitle) {
                Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName '(slide)' -Issue 'SectionTitleSlide' -Details 'Single centered title slide detected; section title size preserved.'
            }
            if (Test-IsEmptySlideCandidate -Slide $slide) {
                Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName '(slide)' -Issue 'EmptySlideCandidate' -Details 'No visible text, picture, media, group, or non-placeholder shape detected.'
            }
            if (-not $ReportOnly) {
                Disable-SlideAdvanceOnClick -Slide $slide -SlideNumber $i -FileName $File.Name
            }

            foreach ($shape in $slide.Shapes) {
                try {
                    $shapeName = Get-ShapeName $shape
                    if ($shape.Type -eq $script:MsoGroup) {
                        Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName $shapeName -Issue 'GroupShapeSkipped' -Details 'Grouped shapes are not modified to avoid layout damage.'
                        continue
                    }

                    if (-not $ReportOnly) {
                        Clear-DecorativeEffects -Shape $shape
                    }

                    if ($shape.Type -eq $script:MsoTextEffect) {
                        Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName $shapeName -Issue 'WordArtStylePreserved' -Details 'WordArt object detected; decorative effects are cleared but object is not converted.'
                    }

                    if (Test-IsLargePictureShape $shape) {
                        Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName $shapeName -Issue 'RasterPicturePreserved' -Details 'Large picture detected; embedded text inside the bitmap is not rewritten automatically.'
                    }

                    if ($shape.Type -eq $script:MsoTable -or (Test-ShapeHasTable $shape)) {
                        Normalize-TableShape -Shape $shape -SlideNumber $i -FileName $File.Name
                        continue
                    }

                    $text = Get-ShapeText $shape
                    if (-not [string]::IsNullOrWhiteSpace($text)) {
                        if ($ReportOnly) {
                            if (Test-IsFormulaCandidateText $text) {
                                Add-FormulaCandidateReport -FileName $File.Name -SlideNumber $i -ShapeName $shapeName -Text $text | Out-Null
                            }
                            try {
                                $fontSize = $shape.TextFrame2.TextRange.Font.Size
                                if ($fontSize -lt $script:Style.SizeMinimum) {
                                    Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName $shapeName -Issue 'SmallText' -Details "$fontSize pt"
                                }
                            } catch { }
                        } else {
                            if (-not $preserveSlideStyle) {
                                Normalize-TextShape -Shape $shape -Text $text -SlideNumber $i -FileName $File.Name -IsVideoSlide $isVideo -IsSectionTitleSlide $isSectionTitle -SlideWidth $slideWidth
                                Normalize-HighlightBox -Shape $shape -SlideNumber $i -FileName $File.Name
                            }
                            Clear-DecorativeEffects -Shape $shape
                        }
                    }
                } catch {
                    Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName (Get-ShapeName $shape) -Issue (Get-ComFailureCategory $_) -Details (Format-ComFailureDetails $_)
                }
            }
            if (-not $ReportOnly -and -not $preserveSlideStyle) {
                try {
                    Align-AnswerTextBoxes -Slide $slide -SlideNumber $i -FileName $File.Name
                } catch {
                    Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName '(slide)' -Issue 'AnswerTextPassFailed' -Details $_.Exception.Message
                }
            }
        }

        if (-not $ReportOnly) {
            Invoke-WithComRetry { $pres.SaveAs($outFile) }
        Add-ReportRow -File $File.Name -FilePath $File.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'SavedAs' -Details $outFile
            if (-not $NoPdf) {
                Export-PresentationPdf -Presentation $pres -PdfPath $pdfFile -FileName $File.Name
            }
            if (-not [string]::IsNullOrWhiteSpace($ImageOutputDir)) {
                $imageDir = Join-Path $ImageOutputDir $safeName
                Export-PresentationImages -Presentation $pres -ImageDir $imageDir -FileName $File.Name
            }
        }
    } finally {
        $script:CurrentFilePath = ''
        if ($null -ne $pres) {
            try { $pres.Close() | Out-Null } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# --- Main ---
$InputPath  = [System.IO.Path]::GetFullPath($InputPath)
$OutputDir  = [System.IO.Path]::GetFullPath($OutputDir)

$files = @(Get-PptFiles -Path $InputPath -Pattern $FilePattern -Recurse:$Recurse)
if ($files.Count -eq 0) { throw "No .pptx/.pptm files found in $InputPath" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

# -WhatIf: list files that would be processed, then exit
if (-not $PSCmdlet.ShouldProcess($InputPath, 'Normalize PPT files')) {
    Write-Host "WhatIf: would process $($files.Count) file(s):"
    foreach ($f in $files) { Write-Host "  - $($f.FullName)" }
    return
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$failedCount = 0
$total = $files.Count
$activity = if ($ReportOnly) { 'Inspecting PPT style' } else { 'Normalizing PPT style' }

# --- Parallel path: each file in a separate PS process (safe for STA COM) ---
if ($DegreeOfParallelism -gt 1 -and $files.Count -gt 1) {
    Write-Verbose "Parallel mode: $DegreeOfParallelism worker(s) for $($files.Count) file(s)"
    $scriptPath = $PSCommandPath
    $parallelTempRoot = Join-Path $OutputDir '_parallel_workers'
    if (-not (Test-Path -LiteralPath $parallelTempRoot)) { New-Item -ItemType Directory -Path $parallelTempRoot -Force | Out-Null }
    $runningJobs = New-Object System.Collections.Generic.List[object]
    $fileQueue = New-Object System.Collections.Generic.Queue[object]
    foreach ($f in $files) { $fileQueue.Enqueue($f) }

    function Start-NextJob {
        param([System.IO.FileInfo]$FileItem)
        $workerName = '{0}_{1}' -f (Get-RelativePathSafeStem -RootPath $InputPath -TargetPath $FileItem.FullName), ([Guid]::NewGuid().ToString('N'))
        $workerOutputDir = Join-Path $parallelTempRoot $workerName
        New-Item -ItemType Directory -Path $workerOutputDir -Force | Out-Null
        $childArgs = @(
            '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', $scriptPath,
            '-InputPath', $FileItem.FullName,
            '-OutputDir', $workerOutputDir,
            '-DegreeOfParallelism', '1'
        )
        if ($ReportOnly)    { $childArgs += '-ReportOnly' }
        if ($NoBackup)      { $childArgs += '-NoBackup' }
        if ($NoPdf)         { $childArgs += '-NoPdf' }
        if ($UpdateMaster)  { $childArgs += '-UpdateMaster' }
        if ($Force)         { $childArgs += '-Force' }
        if (-not [string]::IsNullOrWhiteSpace($ImageOutputDir)) {
            $childArgs += @('-ImageOutputDir', $ImageOutputDir)
        }
        $job = Start-Job -ScriptBlock {
            & powershell.exe @args
            if ($LASTEXITCODE -ne 0) {
                throw "Child PowerShell exited with code $LASTEXITCODE."
            }
        } -ArgumentList $childArgs -Name "PPT_$($FileItem.Name)"

        return [pscustomobject]@{
            Job = $job
            File = $FileItem
            OutputDir = $workerOutputDir
        }
    }

    function Merge-WorkerOutput {
        param(
            [string]$WorkerOutputDir,
            [System.IO.FileInfo]$FileItem
        )

        if (-not (Test-Path -LiteralPath $WorkerOutputDir)) { return }

        $childReport = Join-Path $WorkerOutputDir 'physics-ppt-normalize-report.csv'
        if (Test-Path -LiteralPath $childReport) {
            try {
                foreach ($row in @(Import-Csv -LiteralPath $childReport -Encoding UTF8)) {
                    Add-ReportRow -File ([string]$row.File) -FilePath ([string]$row.FilePath) -SlideNumber ([int]$row.Slide) -ShapeName ([string]$row.Shape) -Issue ([string]$row.Issue) -Details ([string]$row.Details)
                }
            } catch {
                Add-ReportRow -File $FileItem.Name -FilePath $FileItem.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildReportMergeFailed' -Details $_.Exception.Message
            }
        }

        $workerBackupDir = Join-Path $WorkerOutputDir '_backup_originals'
        if (Test-Path -LiteralPath $workerBackupDir) {
            $backupDir = Join-Path $OutputDir '_backup_originals'
            if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
            Get-ChildItem -LiteralPath $workerBackupDir -File -ErrorAction SilentlyContinue |
                Move-Item -Destination $backupDir -Force
        }

        Get-ChildItem -LiteralPath $WorkerOutputDir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne 'physics-ppt-normalize-report.csv' } |
            Move-Item -Destination $OutputDir -Force

        Remove-Item -LiteralPath $WorkerOutputDir -Recurse -Force
    }

    # Seed initial jobs
    while ($fileQueue.Count -gt 0 -and $runningJobs.Count -lt $DegreeOfParallelism) {
        $nextFile = $fileQueue.Dequeue()
        $entry = Start-NextJob -FileItem $nextFile
        $runningJobs.Add($entry)
        Write-Verbose "Started job for: $($nextFile.Name)"
    }

    # Process completions
    $completedCount = 0
    while ($runningJobs.Count -gt 0) {
        $done = $runningJobs | Where-Object { $_.Job.State -in 'Completed', 'Failed' }
        if ($null -eq $done) {
            Start-Sleep -Milliseconds 500
            continue
        }
        foreach ($entry in @($done)) {
            $runningJobs.Remove($entry) | Out-Null
            $completedCount++
            Write-Progress -Activity $activity -Status "[$completedCount/$total] $($entry.File.Name)" -PercentComplete ([int](($completedCount / $total) * 100))
            if ($entry.Job.State -eq 'Failed' -or $entry.Job.ChildJobs[0].JobStateInfo.Reason) {
                $failedCount++
                $errMsg = try { $entry.Job.ChildJobs[0].JobStateInfo.Reason.Message } catch { 'Unknown error' }
                Write-Warning "Failed: $($entry.File.Name) — $errMsg"
                Add-ReportRow -File $entry.File.Name -FilePath $entry.File.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildProcessFailed' -Details $errMsg
            } else {
                Write-Verbose "Completed: $($entry.File.Name)"
            }
            Merge-WorkerOutput -WorkerOutputDir $entry.OutputDir -FileItem $entry.File
            Remove-Job -Job $entry.Job -Force
            # Start next queued file
            if ($fileQueue.Count -gt 0) {
                $nextFile = $fileQueue.Dequeue()
                $nextEntry = Start-NextJob -FileItem $nextFile
                $runningJobs.Add($nextEntry)
                Write-Verbose "Started job for: $($nextFile.Name)"
            }
        }
    }
    Write-Progress -Activity $activity -Completed
    if ((Test-Path -LiteralPath $parallelTempRoot) -and @(Get-ChildItem -LiteralPath $parallelTempRoot -Force).Count -eq 0) {
        Remove-Item -LiteralPath $parallelTempRoot -Force
    }
} else {
    # --- Sequential path (default): single COM instance ---
    $pp = $null
    $current = 0
    try {
        $pp = New-Object -ComObject PowerPoint.Application
        $pp.Visible = $script:MsoTrue
        foreach ($file in $files) {
            $current++
            $pct = [int](($current / $total) * 100)
            Write-Progress -Activity $activity -Status "[$current/$total] $($file.Name)" -PercentComplete $pct
            Write-Verbose "[$current/$total] Processing: $($file.FullName)"
            $attempt = 0
            $completed = $false
            while (-not $completed -and $attempt -le $FileRetryCount) {
                $attempt++
                $rowSnapshot = $script:ReportRows.Count
                try {
                    Normalize-Presentation -PowerPoint $pp -File $file
                    if ($attempt -gt 1) {
                        Add-ReportRow -File $file.Name -FilePath $file.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'RetrySucceeded' -Details "Succeeded on attempt $attempt."
                    }
                    $completed = $true
                } catch {
                    Reset-ReportRowsToCount -Count $rowSnapshot
                    $details = Format-ComFailureDetails $_
                    $category = Get-ComFailureCategory $_
                    $message = if ($null -ne $_.Exception) { $_.Exception.Message } else { '' }
                    $retryable = Test-IsRetryablePresentationFailure -Category $category -Message $message
                    if ($retryable -and $attempt -le $FileRetryCount) {
                        Write-Warning "Retrying $($file.Name) after $details"
                        Add-ReportRow -File $file.Name -FilePath $file.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'RetryAfterFailure' -Details "Attempt $attempt failed: $details"
                        $pp = Restart-PowerPointApplication -Current $pp
                        continue
                    }

                    $failedCount++
                    Write-Warning "Failed: $($file.Name) — $details"
                    Add-ReportRow -File $file.Name -FilePath $file.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue $category -Details $details
                    $completed = $true
                }
            }
        }
        Write-Progress -Activity $activity -Completed
    } catch {
        $remainingCount = if ($current -gt 0) { $total - $current + 1 } else { $total }
        $failedCount = [Math]::Min($total, $failedCount + $remainingCount)
        $details = Format-ComFailureDetails $_
        Write-Warning "PowerPoint processing failed — $details"
        foreach ($file in @($files | Select-Object -Skip ([Math]::Max(0, $current - 1)))) {
                        Add-ReportRow -File $file.Name -FilePath $file.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue (Get-ComFailureCategory $_) -Details $details
        }
        Write-Progress -Activity $activity -Completed
    } finally {
        if ($null -ne $pp) {
            try { $pp.Quit() | Out-Null } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
$sw.Stop()

$reportPath = Join-Path $OutputDir 'physics-ppt-normalize-report.csv'
# Write CSV with BOM for Excel Chinese compatibility (reliable across PS 5.x and 7.x)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$csvLines = $script:ReportRows | ConvertTo-Csv -NoTypeInformation
[System.IO.File]::WriteAllLines($reportPath, $csvLines, $utf8Bom)

Write-Host "Report saved: $reportPath"
$successCount = $total - $failedCount
Write-Host "Done. $successCount/$total file(s) succeeded in $($sw.Elapsed.ToString('mm\:ss'))"
if ($failedCount -gt 0) { Write-Warning "$failedCount file(s) failed — see report for details." }
