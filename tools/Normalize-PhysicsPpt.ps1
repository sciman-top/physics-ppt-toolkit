<#
.SYNOPSIS
  Normalize junior-middle-school physics PPT visual style with low-risk PowerPoint COM automation.

.DESCRIPTION
  This script opens .pptx files in Microsoft PowerPoint, applies low-risk style normalization,
  and saves normalized copies to an output folder. It does not change text content,
  animations, picture crops, video links, object geometry, or slide transitions.
  Text-box expansion and answer alignment/animation behaviors are disabled by default.
  Click-to-advance is preserved by default; use -DisableAdvanceOnClick only for a
  deliberate anti-misclick presentation mode.

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

.PARAMETER DisableAdvanceOnClick
  Explicitly disable mouse-click slide advance on every slide. The default is to
  preserve the source presentation's click behavior.

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
    [switch]$DisableAdvanceOnClick,
    [switch]$Force,
    [switch]$FailOnError,
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

$script:NormalizeScriptPath = $PSCommandPath

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
$script:PpPlaceholderSubtitle = 2
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
            $category = Get-ComFailureCategory -ErrorRecord $_
            $message = if ($null -ne $_.Exception) { $_.Exception.Message } else { '' }
            if ($attempt -ge $MaxRetries -or -not (Test-IsRetryablePresentationFailure -Category $category -Message $message)) {
                throw
            }
            Start-Sleep -Milliseconds $DelayMs
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

$script:StyleRuleEnabled = @{}
if ($null -ne $script:ConfigJson -and $null -ne $script:ConfigJson.styleRules) {
    foreach ($styleRule in @($script:ConfigJson.styleRules)) {
        $ruleId = [string]$styleRule.id
        if (-not [string]::IsNullOrWhiteSpace($ruleId)) {
            $script:StyleRuleEnabled[$ruleId] = [bool]$styleRule.enabled
        }
    }
}

function Test-StyleRuleEnabled {
    param([Parameter(Mandatory = $true)][string]$RuleId)
    if ($script:StyleRuleEnabled.ContainsKey($RuleId)) {
        return [bool]$script:StyleRuleEnabled[$RuleId]
    }
    return $true
}

function Add-RuleSkippedReport {
    param(
        [string]$FileName,
        [int]$SlideNumber,
        [string]$ShapeName,
        [string]$RuleId,
        [string]$Property,
        [string]$Details,
        [string]$RiskLevel = 'R1'
    )
    if ($FileName -ne '') {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
            -Issue 'RuleDisabled' -Details $Details -RuleId $RuleId -Property $Property `
            -RiskLevel $RiskLevel -Result 'Skipped'
    }
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
    FontCompactChinese    = Get-ConfigValue 'fonts' 'compactChinese'    '微软雅黑 UI'
    FontLatin             = Get-ConfigValue 'fonts' 'latin'             'Arial'
    FontMath              = Get-ConfigValue 'fonts' 'math'              'Cambria Math'
    SizeTitle1            = Get-ConfigValue 'fontSizes' 'title1'        46
    SizeSectionTitle      = Get-ConfigValue 'fontSizes' 'sectionTitle'  56
    SizeTitle2            = Get-ConfigValue 'fontSizes' 'title2'        38
    SizeBody              = Get-ConfigValue 'fontSizes' 'body'          32
    SizeBodyMax           = Get-ConfigValue 'fontSizes' 'bodyMax'       36
    SizeDisplayTitleMin   = Get-ConfigValue 'fontSizes' 'displayTitleMin' 56
    SizeAuxiliary         = Get-ConfigValue 'fontSizes' 'auxiliary'     28
    SizeMinimum           = Get-ConfigValue 'fontSizes' 'minimum'       24
    SizeTableHeader       = Get-ConfigValue 'fontSizes' 'tableHeader'   30
    SizeTableBody         = Get-ConfigValue 'fontSizes' 'tableBody'     28
    SizeFormulaInline     = Get-ConfigValue 'fontSizes' 'formulaInline' 34
    SizeFormulaStandalone = Get-ConfigValue 'fontSizes' 'formulaStandalone' 38
    SizeFormulaCore       = Get-ConfigValue 'fontSizes' 'formulaCore'   42
    SizeFooter            = Get-ConfigValue 'fontSizes' 'footer'        20
    ColorWhite            = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'white'         '#FFFFFF')
    ColorBlack            = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'black'         '#000000')
    ColorBody             = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'body'          '#000000')
    ColorDarkGray         = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'darkGray'      '#222222')
    ColorSectionTitle     = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'sectionTitle'  '#1F4E79')
    ColorSectionTitleHex  = Get-ConfigValue 'colors' 'sectionTitle'  '#1F4E79'
    ColorExtensionTitle   = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'extensionTitle' '#9C1C1C')
    ColorExtensionTitleHex = Get-ConfigValue 'colors' 'extensionTitle' '#9C1C1C'
    ColorFormulaBlue      = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'formulaBlue'   '#004C99')
    ColorFormulaBlueHex   = Get-ConfigValue 'colors' 'formulaBlue'   '#004C99'
    ColorExperimentGreen  = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'experimentGreen' '#006B3C')
    ColorExperimentGreenHex = Get-ConfigValue 'colors' 'experimentGreen' '#006B3C'
    ColorYellowFill       = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'yellowFill'    '#FFF2CC')
    ColorYellowBorder     = Convert-HexToRgbInt (Get-ConfigValue 'colors' 'yellowBorder'  '#D6A300')
}

$script:AllowTextBoxWidthExpansion = [bool](Get-ConfigValue 'rules' 'allowTextBoxWidthExpansion' $false)
$script:FormulaTextStyleDefault = [bool](Get-ConfigValue 'rules' 'formulaTextStyleDefault' $true)
$script:FormulaConversionDefault = [bool](Get-ConfigValue 'rules' 'formulaConversionDefault' $false)
$script:DoNotMoveShapes = [bool](Get-ConfigValue 'rules' 'doNotMoveShapes' $true)
$script:DoNotResizeShapes = [bool](Get-ConfigValue 'rules' 'doNotResizeShapes' $true)
$script:DoNotModifyAnimations = [bool](Get-ConfigValue 'rules' 'doNotModifyAnimations' $true)
# Guard anchor required by Test-ToolkitFiles: this tool never touches slide
# transitions, so the flag is honored trivially; kept explicit for the contract.
$script:DoNotModifySlideTransitions = [bool](Get-ConfigValue 'rules' 'doNotModifySlideTransitions' $true)
$script:DisableAdvanceOnClick = [bool](Get-ConfigValue 'rules' 'disableAdvanceOnClick' $false) -or [bool]$DisableAdvanceOnClick
# Host font availability cannot change mid-run; the check result is cached so
# large batches do not re-enumerate every installed family per file.
$script:ConfiguredFontCheckCache = $null

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
        [string]$Details,
        [string]$RuleId = '',
        [string]$Property = '',
        [string]$Before = '',
        [string]$After = '',
        [string]$RiskLevel = '',
        [string]$Result = ''
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
        RuleId = $RuleId
        Property = $Property
        Before = $Before
        After = $After
        RiskLevel = $RiskLevel
        Result = $Result
    }) | Out-Null
}

function Get-FileSha256 {
    param([string]$Path)
    $stream = $null
    $sha256 = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        return ($sha256.ComputeHash($stream) | ForEach-Object { $_.ToString('x2') }) -join ''
    } finally {
        if ($null -ne $sha256) { $sha256.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-NormalizeSignature {
    param(
        [System.IO.FileInfo]$File,
        [string]$SafeName
    )
    $sourceHash = Get-FileSha256 -Path $File.FullName
    $configHash = if (Test-Path -LiteralPath $ConfigPath) { Get-FileSha256 -Path $ConfigPath } else { 'missing' }
    $scriptHash = Get-FileSha256 -Path $script:NormalizeScriptPath
    $payload = [ordered]@{
        sourceSha256 = $sourceHash
        configSha256 = $configHash
        scriptSha256 = $scriptHash
        safeName = $SafeName
        noPdf = [bool]$NoPdf
        reportOnly = [bool]$ReportOnly
        updateMaster = [bool]$UpdateMaster
        disableAdvanceOnClick = [bool]$DisableAdvanceOnClick
        filePattern = [string]$FilePattern
        imageOutputDir = if ([string]::IsNullOrWhiteSpace($ImageOutputDir)) { '' } else { [System.IO.Path]::GetFullPath($ImageOutputDir) }
    }
    $json = $payload | ConvertTo-Json -Compress -Depth 6
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    try { $signature = ($hasher.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join '' } finally { $hasher.Dispose() }
    return [pscustomobject]@{ Signature = $signature; Payload = $payload }
}

function Test-PageImageSet {
    param([string]$Directory, [int]$ExpectedCount)
    if ($ExpectedCount -le 0 -or [string]::IsNullOrWhiteSpace($Directory) -or -not (Test-Path -LiteralPath $Directory)) { return $false }
    $files = @(Get-ChildItem -LiteralPath $Directory -Filter '*.png' -File)
    $numbers = New-Object System.Collections.Generic.List[int]
    foreach ($file in $files) {
        if ($file.BaseName -notmatch '^page-(\d+)$') { return $false }
        if ($file.Length -le 0) { return $false }
        try {
            $imageInfo = Get-BasicImageInfo -Path $file.FullName
            if ($imageInfo.Width -le 0 -or $imageInfo.Height -le 0) { return $false }
        } catch {
            return $false
        }
        $numbers.Add([int]$Matches[1]) | Out-Null
    }
    if ($numbers.Count -ne $ExpectedCount -or @($numbers | Sort-Object -Unique).Count -ne $numbers.Count) { return $false }
    $numbers = @($numbers | Sort-Object)
    return (($numbers | ForEach-Object { [string]$_ }) -join ',') -eq ((1..$ExpectedCount) -join ',')
}

function Test-ShapeUsesAutomaticSizing {
    param($Shape)
    try {
        return ([int]$Shape.TextFrame2.AutoSize -ne 0)
    } catch {
        return $false
    }
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
        Release-ComObjectSafe -ComObject $Current
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if ($FileRetryDelayMs -gt 0) { Start-Sleep -Milliseconds $FileRetryDelayMs }

    $next = New-PowerPointApplication
    return $next
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

function Test-IsPlaceholderTitleShape {
    param($Shape)
    try {
        if ($Shape.Type -eq $script:MsoPlaceholder) {
            $phType = $Shape.PlaceholderFormat.Type
            return ($phType -eq $script:PpPlaceholderTitle -or $phType -eq $script:PpPlaceholderCenterTitle)
        }
    } catch { }
    return $false
}

function Test-IsTitleShape {
    param($Shape)
    # Primary: PowerPoint placeholder type (most reliable)
    if (Test-IsPlaceholderTitleShape $Shape) { return $true }
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
        'Cover'        { return 'Cover slide detected; fonts, sizes, emphasis, colors, and layout are preserved. Only decorative shape effects are cleared.' }
        'Ending'       { return 'Ending slide detected; fonts, sizes, emphasis, colors, and layout are preserved. Only decorative shape effects are cleared.' }
        'Resource'     { return 'Resource/download slide detected; fonts, sizes, emphasis, colors, and layout are preserved. Only decorative shape effects are cleared.' }
        'AppendixText' { return 'Appendix explanation slide detected; original text style is preserved. Only decorative shape effects are cleared.' }
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

function Get-FormulaWhitelistMatch {
    param($Profile)

    if ($null -eq $Profile -or [string]::IsNullOrWhiteSpace([string]$Profile.Normalized)) {
        return $null
    }

    foreach ($rule in @($script:FormulaWhitelist)) {
        $pattern = Get-FormulaRuleValue -Rule $rule -Name 'sourcePattern'
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        try {
            # Case-sensitive on purpose: P=W/t (power) and p=F/S (pressure) differ only by case.
            if ([string]$Profile.Normalized -cmatch $pattern) {
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

    if (-not (Test-StyleRuleEnabled -RuleId 'STYLE.FORMULA.TEXT')) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) `
            -RuleId 'STYLE.FORMULA.TEXT' -Property 'FontName/NameFarEast/Size/Color/ParagraphAlignment' `
            -Details 'Formula text style rule is disabled by configuration.'
        return
    }

    try {
        $textRange = $Shape.TextFrame2.TextRange
        $font = $textRange.Font
        $beforeName = [string]$font.Name
        $beforeFarEast = [string]$font.NameFarEast
        $beforeSize = Get-TextRangeFontSize $textRange
        $beforeBold = [string]$font.Bold
        $beforeColor = ''
        try { $beforeColor = [string]$font.Fill.ForeColor.RGB } catch { }
        $targetSize = Resolve-FormulaTargetSize -Profile $Profile
        $safeSize = Resolve-SafeFontSize -TextRange $textRange -TargetSize $targetSize -FileName $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name
        $beforeAlignment = ''
        try { $beforeAlignment = [string]$textRange.ParagraphFormat.Alignment } catch { }

        $font.Name = $script:Style.FontMath
        $font.NameFarEast = $script:Style.FontChinese
        $font.Size = $safeSize
        $font.Bold = $script:MsoFalse
        $font.Fill.Visible = $script:MsoTrue
        $font.Fill.ForeColor.RGB = $script:Style.ColorFormulaBlue
        $textRange.ParagraphFormat.Alignment = $script:PpAlignCenter

        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'FormulaTextStyleNormalized' `
            -Details ("kind={0}; size={1:N1} pt; font={2}; color={3}; text preserved." -f $Profile.Kind, $safeSize, $script:Style.FontMath, $script:Style.ColorFormulaBlueHex) `
            -RuleId 'STYLE.FORMULA.TEXT' -Property 'FontName/NameFarEast/Size/Bold/Color/ParagraphAlignment' `
            -Before ("{0}|{1}|{2}|{3}|{4}|{5}" -f $beforeName, $beforeFarEast, $beforeSize, $beforeBold, $beforeColor, $beforeAlignment) `
            -After ("{0}|{1}|{2}|{3}|{4}|{5}" -f $script:Style.FontMath, $script:Style.FontChinese, $safeSize, $script:MsoFalse, $script:Style.ColorFormulaBlue, $script:PpAlignCenter) `
            -RiskLevel 'R1' -Result 'Applied'
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'FormulaTextStyleFailed' -Details $_.Exception.Message `
            -RuleId 'STYLE.FORMULA.TEXT' -Property 'FontName/NameFarEast/Size/Bold/Color/ParagraphAlignment' -RiskLevel 'R1' -Result 'Failed'
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
        [switch]$ForceTargetSize,
        [double]$MaxSize = 0
    )

    $currentSize = Get-TextRangeFontSize $TextRange
    $chosenSize = $TargetSize
    if (-not $ForceTargetSize) {
        if ($null -ne $currentSize -and $currentSize -lt $TargetSize) {
            if ($FileName -ne '' -and $currentSize -lt $script:Style.SizeMinimum) {
                Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
                    -Issue 'SmallTextPreserved' -Details "$currentSize pt; not increased to avoid layout overflow."
            }
            $chosenSize = $currentSize
        } elseif ($null -ne $currentSize -and $currentSize -le 72) {
            $chosenSize = $currentSize
        }
    }
    if ($MaxSize -gt 0 -and $chosenSize -gt $MaxSize) {
        if ($FileName -ne '') {
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
                -Issue 'BodyFontSizeCapped' -Details "$chosenSize pt exceeded the configured body ceiling; capped to $MaxSize pt (smaller sizes reduce overflow risk)."
        }
        return $MaxSize
    }
    return $chosenSize
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
        [switch]$ForceTargetSize,
        [double]$MaxSize = 0,
        [switch]$FontOnly
    )
    if (-not (Test-StyleRuleEnabled -RuleId 'STYLE.TEXT.FONT')) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
            -RuleId 'STYLE.TEXT.FONT' -Property $(if ($FontOnly) { 'FontName/NameFarEast' } else { 'FontName/NameFarEast/Size/Bold/Color' }) `
            -Details 'Text style rule is disabled by configuration.'
        return
    }
    try {
        $font = $TextRange.Font
        $beforeName = [string]$font.Name
        $beforeFarEast = [string]$font.NameFarEast
        $beforeSize = Get-TextRangeFontSize $TextRange
        $beforeBold = [string]$font.Bold
        $beforeColor = try { [string]$font.Fill.ForeColor.RGB } catch { '' }
        # A missing size (mixed runs report a negative COM value) means a size
        # write would flatten the emphasis hierarchy; preserve it like mixed fonts.
        $mixedFontSize = ($null -eq $beforeSize -and -not $ForceTargetSize)
        $safeSize = Resolve-SafeFontSize -TextRange $TextRange -TargetSize $Size -FileName $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -ForceTargetSize:$ForceTargetSize -MaxSize $MaxSize
        $font.Name = $script:Style.FontLatin
        $font.NameFarEast = $script:Style.FontChinese
        if (-not $FontOnly) {
            if ($mixedFontSize) {
                Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Issue 'TextStyleSkippedMixedFontSize' `
                    -Details 'Text range mixes font sizes; size, bold, and color were preserved to keep the emphasis hierarchy.' `
                    -RuleId 'STYLE.TEXT.FONT' -Property 'Size/Bold/Color' -Before "$beforeSize|$beforeBold|$beforeColor" -After "$beforeSize|$beforeBold|$beforeColor" `
                    -RiskLevel 'R0' -Result 'Skipped'
            } else {
                $font.Size = $safeSize
                $font.Bold = $(if ($Bold) { $script:MsoTrue } else { $script:MsoFalse })
                $font.Fill.Visible = $script:MsoTrue
                $font.Fill.ForeColor.RGB = $Color
            }
        }
        if ($FileName -ne '') {
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName -Issue 'TextStyleNormalized' `
                -Details $(if ($FontOnly) { 'Font family normalized; size, emphasis, color, text content, and geometry preserved.' } else { 'Font family and safe font size normalized; text content preserved.' }) `
                -RuleId 'STYLE.TEXT.FONT' -Property $(if ($FontOnly) { 'FontName/NameFarEast' } else { 'FontName/NameFarEast/Size/Bold/Color' }) `
                -Before ("{0}|{1}|{2}|{3}|{4}" -f $beforeName, $beforeFarEast, $beforeSize, $beforeBold, $beforeColor) `
                -After $(if ($FontOnly -or $mixedFontSize) { "{0}|{1}|{2}|{3}|{4}" -f $script:Style.FontLatin, $script:Style.FontChinese, $beforeSize, $beforeBold, $beforeColor } else { "{0}|{1}|{2}|{3}|{4}" -f $script:Style.FontLatin, $script:Style.FontChinese, $safeSize, $Bold, $Color }) `
                -RiskLevel 'R1' -Result 'Applied'
        }
    } catch {
        if ($FileName -ne '') {
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
                -Issue 'TextStyleFailed' -Details $_.Exception.Message -RuleId 'STYLE.TEXT.FONT' -RiskLevel 'R1' -Result 'Failed'
        }
    }
}

function Test-IsSecondaryTitleShape {
    param($Shape)
    try {
        if ($Shape.Type -eq $script:MsoPlaceholder) {
            return ([int]$Shape.PlaceholderFormat.Type -eq $script:PpPlaceholderSubtitle)
        }
    } catch { }
    $name = Get-ShapeName $Shape
    return ($name -match '(?i)(subtitle|副标题|二级标题|标题\s*2)')
}

function Test-IsAuxiliaryTextShape {
    param($Shape)
    $name = Get-ShapeName $Shape
    return ($name -match '(?i)(footer|source|页脚|来源|参考|备注|注释|辅助|说明|提示)')
}

function Get-InstalledFontFamilyNames {
    try {
        Add-Type -AssemblyName System.Drawing
        $collection = New-Object System.Drawing.Text.InstalledFontCollection
        try {
            return @($collection.Families | ForEach-Object { [string]$_.Name })
        } finally {
            $collection.Dispose()
        }
    } catch {
        throw "Installed font enumeration is unavailable: $($_.Exception.Message)"
    }
}

function Get-MissingConfiguredFonts {
    # Cached: installed families cannot change mid-run, and re-enumerating
    # them for every file in a batch is pure overhead.
    if ($null -eq $script:ConfiguredFontCheckCache) {
        $script:ConfiguredFontCheckCache = Get-MissingConfiguredFontsUncached
    }
    return $script:ConfiguredFontCheckCache
}

function Get-MissingConfiguredFontsUncached {
    $configured = @($script:Style.FontChinese, $script:Style.FontCompactChinese, $script:Style.FontLatin, $script:Style.FontMath) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        Select-Object -Unique
    try {
        $installed = @(Get-InstalledFontFamilyNames)
        if ($installed.Count -eq 0) {
            return [pscustomobject]@{
                Status = 'Unavailable'
                Configured = @($configured)
                InstalledCount = 0
                Missing = @()
                Error = 'No installed font families were returned by the host.'
            }
        }
    } catch {
        return [pscustomobject]@{
            Status = 'Unavailable'
            Configured = @($configured)
            InstalledCount = 0
            Missing = @()
            Error = $_.Exception.Message
        }
    }
    $aliases = @{
        '微软雅黑' = @('微软雅黑', 'Microsoft YaHei')
        '微软雅黑 UI' = @('微软雅黑 UI', 'Microsoft YaHei UI')
    }
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($fontName in $configured) {
        $candidates = if ($aliases.ContainsKey([string]$fontName)) { @($aliases[[string]$fontName]) } else { @([string]$fontName) }
        $found = @($installed | Where-Object { $candidate = [string]$_; @($candidates | Where-Object { $candidate -ieq [string]$_ }).Count -gt 0 }).Count -gt 0
        if (-not $found) { $missing.Add([string]$fontName) | Out-Null }
    }
    return [pscustomobject]@{
        Status = if ($missing.Count -gt 0) { 'Missing' } else { 'Available' }
        Configured = @($configured)
        InstalledCount = $installed.Count
        Missing = @($missing.ToArray())
        Error = ''
    }
}

function Get-SlideAspectRatioCheck {
    param($Presentation)
    try {
        $width = [double]$Presentation.PageSetup.SlideWidth
        $height = [double]$Presentation.PageSetup.SlideHeight
        if ($width -le 0 -or $height -le 0) {
            return [pscustomobject]@{
                Status = 'Unavailable'
                Width = $width
                Height = $height
                Ratio = $null
                ExpectedRatio = (16.0 / 9.0)
                Is16By9 = $false
                Error = 'PowerPoint returned a non-positive slide dimension.'
            }
        }
        $ratio = $width / $height
        $expected = 16.0 / 9.0
        return [pscustomobject]@{
            Status = 'Checked'
            Width = $width
            Height = $height
            Ratio = $ratio
            ExpectedRatio = $expected
            Is16By9 = ([math]::Abs($ratio - $expected) -le 0.015)
            Error = ''
        }
    } catch {
        return [pscustomobject]@{
            Status = 'Unavailable'
            Width = $null
            Height = $null
            Ratio = $null
            ExpectedRatio = (16.0 / 9.0)
            Is16By9 = $false
            Error = $_.Exception.Message
        }
    }
}

function Add-PresentationPreflightReports {
    param(
        $Presentation,
        [string]$FileName
    )

    $fontCheck = Get-MissingConfiguredFonts
    $fontRule = 'CHECK.FONT.AVAILABILITY'
    if ($fontCheck.Status -eq 'Available') {
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ConfiguredFontsAvailable' `
            -Details ("All configured font families are available on this host ({0} configured; {1} installed families inspected)." -f $fontCheck.Configured.Count, $fontCheck.InstalledCount) `
            -RuleId $fontRule -Property 'FontFamily' -Before (($fontCheck.Configured -join '|')) -After 'AvailableOnHost' `
            -RiskLevel 'R1' -Result 'Passed'
    } elseif ($fontCheck.Status -eq 'Missing') {
        $missingText = $fontCheck.Missing -join '|'
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ConfiguredFontsMissing' `
            -Details ("Configured font families are not installed on this host: {0}. PowerPoint may substitute fonts; inspect a real-host render before delivery." -f $missingText) `
            -RuleId $fontRule -Property 'FontFamily' -Before $missingText -After 'HostFontSubstitutionPossible' `
            -RiskLevel 'R1' -Result 'NeedsReview'
    } else {
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ConfiguredFontCheckUnavailable' `
            -Details ("Configured font availability could not be inspected: {0}" -f $fontCheck.Error) `
            -RuleId $fontRule -Property 'FontFamily' -Before 'Unavailable' -After 'NoWriteBack' `
            -RiskLevel 'R1' -Result 'NeedsReview'
    }

    $aspectCheck = Get-SlideAspectRatioCheck -Presentation $Presentation
    $aspectRule = 'CHECK.SLIDE.ASPECT_RATIO'
    if ($aspectCheck.Status -eq 'Checked') {
        $dimensions = "{0:N1}x{1:N1} pt (ratio {2:N4})" -f $aspectCheck.Width, $aspectCheck.Height, $aspectCheck.Ratio
        if ($aspectCheck.Is16By9) {
            Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'SlideAspectRatio16By9' `
                -Details ("Presentation canvas is 16:9: {0}." -f $dimensions) `
                -RuleId $aspectRule -Property 'SlideWidth/SlideHeight' -Before $dimensions -After '16:9' `
                -RiskLevel 'R1' -Result 'Passed'
        } else {
            Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'SlideAspectRatioMismatch' `
                -Details ("Presentation canvas is not 16:9: {0}. No canvas resize was performed because geometry and content are protected." -f $dimensions) `
                -RuleId $aspectRule -Property 'SlideWidth/SlideHeight' -Before $dimensions -After '16:9 expected; no write-back' `
                -RiskLevel 'R1' -Result 'NeedsReview'
        }
    } else {
        Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'SlideAspectRatioCheckUnavailable' `
            -Details ("Slide aspect-ratio check could not be completed: {0}" -f $aspectCheck.Error) `
            -RuleId $aspectRule -Property 'SlideWidth/SlideHeight' -Before 'Unavailable' -After 'NoWriteBack' `
            -RiskLevel 'R1' -Result 'NeedsReview'
    }

    return [pscustomobject]@{
        Font = $fontCheck
        AspectRatio = $aspectCheck
    }
}

function Set-AutoSizeFontSizeCapSafely {
    param(
        $Shape,
        [double]$MaxSize,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$FileName,
        [int]$SlideNumber,
        [string]$ShapeName
    )
    if ($MaxSize -le 0) { return }
    if (-not (Test-StyleRuleEnabled -RuleId 'STYLE.TEXT.FONT')) { return }
    $textRange = $Shape.TextFrame2.TextRange
    $beforeSize = Get-TextRangeFontSize $textRange
    if ($null -eq $beforeSize -or $beforeSize -le $MaxSize) { return }
    try {
        $textRange.Font.Size = $MaxSize
        # Measure the AutoSize reflow first; restore the captured bounds only if
        # the size change actually moved the shape.
        $reflowDrift = Get-AutoSizeGeometryDrift -Shape $Shape -Left $Left -Top $Top -Width $Width -Height $Height
        if ($reflowDrift -gt 0.05) {
            $Shape.Left = [single]$Left
            $Shape.Top = [single]$Top
            $Shape.Width = [single]$Width
            $Shape.Height = [single]$Height
        }
        $geometryDrift = (
            [Math]::Abs(([double]$Shape.Left) - $Left) -gt 0.05 -or
            [Math]::Abs(([double]$Shape.Top) - $Top) -gt 0.05 -or
            [Math]::Abs(([double]$Shape.Width) - $Width) -gt 0.05 -or
            [Math]::Abs(([double]$Shape.Height) - $Height) -gt 0.05
        )
        if ($geometryDrift) {
            $textRange.Font.Size = $beforeSize
            $Shape.Left = [single]$Left
            $Shape.Top = [single]$Top
            $Shape.Width = [single]$Width
            $Shape.Height = [single]$Height
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
                -Issue 'BodyFontSizeCapRolledBack' -Details "AutoSize geometry could not be restored after capping to $MaxSize pt; font size rolled back to $beforeSize pt." `
                -RuleId 'STYLE.TEXT.FONT' -Property 'Size' -Before "$beforeSize" -After "$beforeSize" -RiskLevel 'R0' -Result 'Skipped'
            return
        }
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
            -Issue 'BodyFontSizeCapped' -Details "$beforeSize pt exceeded the configured body ceiling; capped to $MaxSize pt with AutoSize geometry restored to the original bounds." `
            -RuleId 'STYLE.TEXT.FONT' -Property 'Size' -Before "$beforeSize" -After "$MaxSize" -RiskLevel 'R1' -Result 'Applied'
    } catch {
        try { $textRange.Font.Size = $beforeSize } catch { }
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $ShapeName `
            -Issue 'BodyFontSizeCapFailed' -Details $_.Exception.Message `
            -RuleId 'STYLE.TEXT.FONT' -Property 'Size' -RiskLevel 'R1' -Result 'Failed'
    }
}

function Get-AutoSizeGeometryDrift {
    param($Shape, [double]$Left, [double]$Top, [double]$Width, [double]$Height)
    return [Math]::Max(
        [Math]::Max([Math]::Abs(([double]$Shape.Left) - $Left), [Math]::Abs(([double]$Shape.Top) - $Top)),
        [Math]::Max([Math]::Abs(([double]$Shape.Width) - $Width), [Math]::Abs(([double]$Shape.Height) - $Height)))
}

function Set-AutoSizeTextFontSafely {
    param(
        $Shape,
        [int]$SlideNumber,
        [string]$FileName,
        [switch]$SpecialSlide
    )

    $shapeName = Get-ShapeName $Shape
    if (-not (Test-StyleRuleEnabled -RuleId 'STYLE.TEXT.FONT')) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
            -RuleId 'STYLE.TEXT.FONT' -Property 'FontName/NameFarEast' `
            -Details 'Text style rule is disabled by configuration.'
        return
    }
    $left = [double]$Shape.Left
    $top = [double]$Shape.Top
    $width = [double]$Shape.Width
    $height = [double]$Shape.Height
    $autoSize = [int]$Shape.TextFrame2.AutoSize
    $font = $Shape.TextFrame2.TextRange.Font
    $beforeName = [string]$font.Name
    $beforeFarEast = [string]$font.NameFarEast

    if ([string]::IsNullOrWhiteSpace($beforeName) -or [string]::IsNullOrWhiteSpace($beforeFarEast)) {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
            -Issue 'TextStyleSkippedMixedFont' -Details 'Mixed or unresolved font runs were preserved because a lossless rollback cannot be guaranteed.' `
            -RuleId 'STYLE.TEXT.FONT' -Property 'FontName/NameFarEast' -Before "$beforeName|$beforeFarEast" -After "$beforeName|$beforeFarEast" `
            -RiskLevel 'R0' -Result 'Skipped'
        return
    }

    try {
        $font.Name = $script:Style.FontLatin
        $font.NameFarEast = $script:Style.FontChinese

        $drift = Get-AutoSizeGeometryDrift -Shape $Shape -Left $left -Top $top -Width $width -Height $height

        if ($drift -gt 0.05) {
            # AutoSize reflowed after the font change. Restore the original bounds first;
            # keeping the normalized font is safe when the reflow was negligible.
            $Shape.Left = [single]$left
            $Shape.Top = [single]$top
            $Shape.Width = [single]$width
            $Shape.Height = [single]$height
            $driftTolerance = [Math]::Max(2.0, 0.1 * [Math]::Min($width, $height))
            if ($drift -le $driftTolerance) {
                Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
                    -Issue 'TextStyleNormalizedGeometryRestored' `
                    -Details ("Microsoft YaHei reflowed the AutoSize shape by {0:N2} pt; the normalized font was kept and the original geometry was restored." -f $drift) `
                    -RuleId 'STYLE.TEXT.FONT' -Property 'FontName/NameFarEast' -Before "$beforeName|$beforeFarEast" `
                    -After "$($script:Style.FontLatin)|$($script:Style.FontChinese)" -RiskLevel 'R1' -Result 'Applied'
                Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
                    -Issue 'AutoSizeGeometryRestored' -Details 'Font normalization completed with AutoSize geometry restored to the original bounds.' `
                    -RuleId 'SAFETY.GEOMETRY.AUTOSIZE' -Property 'Left/Top/Width/Height/AutoSize' `
                    -Before ("{0}|{1}|{2}|{3}|{4}" -f $left, $top, $width, $height, $autoSize) `
                    -After ("{0}|{1}|{2}|{3}|{4}" -f $Shape.Left, $Shape.Top, $Shape.Width, $Shape.Height, $Shape.TextFrame2.AutoSize) `
                    -RiskLevel 'R1' -Result 'Applied'
                return
            }
            if (-not $SpecialSlide) {
                $compactChineseFont = $script:Style.FontCompactChinese
                $font.Name = $script:Style.FontLatin
                $font.NameFarEast = $compactChineseFont
                $fallbackDrift = Get-AutoSizeGeometryDrift -Shape $Shape -Left $left -Top $top -Width $width -Height $height
                if ($fallbackDrift -gt 0.05) {
                    $Shape.Left = [single]$left
                    $Shape.Top = [single]$top
                    $Shape.Width = [single]$width
                    $Shape.Height = [single]$height
                    if ($fallbackDrift -le $driftTolerance) {
                        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
                            -Issue 'TextStyleNormalizedCompactFallbackGeometryRestored' `
                            -Details ("Microsoft YaHei UI reflowed the AutoSize shape by {0:N2} pt; the compatible fallback font was kept and the original geometry was restored." -f $fallbackDrift) `
                            -RuleId 'STYLE.TEXT.FONT' -Property 'FontName/NameFarEast' -Before "$beforeName|$beforeFarEast" `
                            -After "$($script:Style.FontLatin)|$compactChineseFont" -RiskLevel 'R1' -Result 'Applied'
                        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
                            -Issue 'AutoSizeGeometryRestored' -Details 'Compatible font fallback completed with AutoSize geometry restored to the original bounds.' `
                            -RuleId 'SAFETY.GEOMETRY.AUTOSIZE' -Property 'Left/Top/Width/Height/AutoSize' `
                            -Before ("{0}|{1}|{2}|{3}|{4}" -f $left, $top, $width, $height, $autoSize) `
                            -After ("{0}|{1}|{2}|{3}|{4}" -f $Shape.Left, $Shape.Top, $Shape.Width, $Shape.Height, $Shape.TextFrame2.AutoSize) `
                            -RiskLevel 'R1' -Result 'Applied'
                        return
                    }
                } else {
                    Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
                        -Issue 'TextStyleNormalizedCompactFallback' `
                        -Details 'Microsoft YaHei changed AutoSize geometry; Microsoft YaHei UI preserved geometry and was used as the compatible fallback.' `
                        -RuleId 'STYLE.TEXT.FONT' -Property 'FontName/NameFarEast' -Before "$beforeName|$beforeFarEast" `
                        -After "$($script:Style.FontLatin)|$compactChineseFont" -RiskLevel 'R1' -Result 'Applied'
                    Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
                        -Issue 'AutoSizeGeometryRestored' -Details 'Compatible font fallback completed without changing AutoSize or object geometry.' `
                        -RuleId 'SAFETY.GEOMETRY.AUTOSIZE' -Property 'Left/Top/Width/Height/AutoSize' `
                        -Before ("{0}|{1}|{2}|{3}|{4}" -f $left, $top, $width, $height, $autoSize) `
                        -After ("{0}|{1}|{2}|{3}|{4}" -f $Shape.Left, $Shape.Top, $Shape.Width, $Shape.Height, $Shape.TextFrame2.AutoSize) `
                        -RiskLevel 'R1' -Result 'Applied'
                    return
                }
                $font.Name = $beforeName
                $font.NameFarEast = $beforeFarEast
                $Shape.Left = [single]$left
                $Shape.Top = [single]$top
                $Shape.Width = [single]$width
                $Shape.Height = [single]$height
            }
            Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
                -Issue 'TextStyleSkippedGeometryRisk' -Details 'Target font reflowed AutoSize geometry beyond the accepted tolerance; font and geometry were rolled back before save.' `
                -RuleId 'STYLE.TEXT.FONT' -Property 'FontName/NameFarEast' -Before "$beforeName|$beforeFarEast" -After "$beforeName|$beforeFarEast" `
                -RiskLevel 'R0' -Result 'Skipped'
            return
        }

        # Restore geometry only when AutoSize actually reflowed; writing bounds
        # back on an unchanged shape can only introduce single-precision rounding.
        if ($drift -gt 0.05) {
            $Shape.Left = [single]$left
            $Shape.Top = [single]$top
            $Shape.Width = [single]$width
            $Shape.Height = [single]$height
        }
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
            -Issue 'TextStyleNormalized' `
            -Details $(if ($SpecialSlide) { 'Special slide font family normalized; original size, emphasis, color, AutoSize, and geometry were preserved.' } else { 'Font family normalized; original size, emphasis, color, AutoSize, and geometry were preserved.' }) `
            -RuleId 'STYLE.TEXT.FONT' -Property 'FontName/NameFarEast' -Before "$beforeName|$beforeFarEast" `
            -After "$($script:Style.FontLatin)|$($script:Style.FontChinese)" -RiskLevel 'R1' -Result 'Applied'
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
            -Issue $(if ($SpecialSlide) { 'SpecialSlideFontNormalized' } else { 'AutoSizeGeometryRestored' }) `
            -Details 'Font-only normalization completed without changing AutoSize or object geometry.' `
            -RuleId 'SAFETY.GEOMETRY.AUTOSIZE' -Property 'Left/Top/Width/Height/AutoSize' `
            -Before ("{0}|{1}|{2}|{3}|{4}" -f $left, $top, $width, $height, $autoSize) `
            -After ("{0}|{1}|{2}|{3}|{4}" -f $Shape.Left, $Shape.Top, $Shape.Width, $Shape.Height, $Shape.TextFrame2.AutoSize) `
            -RiskLevel 'R1' -Result 'Applied'
    } catch {
        try {
            $font.Name = $beforeName
            $font.NameFarEast = $beforeFarEast
            $Shape.Left = [single]$left
            $Shape.Top = [single]$top
            $Shape.Width = [single]$width
            $Shape.Height = [single]$height
        } catch { }
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
            -Issue 'TextStyleFailed' -Details $_.Exception.Message -RuleId 'STYLE.TEXT.FONT' -RiskLevel 'R1' -Result 'Failed'
    }
}

function Set-SectionTitleTextStyle {
    param($Shape, [int]$SlideNumber, [string]$FileName)
    if (-not (Test-StyleRuleEnabled -RuleId 'STYLE.SECTION_TITLE.EMPHASIS')) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) `
            -RuleId 'STYLE.SECTION_TITLE.EMPHASIS' -Property 'Bold/Color' `
            -Details 'Section-title emphasis rule is disabled by configuration.'
        return
    }
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
        $beforeBold = [string]$font.Bold
        $beforeColor = [string]$font.Fill.ForeColor.RGB
        $font.Bold = $script:MsoTrue
        $font.Fill.Visible = $script:MsoTrue
        $font.Fill.ForeColor.RGB = $targetColor
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue $issue -Details $details `
            -RuleId 'STYLE.SECTION_TITLE.EMPHASIS' -Property 'Bold/Color' -Before ("$beforeBold|$beforeColor") -After ("$($script:MsoTrue)|$targetColor") -RiskLevel 'R1' -Result 'Applied'
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'SectionTitleStyleFailed' -Details $_.Exception.Message `
            -RuleId 'STYLE.SECTION_TITLE.EMPHASIS' -Property 'Bold/Color' -RiskLevel 'R1' -Result 'Failed'
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
        if ($script:DoNotResizeShapes) {
            return $false
        }
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
    param($Shape, [string]$Text, [int]$SlideNumber, [string]$FileName, [bool]$IsVideoSlide, [bool]$IsSectionTitleSlide, [double]$SlideWidth, [switch]$FontOnly)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    $isTitle = Test-IsTitleShape $Shape
    $isSecondaryTitle = (-not $isTitle -and (Test-IsSecondaryTitleShape $Shape))
    $isAuxiliary = (-not $isTitle -and -not $isSecondaryTitle -and (Test-IsAuxiliaryTextShape $Shape))
    $isFooter = ($isAuxiliary -and (Get-ShapeName $Shape) -match '(?i)(footer|页脚|来源|source)')
    $maxSize = 0
    if (-not $IsSectionTitleSlide -and -not $isSecondaryTitle) {
        $currentSize = Get-TextRangeFontSize $Shape.TextFrame2.TextRange
        $isDisplayTitle = (Test-IsPlaceholderTitleShape $Shape) -or `
            ($null -ne $currentSize -and $currentSize -ge $script:Style.SizeDisplayTitleMin)
        if (-not $isDisplayTitle) { $maxSize = [double]$script:Style.SizeBodyMax }
    }

    if (Test-ShapeUsesAutomaticSizing -Shape $Shape) {
        $capLeft = [double]$Shape.Left
        $capTop = [double]$Shape.Top
        $capWidth = [double]$Shape.Width
        $capHeight = [double]$Shape.Height
        Set-AutoSizeTextFontSafely -Shape $Shape -SlideNumber $SlideNumber -FileName $FileName -SpecialSlide:$FontOnly
        Set-AutoSizeFontSizeCapSafely -Shape $Shape -MaxSize $maxSize -Left $capLeft -Top $capTop -Width $capWidth -Height $capHeight `
            -FileName $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape)
        return
    }

    if ($FontOnly) {
        Set-TextRangeStyle -TextRange $Shape.TextFrame2.TextRange -Size 0 -Color 0 -Bold:$false `
            -FileName $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -FontOnly
        return
    }

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
        $size = if ($IsSectionTitleSlide) {
            $script:Style.SizeSectionTitle
        } elseif ($isTitle) {
            $script:Style.SizeTitle1
        } elseif ($isSecondaryTitle) {
            $script:Style.SizeTitle2
        } elseif ($isFooter) {
            $script:Style.SizeFooter
        } elseif ($isAuxiliary) {
            $script:Style.SizeAuxiliary
        } else {
            $script:Style.SizeBody
        }
        $bold = [bool]($isTitle -or $isSecondaryTitle)
        # White text is only safe when the matching black video-slide background
        # will actually be written; SLIDE.BACKGROUND is disabled by default, and
        # white text over an untouched light background would be unreadable.
        $videoBackgroundNormalized = $IsVideoSlide -and (Test-StyleRuleEnabled -RuleId 'SLIDE.BACKGROUND')
        $color = if ($videoBackgroundNormalized) { $script:Style.ColorWhite } elseif ($isAuxiliary) { $script:Style.ColorDarkGray } else { $script:Style.ColorBody }
        Set-TextRangeStyle -TextRange $Shape.TextFrame2.TextRange -Size $size -Color $color -Bold $bold `
            -FileName $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -MaxSize $maxSize -ForceTargetSize:$IsSectionTitleSlide
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

    if ($script:DoNotModifyAnimations) {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) -Issue 'AnswerAnimationSkipped' -Details 'Animation modification is disabled by safety policy.'
        return
    }

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

    if ($script:DoNotMoveShapes) {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'AnswerTextAlignmentSkipped' -Details 'Object movement is disabled by safety policy.'
        return
    }

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
    if (-not (Test-StyleRuleEnabled -RuleId 'STYLE.HIGHLIGHT.TEXT_COLOR')) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName (Get-ShapeName $Shape) `
            -RuleId 'STYLE.HIGHLIGHT.TEXT_COLOR' -Property 'Color' `
            -Details 'Highlight text-color rule is disabled by configuration.'
        return
    }
    try {
        if ($Shape.Fill.Visible -eq $script:MsoTrue) {
            $rgb = $Shape.Fill.ForeColor.RGB
            if (Test-IsYellowishFill $rgb) {
                if ($Shape.TextFrame2.HasText -eq $script:MsoTrue) {
                    $beforeColor = [string]$Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    $Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $script:Style.ColorBody
                    if ($FileName -ne '') {
                        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $Shape.Name -Issue 'HighlightTextColorFixed' -Details 'Yellow highlight text color was set to body color for readability.' `
                            -RuleId 'STYLE.HIGHLIGHT.TEXT_COLOR' -Property 'Color' -Before $beforeColor -After ([string]$script:Style.ColorBody) -RiskLevel 'R1' -Result 'Applied'
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
    param($Shape, [int]$SlideNumber = 0, [string]$FileName = '')
    $ruleId = 'STYLE.DECORATIVE.EFFECTS'
    $shapeName = Get-ShapeName $Shape
    if (-not (Test-StyleRuleEnabled -RuleId $ruleId)) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName $shapeName `
            -RuleId $ruleId -Property 'Shadow/Glow/SoftEdge/FontShadow/FontOutline' `
            -Details 'Decorative-effect cleanup rule is disabled by configuration.'
        return
    }

    $before = [ordered]@{}
    $after = [ordered]@{}
    $changed = $false
    try { $before.ShadowVisible = [string]$Shape.Shadow.Visible } catch { }
    try { $before.ShadowTransparency = [string]$Shape.Shadow.Transparency } catch { }
    try { $before.ShadowBlur = [string]$Shape.Shadow.Blur } catch { }
    try { $before.ShadowOffsetX = [string]$Shape.Shadow.OffsetX } catch { }
    try { $before.ShadowOffsetY = [string]$Shape.Shadow.OffsetY } catch { }
    try { $before.GlowRadius = [string]$Shape.Glow.Radius } catch { }
    try { $before.SoftEdgeRadius = [string]$Shape.SoftEdge.Radius } catch { }
    try {
        if ($Shape.Shadow.Visible -ne $script:MsoFalse) { $Shape.Shadow.Visible = $script:MsoFalse; $changed = $true }
        if ([double]$Shape.Shadow.Transparency -ne 1) { $Shape.Shadow.Transparency = 1; $changed = $true }
        if ([double]$Shape.Shadow.Blur -ne 0) { $Shape.Shadow.Blur = 0; $changed = $true }
        if ([double]$Shape.Shadow.OffsetX -ne 0) { $Shape.Shadow.OffsetX = 0; $changed = $true }
        if ([double]$Shape.Shadow.OffsetY -ne 0) { $Shape.Shadow.OffsetY = 0; $changed = $true }
    } catch { }
    try { if ([double]$Shape.Glow.Radius -ne 0) { $Shape.Glow.Radius = 0; $changed = $true } } catch { }
    try { if ([double]$Shape.SoftEdge.Radius -ne 0) { $Shape.SoftEdge.Radius = 0; $changed = $true } } catch { }
    try {
        if ($Shape.TextFrame2.HasText -eq $script:MsoTrue) {
            $font = $Shape.TextFrame2.TextRange.Font
            try { $before.FontShadowVisible = [string]$font.Shadow.Visible } catch { }
            try { $before.FontShadowTransparency = [string]$font.Shadow.Transparency } catch { }
            try { $before.FontOutlineVisible = [string]$font.Line.Visible } catch { }
            if ($font.Shadow.Visible -ne $script:MsoFalse) { $font.Shadow.Visible = $script:MsoFalse; $changed = $true }
            if ([double]$font.Shadow.Transparency -ne 1) { $font.Shadow.Transparency = 1; $changed = $true }
            if ($font.Line.Visible -ne $script:MsoFalse) { $font.Line.Visible = $script:MsoFalse; $changed = $true }
        }
    } catch { }
    try { $after.ShadowVisible = [string]$Shape.Shadow.Visible } catch { }
    try { $after.ShadowTransparency = [string]$Shape.Shadow.Transparency } catch { }
    try { $after.ShadowBlur = [string]$Shape.Shadow.Blur } catch { }
    try { $after.ShadowOffsetX = [string]$Shape.Shadow.OffsetX } catch { }
    try { $after.ShadowOffsetY = [string]$Shape.Shadow.OffsetY } catch { }
    try { $after.GlowRadius = [string]$Shape.Glow.Radius } catch { }
    try { $after.SoftEdgeRadius = [string]$Shape.SoftEdge.Radius } catch { }
    if ($FileName -ne '' -and $changed) {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName $shapeName -Issue 'DecorativeEffectsCleared' `
            -Details 'Shadow, glow, soft-edge, and text-outline effects were cleared.' `
            -RuleId $ruleId -Property 'Shadow/Glow/SoftEdge/FontShadow/FontOutline' `
            -Before (($before | ConvertTo-Json -Compress)) -After (($after | ConvertTo-Json -Compress)) `
            -RiskLevel 'R1' -Result 'Applied'
    }
}

function Set-SlideBackground {
    param($Slide, [bool]$IsVideoSlide, [int]$SlideNumber = 0, [string]$FileName = '')
    $ruleId = 'SLIDE.BACKGROUND'
    if (-not (Test-StyleRuleEnabled -RuleId $ruleId)) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' `
            -RuleId $ruleId -Property 'FollowMasterBackground/Fill.ForeColor.RGB' `
            -Details 'Slide background normalization is disabled by configuration; original background preserved.' -RiskLevel 'R2'
        return
    }
    $beforeFollow = ''
    $beforeColor = ''
    try {
        $beforeFollow = [string]$Slide.FollowMasterBackground
        $beforeColor = [string]$Slide.Background.Fill.ForeColor.RGB
        $Slide.FollowMasterBackground = $script:MsoFalse
        $Slide.Background.Fill.Solid() | Out-Null
        $Slide.Background.Fill.ForeColor.RGB = $(if ($IsVideoSlide) { $script:Style.ColorBlack } else { $script:Style.ColorWhite })
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'SlideBackgroundNormalized' `
            -Details $(if ($IsVideoSlide) { 'Video/media candidate background normalized to black.' } else { 'Normal slide background normalized to white.' }) `
            -RuleId $ruleId -Property 'FollowMasterBackground/Fill.ForeColor.RGB' `
            -Before "$beforeFollow|$beforeColor" -After ("$($script:MsoFalse)|" + [string]$(if ($IsVideoSlide) { $script:Style.ColorBlack } else { $script:Style.ColorWhite })) `
            -RiskLevel 'R2' -Result 'Applied'
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'SlideBackgroundNormalizeFailed' `
            -Details (Format-ComFailureDetails $_) -RuleId $ruleId -Property 'FollowMasterBackground/Fill.ForeColor.RGB' -RiskLevel 'R2' -Result 'Failed'
    }
}

function Report-SlideBackgroundCandidate {
    param($Slide, [bool]$IsVideoSlide, [int]$SlideNumber = 0, [string]$FileName = '')
    $beforeFollow = ''
    $beforeColor = ''
    try { $beforeFollow = [string]$Slide.FollowMasterBackground } catch { }
    try { $beforeColor = [string]$Slide.Background.Fill.ForeColor.RGB } catch { }
    $targetColor = if ($IsVideoSlide) { $script:Style.ColorBlack } else { $script:Style.ColorWhite }
    $reason = if (Test-StyleRuleEnabled -RuleId 'SLIDE.BACKGROUND') {
        'Background candidate reported only; enable SLIDE.BACKGROUND explicitly after visual review to write it.'
    } else {
        'Background normalization is disabled by configuration; original background preserved.'
    }
    Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'SlideBackgroundCandidate' `
        -Details $reason -RuleId 'SLIDE.BACKGROUND' -Property 'FollowMasterBackground/Fill.ForeColor.RGB' `
        -Before "$beforeFollow|$beforeColor" -After ("$($script:MsoFalse)|$targetColor") -RiskLevel 'R2' -Result 'Skipped'
}

function Disable-SlideAdvanceOnClick {
    param($Slide, [int]$SlideNumber, [string]$FileName)

    if (-not (Test-StyleRuleEnabled -RuleId 'SLIDE.TRANSITION.ADVANCE_ON_CLICK')) {
        Add-RuleSkippedReport -FileName $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' `
            -RuleId 'SLIDE.TRANSITION.ADVANCE_ON_CLICK' -Property 'AdvanceOnClick' `
            -Details 'Click-to-advance rule is disabled by configuration.'
        return
    }
    if (-not $script:DisableAdvanceOnClick) {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'AdvanceOnClickPreserved' -Details 'Click-to-advance behavior preserved by default; use -DisableAdvanceOnClick for anti-misclick mode.' -RuleId 'SLIDE.TRANSITION.ADVANCE_ON_CLICK' -Property 'AdvanceOnClick' -RiskLevel 'R1' -Result 'Skipped'
        return
    }
    try {
        $before = [string]$Slide.SlideShowTransition.AdvanceOnClick
        $Slide.SlideShowTransition.AdvanceOnClick = $script:MsoFalse
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'AdvanceOnClickDisabled' -Details 'Advance on click disabled; other transition properties are preserved.' -RuleId 'SLIDE.TRANSITION.ADVANCE_ON_CLICK' -Property 'AdvanceOnClick' -Before $before -After ([string]$script:MsoFalse) -RiskLevel 'R1' -Result 'Applied'
    } catch {
        Add-ReportRow -File $FileName -SlideNumber $SlideNumber -ShapeName '(slide)' -Issue 'AdvanceOnClickDisableFailed' -Details (Format-ComFailureDetails $_) -RuleId 'SLIDE.TRANSITION.ADVANCE_ON_CLICK' -Property 'AdvanceOnClick' -RiskLevel 'R1' -Result 'Failed'
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
        if (-not (Test-Path -LiteralPath $PdfPath) -or (Get-Item -LiteralPath $PdfPath).Length -le 0) {
            throw "PowerPoint did not create a non-empty PDF: $PdfPath"
        }
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
        $expectedCount = [int]$Presentation.Slides.Count
        Get-ChildItem -LiteralPath $ImageDir -Filter 'page-*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force
        $failedSlides = New-Object System.Collections.Generic.List[int]
        for ($slideNo = 1; $slideNo -le $expectedCount; $slideNo++) {
            $target = Join-Path $ImageDir ('page-{0:000}.png' -f $slideNo)
            try {
                $slide = $Presentation.Slides.Item($slideNo)
                Invoke-WithComRetry { $slide.Export($target, 'PNG') }
                if (-not (Test-Path -LiteralPath $target) -or (Get-Item -LiteralPath $target).Length -le 0) {
                    throw "PowerPoint did not create a non-empty PNG: $target"
                }
                $imageInfo = Get-BasicImageInfo -Path $target
                if ($imageInfo.Width -le 0 -or $imageInfo.Height -le 0) {
                    throw "PowerPoint created an undecodable PNG: $target"
                }
            } catch {
                $failedSlides.Add($slideNo) | Out-Null
                Add-ReportRow -File $FileName -SlideNumber $slideNo -ShapeName '(slide)' -Issue 'SlidePngExportFailed' -Details (Format-ComFailureDetails $_)
            }
        }
        $actualNumbers = @(Get-ChildItem -LiteralPath $ImageDir -Filter 'page-*.png' -File | ForEach-Object {
            if ($_.BaseName -match '^page-(\d+)$') { [int]$Matches[1] }
        } | Sort-Object -Unique)
        $actualCount = $actualNumbers.Count
        $expectedKey = if ($expectedCount -gt 0) { ((1..$expectedCount) -join ',') } else { '' }
        $actualKey = (($actualNumbers | ForEach-Object { [string]$_ }) -join ',')
        if ($failedSlides.Count -gt 0 -or $actualKey -ne $expectedKey) {
            Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ImageExportCountMismatch' -Details "Expected pages $expectedKey but found $actualKey; failed slides=$($failedSlides -join ',')."
            Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ImagesExportFailed' -Details "Page image export did not produce the complete expected set under $ImageDir."
        } else {
            Add-ReportRow -File $FileName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ImagesExported' -Details $ImageDir
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
    $cachePath = $outFile + '.cache.json'
    $backupDir = Join-Path $OutputDir '_backup_originals'
    $signature = Get-NormalizeSignature -File $File -SafeName $safeName

    # Timestamps alone cannot distinguish flags, config/script revisions, or a
    # workflow that intentionally cleared an output artifact.
    if (-not $ReportOnly -and -not $Force -and (Test-Path -LiteralPath $outFile)) {
        $cache = $null
        try { $cache = Get-Content -LiteralPath $cachePath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $cache = $null }
        $cacheImageDir = if ($null -ne $cache) { [string]$cache.imageDir } else { '' }
        $cacheExpectedSlides = 0
        if ($null -ne $cache -and $null -ne $cache.PSObject.Properties['expectedSlides']) {
            [void][int]::TryParse([string]$cache.expectedSlides, [ref]$cacheExpectedSlides)
        }
        $cacheOutputReady = (Test-Path -LiteralPath $outFile) -and (Get-Item -LiteralPath $outFile).Length -gt 0
        $cacheArtifactsReady = $NoPdf -or ((Test-Path -LiteralPath $pdfFile) -and ((Get-Item -LiteralPath $pdfFile).Length -gt 0))
        if (-not [string]::IsNullOrWhiteSpace($ImageOutputDir)) {
            $cacheArtifactsReady = $cacheArtifactsReady -and (Test-PageImageSet -Directory $cacheImageDir -ExpectedCount $cacheExpectedSlides)
        }
        if ($null -ne $cache -and $cacheOutputReady -and [string]$cache.signature -eq $signature.Signature -and
            [string]$cache.outputPath -eq $outFile -and $cacheArtifactsReady) {
            Write-Verbose "Skip (up-to-date): $($File.Name)"
            Add-ReportRow -File $File.Name -FilePath $File.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'SkippedUpToDate' -Details $outFile
            return
        }
    }

    if (-not $NoBackup -and -not $ReportOnly) {
        if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
        # The backup name mirrors the collision-safe output stem so same-named
        # files from different subdirectories (-Recurse) cannot overwrite each
        # other's only original backup.
        Copy-Item -LiteralPath $File.FullName -Destination (Join-Path $backupDir ($safeName + $File.Extension)) -Force
    }

    $pres = $null
    try {
        $script:CurrentFilePath = $File.FullName
        $pres = Invoke-WithComRetry {
            $PowerPoint.Presentations.Open($File.FullName, $(if ($ReportOnly) { $script:MsoTrue } else { $script:MsoFalse }), $script:MsoFalse, $script:MsoFalse)
        }

        Add-PresentationPreflightReports -Presentation $pres -FileName $File.Name | Out-Null
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
            if ($ReportOnly) {
                Report-SlideBackgroundCandidate -Slide $slide -IsVideoSlide $isVideo -SlideNumber $i -FileName $File.Name
            } else {
                Set-SlideBackground -Slide $slide -IsVideoSlide $isVideo -SlideNumber $i -FileName $File.Name
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
                        if ($shape.Type -eq $script:MsoTable -or (Test-ShapeHasTable $shape)) {
                            Normalize-TableShape -Shape $shape -SlideNumber $i -FileName $File.Name
                            continue
                        }
                        Clear-DecorativeEffects -Shape $shape -SlideNumber $i -FileName $File.Name
                    }

                    if ($shape.Type -eq $script:MsoTextEffect) {
                        Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName $shapeName -Issue 'WordArtStylePreserved' -Details 'WordArt object detected; decorative effects are cleared but object is not converted.'
                    }

                    if (Test-IsLargePictureShape $shape) {
                        Add-ReportRow -File $File.Name -SlideNumber $i -ShapeName $shapeName -Issue 'RasterPicturePreserved' -Details 'Large picture detected; embedded text inside the bitmap is not rewritten automatically.'
                    }

                    $text = Get-ShapeText $shape
                    if (-not [string]::IsNullOrWhiteSpace($text)) {
                        if ($ReportOnly -or $preserveSlideStyle) {
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
                            Normalize-TextShape -Shape $shape -Text $text -SlideNumber $i -FileName $File.Name -IsVideoSlide $isVideo -IsSectionTitleSlide $isSectionTitle -SlideWidth $slideWidth
                            Normalize-HighlightBox -Shape $shape -SlideNumber $i -FileName $File.Name
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
            $fileFailureIssues = @($script:ReportRows | Where-Object {
                $_.FilePath -eq $File.FullName -and $_.Issue -in @('PdfExportFailed', 'ImagesExportFailed', 'ImageExportCountMismatch', 'SlidePngExportFailed', 'SavedAsFailed')
            })
            $cacheImageDir = if ([string]::IsNullOrWhiteSpace($ImageOutputDir)) { '' } else { Join-Path $ImageOutputDir $safeName }
            $cacheReady = (Test-Path -LiteralPath $outFile) -and (Get-Item -LiteralPath $outFile).Length -gt 0 -and $fileFailureIssues.Count -eq 0 -and
                ($NoPdf -or ((Test-Path -LiteralPath $pdfFile) -and (Get-Item -LiteralPath $pdfFile).Length -gt 0)) -and
                ([string]::IsNullOrWhiteSpace($ImageOutputDir) -or (Test-PageImageSet -Directory $cacheImageDir -ExpectedCount ([int]$pres.Slides.Count)))
            if ($cacheReady) {
                $cacheRecord = [ordered]@{
                    schemaVersion = 1
                    generatedAt = (Get-Date).ToUniversalTime().ToString('o')
                    signature = $signature.Signature
                    sourceSha256 = $signature.Payload.sourceSha256
                    outputPath = $outFile
                    pdfPath = if ($NoPdf) { '' } else { $pdfFile }
                    imageDir = $cacheImageDir
                    expectedSlides = [int]$pres.Slides.Count
                    parameters = $signature.Payload
                }
                [System.IO.File]::WriteAllText($cachePath, ($cacheRecord | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
            }
        }
    } finally {
        $script:CurrentFilePath = ''
        if ($null -ne $pres) {
            try { $pres.Close() | Out-Null } catch { }
            Release-ComObjectSafe -ComObject $pres
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# --- Main ---
$InputPath  = [System.IO.Path]::GetFullPath($InputPath)
$OutputDir  = [System.IO.Path]::GetFullPath($OutputDir)

$files = @(Get-PresentationFiles -Path $InputPath -Pattern $FilePattern -Recurse:$Recurse -SupportedExtensions @('.pptx', '.pptm') -ExcludedRoots @($OutputDir, $ImageOutputDir))
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
    $workerPowerShell = Resolve-PowerShellHost
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
        if ($DisableAdvanceOnClick) { $childArgs += '-DisableAdvanceOnClick' }
        if ($Force)         { $childArgs += '-Force' }
        $childArgs += @('-FileRetryCount', [string]$FileRetryCount, '-FileRetryDelayMs', [string]$FileRetryDelayMs)
        if (-not [string]::IsNullOrWhiteSpace($ImageOutputDir)) {
            $childArgs += @('-ImageOutputDir', $ImageOutputDir)
        }
        $job = Start-Job -ScriptBlock {
            param(
                [string]$hostPath,
                [object[]]$workerArguments
            )
            & $hostPath @workerArguments
            if ($LASTEXITCODE -ne 0) {
                throw "Child PowerShell exited with code $LASTEXITCODE."
            }
        } -ArgumentList @($workerPowerShell, (,$childArgs)) -Name "PPT_$($FileItem.Name)"

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

        $mergeSucceeded = $true
        $childRows = @()
        if (-not (Test-Path -LiteralPath $WorkerOutputDir)) {
            return [pscustomobject]@{
                Succeeded = $false
                ChildFailureCount = 0
                ChildFailureIssues = @()
            }
        }

        $childReport = Join-Path $WorkerOutputDir 'physics-ppt-normalize-report.csv'
        if (Test-Path -LiteralPath $childReport) {
            try {
                $childRows = @(Import-Csv -LiteralPath $childReport -Encoding UTF8)
                foreach ($row in $childRows) {
                    $ruleId = if ($null -ne $row.PSObject.Properties['RuleId']) { [string]$row.RuleId } else { '' }
                    $property = if ($null -ne $row.PSObject.Properties['Property']) { [string]$row.Property } else { '' }
                    $before = if ($null -ne $row.PSObject.Properties['Before']) { [string]$row.Before } else { '' }
                    $after = if ($null -ne $row.PSObject.Properties['After']) { [string]$row.After } else { '' }
                    $riskLevel = if ($null -ne $row.PSObject.Properties['RiskLevel']) { [string]$row.RiskLevel } else { '' }
                    $result = if ($null -ne $row.PSObject.Properties['Result']) { [string]$row.Result } else { '' }
                    Add-ReportRow -File ([string]$row.File) -FilePath ([string]$row.FilePath) -SlideNumber ([int]$row.Slide) -ShapeName ([string]$row.Shape) -Issue ([string]$row.Issue) -Details ([string]$row.Details) `
                        -RuleId $ruleId -Property $property -Before $before -After $after -RiskLevel $riskLevel -Result $result
                }
            } catch {
                $mergeSucceeded = $false
                Add-ReportRow -File $FileItem.Name -FilePath $FileItem.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildReportMergeFailed' -Details $_.Exception.Message
            }
        } else {
            $mergeSucceeded = $false
            Add-ReportRow -File $FileItem.Name -FilePath $FileItem.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildReportMissing' -Details "Worker report was not produced: $childReport"
        }

        # The child normalizer records per-file failures in its report but does
        # not exit non-zero unless -FailOnError is supplied.  Treat explicit
        # Failed results and failure-shaped issue IDs as a failed file so the
        # parent summary cannot report a false success.
        $childFailureRows = @($childRows | Where-Object {
            ([string]$_.Result -eq 'Failed') -or
            ([string]$_.Issue -match '(?i)(Failed|Failure)$') -or
            ([string]$_.Issue -in @(
                'PowerPointBusyOrRejectedCall', 'FileInUseOrSharingViolation',
                'PowerPointComNotRegistered', 'FileNotFoundOrUnavailable',
                'PowerPointComFailure', 'UnhandledFailure'
            ))
        })

        $workerBackupDir = Join-Path $WorkerOutputDir '_backup_originals'
        if (Test-Path -LiteralPath $workerBackupDir) {
            $backupDir = Join-Path $OutputDir '_backup_originals'
            if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
            try {
                Get-ChildItem -LiteralPath $workerBackupDir -File -ErrorAction SilentlyContinue |
                    Move-Item -Destination $backupDir -Force
            } catch {
                $mergeSucceeded = $false
                Add-ReportRow -File $FileItem.Name -FilePath $FileItem.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'WorkerMergeFailed' -Details $_.Exception.Message
            }
        }

        try {
            Get-ChildItem -LiteralPath $WorkerOutputDir -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne 'physics-ppt-normalize-report.csv' } |
                Move-Item -Destination $OutputDir -Force
            Remove-Item -LiteralPath $WorkerOutputDir -Recurse -Force
        } catch {
            $mergeSucceeded = $false
            # A locked file (antivirus scan, preview window) must not abort the
            # whole batch; keep the worker dir for inspection and report it.
            Add-ReportRow -File $FileItem.Name -FilePath $FileItem.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'WorkerMergeFailed' -Details ("{0}; worker output kept at {1}" -f $_.Exception.Message, $WorkerOutputDir)
        }

        return [pscustomobject]@{
            Succeeded = $mergeSucceeded
            ChildFailureCount = $childFailureRows.Count
            ChildFailureIssues = @($childFailureRows | ForEach-Object { [string]$_.Issue } | Sort-Object -Unique)
        }
    }

    # Seed initial jobs
    while ($fileQueue.Count -gt 0 -and $runningJobs.Count -lt $DegreeOfParallelism) {
        $nextFile = $fileQueue.Dequeue()
        try {
            $entry = Start-NextJob -FileItem $nextFile
            $runningJobs.Add($entry)
            Write-Verbose "Started job for: $($nextFile.Name)"
        } catch {
            $failedCount++
            $details = $_.Exception.Message
            Write-Warning "Failed to start worker for $($nextFile.Name) — $details"
            Add-ReportRow -File $nextFile.Name -FilePath $nextFile.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildProcessStartFailed' -Details $details
        }
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
            $fileFailed = $false
            try {
                if ($entry.Job.State -eq 'Failed' -or $entry.Job.ChildJobs[0].JobStateInfo.Reason) {
                    $fileFailed = $true
                    $errMsg = try { $entry.Job.ChildJobs[0].JobStateInfo.Reason.Message } catch { 'Unknown error' }
                    Write-Warning "Failed: $($entry.File.Name) — $errMsg"
                    Add-ReportRow -File $entry.File.Name -FilePath $entry.File.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildProcessFailed' -Details $errMsg
                } else {
                    Write-Verbose "Completed: $($entry.File.Name)"
                }
                $mergeResult = Merge-WorkerOutput -WorkerOutputDir $entry.OutputDir -FileItem $entry.File
                if (-not $mergeResult.Succeeded) {
                    $fileFailed = $true
                }
                if ($mergeResult.ChildFailureCount -gt 0) {
                    $fileFailed = $true
                    Add-ReportRow -File $entry.File.Name -FilePath $entry.File.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildReportedFailure' -Details ("{0} failure row(s): {1}" -f $mergeResult.ChildFailureCount, ($mergeResult.ChildFailureIssues -join ', '))
                }
            } catch {
                # Drain-loop resilience: a single merge/start failure must not
                # abort the batch before the report is written.
                $fileFailed = $true
                Write-Warning "Worker completion failed for $($entry.File.Name) — $($_.Exception.Message)"
                Add-ReportRow -File $entry.File.Name -FilePath $entry.File.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'WorkerMergeFailed' -Details $_.Exception.Message
            } finally {
                if ($fileFailed) { $failedCount++ }
                try { Remove-Job -Job $entry.Job -Force } catch { }
            }
            # Start next queued file
            if ($fileQueue.Count -gt 0) {
                try {
                    $nextFile = $fileQueue.Dequeue()
                    $nextEntry = Start-NextJob -FileItem $nextFile
                    $runningJobs.Add($nextEntry)
                    Write-Verbose "Started job for: $($nextFile.Name)"
                } catch {
                    $failedCount++
                    Write-Warning "Failed to start next worker — $($_.Exception.Message)"
                    Add-ReportRow -File $nextFile.Name -FilePath $nextFile.FullName -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ChildProcessStartFailed' -Details $_.Exception.Message
                }
            }
        }
    }
    Write-Progress -Activity $activity -Completed
    if (Test-Path -LiteralPath $parallelTempRoot) {
        $leftover = @(Get-ChildItem -LiteralPath $parallelTempRoot -Force -ErrorAction SilentlyContinue)
        if ($leftover.Count -eq 0) {
            Remove-Item -LiteralPath $parallelTempRoot -Force
        } else {
            Add-ReportRow -File '(batch)' -SlideNumber 0 -ShapeName '(presentation)' -Issue 'ParallelTempLeftover' -Details ("{0} worker output dir(s) kept for inspection under {1}" -f $leftover.Count, $parallelTempRoot)
        }
    }
} else {
    # --- Sequential path (default): single COM instance ---
    $pp = $null
    $current = 0
    try {
        $pp = New-PowerPointApplication
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
            Release-ComObjectSafe -ComObject $pp
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
$sw.Stop()

$reportPath = Join-Path $OutputDir 'physics-ppt-normalize-report.csv'
try {
    # Shared helper keeps the BOM for Excel Chinese compatibility across PS 5.x and 7.x,
    # and guards the final write so a locked report file cannot crash the run summary.
    Write-Utf8BomCsv -InputObject $script:ReportRows.ToArray() -Path $reportPath
} catch {
    Write-Warning "Failed to write report CSV '$reportPath' — $($_.Exception.Message)"
}

Write-Host "Report saved: $reportPath"
$successCount = $total - $failedCount
Write-Host "Done. $successCount/$total file(s) succeeded in $($sw.Elapsed.ToString('mm\:ss'))"
if ($failedCount -gt 0) {
    Write-Warning "$failedCount file(s) failed — see report for details."
    if ($FailOnError) { throw "Normalization failed for $failedCount file(s)." }
}
