<#
.SYNOPSIS
  Basic toolkit self-check. This does not require PowerPoint.

.DESCRIPTION
  Validates file existence, JSON config schema, PowerShell syntax,
  VBA Option Explicit presence, and color format correctness.

.EXAMPLE
  .\Test-ToolkitFiles.ps1

.NOTES
  Run from any directory. It resolves paths relative to the script location.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot

if (-not ($PSVersionTable.PSEdition -eq 'Core' -and $PSVersionTable.PSVersion.Major -ge 7)) {
    Write-Warning 'Toolkit self-check is running under Windows PowerShell 5.1; use pwsh for the primary workflow. The legacy host remains supported as a fallback.'
}

# --- 1. Required files existence check ---
$required = @(
    'AGENTS.md',
    '.gitattributes',
    'README.md',
    'package.json',
    'docs\初中物理PPT统一排版规范.md',
    'docs\使用方法.md',
    'docs\GitHub远端迁移说明.md',
    'docs\PPT母版制作说明.md',
    'docs\自动化边界与风险控制.md',
    'docs\公式处理说明.md',
    'docs\公式排版优化路线图.md',
    'docs\编码与兼容性规范.md',
    'docs\媒体优化路线图.md',
    'docs\产品需求与工程路线图.md',
    'manual\physics-ppt-visual-review\SKILL.md',
    'manual\physics-ppt-visual-review\references\review-result.schema.json',
    'config\physics-ppt-style.config.json',
    'config\presentation-snapshot.schema.json',
    'config\invariant-comparison.schema.json',
    'tools\Normalize-PhysicsPpt.ps1',
    'tools\Apply-FormulaSvgWhitelist.ps1',
    'tools\Export-FormulaOmmlCandidates.ps1',
    'tools\Apply-FormulaOmmlWhitelist.ps1',
    'tools\Export-FormulaWhitelistSuggestions.ps1',
    'tools\Export-FormulaImageCandidates.ps1',
    'tools\Invoke-FormulaImageOcrProbe.ps1',
    'tools\formula_image_ocr_probe.py',
    'tools\PhysicsPpt.Common.ps1',
    'tools\Export-PptxVisualAudit.ps1',
    'tools\Export-PptxVisualConfirmation.ps1',
    'tools\Apply-PptxVisualAuditFixes.ps1',
    'tools\FormulaOfficeMathValidator\FormulaOfficeMathValidator.csproj',
    'tools\FormulaOfficeMathValidator\Program.cs',
    'tools\Report-PhysicsPptStyle.ps1',
    'tools\Test-ToolkitFiles.ps1',
    'tools\Assert-Toolchain.ps1',
    'tools\Invoke-PhysicsPptWorkflow.ps1',
    'tools\Optimize-PptxMedia.ps1',
    'tools\Optimize-PptxMedia.worker.js',
    'tools\Export-PptxImageCandidates.ps1',
    'tools\Invoke-PptxImageEnhancementProbe.ps1',
    'tools\Apply-PptxImageEnhancement.ps1',
    'tools\Render-FormulaSvg.mjs',
    'tools\Export-PptxExternalLinks.ps1',
    'tools\Export-PptxInvariantSnapshot.ps1',
    'tools\Compare-PptxInvariantSnapshot.ps1',
    'tools\Export-PptxAiReviewPacket.ps1',
    'tools\Import-PptxAiReviewResult.ps1',
    'tools\Build-PptxAiReviewResult.ps1',
    'tools\Test-PhysicsPptPolicy.ps1',
    'vba\PhysicsPptCommon.bas',
    'vba\PhysicsPptNormalize.bas',
    'vba\PhysicsPptReportOnly.bas',
    'vba\ApplyPhysicsPptMasterStyle.bas',
    'examples\fixtures\minimal-physics-sample.pptx',
    'examples\sample-run-commands.ps1',
    '一键规范化并导出PDF.cmd',
    '一键规范化导出并转换可编辑公式.cmd',
    '一键检查PPT.cmd'
)

foreach ($rel in $required) {
    $path = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing required file: $rel" }
}

# --- 2. JSON config schema validation ---
$configPath = Join-Path $root 'config\physics-ppt-style.config.json'
$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ($config.schemaVersion -ne 1) { throw "Unsupported config schemaVersion: $($config.schemaVersion)" }

$requiredSections = @('fonts', 'fontSizes', 'colors', 'styleRules', 'formulaWhitelist', 'rules')
foreach ($section in $requiredSections) {
    if ($null -eq $config.$section) { throw "Config missing required section: $section" }
}

$requiredFontFields = @('chinese', 'compactChinese', 'latin', 'math')
foreach ($field in $requiredFontFields) {
    if ([string]::IsNullOrWhiteSpace($config.fonts.$field)) { throw "Config fonts.$field is missing or empty" }
}

$requiredSizeFields = @('title1', 'sectionTitle', 'title2', 'body', 'auxiliary', 'minimum', 'tableHeader', 'tableBody', 'formulaInline', 'formulaStandalone', 'formulaCore', 'footer')
foreach ($field in $requiredSizeFields) {
    if ($null -eq $config.fontSizes.$field -or $config.fontSizes.$field -isnot [ValueType] -or [double]$config.fontSizes.$field -ne [math]::Floor([double]$config.fontSizes.$field) -or $config.fontSizes.$field -lt 8 -or $config.fontSizes.$field -gt 96) { throw "Config fontSizes.$field must be an integer between 8 and 96" }
}
if ($config.fontSizes.minimum -gt $config.fontSizes.body) { throw 'Config fontSizes.minimum must not exceed body.' }

# Validate color values are hex strings
$requiredColorFields = @('white', 'black', 'body', 'darkGray', 'emphasisRed', 'sectionTitle', 'extensionTitle', 'formulaBlue', 'experimentGreen', 'yellowFill', 'yellowBorder', 'blueFill', 'grayFill', 'videoYellow', 'videoBlue', 'videoRed', 'videoGreen')
foreach ($field in $requiredColorFields) {
    $val = $config.colors.$field
    if ([string]::IsNullOrWhiteSpace($val)) { throw "Config colors.$field is missing or empty" }
    if ($val -notmatch '^#[0-9A-Fa-f]{6}$') { throw "Config colors.$field is not a valid hex color: $val" }
}

function Get-RelativeLuminance {
    param([Parameter(Mandatory = $true)][string]$Hex)
    $value = $Hex.TrimStart('#')
    $channels = @(
        [Convert]::ToInt32($value.Substring(0, 2), 16),
        [Convert]::ToInt32($value.Substring(2, 2), 16),
        [Convert]::ToInt32($value.Substring(4, 2), 16)
    )
    $linear = New-Object double[] 3
    for ($i = 0; $i -lt 3; $i++) {
        $channel = $channels[$i] / 255.0
        $linear[$i] = if ($channel -le 0.04045) { $channel / 12.92 } else { [Math]::Pow(($channel + 0.055) / 1.055, 2.4) }
    }
    return (0.2126 * $linear[0]) + (0.7152 * $linear[1]) + (0.0722 * $linear[2])
}

function Get-ContrastRatio {
    param([string]$Foreground, [string]$Background)
    $foregroundLuminance = Get-RelativeLuminance $Foreground
    $backgroundLuminance = Get-RelativeLuminance $Background
    $lighter = [Math]::Max($foregroundLuminance, $backgroundLuminance)
    $darker = [Math]::Min($foregroundLuminance, $backgroundLuminance)
    return (($lighter + 0.05) / ($darker + 0.05))
}

$contrastChecks = @(
    @{ Name = 'body on white'; Foreground = $config.colors.body; Background = $config.colors.white; Minimum = 7.0 },
    @{ Name = 'darkGray on white'; Foreground = $config.colors.darkGray; Background = $config.colors.white; Minimum = 7.0 },
    @{ Name = 'emphasisRed on white'; Foreground = $config.colors.emphasisRed; Background = $config.colors.white; Minimum = 4.5 },
    @{ Name = 'sectionTitle on white'; Foreground = $config.colors.sectionTitle; Background = $config.colors.white; Minimum = 7.0 },
    @{ Name = 'extensionTitle on white'; Foreground = $config.colors.extensionTitle; Background = $config.colors.white; Minimum = 7.0 },
    @{ Name = 'formulaBlue on white'; Foreground = $config.colors.formulaBlue; Background = $config.colors.white; Minimum = 7.0 },
    @{ Name = 'experimentGreen on white'; Foreground = $config.colors.experimentGreen; Background = $config.colors.white; Minimum = 6.0 },
    @{ Name = 'formulaBlue on blueFill'; Foreground = $config.colors.formulaBlue; Background = $config.colors.blueFill; Minimum = 4.5 },
    @{ Name = 'black on yellowFill'; Foreground = $config.colors.black; Background = $config.colors.yellowFill; Minimum = 4.5 },
    @{ Name = 'emphasisRed on yellowFill'; Foreground = $config.colors.emphasisRed; Background = $config.colors.yellowFill; Minimum = 4.5 },
    @{ Name = 'videoYellow on black'; Foreground = $config.colors.videoYellow; Background = $config.colors.black; Minimum = 4.5 },
    @{ Name = 'videoBlue on black'; Foreground = $config.colors.videoBlue; Background = $config.colors.black; Minimum = 4.5 },
    @{ Name = 'videoRed on black'; Foreground = $config.colors.videoRed; Background = $config.colors.black; Minimum = 4.5 },
    @{ Name = 'videoGreen on black'; Foreground = $config.colors.videoGreen; Background = $config.colors.black; Minimum = 4.5 }
)
foreach ($contrastCheck in $contrastChecks) {
    $ratio = Get-ContrastRatio -Foreground $contrastCheck.Foreground -Background $contrastCheck.Background
    if ($ratio -lt [double]$contrastCheck.Minimum) {
        throw "Color contrast is below the classroom threshold for $($contrastCheck.Name): $([Math]::Round($ratio, 2)):1 < $($contrastCheck.Minimum):1"
    }
}

if ($null -eq $config.rules.formulaTextStyleDefault) {
    throw "Config rules.formulaTextStyleDefault is missing"
}
foreach ($safetyRule in @('doNotModifyTextContent', 'doNotMoveShapes', 'doNotResizeShapes', 'doNotModifyAnimations', 'doNotModifySlideTransitions', 'doNotCropImages', 'allowTextBoxWidthExpansion', 'disableAdvanceOnClick')) {
    $safetyProp = $config.rules.PSObject.Properties[$safetyRule]
    if ($null -eq $safetyProp -or $safetyProp.Value -isnot [bool]) {
        throw "Config rules.$safetyRule must be a boolean"
    }
}
if ($config.rules.disableAdvanceOnClick) { throw 'Config must preserve advance-on-click by default; use the explicit workflow switch for anti-misclick mode.' }
if ($config.rules.allowTextBoxWidthExpansion) { throw 'Config must disable text-box width expansion by default.' }
foreach ($requiredTrueRule in @('doNotModifyTextContent', 'doNotMoveShapes', 'doNotResizeShapes', 'doNotModifyAnimations', 'doNotModifySlideTransitions', 'doNotCropImages')) {
    if (-not $config.rules.$requiredTrueRule) { throw "Config safety rule must default to true: $requiredTrueRule" }
}

$styleRules = @($config.styleRules)
if ($styleRules.Count -lt 3) { throw 'Config styleRules must define the supported low-risk rules.' }
$ruleIds = @{}
foreach ($styleRule in $styleRules) {
    foreach ($field in @('id', 'scope', 'riskLevel')) {
        if ([string]::IsNullOrWhiteSpace([string]$styleRule.$field)) { throw "Config styleRules entry missing $field" }
    }
    if ($styleRule.id -notmatch '^[A-Z]+(\.[A-Z_]+)+$') { throw "Config styleRules id is invalid: $($styleRule.id)" }
    if ($ruleIds.ContainsKey([string]$styleRule.id)) { throw "Config styleRules contains duplicate id: $($styleRule.id)" }
    $ruleIds[[string]$styleRule.id] = $true
    if ($styleRule.enabled -isnot [bool]) { throw "Config styleRules.$($styleRule.id).enabled must be boolean" }
    if ($styleRule.riskLevel -notin @('R1', 'R2')) { throw "Config styleRules.$($styleRule.id).riskLevel must be R1 or R2" }
    if (@($styleRule.allowedProperties).Count -eq 0) { throw "Config styleRules.$($styleRule.id).allowedProperties must not be empty" }
}
foreach ($requiredStyleRule in @('STYLE.TEXT.FONT', 'STYLE.FORMULA.TEXT', 'STYLE.HIGHLIGHT.TEXT_COLOR', 'STYLE.SECTION_TITLE.EMPHASIS', 'STYLE.DECORATIVE.EFFECTS', 'SLIDE.BACKGROUND', 'SLIDE.TRANSITION.ADVANCE_ON_CLICK')) {
    if (-not $ruleIds.ContainsKey($requiredStyleRule)) { throw "Config styleRules is missing required rule: $requiredStyleRule" }
}
if (($config.styleRules | Where-Object { $_.id -eq 'SLIDE.BACKGROUND' }).enabled) {
    throw 'Slide background normalization must remain disabled by default until a real visual baseline authorizes it.'
}

$formulaWhitelist = @($config.formulaWhitelist)
if ($formulaWhitelist.Count -eq 0) {
    throw "Config formulaWhitelist must contain at least one formula rule"
}

foreach ($rule in $formulaWhitelist) {
    foreach ($field in @('name', 'sourcePattern', 'targetUnicodeMath', 'targetTex')) {
        if ([string]::IsNullOrWhiteSpace([string]$rule.$field)) {
            throw "Config formulaWhitelist entry missing field: $field"
        }
    }
    try {
        [regex]$rule.sourcePattern | Out-Null
    } catch {
        throw "Config formulaWhitelist sourcePattern is invalid: $($rule.sourcePattern)"
    }
}

$packagePath = Join-Path $root 'package.json'
$packageJson = Get-Content -LiteralPath $packagePath -Raw -Encoding UTF8 | ConvertFrom-Json
$mathJaxDependency = $packageJson.dependencies.PSObject.Properties['@mathjax/src']
if ($null -eq $mathJaxDependency -or [string]::IsNullOrWhiteSpace([string]$mathJaxDependency.Value)) {
    throw "package.json dependencies.@mathjax/src is missing"
}

$visualSkillPath = Join-Path $root 'manual\physics-ppt-visual-review\SKILL.md'
$visualSkill = Get-Content -LiteralPath $visualSkillPath -Raw -Encoding UTF8
if ($visualSkill -notmatch '(?s)^---\s*\r?\nname:\s*physics-ppt-visual-review\s*\r?\ndescription:\s*.+?\r?\n---') {
    throw 'physics-ppt-visual-review SKILL.md frontmatter is invalid.'
}
if ($visualSkill -match 'TODO|PLACEHOLDER') { throw 'physics-ppt-visual-review contains unfinished scaffold text.' }
$visualReviewSchemaPath = Join-Path $root 'manual\physics-ppt-visual-review\references\review-result.schema.json'
$visualReviewSchema = Get-Content -LiteralPath $visualReviewSchemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($visualReviewSchema.title -ne 'Physics PPT visual review result') { throw 'Visual review result schema is invalid.' }
foreach ($schemaFile in @('config\presentation-snapshot.schema.json', 'config\invariant-comparison.schema.json')) {
    $schema = Get-Content -LiteralPath (Join-Path $root $schemaFile) -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int]$schema.properties.schemaVersion.const -ne 1) { throw "Schema version guard is invalid: $schemaFile" }
}

# --- 3. PowerShell syntax check ---
$psFiles = @(
    Get-ChildItem -LiteralPath (Join-Path $root 'tools') -Filter '*.ps1' -Recurse -File
    Get-ChildItem -LiteralPath (Join-Path $root 'examples') -Filter '*.ps1' -Recurse -File
)
foreach ($file in $psFiles) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count -gt 0) {
        $msg = ($errors | ForEach-Object { $_.Message }) -join '; '
        throw "PowerShell parse errors in $($file.Name): $msg"
    }

    $content = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    if ($content -match '\[[^\]]+\]\s*\$[A-Za-z_][A-Za-z0-9_]*\s*=\s*\(Join-Path\s+\$PSScriptRoot') {
        throw "PowerShell parameter defaults must not depend on PSScriptRoot; resolve after param binding: $($file.Name)"
    }
    if ($content -match '\}function\s+[A-Za-z_][A-Za-z0-9_-]*') {
        throw "PowerShell function declarations must be separated by whitespace: $($file.Name)"
    }
    if ($content -match '(?s)Export-Csv.{0,160}-Encoding\s+UTF8') {
        throw "Excel-facing CSV must use Write-Utf8BomCsv for host-independent BOM output: $($file.Name)"
    }
    if ($content -match 'ExtractToDirectory\s*\(') {
        throw "PPTX extraction must use Expand-PptxPackageSafely: $($file.Name)"
    }
    if ($file.Name -ne 'PhysicsPpt.Common.ps1' -and $content -match 'Marshal\]::ReleaseComObject') {
        throw "COM cleanup must use Release-ComObjectSafe: $($file.Name)"
    }
}

# --- 4. Encoding guard for Windows PowerShell 5.1 ---
foreach ($file in $psFiles) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
        throw "PowerShell file must be UTF-8 with BOM for Windows PowerShell 5.1 compatibility: $($file.Name)"
    }
}

# --- 5. VBA Option Explicit check ---
$vbaFiles = Get-ChildItem -LiteralPath (Join-Path $root 'vba') -Filter '*.bas'
foreach ($file in $vbaFiles) {
    $content = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    if ($content -notmatch '(?m)^\s*Option\s+Explicit\s*$') {
        throw "VBA file missing 'Option Explicit': $($file.Name)"
    }
}

# --- 6. Minimal PPTX fixture structure check ---
$samplePptx = Join-Path $root 'examples\fixtures\minimal-physics-sample.pptx'
$zip = $null
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($samplePptx)
    $entries = @($zip.Entries | ForEach-Object { $_.FullName })
    $requiredEntries = @(
        '[Content_Types].xml',
        '_rels/.rels',
        'ppt/presentation.xml',
        'ppt/slides/slide1.xml',
        'ppt/slides/slide2.xml'
    )
    foreach ($entry in $requiredEntries) {
        if ($entry -notin $entries) { throw "Sample PPTX missing required entry: $entry" }
    }
} finally {
    if ($null -ne $zip) { $zip.Dispose() }
}

# --- 7. JSON <-> VBA constants sync check ---
$vbaCommonPath = Join-Path $root 'vba\PhysicsPptCommon.bas'
$vbaContent = Get-Content -LiteralPath $vbaCommonPath -Raw -Encoding UTF8
if ($vbaContent -notmatch '(?m)^Public\s+Const\s+NORMALIZE_SLIDE_BACKGROUND\s+As\s+Boolean\s*=\s*False\s*$') {
    throw 'VBA fallback must preserve slide backgrounds by default.'
}

# Mapping: VBA constant -> JSON path -> expected value
$syncChecks = @(
    @{ VbaConst = 'FONT_CN';              JsonCat = 'fonts';     JsonKey = 'chinese';    Type = 'String' },
    @{ VbaConst = 'FONT_LATIN';           JsonCat = 'fonts';     JsonKey = 'latin';      Type = 'String' },
    @{ VbaConst = 'FONT_MATH';            JsonCat = 'fonts';     JsonKey = 'math';       Type = 'String' },
    @{ VbaConst = 'SIZE_TITLE';           JsonCat = 'fontSizes'; JsonKey = 'title1';     Type = 'Numeric' },
    @{ VbaConst = 'SIZE_BODY';            JsonCat = 'fontSizes'; JsonKey = 'body';       Type = 'Numeric' },
    @{ VbaConst = 'SIZE_TABLE_HEADER';    JsonCat = 'fontSizes'; JsonKey = 'tableHeader'; Type = 'Numeric' },
    @{ VbaConst = 'SIZE_TABLE_BODY';      JsonCat = 'fontSizes'; JsonKey = 'tableBody';  Type = 'Numeric' },
    @{ VbaConst = 'SIZE_MINIMUM';         JsonCat = 'fontSizes'; JsonKey = 'minimum';    Type = 'Numeric' }
)

foreach ($check in $syncChecks) {
    $jsonVal = $config.($check.JsonCat).($check.JsonKey)

    # Extract VBA constant value
    $pattern = '(?m)Public\s+Const\s+' + [regex]::Escape($check.VbaConst) + '\s+As\s+\w+\s*=\s*(.+?)\s*$'
    if ($vbaContent -match $pattern) {
        $vbaRaw = $Matches[1].Trim()
        if ($check.Type -eq 'String') {
            # VBA strings are quoted: "微软雅黑"
            $vbaVal = $vbaRaw.Trim('"')
        } else {
            # VBA numeric: 46, 32 etc. (may have type suffix like 46@)
            if ($vbaRaw -match '^(\d+)') { $vbaVal = $Matches[1] } else { $vbaVal = $vbaRaw }
        }

        if ("$vbaVal" -ne "$jsonVal") {
            throw "JSON <-> VBA mismatch: $($check.VbaConst) = $vbaRaw in VBA but config $($check.JsonCat).$($check.JsonKey) = $jsonVal"
        }
    } else {
        Write-Warning "Sync check: could not find VBA constant $($check.VbaConst) in PhysicsPptCommon.bas"
    }
}

# --- 8. Shared helper behavior check (Windows PowerShell 5.1 compatible) ---
. (Join-Path $root 'tools\PhysicsPpt.Common.ps1')
$powerShellHostInfo = Get-PowerShellHostInfo
if ($null -eq $powerShellHostInfo -or [string]::IsNullOrWhiteSpace([string]$powerShellHostInfo.Path)) {
    throw 'No PowerShell host is available for child workflow processes.'
}
if (-not $powerShellHostInfo.IsPrimary) {
    Write-Warning 'PowerShell host resolver selected the Windows PowerShell 5.1 compatibility fallback; install/use pwsh for the primary path.'
}
if ([string]::IsNullOrWhiteSpace((Resolve-PowerShellHost))) {
    throw 'Resolve-PowerShellHost returned an empty executable path.'
}
$tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$selfCheckRoot = Join-Path $tempBase ('physics-ppt-toolkit-selfcheck-' + [Guid]::NewGuid().ToString('N'))
try {
    $inputDir = Join-Path $selfCheckRoot 'input'
    $excludedDir = Join-Path $inputDir 'custom-output'
    $historicalDir = Join-Path $inputDir '_physics_ppt_output_20260101_010101'
    New-Item -ItemType Directory -Path $excludedDir -Force | Out-Null
    New-Item -ItemType Directory -Path $historicalDir -Force | Out-Null
    [System.IO.File]::WriteAllBytes((Join-Path $inputDir 'lesson.pptx'), [byte[]]@(1))
    [System.IO.File]::WriteAllBytes((Join-Path $excludedDir 'generated.pptx'), [byte[]]@(1))
    [System.IO.File]::WriteAllBytes((Join-Path $historicalDir 'historical.pptx'), [byte[]]@(1))
    [System.IO.File]::WriteAllBytes((Join-Path $inputDir '~$lesson.pptx'), [byte[]]@(1))

    $discovered = @(Get-PresentationFiles -Path $inputDir -Recurse -SupportedExtensions @('.pptx') -ExcludedRoots @($excludedDir))
    if ($discovered.Count -ne 1 -or $discovered[0].Name -ne 'lesson.pptx') {
        throw 'Get-PresentationFiles failed to exclude temporary or generated presentations.'
    }

    $csvPath = Join-Path $selfCheckRoot 'utf8-bom.csv'
    $csvRows = New-Object System.Collections.Generic.List[object]
    $csvRows.Add([pscustomobject]@{ Text = '中文' }) | Out-Null
    Write-Utf8BomCsv -InputObject $csvRows -Path $csvPath
    $csvBytes = [System.IO.File]::ReadAllBytes($csvPath)
    if ($csvBytes.Length -lt 3 -or $csvBytes[0] -ne 0xEF -or $csvBytes[1] -ne 0xBB -or $csvBytes[2] -ne 0xBF) {
        throw 'Write-Utf8BomCsv did not write a UTF-8 BOM.'
    }
    if (@(Import-Csv -LiteralPath $csvPath -Encoding UTF8).Count -ne 1) {
        throw 'Write-Utf8BomCsv did not preserve the input row.'
    }

    $extractDir = Join-Path $selfCheckRoot 'package'
    $roundTripPptx = Join-Path $selfCheckRoot 'roundtrip.pptx'
    Expand-PptxPackageSafely -PptxPath $samplePptx -DestinationDir $extractDir
    New-PptxPackageFromDirectory -SourceDir $extractDir -DestinationPath $roundTripPptx
    $roundTripZip = $null
    try {
        $roundTripZip = [System.IO.Compression.ZipFile]::OpenRead($roundTripPptx)
        $roundTripEntries = @($roundTripZip.Entries | ForEach-Object { $_.FullName })
        foreach ($entry in $requiredEntries) {
            if ($entry -notin $roundTripEntries) { throw "Round-trip PPTX missing required entry: $entry" }
        }
    } finally {
        if ($null -ne $roundTripZip) { $roundTripZip.Dispose() }
    }
} finally {
    $resolvedSelfCheckRoot = [System.IO.Path]::GetFullPath($selfCheckRoot)
    if ((Test-Path -LiteralPath $resolvedSelfCheckRoot) -and
        (Test-PathInsideDirectory -ChildPath $resolvedSelfCheckRoot -ParentPath $tempBase) -and
        ([System.IO.Path]::GetFileName($resolvedSelfCheckRoot) -like 'physics-ppt-toolkit-selfcheck-*')) {
        Remove-Item -LiteralPath $resolvedSelfCheckRoot -Recurse -Force
    }
}

$normalizeContent = Get-Content -LiteralPath (Join-Path $root 'tools\Normalize-PhysicsPpt.ps1') -Raw -Encoding UTF8
$commonContent = Get-Content -LiteralPath (Join-Path $root 'tools\PhysicsPpt.Common.ps1') -Raw -Encoding UTF8
foreach ($powerShellHostMarker in @('Get-PowerShellHostInfo', 'Resolve-PowerShellHost', 'pwsh.exe', 'powershell.exe')) {
    if ($commonContent -notmatch [regex]::Escape($powerShellHostMarker)) { throw "PowerShell host resolver marker is missing: $powerShellHostMarker" }
}
foreach ($silentAutomationMarker in @('New-PowerPointApplication', 'DisplayAlerts = 1', '$application.Visible')) {
    if ($commonContent -notmatch [regex]::Escape($silentAutomationMarker)) { throw "Silent PowerPoint automation marker is missing: $silentAutomationMarker" }
}
if ($normalizeContent -match '&\s+powershell\.exe') { throw 'Parallel PowerShell workers must resolve the PS7-first host instead of hard-coding powershell.exe.' }
if ($normalizeContent -notmatch 'Resolve-PowerShellHost') { throw 'Normalize script must use Resolve-PowerShellHost for child workers.' }

foreach ($launcher in @('一键规范化并导出PDF.cmd', '一键检查PPT.cmd', '一键规范化导出并转换可编辑公式.cmd')) {
    $launcherPath = Join-Path $root $launcher
    $launcherContent = Get-Content -LiteralPath $launcherPath -Raw -Encoding UTF8
    if ($launcherContent -notmatch '(?im)where\s+pwsh\.exe') { throw "Launcher must probe pwsh first: $launcher" }
    if ($launcherContent -notmatch '(?im)set\s+"PS_HOST=pwsh\.exe"') { throw "Launcher must default to pwsh.exe: $launcher" }
    if ($launcherContent -notmatch '(?im)powershell\.exe') { throw "Launcher must retain the explicit Windows PowerShell fallback: $launcher" }
}
foreach ($automationScript in @(
    'tools\Normalize-PhysicsPpt.ps1',
    'tools\Invoke-PhysicsPptWorkflow.ps1',
    'tools\Export-PptxInvariantSnapshot.ps1',
    'tools\Export-PptxVisualAudit.ps1',
    'tools\Apply-FormulaSvgWhitelist.ps1',
    'tools\Apply-PptxVisualAuditFixes.ps1',
    'tools\Assert-Toolchain.ps1'
)) {
    $automationContent = Get-Content -LiteralPath (Join-Path $root $automationScript) -Raw -Encoding UTF8
    if ($automationContent -match 'New-Object\s+-ComObject\s+PowerPoint\.Application') {
        throw "PowerPoint automation script must use New-PowerPointApplication: $automationScript"
    }
}
foreach ($requiredGuard in @('DoNotMoveShapes', 'DoNotResizeShapes', 'DoNotModifyAnimations', 'DoNotModifySlideTransitions', 'Test-ShapeUsesAutomaticSizing')) {
    if ($normalizeContent -notmatch [regex]::Escape($requiredGuard)) { throw "Normalize safety guard is missing: $requiredGuard" }
}
foreach ($requiredRuleMarker in @('Test-StyleRuleEnabled', 'STYLE.FORMULA.TEXT', 'STYLE.DECORATIVE.EFFECTS', 'SLIDE.BACKGROUND', 'DisableAdvanceOnClick', 'AdvanceOnClickPreserved')) {
    if ($normalizeContent -notmatch [regex]::Escape($requiredRuleMarker)) { throw "Normalize rule marker is missing: $requiredRuleMarker" }
}
if ($normalizeContent -match "Issue 'TextStyleSkippedAutoSize'") { throw 'AutoSize text must use font-only normalization with geometry restoration, not be skipped.' }
foreach ($fontSafetyMarker in @('AutoSizeGeometryRestored', 'Set-AutoSizeTextFontSafely', 'TextStyleSkippedGeometryRisk', 'font and geometry were rolled back before save')) {
    if ($normalizeContent -notmatch [regex]::Escape($fontSafetyMarker)) { throw "AutoSize font normalization safety marker is missing: $fontSafetyMarker" }
}
foreach ($preflightMarker in @('Get-MissingConfiguredFonts', 'Get-SlideAspectRatioCheck', 'CHECK.FONT.AVAILABILITY', 'CHECK.SLIDE.ASPECT_RATIO', 'ConfiguredFontCheckUnavailable', 'SlideAspectRatioCheckUnavailable')) {
    if ($normalizeContent -notmatch [regex]::Escape($preflightMarker)) { throw "Preflight check marker is missing: $preflightMarker" }
}

$aiImportContent = Get-Content -LiteralPath (Join-Path $root 'tools\Import-PptxAiReviewResult.ps1') -Raw -Encoding UTF8
if ($aiImportContent -match 'PowerPoint\.Application|Presentations\.Open|SaveAs|Normalize-PhysicsPpt') { throw 'AI review import must remain read-only and must not access PPTX automation.' }

$workflowContent = Get-Content -LiteralPath (Join-Path $root 'tools\Invoke-PhysicsPptWorkflow.ps1') -Raw -Encoding UTF8
foreach ($requiredWorkflowMarker in @('BlockedMissingNormalizedPptx', 'deliveryBlocked', 'deliveryStatus', 'invariantDeliveryBlocked', "'Ready'", "'BlockedAiVisualReview'")) {
    if ($workflowContent -notmatch [regex]::Escape($requiredWorkflowMarker)) { throw "Workflow delivery gate marker is missing: $requiredWorkflowMarker" }
}
foreach ($requiredAiDeliveryMarker in @('deliveryStatus', "'Ready'", '交付状态')) {
    if ($aiImportContent -notmatch [regex]::Escape($requiredAiDeliveryMarker)) { throw "AI import delivery marker is missing: $requiredAiDeliveryMarker" }
}
foreach ($summaryCheckMarker in @('configuredFontsMissingCount', 'slideAspectMismatchCount', '字体回退风险', '非 16:9')) {
    if ($workflowContent -notmatch [regex]::Escape($summaryCheckMarker)) { throw "Summary preflight marker is missing: $summaryCheckMarker" }
}

& (Join-Path $root 'tools\Test-PhysicsPptPolicy.ps1')

Write-Host 'Toolkit self-check passed.'
