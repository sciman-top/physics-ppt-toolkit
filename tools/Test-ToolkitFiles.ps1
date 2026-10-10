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
    'docs\公式识别转换实施计划.md',
    'docs\公式OLE转换执行计划.md',
    'docs\公式GoldSet编制SOP.md',
    'docs\编码与兼容性规范.md',
    'docs\媒体优化路线图.md',
    'docs\产品需求与工程路线图.md',
    'manual\physics-ppt-visual-review\SKILL.md',
    'manual\physics-ppt-visual-review\references\review-result.schema.json',
    'config\physics-ppt-style.config.json',
    'tools\Normalize-PhysicsPpt.ps1',
    'tools\Export-FormulaOmmlCandidates.ps1',
    'tools\Test-FormulaIr.ps1',
    'tools\Resolve-FormulaCanonicalContext.ps1',
    'tools\Test-FormulaCanonicalContext.ps1',
    'tools\Test-FormulaOfficeMathValidator.ps1',
    'tools\Test-ClosedWorldCircuitBreaker.ps1',
    'tools\Plan-ClosedWorldUnattended.ps1',
    'tools\Test-FormulaPowerPointRepairHandling.ps1',
    'tools\Apply-FormulaOmmlWhitelist.ps1',
    'tools\Apply-FormulaOmmlForOle.ps1',
    'tools\Export-FormulaGoldSet.ps1',
    'tools\Invoke-FormulaOleBatch.ps1',
    'tools\Export-FormulaWhitelistSuggestions.ps1',
    'tools\Export-FormulaImageCandidates.ps1',
    'tools\Export-FormulaImageCrops.ps1',
    'tools\Compare-FormulaConverters.ps1',
    'tools\Export-FormulaCarrierInventory.ps1',
    'tools\Export-FormulaOleMapping.ps1',
    'tools\Export-FormulaOleCrops.ps1',
    'tools\Export-FormulaOleVisualAdjudication.ps1',
    'tools\Test-FormulaOleVisualAdjudication.ps1',
    'tools\Export-FormulaEvidenceManifest.ps1',
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
    'tools\Export-PptxInvariantSnapshot.ps1',
    'tools\Compare-PptxInvariantSnapshot.ps1',
    'tools\Export-PptxAiReviewPacket.ps1',
    'tools\Import-PptxAiReviewResult.ps1',
    'tools\Test-PhysicsPptPolicy.ps1',
    'tools\Apply-PptxBrandVisualRefresh.ps1',
    'tools\Apply-PptxHighlightBoxStyle.ps1',
    'tools\generate_brand_assets.py',
    'assets\brand\sciman-icon.png',
    'assets\brand\sciman-icon-shadow.png',
    'assets\brand\sciman-icon-watermark.png',
    'assets\brand\bg-16x9.jpg',
    'vba\PhysicsPptCommon.bas',
    'vba\PhysicsPptNormalize.bas',
    'vba\PhysicsPptReportOnly.bas',
    'vba\ApplyPhysicsPptMasterStyle.bas',
    'examples\fixtures\minimal-physics-sample.pptx',
    'examples\fixtures\formula-ir.valid.json',
    'examples\fixtures\formula-ir.invalid.json',
    'examples\fixtures\formula-unicodemath.valid.json',
    'examples\fixtures\formula-unicodemath.invalid.json',
    'examples\fixtures\formula-goldset.sample.csv',
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

$formulaProcessing = $config.PSObject.Properties['formulaProcessing']
if ($null -eq $formulaProcessing -or $null -eq $formulaProcessing.Value) {
    throw 'Config formulaProcessing section is missing.'
}
$formulaProcessingValue = $formulaProcessing.Value
if ([int]$formulaProcessingValue.schemaVersion -ne 1) {
    throw "Unsupported formulaProcessing schemaVersion: $($formulaProcessingValue.schemaVersion)"
}
$allowedFormulaModes = @('ReportOnly', 'CandidateOnly', 'ReviewRequired', 'ClosedWorldUnattended', 'ExplicitMigration')
if ([string]$formulaProcessingValue.defaultMode -notin $allowedFormulaModes) {
    throw "Config formulaProcessing.defaultMode is invalid: $($formulaProcessingValue.defaultMode)"
}
if ([string]$formulaProcessingValue.defaultMode -ne 'ReportOnly') {
    throw 'Config formulaProcessing.defaultMode must remain ReportOnly for the safe default mainline.'
}
foreach ($mode in @($formulaProcessingValue.allowedModes)) {
    if ([string]$mode -notin $allowedFormulaModes) { throw "Config formulaProcessing.allowedModes contains an invalid mode: $mode" }
}
foreach ($safeDefault in @('writeBackEnabled', 'imageFormulaWriteBackEnabled', 'mathTypeMigrationEnabled', 'ocrEnabled', 'modelInferenceEnabled')) {
    $safeDefaultProperty = $formulaProcessingValue.PSObject.Properties[$safeDefault]
    if ($null -eq $safeDefaultProperty -or $safeDefaultProperty.Value -isnot [bool]) {
        throw "Config formulaProcessing.$safeDefault must be boolean"
    }
    if ([bool]$safeDefaultProperty.Value) {
        throw "Config formulaProcessing.$safeDefault must remain false by default"
    }
}

function Test-FormulaIrFixtureContract {
    param([Parameter(Mandatory = $true)]$Record)
    $issues = New-Object System.Collections.Generic.List[string]
    if ([int]$Record.schemaVersion -ne 1) { $issues.Add('schemaVersion') | Out-Null }
    if ([string]::IsNullOrWhiteSpace([string]$Record.recordId)) { $issues.Add('recordId') | Out-Null }
    $source = $Record.source
    foreach ($field in @('filePath', 'fileSha256', 'sourceSha256', 'carrier', 'slide', 'shapeId', 'shapeName', 'bbox')) {
        if ($null -eq $source -or $null -eq $source.PSObject.Properties[$field]) { $issues.Add("source.$field") | Out-Null }
    }
    if ($null -ne $source -and $null -ne $source.PSObject.Properties['fileSha256'] -and [string]$source.fileSha256 -notmatch '^[A-Fa-f0-9]{64}$') { $issues.Add('source.fileSha256') | Out-Null }
    if ($null -ne $source -and $null -ne $source.PSObject.Properties['sourceSha256'] -and [string]$source.sourceSha256 -notmatch '^[A-Fa-f0-9]{64}$') { $issues.Add('source.sourceSha256') | Out-Null }
    $detection = $Record.detection
    foreach ($field in @('method', 'status', 'confidence', 'rawCandidates')) {
        if ($null -eq $detection -or $null -eq $detection.PSObject.Properties[$field]) { $issues.Add("detection.$field") | Out-Null }
    }
    $canonical = $Record.canonical
    foreach ($field in @('status', 'source', 'unicodeMath', 'tex', 'mathml', 'tokens')) {
        if ($null -eq $canonical -or $null -eq $canonical.PSObject.Properties[$field]) { $issues.Add("canonical.$field") | Out-Null }
    }
    $decision = $Record.decision
    foreach ($field in @('mode', 'status', 'targetCarrier', 'reason')) {
        if ($null -eq $decision -or $null -eq $decision.PSObject.Properties[$field]) { $issues.Add("decision.$field") | Out-Null }
    }
    $evidence = $Record.evidence
    if ($null -eq $evidence -or $null -eq $evidence.PSObject.Properties['paths'] -or @($evidence.paths).Count -lt 1) { $issues.Add('evidence.paths') | Out-Null }
    if ($null -eq $evidence -or [string]::IsNullOrWhiteSpace([string]$evidence.rollbackPath)) { $issues.Add('evidence.rollbackPath') | Out-Null }
    if ($null -ne $decision -and $null -ne $decision.PSObject.Properties['status'] -and [string]$decision.status -in @('Converted', 'FallbackSvg')) {
        if ($null -eq $canonical -or [string]$canonical.status -ne 'Resolved') { $issues.Add('converted-without-resolved-canonical') | Out-Null }
        if ($null -eq $canonical -or $null -eq $canonical.PSObject.Properties['source'] -or $null -eq $canonical.source -or [string]$canonical.source.kind -in @('', 'None')) { $issues.Add('converted-without-canonical-source') | Out-Null }
    }
    return @($issues.ToArray())
}

$validFormulaIr = Get-Content -LiteralPath (Join-Path $root 'examples\fixtures\formula-ir.valid.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$validFormulaIrIssues = @(Test-FormulaIrFixtureContract -Record $validFormulaIr)
if ($validFormulaIrIssues.Count -gt 0) { throw "Valid FormulaIR fixture failed: $($validFormulaIrIssues -join ', ')" }
$invalidFormulaIr = Get-Content -LiteralPath (Join-Path $root 'examples\fixtures\formula-ir.invalid.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$invalidFormulaIrIssues = @(Test-FormulaIrFixtureContract -Record $invalidFormulaIr)
if ($invalidFormulaIrIssues.Count -eq 0) { throw 'Invalid FormulaIR fixture was accepted.' }

# --- 2b. UnicodeMath parser fixture contract (executes the exporter's own parser) ---
# Common is dotted early because the extracted parser functions call shared
# classifiers (Get-PhysicsTokenRole); section 8 re-dots it harmlessly.
. (Join-Path $root 'tools\PhysicsPpt.Common.ps1')
$unicodeMathValidPath = Join-Path $root 'examples\fixtures\formula-unicodemath.valid.json'
$unicodeMathInvalidPath = Join-Path $root 'examples\fixtures\formula-unicodemath.invalid.json'
$unicodeMathValid = Get-Content -LiteralPath $unicodeMathValidPath -Raw -Encoding UTF8 | ConvertFrom-Json
$unicodeMathInvalid = Get-Content -LiteralPath $unicodeMathInvalidPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$unicodeMathValid.schemaVersion -ne 1 -or [int]$unicodeMathInvalid.schemaVersion -ne 1) {
    throw 'UnicodeMath fixtures must use schemaVersion 1.'
}
$ommlCandidatesParserPath = Join-Path $root 'tools\Export-FormulaOmmlCandidates.ps1'
$parserTokens = $null
$parserParseErrors = $null
$ommlCandidatesParserAst = [System.Management.Automation.Language.Parser]::ParseFile($ommlCandidatesParserPath, [ref]$parserTokens, [ref]$parserParseErrors)
if (@($parserParseErrors).Count -gt 0) { throw "OMML candidate exporter failed to parse: $($ommlCandidatesParserPath)" }
$requiredParserFunctions = @(
    'Convert-SuperscriptSubscriptCharsToAscii', 'Get-FormulaTokens', 'Parse-FormulaAtom',
    'Parse-FormulaSequence', 'Parse-FormulaExpression', 'Parse-FormulaAst', 'Find-FormulaStructureIssue',
    'Get-FormulaTokenRole', 'Convert-FormulaTokenToTex', 'Convert-AstToUnicodeMath',
    'New-OmmlFragment', 'New-Element', 'Add-TextElement', 'Add-OmmlRun', 'New-RunProperties', 'Add-OmmlNode'
)
$parserFunctions = @{}
foreach ($functionAst in @($ommlCandidatesParserAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))) {
    if ($functionAst.Name -in $requiredParserFunctions -and -not $parserFunctions.ContainsKey($functionAst.Name)) {
        $parserFunctions[$functionAst.Name] = $functionAst
    }
}
$missingParserFunctions = @($requiredParserFunctions | Where-Object { -not $parserFunctions.ContainsKey($_) })
if ($missingParserFunctions.Count -gt 0) {
    throw "OMML candidate exporter is missing UnicodeMath parser functions: $($missingParserFunctions -join ', ')"
}
foreach ($parserName in $requiredParserFunctions) {
    # Invoke-Expression of the full definition text is the reliable way to load
    # these functions for execution; Set-Item Function: with a body scriptblock
    # misbinds under pwsh 7.6 (calls emit the body text instead of executing it).
    Invoke-Expression $parserFunctions[$parserName].Extent.Text
}
function Test-UnicodeMathRoundTrip {
    param([string]$Text)
    $mathAst = Parse-FormulaAst -UnicodeMath $Text
    $canonical = Convert-AstToUnicodeMath -Node $mathAst
    $reparseAst = Parse-FormulaAst -UnicodeMath $canonical
    $recanonical = Convert-AstToUnicodeMath -Node $reparseAst
    return @{ Canonical = $canonical; RoundTripStable = ($canonical -ceq $recanonical) }
}
foreach ($validCase in @($unicodeMathValid.cases)) {
    $caseResult = Test-UnicodeMathRoundTrip -Text ([string]$validCase.unicodeMath)
    if (-not $caseResult.RoundTripStable) { throw "UnicodeMath valid case is not round-trip stable: $($validCase.name)" }
    $expectedProperty = $validCase.PSObject.Properties['expectedCanonicalUnicodeMath']
    if ($null -ne $expectedProperty -and $caseResult.Canonical -cne [string]$expectedProperty.Value) {
        throw ("UnicodeMath valid case canonical mismatch: {0}; expected '{1}', actual '{2}'" -f $validCase.name, $expectedProperty.Value, $caseResult.Canonical)
    }
}
foreach ($invalidCase in @($unicodeMathInvalid.cases)) {
    $threw = $false
    try { $null = Test-UnicodeMathRoundTrip -Text ([string]$invalidCase.unicodeMath) } catch { $threw = $true }
    if (-not $threw) { throw "UnicodeMath invalid case was accepted: $($invalidCase.name)" }
}

# Structural OMML regression guard: superscript groups must render as raised
# scripts without visible parentheses, and the escaped linear slash must stay
# a division run instead of becoming a stacked fraction.
# These mirror Export-FormulaOmmlCandidates.ps1's own $script: constants: the
# parser functions are Invoke-Expression'd into THIS scope (section 2b), so
# their bodies resolve these names here. Do not remove them as "unused" — the
# references are invisible to single-file scans.
$script:NsA14 = 'http://schemas.microsoft.com/office/drawing/2010/main'
$script:NsA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$script:NsM = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
$script:FormulaColorHex = '000000'
$script:FormulaSizeHundredths = 3800
$script:FormulaEastAsianFontName = '宋体'
$supFragmentXml = (New-OmmlFragment -UnicodeMath '10^(-3)').OuterXml
if ($supFragmentXml -notmatch ':sSup') { throw 'Superscript OMML regression: 10^(-3) no longer produces m:sSup.' }
if ($supFragmentXml -match '\(|\)') { throw 'Superscript OMML regression: parentheses leaked into the m:sup script runs.' }
$slashFragmentXml = (New-OmmlFragment -UnicodeMath 'J\/(kg·℃)').OuterXml
if ($slashFragmentXml -match '<m:f>') { throw 'Linear slash OMML regression: J\/(kg·℃) produced a stacked fraction.' }
if ($slashFragmentXml -notmatch '>/<') { throw 'Linear slash OMML regression: escaped slash run text is missing.' }

# --- 2c. Formula recognition gold set contract ---

# The exporter covers both parameter sets: GoldSet construction and the
# evidence-availability verification mode (former standalone availability
# exporter). Evidence-availability receipts keep their schema above.
$goldSetSampleRows = @(Import-Csv -LiteralPath (Join-Path $root 'examples\fixtures\formula-goldset.sample.csv'))
if ($goldSetSampleRows.Count -lt 1) { throw 'Formula gold set sample fixture must contain at least one row.' }
$goldSetSampleHeader = @($goldSetSampleRows[0].PSObject.Properties.Name)
foreach ($goldSetColumn in @('ReviewStatus', 'Slide', 'ShapeIds', 'WhitelistName', 'SourceFormulaText', 'SizePt', 'MainColorHex', 'SubColorHex', 'Note', 'EvidencePath')) {
    if ($goldSetColumn -notin $goldSetSampleHeader) { throw "Formula gold set sample fixture is missing column: $goldSetColumn" }
}
foreach ($goldSetSampleRow in $goldSetSampleRows) {
    if ([string]$goldSetSampleRow.ReviewStatus -notin @('Draft', 'Approved')) { throw 'Formula gold set sample fixture has invalid ReviewStatus.' }
}

& (Join-Path $root 'tools\Test-FormulaCanonicalContext.ps1')
if (-not $?) { throw 'Formula context resolver fixture test failed.' }
& (Join-Path $root 'tools\Test-FormulaOleVisualAdjudication.ps1')
if (-not $?) { throw 'Formula OLE visual adjudication fixture test failed.' }
& (Join-Path $root 'tools\Test-FormulaOfficeMathValidator.ps1')
if (-not $?) { throw 'FormulaOfficeMathValidator fault-injection test failed.' }
& (Join-Path $root 'tools\Test-ClosedWorldCircuitBreaker.ps1')
if (-not $?) { throw 'Closed-world false-acceptance circuit-breaker test failed.' }

$requiredFontFields = @('chinese', 'compactChinese', 'latin', 'math')
foreach ($field in $requiredFontFields) {
    if ([string]::IsNullOrWhiteSpace($config.fonts.$field)) { throw "Config fonts.$field is missing or empty" }
}

$requiredSizeFields = @('title1', 'sectionTitle', 'title2', 'body', 'bodyMax', 'displayTitleMin', 'auxiliary', 'minimum', 'tableHeader', 'tableBody', 'formulaInline', 'formulaStandalone', 'formulaCore', 'footer')
foreach ($field in $requiredSizeFields) {
    if ($null -eq $config.fontSizes.$field -or $config.fontSizes.$field -isnot [ValueType] -or [double]$config.fontSizes.$field -ne [math]::Floor([double]$config.fontSizes.$field) -or $config.fontSizes.$field -lt 8 -or $config.fontSizes.$field -gt 96) { throw "Config fontSizes.$field must be an integer between 8 and 96" }
}
if ($config.fontSizes.minimum -gt $config.fontSizes.body) { throw 'Config fontSizes.minimum must not exceed body.' }
if ($config.fontSizes.bodyMax -lt $config.fontSizes.body) { throw 'Config fontSizes.bodyMax must not be smaller than body.' }
if ($config.fontSizes.displayTitleMin -lt $config.fontSizes.bodyMax) { throw 'Config fontSizes.displayTitleMin must not be smaller than bodyMax.' }

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
    # Every target must render through the canonical parser: a config typo
    # would otherwise only fail closed mid-production batch.
    $renderThrew = $false
    $ommlFragment = $null
    try { $ommlFragment = New-OmmlFragment -UnicodeMath ([string]$rule.targetUnicodeMath) } catch { $renderThrew = $true }
    if ($renderThrew -or $null -eq $ommlFragment -or $ommlFragment.OuterXml -notmatch ':oMath') {
        throw ("Config formulaWhitelist targetUnicodeMath is not renderable by the canonical parser: {0} => {1}" -f $rule.name, $rule.targetUnicodeMath)
    }
}

# Strong layout constraint for reports/: the root may only contain versioned
# delivery directories (<deck>_v<N>) and _archive. Anything else (run_*,
# ad-hoc names, legacy dotted variants) fails the minimum gate. A _v<N> suffix
# directly attached to an ASCII word (normalized_v1, brand_v2, __pptx_v1) is
# the legacy dotted-naming fingerprint and is rejected as well.
$reportsRoot = Join-Path $root 'reports'
if (Test-Path -LiteralPath $reportsRoot) {
    $layoutViolations = @()
    foreach ($entry in @(Get-ChildItem -LiteralPath $reportsRoot -Force)) {
        if (-not $entry.PSIsContainer) { continue }
        if ($entry.Name -eq '_archive') { continue }
        if ($entry.Name -match '^.+?_v(\d+)$' -and $entry.Name -notmatch '^.+[A-Za-z]_v(\d+)$') { continue }
        $layoutViolations += $entry.Name
    }
    if ($layoutViolations.Count -gt 0) {
        throw ("reports/ root layout violation: only <deck>_v<N> delivery dirs and _archive are allowed; move offenders into reports/_archive first (offenders: " + ($layoutViolations -join ', ') + ")")
    }
}

$packagePath = Join-Path $root 'package.json'
$packageJson = Get-Content -LiteralPath $packagePath -Raw -Encoding UTF8 | ConvertFrom-Json
$sharpDependency = $packageJson.dependencies.PSObject.Properties['sharp']
if ($null -eq $sharpDependency -or [string]::IsNullOrWhiteSpace([string]$sharpDependency.Value)) {
    throw "package.json dependencies.sharp is missing"
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
    @{ VbaConst = 'SIZE_BODY_MAX';        JsonCat = 'fontSizes'; JsonKey = 'bodyMax';    Type = 'Numeric' },
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
        # A renamed or deleted constant must fail the gate, not silently drop
        # the sync guarantee.
        throw "Sync check: could not find VBA constant $($check.VbaConst) in PhysicsPptCommon.bas"
    }
}

# Highlight-box colors: the VBA constants carry precomputed RGB Long values
# (R + G*256 + B*65536). Verify them against the configured hex values.
$colorSyncChecks = @(
    @{ VbaConst = 'COLOR_WHITE';         JsonKey = 'white' },
    @{ VbaConst = 'COLOR_BODY';          JsonKey = 'body' },
    @{ VbaConst = 'COLOR_YELLOW_FILL';   JsonKey = 'yellowFill' },
    @{ VbaConst = 'COLOR_YELLOW_BORDER'; JsonKey = 'yellowBorder' }
)
foreach ($colorCheck in $colorSyncChecks) {
    $jsonHex = [string]$config.colors.($colorCheck.JsonKey)
    if ($jsonHex -notmatch '^#[0-9A-Fa-f]{6}$') { throw "Config color colors.$($colorCheck.JsonKey) is not a six-digit #RRGGBB value." }
    $hexDigits = $jsonHex.TrimStart('#')
    # VBA RGB Long layout: R + G*256 + B*65536 (e.g. #FFF2CC -> 0x00CCF2FF).
    $expected = ([int]::Parse($hexDigits.Substring(0, 2), [System.Globalization.NumberStyles]::HexNumber)) +
        ([int]::Parse($hexDigits.Substring(2, 2), [System.Globalization.NumberStyles]::HexNumber) * 256) +
        ([int]::Parse($hexDigits.Substring(4, 2), [System.Globalization.NumberStyles]::HexNumber) * 65536)
    $pattern = '(?m)Public\s+Const\s+' + [regex]::Escape($colorCheck.VbaConst) + '\s+As\s+Long\s*=\s*(\d+)'
    if ($vbaContent -match $pattern) {
        if ([int64]$Matches[1] -ne [int64]$expected) {
            throw "JSON <-> VBA color mismatch: $($colorCheck.VbaConst) = $($Matches[1]) in VBA but config colors.$($colorCheck.JsonKey) ($jsonHex) = $expected"
        }
    } else {
        throw "Sync check: could not find VBA color constant $($colorCheck.VbaConst) in PhysicsPptCommon.bas"
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
if ($normalizeContent -match '&\s+powershell\.exe') { throw 'Parallel PowerShell workers must resolve the PS7-first host instead of hard-coding powershell.exe.' }
if ($normalizeContent -notmatch 'Resolve-PowerShellHost') { throw 'Normalize script must use Resolve-PowerShellHost for child workers.' }

foreach ($launcher in @('一键规范化并导出PDF.cmd', '一键检查PPT.cmd', '一键规范化导出并转换可编辑公式.cmd')) {
    $launcherPath = Join-Path $root $launcher
    # cmd.exe does not strip a UTF-8 BOM: the first line turns into garbage
    # before @echo off and the launcher prints an error on every run.
    $launcherBytes = [System.IO.File]::ReadAllBytes($launcherPath)
    if ($launcherBytes.Length -ge 3 -and $launcherBytes[0] -eq 0xEF -and $launcherBytes[1] -eq 0xBB -and $launcherBytes[2] -eq 0xBF) { throw "Launcher must not carry a UTF-8 BOM: $launcher" }
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
    'tools\Apply-PptxVisualAuditFixes.ps1',
    'tools\Apply-PptxHighlightBoxStyle.ps1',
    'tools\Assert-Toolchain.ps1'
)) {
    $automationContent = Get-Content -LiteralPath (Join-Path $root $automationScript) -Raw -Encoding UTF8
    if ($automationContent -match 'New-Object\s+-ComObject\s+PowerPoint\.Application') {
        throw "PowerPoint automation script must use New-PowerPointApplication: $automationScript"
    }
}
if ($normalizeContent -match [regex]::Escape('$font.Fill.ForeColor.RGB = $Color')) { throw 'Normalize body path must never write author text colors.' }
if ($normalizeContent -match [regex]::Escape('$font.Bold = $(if ($Bold) { $script:MsoTrue } else { $script:MsoFalse })')) { throw 'Normalize body path must never write Bold=false over author bold runs.' }

# Geometry heal write-back re-serializes slide XML verbatim; a
# PreserveWhitespace=$false Load drops whitespace-only <a:t> runs and the
# normalized copy silently loses formula spacing (regression: 14.1 slide9
# 'Q放 = qm' arrived as 'Q放= qm' in the v39 invariant gate).
if ($normalizeContent -notmatch '(?s)function Read-GeometrySlideXmlDocument \{.*?PreserveWhitespace\s*=\s*\$true') {
    throw 'Geometry slide XML loader must set PreserveWhitespace=$true before Load to keep whitespace-only a:t runs.'
}
$whitespaceRunDoc = New-Object System.Xml.XmlDocument
$whitespaceRunDoc.PreserveWhitespace = $true
$whitespaceRunDoc.LoadXml('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:txBody><a:p><a:r><a:t> </a:t></a:r><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>')
$whitespaceRunSettings = New-Object System.Xml.XmlWriterSettings
$whitespaceRunSettings.Indent = $false
$whitespaceRunStream = New-Object System.IO.MemoryStream
try {
    $whitespaceRunWriter = [System.Xml.XmlWriter]::Create($whitespaceRunStream, $whitespaceRunSettings)
    try {
        $whitespaceRunDoc.Save($whitespaceRunWriter)
        $whitespaceRunWriter.Flush()
    } finally {
        $whitespaceRunWriter.Dispose()
    }
    $whitespaceRunXml = [System.Text.Encoding]::UTF8.GetString($whitespaceRunStream.ToArray())
} finally {
    $whitespaceRunStream.Dispose()
}
if ($whitespaceRunXml -notmatch '<a:t> </a:t>') { throw 'Whitespace-only a:t round trip failed; geometry heal write-back would erase formula spacing.' }

# OLE content-fingerprint guard: GoldSet SourceFormulaText must agree with
# the MathType embedding's actual CJK characters. The v37-era goldset
# attributed slide16 sh13/sh16 to each other's formulas (Q吸=cmΔt vs Q放=qm)
# and passed every schema check — only the embedding bytes know what the OLE
# really holds, so the mapping must verify the fingerprint.
$fingerprintCommonAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'tools\PhysicsPpt.Common.ps1'), [ref]$null, [ref]$null)

# Deck-stem canonicalization probe: stage-suffix stripping must fully resolve
# compound lineage names (X.brand.callout) to the clean stem. A half-stripped
# stem (X.brand) yields reports/X.brand_v<N>, a layout-gate violation
# (a letter directly before _vN).
$stemFunctionAst = $fingerprintCommonAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-CanonicalDeckStem' }, $true) | Select-Object -First 1
if ($null -eq $stemFunctionAst) { throw 'Get-CanonicalDeckStem function definition not found for the behavioral probe.' }
Invoke-Expression $stemFunctionAst.Extent.Text
$stemProbeCases = @(
    @('plain deck stem', '13.3比热容（王耀强）.pptx', '13.3比热容（王耀强）'),
    @('normalized input', 'X.normalized.pptx', 'X'),
    @('brand input', 'X.brand.pptx', 'X'),
    @('callout input', 'X.callout.pptx', 'X'),
    @('compound lineage input', 'X.brand.callout.pptx', 'X'),
    @('full lineage input', 'X.normalized.brand.callout.pptx', 'X')
)
foreach ($stemCase in $stemProbeCases) {
    $stemActual = Get-CanonicalDeckStem -FileName ([string]$stemCase[1])
    if ([string]$stemActual -cne [string]$stemCase[2]) {
        throw ("Deck-stem probe failed for {0}: expected '{1}', got '{2}'." -f $stemCase[0], $stemCase[2], $stemActual)
    }
}

$fingerprintFunctionAst = $fingerprintCommonAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-OleEquationCjkFingerprint' }, $true) | Select-Object -First 1
if ($null -eq $fingerprintFunctionAst) { throw 'Get-OleEquationCjkFingerprint function definition not found for the behavioral probe.' }
Invoke-Expression $fingerprintFunctionAst.Extent.Text
$fingerprintProbeDir = Join-Path $env:TEMP ("fingerprint-probe-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fingerprintProbeDir -Force | Out-Null
try {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $fingerprintProbeZip = Join-Path $fingerprintProbeDir 'probe.pptx'
    $probeArchive = [System.IO.Compression.ZipFile]::Open($fingerprintProbeZip, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $probeUtf8 = New-Object System.Text.UTF8Encoding($false)
        $probeSlideEntry = $probeArchive.CreateEntry('ppt/slides/slide1.xml')
        $probeSlideWriter = New-Object System.IO.StreamWriter($probeSlideEntry.Open(), $probeUtf8)
        try {
            $probeSlideWriter.Write('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="probe"/></p:nvGraphicFramePr><p:graphicData><oleObj r:id="rId5"/></p:graphicData></p:graphicFrame>')
        } finally { $probeSlideWriter.Dispose() }
        $probeRelsEntry = $probeArchive.CreateEntry('ppt/slides/_rels/slide1.xml.rels')
        $probeRelsWriter = New-Object System.IO.StreamWriter($probeRelsEntry.Open(), $probeUtf8)
        try {
            $probeRelsWriter.Write('<Relationships><Relationship Id="rId5" Type="http://p.invalid/ole" Target="../embeddings/ole1.bin"/></Relationships>')
        } finally { $probeRelsWriter.Dispose() }
        $probeBinEntry = $probeArchive.CreateEntry('ppt/embeddings/ole1.bin')
        $probeBinStream = $probeBinEntry.Open()
        try {
            $probePayload = [System.Text.Encoding]::Unicode.GetBytes('Q放=qm')
            $probeBinStream.Write($probePayload, 0, $probePayload.Length)
        } finally { $probeBinStream.Dispose() }
        $probePresEntry = $probeArchive.CreateEntry('ppt/presentation.xml')
        $probePresWriter = New-Object System.IO.StreamWriter($probePresEntry.Open(), $probeUtf8)
        try {
            $probePresWriter.Write('<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>')
        } finally { $probePresWriter.Dispose() }
        $probePresRelsEntry = $probeArchive.CreateEntry('ppt/_rels/presentation.xml.rels')
        $probePresRelsWriter = New-Object System.IO.StreamWriter($probePresRelsEntry.Open(), $probeUtf8)
        try {
            $probePresRelsWriter.Write('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>')
        } finally { $probePresRelsWriter.Dispose() }
    } finally { $probeArchive.Dispose() }
    $probeFingerprint = Get-OleEquationCjkFingerprint -PptxPath $fingerprintProbeZip -SlideNumber 1 -ShapeId 7
    if ($probeFingerprint -cne '放') { throw "OLE fingerprint probe failed: expected 放, got '$probeFingerprint'." }
    $probeMissing = @(@('放', '吸') | Where-Object { -not $probeFingerprint.Contains($_) })
    if ($probeMissing.Count -ne 1 -or $probeMissing[0] -cne '吸') { throw 'OLE fingerprint subset probe failed: 吸 must be reported missing from a Q放=qm embedding.' }
} finally {
    if (Test-Path -LiteralPath $fingerprintProbeDir) { Remove-Item -LiteralPath $fingerprintProbeDir -Recurse -Force }
}

# Slide-order domain probe: slideN.xml names are creation order and drift
# from the presentation order once a deck is reordered in PowerPoint. Every
# slide-locating helper must resolve through the sldIdLst map, or guards
# silently miss on reordered decks (18.2 blipFill re-bake class).
$normalizeReorderAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'tools\Normalize-PhysicsPpt.ps1'), [ref]$null, [ref]$null)
$reorderNormalizeFunctions = @('Read-GeometrySlideXmlDocument', 'Get-PictureFillShapeIds')
$reorderFunctionText = ''
foreach ($fnName in $reorderNormalizeFunctions) {
    $fnAst = $normalizeReorderAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $fnName }, $true) | Select-Object -First 1
    if ($null -eq $fnAst) { throw "Normalize slide-order function not found for the behavioral probe: $fnName" }
    $reorderFunctionText += "`n" + $fnAst.Extent.Text
}
Invoke-Expression $reorderFunctionText
$reorderProbeDir = Join-Path $env:TEMP ("reorder-probe-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $reorderProbeDir -Force | Out-Null
try {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $reorderProbeZip = Join-Path $reorderProbeDir 'reordered.pptx'
    $reorderArchive = [System.IO.Compression.ZipFile]::Open($reorderProbeZip, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $reorderUtf8 = New-Object System.Text.UTF8Encoding($false)
        # Physical slide1.xml holds the OLE (id=7); physical slide2.xml holds
        # the picture-filled text box (id=9). Presentation order swaps them:
        # presentation page 1 = slide2.xml, page 2 = slide1.xml.
        $r1Entry = $reorderArchive.CreateEntry('ppt/slides/slide1.xml')
        $r1Writer = New-Object System.IO.StreamWriter($r1Entry.Open(), $reorderUtf8)
        try {
            $r1Writer.Write('<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="probe"/></p:nvGraphicFramePr><p:graphicData><oleObj r:id="rId5"/></p:graphicData></p:graphicFrame></p:spTree>')
        } finally { $r1Writer.Dispose() }
        $r2Entry = $reorderArchive.CreateEntry('ppt/slides/slide2.xml')
        $r2Writer = New-Object System.IO.StreamWriter($r2Entry.Open(), $reorderUtf8)
        try {
            $r2Writer.Write('<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sp><p:nvSpPr><p:cNvPr id="9" name="probe-fill"/></p:nvSpPr><p:spPr><a:blipFill><a:blip r:embed="rId3"/></a:blipFill></p:spPr></p:sp></p:spTree>')
        } finally { $r2Writer.Dispose() }
        $rPresEntry = $reorderArchive.CreateEntry('ppt/presentation.xml')
        $rPresWriter = New-Object System.IO.StreamWriter($rPresEntry.Open(), $reorderUtf8)
        try {
            $rPresWriter.Write('<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rIdB"/><p:sldId id="257" r:id="rIdA"/></p:sldIdLst></p:presentation>')
        } finally { $rPresWriter.Dispose() }
        $rPresRelsEntry = $reorderArchive.CreateEntry('ppt/_rels/presentation.xml.rels')
        $rPresRelsWriter = New-Object System.IO.StreamWriter($rPresRelsEntry.Open(), $reorderUtf8)
        try {
            $rPresRelsWriter.Write('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rIdB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>')
        } finally { $rPresRelsWriter.Dispose() }
        $rS1RelsEntry = $reorderArchive.CreateEntry('ppt/slides/_rels/slide1.xml.rels')
        $rS1RelsWriter = New-Object System.IO.StreamWriter($rS1RelsEntry.Open(), $reorderUtf8)
        try {
            $rS1RelsWriter.Write('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://p.invalid/ole" Target="../embeddings/ole1.bin"/></Relationships>')
        } finally { $rS1RelsWriter.Dispose() }
        $rBinEntry = $reorderArchive.CreateEntry('ppt/embeddings/ole1.bin')
        $rBinStream = $rBinEntry.Open()
        try {
            $rBinPayload = [System.Text.Encoding]::Unicode.GetBytes('Q放=qm')
            $rBinStream.Write($rBinPayload, 0, $rBinPayload.Length)
        } finally { $rBinStream.Dispose() }
    } finally { $reorderArchive.Dispose() }
    # The OLE lives on presentation page 2 (physical slide1.xml). A physical-
    # name implementation would look inside slide2.xml and fail.
    $reorderFingerprint = Get-OleEquationCjkFingerprint -PptxPath $reorderProbeZip -SlideNumber 2 -ShapeId 7
    if ($reorderFingerprint -cne '放') { throw "Reordered-deck fingerprint probe failed: expected 放 via presentation order, got '$reorderFingerprint'." }
    # The picture-fill map must key by presentation order: the fill shape is
    # on presentation page 1 (physical slide2.xml), shape id 9.
    $reorderFillMap = Get-PictureFillShapeIds -SourcePath $reorderProbeZip
    if (-not $reorderFillMap.ContainsKey(1) -or @($reorderFillMap[1]) -notcontains 9) {
        throw "Reordered-deck picture-fill probe failed: expected shape 9 under presentation page 1, got keys [$(@($reorderFillMap.Keys) -join ',')]."
    }
    if ($reorderFillMap.ContainsKey(2)) { throw 'Reordered-deck picture-fill probe failed: presentation page 2 must not carry picture-fill ids.' }
} finally {
    if (Test-Path -LiteralPath $reorderProbeDir) { Remove-Item -LiteralPath $reorderProbeDir -Recurse -Force }
}

# OLE block id probe: the apply step resolves slide+shapeId as the primary
# mapping identity, and the reused cNvPr id keeps animation spids valid. The
# id extractor must cover every authoring layout — before the 15.4电流的测量
# fix, a bare p:graphicFrame with mc:AlternateContent inside a:graphicData
# (legacy authoring) returned 0 and failed mapping with OleIndexOutOfRange.
$oleApplyAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'tools\Apply-FormulaOmmlForOle.ps1'), [ref]$null, [ref]$null)
$oleBlockIdFn = $oleApplyAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-OleBlockShapeId' }, $true) | Select-Object -First 1
if ($null -eq $oleBlockIdFn) { throw 'OLE apply function Get-OleBlockShapeId not found for the behavioral probe.' }
Invoke-Expression $oleBlockIdFn.Extent.Text
$oleIdProbeNs = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
$oleIdXmlLegacy = '<p:graphicFrame ' + $oleIdProbeNs + '><p:nvGraphicFramePr><p:cNvPr id="8" name="legacy"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="ole"><mc:AlternateContent><mc:Choice Requires="v"><p:oleObj progId="Equation.DSMT4"/></mc:Choice><mc:Fallback><p:oleObj progId="Equation.DSMT4"><p:pic><p:nvPicPr><p:cNvPr id="8" name="legacy"/></p:nvPicPr></p:pic></p:oleObj></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></p:graphicFrame>'
$oleIdXmlAcFrameChild = '<mc:AlternateContent ' + $oleIdProbeNs + '><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="13" name="ac-frame"/></p:nvGraphicFramePr></p:graphicFrame><mc:Fallback/></mc:AlternateContent>'
$oleIdXmlAcChoice = '<mc:AlternateContent ' + $oleIdProbeNs + '><mc:Choice Requires="v"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="11" name="modern"/></p:nvGraphicFramePr></p:graphicFrame></mc:Choice><mc:Fallback/></mc:AlternateContent>'
$oleIdXmlFallbackOnly = '<mc:AlternateContent ' + $oleIdProbeNs + '><mc:Choice Requires="v"><p:oleObj progId="Equation.DSMT4"/></mc:Choice><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="21" name="fallback-only"/></p:nvPicPr></p:pic></mc:Fallback></mc:AlternateContent>'
$oleIdProbeCases = @(
    @('legacy bare graphicFrame (AC inside a:graphicData)', $oleIdXmlLegacy, 8),
    @('top-level AC with bare graphicFrame child', $oleIdXmlAcFrameChild, 13),
    @('top-level AC with graphicFrame inside Choice (PowerPoint 2016+)', $oleIdXmlAcChoice, 11),
    @('AC without a resolvable frame id must stay 0 (Fallback pic is never used)', $oleIdXmlFallbackOnly, 0)
)
foreach ($oleIdCase in $oleIdProbeCases) {
    $oleIdDoc = New-Object System.Xml.XmlDocument
    $oleIdDoc.LoadXml([string]$oleIdCase[1])
    $oleIdActual = Get-OleBlockShapeId -Block $oleIdDoc.DocumentElement -NsManager $null
    if ([int]$oleIdActual -ne [int]$oleIdCase[2]) {
        throw ("OLE block id probe failed for {0}: expected id {1}, got {2}." -f $oleIdCase[0], $oleIdCase[2], $oleIdActual)
    }
}

# Line-layout guard decision probe: the wrap-point comparison is the last
# defense against same-geometry re-wraps that hide sibling answer text
# (slide 16 blank-collision class). The COM capture (Lines()) is real-host
# only, but the decision semantics must hold: an unavailable probe stays
# inert, and any line-count or per-line text change is a relayout.
$layoutDecisionText = ''
foreach ($layoutFnName in @('Test-TextRangeLayoutUnchanged', 'Get-TextRangeLayoutChangeText', 'Get-AutoSizeGeometryDrift')) {
    $layoutFnAst = $normalizeReorderAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $layoutFnName }, $true) | Select-Object -First 1
    if ($null -eq $layoutFnAst) { throw "Normalize guard function not found for the behavioral probe: $layoutFnName" }
    $layoutDecisionText += "`n" + $layoutFnAst.Extent.Text
}
Invoke-Expression $layoutDecisionText
$layoutBefore = @('答：', 'A', 'B')
if (-not (Test-TextRangeLayoutUnchanged -Before $layoutBefore -After @('答：', 'A', 'B'))) { throw 'Line-layout probe: identical layouts must be unchanged.' }
if (-not (Test-TextRangeLayoutUnchanged -Before $null -After $null)) { throw 'Line-layout probe: an unavailable capture must stay inert (unchanged).' }
if (Test-TextRangeLayoutUnchanged -Before $layoutBefore -After @('答：', 'A')) { throw 'Line-layout probe: line-count change must be detected as a relayout.' }
if (Test-TextRangeLayoutUnchanged -Before $layoutBefore -After @('答：', 'A、', 'B')) { throw 'Line-layout probe: same-count per-line text change must be detected as a relayout.' }
$layoutCountChangeText = Get-TextRangeLayoutChangeText -Before $layoutBefore -After @('答：', 'A')
if ($layoutCountChangeText -notmatch 'rendered lines 3 -> 2; first changed line 3') { throw "Line-layout probe: unexpected count-change report: $layoutCountChangeText" }
$layoutTextChangeReport = Get-TextRangeLayoutChangeText -Before $layoutBefore -After @('答：', 'A、', 'B')
if ($layoutTextChangeReport -notmatch 'first changed line 2') { throw "Line-layout probe: unexpected same-count change report: $layoutTextChangeReport" }

# AutoSize geometry drift probe: the rollback trigger measures the max
# per-axis delta across Left/Top/Width/Height. A partial-axis rewrite or an
# average would silently let an AutoSize reflow through (18.2 AutoSize class).
$driftShape = [pscustomobject]@{ Left = 100.0; Top = 50.0; Width = 300.0; Height = 80.0 }
$zeroDrift = Get-AutoSizeGeometryDrift -Shape $driftShape -Left 100.0 -Top 50.0 -Width 300.0 -Height 80.0
if ([Math]::Abs([double]$zeroDrift) -gt 0.0) { throw 'Geometry drift probe: identical geometry must yield zero drift.' }
$axisDrift = Get-AutoSizeGeometryDrift -Shape $driftShape -Left 100.06 -Top 50.0 -Width 300.0 -Height 80.0
if ([Math]::Abs([double]$axisDrift - 0.06) -gt 0.000001) { throw "Geometry drift probe: single-axis drift mismatch: $axisDrift" }
$maxAxisDrift = Get-AutoSizeGeometryDrift -Shape $driftShape -Left 99.0 -Top 50.0 -Width 300.5 -Height 80.0
if ([Math]::Abs([double]$maxAxisDrift - 1.0) -gt 0.000001) { throw "Geometry drift probe: max-across-axes semantics violated: $maxAxisDrift" }

$aiImportContent = Get-Content -LiteralPath (Join-Path $root 'tools\Import-PptxAiReviewResult.ps1') -Raw -Encoding UTF8
if ($aiImportContent -match 'PowerPoint\.Application|Presentations\.Open|SaveAs|Normalize-PhysicsPpt') { throw 'AI review import must remain read-only and must not access PPTX automation.' }

& (Join-Path $root 'tools\Test-PhysicsPptPolicy.ps1')

Write-Host 'Toolkit self-check passed.'
