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
    'config\formula-ir.schema.json',
    'config\formula-evidence-manifest.schema.json',
    'config\formula-goldset.schema.json',
    'config\formula-recognition-result.schema.json',
    'config\formula-recognition.adapters.json',
    'config\presentation-snapshot.schema.json',
    'config\invariant-comparison.schema.json',
    'tools\Normalize-PhysicsPpt.ps1',
    'tools\Apply-FormulaSvgWhitelist.ps1',
    'tools\Export-FormulaOmmlCandidates.ps1',
    'tools\Export-FormulaIrFromOleMapping.ps1',
    'tools\Test-FormulaIr.ps1',
    'tools\Apply-FormulaOmmlWhitelist.ps1',
    'tools\Apply-FormulaOmmlForOle.ps1',
    'tools\Set-PptxTextBold.ps1',
    'tools\Export-FormulaGoldSet.ps1',
    'tools\Run-FormulaRecognitionAdapter.ps1',
    'tools\Invoke-FormulaOleBatch.ps1',
    'tools\Export-FormulaWhitelistSuggestions.ps1',
    'tools\Export-FormulaImageCandidates.ps1',
    'tools\Export-FormulaImageCrops.ps1',
    'tools\Compare-FormulaConverters.ps1',
    'tools\Export-FormulaCarrierInventory.ps1',
    'tools\Export-FormulaOleMapping.ps1',
    'tools\Export-FormulaOleCrops.ps1',
    'tools\Export-FormulaEvidenceManifest.ps1',
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

$formulaIrSchemaPath = Join-Path $root 'config\formula-ir.schema.json'
$formulaIrSchema = Get-Content -LiteralPath $formulaIrSchemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$formulaIrSchema.title -ne 'Physics PPT FormulaIR' -or [int]$formulaIrSchema.properties.schemaVersion.const -ne 1) {
    throw 'FormulaIR schema metadata is invalid.'
}
$evidenceSchemaPath = Join-Path $root 'config\formula-evidence-manifest.schema.json'
$evidenceSchema = Get-Content -LiteralPath $evidenceSchemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$evidenceSchema.title -ne 'Physics PPT formula evidence manifest' -or
    [int]$evidenceSchema.properties.schemaVersion.const -ne 1) {
    throw 'Formula evidence manifest schema metadata is invalid.'
}
foreach ($evidenceRequired in @('schemaVersion', 'generatedAt', 'policy', 'input', 'inventory', 'counts', 'records')) {
    if ($evidenceRequired -notin @($evidenceSchema.required)) {
        throw "Formula evidence manifest schema is missing required root field: $evidenceRequired"
    }
}
$evidenceManifestScript = Get-Content -LiteralPath (Join-Path $root 'tools\Export-FormulaEvidenceManifest.ps1') -Raw -Encoding UTF8
foreach ($evidenceToken in @('formula-evidence-manifest.json', 'pathBase', 'writeBackAllowed = $false', 'PptxPackageEntry', 'evidenceSetSha256')) {
    if ($evidenceManifestScript -notmatch [regex]::Escape($evidenceToken)) {
        throw "Formula evidence manifest tool is missing required contract token: $evidenceToken"
    }
}
$cropScript = Get-Content -LiteralPath (Join-Path $root 'tools\Export-FormulaImageCrops.ps1') -Raw -Encoding UTF8
foreach ($cropToken in @('PaddingPx', 'transparentTrimOnly', 'MixedImage', 'SplitFailed', 'originalPreservedOnFailure')) {
    if ($cropScript -notmatch [regex]::Escape($cropToken)) {
        throw "Formula image crop protocol is missing required token: $cropToken"
    }
}
$converterScript = Get-Content -LiteralPath (Join-Path $root 'tools\Compare-FormulaConverters.ps1') -Raw -Encoding UTF8
foreach ($converterToken in @('Pandoc', 'ReferenceOnly', 'formula-converter-comparison-manifest.json', 'writesPptx = $false')) {
    if ($converterScript -notmatch [regex]::Escape($converterToken)) {
        throw "Formula converter comparison harness is missing required token: $converterToken"
    }
}
$ommlCandidateScript = Get-Content -LiteralPath (Join-Path $root 'tools\Export-FormulaOmmlCandidates.ps1') -Raw -Encoding UTF8
foreach ($formulaIrToken in @('Convert-AstToFormulaIrToken', 'formulaIrDir', 'FormulaIrStatus', 'CandidateOnly')) {
    if ($ommlCandidateScript -notmatch [regex]::Escape($formulaIrToken)) {
        throw "OMML candidate exporter is missing FormulaIR token contract: $formulaIrToken"
    }
}
$oleIrScript = Get-Content -LiteralPath (Join-Path $root 'tools\Export-FormulaIrFromOleMapping.ps1') -Raw -Encoding UTF8
foreach ($oleIrToken in @('ApprovedMathTypeOleGoldSet', 'CandidateOnly', 'writeBackAllowed = $false', 'Export-FormulaOmmlCandidates.ps1')) {
    if ($oleIrScript -notmatch [regex]::Escape($oleIrToken)) {
        throw "OLE FormulaIR bridge is missing required token: $oleIrToken"
    }
}
$workflowScript = Get-Content -LiteralPath (Join-Path $root 'tools\Invoke-PhysicsPptWorkflow.ps1') -Raw -Encoding UTF8
foreach ($workflowToken in @('FormulaProcessingMode', 'ExplicitMigration', 'ClosedWorldUnattended', 'actualWriteBack', 'writeBackRequested')) {
    if ($workflowScript -notmatch [regex]::Escape($workflowToken)) {
        throw "Workflow formula mode contract is missing required token: $workflowToken"
    }
}
$inventoryScript = Get-Content -LiteralPath (Join-Path $root 'tools\Export-FormulaCarrierInventory.ps1') -Raw -Encoding UTF8
foreach ($inventoryToken in @('OfficeMath', 'MathTypeOle', 'TextFormula', 'FormulaImage', 'MixedImage', 'GroupFormula', 'Unknown', 'formula-carrier-inventory.json', 'ReviewStatus=Approved', 'formula-ole-mapping.csv')) {
    if ($inventoryScript -notmatch [regex]::Escape($inventoryToken)) {
        throw "Formula carrier inventory is missing required carrier/manifest token: $inventoryToken"
    }
}
$oleCropsScript = Get-Content -LiteralPath (Join-Path $root 'tools\Export-FormulaOleCrops.ps1') -Raw -Encoding UTF8
foreach ($oleCropsToken in @('sldSz', 'SuggestedSizePt', 'writeBackAllowed = $false', 'ole-crops.csv', 'formula-ole-crops-manifest.json', 'PaddingPx', 'PageSha256')) {
    if ($oleCropsScript -notmatch [regex]::Escape($oleCropsToken)) {
        throw "OLE crop exporter is missing required evidence token: $oleCropsToken"
    }
}
foreach ($schemaRequired in @('schemaVersion', 'recordId', 'source', 'detection', 'canonical', 'decision', 'evidence')) {
    if ($schemaRequired -notin @($formulaIrSchema.required)) {
        throw "FormulaIR schema is missing required root field: $schemaRequired"
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
    'Parse-FormulaSequence', 'Parse-FormulaExpression', 'Parse-FormulaAst',
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

# --- 2c. Formula recognition gold set / adapter transport contract ---
$goldSetSchemaPath = Join-Path $root 'config\formula-goldset.schema.json'
$goldSetSchema = Get-Content -LiteralPath $goldSetSchemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$goldSetSchema.title -ne 'Physics PPT formula recognition gold set' -or
    [int]$goldSetSchema.properties.schemaVersion.const -ne 1) {
    throw 'Formula gold set schema metadata is invalid.'
}
foreach ($goldSetRequired in @('schemaVersion', 'generatedAt', 'input', 'policy', 'counts', 'evidenceSetSha256', 'records')) {
    if ($goldSetRequired -notin @($goldSetSchema.required)) { throw "Formula gold set schema is missing required root field: $goldSetRequired" }
}
if ([bool]$goldSetSchema.properties.policy.properties.writeBackAllowed.const) { throw 'Formula gold set schema must keep writeBackAllowed=false.' }
if ([string]$goldSetSchema.properties.policy.properties.reviewRequirement.const -ne 'HumanReviewed') { throw 'Formula gold set schema must require HumanReviewed review.' }

$recognitionResultSchemaPath = Join-Path $root 'config\formula-recognition-result.schema.json'
$recognitionResultSchema = Get-Content -LiteralPath $recognitionResultSchemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$recognitionResultSchema.title -ne 'Physics PPT formula recognition adapter result' -or
    [int]$recognitionResultSchema.properties.schemaVersion.const -ne 1) {
    throw 'Formula recognition result schema metadata is invalid.'
}
foreach ($statusValue in @('Passed', 'Unavailable', 'Failed', 'InvalidOutput', 'Timeout')) {
    if ($statusValue -notin @($recognitionResultSchema.properties.status.enum)) { throw "Formula recognition result schema status enum is missing: $statusValue" }
}
if ([bool]$recognitionResultSchema.properties.writeBackAllowed.const) { throw 'Formula recognition result schema must keep writeBackAllowed=false.' }

$adapterConfigPath = Join-Path $root 'config\formula-recognition.adapters.json'
$adapterConfig = Get-Content -LiteralPath $adapterConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$adapterConfig.schemaVersion -ne 1) { throw 'Formula recognition adapter config schemaVersion must be 1.' }
if ([string]$adapterConfig.defaultAdapter -ne 'None') { throw 'Formula recognition adapter config must keep defaultAdapter=None.' }
if ([bool]$adapterConfig.writeBackAllowed) { throw 'Formula recognition adapter config must keep writeBackAllowed=false.' }
if (@($adapterConfig.adapters).Count -lt 1) { throw 'Formula recognition adapter config must declare at least one adapter entry.' }
foreach ($adapterEntry in @($adapterConfig.adapters)) {
    foreach ($adapterField in @('id', 'status', 'entrypoint', 'outputContract', 'licenseStatus', 'disabledReason')) {
        $adapterFieldProperty = $adapterEntry.PSObject.Properties[$adapterField]
        if ($null -eq $adapterFieldProperty -or [string]::IsNullOrWhiteSpace([string]$adapterFieldProperty.Value)) {
            throw "Formula recognition adapter '$($adapterEntry.id)' is missing field: $adapterField"
        }
    }
    if ([string]$adapterEntry.outputContract -ne 'formula-recognition-result.schema.json') {
        throw "Formula recognition adapter '$($adapterEntry.id)' must bind the result schema contract."
    }
    if ([string]$adapterEntry.licenseStatus -ne 'VerifyBeforeInstall') {
        throw "Formula recognition adapter '$($adapterEntry.id)' must keep licenseStatus=VerifyBeforeInstall until provisioned."
    }
}

$goldSetScript = Get-Content -LiteralPath (Join-Path $root 'tools\Export-FormulaGoldSet.ps1') -Raw -Encoding UTF8
foreach ($goldSetToken in @('writeBackAllowed = $false', 'HumanReviewed', 'MixedNonIsolatable', 'formula-goldset-manifest.json', 'evidenceSetSha256')) {
    if ($goldSetScript -notmatch [regex]::Escape($goldSetToken)) {
        throw "Formula gold set exporter is missing required contract token: $goldSetToken"
    }
}
$adapterTransportScript = Get-Content -LiteralPath (Join-Path $root 'tools\Run-FormulaRecognitionAdapter.ps1') -Raw -Encoding UTF8
foreach ($adapterTransportToken in @('Unavailable', 'InvalidOutput', 'Timeout', 'writeBackAllowed', 'RunnerArgumentList', 'schemaVersion')) {
    if ($adapterTransportScript -notmatch [regex]::Escape($adapterTransportToken)) {
        throw "Formula recognition adapter transport is missing required contract token: $adapterTransportToken"
    }
}
$oleBatchScript = Get-Content -LiteralPath (Join-Path $root 'tools\Invoke-FormulaOleBatch.ps1') -Raw -Encoding UTF8
foreach ($oleBatchToken in @('formula-ole-batch-manifest.json', 'Resume', 'Apply-FormulaOmmlForOle', 'FormulaOfficeMathValidator', 'Export-FormulaCarrierInventory')) {
    if ($oleBatchScript -notmatch [regex]::Escape($oleBatchToken)) {
        throw "OLE batch orchestrator is missing required contract token: $oleBatchToken"
    }
}
$goldSetSampleRows = @(Import-Csv -LiteralPath (Join-Path $root 'examples\fixtures\formula-goldset.sample.csv'))
if ($goldSetSampleRows.Count -lt 1) { throw 'Formula gold set sample fixture must contain at least one row.' }
$goldSetSampleHeader = @($goldSetSampleRows[0].PSObject.Properties.Name)
foreach ($goldSetColumn in @('ReviewStatus', 'Slide', 'ShapeIds', 'WhitelistName', 'SourceFormulaText', 'SizePt', 'MainColorHex', 'SubColorHex', 'Note', 'EvidencePath')) {
    if ($goldSetColumn -notin $goldSetSampleHeader) { throw "Formula gold set sample fixture is missing column: $goldSetColumn" }
}
foreach ($goldSetSampleRow in $goldSetSampleRows) {
    if ([string]$goldSetSampleRow.ReviewStatus -notin @('Draft', 'Approved')) { throw 'Formula gold set sample fixture has invalid ReviewStatus.' }
}

# Exercise the adapter transport boundary end to end: no runner, missing
# runner, failed runner, invalid runner output, slow runner and a trusted
# runner must each produce a deterministic, non-write-back JSON result.
$adapterToolPath = Join-Path $root 'tools\Run-FormulaRecognitionAdapter.ps1'
$adapterProbeDir = Join-Path ([IO.Path]::GetTempPath()) ("physics-ppt-adapter-probe-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $adapterProbeDir -Force | Out-Null
$adapterProbeImage = Join-Path $adapterProbeDir 'probe.png'
[IO.File]::WriteAllBytes($adapterProbeImage, [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A))
$runnerHost = (Get-Process -Id $PID).Path
$fixtureRunnerScripts = @{
    'passed.ps1'  = "`$null = [Console]::In.ReadToEnd()`n`$result = [ordered]@{ status = 'Passed'; candidates = @(@{ raw = 'W_{\text{总}}=Fs'; score = 0.93; status = 'Candidate' }); diagnostics = [ordered]@{ reason = 'fixture runner completed' } }`n`$result | ConvertTo-Json -Depth 6"
    'invalid.ps1' = "`$null = [Console]::In.ReadToEnd()`nWrite-Output 'definitely not json'"
    'failed.ps1'  = "`$null = [Console]::In.ReadToEnd()`nexit 4"
    'slow.ps1'    = "`$null = [Console]::In.ReadToEnd()`nStart-Sleep -Seconds 30"
}
try {
    foreach ($fixtureName in $fixtureRunnerScripts.Keys) {
        [IO.File]::WriteAllText((Join-Path $adapterProbeDir $fixtureName), $fixtureRunnerScripts[$fixtureName], (New-Object Text.UTF8Encoding($false)))
    }
    $expectedInputSha256 = (Get-FileHash -LiteralPath $adapterProbeImage -Algorithm SHA256).Hash.ToLowerInvariant()
    $adapterProbes = @(
        @{ Name = 'no-runner';        Args = @{};                                                                                        Expected = 'Unavailable' },
        @{ Name = 'missing-runner';   Args = @{ RunnerPath = (Join-Path $adapterProbeDir 'does-not-exist.ps1') };                        Expected = 'Unavailable' },
        @{ Name = 'failed-runner';    Args = @{ RunnerPath = $runnerHost; RunnerArgumentList = @('-NoProfile', '-NonInteractive', '-File', (Join-Path $adapterProbeDir 'failed.ps1')) }; Expected = 'Failed' },
        @{ Name = 'invalid-runner';   Args = @{ RunnerPath = $runnerHost; RunnerArgumentList = @('-NoProfile', '-NonInteractive', '-File', (Join-Path $adapterProbeDir 'invalid.ps1')) }; Expected = 'InvalidOutput' },
        @{ Name = 'timeout-runner';   Args = @{ RunnerPath = $runnerHost; RunnerArgumentList = @('-NoProfile', '-NonInteractive', '-File', (Join-Path $adapterProbeDir 'slow.ps1')); TimeoutSeconds = 1 }; Expected = 'Timeout' },
        @{ Name = 'trusted-runner';   Args = @{ RunnerPath = $runnerHost; RunnerArgumentList = @('-NoProfile', '-NonInteractive', '-File', (Join-Path $adapterProbeDir 'passed.ps1')) }; Expected = 'Passed' }
    )
    foreach ($probe in $adapterProbes) {
        $probeOutput = Join-Path $adapterProbeDir ("result-" + $probe.Name + ".json")
        $probeParams = @{ Adapter = 'PP-FormulaNet_plus-M'; ImagePath = $adapterProbeImage; OutputPath = $probeOutput }
        foreach ($probeArg in $probe.Args.Keys) { $probeParams[$probeArg] = $probe.Args[$probeArg] }
        & $adapterToolPath @probeParams | Out-Null
        $probeResult = Get-Content -LiteralPath $probeOutput -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]$probeResult.status -ne $probe.Expected) {
            throw ("Adapter transport probe '{0}' returned status '{1}' instead of '{2}' (reason: {3})" -f $probe.Name, $probeResult.status, $probe.Expected, $probeResult.diagnostics.reason)
        }
        if ([bool]$probeResult.writeBackAllowed) { throw "Adapter transport probe '$($probe.Name)' enabled write-back." }
        if ([string]$probeResult.input.sha256 -ne $expectedInputSha256) { throw "Adapter transport probe '$($probe.Name)' lost the input hash binding." }
        if ($probe.Name -eq 'trusted-runner') {
            if (@($probeResult.candidates).Count -ne 1 -or [string]@($probeResult.candidates)[0].status -ne 'Candidate') {
                throw "Adapter transport probe 'trusted-runner' did not carry runner candidates."
            }
        }
    }
} finally {
    if (Test-Path -LiteralPath $adapterProbeDir) { Remove-Item -LiteralPath $adapterProbeDir -Recurse -Force }
}

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
    'tools\Apply-PptxHighlightBoxStyle.ps1',
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
foreach ($requiredRuleMarker in @('Test-StyleRuleEnabled', 'STYLE.FORMULA.TEXT', 'STYLE.DECORATIVE.EFFECTS', 'SLIDE.BACKGROUND', 'DisableAdvanceOnClick', 'AdvanceOnClickPreserved', 'SizeBodyMax', 'BodyFontSizeCapped', 'SizeDisplayTitleMin', 'TextStyleNormalizedGeometryRestored')) {
    if ($normalizeContent -notmatch [regex]::Escape($requiredRuleMarker)) { throw "Normalize rule marker is missing: $requiredRuleMarker" }
}
if ($normalizeContent -match 'Normalize-TextShape[^\r\n]*-FontOnly') { throw 'Special slides must stay report-only; preserve slides must not receive font-only normalization.' }
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
foreach ($requiredWorkflowMarker in @('BlockedMissingNormalizedPptx', 'deliveryBlocked', 'deliveryStatus', 'invariantDeliveryBlocked', "'Ready'", "'BlockedAiVisualReview'", 'New-VersionedDeliveryRoot', 'Get-ExistingPreparedAiEvidence', 'VersionedDelivery')) {
    if ($workflowContent -notmatch [regex]::Escape($requiredWorkflowMarker)) { throw "Workflow delivery gate marker is missing: $requiredWorkflowMarker" }
}
foreach ($requiredAiReuseMarker in @('Test-PreparedManifestMatchesCurrent', 'Test-PreparedPacketMatchesCurrent', 'preparedManifestMatchesCurrent', 'preparedPacketMatchesCurrent', 'existing prepared AI packet does not match')) {
    if ($workflowContent -notmatch [regex]::Escape($requiredAiReuseMarker)) { throw "Workflow AI-reuse integrity marker is missing: $requiredAiReuseMarker" }
}
foreach ($requiredArtifactMarker in @('inputSha256', 'normalizedPptxSha256', 'pdfSha256', 'Test-PageImageSet', 'Test-UsablePageImage')) {
    if ($workflowContent -notmatch [regex]::Escape($requiredArtifactMarker)) { throw "Workflow artifact-integrity marker is missing: $requiredArtifactMarker" }
}
$visualAuditContent = Get-Content -LiteralPath (Join-Path $root 'tools\Export-PptxVisualAudit.ps1') -Raw -Encoding UTF8
foreach ($requiredVisualExportMarker in @('did not create a non-empty PDF', 'PdfExportedBySaveAsFallback', 'Get-ImageVisualMetrics', 'SlidePngExportFailed')) {
    if ($visualAuditContent -notmatch [regex]::Escape($requiredVisualExportMarker)) { throw "Visual export-integrity marker is missing: $requiredVisualExportMarker" }
}
foreach ($requiredPngDecodeMarker in @('Get-BasicImageInfo -Path $target', 'undecodable PNG')) {
    if ($normalizeContent -notmatch [regex]::Escape($requiredPngDecodeMarker)) { throw "Normalize PNG-integrity marker is missing: $requiredPngDecodeMarker" }
    if ($workflowContent -notmatch [regex]::Escape($requiredPngDecodeMarker)) { throw "Workflow PNG-integrity marker is missing: $requiredPngDecodeMarker" }
}
foreach ($requiredAiDeliveryMarker in @('deliveryStatus', "'Ready'", 'preparedManifestSha256', '交付状态')) {
    if ($aiImportContent -notmatch [regex]::Escape($requiredAiDeliveryMarker)) { throw "AI import delivery marker is missing: $requiredAiDeliveryMarker" }
}
foreach ($summaryCheckMarker in @('configuredFontsMissingCount', 'slideAspectMismatchCount', '字体回退风险', '非 16:9')) {
    if ($workflowContent -notmatch [regex]::Escape($summaryCheckMarker)) { throw "Summary preflight marker is missing: $summaryCheckMarker" }
}

& (Join-Path $root 'tools\Test-PhysicsPptPolicy.ps1')

Write-Host 'Toolkit self-check passed.'
