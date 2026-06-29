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
    'config\physics-ppt-style.config.json',
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

$requiredSections = @('fonts', 'fontSizes', 'colors', 'formulaWhitelist', 'rules')
foreach ($section in $requiredSections) {
    if ($null -eq $config.$section) { throw "Config missing required section: $section" }
}

$requiredFontFields = @('chinese', 'latin', 'math')
foreach ($field in $requiredFontFields) {
    if ([string]::IsNullOrWhiteSpace($config.fonts.$field)) { throw "Config fonts.$field is missing or empty" }
}

$requiredSizeFields = @('title1', 'body', 'minimum', 'tableHeader', 'tableBody', 'formulaInline', 'formulaStandalone', 'formulaCore')
foreach ($field in $requiredSizeFields) {
    if ($null -eq $config.fontSizes.$field) { throw "Config fontSizes.$field is missing" }
}

# Validate color values are hex strings
$requiredColorFields = @('white', 'black', 'body', 'emphasisRed', 'formulaBlue', 'yellowFill', 'yellowBorder')
foreach ($field in $requiredColorFields) {
    $val = $config.colors.$field
    if ([string]::IsNullOrWhiteSpace($val)) { throw "Config colors.$field is missing or empty" }
    if ($val -notmatch '^#[0-9A-Fa-f]{6}$') { throw "Config colors.$field is not a valid hex color: $val" }
}

if ($null -eq $config.rules.formulaTextStyleDefault) {
    throw "Config rules.formulaTextStyleDefault is missing"
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

Write-Host 'Toolkit self-check passed.'

# --- 7. JSON <-> VBA constants sync check ---
$vbaCommonPath = Join-Path $root 'vba\PhysicsPptCommon.bas'
$vbaContent = Get-Content -LiteralPath $vbaCommonPath -Raw -Encoding UTF8

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

Write-Host 'JSON <-> VBA sync check passed.'
