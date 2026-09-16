<#
.SYNOPSIS
  Orchestrate one explicit MathType/OLE → OfficeMath migration batch.

.DESCRIPTION
  Chains the OLE pipeline end to end into a versioned delivery directory
  (reports/<stem>_v<N>): carrier inventory (with optional page-kind annotation)
  -> rendered pages -> OLE crops -> GoldSet mapping -> OMML candidates -> apply
  (new copy, mc:Fallback preserved) -> Open XML validator -> visual audit. Any
  step failure stops the batch and records the completed prefix in the batch
  manifest. The source PPTX is never modified.

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
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-BatchManifestPath {
    param([string]$Root)
    return (Join-Path $Root '00_检查报告\formula-ole-batch-manifest.json')
}

function New-StepSignature {
    param([string]$StepName, [string[]]$InputPaths, [string[]]$Flags)
    $material = @($StepName) + @($Flags) + @($InputPaths | ForEach-Object { Get-FileSha256 -Path $_ })
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes(($material -join '|'))
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Test-StepReusable {
    param($Step)
    if ($null -eq $Step) { return $false }
    if ([string]$Step.status -ne 'Passed') { return $false }
    foreach ($output in @($Step.outputs)) {
        if ([string]::IsNullOrWhiteSpace([string]$output) -or -not (Test-Path -LiteralPath [string]$output)) { return $false }
    }
    return (New-StepSignature -StepName ([string]$Step.name) -InputPaths @([string[]]$Step.inputs) -Flags @([string[]]$step.flags) -eq [string]$Step.signature)
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
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($inputFullPath)
    if (-not (Test-Path -LiteralPath $reportsRoot)) { New-Item -ItemType Directory -Path $reportsRoot -Force | Out-Null }
    $maxVersion = 0
    foreach ($dir in @(Get-ChildItem -LiteralPath $reportsRoot -Directory -ErrorAction SilentlyContinue)) {
        if ($dir.Name -match ('^' + [regex]::Escape($stem) + '_v(\d+)$')) {
            $maxVersion = [Math]::Max($maxVersion, [int]$Matches[1])
        }
    }
    $deliveryFullPath = Join-Path $reportsRoot ('{0}_v{1}' -f $stem, ($maxVersion + 1))
    New-Item -ItemType Directory -Path $deliveryFullPath | Out-Null
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
    if ($null -ne $existing -and (Test-StepReusable -Step $existing)) {
        Write-Output ("[resume] step '{0}' reused (signature unchanged, outputs present)." -f $StepName)
        return
    }
    # Drop any stale recording of this step before rerunning it.
    for ($i = $batchSteps.Count - 1; $i -ge 0; $i--) {
        if ([string]$batchSteps[$i].name -eq $StepName) { $batchSteps.RemoveAt($i) }
    }
    Write-Output ("[step] {0}" -f $StepName)
    $null = & $Action
    $batchSteps.Add([pscustomobject]@{
        name = $StepName
        status = 'Passed'
        inputs = @($Inputs)
        flags = @($Flags)
        signature = New-StepSignature -StepName $StepName -InputPaths $Inputs -Flags $Flags
        outputs = @($Outputs)
        at = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }) | Out-Null
}

$inputSha256 = Get-FileSha256 -Path $inputFullPath
$goldSetSha256 = Get-FileSha256 -Path $goldSetPath
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
        Invoke-BatchStep -StepName 'before-audit' -Inputs @($inputFullPath) -Outputs @((Join-Path $beforeAuditDir 'pptx-visual-audit.csv')) -Action {
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
    Invoke-BatchStep -StepName 'crops' -Inputs @($inputFullPath, $inventoryJson, $resolvedPagesDir) -Outputs @($cropsCsv) -Action {
        & (Join-Path $PSScriptRoot 'Export-FormulaOleCrops.ps1') -CarrierInventoryJson $inventoryJson -PagesDir $resolvedPagesDir -OutputDir $cropsDir | Out-Null
    }

    # --- step 4: GoldSet mapping (page-kind exclusion lives here) ---
    $mappingCsv = Join-Path $reportDir 'formula-ole-mapping.csv'
    $reviewCsv = Join-Path $reportDir 'formula-ole-review.csv'
    Invoke-BatchStep -StepName 'mapping' -Inputs @($inventoryJson, $goldSetPath) -Outputs @($mappingCsv, $reviewCsv) -Action {
        & (Join-Path $PSScriptRoot 'Export-FormulaOleMapping.ps1') -CarrierInventoryJson $inventoryJson -GoldSetCsv $goldSetPath -OutputDir $reportDir | Out-Null
    }

    # --- step 5: OMML candidates ---
    $fragmentsDir = Join-Path $reportDir 'omml-fragments'
    $candidatesCsv = Join-Path $fragmentsDir 'formula-omml-candidates.csv'
    Invoke-BatchStep -StepName 'candidates' -Inputs @($reviewCsv, $goldSetPath) -Outputs @($candidatesCsv) -Action {
        & (Join-Path $PSScriptRoot 'Export-FormulaOmmlCandidates.ps1') -FormulaReviewCsv $reviewCsv -OutputDir $fragmentsDir | Out-Null
    }

    # --- step 6: apply to a new copy ---
    $outputName = [System.IO.Path]::GetFileName($inputFullPath)
    $outputPptx = Join-Path $deliveryDir $outputName
    $applyReport = Join-Path $reportDir 'formula-ole-apply-report.csv'
    Invoke-BatchStep -StepName 'apply' -Inputs @($inputFullPath, $mappingCsv, $candidatesCsv) -Outputs @($outputPptx, $applyReport) -Action {
        & (Join-Path $PSScriptRoot 'Apply-FormulaOmmlForOle.ps1') -InputPath $inputFullPath -MappingCsv $mappingCsv -OmmlCandidateCsv $candidatesCsv -OutputPath $outputPptx -ReportPath $applyReport | Out-Null
    }

    # --- step 7: Open XML validator ---
    $validatorJson = Join-Path $reportDir 'validator.json'
    $dotnetPath = Join-Path $env:USERPROFILE '.dotnet\dotnet.exe'
    Invoke-BatchStep -StepName 'validator' -Inputs @($outputPptx) -Outputs @($validatorJson) -Action {
        if (-not (Test-Path -LiteralPath $dotnetPath)) { throw "dotnet not found for the Open XML validator: $dotnetPath" }
        $validatorOutput = & $dotnetPath run --project (Join-Path $PSScriptRoot 'FormulaOfficeMathValidator\FormulaOfficeMathValidator.csproj') -c Release -- $outputPptx --max-errors 10 --json $validatorJson
        $validatorOutput | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "FormulaOfficeMathValidator failed with exit code $LASTEXITCODE." }
    }

    # --- step 8: after visual audit (PowerPoint open proof + pages) ---
    if (-not $SkipVisualAudit) {
        $afterAuditDir = Join-Path $reportDir 'after-audit'
        Invoke-BatchStep -StepName 'after-audit' -Inputs @($outputPptx) -Outputs @((Join-Path $afterAuditDir 'pptx-visual-audit.csv')) -Action {
            & (Join-Path $PSScriptRoot 'Export-PptxVisualAudit.ps1') -InputPath $outputPptx -OutputDir $afterAuditDir | Out-Null
        }
    }

    # --- batch manifest ---
    $applyRows = @()
    if (Test-Path -LiteralPath $applyReport) {
        $applyRows = @(Import-Csv -LiteralPath $applyReport -Encoding UTF8)
    }
    $manifest = [ordered]@{
        schemaVersion = 1
        generatedAt = Get-Date -Format 'yyyy-MM-dd hh:mm:ss'
        status = 'Passed'
        input = [ordered]@{ path = $inputFullPath; sha256 = $inputSha256 }
        goldSet = [ordered]@{ path = $goldSetPath; sha256 = $goldSetSha256 }
        deliveryRoot = $deliveryFullPath
        outputPptx = $outputPptx
        outputSha256 = Get-FileSha256 -Path $outputPptx
        applySummary = [ordered]@{
            replaced = @($applyRows | Where-Object { $_.Issue -eq 'OleFormulaReplaced' }).Count
            refused = @($applyRows | Where-Object { $_.Issue -notin @('OleFormulaReplaced', 'SavedAs') }).Count
        }
        steps = $batchSteps.ToArray()
        writeBackAllowed = $true
        note = 'Explicit migration batch on a new copy. Source PPTX untouched; mc:Fallback keeps original OLE objects.'
    }
    Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 10) -Path $manifestPath
    Write-Output ("Batch passed: {0}" -f $deliveryFullPath)
    Write-Output ("Output: {0} (replaced={1}; refused={2})" -f $outputPptx, $manifest.applySummary.replaced.Value, $manifest.applySummary.refused.Value)
} catch {
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
    Write-Output ("Batch FAILED at: {0}" -f $_.Exception.Message)
    throw
}
