<#
.SYNOPSIS
  Run a local OCR probe on formula-like PPTX image candidates.

.DESCRIPTION
  Reads formula-image-candidates.csv, runs the bundled/free RapidOCR runtime
  through Python, and writes review reports. This script is read-only: it never
  modifies PPTX files and never applies OCR output automatically.

.PARAMETER FormulaImageCandidateCsv
  formula-image-candidates.csv from Export-FormulaImageCandidates.ps1.

.PARAMETER OutputDir
  Directory where OCR reports and previews are written.

.PARAMETER PythonPath
  Python executable. Defaults to the bundled Codex Python when available.

.PARAMETER MaxImages
  Maximum candidate images to OCR in one run.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FormulaImageCandidateCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [ValidateRange(1, 500)]
    [int]$MaxImages = 30,

    [ValidateRange(0, 100)]
    [int]$MinScore = 45,

    [ValidateRange(10000, [int]::MaxValue)]
    [int]$MaxPixels = 1500000,

    [string]$PythonPath = '',

    [switch]$OcrOriginal,

    [switch]$NoContactSheet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-PythonPath {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        $full = [System.IO.Path]::GetFullPath($RequestedPath)
        if (-not (Test-Path -LiteralPath $full)) { throw "PythonPath not found: $full" }
        return $full
    }

    $bundled = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $bundled) { return $bundled }

    $cmd = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source)) { return $cmd.Source }

    throw 'Python was not found. Install Python or pass -PythonPath.'
}

$FormulaImageCandidateCsv = [System.IO.Path]::GetFullPath($FormulaImageCandidateCsv)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
$PythonPath = Resolve-PythonPath -RequestedPath $PythonPath
$helper = Join-Path $PSScriptRoot 'formula_image_ocr_probe.py'

if (-not (Test-Path -LiteralPath $FormulaImageCandidateCsv)) { throw "FormulaImageCandidateCsv not found: $FormulaImageCandidateCsv" }
if (-not (Test-Path -LiteralPath $helper)) { throw "OCR helper not found: $helper" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$args = @(
    $helper,
    '--input-csv', $FormulaImageCandidateCsv,
    '--output-dir', $OutputDir,
    '--max-images', [string]$MaxImages,
    '--min-score', [string]$MinScore,
    '--max-pixels', [string]$MaxPixels
)

if ($OcrOriginal) {
    $args += '--ocr-original'
}

if ($NoContactSheet) {
    $args += '--no-contact-sheet'
}

Write-Host "Using Python: $PythonPath"
& $PythonPath @args
if ($LASTEXITCODE -ne 0) {
    throw "Formula image OCR probe failed with exit code $LASTEXITCODE"
}
