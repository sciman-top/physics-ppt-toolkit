<#
.SYNOPSIS
  Apply verified image enhancement probe outputs to PPTX copies.

.DESCRIPTION
  Reads pptx-image-enhancement-probe.csv and replaces only verified JPEG media
  entries whose delivery image is same-dimension and within the configured size
  multiplier. The script writes new PPTX files and reports; it never overwrites
  input PPTX files.

.PARAMETER InputPath
  Path to a PPTX file or a directory containing PPTX files.

.PARAMETER ProbeCsv
  CSV produced by Invoke-PptxImageEnhancementProbe.ps1.

.PARAMETER OutputDir
  Directory where image-enhanced PPTX copies and reports are written.

.PARAMETER FilePattern
  File name pattern used when InputPath is a directory.

.PARAMETER MaxSizeMultiplier
  Maximum replacement/original byte ratio. Default 1.0 only accepts smaller or
  same-size replacement files.

.PARAMETER RequireSameDimensions
  Require replacement width/height to match the original media. Enabled by
  default because this avoids layout and crop drift in PPTX.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$ProbeCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [switch]$Recurse,

    [string]$FilePattern = '*.pptx',

    [ValidateRange(0.1, 20)]
    [double]$MaxSizeMultiplier = 1.0,

    [bool]$RequireSameDimensions = $true,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-PptxFiles {
    param([string]$Path, [string]$Pattern, [switch]$Recurse)
    if (-not (Test-Path -LiteralPath $Path)) { throw "InputPath not found: $Path" }
    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) {
        $opt = @{ LiteralPath = $item.FullName; Filter = $Pattern; File = $true }
        if ($Recurse) { $opt.Recurse = $true }
        return @(Get-ChildItem @opt | Where-Object { $_.Name -notlike '~$*' })
    }
    if ($item.Extension -ne '.pptx') { throw "Only .pptx files are supported: $($item.FullName)" }
    return @($item)
}

function Get-BasicImageInfo {
    param([string]$Path)
    Add-Type -AssemblyName System.Drawing
    $image = $null
    try {
        $image = [System.Drawing.Image]::FromFile($Path)
        return [pscustomobject]@{
            Width = [int]$image.Width
            Height = [int]$image.Height
            Bytes = [int64](Get-Item -LiteralPath $Path).Length
        }
    } finally {
        if ($null -ne $image) { $image.Dispose() }
    }
}

function Get-ReplacementPath {
    param([object]$Row)

    foreach ($field in @('DeliveryPath', 'BestOutput', 'CompressedPath', 'EnhancedPath')) {
        $prop = $Row.PSObject.Properties[$field]
        if ($null -ne $prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value) -and (Test-Path -LiteralPath ([string]$prop.Value))) {
            return [string]$prop.Value
        }
    }
    return ''
}

function Add-ReportRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Deck,
        [string]$MediaPath,
        [string]$Action,
        [string]$Reason,
        [int64]$OriginalBytes,
        [int64]$ReplacementBytes,
        [string]$ReplacementPath
    )

    $savedBytes = $OriginalBytes - $ReplacementBytes
    $savedPercent = if ($OriginalBytes -gt 0) { [Math]::Round(($savedBytes / [double]$OriginalBytes) * 100, 2) } else { 0 }
    $Rows.Add([pscustomobject]@{
        Deck = $Deck
        MediaPath = $MediaPath
        Action = $Action
        Reason = $Reason
        OriginalBytes = $OriginalBytes
        ReplacementBytes = $ReplacementBytes
        SavedBytes = $savedBytes
        SavedPercent = $savedPercent
        ReplacementPath = $ReplacementPath
    }) | Out-Null
}

function Apply-ReplacementsToPresentation {
    param(
        [System.IO.FileInfo]$File,
        [object[]]$Rows,
        [string]$OutputDir,
        [double]$MaxSizeMultiplier,
        [bool]$RequireSameDimensions,
        [bool]$Force
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $safeBase = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $outPptx = Join-Path $OutputDir ($safeBase + '.image-enhanced.pptx')
    $reportPath = Join-Path $OutputDir ($safeBase + '.image-enhancement-apply-report.csv')
    $manifestPath = Join-Path $OutputDir ($safeBase + '.image-enhancement-apply.json')
    if ((Test-Path -LiteralPath $outPptx) -and -not $Force) {
        Write-Host "Skip existing: $outPptx"
        return
    }

    $workRoot = Join-Path $OutputDir ('_image_enhance_work_' + [Guid]::NewGuid().ToString('N'))
    $reportRows = New-Object System.Collections.Generic.List[object]
    $status = 'success'
    $message = ''

    try {
        New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($File.FullName, $workRoot)
        $videoPosterMap = Get-VideoPosterImageMap -PptxFile $File

        foreach ($row in $Rows) {
            $mediaPath = Resolve-PackagePath -PackagePath ([string]$row.MediaPath)
            $replacementPath = Get-ReplacementPath -Row $row
            if ([string]::IsNullOrWhiteSpace($mediaPath) -or -not $mediaPath.StartsWith('ppt/media/', [System.StringComparison]::OrdinalIgnoreCase)) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath ([string]$row.MediaPath) -Action 'Skipped' -Reason 'Invalid or non-media PPTX package path' -OriginalBytes 0 -ReplacementBytes 0 -ReplacementPath $replacementPath
                continue
            }

            $targetPath = Resolve-ExtractedPackageFilePath -ExtractionRoot $workRoot -PackagePath $mediaPath
            if ([string]::IsNullOrWhiteSpace($targetPath)) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Skipped' -Reason 'Media path resolves outside PPTX work directory' -OriginalBytes 0 -ReplacementBytes 0 -ReplacementPath $replacementPath
                continue
            }

            if ($videoPosterMap.ContainsKey($mediaPath)) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Skipped' -Reason 'Video poster frame is not a standalone image' -OriginalBytes 0 -ReplacementBytes 0 -ReplacementPath $replacementPath
                continue
            }
            if ([string]::IsNullOrWhiteSpace($replacementPath)) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Skipped' -Reason 'Replacement file not found' -OriginalBytes 0 -ReplacementBytes 0 -ReplacementPath ''
                continue
            }

            $originalExt = [System.IO.Path]::GetExtension($targetPath).ToLowerInvariant()
            $replacementExt = [System.IO.Path]::GetExtension($replacementPath).ToLowerInvariant()
            if ($originalExt -notin @('.jpg', '.jpeg') -or $replacementExt -notin @('.jpg', '.jpeg')) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Skipped' -Reason 'Only JPEG-to-JPEG replacement is enabled' -OriginalBytes 0 -ReplacementBytes 0 -ReplacementPath $replacementPath
                continue
            }

            if (-not (Test-Path -LiteralPath $targetPath)) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Skipped' -Reason 'Media path not found in PPTX' -OriginalBytes 0 -ReplacementBytes 0 -ReplacementPath $replacementPath
                continue
            }

            $originalInfo = Get-BasicImageInfo -Path $targetPath
            $replacementInfo = Get-BasicImageInfo -Path $replacementPath
            $ratio = if ($originalInfo.Bytes -gt 0) { $replacementInfo.Bytes / [double]$originalInfo.Bytes } else { 999 }

            if ($RequireSameDimensions -and ($originalInfo.Width -ne $replacementInfo.Width -or $originalInfo.Height -ne $replacementInfo.Height)) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Skipped' -Reason 'Replacement dimensions differ from original' -OriginalBytes $originalInfo.Bytes -ReplacementBytes $replacementInfo.Bytes -ReplacementPath $replacementPath
                continue
            }

            if ($ratio -gt $MaxSizeMultiplier) {
                Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Skipped' -Reason "Replacement exceeds MaxSizeMultiplier $MaxSizeMultiplier" -OriginalBytes $originalInfo.Bytes -ReplacementBytes $replacementInfo.Bytes -ReplacementPath $replacementPath
                continue
            }

            Copy-Item -LiteralPath $replacementPath -Destination $targetPath -Force
            Add-ReportRow -Rows $reportRows -Deck $File.Name -MediaPath $mediaPath -Action 'Replaced' -Reason 'Verified delivery image applied' -OriginalBytes $originalInfo.Bytes -ReplacementBytes $replacementInfo.Bytes -ReplacementPath $replacementPath
        }

        New-PptxPackageFromDirectory -SourceDir $workRoot -DestinationPath $outPptx
    } catch {
        $status = 'failed'
        $message = $_.Exception.Message
    } finally {
        if (Test-Path -LiteralPath $workRoot) {
            $resolvedWork = [System.IO.Path]::GetFullPath($workRoot)
            $resolvedOutput = [System.IO.Path]::GetFullPath($OutputDir)
            if (Test-PathInsideDirectory -ChildPath $resolvedWork -ParentPath $resolvedOutput) {
                Remove-Item -LiteralPath $workRoot -Recurse -Force
            }
        }
    }

    $reportRows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
    $replacedRows = @($reportRows | Where-Object { $_.Action -eq 'Replaced' })
    $savedBytes = 0
    if ($replacedRows.Count -gt 0) {
        $savedMeasure = $replacedRows | Measure-Object -Property SavedBytes -Sum
        if ($null -ne $savedMeasure) { $savedBytes = [int64]$savedMeasure.Sum }
    }
    $manifest = [pscustomobject]@{
        generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        input = $File.FullName
        outputPptx = $outPptx
        report = $reportPath
        status = $status
        message = $message
        replacementCandidateCount = $Rows.Count
        replacedCount = $replacedRows.Count
        skippedCount = $reportRows.Count - $replacedRows.Count
        savedBytes = $savedBytes
        maxSizeMultiplier = $MaxSizeMultiplier
        requireSameDimensions = $RequireSameDimensions
    }
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$ProbeCsv = [System.IO.Path]::GetFullPath($ProbeCsv)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $ProbeCsv)) { throw "ProbeCsv not found: $ProbeCsv" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$probeRows = @(Import-Csv -LiteralPath $ProbeCsv -Encoding UTF8 | Where-Object { $_.Status -eq 'Success' })
if ($probeRows.Count -eq 0) { throw 'ProbeCsv contains no successful rows.' }

$rowsByDeck = @{}
foreach ($row in $probeRows) {
    $deck = [string]$row.Deck
    if (-not $rowsByDeck.ContainsKey($deck)) {
        $rowsByDeck[$deck] = New-Object System.Collections.Generic.List[object]
    }
    $rowsByDeck[$deck].Add($row) | Out-Null
}

$files = @(Get-PptxFiles -Path $InputPath -Pattern $FilePattern -Recurse:$Recurse)
foreach ($file in $files) {
    if (-not $rowsByDeck.ContainsKey($file.Name)) { continue }
    Write-Host "Applying image enhancements: $($file.Name)"
    Apply-ReplacementsToPresentation -File $file -Rows $rowsByDeck[$file.Name].ToArray() -OutputDir $OutputDir -MaxSizeMultiplier $MaxSizeMultiplier -RequireSameDimensions $RequireSameDimensions -Force ([bool]$Force)
}

Write-Host "Image enhancement apply done: $OutputDir"
