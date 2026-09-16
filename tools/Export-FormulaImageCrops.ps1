<#
.SYNOPSIS
  Export deterministic, evidence-bound formula image crops.

.DESCRIPTION
  Consumes formula-image-candidates.csv and creates review crops without
  changing a PPTX.  Standalone FormulaImage rows may use the full image when
  no rectangle is supplied.  MixedImage rows require an explicit crop
  rectangle; otherwise they receive SplitFailed and no crop is written.
  Transparent margins are trimmed only when the source actually contains
  transparency.  Opaque images are never content-trimmed heuristically.

.PARAMETER FormulaImageCandidateCsv
  Output of Export-FormulaImageCandidates.ps1.

.PARAMETER OutputDir
  Directory for crops, CSV, JSON, and the crop manifest.

.PARAMETER CropManifestCsv
  Optional UTF-8 CSV with Deck, MediaPath, X, Y, Width, Height columns.  The
  rectangle is in source-image pixels before padding and scaling.

.PARAMETER PaddingPx
  Padding added around the selected region before transparent-margin trim.

.PARAMETER Scale
  Output scale factor after cropping.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FormulaImageCandidateCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$CropManifestCsv = '',

    [ValidateRange(0, 128)]
    [int]$PaddingPx = 8,

    [ValidateRange(0.25, 4.0)]
    [double]$Scale = 2.0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-RowValue {
    param([object]$Row, [string]$Name, [object]$Default = '')
    if ($null -eq $Row) { return $Default }
    $property = $Row.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-NormalizedKey {
    param([string]$Deck, [string]$MediaPath)
    return (([string]$Deck).Trim().ToLowerInvariant() + '|' + ([string]$MediaPath).Replace('\', '/').Trim().ToLowerInvariant())
}

function Convert-ToInt {
    param([object]$Value, [string]$Name)
    $result = 0
    if (-not [int]::TryParse([string]$Value, [ref]$result)) { throw "$Name must be an integer." }
    return $result
}

function Get-ExplicitCrop {
    param($CropMap, [string]$Deck, [string]$MediaPath)
    $key = Get-NormalizedKey -Deck $Deck -MediaPath $MediaPath
    if (-not $CropMap.ContainsKey($key)) { return $null }
    $row = $CropMap[$key]
    return [pscustomobject]@{
        X = Convert-ToInt -Value (Get-RowValue -Row $row -Name 'X') -Name 'X'
        Y = Convert-ToInt -Value (Get-RowValue -Row $row -Name 'Y') -Name 'Y'
        Width = Convert-ToInt -Value (Get-RowValue -Row $row -Name 'Width') -Name 'Width'
        Height = Convert-ToInt -Value (Get-RowValue -Row $row -Name 'Height') -Name 'Height'
    }
}

function Get-AlphaBounds {
    param([System.Drawing.Bitmap]$Bitmap, [int]$Left, [int]$Top, [int]$Width, [int]$Height)
    $minX = $Width; $minY = $Height; $maxX = -1; $maxY = -1
    for ($y = 0; $y -lt $Height; $y++) {
        for ($x = 0; $x -lt $Width; $x++) {
            $pixel = $Bitmap.GetPixel($Left + $x, $Top + $y)
            if ($pixel.A -gt 0) {
                if ($x -lt $minX) { $minX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }
    if ($maxX -lt 0) { return $null }
    return [pscustomobject]@{ X = $minX; Y = $minY; Width = $maxX - $minX + 1; Height = $maxY - $minY + 1 }
}

function New-CropBitmap {
    param(
        [System.Drawing.Bitmap]$Source,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [int]$Padding,
        [double]$Scale
    )
    $left = [Math]::Max(0, $X - $Padding)
    $top = [Math]::Max(0, $Y - $Padding)
    $right = [Math]::Min($Source.Width, $X + $Width + $Padding)
    $bottom = [Math]::Min($Source.Height, $Y + $Height + $Padding)
    $cropWidth = $right - $left
    $cropHeight = $bottom - $top
    if ($cropWidth -le 0 -or $cropHeight -le 0) { throw 'Crop rectangle is empty after clamping.' }

    # Only transparent pixels may be trimmed automatically.  For opaque
    # JPEG/PNG sources the selected rectangle remains the source of truth.
    $alphaBounds = Get-AlphaBounds -Bitmap $Source -Left $left -Top $top -Width $cropWidth -Height $cropHeight
    if ($null -ne $alphaBounds) {
        $left += $alphaBounds.X
        $top += $alphaBounds.Y
        $cropWidth = $alphaBounds.Width
        $cropHeight = $alphaBounds.Height
    }

    $outWidth = [Math]::Max(1, [int][Math]::Round($cropWidth * $Scale))
    $outHeight = [Math]::Max(1, [int][Math]::Round($cropHeight * $Scale))
    $output = New-Object System.Drawing.Bitmap($outWidth, $outHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = $null
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($output)
        $graphics.Clear([System.Drawing.Color]::Transparent)
        $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $destination = New-Object System.Drawing.Rectangle(0, 0, $outWidth, $outHeight)
        $sourceRect = New-Object System.Drawing.Rectangle($left, $top, $cropWidth, $cropHeight)
        $graphics.DrawImage($Source, $destination, $sourceRect, [System.Drawing.GraphicsUnit]::Pixel)
    } finally {
        if ($null -ne $graphics) { $graphics.Dispose() }
    }
    return [pscustomobject]@{ Bitmap = $output; X = $left; Y = $top; Width = $cropWidth; Height = $cropHeight; Scale = $Scale }
}

function Write-CsvRows {
    param([object[]]$Rows, [string]$Path, [string[]]$Columns)
    if (@($Rows).Count -eq 0) {
        Write-Utf8BomText -Path $Path -Text (($Columns -join ',') + "`r`n")
    } else {
        Write-Utf8BomCsv -InputObject @($Rows | Select-Object $Columns) -Path $Path
    }
}

$csvPath = [System.IO.Path]::GetFullPath($FormulaImageCandidateCsv)
$outputDir = [System.IO.Path]::GetFullPath($OutputDir)
$cropCsvPath = if ([string]::IsNullOrWhiteSpace($CropManifestCsv)) { '' } else { [System.IO.Path]::GetFullPath($CropManifestCsv) }
if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) { throw "Formula image candidate CSV not found: $csvPath" }
if ($cropCsvPath -and -not (Test-Path -LiteralPath $cropCsvPath -PathType Leaf)) { throw "Crop manifest CSV not found: $cropCsvPath" }
if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
$cropDir = Join-Path $outputDir 'crops'
if (-not (Test-Path -LiteralPath $cropDir -PathType Container)) { New-Item -ItemType Directory -Path $cropDir -Force | Out-Null }

Add-Type -AssemblyName System.Drawing
$cropMap = @{}
if ($cropCsvPath) {
    foreach ($cropRow in @(Import-Csv -LiteralPath $cropCsvPath -Encoding UTF8)) {
        $key = Get-NormalizedKey -Deck ([string](Get-RowValue -Row $cropRow -Name 'Deck')) -MediaPath ([string](Get-RowValue -Row $cropRow -Name 'MediaPath'))
        if ($cropMap.ContainsKey($key)) { throw "Duplicate crop manifest key: $key" }
        $cropMap[$key] = $cropRow
    }
}

$results = New-Object System.Collections.Generic.List[object]
$rows = @(Import-Csv -LiteralPath $csvPath -Encoding UTF8 | Where-Object { [string](Get-RowValue -Row $_ -Name 'FormulaImageCandidate') -in @('True', 'true', '1') })
for ($index = 0; $index -lt $rows.Count; $index++) {
    $row = $rows[$index]
    $deck = [string](Get-RowValue -Row $row -Name 'Deck')
    $mediaPath = [string](Get-RowValue -Row $row -Name 'MediaPath')
    $mediaRole = [string](Get-RowValue -Row $row -Name 'MediaRole' -Default 'StandaloneImage')
    $sourcePath = [string](Get-RowValue -Row $row -Name 'FormulaExtractedPath')
    $status = 'Failed'
    $reason = ''
    $cropPath = ''
    $sourceHash = ''
    $cropHash = ''
    $sourceWidth = 0; $sourceHeight = 0
    $cropX = $null; $cropY = $null; $cropWidth = $null; $cropHeight = $null
    $outputWidth = 0; $outputHeight = 0
    $bitmap = $null
    $cropBitmap = $null
    try {
        if ([string]::IsNullOrWhiteSpace($sourcePath) -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw 'Extracted source image is missing.' }
        $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $bitmap = [System.Drawing.Bitmap]::new($sourcePath)
        $sourceWidth = $bitmap.Width; $sourceHeight = $bitmap.Height
        $explicit = Get-ExplicitCrop -CropMap $cropMap -Deck $deck -MediaPath $mediaPath
        if ($mediaRole -ne 'StandaloneImage' -and $null -eq $explicit) {
            $status = 'SplitFailed'
            $reason = 'MixedImage requires an explicit formula region; original image is preserved.'
        } else {
            $region = if ($null -ne $explicit) { $explicit } else { [pscustomobject]@{ X = 0; Y = 0; Width = $sourceWidth; Height = $sourceHeight } }
            if ($region.X -lt 0 -or $region.Y -lt 0 -or $region.Width -le 0 -or $region.Height -le 0 -or
                $region.X + $region.Width -gt $sourceWidth -or $region.Y + $region.Height -gt $sourceHeight) {
                throw 'Crop rectangle is outside source-image bounds.'
            }
            $cropResult = New-CropBitmap -Source $bitmap -X $region.X -Y $region.Y -Width $region.Width -Height $region.Height -Padding $PaddingPx -Scale $Scale
            $cropBitmap = $cropResult.Bitmap
            $cropX = $cropResult.X; $cropY = $cropResult.Y; $cropWidth = $cropResult.Width; $cropHeight = $cropResult.Height
            $fileName = ('crop-{0:0000}-{1}.png' -f ($index + 1), (Convert-ToSafePathSegment -Name ([System.IO.Path]::GetFileNameWithoutExtension($mediaPath))))
            $cropPath = Join-Path $cropDir $fileName
            $cropBitmap.Save($cropPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $cropHash = (Get-FileHash -LiteralPath $cropPath -Algorithm SHA256).Hash.ToLowerInvariant()
            $outputWidth = $cropBitmap.Width; $outputHeight = $cropBitmap.Height
            $status = if ($null -ne $explicit) { 'Cropped' } else { 'FullImageFallback' }
            $reason = if ($null -ne $explicit) { 'Explicit region crop with transparent-only trim.' } else { 'Standalone image has no region annotation; full image retained as candidate.' }
        }
    } catch {
        $status = 'Failed'
        $reason = $_.Exception.Message
        if ($cropPath -and (Test-Path -LiteralPath $cropPath)) { Remove-Item -LiteralPath $cropPath -Force }
        $cropPath = ''; $cropHash = ''
    } finally {
        if ($null -ne $cropBitmap) { $cropBitmap.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
    }
    $results.Add([pscustomobject]@{
        Index = $index + 1
        Deck = $deck
        MediaPath = $mediaPath
        MediaRole = $mediaRole
        SourcePath = $sourcePath
        SourceSha256 = $sourceHash
        SourceWidth = $sourceWidth
        SourceHeight = $sourceHeight
        CropX = $cropX
        CropY = $cropY
        CropWidth = $cropWidth
        CropHeight = $cropHeight
        PaddingPx = $PaddingPx
        Scale = $Scale
        OutputPath = $cropPath
        OutputSha256 = $cropHash
        OutputWidth = $outputWidth
        OutputHeight = $outputHeight
        Status = $status
        Reason = $reason
        Rollback = $sourcePath
    }) | Out-Null
}

$resultRows = @($results.ToArray())
$resultCsv = Join-Path $outputDir 'formula-image-crops.csv'
$resultJson = Join-Path $outputDir 'formula-image-crops.json'
$resultManifest = Join-Path $outputDir 'formula-image-crops-manifest.json'
$columns = @('Index', 'Deck', 'MediaPath', 'MediaRole', 'SourcePath', 'SourceSha256', 'SourceWidth', 'SourceHeight', 'CropX', 'CropY', 'CropWidth', 'CropHeight', 'PaddingPx', 'Scale', 'OutputPath', 'OutputSha256', 'OutputWidth', 'OutputHeight', 'Status', 'Reason', 'Rollback')
Write-CsvRows -Rows $resultRows -Path $resultCsv -Columns $columns
Write-Utf8BomText -Path $resultJson -Text ($resultRows | ConvertTo-Json -Depth 8)
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    inputCsv = [ordered]@{ path = $csvPath; sha256 = (Get-FileHash -LiteralPath $csvPath -Algorithm SHA256).Hash.ToLowerInvariant() }
    cropManifestCsv = if ($cropCsvPath) { [ordered]@{ path = $cropCsvPath; sha256 = (Get-FileHash -LiteralPath $cropCsvPath -Algorithm SHA256).Hash.ToLowerInvariant() } } else { $null }
    protocol = [ordered]@{ paddingPx = $PaddingPx; scale = $Scale; transparentTrimOnly = $true; mixedImageNeedsExplicitRegion = $true; originalPreservedOnFailure = $true }
    candidateCount = $resultRows.Count
    croppedCount = @($resultRows | Where-Object { $_.Status -in @('Cropped', 'FullImageFallback') }).Count
    splitFailedCount = @($resultRows | Where-Object { $_.Status -eq 'SplitFailed' }).Count
    failedCount = @($resultRows | Where-Object { $_.Status -eq 'Failed' }).Count
    csv = $resultCsv
    json = $resultJson
    rows = $resultRows
}
Write-Utf8BomText -Path $resultManifest -Text ($manifest | ConvertTo-Json -Depth 12)
Write-Output ("Formula image crop protocol done: {0}`nCandidates: {1}; Cropped: {2}; SplitFailed: {3}; Failed: {4}" -f $outputDir, $manifest.candidateCount, $manifest.croppedCount, $manifest.splitFailedCount, $manifest.failedCount)
