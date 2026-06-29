<#
.SYNOPSIS
  Run a controlled Real-ESRGAN probe on PPTX image candidates.

.DESCRIPTION
  Reads the CSV from Export-PptxImageCandidates.ps1, selects a small set of
  extracted candidate images, enhances them with realesrgan-ncnn-vulkan, then
  writes before/after reports and a contact sheet. This script does not modify
  any PPTX file.

.PARAMETER CandidateCsv
  Path to pptx-image-candidates.csv.

.PARAMETER OutputDir
  Directory where probe images and reports are written.

.PARAMETER MediaPath
  Optional media paths to probe. When omitted, the script chooses a small,
  low-resolution candidate set.

.PARAMETER DeckPattern
  Optional wildcard filter for deck names.

.PARAMETER MaxImages
  Maximum auto-selected images when MediaPath is omitted.

.PARAMETER MaxInputPixels
  Maximum source pixels for auto-selection. Large images usually do not benefit
  from super-resolution and can bloat outputs.

.PARAMETER RealesrganPath
  Path to realesrgan-ncnn-vulkan.exe.

.PARAMETER ToolRoot
  Optional folder scanned recursively for oxipng.exe.

.PARAMETER DeliveryScale
  Post-process enhanced PNG to a JPEG at OriginalSize * DeliveryScale. The
  default keeps the original pixel size to improve clarity while reducing PPTX
  size; use 2 only when a visibly low-resolution photo needs more pixels.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CandidateCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string[]]$MediaPath = @(),

    [string]$DeckPattern = '*',

    [ValidateRange(1, 50)]
    [int]$MaxImages = 6,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxInputPixels = 1200000,

    [ValidateRange(2, 4)]
    [int]$Scale = 4,

    [ValidateSet('realesrgan-x4plus', 'realesrgan-x4plus-anime', 'realesr-animevideov3', 'realesrnet-x4plus')]
    [string]$ModelName = 'realesrgan-x4plus',

    [ValidateRange(0, 100)]
    [int]$MinPhotoScore = 55,

    [ValidateRange(1, 4)]
    [int]$DeliveryScale = 1,

    [ValidateRange(1, 100)]
    [int]$DeliveryJpegQuality = 85,

    [string]$RealesrganPath = '',

    [string]$ToolRoot = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($RealesrganPath)) {
    $RealesrganPath = Join-Path $PSScriptRoot 'vendor\realesrgan-ncnn-vulkan-20220424\realesrgan-ncnn-vulkan.exe'
}
if ([string]::IsNullOrWhiteSpace($ToolRoot)) {
    $ToolRoot = Join-Path $PSScriptRoot 'vendor'
}

function Convert-ToSafePathSegment {
    param([string]$Name)
    $safe = $Name
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$ch, '_')
    }
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'image' }
    return $safe
}

function Resolve-Tool {
    param([string]$Name, [string]$Root)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $cmd) { return $cmd.Source }
    if (-not [string]::IsNullOrWhiteSpace($Root) -and (Test-Path -LiteralPath $Root)) {
        $exeName = if ($Name.EndsWith('.exe', [System.StringComparison]::OrdinalIgnoreCase)) { $Name } else { "$Name.exe" }
        $match = Get-ChildItem -LiteralPath $Root -Recurse -File -Filter $exeName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $match) { return $match.FullName }
    }
    return ''
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

function Save-ResizedJpeg {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [int]$Width,
        [int]$Height,
        [int]$Quality
    )

    Add-Type -AssemblyName System.Drawing
    $source = $null
    $bitmap = $null
    $graphics = $null
    $encoderParams = $null
    try {
        $source = [System.Drawing.Image]::FromFile($InputPath)
        $bitmap = New-Object System.Drawing.Bitmap($Width, $Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.DrawImage($source, 0, 0, $Width, $Height)

        $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
            Where-Object { $_.MimeType -eq 'image/jpeg' } |
            Select-Object -First 1
        if ($null -eq $codec) { throw 'JPEG codec not available.' }

        $encoder = [System.Drawing.Imaging.Encoder]::Quality
        $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters 1
        $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter $encoder, ([int64]$Quality)
        $bitmap.Save($OutputPath, $codec, $encoderParams)
    } finally {
        if ($null -ne $encoderParams) { $encoderParams.Dispose() }
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
        if ($null -ne $source) { $source.Dispose() }
    }
}

function Invoke-Oxipng {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$OxipngPath
    )
    if ([string]::IsNullOrWhiteSpace($OxipngPath)) { return $false }
    $raw = & $OxipngPath -o 4 --strip safe --out $OutputPath $InputPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Warning ("oxipng failed: " + (($raw | Out-String).Trim()))
        return $false
    }
    return $true
}

function New-ProbeContactSheet {
    param(
        [object[]]$Rows,
        [string]$OutputPath
    )

    Add-Type -AssemblyName System.Drawing
    $items = @($Rows | Where-Object { $_.Status -eq 'Success' -and (Test-Path -LiteralPath $_.OriginalCopy) -and (Test-Path -LiteralPath $_.BestOutput) })
    if ($items.Count -eq 0) { return $false }

    $tileWidth = 360
    $tileHeight = 260
    $labelHeight = 52
    $columns = 2
    $rowsCount = $items.Count
    $sheetWidth = $tileWidth * $columns
    $sheetHeight = $tileHeight * $rowsCount

    $bitmap = $null
    $graphics = $null
    $font = $null
    $smallFont = $null
    $pen = $null
    try {
        $bitmap = New-Object System.Drawing.Bitmap($sheetWidth, $sheetHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::FromArgb(248, 248, 248))
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $font = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5)
        $smallFont = New-Object System.Drawing.Font('Microsoft YaHei UI', 7.5)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(210, 210, 210), 1)

        for ($i = 0; $i -lt $items.Count; $i++) {
            $y = $i * $tileHeight
            $pair = @(
                @{ Path = $items[$i].OriginalCopy; Label = 'Before' },
                @{ Path = $items[$i].BestOutput; Label = 'After' }
            )
            for ($col = 0; $col -lt 2; $col++) {
                $x = $col * $tileWidth
                $graphics.FillRectangle([System.Drawing.Brushes]::White, $x + 6, $y + 6, $tileWidth - 12, $tileHeight - 12)
                $graphics.DrawRectangle($pen, $x + 6, $y + 6, $tileWidth - 12, $tileHeight - 12)
                $img = $null
                try {
                    $img = [System.Drawing.Image]::FromFile($pair[$col].Path)
                    $maxW = $tileWidth - 24
                    $maxH = $tileHeight - $labelHeight - 24
                    $scaleFactor = [Math]::Min($maxW / [double]$img.Width, $maxH / [double]$img.Height)
                    $drawW = [Math]::Max(1, [int][Math]::Round($img.Width * $scaleFactor))
                    $drawH = [Math]::Max(1, [int][Math]::Round($img.Height * $scaleFactor))
                    $drawX = $x + [int](($tileWidth - $drawW) / 2)
                    $drawY = $y + 14 + [int](($maxH - $drawH) / 2)
                    $graphics.DrawImage($img, $drawX, $drawY, $drawW, $drawH)
                } finally {
                    if ($null -ne $img) { $img.Dispose() }
                }
                $labelY = $y + $tileHeight - $labelHeight + 6
                $graphics.DrawString($pair[$col].Label, $font, [System.Drawing.Brushes]::Black, $x + 12, $labelY)
                $graphics.DrawString("$($items[$i].MediaPath) slides=$($items[$i].UsedOnSlides)", $smallFont, [System.Drawing.Brushes]::DimGray, $x + 12, $labelY + 18)
                $graphics.DrawString("$($items[$i].OriginalWidth)x$($items[$i].OriginalHeight) -> $($items[$i].EnhancedWidth)x$($items[$i].EnhancedHeight)", $smallFont, [System.Drawing.Brushes]::DimGray, $x + 12, $labelY + 34)
            }
        }
        $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        return $true
    } finally {
        if ($null -ne $pen) { $pen.Dispose() }
        if ($null -ne $smallFont) { $smallFont.Dispose() }
        if ($null -ne $font) { $font.Dispose() }
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
    }
}

$CandidateCsv = [System.IO.Path]::GetFullPath($CandidateCsv)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
$RealesrganPath = [System.IO.Path]::GetFullPath($RealesrganPath)
$ToolRoot = if ([string]::IsNullOrWhiteSpace($ToolRoot)) { '' } else { [System.IO.Path]::GetFullPath($ToolRoot) }

if (-not (Test-Path -LiteralPath $CandidateCsv)) { throw "CandidateCsv not found: $CandidateCsv" }
if (-not (Test-Path -LiteralPath $RealesrganPath)) { throw "RealesrganPath not found: $RealesrganPath" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$modelDir = Join-Path (Split-Path -Parent $RealesrganPath) 'models'
if (-not (Test-Path -LiteralPath $modelDir)) { throw "Real-ESRGAN model directory not found: $modelDir" }
$oxipngPath = Resolve-Tool -Name 'oxipng' -Root $ToolRoot

$rows = @(Import-Csv -LiteralPath $CandidateCsv | Where-Object {
    $_.EnhancementCandidate -eq 'True' -and
    $_.Deck -like $DeckPattern -and
    -not [string]::IsNullOrWhiteSpace($_.ExtractedPath) -and
    (Test-Path -LiteralPath $_.ExtractedPath) -and
    [int]$_.PhotoScore -ge $MinPhotoScore
})

if ($MediaPath.Count -gt 0) {
    $lookup = @{}
    foreach ($path in $MediaPath) { $lookup[$path] = $true }
    $selected = @($rows | Where-Object { $lookup.ContainsKey($_.MediaPath) })
} else {
    $selected = @($rows | Where-Object {
        [int64]$_.Pixels -le $MaxInputPixels -and
        [int]$_.Width -ge 300 -and
        [int]$_.Height -ge 250 -and
        ([double]$_.Width / [double]$_.Height) -le 2.5 -and
        ([double]$_.Height / [double]$_.Width) -le 2.5
    } | Sort-Object @{ Expression = { if ($_.CandidateLevel -eq 'LikelyPhoto') { 0 } else { 1 } } }, @{ Expression = { [int]$_.PhotoScore }; Descending = $true }, @{ Expression = { [int64]$_.Pixels } } | Select-Object -First $MaxImages)
}

if ($selected.Count -eq 0) { throw 'No eligible candidate images found for enhancement probe.' }

$probeRows = New-Object System.Collections.Generic.List[object]
$probeIndex = 0
foreach ($row in $selected) {
    $probeIndex++
    $safeBase = 'probe-{0:000}' -f $probeIndex
    $inputExt = [System.IO.Path]::GetExtension([string]$row.ExtractedPath).ToLowerInvariant()
    if ($inputExt -eq '.jpeg') { $inputExt = '.jpg' }
    $originalCopy = Join-Path $OutputDir ($safeBase + '.before' + $inputExt)
    $enhancedPath = Join-Path $OutputDir ($safeBase + ".realesrgan-x$Scale.png")
    $compressedPath = Join-Path $OutputDir ($safeBase + ".realesrgan-x$Scale.oxipng.png")
    $deliveryPath = Join-Path $OutputDir ($safeBase + ".delivery-x$DeliveryScale-q$DeliveryJpegQuality.jpg")
    $status = 'Success'
    $message = ''
    try {
        Copy-Item -LiteralPath ([string]$row.ExtractedPath) -Destination $originalCopy -Force
        $before = Get-BasicImageInfo -Path $originalCopy
        Write-Host "Enhancing: $($row.Deck) $($row.MediaPath)"
        $oldErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $raw = & $RealesrganPath -i $originalCopy -o $enhancedPath -n $ModelName -s $Scale -f png -m $modelDir 2>&1
            $exitCode = $LASTEXITCODE
        } finally {
            $ErrorActionPreference = $oldErrorActionPreference
        }
        if ($exitCode -ne 0) {
            throw ("Real-ESRGAN failed: " + (($raw | Out-String).Trim()))
        }
        $bestPath = $enhancedPath
        if (Invoke-Oxipng -InputPath $enhancedPath -OutputPath $compressedPath -OxipngPath $oxipngPath) {
            $bestPath = $compressedPath
        }
        $fullEnhanced = Get-BasicImageInfo -Path $bestPath
        $deliveryWidth = [Math]::Min([int]($before.Width * $DeliveryScale), [int]$fullEnhanced.Width)
        $deliveryHeight = [Math]::Min([int]($before.Height * $DeliveryScale), [int]$fullEnhanced.Height)
        Save-ResizedJpeg -InputPath $bestPath -OutputPath $deliveryPath -Width $deliveryWidth -Height $deliveryHeight -Quality $DeliveryJpegQuality
        $bestPath = $deliveryPath
        $after = Get-BasicImageInfo -Path $bestPath
    } catch {
        $status = 'Failed'
        $message = $_.Exception.Message
        $before = if (Test-Path -LiteralPath $originalCopy) { Get-BasicImageInfo -Path $originalCopy } else { [pscustomobject]@{ Width = 0; Height = 0; Bytes = 0 } }
        $fullEnhanced = if (Test-Path -LiteralPath $enhancedPath) { Get-BasicImageInfo -Path $enhancedPath } else { [pscustomobject]@{ Width = 0; Height = 0; Bytes = 0 } }
        $after = [pscustomobject]@{ Width = 0; Height = 0; Bytes = 0 }
        $bestPath = ''
    }

    $probeRows.Add([pscustomobject]@{
        Deck = $row.Deck
        MediaPath = $row.MediaPath
        UsedOnSlides = $row.UsedOnSlides
        CandidateLevel = $row.CandidateLevel
        PhotoScore = [int]$row.PhotoScore
        Status = $status
        Message = $message
        OriginalCopy = $originalCopy
        EnhancedPath = $enhancedPath
        CompressedPath = if (Test-Path -LiteralPath $compressedPath) { $compressedPath } else { '' }
        DeliveryPath = if (Test-Path -LiteralPath $deliveryPath) { $deliveryPath } else { '' }
        BestOutput = $bestPath
        OriginalWidth = $before.Width
        OriginalHeight = $before.Height
        OriginalBytes = $before.Bytes
        FullEnhancedWidth = $fullEnhanced.Width
        FullEnhancedHeight = $fullEnhanced.Height
        FullEnhancedBytes = $fullEnhanced.Bytes
        EnhancedWidth = $after.Width
        EnhancedHeight = $after.Height
        EnhancedBytes = $after.Bytes
        SizeMultiplier = if ($before.Bytes -gt 0) { [Math]::Round($after.Bytes / [double]$before.Bytes, 2) } else { 0 }
        Scale = $Scale
        Model = $ModelName
    }) | Out-Null
}

$csvPath = Join-Path $OutputDir 'pptx-image-enhancement-probe.csv'
$jsonPath = Join-Path $OutputDir 'pptx-image-enhancement-probe.json'
$manifestPath = Join-Path $OutputDir 'pptx-image-enhancement-probe-manifest.json'
$sheetPath = Join-Path $OutputDir 'pptx-image-enhancement-probe.contact-sheet.png'

$probeRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
$probeRows | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
$sheetCreated = New-ProbeContactSheet -Rows $probeRows.ToArray() -OutputPath $sheetPath

$successRows = @($probeRows | Where-Object { $_.Status -eq 'Success' })
$manifest = [pscustomobject]@{
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    candidateCsv = $CandidateCsv
    outputDir = $OutputDir
    selectedCount = $selected.Count
    successCount = $successRows.Count
    failedCount = $probeRows.Count - $successRows.Count
    scale = $Scale
    deliveryScale = $DeliveryScale
    deliveryJpegQuality = $DeliveryJpegQuality
    model = $ModelName
    realesrganPath = $RealesrganPath
    oxipngPath = $oxipngPath
    csv = $csvPath
    json = $jsonPath
    contactSheet = if ($sheetCreated) { $sheetPath } else { '' }
    note = 'Probe outputs are for visual review only. They are not written back into PPTX files.'
}
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Host "Enhancement probe done: $OutputDir"
Write-Host "Success: $($successRows.Count) / Selected: $($selected.Count)"
