<#
.SYNOPSIS
  Export embedded PPTX image candidates for optional clarity enhancement.

.DESCRIPTION
  This is a read-only scanner. It inspects ppt/media image files, estimates
  whether each image is more likely to be a photo/natural scene or a
  diagram/text-like asset, and writes CSV/JSON reports. It can optionally copy
  likely candidates and build a contact sheet for visual review.

.PARAMETER InputPath
  Path to a PPTX file or a directory containing PPTX files.

.PARAMETER OutputDir
  Directory where reports and optional extracted candidate images are written.

.PARAMETER Recurse
  Search subdirectories when InputPath is a directory.

.PARAMETER FilePattern
  File name pattern used when InputPath is a directory.

.PARAMETER MinBytes
  Minimum media file size considered large enough for enhancement.

.PARAMETER MinPixels
  Minimum pixel count considered large enough for enhancement.

.PARAMETER SampleSize
  Maximum width/height used for fast color and edge sampling.

.PARAMETER ExportCandidates
  Copy LikelyPhoto and MaybePhoto image files to candidate-images.

.PARAMETER ContactSheet
  Build a PNG contact sheet from exported candidates.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [switch]$Recurse,

    [string]$FilePattern = '*.pptx',

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MinBytes = 51200,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MinPixels = 90000,

    [ValidateRange(16, 256)]
    [int]$SampleSize = 64,

    [switch]$ExportCandidates,

    [switch]$ContactSheet,

    [ValidateRange(1, 500)]
    [int]$MaxContactSheetItems = 80
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

function Convert-ToSafePathSegment {
    param([string]$Name)
    $safe = $Name
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$ch, '_')
    }
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'presentation' }
    return $safe
}

function Get-SlideMediaMap {
    param([System.IO.FileInfo]$PptxFile)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $map = @{}
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($PptxFile.FullName)
        $relEntries = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/_rels/slide(\d+)\.xml\.rels$' })
        foreach ($relEntry in $relEntries) {
            if ($relEntry.FullName -notmatch '^ppt/slides/_rels/slide(\d+)\.xml\.rels$') { continue }
            $slideNumber = [int]$Matches[1]
            $sourcePart = "ppt/slides/slide$slideNumber.xml"
            $xmlText = Read-ZipEntryText -Zip $zip -EntryName $relEntry.FullName
            if ([string]::IsNullOrWhiteSpace($xmlText)) { continue }
            $xml = New-Object System.Xml.XmlDocument
            $xml.PreserveWhitespace = $false
            $xml.LoadXml($xmlText)
            foreach ($rel in $xml.GetElementsByTagName('Relationship')) {
                $type = [string]$rel.GetAttribute('Type')
                $targetMode = [string]$rel.GetAttribute('TargetMode')
                if ($targetMode -eq 'External') { continue }
                if (-not $type.EndsWith('/image', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                $resolved = Resolve-PackageTarget -SourcePart $sourcePart -Target ([string]$rel.GetAttribute('Target'))
                if (-not $resolved.StartsWith('ppt/media/', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if (-not $map.ContainsKey($resolved)) {
                    $map[$resolved] = New-Object System.Collections.Generic.List[int]
                }
                if (-not $map[$resolved].Contains($slideNumber)) {
                    $map[$resolved].Add($slideNumber) | Out-Null
                }
            }
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
    return $map
}

function Get-ImageAnalysis {
    param(
        [string]$Path,
        [int]$SampleSize
    )

    Add-Type -AssemblyName System.Drawing
    $image = $null
    $thumb = $null
    $graphics = $null
    try {
        $image = [System.Drawing.Image]::FromFile($Path)
        $width = [int]$image.Width
        $height = [int]$image.Height
        if ($width -le 0 -or $height -le 0) { throw 'Invalid image size.' }

        $scale = [Math]::Min($SampleSize / [double]$width, $SampleSize / [double]$height)
        if ($scale -gt 1) { $scale = 1 }
        $sampleWidth = [Math]::Max(1, [int][Math]::Round($width * $scale))
        $sampleHeight = [Math]::Max(1, [int][Math]::Round($height * $scale))

        $thumb = New-Object System.Drawing.Bitmap($sampleWidth, $sampleHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [System.Drawing.Graphics]::FromImage($thumb)
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighSpeed
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::Bilinear
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighSpeed
        $graphics.DrawImage($image, 0, 0, $sampleWidth, $sampleHeight)

        $bins = New-Object 'System.Collections.Generic.HashSet[string]'
        $opaque = 0
        $transparent = 0
        $gray = 0
        $edgeCount = 0
        $edgeComparisons = 0
        $prevRowLum = New-Object double[] $sampleWidth

        for ($y = 0; $y -lt $sampleHeight; $y++) {
            $leftLum = $null
            for ($x = 0; $x -lt $sampleWidth; $x++) {
                $color = $thumb.GetPixel($x, $y)
                if ($color.A -lt 240) {
                    $transparent++
                    continue
                }

                $opaque++
                $rBin = [int]($color.R / 16)
                $gBin = [int]($color.G / 16)
                $bBin = [int]($color.B / 16)
                $bins.Add("$rBin,$gBin,$bBin") | Out-Null

                if (([Math]::Abs($color.R - $color.G) + [Math]::Abs($color.G - $color.B) + [Math]::Abs($color.R - $color.B)) -lt 42) {
                    $gray++
                }

                $lum = (0.2126 * $color.R) + (0.7152 * $color.G) + (0.0722 * $color.B)
                if ($null -ne $leftLum) {
                    $edgeComparisons++
                    if ([Math]::Abs($lum - [double]$leftLum) -gt 32) { $edgeCount++ }
                }
                if ($y -gt 0) {
                    $edgeComparisons++
                    if ([Math]::Abs($lum - $prevRowLum[$x]) -gt 32) { $edgeCount++ }
                }
                $prevRowLum[$x] = $lum
                $leftLum = $lum
            }
        }

        $samplePixels = [Math]::Max(1, $sampleWidth * $sampleHeight)
        $colorDiversity = if ($opaque -gt 0) { [Math]::Round($bins.Count / [double]$opaque, 4) } else { 0 }
        $grayRatio = if ($opaque -gt 0) { [Math]::Round($gray / [double]$opaque, 4) } else { 0 }
        $transparentRatio = [Math]::Round($transparent / [double]$samplePixels, 4)
        $edgeDensity = if ($edgeComparisons -gt 0) { [Math]::Round($edgeCount / [double]$edgeComparisons, 4) } else { 0 }

        return [pscustomobject]@{
            Width = $width
            Height = $height
            Pixels = [int64]$width * [int64]$height
            PixelFormat = $image.PixelFormat.ToString()
            SampleWidth = $sampleWidth
            SampleHeight = $sampleHeight
            UniqueColorBins = $bins.Count
            ColorDiversity = $colorDiversity
            GrayRatio = $grayRatio
            TransparentRatio = $transparentRatio
            EdgeDensity = $edgeDensity
        }
    } finally {
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $thumb) { $thumb.Dispose() }
        if ($null -ne $image) { $image.Dispose() }
    }
}

function Get-CandidateAssessment {
    param(
        [string]$Extension,
        [int64]$Bytes,
        [object]$Analysis,
        [int]$MinBytes,
        [int]$MinPixels
    )

    $score = 0
    $reasons = New-Object System.Collections.Generic.List[string]
    $ext = $Extension.ToLowerInvariant()

    if ($ext -in @('.jpg', '.jpeg')) {
        $score += 35
        $reasons.Add('JPEG source') | Out-Null
    } elseif ($ext -eq '.png') {
        $score += 10
        $reasons.Add('PNG source') | Out-Null
    } else {
        return [pscustomobject]@{
            PhotoScore = 0
            CandidateLevel = 'Unsupported'
            EnhancementCandidate = $false
            Reason = 'Unsupported image extension'
        }
    }

    if ($Bytes -ge $MinBytes) {
        $score += 8
    } else {
        $score -= 20
        $reasons.Add("Below MinBytes $MinBytes") | Out-Null
    }

    if ($Analysis.Pixels -ge $MinPixels) {
        $score += 10
    } else {
        $score -= 20
        $reasons.Add("Below MinPixels $MinPixels") | Out-Null
    }

    if ($Analysis.Width -ge 240 -and $Analysis.Height -ge 240) {
        $score += 10
    } else {
        $score -= 20
        $reasons.Add('Small width or height') | Out-Null
    }

    if ($Analysis.TransparentRatio -gt 0.02) {
        $score -= 25
        $reasons.Add('Transparency suggests overlay/diagram asset') | Out-Null
    }

    if ($Analysis.UniqueColorBins -ge 700) {
        $score += 25
        $reasons.Add('High color variety') | Out-Null
    } elseif ($Analysis.UniqueColorBins -ge 300) {
        $score += 15
        $reasons.Add('Moderate color variety') | Out-Null
    } elseif ($Analysis.UniqueColorBins -ge 120) {
        $score += 5
        $reasons.Add('Limited color variety') | Out-Null
    } else {
        $score -= 15
        $reasons.Add('Very low color variety') | Out-Null
    }

    if ($Analysis.ColorDiversity -ge 0.14) {
        $score += 15
    } elseif ($Analysis.ColorDiversity -ge 0.07) {
        $score += 8
    } else {
        $score -= 10
        $reasons.Add('Low sampled color diversity') | Out-Null
    }

    if ($Analysis.EdgeDensity -ge 0.04 -and $Analysis.EdgeDensity -le 0.45) {
        $score += 10
    } elseif ($Analysis.EdgeDensity -gt 0.55) {
        $score -= 15
        $reasons.Add('Very high edge density suggests text/diagram') | Out-Null
    } elseif ($Analysis.EdgeDensity -lt 0.015) {
        $score -= 5
        $reasons.Add('Very low edge density') | Out-Null
    }

    if ($Analysis.GrayRatio -le 0.70) {
        $score += 8
    } elseif ($Analysis.GrayRatio -gt 0.90) {
        $score -= 12
        $reasons.Add('Mostly grayscale') | Out-Null
    }

    if ($score -ge 70 -and $Bytes -ge $MinBytes -and $Analysis.Pixels -ge $MinPixels) {
        $level = 'LikelyPhoto'
        $candidate = $true
    } elseif ($score -ge 55 -and $Bytes -ge $MinBytes -and $Analysis.Pixels -ge $MinPixels) {
        $level = 'MaybePhoto'
        $candidate = $true
    } elseif ($Bytes -lt $MinBytes -or $Analysis.Pixels -lt $MinPixels) {
        $level = 'TooSmall'
        $candidate = $false
    } else {
        $level = 'LikelyDiagramOrText'
        $candidate = $false
    }

    if ($reasons.Count -eq 0) { $reasons.Add('No specific risk flags') | Out-Null }
    return [pscustomobject]@{
        PhotoScore = [Math]::Max(0, [Math]::Min(100, $score))
        CandidateLevel = $level
        EnhancementCandidate = $candidate
        Reason = ($reasons -join '; ')
    }
}

function New-CandidateContactSheet {
    param(
        [object[]]$Rows,
        [string]$OutputPath,
        [int]$MaxItems
    )

    Add-Type -AssemblyName System.Drawing
    $items = @($Rows | Where-Object { $_.EnhancementCandidate -and -not [string]::IsNullOrWhiteSpace($_.ExtractedPath) -and (Test-Path -LiteralPath $_.ExtractedPath) } |
        Sort-Object PhotoScore -Descending |
        Select-Object -First $MaxItems)
    if ($items.Count -eq 0) { return $false }

    $tileWidth = 280
    $tileHeight = 230
    $imageHeight = 156
    $columns = 4
    if ($items.Count -lt 4) { $columns = [Math]::Max(1, $items.Count) }
    $rowsCount = [int][Math]::Ceiling($items.Count / [double]$columns)
    $sheetWidth = $columns * $tileWidth
    $sheetHeight = $rowsCount * $tileHeight

    $bitmap = $null
    $graphics = $null
    $font = $null
    $smallFont = $null
    $brush = $null
    $mutedBrush = $null
    $pen = $null
    try {
        $bitmap = New-Object System.Drawing.Bitmap($sheetWidth, $sheetHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::FromArgb(248, 248, 248))
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $font = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5)
        $smallFont = New-Object System.Drawing.Font('Microsoft YaHei UI', 7.5)
        $brush = [System.Drawing.Brushes]::Black
        $mutedBrush = [System.Drawing.Brushes]::DimGray
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(210, 210, 210), 1)

        for ($i = 0; $i -lt $items.Count; $i++) {
            $row = [int][Math]::Floor($i / $columns)
            $col = $i % $columns
            $x = $col * $tileWidth
            $y = $row * $tileHeight
            $graphics.FillRectangle([System.Drawing.Brushes]::White, $x + 6, $y + 6, $tileWidth - 12, $tileHeight - 12)
            $graphics.DrawRectangle($pen, $x + 6, $y + 6, $tileWidth - 12, $tileHeight - 12)

            $img = $null
            try {
                $img = [System.Drawing.Image]::FromFile($items[$i].ExtractedPath)
                $maxW = $tileWidth - 24
                $maxH = $imageHeight
                $scale = [Math]::Min($maxW / [double]$img.Width, $maxH / [double]$img.Height)
                $drawW = [Math]::Max(1, [int][Math]::Round($img.Width * $scale))
                $drawH = [Math]::Max(1, [int][Math]::Round($img.Height * $scale))
                $drawX = $x + [int](($tileWidth - $drawW) / 2)
                $drawY = $y + 14 + [int](($imageHeight - $drawH) / 2)
                $graphics.DrawImage($img, $drawX, $drawY, $drawW, $drawH)
            } finally {
                if ($null -ne $img) { $img.Dispose() }
            }

            $labelY = $y + $imageHeight + 18
            $name = [System.IO.Path]::GetFileNameWithoutExtension([string]$items[$i].Deck)
            if ($name.Length -gt 18) { $name = $name.Substring(0, 18) + '...' }
            $line1 = "$($items[$i].CandidateLevel) score=$($items[$i].PhotoScore) $($items[$i].Width)x$($items[$i].Height)"
            $line2 = "$name / $($items[$i].MediaPath)"
            $line3 = "slides: $($items[$i].UsedOnSlides)"
            $graphics.DrawString($line1, $font, $brush, $x + 12, $labelY)
            $graphics.DrawString($line2, $smallFont, $mutedBrush, $x + 12, $labelY + 20)
            $graphics.DrawString($line3, $smallFont, $mutedBrush, $x + 12, $labelY + 38)
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

function Scan-PresentationImages {
    param(
        [System.IO.FileInfo]$File,
        [string]$OutputDir,
        [bool]$ExportCandidates,
        [int]$MinBytes,
        [int]$MinPixels,
        [int]$SampleSize
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.Drawing

    $rows = New-Object System.Collections.Generic.List[object]
    $safeDeck = Convert-ToSafePathSegment -Name ([System.IO.Path]::GetFileNameWithoutExtension($File.Name))
    $extractDir = Join-Path $OutputDir 'candidate-images'
    if ($ExportCandidates -and -not (Test-Path -LiteralPath $extractDir)) {
        New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    }

    $workRoot = Join-Path $OutputDir ('_candidate_scan_' + [Guid]::NewGuid().ToString('N'))
    try {
        New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($File.FullName, $workRoot)
        $slideMap = Get-SlideMediaMap -PptxFile $File
        $videoPosterMap = Get-VideoPosterImageMap -PptxFile $File
        $mediaRoot = Join-Path $workRoot 'ppt\media'
        if (-not (Test-Path -LiteralPath $mediaRoot)) { return @() }

        foreach ($media in @(Get-ChildItem -LiteralPath $mediaRoot -File)) {
            $ext = $media.Extension.ToLowerInvariant()
            if ($ext -notin @('.jpg', '.jpeg', '.png')) { continue }
            $rel = 'ppt/media/' + $media.Name
            if ($videoPosterMap.ContainsKey($rel)) {
                $slides = (@($videoPosterMap[$rel]) | Sort-Object | ForEach-Object { [string]$_ }) -join ','
                $rows.Add([pscustomobject]@{
                    Deck = $File.Name
                    DeckPath = $File.FullName
                    MediaPath = $rel
                    Extension = $ext
                    MediaRole = 'VideoPosterFrame'
                    OriginalBytes = [int64]$media.Length
                    Width = 0
                    Height = 0
                    Pixels = 0
                    PixelFormat = ''
                    UsedOnSlides = $slides
                    UniqueColorBins = 0
                    ColorDiversity = 0
                    GrayRatio = 0
                    TransparentRatio = 0
                    EdgeDensity = 0
                    PhotoScore = 0
                    CandidateLevel = 'VideoPosterFrame'
                    EnhancementCandidate = $false
                    Reason = 'Video poster frame; handled as video media, not a standalone image'
                    ExtractedPath = ''
                }) | Out-Null
                continue
            }

            try {
                $analysis = Get-ImageAnalysis -Path $media.FullName -SampleSize $SampleSize
                $assessment = Get-CandidateAssessment -Extension $ext -Bytes ([int64]$media.Length) -Analysis $analysis -MinBytes $MinBytes -MinPixels $MinPixels
                $slides = if ($slideMap.ContainsKey($rel)) {
                    (@($slideMap[$rel]) | Sort-Object | ForEach-Object { [string]$_ }) -join ','
                } else {
                    ''
                }
                $extractedPath = ''
                if ($ExportCandidates -and $assessment.EnhancementCandidate) {
                    $outName = Convert-ToSafePathSegment -Name ($safeDeck + '__' + $media.Name)
                    $extractedPath = Join-Path $extractDir $outName
                    Copy-Item -LiteralPath $media.FullName -Destination $extractedPath -Force
                }

                $rows.Add([pscustomobject]@{
                    Deck = $File.Name
                    DeckPath = $File.FullName
                    MediaPath = $rel
                    Extension = $ext
                    MediaRole = 'StandaloneImage'
                    OriginalBytes = [int64]$media.Length
                    Width = $analysis.Width
                    Height = $analysis.Height
                    Pixels = $analysis.Pixels
                    PixelFormat = $analysis.PixelFormat
                    UsedOnSlides = $slides
                    UniqueColorBins = $analysis.UniqueColorBins
                    ColorDiversity = $analysis.ColorDiversity
                    GrayRatio = $analysis.GrayRatio
                    TransparentRatio = $analysis.TransparentRatio
                    EdgeDensity = $analysis.EdgeDensity
                    PhotoScore = $assessment.PhotoScore
                    CandidateLevel = $assessment.CandidateLevel
                    EnhancementCandidate = [bool]$assessment.EnhancementCandidate
                    Reason = $assessment.Reason
                    ExtractedPath = $extractedPath
                }) | Out-Null
            } catch {
                $rows.Add([pscustomobject]@{
                    Deck = $File.Name
                    DeckPath = $File.FullName
                    MediaPath = $rel
                    Extension = $ext
                    MediaRole = 'StandaloneImage'
                    OriginalBytes = [int64]$media.Length
                    Width = 0
                    Height = 0
                    Pixels = 0
                    PixelFormat = ''
                    UsedOnSlides = ''
                    UniqueColorBins = 0
                    ColorDiversity = 0
                    GrayRatio = 0
                    TransparentRatio = 0
                    EdgeDensity = 0
                    PhotoScore = 0
                    CandidateLevel = 'Unreadable'
                    EnhancementCandidate = $false
                    Reason = $_.Exception.Message
                    ExtractedPath = ''
                }) | Out-Null
            }
        }
    } finally {
        if (Test-Path -LiteralPath $workRoot) {
            $resolvedWork = [System.IO.Path]::GetFullPath($workRoot)
            $resolvedOutput = [System.IO.Path]::GetFullPath($OutputDir)
            if (Test-PathInsideDirectory -ChildPath $resolvedWork -ParentPath $resolvedOutput) {
                Remove-Item -LiteralPath $workRoot -Recurse -Force
            }
        }
    }

    return $rows.ToArray()
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$files = @(Get-PptxFiles -Path $InputPath -Pattern $FilePattern -Recurse:$Recurse)
if ($files.Count -eq 0) { throw "No PPTX files found in $InputPath" }

$allRows = New-Object System.Collections.Generic.List[object]
foreach ($file in $files) {
    Write-Host "Scanning images: $($file.Name)"
    foreach ($row in @(Scan-PresentationImages -File $file -OutputDir $OutputDir -ExportCandidates ([bool]$ExportCandidates) -MinBytes $MinBytes -MinPixels $MinPixels -SampleSize $SampleSize)) {
        $allRows.Add($row) | Out-Null
    }
}

$csvPath = Join-Path $OutputDir 'pptx-image-candidates.csv'
$jsonPath = Join-Path $OutputDir 'pptx-image-candidates.json'
$manifestPath = Join-Path $OutputDir 'pptx-image-candidates-manifest.json'
$contactSheetPath = Join-Path $OutputDir 'pptx-image-candidates.contact-sheet.png'

$allRows | Sort-Object Deck, @{ Expression = 'EnhancementCandidate'; Descending = $true }, @{ Expression = 'PhotoScore'; Descending = $true }, MediaPath |
    Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
$allRows | Sort-Object Deck, @{ Expression = 'EnhancementCandidate'; Descending = $true }, @{ Expression = 'PhotoScore'; Descending = $true }, MediaPath |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

$contactSheetCreated = $false
if ($ContactSheet) {
    $contactSheetCreated = New-CandidateContactSheet -Rows $allRows.ToArray() -OutputPath $contactSheetPath -MaxItems $MaxContactSheetItems
}

$candidateRows = @($allRows | Where-Object { $_.EnhancementCandidate })
$likelyRows = @($allRows | Where-Object { $_.CandidateLevel -eq 'LikelyPhoto' })
$maybeRows = @($allRows | Where-Object { $_.CandidateLevel -eq 'MaybePhoto' })
$diagramRows = @($allRows | Where-Object { $_.CandidateLevel -eq 'LikelyDiagramOrText' })
$videoPosterRows = @($allRows | Where-Object { $_.CandidateLevel -eq 'VideoPosterFrame' })
$manifest = [pscustomobject]@{
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    input = $InputPath
    outputDir = $OutputDir
    fileCount = $files.Count
    imageCount = $allRows.Count
    enhancementCandidateCount = $candidateRows.Count
    likelyPhotoCount = $likelyRows.Count
    maybePhotoCount = $maybeRows.Count
    likelyDiagramOrTextCount = $diagramRows.Count
    videoPosterFrameCount = $videoPosterRows.Count
    minBytes = $MinBytes
    minPixels = $MinPixels
    sampleSize = $SampleSize
    csv = $csvPath
    json = $jsonPath
    contactSheet = if ($contactSheetCreated) { $contactSheetPath } else { '' }
    note = 'Scanner is heuristic and read-only. Video poster frames are excluded from image enhancement candidates.'
}
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Host "Image candidate scan done: $OutputDir"
Write-Host "Candidates: $($candidateRows.Count) / Images: $($allRows.Count)"
