<#
.SYNOPSIS
  Export per-OLE formula crops from rendered slide pages for GoldSet evidence.

.DESCRIPTION
  Reads formula-carrier-inventory.json, validates the inventory hash against the
  source PPTX, and crops each OLE record's bounding box out of the rendered page
  PNGs produced by Export-PptxVisualAudit (pages/slide-NNN.png). Output is review
  evidence only: crops, a CSV index with sha256 bindings, and a manifest. The
  source PPTX is never modified and no candidate text is produced here.

  SuggestedSizePt is an advisory glyph-height estimate (content pixels after
  near-white trimming, converted through the slide height). It is a starting
  point for the GoldSet SizePt column, never an authority; final sizes are
  decided by judge visual acceptance.

.PARAMETER CarrierInventoryJson
  formula-carrier-inventory.json produced by Export-FormulaCarrierInventory.

.PARAMETER PagesDir
  Directory of rendered page PNGs named slide-NNN.png (Export-PptxVisualAudit pages/).

.PARAMETER OutputDir
  Directory for ole-crops/ PNGs, ole-crops.csv, and the manifest JSON.

.PARAMETER PageImageWidth
  Pixel width of the rendered page images. Defaults to 1600.

.PARAMETER PaddingPx
  Extra margin added around the bbox crop, clamped to the page. Defaults to 4.

.PARAMETER Carrier
  Carrier filter from the inventory. Defaults to MathTypeOle; 'All' keeps every record.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CarrierInventoryJson,

    [Parameter(Mandatory = $true)]
    [string]$PagesDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [ValidateRange(320, 4800)]
    [int]$PageImageWidth = 1600,

    [ValidateRange(0, 64)]
    [int]$PaddingPx = 4,

    [string]$Carrier = 'MathTypeOle'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$script:NsP = 'http://schemas.openxmlformats.org/presentationml/2006/main'
$script:EmuPerPoint = 12700.0

$inventoryPath = [System.IO.Path]::GetFullPath($CarrierInventoryJson)
$pagesFullPath = [System.IO.Path]::GetFullPath($PagesDir)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputDir)
foreach ($path in @($inventoryPath, $pagesFullPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required input not found: $path" }
}
if (-not (Test-Path -LiteralPath $outputFullPath)) { New-Item -ItemType Directory -Path $outputFullPath -Force | Out-Null }

$inventory = Get-Content -LiteralPath $inventoryPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$inventory.schemaVersion -ne 1) { throw "Unsupported carrier inventory schemaVersion: $($inventory.schemaVersion)" }
$inputPath = [System.IO.Path]::GetFullPath([string]$inventory.input.path)
if (-not (Test-Path -LiteralPath $inputPath)) { throw "Inventory source PPTX not found: $inputPath" }
$actualInputSha256 = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualInputSha256 -ne ([string]$inventory.input.sha256).ToLowerInvariant()) {
    throw "Inventory source hash does not match the current PPTX: expected $($inventory.input.sha256), actual $actualInputSha256"
}

function Get-SlideSizeEmu {
    param([string]$PptxPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $zip = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
    try {
        $entry = $zip.GetEntry('ppt/presentation.xml')
        if ($null -eq $entry) { throw 'ppt/presentation.xml missing' }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
        $reader.Dispose()
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($text)
        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace('p', $script:NsP)
        $sldSz = $doc.SelectSingleNode('//p:sldSz', $ns)
        if ($null -eq $sldSz) { throw 'p:sldSz missing from presentation.xml' }
        return [pscustomobject]@{ Cx = [int64]$sldSz.GetAttribute('cx'); Cy = [int64]$sldSz.GetAttribute('cy') }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

function Get-ContentHeightPx {
    # Near-white trim: returns the tight content height in pixels, or 0 when the
    # crop is blank. Advisory input for SuggestedSizePt only.
    param([System.Drawing.Bitmap]$Bitmap)
    $maxX = -1; $minY = $Bitmap.Height; $maxY = -1
    for ($y = 0; $y -lt $Bitmap.Height; $y++) {
        for ($x = 0; $x -lt $Bitmap.Width; $x++) {
            $pixel = $Bitmap.GetPixel($x, $y)
            $brightness = ($pixel.R + $pixel.G + $pixel.B) / 3
            if ($brightness -lt 245 -or $pixel.A -lt 32) {
                if ($y -lt $minY) { $minY = $y }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }
    if ($maxY -lt $minY) { return 0 }
    return ($maxY - $minY + 1)
}

$slideSize = Get-SlideSizeEmu -PptxPath $inputPath
$slideWidthPt = $slideSize.Cx / $script:EmuPerPoint
$slideHeightPt = $slideSize.Cy / $script:EmuPerPoint
$pageWidthPx = [double]$PageImageWidth
$pageHeightPx = [Math]::Round($pageWidthPx * $slideSize.Cy / $slideSize.Cx)
$ptPerPx = $slideHeightPt / $pageHeightPx

$allRecords = @($inventory.records)
if ($Carrier -ne 'All') {
    $allRecords = @($allRecords | Where-Object { [string]$_.source.carrier -eq $Carrier })
}

$cropDir = Join-Path $outputFullPath 'ole-crops'
if (-not (Test-Path -LiteralPath $cropDir)) { New-Item -ItemType Directory -Path $cropDir -Force | Out-Null }

$rows = New-Object System.Collections.Generic.List[object]
$cropHashes = New-Object System.Collections.Generic.List[string]
$missingPages = New-Object System.Collections.Generic.List[string]
Add-Type -AssemblyName System.Drawing | Out-Null

foreach ($record in $allRecords) {
    $slideNo = [int]$record.source.slide
    $pageName = 'slide-{0:d3}.png' -f $slideNo
    $pagePath = Join-Path $pagesFullPath $pageName
    $recordId = [string]$record.recordId
    if (-not (Test-Path -LiteralPath $pagePath)) {
        $missingPages.Add("$recordId -> $pageName") | Out-Null
        continue
    }
    $bbox = $record.source.bbox
    if ($null -eq $bbox -or $null -eq $bbox.left -or $null -eq $bbox.width) {
        $missingPages.Add("$recordId -> bbox missing") | Out-Null
        continue
    }
    $leftPx = [Math]::Round([double]$bbox.left * $pageWidthPx / $slideWidthPt) - $PaddingPx
    $topPx = [Math]::Round([double]$bbox.top * $pageHeightPx / $slideHeightPt) - $PaddingPx
    $rightPx = [Math]::Round(([double]$bbox.left + [double]$bbox.width) * $pageWidthPx / $slideWidthPt) + $PaddingPx
    $bottomPx = [Math]::Round(([double]$bbox.top + [double]$bbox.height) * $pageHeightPx / $slideHeightPt) + $PaddingPx
    $leftPx = [Math]::Max(0, $leftPx); $topPx = [Math]::Max(0, $topPx)
    $rightPx = [Math]::Min([int]$pageWidthPx, $rightPx); $bottomPx = [Math]::Min([int]$pageHeightPx, $bottomPx)
    $cropW = $rightPx - $leftPx; $cropH = $bottomPx - $topPx
    if ($cropW -le 0 -or $cropH -le 0) {
        $missingPages.Add("$recordId -> empty crop rect") | Out-Null
        continue
    }

    $pageImage = [System.Drawing.Bitmap]::new($pagePath)
    try {
        $cropRect = New-Object System.Drawing.Rectangle($leftPx, $topPx, $cropW, $cropH)
        $cropped = $pageImage.Clone($cropRect, $pageImage.PixelFormat)
        $suggestedPt = 0
        try {
            $contentPx = Get-ContentHeightPx -Bitmap $cropped
            if ($contentPx -gt 0) {
                $suggestedPt = [Math]::Max(8, [Math]::Min(96, [int][Math]::Round($contentPx * $ptPerPx)))
            }
        } catch { $suggestedPt = 0 }
        $cropName = '{0}_slide-{1:d3}.png' -f $recordId, $slideNo
        $cropPath = Join-Path $cropDir $cropName
        $cropped.Save($cropPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $cropped.Dispose()
    } finally {
        $pageImage.Dispose()
    }
    $cropSha256 = (Get-FileHash -LiteralPath $cropPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $pageSha256 = (Get-FileHash -LiteralPath $pagePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $cropHashes.Add($cropSha256) | Out-Null
    $rows.Add([pscustomobject]@{
        RecordId = $recordId
        Carrier = [string]$record.source.carrier
        Slide = $slideNo
        SlideKind = if ($null -ne $record.source.PSObject.Properties['slideKind']) { [string]$record.source.slideKind } else { '' }
        ShapeId = [int]$record.source.shapeId
        ShapeName = [string]$record.source.shapeName
        OleProgId = [string]$record.source.oleProgId
        BboxLeftPt = [Math]::Round([double]$bbox.left, 2)
        BboxTopPt = [Math]::Round([double]$bbox.top, 2)
        BboxWidthPt = [Math]::Round([double]$bbox.width, 2)
        BboxHeightPt = [Math]::Round([double]$bbox.height, 2)
        CropRectPx = "$leftPx,$topPx,$cropW,$cropH"
        SuggestedSizePt = $suggestedPt
        CropPng = $cropPath
        CropSha256 = $cropSha256
        PagePng = $pagePath
        PageSha256 = $pageSha256
    }) | Out-Null
}

$csvPath = Join-Path $outputFullPath 'ole-crops.csv'
if ($rows.Count -gt 0) {
    Write-Utf8BomCsv -InputObject $rows.ToArray() -Path $csvPath
} else {
    Write-Utf8BomText -Text "RecordId,Carrier,Slide,SlideKind,ShapeId,ShapeName,OleProgId,BboxLeftPt,BboxTopPt,BboxWidthPt,BboxHeightPt,CropRectPx,SuggestedSizePt,CropPng,CropSha256,PagePng,PageSha256`r`n" -Path $csvPath
}

$cropHashInput = ($cropHashes | Sort-Object) -join '|'
$shaForEvidenceSet = [System.Security.Cryptography.SHA256]::Create()
try {
    $evidenceBytes = [System.Text.Encoding]::UTF8.GetBytes($cropHashInput)
    $evidenceSetHash = ([System.BitConverter]::ToString($shaForEvidenceSet.ComputeHash($evidenceBytes)) -replace '-', '').ToLowerInvariant()
} finally {
    $shaForEvidenceSet.Dispose()
}

$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    input = [ordered]@{ path = $inputPath; sha256 = $actualInputSha256 }
    inventory = [ordered]@{ path = $inventoryPath; sha256 = (Get-FileHash -LiteralPath $inventoryPath -Algorithm SHA256).Hash.ToLowerInvariant() }
    pages = [ordered]@{ dir = $pagesFullPath; imageWidthPx = $PageImageWidth; slideSizeEmu = "$($slideSize.Cx)x$($slideSize.Cy)" }
    carrier = $Carrier
    recordCount = $allRecords.Count
    croppedCount = $rows.Count
    missingCount = $missingPages.Count
    missing = @($missingPages.ToArray())
    evidenceSetSha256 = $evidenceSetHash
    csv = $csvPath
    cropDir = $cropDir
    writeBackAllowed = $false
    note = 'Crops are GoldSet evidence only. SuggestedSizePt is an advisory glyph-height estimate; final SizePt comes from review and judge visual acceptance.'
}
$manifestPath = Join-Path $outputFullPath 'formula-ole-crops-manifest.json'
Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 8) -Path $manifestPath

Write-Output ("OLE crops done: {0}`nCropped: {1} / Records: {2}; missing: {3}; manifest: {4}" -f $cropDir, $rows.Count, $allRecords.Count, $missingPages.Count, $manifestPath)
