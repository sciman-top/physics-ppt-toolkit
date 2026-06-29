<#
.SYNOPSIS
  Optimize embedded PPTX media with conservative, reversible rules.

.DESCRIPTION
  Creates .media-optimized.pptx copies and media optimization reports.
  The script never overwrites input PPTX files. By default it only recompresses
  JPEG/JPG media with a conservative quality setting. PNG and video files are
  reported but not modified unless a future tool-specific path is enabled.

.PARAMETER InputPath
  Path to a .pptx file or directory containing PPTX files.

.PARAMETER OutputDir
  Directory where optimized PPTX copies and reports are written.

.PARAMETER Recurse
  Search subdirectories when InputPath is a directory.

.PARAMETER FilePattern
  File name pattern used when InputPath is a directory.

.PARAMETER MinBytes
  Minimum embedded media size eligible for optimization.

.PARAMETER MinSavingsPercent
  Minimum percentage reduction required before replacing a media file.

.PARAMETER JpegQuality
  JPEG re-encoding quality, 1-100.

.PARAMETER SharpenJpeg
  Apply a mild 3x3 sharpen filter before JPEG re-encoding. Off by default.

.PARAMETER UseSharp
  Use the optional open-source sharp/libvips path when Node.js and sharp are
  available. This enables conservative PNG recompression and faster JPEG
  recompression. If unavailable, the script falls back to built-in behavior.

.PARAMETER NodePath
  Optional path to node.exe for the sharp/libvips path.

.PARAMETER NodeModulesPath
  Optional node_modules path used to resolve sharp.

.PARAMETER ToolRoot
  Optional folder scanned recursively for portable external tools.

.PARAMETER Force
  Recreate outputs even when they already exist.
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
    [int]$MinBytes = 204800,

    [ValidateRange(0, 100)]
    [double]$MinSavingsPercent = 5,

    [ValidateRange(1, 100)]
    [int]$JpegQuality = 90,

    [switch]$SharpenJpeg,

    [switch]$UseSharp,

    [string]$NodePath = '',

    [string]$NodeModulesPath = '',

    [string]$ToolRoot = '',

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

if ([string]::IsNullOrWhiteSpace($ToolRoot)) {
    $ToolRoot = Join-Path $PSScriptRoot 'vendor'
}

function Get-PptxFiles {
    param([string]$Path, [string]$Pattern, [switch]$Recurse)
    if (-not (Test-Path -LiteralPath $Path)) { throw "InputPath not found: $Path" }
    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) {
        $opt = @{ LiteralPath = $item.FullName; Filter = $Pattern; File = $true }
        if ($Recurse) { $opt.Recurse = $true }
        return @(Get-ChildItem @opt | Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '*.media-optimized.pptx' })
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

function Get-ImageInfo {
    param([string]$Path)
    $image = $null
    try {
        Add-Type -AssemblyName System.Drawing
        $image = [System.Drawing.Image]::FromFile($Path)
        return [pscustomobject]@{
            Width = [int]$image.Width
            Height = [int]$image.Height
            Format = $image.RawFormat.ToString()
        }
    } catch {
        return [pscustomobject]@{
            Width = 0
            Height = 0
            Format = ''
        }
    } finally {
        if ($null -ne $image) { $image.Dispose() }
    }
}

function Save-Jpeg {
    param(
        [System.Drawing.Image]$Image,
        [string]$Path,
        [int]$Quality
    )

    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
        Where-Object { $_.MimeType -eq 'image/jpeg' } |
        Select-Object -First 1
    if ($null -eq $codec) { throw 'JPEG codec not available.' }

    $encoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters 1
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter $encoder, ([int64]$Quality)
    try {
        $Image.Save($Path, $codec, $encoderParams)
    } finally {
        $encoderParams.Dispose()
    }
}

function New-SharpenedBitmap {
    param([System.Drawing.Bitmap]$Source)

    $width = $Source.Width
    $height = $Source.Height
    $target = New-Object System.Drawing.Bitmap $width, $height
    for ($y = 0; $y -lt $height; $y++) {
        for ($x = 0; $x -lt $width; $x++) {
            if ($x -eq 0 -or $y -eq 0 -or $x -eq ($width - 1) -or $y -eq ($height - 1)) {
                $target.SetPixel($x, $y, $Source.GetPixel($x, $y))
                continue
            }

            $center = $Source.GetPixel($x, $y)
            $left = $Source.GetPixel($x - 1, $y)
            $right = $Source.GetPixel($x + 1, $y)
            $top = $Source.GetPixel($x, $y - 1)
            $bottom = $Source.GetPixel($x, $y + 1)

            $r = [Math]::Min(255, [Math]::Max(0, (5 * $center.R) - $left.R - $right.R - $top.R - $bottom.R))
            $g = [Math]::Min(255, [Math]::Max(0, (5 * $center.G) - $left.G - $right.G - $top.G - $bottom.G))
            $b = [Math]::Min(255, [Math]::Max(0, (5 * $center.B) - $left.B - $right.B - $top.B - $bottom.B))
            $target.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($center.A, $r, $g, $b))
        }
    }
    return $target
}

function Optimize-JpegMedia {
    param(
        [string]$Path,
        [int]$Quality,
        [switch]$Sharpen
    )

    Add-Type -AssemblyName System.Drawing
    $source = $null
    $bitmap = $null
    $working = $null
    $temp = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.jpg')
    try {
        $source = [System.Drawing.Image]::FromFile($Path)
        $bitmap = New-Object System.Drawing.Bitmap $source
        if ($Sharpen) {
            $working = New-SharpenedBitmap -Source $bitmap
        } else {
            $working = $bitmap
        }
        Save-Jpeg -Image $working -Path $temp -Quality $Quality
        return $temp
    } catch {
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
        throw
    } finally {
        if ($null -ne $working -and -not [object]::ReferenceEquals($working, $bitmap)) { $working.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
        if ($null -ne $source) { $source.Dispose() }
    }
}

function Resolve-NodeExecutable {
    param([string]$Candidate)

    if (-not [string]::IsNullOrWhiteSpace($Candidate)) {
        if (-not (Test-Path -LiteralPath $Candidate)) { throw "NodePath not found: $Candidate" }
        return (Get-Item -LiteralPath $Candidate).FullName
    }

    $cmd = Get-Command node -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $cmd) { return $cmd.Source }
    return ''
}

function Invoke-WithNodeModulePath {
    param(
        [string]$ModulePath,
        [scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )

    $oldNodePath = $env:NODE_PATH
    try {
        if (-not [string]::IsNullOrWhiteSpace($ModulePath)) {
            if (-not (Test-Path -LiteralPath $ModulePath)) { throw "NodeModulesPath not found: $ModulePath" }
            $moduleEntries = New-Object System.Collections.Generic.List[string]
            $moduleEntries.Add($ModulePath) | Out-Null
            $pnpmModulePath = Join-Path $ModulePath '.pnpm\node_modules'
            if (Test-Path -LiteralPath $pnpmModulePath) { $moduleEntries.Add($pnpmModulePath) | Out-Null }
            $combinedModulePath = ($moduleEntries -join [System.IO.Path]::PathSeparator)
            if ([string]::IsNullOrWhiteSpace($oldNodePath)) {
                $env:NODE_PATH = $combinedModulePath
            } else {
                $env:NODE_PATH = $combinedModulePath + [System.IO.Path]::PathSeparator + $oldNodePath
            }
        }
        & $ScriptBlock @ArgumentList
    } finally {
        $env:NODE_PATH = $oldNodePath
    }
}

function Test-SharpAvailable {
    param(
        [string]$ResolvedNodePath,
        [string]$ModulePath
    )

    if ([string]::IsNullOrWhiteSpace($ResolvedNodePath)) { return $false }
    try {
        Invoke-WithNodeModulePath -ModulePath $ModulePath -ArgumentList @($ResolvedNodePath) -ScriptBlock {
            param([string]$NodeExe)
            $null = & $NodeExe -e "require('sharp'); process.stdout.write('ok')" 2>$null
            $LASTEXITCODE -eq 0
        }
    } catch {
        return $false
    }
}

function Invoke-SharpMedia {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Kind,
        [int]$Quality,
        [switch]$Sharpen,
        [string]$ResolvedNodePath,
        [string]$ModulePath,
        [string]$WorkerPath
    )

    if (-not (Test-Path -LiteralPath $WorkerPath)) { throw "Sharp worker not found: $WorkerPath" }

    $argsPath = [System.IO.Path]::GetTempFileName()
    try {
        [pscustomobject]@{
            input = $InputFile
            output = $OutputFile
            kind = $Kind
            quality = $Quality
            sharpen = [bool]$Sharpen
        } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $argsPath -Encoding UTF8

        $raw = Invoke-WithNodeModulePath -ModulePath $ModulePath -ArgumentList @($ResolvedNodePath, $WorkerPath, $argsPath) -ScriptBlock {
            param([string]$NodeExe, [string]$Worker, [string]$ArgsFile)
            & $NodeExe $Worker $ArgsFile 2>&1
        }
        if ($LASTEXITCODE -ne 0) {
            throw ("sharp/libvips failed: " + (($raw | Out-String).Trim()))
        }
        return (($raw | Out-String).Trim() | ConvertFrom-Json)
    } finally {
        if (Test-Path -LiteralPath $argsPath) { Remove-Item -LiteralPath $argsPath -Force }
    }
}

function Invoke-OxipngMedia {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$ToolPath
    )

    $raw = & $ToolPath -o 4 --strip safe --out $OutputFile $InputFile 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ("oxipng failed: " + (($raw | Out-String).Trim()))
    }
}

function Resolve-ExternalTool {
    param(
        [string]$Name,
        [string]$Root
    )

    $cmd = Get-Command $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $cmd) { return $cmd.Source }

    if (-not [string]::IsNullOrWhiteSpace($Root) -and (Test-Path -LiteralPath $Root)) {
        $exeName = if ($Name.EndsWith('.exe', [System.StringComparison]::OrdinalIgnoreCase)) { $Name } else { "$Name.exe" }
        $match = Get-ChildItem -LiteralPath $Root -Recurse -File -Filter $exeName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $match) { return $match.FullName }
    }

    return ''
}

function Get-ExternalToolMap {
    param(
        [string]$ResolvedNodePath,
        [string]$ModulePath,
        [bool]$SharpAvailable,
        [string]$ToolRoot
    )

    $names = @('oxipng', 'pngquant', 'magick', 'cjpeg', 'jpegtran', 'realesrgan-ncnn-vulkan', 'ffmpeg')
    $map = @{}
    foreach ($name in $names) {
        $toolPath = Resolve-ExternalTool -Name $name -Root $ToolRoot
        if (-not [string]::IsNullOrWhiteSpace($toolPath)) { $map[$name] = $toolPath }
    }
    if (-not [string]::IsNullOrWhiteSpace($ResolvedNodePath)) { $map['node'] = $ResolvedNodePath }
    if (-not [string]::IsNullOrWhiteSpace($ModulePath)) { $map['node_modules'] = $ModulePath }
    if ($SharpAvailable) { $map['sharp'] = 'available via Node.js' }
    return $map
}

function Add-ReportRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$File,
        [string]$MediaPath,
        [string]$Extension,
        [int64]$OriginalBytes,
        [int64]$OptimizedBytes,
        [string]$Action,
        [string]$Reason,
        [string]$Tool,
        [int]$Width,
        [int]$Height
    )

    $savedBytes = [Math]::Max(0, $OriginalBytes - $OptimizedBytes)
    $savedPercent = if ($OriginalBytes -gt 0) { [Math]::Round(($savedBytes / $OriginalBytes) * 100, 2) } else { 0 }
    $Rows.Add([pscustomobject]@{
        File = $File
        MediaPath = $MediaPath
        Extension = $Extension
        OriginalBytes = $OriginalBytes
        OptimizedBytes = $OptimizedBytes
        SavedBytes = $savedBytes
        SavedPercent = $savedPercent
        Action = $Action
        Reason = $Reason
        Tool = $Tool
        Width = $Width
        Height = $Height
    }) | Out-Null
}

function Optimize-PresentationMedia {
    param(
        [System.IO.FileInfo]$File,
        [string]$OutputDir,
        [hashtable]$Tools,
        [bool]$SharpAvailable,
        [string]$ResolvedNodePath,
        [string]$ResolvedNodeModulesPath,
        [string]$SharpWorkerPath,
        [int]$MinBytes,
        [double]$MinSavingsPercent,
        [int]$JpegQuality,
        [bool]$SharpenJpeg,
        [bool]$UseSharp,
        [bool]$Force
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.Drawing

    $safeName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $outPptx = Join-Path $OutputDir ($safeName + '.media-optimized.pptx')
    $reportPath = Join-Path $OutputDir ($safeName + '.media-optimization-report.csv')
    $manifestPath = Join-Path $OutputDir ($safeName + '.media-optimization.json')

    if ((Test-Path -LiteralPath $outPptx) -and -not $Force) {
        Write-Host "Skip existing: $outPptx"
        return
    }

    $workRoot = Join-Path $OutputDir ('_media_work_' + [Guid]::NewGuid().ToString('N'))
    $rows = New-Object System.Collections.Generic.List[object]
    $status = 'success'
    $message = ''

    try {
        New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($File.FullName, $workRoot)

        $mediaRoot = Join-Path $workRoot 'ppt\media'
        $mediaFiles = @()
        if (Test-Path -LiteralPath $mediaRoot) {
            $mediaFiles = @(Get-ChildItem -LiteralPath $mediaRoot -File)
        }
        $videoPosterMap = Get-VideoPosterImageMap -PptxFile $File

        foreach ($media in $mediaFiles) {
            $ext = $media.Extension.ToLowerInvariant()
            $rel = ('ppt/media/' + $media.Name)
            $originalBytes = [int64]$media.Length

            if ($videoPosterMap.ContainsKey($rel)) {
                Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $originalBytes -Action 'Skipped' -Reason 'Video poster frame is not a standalone image' -Tool 'none' -Width 0 -Height 0
                continue
            }

            $info = Get-ImageInfo -Path $media.FullName

            if ($originalBytes -lt $MinBytes) {
                Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $originalBytes -Action 'Skipped' -Reason "Below MinBytes $MinBytes" -Tool 'none' -Width $info.Width -Height $info.Height
                continue
            }

            if ($ext -in @('.jpg', '.jpeg')) {
                $temp = $null
                try {
                    $toolName = 'System.Drawing JPEG'
                    if ($UseSharp -and $SharpAvailable) {
                        $temp = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.jpg')
                        $null = Invoke-SharpMedia -InputFile $media.FullName -OutputFile $temp -Kind 'jpeg' -Quality $JpegQuality -Sharpen:$SharpenJpeg -ResolvedNodePath $ResolvedNodePath -ModulePath $ResolvedNodeModulesPath -WorkerPath $SharpWorkerPath
                        $toolName = 'sharp/libvips'
                    } else {
                        $temp = Optimize-JpegMedia -Path $media.FullName -Quality $JpegQuality -Sharpen:$SharpenJpeg
                    }
                    $newBytes = (Get-Item -LiteralPath $temp).Length
                    $savedPercent = if ($originalBytes -gt 0) { (($originalBytes - $newBytes) / $originalBytes) * 100 } else { 0 }
                    if ($newBytes -lt $originalBytes -and $savedPercent -ge $MinSavingsPercent) {
                        Copy-Item -LiteralPath $temp -Destination $media.FullName -Force
                        Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $newBytes -Action 'Replaced' -Reason 'JPEG recompressed and savings threshold met' -Tool $toolName -Width $info.Width -Height $info.Height
                    } else {
                        Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $newBytes -Action 'Kept' -Reason 'Savings below threshold' -Tool $toolName -Width $info.Width -Height $info.Height
                    }
                } catch {
                    Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $originalBytes -Action 'Failed' -Reason $_.Exception.Message -Tool 'JPEG optimizer' -Width $info.Width -Height $info.Height
                } finally {
                    if ($null -ne $temp -and (Test-Path -LiteralPath $temp)) { Remove-Item -LiteralPath $temp -Force }
                }
                continue
            }

            if ($ext -eq '.png') {
                if ($Tools.ContainsKey('oxipng') -or ($UseSharp -and $SharpAvailable)) {
                    $temp = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.png')
                    try {
                        $toolName = 'sharp/libvips'
                        if ($Tools.ContainsKey('oxipng')) {
                            Invoke-OxipngMedia -InputFile $media.FullName -OutputFile $temp -ToolPath $Tools['oxipng']
                            $toolName = 'oxipng'
                        } else {
                            $null = Invoke-SharpMedia -InputFile $media.FullName -OutputFile $temp -Kind 'png' -Quality $JpegQuality -ResolvedNodePath $ResolvedNodePath -ModulePath $ResolvedNodeModulesPath -WorkerPath $SharpWorkerPath
                        }
                        $newBytes = (Get-Item -LiteralPath $temp).Length
                        $savedPercent = if ($originalBytes -gt 0) { (($originalBytes - $newBytes) / $originalBytes) * 100 } else { 0 }
                        if ($newBytes -lt $originalBytes -and $savedPercent -ge $MinSavingsPercent) {
                            Copy-Item -LiteralPath $temp -Destination $media.FullName -Force
                            Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $newBytes -Action 'Replaced' -Reason 'PNG recompressed and savings threshold met' -Tool $toolName -Width $info.Width -Height $info.Height
                        } else {
                            Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $newBytes -Action 'Kept' -Reason 'Savings below threshold' -Tool $toolName -Width $info.Width -Height $info.Height
                        }
                    } catch {
                        Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $originalBytes -Action 'Failed' -Reason $_.Exception.Message -Tool 'PNG optimizer' -Width $info.Width -Height $info.Height
                    } finally {
                        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
                    }
                } else {
                    Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $originalBytes -Action 'Skipped' -Reason 'PNG optimization requires an external lossless optimizer such as oxipng or sharp/libvips' -Tool 'none' -Width $info.Width -Height $info.Height
                }
                continue
            }

            Add-ReportRow -Rows $rows -File $File.Name -MediaPath $rel -Extension $ext -OriginalBytes $originalBytes -OptimizedBytes $originalBytes -Action 'Skipped' -Reason 'Unsupported or high-risk media type' -Tool 'none' -Width $info.Width -Height $info.Height
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

    $rows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8

    $originalBytesTotal = [int64]$File.Length
    $optimizedBytesTotal = if (Test-Path -LiteralPath $outPptx) { [int64](Get-Item -LiteralPath $outPptx).Length } else { 0 }
    $savedBytesTotal = [Math]::Max(0, $originalBytesTotal - $optimizedBytesTotal)
    $savedPercentTotal = if ($originalBytesTotal -gt 0) { [Math]::Round(($savedBytesTotal / $originalBytesTotal) * 100, 2) } else { 0 }
    $replacedCount = @($rows | Where-Object { $_.Action -eq 'Replaced' }).Count
    $failedCount = @($rows | Where-Object { $_.Action -eq 'Failed' }).Count

    $manifest = [pscustomobject]@{
        generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        input = $File.FullName
        outputPptx = $outPptx
        report = $reportPath
        status = $status
        message = $message
        originalBytes = $originalBytesTotal
        optimizedBytes = $optimizedBytesTotal
        savedBytes = $savedBytesTotal
        savedPercent = $savedPercentTotal
        mediaCount = $rows.Count
        replacedCount = $replacedCount
        failedCount = $failedCount
        jpegQuality = $JpegQuality
        sharpenJpeg = [bool]$SharpenJpeg
        useSharp = [bool]$UseSharp
        sharpAvailable = [bool]$SharpAvailable
        minBytes = $MinBytes
        minSavingsPercent = $MinSavingsPercent
        externalTools = $Tools
    }
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
$ToolRoot = if ([string]::IsNullOrWhiteSpace($ToolRoot)) { '' } else { [System.IO.Path]::GetFullPath($ToolRoot) }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$resolvedNodePath = Resolve-NodeExecutable -Candidate $NodePath
$resolvedNodeModulesPath = ''
if (-not [string]::IsNullOrWhiteSpace($NodeModulesPath)) {
    if (-not (Test-Path -LiteralPath $NodeModulesPath)) { throw "NodeModulesPath not found: $NodeModulesPath" }
    $resolvedNodeModulesPath = (Get-Item -LiteralPath $NodeModulesPath).FullName
}
$sharpAvailable = Test-SharpAvailable -ResolvedNodePath $resolvedNodePath -ModulePath $resolvedNodeModulesPath
if ($UseSharp -and -not $sharpAvailable) {
    Write-Warning 'UseSharp was requested, but sharp/libvips is not available. Falling back to built-in JPEG behavior and PNG reporting.'
}
$sharpWorkerPath = Join-Path $PSScriptRoot 'Optimize-PptxMedia.worker.js'
$tools = Get-ExternalToolMap -ResolvedNodePath $resolvedNodePath -ModulePath $resolvedNodeModulesPath -SharpAvailable $sharpAvailable -ToolRoot $ToolRoot
$files = @(Get-PptxFiles -Path $InputPath -Pattern $FilePattern -Recurse:$Recurse)
if ($files.Count -eq 0) { throw "No .pptx files found in $InputPath" }

foreach ($file in $files) {
    Write-Host "Optimizing media: $($file.Name)"
    Optimize-PresentationMedia -File $file -OutputDir $OutputDir -Tools $tools -SharpAvailable $sharpAvailable -ResolvedNodePath $resolvedNodePath -ResolvedNodeModulesPath $resolvedNodeModulesPath -SharpWorkerPath $sharpWorkerPath -MinBytes $MinBytes -MinSavingsPercent $MinSavingsPercent -JpegQuality $JpegQuality -SharpenJpeg ([bool]$SharpenJpeg) -UseSharp ([bool]$UseSharp) -Force ([bool]$Force)
}

Write-Host "Media optimization done: $OutputDir"
