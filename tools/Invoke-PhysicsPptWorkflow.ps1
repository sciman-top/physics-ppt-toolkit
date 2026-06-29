<#
.SYNOPSIS
  One-command workflow for checking, normalizing, exporting PDF, and summarizing physics PPT files.

.DESCRIPTION
  Creates a timestamped output workspace, runs toolkit self-check, optionally runs report-only inspection,
  normalizes PPTX/PPTM files, exports same-name PDFs, and writes summary.md plus review-manifest.json.
  SafeNormalize can be used explicitly when PDF/page image export is not needed.
  Detailed visual review artifacts can be generated explicitly with -IncludeReviewArtifacts.
  Rule-based rendered visual audits and deterministic low-risk layout fixes can be enabled with
  -IncludeVisualAudit and -ApplyVisualAuditFixes.

.PARAMETER InputPath
  Path to a .pptx/.pptm file or a directory containing PPT files.

.PARAMETER OutputRoot
  Optional output root. Defaults to a timestamped folder next to InputPath.

.PARAMETER Mode
  CheckOnly: only generate report.
  SafeNormalize: generate normalized PPTX without PDF.
  NormalizeAndPdf: generate normalized PPTX plus same-name PDF.
  ForceRebuild: regenerate outputs even when existing files are newer than input.

.PARAMETER Recurse
  Search subdirectories when InputPath is a directory.

.PARAMETER UpdateMaster
  Also normalize the slide master text styles.

.PARAMETER OpenOutput
  Open the output folder when the workflow completes.

.PARAMETER OpenGeneratedPptx
  Open the generated PPTX when exactly one file was processed; otherwise open the generated PPTX folder.

.PARAMETER IncludeReviewArtifacts
  Also export page images, contact sheets, before/after sheets, and review indexes for detailed visual review.

.PARAMETER IncludeVisualAudit
  Also export rendered page images/PDF and run rule-based visual/layout checks on generated PPTX files.

.PARAMETER ApplyVisualAuditFixes
  Copy generated PPTX files and apply only deterministic layout fixes from the visual audit.

.PARAMETER ApplyFormulaOmmlWhitelist
  Generate editable OfficeMath/OMML copies for formula whitelist hits, then validate them with Open XML.

.PARAMETER FormulaOmmlVisualAudit
  Also run the slow rendered visual audit on OMML copies. The default fast formula path uses Open XML validation only.

.PARAMETER SkipPreflightReport
  Skip the initial report-only pass for NormalizeAndPdf/SafeNormalize runs. The normalize pass still writes the final report.

.PARAMETER FormulaOmmlMaxItems
  Maximum formula whitelist rows to convert per run.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,

    [string]$OutputRoot,

    [ValidateSet('CheckOnly', 'SafeNormalize', 'NormalizeAndPdf', 'ForceRebuild')]
    [string]$Mode = 'NormalizeAndPdf',

    [string]$FilePattern = '*.ppt*',

    [switch]$Recurse,
    [switch]$UpdateMaster,
    [switch]$OpenOutput,
    [switch]$OpenGeneratedPptx,
    [switch]$IncludeReviewArtifacts,
    [switch]$IncludeVisualAudit,
    [switch]$ApplyVisualAuditFixes,
    [switch]$ApplyFormulaOmmlWhitelist,
    [switch]$FormulaOmmlVisualAudit,
    [switch]$SkipPreflightReport,

    [ValidateRange(1, 1000)]
    [int]$FormulaOmmlMaxItems = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

if ($ApplyVisualAuditFixes -and -not $IncludeVisualAudit) {
    $IncludeVisualAudit = [System.Management.Automation.SwitchParameter]::Present
}

function Get-PptFiles {
    param([string]$Path, [string]$Pattern, [switch]$Recurse)
    if (-not (Test-Path -LiteralPath $Path)) { throw "InputPath not found: $Path" }
    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) {
        $opt = @{ LiteralPath = $item.FullName; Filter = $Pattern; File = $true }
        if ($Recurse) { $opt.Recurse = $true }
        return @(Get-ChildItem @opt | Where-Object { $_.Name -notlike '~$*' -and $_.Extension -in '.pptx', '.pptm' })
    }
    if ($item.Extension -notin '.pptx', '.pptm') { throw "Only .pptx/.pptm files are supported: $($item.FullName)" }
    return @($item)
}

function Get-PowerPointAutomationBlockers {
    try {
        $processes = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
        return @($processes | Where-Object {
            $_.MainWindowTitle -like '*受保护的视图*' -or
            $_.MainWindowTitle -like '*Protected View*'
        } | ForEach-Object {
            [pscustomobject]@{
                Id = $_.Id
                Title = $_.MainWindowTitle
            }
        })
    } catch {
        return @()
    }
}

function Assert-PowerPointAutomationReady {
    $blockers = @(Get-PowerPointAutomationBlockers)
    if ($blockers.Count -eq 0) { return }

    $details = ($blockers | ForEach-Object { "PID $($_.Id): $($_.Title)" }) -join '; '
    throw "PowerPoint automation is blocked by an existing Protected View window. Close it and rerun. Blockers: $details"
}

function Get-DefaultOutputRoot {
    param([System.IO.FileSystemInfo]$InputItem)
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseDir = if ($InputItem.PSIsContainer) { $InputItem.FullName } else { $InputItem.DirectoryName }
    return Join-Path $baseDir ("_physics_ppt_output_$stamp")
}

function Convert-ReportCsv {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    return @(Import-Csv -LiteralPath $Path -Encoding UTF8)
}

function Get-IssueCount {
    param($Rows, [string]$Issue)
    return @($Rows | Where-Object { $_.Issue -eq $Issue }).Count
}

function Get-VideoCandidateCount {
    param($Rows)
    return @($Rows | Where-Object { $_.Issue -eq 'SlideType' -and $_.Details -eq 'VideoOrMediaCandidate' }).Count
}

function Get-SmallTextCount {
    param($Rows)
    return @($Rows | Where-Object { $_.Issue -in @('SmallText', 'SmallTextPreserved', 'SmallTextAfterNormalize') }).Count
}

function Get-ImageExportMismatchCount {
    param($Rows)
    return @($Rows | Where-Object { $_.Issue -eq 'ImageExportCountMismatch' }).Count
}

function Get-FileIdentity {
    param(
        [string]$InputRoot,
        [System.IO.FileSystemInfo]$FileItem
    )

    $fullPath = [System.IO.Path]::GetFullPath($FileItem.FullName)
    $displayName = $FileItem.Name
    $relativePath = $displayName

    if (-not [string]::IsNullOrWhiteSpace($InputRoot)) {
        $rootFull = [System.IO.Path]::GetFullPath($InputRoot)
        $parentDir = Split-Path -Parent $fullPath
        if ($parentDir.Equals($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            $relativePath = $displayName
        } elseif (Test-PathInsideDirectory -ChildPath $fullPath -ParentPath $rootFull) {
            $relativePath = $fullPath.Substring($rootFull.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
        }
    }

    return [pscustomobject]@{
        key = $fullPath
        displayName = $displayName
        relativePath = $relativePath -replace '\\', '/'
        safeStem = Get-RelativePathSafeStem -RootPath $InputRoot -TargetPath $fullPath
    }
}

function Get-IdentityMap {
    param(
        [string]$InputRoot,
        [System.IO.FileInfo[]]$Files
    )

    $map = @{}
    foreach ($file in @($Files)) {
        $map[$file.FullName] = Get-FileIdentity -InputRoot $InputRoot -FileItem $file
    }
    return $map
}

function Get-ReportRowFileKey {
    param($Row)

    if ($null -eq $Row) { return '' }
    $pathProp = $Row.PSObject.Properties['FilePath']
    if ($null -ne $pathProp -and -not [string]::IsNullOrWhiteSpace([string]$pathProp.Value)) {
        return [string]$pathProp.Value
    }
    $fileProp = $Row.PSObject.Properties['File']
    if ($null -ne $fileProp -and -not [string]::IsNullOrWhiteSpace([string]$fileProp.Value)) {
        return [string]$fileProp.Value
    }
    return ''
}

function New-ImageContactSheet {
    param(
        [string]$ImageDir,
        [string]$OutputPath,
        [int]$Columns = 4,
        [int]$ThumbWidth = 240,
        [int]$ThumbHeight = 135
    )

    if (-not (Test-Path -LiteralPath $ImageDir)) { return $null }
    $files = @(Get-ChildItem -LiteralPath $ImageDir -Filter 'page-*.png' -File | Sort-Object Name)
    if ($files.Count -eq 0) { return $null }

    Add-Type -AssemblyName System.Drawing
    $labelHeight = 24
    $rows = [int][Math]::Ceiling($files.Count / [double]$Columns)
    $sheet = $null
    $graphics = $null
    $font = $null
    try {
        $sheet = New-Object System.Drawing.Bitmap ($Columns * $ThumbWidth), ($rows * ($ThumbHeight + $labelHeight))
        $graphics = [System.Drawing.Graphics]::FromImage($sheet)
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $font = New-Object System.Drawing.Font('Arial', 10)
        for ($i = 0; $i -lt $files.Count; $i++) {
            $image = $null
            try {
                $image = [System.Drawing.Image]::FromFile($files[$i].FullName)
                $x = ($i % $Columns) * $ThumbWidth
                $y = [int][Math]::Floor($i / $Columns) * ($ThumbHeight + $labelHeight)
                $graphics.DrawImage($image, $x, $y, $ThumbWidth, $ThumbHeight)
                $graphics.DrawString($files[$i].BaseName, $font, [System.Drawing.Brushes]::Black, $x + 6, $y + $ThumbHeight + 4)
            } finally {
                if ($null -ne $image) { $image.Dispose() }
            }
        }
        $parent = Split-Path -Parent $OutputPath
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $sheet.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        return $OutputPath
    } finally {
        if ($null -ne $font) { $font.Dispose() }
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $sheet) { $sheet.Dispose() }
    }
}

function New-ContactSheets {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$OutputRoot,
        $IdentityMap
    )

    $sheetDir = Join-Path $OutputRoot '05_页面总览'
    $items = @{}
    foreach ($file in $Files) {
        $identity = $IdentityMap[$file.FullName]
        $sourceImageDir = Join-Path $OutputRoot ('04_页面图片\' + $identity.safeStem)
        $sheetPath = Join-Path $sheetDir ("$($identity.safeStem).contact-sheet.png")
        try {
            $created = New-ImageContactSheet -ImageDir $sourceImageDir -OutputPath $sheetPath
            if (-not [string]::IsNullOrWhiteSpace($created)) {
                $items[$file.FullName] = $created
            }
        } catch {
            Write-Warning "Contact sheet failed for $($file.Name): $($_.Exception.Message)"
        }
    }
    return $items
}

function Get-ReviewIssueNames {
    return @(
        'EmptySlideCandidate',
        'SmallText',
        'SmallTextPreserved',
        'SmallTextAfterNormalize',
        'FormulaStyleSkipped',
        'FormulaWhitelistCandidate',
        'FormulaTextStyleFailed',
        'FormulaConversionPending',
        'GroupShapeSkipped',
        'RasterPicturePreserved',
        'TextBoxWidthExpandFailed',
        'PdfExportFailed',
        'ImagesExportFailed',
        'ImageExportCountMismatch'
    )
}

function Get-ReviewSlideIssueMap {
    param($Rows)

    $reviewIssues = Get-ReviewIssueNames

    $map = @{}
    foreach ($row in @($Rows | Where-Object { $_.Slide -match '^\d+$' -and [int]$_.Slide -gt 0 -and $_.Issue -in $reviewIssues })) {
        $fileKey = Get-ReportRowFileKey -Row $row
        if ([string]::IsNullOrWhiteSpace($fileKey)) { continue }
        if (-not $map.ContainsKey($fileKey)) { $map[$fileKey] = @{} }
        $slideNo = [int]$row.Slide
        if (-not $map[$fileKey].ContainsKey($slideNo)) { $map[$fileKey][$slideNo] = New-Object System.Collections.Generic.List[string] }
        if (-not $map[$fileKey][$slideNo].Contains($row.Issue)) { $map[$fileKey][$slideNo].Add($row.Issue) }
    }
    return $map
}

function Get-ReviewSlideRecords {
    param(
        $Rows,
        [string]$ImageDir,
        [string]$SourceImageDir
    )

    $reviewIssues = Get-ReviewIssueNames

    return @(
        $Rows |
            Where-Object { $_.Slide -match '^\d+$' -and [int]$_.Slide -gt 0 -and $_.Issue -in $reviewIssues } |
            Group-Object Slide |
            Sort-Object { [int]$_.Name } |
            ForEach-Object {
                $slideNo = [int]$_.Name
                $pageImage = Join-Path $ImageDir ('page-{0:000}.png' -f $slideNo)
                $sourcePageImage = if ([string]::IsNullOrWhiteSpace($SourceImageDir)) { $null } else { Join-Path $SourceImageDir ('page-{0:000}.png' -f $slideNo) }
                $pageImageExists = Test-Path -LiteralPath $pageImage
                $sourcePageImageExists = (-not [string]::IsNullOrWhiteSpace($sourcePageImage) -and (Test-Path -LiteralPath $sourcePageImage))
                $pageImageValue = if ($pageImageExists) { $pageImage } else { $null }
                $sourcePageImageValue = if ($sourcePageImageExists) { $sourcePageImage } else { $null }
                $normalizedWhitePercent = if ($pageImageExists) { Get-ImageWhitePercent -ImagePath $pageImage } else { $null }
                [pscustomobject]@{
                    slide = $slideNo
                    issues = @(($_.Group | Select-Object -ExpandProperty Issue -Unique | Sort-Object))
                    pageImage = $pageImageValue
                    sourcePageImage = $sourcePageImageValue
                    visualDeltaPercent = if ($sourcePageImageExists -and $pageImageExists) { Get-ImageDeltaPercent -BeforePath $sourcePageImage -AfterPath $pageImage } else { $null }
                    sourceWhitePercent = if ($sourcePageImageExists) { Get-ImageWhitePercent -ImagePath $sourcePageImage } else { $null }
                    normalizedWhitePercent = $normalizedWhitePercent
                    isVisuallyBlank = ($null -ne $normalizedWhitePercent -and $normalizedWhitePercent -ge 98)
                    findings = @($_.Group | ForEach-Object {
                        [pscustomobject]@{
                            shape = $_.Shape
                            issue = $_.Issue
                            details = $_.Details
                        }
                    })
                }
            }
    )
}

function New-ReviewContactSheet {
    param(
        [string]$ImageDir,
        [string]$OutputPath,
        [hashtable]$SlideIssues,
        [int]$Columns = 3,
        [int]$ThumbWidth = 320,
        [int]$ThumbHeight = 180
    )

    if ($null -eq $SlideIssues -or $SlideIssues.Count -eq 0) { return $null }
    if (-not (Test-Path -LiteralPath $ImageDir)) { return $null }

    Add-Type -AssemblyName System.Drawing
    $slides = @($SlideIssues.Keys | Sort-Object { [int]$_ })
    $labelHeight = 48
    $rows = [int][Math]::Ceiling($slides.Count / [double]$Columns)
    $sheet = $null
    $graphics = $null
    $font = $null
    $smallFont = $null
    $borderPen = $null
    $drawn = 0
    try {
        $sheet = New-Object System.Drawing.Bitmap ($Columns * $ThumbWidth), ($rows * ($ThumbHeight + $labelHeight))
        $graphics = [System.Drawing.Graphics]::FromImage($sheet)
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $font = New-Object System.Drawing.Font('Arial', 10, [System.Drawing.FontStyle]::Bold)
        $smallFont = New-Object System.Drawing.Font('Arial', 8)
        $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(192, 0, 0), 3)

        for ($i = 0; $i -lt $slides.Count; $i++) {
            $slideNo = [int]$slides[$i]
            $source = Join-Path $ImageDir ('page-{0:000}.png' -f $slideNo)
            if (-not (Test-Path -LiteralPath $source)) { continue }
            $image = $null
            try {
                $image = [System.Drawing.Image]::FromFile($source)
                $x = ($i % $Columns) * $ThumbWidth
                $y = [int][Math]::Floor($i / $Columns) * ($ThumbHeight + $labelHeight)
                $graphics.DrawImage($image, $x, $y, $ThumbWidth, $ThumbHeight)
                $graphics.DrawRectangle($borderPen, $x + 1, $y + 1, $ThumbWidth - 3, $ThumbHeight - 3)
                $issueText = (@($SlideIssues[$slideNo]) | Sort-Object) -join ', '
                $graphics.DrawString(('page-{0:000}' -f $slideNo), $font, [System.Drawing.Brushes]::Black, $x + 6, $y + $ThumbHeight + 4)
                $graphics.DrawString($issueText, $smallFont, [System.Drawing.Brushes]::DarkRed, $x + 6, $y + $ThumbHeight + 24)
                $drawn++
            } finally {
                if ($null -ne $image) { $image.Dispose() }
            }
        }

        if ($drawn -eq 0) { return $null }
        $parent = Split-Path -Parent $OutputPath
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $sheet.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        return $OutputPath
    } finally {
        if ($null -ne $borderPen) { $borderPen.Dispose() }
        if ($null -ne $smallFont) { $smallFont.Dispose() }
        if ($null -ne $font) { $font.Dispose() }
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $sheet) { $sheet.Dispose() }
    }
}

function New-ReviewContactSheets {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$OutputRoot,
        $Rows,
        $IdentityMap
    )

    $reviewMap = Get-ReviewSlideIssueMap -Rows $Rows
    $sheetDir = Join-Path $OutputRoot '06_重点复核'
    $items = @{}
    foreach ($file in $Files) {
        if (-not $reviewMap.ContainsKey($file.FullName)) { continue }
        $identity = $IdentityMap[$file.FullName]
        $sourceImageDir = Join-Path $OutputRoot ('04_页面图片\' + $identity.safeStem)
        $sheetPath = Join-Path $sheetDir ("$($identity.safeStem).review-sheet.png")
        try {
            $created = New-ReviewContactSheet -ImageDir $sourceImageDir -OutputPath $sheetPath -SlideIssues $reviewMap[$file.FullName]
            if (-not [string]::IsNullOrWhiteSpace($created)) {
                $items[$file.FullName] = $created
            }
        } catch {
            Write-Warning "Review sheet failed for $($file.Name): $($_.Exception.Message)"
        }
    }
    return $items
}

function Rename-ExportedPageImages {
    param([string]$ImageDir)

    if (-not (Test-Path -LiteralPath $ImageDir)) { return }
    Get-ChildItem -LiteralPath $ImageDir -Filter '*.PNG' -File |
        ForEach-Object {
            if ($_.BaseName -match '(\d+)$') {
                $pageNo = [int]$Matches[1]
                $target = Join-Path $ImageDir ('page-{0:000}.png' -f $pageNo)
                if ($_.FullName -ne $target) {
                    Move-Item -LiteralPath $_.FullName -Destination $target -Force
                }
            }
        }
}

function Export-SourcePageImages {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$OutputRoot,
        $IdentityMap
    )

    $sourceImageRoot = Join-Path $OutputRoot '07_原始页面图片'
    $items = @{}
    $pp = $null
    try {
        $pp = New-Object -ComObject PowerPoint.Application
        foreach ($file in $Files) {
            $identity = $IdentityMap[$file.FullName]
            $imageDir = Join-Path $sourceImageRoot $identity.safeStem
            $pres = $null
            try {
                if (-not (Test-Path -LiteralPath $imageDir)) { New-Item -ItemType Directory -Path $imageDir -Force | Out-Null }
                Get-ChildItem -LiteralPath $imageDir -Filter '*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force
                $pres = $pp.Presentations.Open($file.FullName, 0, 0, 0)
                $pres.Export($imageDir, 'PNG')
                Rename-ExportedPageImages -ImageDir $imageDir
                $items[$file.FullName] = $imageDir
            } catch {
                Write-Warning "Original page image export failed for $($file.Name): $($_.Exception.Message)"
            } finally {
                if ($null -ne $pres) {
                    try { $pres.Close() | Out-Null } catch { }
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
                }
            }
        }
    } catch {
        Write-Warning "PowerPoint original export setup failed: $($_.Exception.Message)"
    } finally {
        if ($null -ne $pp) {
            try { $pp.Quit() | Out-Null } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    return $items
}

function New-BeforeAfterReviewSheet {
    param(
        [string]$SourceImageDir,
        [string]$NormalizedImageDir,
        [string]$OutputPath,
        [hashtable]$SlideIssues,
        [int]$ThumbWidth = 300,
        [int]$ThumbHeight = 169
    )

    if ($null -eq $SlideIssues -or $SlideIssues.Count -eq 0) { return $null }
    if (-not (Test-Path -LiteralPath $SourceImageDir)) { return $null }
    if (-not (Test-Path -LiteralPath $NormalizedImageDir)) { return $null }

    Add-Type -AssemblyName System.Drawing
    $slides = @($SlideIssues.Keys | Sort-Object { [int]$_ })
    $labelHeight = 46
    $gap = 18
    $sheetWidth = ($ThumbWidth * 2) + $gap
    $sheetHeight = $slides.Count * ($ThumbHeight + $labelHeight)
    $sheet = $null
    $graphics = $null
    $font = $null
    $smallFont = $null
    $linePen = $null
    $drawn = 0
    try {
        $sheet = New-Object System.Drawing.Bitmap $sheetWidth, $sheetHeight
        $graphics = [System.Drawing.Graphics]::FromImage($sheet)
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $font = New-Object System.Drawing.Font('Arial', 10, [System.Drawing.FontStyle]::Bold)
        $smallFont = New-Object System.Drawing.Font('Arial', 8)
        $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(160, 160, 160), 1)

        for ($i = 0; $i -lt $slides.Count; $i++) {
            $slideNo = [int]$slides[$i]
            $sourceImage = Join-Path $SourceImageDir ('page-{0:000}.png' -f $slideNo)
            $normalizedImage = Join-Path $NormalizedImageDir ('page-{0:000}.png' -f $slideNo)
            if (-not (Test-Path -LiteralPath $sourceImage) -or -not (Test-Path -LiteralPath $normalizedImage)) { continue }
            $y = $i * ($ThumbHeight + $labelHeight)
            $leftImage = $null
            $rightImage = $null
            try {
                $leftImage = [System.Drawing.Image]::FromFile($sourceImage)
                $rightImage = [System.Drawing.Image]::FromFile($normalizedImage)
                $graphics.DrawImage($leftImage, 0, $y, $ThumbWidth, $ThumbHeight)
                $graphics.DrawImage($rightImage, $ThumbWidth + $gap, $y, $ThumbWidth, $ThumbHeight)
                $graphics.DrawLine($linePen, 0, $y + $ThumbHeight + $labelHeight - 1, $sheetWidth, $y + $ThumbHeight + $labelHeight - 1)
                $issueText = (@($SlideIssues[$slideNo]) | Sort-Object) -join ', '
                $delta = Get-ImageDeltaPercent -BeforePath $sourceImage -AfterPath $normalizedImage
                $white = Get-ImageWhitePercent -ImagePath $normalizedImage
                $metricText = "Δ $delta%; white $white%"
                $graphics.DrawString(('page-{0:000}  原始' -f $slideNo), $font, [System.Drawing.Brushes]::Black, 6, $y + $ThumbHeight + 4)
                $graphics.DrawString('规范化后', $font, [System.Drawing.Brushes]::Black, $ThumbWidth + $gap + 6, $y + $ThumbHeight + 4)
                $graphics.DrawString($issueText, $smallFont, [System.Drawing.Brushes]::DarkRed, 6, $y + $ThumbHeight + 24)
                $graphics.DrawString($metricText, $smallFont, [System.Drawing.Brushes]::DimGray, $ThumbWidth + $gap + 6, $y + $ThumbHeight + 24)
                $drawn++
            } finally {
                if ($null -ne $leftImage) { $leftImage.Dispose() }
                if ($null -ne $rightImage) { $rightImage.Dispose() }
            }
        }

        if ($drawn -eq 0) { return $null }
        $parent = Split-Path -Parent $OutputPath
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $sheet.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        return $OutputPath
    } finally {
        if ($null -ne $linePen) { $linePen.Dispose() }
        if ($null -ne $smallFont) { $smallFont.Dispose() }
        if ($null -ne $font) { $font.Dispose() }
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $sheet) { $sheet.Dispose() }
    }
}

function New-BeforeAfterReviewSheets {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$OutputRoot,
        $Rows,
        $SourceImageDirs,
        $IdentityMap
    )

    $reviewMap = Get-ReviewSlideIssueMap -Rows $Rows
    $sheetDir = Join-Path $OutputRoot '08_前后对比'
    $items = @{}
    foreach ($file in $Files) {
        if (-not $reviewMap.ContainsKey($file.FullName)) { continue }
        if ($null -eq $SourceImageDirs -or -not $SourceImageDirs.ContainsKey($file.FullName)) { continue }
        $identity = $IdentityMap[$file.FullName]
        $normalizedImageDir = Join-Path $OutputRoot ('04_页面图片\' + $identity.safeStem)
        $sheetPath = Join-Path $sheetDir ("$($identity.safeStem).before-after-review.png")
        try {
            $created = New-BeforeAfterReviewSheet -SourceImageDir $SourceImageDirs[$file.FullName] -NormalizedImageDir $normalizedImageDir -OutputPath $sheetPath -SlideIssues $reviewMap[$file.FullName]
            if (-not [string]::IsNullOrWhiteSpace($created)) {
                $items[$file.FullName] = $created
            }
        } catch {
            Write-Warning "Before/after sheet failed for $($file.Name): $($_.Exception.Message)"
        }
    }
    return $items
}

function New-ReviewPagePackages {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$OutputRoot,
        $Rows,
        $IdentityMap
    )

    $reviewMap = Get-ReviewSlideIssueMap -Rows $Rows
    $packageRoot = Join-Path $OutputRoot '09_重点单页'
    $items = @{}
    foreach ($file in $Files) {
        if (-not $reviewMap.ContainsKey($file.FullName)) { continue }
        $identity = $IdentityMap[$file.FullName]
        $sourceImageDir = Join-Path $OutputRoot ('04_页面图片\' + $identity.safeStem)
        $targetDir = Join-Path $packageRoot $identity.safeStem
        if (-not (Test-Path -LiteralPath $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
        Get-ChildItem -LiteralPath $targetDir -Filter '*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force
        $copied = 0
        foreach ($slideNo in @($reviewMap[$file.FullName].Keys | Sort-Object { [int]$_ })) {
            $source = Join-Path $sourceImageDir ('page-{0:000}.png' -f ([int]$slideNo))
            if (-not (Test-Path -LiteralPath $source)) { continue }
            $issueText = (@($reviewMap[$file.FullName][[int]$slideNo]) | Sort-Object) -join '+'
            $target = Join-Path $targetDir ('page-{0:000}_{1}.png' -f ([int]$slideNo), $issueText)
            Copy-Item -LiteralPath $source -Destination $target -Force
            $copied++
        }
        if ($copied -gt 0) { $items[$file.FullName] = $targetDir }
    }
    return $items
}

function New-ReviewIndexes {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$OutputRoot,
        $Rows,
        $SourceImageDirs,
        $IdentityMap
    )

    $indexRoot = Join-Path $OutputRoot '10_复核索引'
    $items = @{}
    foreach ($file in $Files) {
        $identity = $IdentityMap[$file.FullName]
        $imageDir = Join-Path $OutputRoot ('04_页面图片\' + $identity.safeStem)
        $sourceImageDir = if ($null -ne $SourceImageDirs -and $SourceImageDirs.ContainsKey($file.FullName)) { $SourceImageDirs[$file.FullName] } else { $null }
        $fileRows = @($Rows | Where-Object { (Get-ReportRowFileKey -Row $_) -eq $file.FullName })
        $reviewSlides = @(Get-ReviewSlideRecords -Rows $fileRows -ImageDir $imageDir -SourceImageDir $sourceImageDir)
        if ($reviewSlides.Count -eq 0) { continue }
        if (-not (Test-Path -LiteralPath $indexRoot)) { New-Item -ItemType Directory -Path $indexRoot -Force | Out-Null }
        $indexPath = Join-Path $indexRoot ("$($identity.safeStem).review-pages.csv")
        $reviewSlides |
            ForEach-Object {
                [pscustomobject]@{
                    File = $identity.relativePath
                    Slide = $_.slide
                    Issues = ($_.issues -join ';')
                    VisualDeltaPercent = $_.visualDeltaPercent
                    NormalizedWhitePercent = $_.normalizedWhitePercent
                    IsVisuallyBlank = $_.isVisuallyBlank
                    PageImage = $_.pageImage
                    SourcePageImage = $_.sourcePageImage
                    Findings = (@($_.findings | ForEach-Object { "$($_.shape):$($_.issue)" }) -join ';')
                }
            } |
            Export-Csv -LiteralPath $indexPath -NoTypeInformation -Encoding UTF8
        $items[$file.FullName] = $indexPath
    }
    return $items
}

function Get-ExpectedSlideCount {
    param($Rows)
    return @($Rows | Where-Object { $_.Issue -eq 'SlideType' -and $_.Slide -match '^\d+$' } | Select-Object -ExpandProperty Slide -Unique).Count
}

function Get-PageImageCount {
    param([string]$ImageDir)
    if ([string]::IsNullOrWhiteSpace($ImageDir) -or -not (Test-Path -LiteralPath $ImageDir)) { return 0 }
    return @(Get-ChildItem -LiteralPath $ImageDir -Filter 'page-*.png' -File).Count
}

function Get-FileLength {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return 0 }
    return (Get-Item -LiteralPath $Path).Length
}

function Get-PdfPageCount {
    param([string]$PdfPath)
    if ([string]::IsNullOrWhiteSpace($PdfPath) -or -not (Test-Path -LiteralPath $PdfPath)) { return 0 }
    try {
        $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($PdfPath))
        return ([regex]::Matches($text, '/Type\s*/Page\b')).Count
    } catch {
        return 0
    }
}

function Get-ImageDeltaPercent {
    param(
        [string]$BeforePath,
        [string]$AfterPath
    )

    if ([string]::IsNullOrWhiteSpace($BeforePath) -or [string]::IsNullOrWhiteSpace($AfterPath)) { return $null }
    if (-not (Test-Path -LiteralPath $BeforePath) -or -not (Test-Path -LiteralPath $AfterPath)) { return $null }

    Add-Type -AssemblyName System.Drawing
    $before = $null
    $after = $null
    $beforeThumb = $null
    $afterThumb = $null
    try {
        $before = [System.Drawing.Image]::FromFile($BeforePath)
        $after = [System.Drawing.Image]::FromFile($AfterPath)
        $width = 120
        $height = 68
        $beforeThumb = New-Object System.Drawing.Bitmap $before, $width, $height
        $afterThumb = New-Object System.Drawing.Bitmap $after, $width, $height
        [double]$sum = 0
        for ($y = 0; $y -lt $height; $y++) {
            for ($x = 0; $x -lt $width; $x++) {
                $a = $beforeThumb.GetPixel($x, $y)
                $b = $afterThumb.GetPixel($x, $y)
                $sum += ([Math]::Abs($a.R - $b.R) + [Math]::Abs($a.G - $b.G) + [Math]::Abs($a.B - $b.B)) / (3 * 255)
            }
        }
        return [Math]::Round(($sum / ($width * $height)) * 100, 2)
    } catch {
        return $null
    } finally {
        if ($null -ne $afterThumb) { $afterThumb.Dispose() }
        if ($null -ne $beforeThumb) { $beforeThumb.Dispose() }
        if ($null -ne $after) { $after.Dispose() }
        if ($null -ne $before) { $before.Dispose() }
    }
}

function Get-ImageWhitePercent {
    param([string]$ImagePath)

    if ([string]::IsNullOrWhiteSpace($ImagePath) -or -not (Test-Path -LiteralPath $ImagePath)) { return $null }

    Add-Type -AssemblyName System.Drawing
    $image = $null
    $thumb = $null
    try {
        $image = [System.Drawing.Image]::FromFile($ImagePath)
        $width = 120
        $height = 68
        $thumb = New-Object System.Drawing.Bitmap $image, $width, $height
        $white = 0
        for ($y = 0; $y -lt $height; $y++) {
            for ($x = 0; $x -lt $width; $x++) {
                $pixel = $thumb.GetPixel($x, $y)
                if ($pixel.R -ge 245 -and $pixel.G -ge 245 -and $pixel.B -ge 245) { $white++ }
            }
        }
        return [Math]::Round(($white / ($width * $height)) * 100, 2)
    } catch {
        return $null
    } finally {
        if ($null -ne $thumb) { $thumb.Dispose() }
        if ($null -ne $image) { $image.Dispose() }
    }
}

function Get-PptxMediaSummary {
    param([string]$PptxPath)

    if ([string]::IsNullOrWhiteSpace($PptxPath) -or -not (Test-Path -LiteralPath $PptxPath)) {
        return [pscustomobject]@{ mediaCount = 0; totalBytes = 0; largest = @() }
    }

    $zip = $null
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
        $media = @($zip.Entries | Where-Object { $_.FullName -like 'ppt/media/*' -and $_.Length -gt 0 })
        $total = @($media | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $total) { $total = 0 }
        $largest = @(
            $media |
                Sort-Object Length -Descending |
                Select-Object -First 5 |
                ForEach-Object {
                    [pscustomobject]@{
                        name = $_.FullName
                        bytes = $_.Length
                    }
                }
        )
        return [pscustomobject]@{
            mediaCount = $media.Count
            totalBytes = [int64]$total
            largest = $largest
        }
    } catch {
        return [pscustomobject]@{ mediaCount = 0; totalBytes = 0; largest = @() }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

function Get-ReviewSlides {
    param($Rows)
    $reviewIssues = Get-ReviewIssueNames
    return @(
        $Rows |
            Where-Object { $_.Slide -match '^\d+$' -and [int]$_.Slide -gt 0 -and $_.Issue -in $reviewIssues } |
            Group-Object { Get-ReportRowFileKey -Row $_ } |
            ForEach-Object {
                $first = $_.Group | Select-Object -First 1
                [pscustomobject]@{
                    File = if ($null -ne $first.PSObject.Properties['FilePath'] -and -not [string]::IsNullOrWhiteSpace([string]$first.FilePath)) { [string]$first.FilePath } else { $_.Name }
                    Slides = @(($_.Group | Select-Object -ExpandProperty Slide -Unique | Sort-Object { [int]$_ }))
                    Issues = @(($_.Group | Select-Object -ExpandProperty Issue -Unique | Sort-Object))
                }
            }
    )
}

function Test-IsFinalFailureIssue {
    param([string]$Issue)
    if ([string]::IsNullOrWhiteSpace($Issue)) { return $false }

    $finalFailureIssues = @(
        'PowerPointBusyOrRejectedCall',
        'PowerPointComFailure',
        'PowerPointComNotRegistered',
        'FileInUseOrSharingViolation',
        'FileNotFoundOrUnavailable',
        'ChildProcessFailed',
        'UnhandledFailure'
    )
    if ($Issue -in $finalFailureIssues) { return $true }
    if ($Issue -in @('RetryAfterFailure', 'RetrySucceeded')) { return $false }
    return ($Issue -match 'Failed$|Failure$|NotRegistered$|SharingViolation$|Unavailable$')
}

function Get-ObjectPropertyValue {
    param(
        $Object,
        [string]$PropertyName,
        $DefaultValue = $null
    )

    if ($null -eq $Object) { return $DefaultValue }
    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return $DefaultValue }
    return $property.Value
}

function Get-VisualArtifactCountSum {
    param(
        $Items,
        [string]$PropertyName
    )

    [int]$sum = 0
    foreach ($item in @($Items)) {
        $value = Get-ObjectPropertyValue -Object $item -PropertyName $PropertyName -DefaultValue $null
        if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { continue }
        [int]$parsed = 0
        if ([int]::TryParse([string]$value, [ref]$parsed)) {
            $sum += $parsed
        }
    }
    return $sum
}

function Read-JsonObject {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Invoke-VisualConfirmationArtifact {
    param(
        [string]$VisualAuditDir,
        [string]$BaselineVisualAuditDir
    )

    $confirmationCsv = if ([string]::IsNullOrWhiteSpace($VisualAuditDir)) { $null } else { Join-Path $VisualAuditDir 'pptx-visual-confirmation.csv' }
    $confirmationManifestPath = if ([string]::IsNullOrWhiteSpace($VisualAuditDir)) { $null } else { Join-Path $VisualAuditDir 'pptx-visual-confirmation-manifest.json' }
    $confirmationContactSheet = if ([string]::IsNullOrWhiteSpace($VisualAuditDir)) { $null } else { Join-Path $VisualAuditDir 'pptx-visual-confirmation.contact-sheet.png' }
    $result = [ordered]@{
        status = 'SkippedNoVisualAuditDir'
        gateStatus = 'Skipped'
        error = ''
        csv = $null
        manifest = $null
        contactSheet = $null
        passedCount = 0
        passedWithKnownIssuesCount = 0
        needsReviewCount = 0
        failedCount = 0
        newErrorCount = 0
        newWarningCount = 0
    }

    if ([string]::IsNullOrWhiteSpace($VisualAuditDir) -or -not (Test-Path -LiteralPath $VisualAuditDir)) {
        return [pscustomobject]$result
    }

    try {
        $confirmationArgs = @{
            VisualAuditDir = $VisualAuditDir
            OutputDir = $VisualAuditDir
            ContactSheet = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($BaselineVisualAuditDir) -and (Test-Path -LiteralPath $BaselineVisualAuditDir)) {
            $confirmationArgs['BaselineVisualAuditDir'] = $BaselineVisualAuditDir
        }
        & (Join-Path $PSScriptRoot 'Export-PptxVisualConfirmation.ps1') @confirmationArgs
        $manifest = Read-JsonObject -Path $confirmationManifestPath
        $result.status = [string](Get-ObjectPropertyValue -Object $manifest -PropertyName 'confirmationStatus' -DefaultValue 'Completed')
        $result.gateStatus = [string](Get-ObjectPropertyValue -Object $manifest -PropertyName 'automationGateStatus' -DefaultValue 'Skipped')
        $result.passedCount = [int](Get-ObjectPropertyValue -Object $manifest -PropertyName 'passedCount' -DefaultValue 0)
        $result.passedWithKnownIssuesCount = [int](Get-ObjectPropertyValue -Object $manifest -PropertyName 'passedWithKnownIssuesCount' -DefaultValue 0)
        $result.needsReviewCount = [int](Get-ObjectPropertyValue -Object $manifest -PropertyName 'needsReviewCount' -DefaultValue 0)
        $result.failedCount = [int](Get-ObjectPropertyValue -Object $manifest -PropertyName 'failedCount' -DefaultValue 0)
        $result.newErrorCount = [int](Get-ObjectPropertyValue -Object $manifest -PropertyName 'newErrorCount' -DefaultValue 0)
        $result.newWarningCount = [int](Get-ObjectPropertyValue -Object $manifest -PropertyName 'newWarningCount' -DefaultValue 0)
    } catch {
        $result.status = 'Failed'
        $result.gateStatus = 'Failed'
        $result.error = $_.Exception.Message
    }

    $result.csv = if ($null -ne $confirmationCsv -and (Test-Path -LiteralPath $confirmationCsv)) { $confirmationCsv } else { $null }
    $result.manifest = if ($null -ne $confirmationManifestPath -and (Test-Path -LiteralPath $confirmationManifestPath)) { $confirmationManifestPath } else { $null }
    $result.contactSheet = if ($null -ne $confirmationContactSheet -and (Test-Path -LiteralPath $confirmationContactSheet)) { $confirmationContactSheet } else { $null }
    return [pscustomobject]$result
}

function Resolve-DotNetCommand {
    $candidates = New-Object System.Collections.Generic.List[string]
    $userDotnet = Join-Path $env:USERPROFILE '.dotnet\dotnet.exe'
    if (Test-Path -LiteralPath $userDotnet) { $candidates.Add($userDotnet) | Out-Null }

    $pathDotnet = Get-Command dotnet -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $pathDotnet) {
        $dotnetPath = if (-not [string]::IsNullOrWhiteSpace([string]$pathDotnet.Source)) { [string]$pathDotnet.Source } else { [string]$pathDotnet.Path }
        if (-not [string]::IsNullOrWhiteSpace($dotnetPath) -and $dotnetPath -notin $candidates) {
            $candidates.Add($dotnetPath) | Out-Null
        }
    }

    foreach ($candidate in $candidates) {
        try {
            $sdks = & $candidate --list-sdks 2>&1
            if ($LASTEXITCODE -eq 0 -and @($sdks | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) {
                return $candidate
            }
        } catch {
            continue
        }
    }

    return ''
}

function Get-FormulaOmmlFiles {
    param($FormulaOmmlArtifacts)

    if ($null -eq $FormulaOmmlArtifacts) { return @() }
    $files = Get-ObjectPropertyValue -Object $FormulaOmmlArtifacts -PropertyName 'files' -DefaultValue @()
    return @($files)
}

function Invoke-FormulaWhitelistSuggestionArtifacts {
    param(
        [string]$FormulaReviewIndex,
        [string]$OutputRoot,
        [int]$MaxSuggestions = 120
    )

    $suggestionDir = Join-Path $OutputRoot '10_复核索引\formula-whitelist-suggestions'
    $suggestionCsv = Join-Path $suggestionDir 'formula-whitelist-suggestions.csv'
    $suggestionManifest = Join-Path $suggestionDir 'formula-whitelist-suggestions-manifest.json'

    if ([string]::IsNullOrWhiteSpace($FormulaReviewIndex) -or -not (Test-Path -LiteralPath $FormulaReviewIndex)) {
        return [pscustomobject]@{
            status = 'SkippedNoFormulaReviewIndex'
            outputDir = $suggestionDir
            suggestionCsv = $null
            suggestionManifest = $null
            promoteCandidateCount = 0
            reviewCandidateCount = 0
            reviewSubscriptOnlyCount = 0
            existingWhitelistCount = 0
        }
    }

    try {
        & (Join-Path $PSScriptRoot 'Export-FormulaWhitelistSuggestions.ps1') -FormulaReviewCsv $FormulaReviewIndex -OutputDir $suggestionDir -MaxSuggestions $MaxSuggestions
        $manifest = Read-JsonObject -Path $suggestionManifest
        return [pscustomobject]@{
            status = 'Completed'
            outputDir = $suggestionDir
            suggestionCsv = if (Test-Path -LiteralPath $suggestionCsv) { $suggestionCsv } else { $null }
            suggestionManifest = if (Test-Path -LiteralPath $suggestionManifest) { $suggestionManifest } else { $null }
            promoteCandidateCount = Get-ObjectPropertyValue -Object $manifest -PropertyName 'promoteCandidateCount' -DefaultValue 0
            reviewCandidateCount = Get-ObjectPropertyValue -Object $manifest -PropertyName 'reviewCandidateCount' -DefaultValue 0
            reviewSubscriptOnlyCount = Get-ObjectPropertyValue -Object $manifest -PropertyName 'reviewSubscriptOnlyCount' -DefaultValue 0
            existingWhitelistCount = Get-ObjectPropertyValue -Object $manifest -PropertyName 'existingWhitelistCount' -DefaultValue 0
        }
    } catch {
        Write-Warning "Formula whitelist suggestions failed: $($_.Exception.Message)"
        return [pscustomobject]@{
            status = 'Failed'
            outputDir = $suggestionDir
            suggestionCsv = $null
            suggestionManifest = $null
            promoteCandidateCount = 0
            reviewCandidateCount = 0
            reviewSubscriptOnlyCount = 0
            existingWhitelistCount = 0
            error = $_.Exception.Message
        }
    }
}

function New-FormulaReviewIndex {
    param(
        $Rows,
        [string]$OutputRoot,
        $IdentityMap
    )

    $formulaIssues = @(
        'FormulaCandidate',
        'FormulaCandidateClass',
        'FormulaWhitelistCandidate',
        'FormulaConversionSkipped',
        'FormulaConversionPending',
        'FormulaTextStyleNormalized',
        'FormulaStyleSkipped',
        'FormulaTextStyleFailed'
    )

    $formulaRows = @($Rows | Where-Object { $_.Issue -in $formulaIssues })
    if ($formulaRows.Count -eq 0) { return $null }

    $items = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $formulaRows.Count; $i++) {
        $row = $formulaRows[$i]
        if ($row.Issue -ne 'FormulaCandidate') { continue }
        $fileKey = Get-ReportRowFileKey -Row $row
        $identity = if (-not [string]::IsNullOrWhiteSpace($fileKey) -and $null -ne $IdentityMap -and $IdentityMap.ContainsKey($fileKey)) { $IdentityMap[$fileKey] } else { $null }

        $details = @{
            FormulaCandidateClass = ''
            FormulaWhitelistCandidate = ''
            FormulaConversionSkipped = ''
            FormulaConversionPending = ''
            FormulaTextStyleNormalized = ''
            FormulaStyleSkipped = ''
            FormulaTextStyleFailed = ''
        }

        $j = $i + 1
        while ($j -lt $formulaRows.Count -and $formulaRows[$j].Issue -ne 'FormulaCandidate') {
            $next = $formulaRows[$j]
            if ((Get-ReportRowFileKey -Row $next) -eq $fileKey -and $next.Slide -eq $row.Slide -and $next.Shape -eq $row.Shape -and $details.ContainsKey($next.Issue)) {
                $details[$next.Issue] = $next.Details
            }
            $j++
        }

        $suggestedAction = 'ReviewCandidate'
        if (-not [string]::IsNullOrWhiteSpace($details.FormulaTextStyleFailed)) {
            $suggestedAction = 'FixStyleFailure'
        } elseif (-not [string]::IsNullOrWhiteSpace($details.FormulaWhitelistCandidate)) {
            $suggestedAction = 'ReviewWhitelistConversion'
        } elseif (-not [string]::IsNullOrWhiteSpace($details.FormulaStyleSkipped)) {
            $suggestedAction = 'ManualReview'
        } elseif (-not [string]::IsNullOrWhiteSpace($details.FormulaTextStyleNormalized)) {
            $suggestedAction = 'TextStyleNormalized'
        }

        $items.Add([pscustomobject]@{
            File = if ($null -ne $identity) { $identity.displayName } else { $row.File }
            FilePath = if ($null -ne $identity) { $identity.key } elseif (-not [string]::IsNullOrWhiteSpace($fileKey)) { $fileKey } else { $row.File }
            FileRelativePath = if ($null -ne $identity) { $identity.relativePath } else { $row.File }
            Slide = $row.Slide
            Shape = $row.Shape
            FormulaText = $row.Details
            CandidateClass = $details.FormulaCandidateClass
            WhitelistCandidate = $details.FormulaWhitelistCandidate
            ConversionStatus = if (-not [string]::IsNullOrWhiteSpace($details.FormulaConversionPending)) { $details.FormulaConversionPending } else { $details.FormulaConversionSkipped }
            StyleStatus = if (-not [string]::IsNullOrWhiteSpace($details.FormulaTextStyleNormalized)) { $details.FormulaTextStyleNormalized } elseif (-not [string]::IsNullOrWhiteSpace($details.FormulaTextStyleFailed)) { $details.FormulaTextStyleFailed } else { $details.FormulaStyleSkipped }
            SuggestedAction = $suggestedAction
        }) | Out-Null
    }

    if ($items.Count -eq 0) { return $null }

    $indexRoot = Join-Path $OutputRoot '10_复核索引'
    if (-not (Test-Path -LiteralPath $indexRoot)) { New-Item -ItemType Directory -Path $indexRoot -Force | Out-Null }
    $indexPath = Join-Path $indexRoot 'formula-review.csv'
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $csvLines = $items | ConvertTo-Csv -NoTypeInformation
    [System.IO.File]::WriteAllLines($indexPath, $csvLines, $utf8Bom)
    return $indexPath
}

function New-Manifest {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$InputFullPath,
        [string]$OutputRoot,
        [string]$ReportPath,
        [string]$Mode,
        [bool]$IncludeReviewArtifacts,
        $ReportRows,
        $ContactSheets,
        $ReviewSheets,
        $SourceImageDirs,
        $BeforeAfterSheets,
        $ReviewPagePackages,
        $ReviewIndexes,
        [string]$FormulaReviewIndex,
        $IdentityMap
    )

    $items = foreach ($file in $Files) {
        $identity = $IdentityMap[$file.FullName]
        $safeName = $identity.safeStem
        $pptxPath = Join-Path $OutputRoot "01_规范化PPTX\$safeName.normalized$($file.Extension)"
        $pdfPath = Join-Path $OutputRoot "02_导出PDF\$safeName.normalized.pdf"
        $imageDir = Join-Path $OutputRoot ('04_页面图片\' + $safeName)
        $fileRows = @($ReportRows | Where-Object { (Get-ReportRowFileKey -Row $_) -eq $file.FullName })
        $failed = @($fileRows | Where-Object { Test-IsFinalFailureIssue -Issue $_.Issue }).Count -gt 0
        $saved = Test-Path -LiteralPath $pptxPath
        $pdf = Test-Path -LiteralPath $pdfPath
        $pageImageCount = Get-PageImageCount -ImageDir $imageDir
        $images = ($pageImageCount -gt 0)
        $contactSheet = if ($null -ne $ContactSheets -and $ContactSheets.ContainsKey($file.FullName)) { $ContactSheets[$file.FullName] } else { $null }
        $reviewSheet = if ($null -ne $ReviewSheets -and $ReviewSheets.ContainsKey($file.FullName)) { $ReviewSheets[$file.FullName] } else { $null }
        $sourceImageDir = if ($null -ne $SourceImageDirs -and $SourceImageDirs.ContainsKey($file.FullName)) { $SourceImageDirs[$file.FullName] } else { $null }
        $beforeAfterSheet = if ($null -ne $BeforeAfterSheets -and $BeforeAfterSheets.ContainsKey($file.FullName)) { $BeforeAfterSheets[$file.FullName] } else { $null }
        $reviewPagePackage = if ($null -ne $ReviewPagePackages -and $ReviewPagePackages.ContainsKey($file.FullName)) { $ReviewPagePackages[$file.FullName] } else { $null }
        $reviewIndex = if ($null -ne $ReviewIndexes -and $ReviewIndexes.ContainsKey($file.FullName)) { $ReviewIndexes[$file.FullName] } else { $null }
        $reviewSlides = @(Get-ReviewSlideRecords -Rows $fileRows -ImageDir $imageDir -SourceImageDir $sourceImageDir)
        $expectedSlides = Get-ExpectedSlideCount -Rows $fileRows
        $sourcePageImageCount = Get-PageImageCount -ImageDir $sourceImageDir
        $pdfPageCount = Get-PdfPageCount -PdfPath $pdfPath
        $status = if ($failed) { 'failed' } elseif ($Mode -eq 'CheckOnly') { 'checked' } elseif ($saved) { 'success' } else { 'unknown' }

            [pscustomobject]@{
                input = $file.FullName
                inputRelativePath = $identity.relativePath
                displayName = $identity.displayName
                outputStem = $safeName
            inputBytes = Get-FileLength -Path $file.FullName
            normalizedPptx = if ($saved) { $pptxPath } elseif ($Mode -eq 'CheckOnly') { $null } else { $pptxPath }
            normalizedPptxBytes = Get-FileLength -Path $pptxPath
            pdf = if ($pdf) { $pdfPath } elseif ($Mode -in @('CheckOnly', 'SafeNormalize')) { $null } else { $pdfPath }
            pdfBytes = Get-FileLength -Path $pdfPath
            pageImages = if ($images) { $imageDir } elseif ($Mode -eq 'CheckOnly' -or -not $IncludeReviewArtifacts) { $null } else { $imageDir }
            sourcePageImages = if ($IncludeReviewArtifacts) { $sourceImageDir } else { $null }
            contactSheet = $contactSheet
            reviewSheet = $reviewSheet
            beforeAfterSheet = $beforeAfterSheet
            reviewPagePackage = $reviewPagePackage
            reviewIndex = $reviewIndex
            reviewSlides = $reviewSlides
            inputMedia = Get-PptxMediaSummary -PptxPath $file.FullName
            normalizedMedia = Get-PptxMediaSummary -PptxPath $pptxPath
            validation = [pscustomobject]@{
                expectedSlides = $expectedSlides
                pageImageCount = $pageImageCount
                sourcePageImageCount = $sourcePageImageCount
                pdfPageCount = $pdfPageCount
                pageImagesMatchExpected = (-not $IncludeReviewArtifacts -or $expectedSlides -eq 0 -or $pageImageCount -eq $expectedSlides)
                sourceImagesMatchExpected = (-not $IncludeReviewArtifacts -or $expectedSlides -eq 0 -or $sourcePageImageCount -eq 0 -or $sourcePageImageCount -eq $expectedSlides)
                pdfPagesMatchExpected = ($expectedSlides -eq 0 -or $pdfPageCount -eq 0 -or $pdfPageCount -eq $expectedSlides)
            }
            report = $ReportPath
            status = $status
            issues = @($fileRows | Group-Object Issue | ForEach-Object {
                [pscustomobject]@{ issue = $_.Name; count = $_.Count }
            })
        }
    }

    [pscustomobject]@{
        generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        mode = $Mode
        source = $InputFullPath
        output = $OutputRoot
        report = $ReportPath
        reviewArtifactsEnabled = $IncludeReviewArtifacts
        formulaReviewIndex = $FormulaReviewIndex
        files = @($items)
    }
}

function Write-Summary {
    param(
        [string]$Path,
        [string]$Mode,
        [string]$InputFullPath,
        [string]$OutputRoot,
        [string]$ReportPath,
        [string]$ManifestPath,
        [System.IO.FileInfo[]]$Files,
        [bool]$IncludeReviewArtifacts,
        $Rows,
        $Manifest
    )

    $reviewItems = @(Get-ReviewSlides -Rows $Rows)
    $manifestFiles = @($Manifest.files)
    $visualAuditItems = @()
    if ($null -ne $Manifest -and ($Manifest.PSObject.Properties.Name -contains 'visualAuditArtifacts')) {
        $visualAuditItems = @($Manifest.visualAuditArtifacts)
    }
    $formulaOmmlArtifacts = Get-ObjectPropertyValue -Object $Manifest -PropertyName 'formulaOmmlArtifacts' -DefaultValue $null
    $formulaOmmlVisualAuditEnabled = [bool](Get-ObjectPropertyValue -Object $Manifest -PropertyName 'formulaOmmlVisualAuditEnabled' -DefaultValue $false)
    $formulaOmmlFiles = @(Get-FormulaOmmlFiles -FormulaOmmlArtifacts $formulaOmmlArtifacts)
    $successCount = @($manifestFiles | Where-Object { $_.status -eq 'success' -or $_.status -eq 'checked' }).Count
    $failedCount = @($manifestFiles | Where-Object { $_.status -eq 'failed' }).Count
    $pdfCount = @($manifestFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.pdf) -and (Test-Path -LiteralPath $_.pdf) }).Count
    $visualAuditCount = @($visualAuditItems | Where-Object { (Get-ObjectPropertyValue -Object $_ -PropertyName 'visualAuditStatus' -DefaultValue '') -eq 'Completed' }).Count
    $visualAuditFailedCount = @($visualAuditItems | Where-Object { (Get-ObjectPropertyValue -Object $_ -PropertyName 'visualAuditStatus' -DefaultValue '') -eq 'Failed' }).Count
    $visualAuditErrorCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'visualAuditErrorCount'
    $visualAuditWarningCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'visualAuditWarningCount'
    $visualConfirmationGatePassedCount = @($visualAuditItems | Where-Object { (Get-ObjectPropertyValue -Object $_ -PropertyName 'visualConfirmationGateStatus' -DefaultValue '') -eq 'Passed' }).Count
    $visualConfirmationNeedsReviewCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'visualConfirmationNeedsReviewCount'
    $visualConfirmationFailedCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'visualConfirmationFailedCount'
    $visualFixedCount = @($visualAuditItems | Where-Object {
        $fixed = Get-ObjectPropertyValue -Object $_ -PropertyName 'fixedPptx' -DefaultValue ''
        -not [string]::IsNullOrWhiteSpace($fixed) -and (Test-Path -LiteralPath $fixed)
    }).Count
    $fixedVisualAuditErrorCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'fixedVisualAuditErrorCount'
    $fixedVisualAuditWarningCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'fixedVisualAuditWarningCount'
    $fixedVisualConfirmationGatePassedCount = @($visualAuditItems | Where-Object { (Get-ObjectPropertyValue -Object $_ -PropertyName 'fixedVisualConfirmationGateStatus' -DefaultValue '') -eq 'Passed' }).Count
    $fixedVisualConfirmationNeedsReviewCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'fixedVisualConfirmationNeedsReviewCount'
    $fixedVisualConfirmationFailedCount = Get-VisualArtifactCountSum -Items $visualAuditItems -PropertyName 'fixedVisualConfirmationFailedCount'
    $formulaOmmlCopyCount = @($formulaOmmlFiles | Where-Object {
        $outputPptx = Get-ObjectPropertyValue -Object $_ -PropertyName 'outputPptx' -DefaultValue ''
        -not [string]::IsNullOrWhiteSpace($outputPptx) -and (Test-Path -LiteralPath $outputPptx)
    }).Count
    $formulaOmmlInsertedCount = Get-VisualArtifactCountSum -Items $formulaOmmlFiles -PropertyName 'insertedCount'
    $formulaOmmlApplyIssueCount = Get-VisualArtifactCountSum -Items $formulaOmmlFiles -PropertyName 'applyIssueCount'
    $formulaOmmlOpenXmlErrorCount = Get-VisualArtifactCountSum -Items $formulaOmmlFiles -PropertyName 'openXmlErrorCount'
    $formulaOmmlVisualErrorCount = Get-VisualArtifactCountSum -Items $formulaOmmlFiles -PropertyName 'visualAuditErrorCount'
    $formulaOmmlVisualWarningCount = Get-VisualArtifactCountSum -Items $formulaOmmlFiles -PropertyName 'visualAuditWarningCount'
    $formulaOmmlVisualConfirmationFailedCount = Get-VisualArtifactCountSum -Items $formulaOmmlFiles -PropertyName 'visualConfirmationFailedCount'
    $formulaOmmlVisualConfirmationNeedsReviewCount = Get-VisualArtifactCountSum -Items $formulaOmmlFiles -PropertyName 'visualConfirmationNeedsReviewCount'
    $formulaOmmlGatePassedCount = @($formulaOmmlFiles | Where-Object { (Get-ObjectPropertyValue -Object $_ -PropertyName 'automationGateStatus' -DefaultValue '') -eq 'Passed' }).Count
    $formulaOmmlGateFailedCount = @($formulaOmmlFiles | Where-Object { (Get-ObjectPropertyValue -Object $_ -PropertyName 'automationGateStatus' -DefaultValue '') -eq 'Failed' }).Count
    $imageCount = if ($IncludeReviewArtifacts) { @($manifestFiles | Where-Object { $null -ne $_.validation -and $_.validation.pageImageCount -gt 0 }).Count } else { 0 }
    $sourceImageCount = if ($IncludeReviewArtifacts) { @($manifestFiles | Where-Object { $null -ne $_.validation -and $_.validation.sourcePageImageCount -gt 0 }).Count } else { 0 }
    $contactSheetCount = if ($IncludeReviewArtifacts) { @($manifestFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.contactSheet) -and (Test-Path -LiteralPath $_.contactSheet) }).Count } else { 0 }
    $reviewSheetCount = if ($IncludeReviewArtifacts) { @($manifestFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.reviewSheet) -and (Test-Path -LiteralPath $_.reviewSheet) }).Count } else { 0 }
    $beforeAfterSheetCount = if ($IncludeReviewArtifacts) { @($manifestFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.beforeAfterSheet) -and (Test-Path -LiteralPath $_.beforeAfterSheet) }).Count } else { 0 }
    $reviewPagePackageCount = if ($IncludeReviewArtifacts) { @($manifestFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.reviewPagePackage) -and (Test-Path -LiteralPath $_.reviewPagePackage) }).Count } else { 0 }
    $reviewIndexCount = if ($IncludeReviewArtifacts) { @($manifestFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.reviewIndex) -and (Test-Path -LiteralPath $_.reviewIndex) }).Count } else { 0 }
        $formulaReviewIndex = if (-not [string]::IsNullOrWhiteSpace($Manifest.formulaReviewIndex) -and (Test-Path -LiteralPath $Manifest.formulaReviewIndex)) { $Manifest.formulaReviewIndex } else { $null }
    $formulaReviewCount = if ($null -ne $formulaReviewIndex) { @((Import-Csv -LiteralPath $formulaReviewIndex -Encoding UTF8)).Count } else { 0 }
    $formulaWhitelistSuggestionArtifacts = Get-ObjectPropertyValue -Object $Manifest -PropertyName 'formulaWhitelistSuggestionArtifacts' -DefaultValue $null
    $formulaWhitelistSuggestionStatus = Get-ObjectPropertyValue -Object $formulaWhitelistSuggestionArtifacts -PropertyName 'status' -DefaultValue ''
    $formulaWhitelistSuggestionCsv = Get-ObjectPropertyValue -Object $formulaWhitelistSuggestionArtifacts -PropertyName 'suggestionCsv' -DefaultValue $null
    $formulaWhitelistPromoteCount = Get-ObjectPropertyValue -Object $formulaWhitelistSuggestionArtifacts -PropertyName 'promoteCandidateCount' -DefaultValue 0
    $formulaWhitelistReviewCount = Get-ObjectPropertyValue -Object $formulaWhitelistSuggestionArtifacts -PropertyName 'reviewCandidateCount' -DefaultValue 0
    $formulaWhitelistSubscriptOnlyCount = Get-ObjectPropertyValue -Object $formulaWhitelistSuggestionArtifacts -PropertyName 'reviewSubscriptOnlyCount' -DefaultValue 0
    $reviewSlideCount = @($manifestFiles | ForEach-Object { @($_.reviewSlides).Count } | Measure-Object -Sum).Sum
    $highVisualDeltaCount = if ($IncludeReviewArtifacts) { @($manifestFiles | ForEach-Object { @($_.reviewSlides) } | Where-Object { $null -ne $_.visualDeltaPercent -and $_.visualDeltaPercent -ge 8 }).Count } else { 0 }
    $visuallyBlankCount = if ($IncludeReviewArtifacts) { @($manifestFiles | ForEach-Object { @($_.reviewSlides) } | Where-Object { $_.isVisuallyBlank }).Count } else { 0 }
    $validationMismatchCount = @(
        $manifestFiles |
            Where-Object {
                $null -ne $_.validation -and (
                    -not $_.validation.pageImagesMatchExpected -or
                    -not $_.validation.sourceImagesMatchExpected -or
                    -not $_.validation.pdfPagesMatchExpected
                )
            }
    ).Count

    $lines = New-Object System.Collections.Generic.List[string]
    $exportsPdf = $Mode -notin @('CheckOnly', 'SafeNormalize')
    $lines.Add('# 物理 PPT 处理结果总览')
    $lines.Add('')
    $lines.Add("- 生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add("- 模式：$Mode")
    $lines.Add("- 输入：$InputFullPath")
    $lines.Add("- 输出目录：$OutputRoot")
    $lines.Add("- 报告：$ReportPath")
    $lines.Add("- 结构化清单：$ManifestPath")
    $lines.Add("- 复核方式：必要时打开生成的 PPTX 全文件复核")
    $lines.Add('')
    $lines.Add('## 处理概况')
    $lines.Add('')
    $lines.Add("- 文件数：$(@($Files).Count)")
    $lines.Add("- 成功/已检查：$successCount")
    $lines.Add("- 失败：$failedCount")
    if ($exportsPdf) {
        $lines.Add("- 已导出 PDF：$pdfCount")
    } else {
        $lines.Add("- PDF 导出：已跳过")
    }
    if ($IncludeReviewArtifacts) {
        $lines.Add("- 已导出页面图片目录：$imageCount")
        $lines.Add("- 已导出原始页面图片目录：$sourceImageCount")
    }
    if ($visualAuditItems.Count -gt 0) {
        $lines.Add("- 视觉审查完成/失败：$visualAuditCount / $visualAuditFailedCount")
        $lines.Add("- 视觉审查错误/警告：$visualAuditErrorCount / $visualAuditWarningCount")
        $lines.Add("- 视觉自动确认通过/需复核/失败：$visualConfirmationGatePassedCount / $visualConfirmationNeedsReviewCount / $visualConfirmationFailedCount")
        if ($visualFixedCount -gt 0) {
            $lines.Add("- 已生成视觉修复副本：$visualFixedCount")
            $lines.Add("- 修复后视觉审查错误/警告：$fixedVisualAuditErrorCount / $fixedVisualAuditWarningCount")
            $lines.Add("- 修复后视觉自动确认通过/需复核/失败：$fixedVisualConfirmationGatePassedCount / $fixedVisualConfirmationNeedsReviewCount / $fixedVisualConfirmationFailedCount")
        }
    }
    if ($null -ne $formulaOmmlArtifacts) {
        $candidateStatus = Get-ObjectPropertyValue -Object $formulaOmmlArtifacts -PropertyName 'candidateStatus' -DefaultValue ''
        $candidateCount = Get-ObjectPropertyValue -Object $formulaOmmlArtifacts -PropertyName 'candidateCount' -DefaultValue 0
        $generatedCandidateCount = Get-ObjectPropertyValue -Object $formulaOmmlArtifacts -PropertyName 'generatedCandidateCount' -DefaultValue 0
        $lines.Add("- OMML 候选生成：$candidateStatus，$generatedCandidateCount / $candidateCount")
        $lines.Add("- 已生成可编辑公式副本：$formulaOmmlCopyCount")
        $lines.Add("- 已写入 OMML 公式：$formulaOmmlInsertedCount")
        $lines.Add("- OMML 写入需复核项：$formulaOmmlApplyIssueCount")
        $lines.Add("- OMML Open XML 错误：$formulaOmmlOpenXmlErrorCount")
        if ($formulaOmmlVisualAuditEnabled) {
            $lines.Add("- OMML 副本视觉审查错误/警告：$formulaOmmlVisualErrorCount / $formulaOmmlVisualWarningCount")
            $lines.Add("- OMML 副本视觉自动确认需复核/失败：$formulaOmmlVisualConfirmationNeedsReviewCount / $formulaOmmlVisualConfirmationFailedCount")
        } else {
            $lines.Add("- OMML 副本视觉审查：已跳过（快速公式闭环）")
        }
        $lines.Add("- OMML 自动门禁通过/失败：$formulaOmmlGatePassedCount / $formulaOmmlGateFailedCount")
    }
    $lines.Add("- 疑似公式：$(Get-IssueCount -Rows $Rows -Issue 'FormulaCandidate')")
    $lines.Add("- 已归一低风险文本公式：$(Get-IssueCount -Rows $Rows -Issue 'FormulaTextStyleNormalized')")
    $lines.Add("- 公式样式跳过：$(Get-IssueCount -Rows $Rows -Issue 'FormulaStyleSkipped')")
    $lines.Add("- 白名单公式候选：$(Get-IssueCount -Rows $Rows -Issue 'FormulaWhitelistCandidate')")
    $lines.Add("- 公式转换跳过：$(Get-IssueCount -Rows $Rows -Issue 'FormulaConversionSkipped')")
    $lines.Add("- 小字号：$(Get-SmallTextCount -Rows $Rows)")
    $lines.Add("- 大图/位图文字需复核：$(Get-IssueCount -Rows $Rows -Issue 'RasterPicturePreserved')")
    $lines.Add("- 分节标题页：$(Get-IssueCount -Rows $Rows -Issue 'SectionTitleSlide')")
    $lines.Add("- 封面页保留样式：$(Get-IssueCount -Rows $Rows -Issue 'CoverSlideStylePreserved')")
    $lines.Add("- 结束页保留样式：$(Get-IssueCount -Rows $Rows -Issue 'EndingSlideStylePreserved')")
    $lines.Add("- 资源页保留样式：$(Get-IssueCount -Rows $Rows -Issue 'ResourceSlideStylePreserved')")
    $lines.Add("- 补充说明页保留样式：$(Get-IssueCount -Rows $Rows -Issue 'AppendixTextSlideStylePreserved')")
    $lines.Add("- 已横向扩展文本框：$(Get-IssueCount -Rows $Rows -Issue 'TextBoxWidthExpanded')")
    $lines.Add("- 已上下对齐答案框：$(Get-IssueCount -Rows $Rows -Issue 'AnswerTextAligned')")
    $lines.Add("- 已检查答案框对齐：$(Get-IssueCount -Rows $Rows -Issue 'AnswerTextAlignmentChecked')")
    $lines.Add("- 已设置答案劈裂动画：$((Get-IssueCount -Rows $Rows -Issue 'AnswerAnimationSet') + (Get-IssueCount -Rows $Rows -Issue 'AnswerAnimationAdded'))")
    $lines.Add("- 组合对象跳过：$(Get-IssueCount -Rows $Rows -Issue 'GroupShapeSkipped')")
    $lines.Add("- 空白页候选：$(Get-IssueCount -Rows $Rows -Issue 'EmptySlideCandidate')")
    $lines.Add("- 已禁用单击换片：$(Get-IssueCount -Rows $Rows -Issue 'AdvanceOnClickDisabled')")
    $lines.Add("- 视频页候选：$(Get-VideoCandidateCount -Rows $Rows)")
    if ($exportsPdf) {
        $lines.Add("- PDF 导出失败：$(Get-IssueCount -Rows $Rows -Issue 'PdfExportFailed')")
    }
    if ($IncludeReviewArtifacts) {
        $lines.Add("- 页面图片数量异常：$(Get-ImageExportMismatchCount -Rows $Rows)")
        $lines.Add("- 页面总览图：$contactSheetCount")
        $lines.Add("- 重点复核图：$reviewSheetCount")
        $lines.Add("- 前后对比图：$beforeAfterSheetCount")
        $lines.Add("- 重点单页包：$reviewPagePackageCount")
        $lines.Add("- 复核索引：$reviewIndexCount")
    }
    $lines.Add("- 公式复核清单：$formulaReviewCount")
    if (-not [string]::IsNullOrWhiteSpace($formulaWhitelistSuggestionStatus)) {
        $lines.Add("- 公式白名单建议：$formulaWhitelistSuggestionStatus；可晋升/需复核/下标待看：$formulaWhitelistPromoteCount / $formulaWhitelistReviewCount / $formulaWhitelistSubscriptOnlyCount")
    }
    $lines.Add("- 重点复核页：$reviewSlideCount")
    if ($IncludeReviewArtifacts) {
        $lines.Add("- 视觉变化较大页：$highVisualDeltaCount")
        $lines.Add("- 视觉空白页：$visuallyBlankCount")
    }
    $lines.Add("- 输出校验异常：$validationMismatchCount")
    $lines.Add('')
    $lines.Add('## 需要人工复核')
    $lines.Add('')

    if ($reviewItems.Count -eq 0) {
        $lines.Add('- 未从报告中发现需要优先复核的页。')
    } else {
        foreach ($item in $reviewItems) {
            $lines.Add("- $($item.File)：第 $($item.Slides -join '、') 页；问题：$($item.Issues -join '、')")
        }
    }

    $lines.Add('')
    $lines.Add('## 输出校验')
    $lines.Add('')
    foreach ($file in $manifestFiles) {
        $validation = $file.validation
        $inputName = if (-not [string]::IsNullOrWhiteSpace($file.inputRelativePath)) { [string]$file.inputRelativePath } else { Split-Path $file.input -Leaf }
        $inputMb = [Math]::Round(([double]$file.inputBytes / 1MB), 1)
        $normalizedMb = [Math]::Round(([double]$file.normalizedPptxBytes / 1MB), 1)
        $pdfMb = [Math]::Round(([double]$file.pdfBytes / 1MB), 1)
        $imageChecksOk = ($null -ne $validation -and (-not $IncludeReviewArtifacts -or ($validation.pageImagesMatchExpected -and $validation.sourceImagesMatchExpected)))
        $checksOk = ($null -ne $validation -and $imageChecksOk -and (-not $exportsPdf -or $validation.pdfPagesMatchExpected))
        $checkText = if ($checksOk) { '通过' } else { '需复核' }
        if ($null -eq $validation) {
            $lines.Add("- $inputName：无输出校验数据。")
        } elseif ($IncludeReviewArtifacts) {
            $line = "- $inputName：页数 $($validation.expectedSlides)；规范化图 $($validation.pageImageCount)；原始图 $($validation.sourcePageImageCount)；输入 $inputMb MB；规范化 PPTX $normalizedMb MB"
            if ($exportsPdf) { $line += "；PDF 页 $($validation.pdfPageCount)；PDF $pdfMb MB" }
            $line += "；校验 $checkText"
            $lines.Add($line)
        } else {
            $line = "- $inputName：页数 $($validation.expectedSlides)；输入 $inputMb MB；规范化 PPTX $normalizedMb MB"
            if ($exportsPdf) { $line += "；PDF 页 $($validation.pdfPageCount)；PDF $pdfMb MB" }
            $line += "；校验 $checkText"
            $lines.Add($line)
            $mediaMb = [Math]::Round(([double]$file.inputMedia.totalBytes / 1MB), 1)
            $largest = @($file.inputMedia.largest | Select-Object -First 1)
            if ($largest.Count -gt 0) {
                $largestName = ([string]$largest[0].name) -replace '^ppt/media/', ''
                $largestMb = [Math]::Round(([double]$largest[0].bytes / 1MB), 1)
                $lines.Add("- $inputName 媒体：$($file.inputMedia.mediaCount) 个；合计 $mediaMb MB；最大 $largestName $largestMb MB")
            } else {
                $lines.Add("- $inputName 媒体：0 个；合计 0 MB")
            }
        }
    }

    $lines.Add('')
    $lines.Add('## 输出结构')
    $lines.Add('')
    $lines.Add('- `00_检查报告/`：检查报告与规范化报告')
    $lines.Add('- `01_规范化PPTX/`：处理后的 PPTX/PPTM')
    if ($exportsPdf) {
        $lines.Add('- `02_导出PDF/`：处理后导出的同名 PDF')
    }
    $lines.Add('- `03_原始备份/`：原始文件备份')
    if ($IncludeReviewArtifacts) {
        $lines.Add('- `04_页面图片/`：处理后 PPTX 导出的逐页 PNG')
        $lines.Add('- `05_页面总览/`：逐页 PNG 的缩略总览图')
        $lines.Add('- `06_重点复核/`：报告命中的高风险页缩略总览图')
        $lines.Add('- `07_原始页面图片/`：原始 PPTX 导出的逐页 PNG，供前后对比')
        $lines.Add('- `08_前后对比/`：重点复核页的原始/规范化后并排总览图')
        $lines.Add('- `09_重点单页/`：报告命中的高风险页单页 PNG 包')
        $lines.Add('- `10_复核索引/`：重点复核页 CSV 索引，包含问题、图片路径和视觉指标')
    }
    if ($null -ne $formulaReviewIndex) {
        $lines.Add("- 公式复核清单：``$formulaReviewIndex``")
    }
    if (-not [string]::IsNullOrWhiteSpace($formulaWhitelistSuggestionCsv) -and (Test-Path -LiteralPath $formulaWhitelistSuggestionCsv)) {
        $lines.Add("- 公式白名单建议清单：``$formulaWhitelistSuggestionCsv``")
    }
    if ($visualAuditItems.Count -gt 0) {
        $lines.Add('- `11_视觉审查/`：对生成 PPTX 导出的 PNG/PDF、形状越界、视觉指标和自动确认门禁')
        $lines.Add('- `12_视觉修复/`：仅包含规则可确定的低风险修复副本，不覆盖规范化 PPTX')
        $lines.Add('- `13_视觉修复审查/`：对视觉修复副本再次导出的审查报告和自动确认门禁')
    }
    if ($null -ne $formulaOmmlArtifacts) {
        $lines.Add('- `14_公式OMML副本/`：白名单公式转换为可编辑 OfficeMath/OMML 的 PPTX 副本')
        if ($formulaOmmlVisualAuditEnabled) {
            $lines.Add('- `15_公式OMML审查/`：OMML 副本的 Open XML 校验、PDF/PNG 和视觉审查报告')
        } else {
            $lines.Add('- `15_公式OMML审查/`：OMML 副本的 Open XML 校验报告；逐页视觉审查默认跳过')
        }
    }
    $lines.Add('- `review-manifest.json`：结构化处理清单')

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($Path, $lines, $utf8Bom)
}

function Open-GeneratedPptxResult {
    param(
        $Manifest,
        [string]$NormalizedDir,
        [string]$VisualFixDir,
        [string]$OutputRoot
    )

    $formulaOutputs = @()
    if ($null -ne $Manifest -and ($Manifest.PSObject.Properties.Name -contains 'formulaOmmlArtifacts')) {
        $formulaOutputs = @(
            Get-FormulaOmmlFiles -FormulaOmmlArtifacts $Manifest.formulaOmmlArtifacts |
                Where-Object {
                    $outputPptx = Get-ObjectPropertyValue -Object $_ -PropertyName 'outputPptx' -DefaultValue ''
                    $gate = Get-ObjectPropertyValue -Object $_ -PropertyName 'automationGateStatus' -DefaultValue ''
                    -not [string]::IsNullOrWhiteSpace($outputPptx) -and (Test-Path -LiteralPath $outputPptx) -and ($gate -eq 'Passed')
                } |
                ForEach-Object { Get-ObjectPropertyValue -Object $_ -PropertyName 'outputPptx' -DefaultValue '' } |
                Select-Object -Unique
        )
    }

    if ($formulaOutputs.Count -eq 1) {
        Invoke-Item -LiteralPath $formulaOutputs[0]
        return
    } elseif ($formulaOutputs.Count -gt 1) {
        $formulaDir = Join-Path $OutputRoot '14_公式OMML副本'
        if (Test-Path -LiteralPath $formulaDir) {
            Invoke-Item -LiteralPath $formulaDir
            return
        }
    }

    $fixedOutputs = @()
    if ($null -ne $Manifest -and ($Manifest.PSObject.Properties.Name -contains 'visualAuditArtifacts')) {
        $fixedOutputs = @(
            @($Manifest.visualAuditArtifacts) |
                Where-Object {
                    $fixed = Get-ObjectPropertyValue -Object $_ -PropertyName 'fixedPptx' -DefaultValue ''
                    -not [string]::IsNullOrWhiteSpace($fixed) -and (Test-Path -LiteralPath $fixed)
                } |
                ForEach-Object { Get-ObjectPropertyValue -Object $_ -PropertyName 'fixedPptx' -DefaultValue '' } |
                Select-Object -Unique
        )
    }

    if ($fixedOutputs.Count -eq 1) {
        Invoke-Item -LiteralPath $fixedOutputs[0]
        return
    } elseif ($fixedOutputs.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($VisualFixDir) -and (Test-Path -LiteralPath $VisualFixDir)) {
        Invoke-Item -LiteralPath $VisualFixDir
        return
    }

    $outputs = @(
        @($Manifest.files) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.normalizedPptx) -and (Test-Path -LiteralPath $_.normalizedPptx) } |
            Select-Object -ExpandProperty normalizedPptx -Unique
    )

    if ($outputs.Count -eq 1) {
        Invoke-Item -LiteralPath $outputs[0]
    } elseif ($outputs.Count -gt 1 -and (Test-Path -LiteralPath $NormalizedDir)) {
        Invoke-Item -LiteralPath $NormalizedDir
    } else {
        Invoke-Item -LiteralPath $OutputRoot
    }
}

function Invoke-VisualAuditArtifacts {
    param(
        $Manifest,
        [string]$OutputRoot,
        [bool]$ApplyFixes,
        [string]$Mode
    )

    $items = New-Object System.Collections.Generic.List[object]
    if ($Mode -eq 'CheckOnly') {
        foreach ($file in @($Manifest.files)) {
            $items.Add([pscustomobject]@{
                file = $file.input
                fileRelativePath = $file.inputRelativePath
                sourcePptx = $null
                visualAuditDir = $null
                visualAuditCsv = $null
                visualAuditContactSheet = $null
                visualAuditStatus = 'SkippedCheckOnly'
                visualAuditError = ''
                visualAuditErrorCount = 0
                visualAuditWarningCount = 0
                visualConfirmationStatus = 'SkippedCheckOnly'
                visualConfirmationGateStatus = 'Skipped'
                visualConfirmationCsv = $null
                visualConfirmationManifest = $null
                visualConfirmationContactSheet = $null
                visualConfirmationError = ''
                fixedPptx = $null
                fixReport = $null
                fixStatus = 'SkippedCheckOnly'
                fixError = ''
                fixedVisualAuditDir = $null
                fixedVisualAuditCsv = $null
                fixedVisualAuditContactSheet = $null
                fixedVisualAuditErrorCount = 0
                fixedVisualAuditWarningCount = 0
                fixedVisualConfirmationStatus = 'SkippedCheckOnly'
                fixedVisualConfirmationGateStatus = 'Skipped'
                fixedVisualConfirmationCsv = $null
                fixedVisualConfirmationManifest = $null
                fixedVisualConfirmationContactSheet = $null
                fixedVisualConfirmationError = ''
            }) | Out-Null
        }
        return @($items.ToArray())
    }

    $auditRoot = Join-Path $OutputRoot '11_视觉审查'
    $fixRoot = Join-Path $OutputRoot '12_视觉修复'
    $fixedAuditRoot = Join-Path $OutputRoot '13_视觉修复审查'
    foreach ($dir in @($auditRoot, $(if ($ApplyFixes) { $fixRoot }), $(if ($ApplyFixes) { $fixedAuditRoot }))) {
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }

    foreach ($file in @($Manifest.files)) {
        $inputName = Split-Path $file.input -Leaf
        $sourcePptx = [string]$file.normalizedPptx
        $safeName = [string]$file.outputStem
        $auditDir = Join-Path $auditRoot $safeName
        $auditCsv = Join-Path $auditDir 'pptx-visual-audit.csv'
        $auditManifestPath = Join-Path $auditDir 'pptx-visual-audit-manifest.json'
        $auditContactSheet = Join-Path $auditDir 'pptx-visual-audit.contact-sheet.png'
        $auditStatus = 'SkippedNoNormalizedPptx'
        $auditError = ''
        [int]$auditErrorCount = 0
        [int]$auditWarningCount = 0
        $visualConfirmation = $null

        $fixedPptx = $null
        $fixReport = $null
        $fixStatus = if ($ApplyFixes) { 'SkippedNoAudit' } else { 'Disabled' }
        $fixError = ''
        $fixedAuditDir = $null
        $fixedAuditCsv = $null
        $fixedAuditContactSheet = $null
        [int]$fixedAuditErrorCount = 0
        [int]$fixedAuditWarningCount = 0
        $fixedVisualConfirmation = $null

        if (-not [string]::IsNullOrWhiteSpace($sourcePptx) -and (Test-Path -LiteralPath $sourcePptx)) {
            try {
                & (Join-Path $PSScriptRoot 'Export-PptxVisualAudit.ps1') -InputPath $sourcePptx -OutputDir $auditDir -ContactSheet
                $auditStatus = 'Completed'
                $auditManifest = Read-JsonObject -Path $auditManifestPath
                $auditErrorCount = [int](Get-ObjectPropertyValue -Object $auditManifest -PropertyName 'errorCount' -DefaultValue 0)
                $auditWarningCount = [int](Get-ObjectPropertyValue -Object $auditManifest -PropertyName 'warningCount' -DefaultValue 0)
                $visualConfirmation = Invoke-VisualConfirmationArtifact -VisualAuditDir $auditDir
            } catch {
                $auditStatus = 'Failed'
                $auditError = $_.Exception.Message
            }
        }

        if ($ApplyFixes -and $auditStatus -eq 'Completed' -and (Test-Path -LiteralPath $auditCsv)) {
            $fileFixDir = Join-Path $fixRoot $safeName
            if (-not (Test-Path -LiteralPath $fileFixDir)) { New-Item -ItemType Directory -Path $fileFixDir -Force | Out-Null }
            $sourceBase = [System.IO.Path]::GetFileNameWithoutExtension($sourcePptx)
            $sourceExt = [System.IO.Path]::GetExtension($sourcePptx)
            $fixedPptx = Join-Path $fileFixDir ($sourceBase + '.visual-fixed' + $sourceExt)
            $fixReport = Join-Path $fileFixDir 'pptx-visual-audit-fixes.csv'
            try {
                & (Join-Path $PSScriptRoot 'Apply-PptxVisualAuditFixes.ps1') -InputPath $sourcePptx -VisualAuditCsv $auditCsv -OutputPath $fixedPptx
                $fixStatus = if (Test-Path -LiteralPath $fixedPptx) { 'Completed' } else { 'FailedNoOutput' }
            } catch {
                $fixStatus = 'Failed'
                $fixError = $_.Exception.Message
            }

            if ($fixStatus -eq 'Completed') {
                $fixedAuditDir = Join-Path $fixedAuditRoot $safeName
                $fixedAuditCsv = Join-Path $fixedAuditDir 'pptx-visual-audit.csv'
                $fixedAuditContactSheet = Join-Path $fixedAuditDir 'pptx-visual-audit.contact-sheet.png'
                $fixedAuditManifestPath = Join-Path $fixedAuditDir 'pptx-visual-audit-manifest.json'
                try {
                    & (Join-Path $PSScriptRoot 'Export-PptxVisualAudit.ps1') -InputPath $fixedPptx -OutputDir $fixedAuditDir -ContactSheet
                    $fixedAuditManifest = Read-JsonObject -Path $fixedAuditManifestPath
                    $fixedAuditErrorCount = [int](Get-ObjectPropertyValue -Object $fixedAuditManifest -PropertyName 'errorCount' -DefaultValue 0)
                    $fixedAuditWarningCount = [int](Get-ObjectPropertyValue -Object $fixedAuditManifest -PropertyName 'warningCount' -DefaultValue 0)
                    $fixedVisualConfirmation = Invoke-VisualConfirmationArtifact -VisualAuditDir $fixedAuditDir -BaselineVisualAuditDir $auditDir
                } catch {
                    $fixStatus = 'FixedAuditFailed'
                    $fixError = $_.Exception.Message
                }
            }
        }

        $items.Add([pscustomobject]@{
            file = $file.input
            fileRelativePath = $file.inputRelativePath
            sourcePptx = $sourcePptx
            visualAuditDir = $auditDir
            visualAuditCsv = if (Test-Path -LiteralPath $auditCsv) { $auditCsv } else { $null }
            visualAuditContactSheet = if (Test-Path -LiteralPath $auditContactSheet) { $auditContactSheet } else { $null }
            visualAuditStatus = $auditStatus
            visualAuditError = $auditError
            visualAuditErrorCount = $auditErrorCount
            visualAuditWarningCount = $auditWarningCount
            visualConfirmationStatus = if ($null -ne $visualConfirmation) { $visualConfirmation.status } else { 'SkippedNoAudit' }
            visualConfirmationGateStatus = if ($null -ne $visualConfirmation) { $visualConfirmation.gateStatus } else { 'Skipped' }
            visualConfirmationCsv = if ($null -ne $visualConfirmation) { $visualConfirmation.csv } else { $null }
            visualConfirmationManifest = if ($null -ne $visualConfirmation) { $visualConfirmation.manifest } else { $null }
            visualConfirmationContactSheet = if ($null -ne $visualConfirmation) { $visualConfirmation.contactSheet } else { $null }
            visualConfirmationError = if ($null -ne $visualConfirmation) { $visualConfirmation.error } else { '' }
            visualConfirmationNeedsReviewCount = if ($null -ne $visualConfirmation) { $visualConfirmation.needsReviewCount } else { 0 }
            visualConfirmationFailedCount = if ($null -ne $visualConfirmation) { $visualConfirmation.failedCount } else { 0 }
            fixedPptx = if (-not [string]::IsNullOrWhiteSpace($fixedPptx) -and (Test-Path -LiteralPath $fixedPptx)) { $fixedPptx } else { $null }
            fixReport = if (-not [string]::IsNullOrWhiteSpace($fixReport) -and (Test-Path -LiteralPath $fixReport)) { $fixReport } else { $null }
            fixStatus = $fixStatus
            fixError = $fixError
            fixedVisualAuditDir = $fixedAuditDir
            fixedVisualAuditCsv = if (-not [string]::IsNullOrWhiteSpace($fixedAuditCsv) -and (Test-Path -LiteralPath $fixedAuditCsv)) { $fixedAuditCsv } else { $null }
            fixedVisualAuditContactSheet = if (-not [string]::IsNullOrWhiteSpace($fixedAuditContactSheet) -and (Test-Path -LiteralPath $fixedAuditContactSheet)) { $fixedAuditContactSheet } else { $null }
            fixedVisualAuditErrorCount = $fixedAuditErrorCount
            fixedVisualAuditWarningCount = $fixedAuditWarningCount
            fixedVisualConfirmationStatus = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.status } else { 'SkippedNoFixedAudit' }
            fixedVisualConfirmationGateStatus = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.gateStatus } else { 'Skipped' }
            fixedVisualConfirmationCsv = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.csv } else { $null }
            fixedVisualConfirmationManifest = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.manifest } else { $null }
            fixedVisualConfirmationContactSheet = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.contactSheet } else { $null }
            fixedVisualConfirmationError = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.error } else { '' }
            fixedVisualConfirmationNeedsReviewCount = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.needsReviewCount } else { 0 }
            fixedVisualConfirmationFailedCount = if ($null -ne $fixedVisualConfirmation) { $fixedVisualConfirmation.failedCount } else { 0 }
        }) | Out-Null
    }

    return @($items.ToArray())
}

function Invoke-FormulaOmmlArtifacts {
    param(
        $Manifest,
        [string]$FormulaReviewIndex,
        [string]$OutputRoot,
        [int]$MaxItems,
        [bool]$RunVisualAudit
    )

    $copyRoot = Join-Path $OutputRoot '14_公式OMML副本'
    $candidateDir = Join-Path $copyRoot '_candidates'
    $reviewRoot = Join-Path $OutputRoot '15_公式OMML审查'
    foreach ($dir in @($copyRoot, $candidateDir, $reviewRoot)) {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    }

    $candidateCsv = Join-Path $candidateDir 'formula-omml-candidates.csv'
    $candidateManifestPath = Join-Path $candidateDir 'formula-omml-candidates-manifest.json'
    $result = [ordered]@{
        formulaReviewCsv = $FormulaReviewIndex
        candidateDir = $candidateDir
        candidateCsv = $null
        candidateStatus = 'Pending'
        candidateError = ''
        candidateCount = 0
        generatedCandidateCount = 0
        failedCandidateCount = 0
        visualAuditEnabled = $RunVisualAudit
        files = @()
    }

    if ([string]::IsNullOrWhiteSpace($FormulaReviewIndex) -or -not (Test-Path -LiteralPath $FormulaReviewIndex)) {
        $result.candidateStatus = 'SkippedNoFormulaReviewCsv'
        return [pscustomobject]$result
    }

    try {
        & (Join-Path $PSScriptRoot 'Export-FormulaOmmlCandidates.ps1') -FormulaReviewCsv $FormulaReviewIndex -OutputDir $candidateDir -MaxItems $MaxItems
        $candidateManifest = Read-JsonObject -Path $candidateManifestPath
        $result.candidateCsv = if (Test-Path -LiteralPath $candidateCsv) { $candidateCsv } else { $null }
        $result.candidateStatus = 'Completed'
        $result.candidateCount = [int](Get-ObjectPropertyValue -Object $candidateManifest -PropertyName 'candidateCount' -DefaultValue 0)
        $result.generatedCandidateCount = [int](Get-ObjectPropertyValue -Object $candidateManifest -PropertyName 'generatedCount' -DefaultValue 0)
        $result.failedCandidateCount = [int](Get-ObjectPropertyValue -Object $candidateManifest -PropertyName 'failedCount' -DefaultValue 0)
    } catch {
        $result.candidateStatus = 'Failed'
        $result.candidateError = $_.Exception.Message
        return [pscustomobject]$result
    }

    if (-not (Test-Path -LiteralPath $candidateCsv)) {
        $result.candidateStatus = 'FailedNoCandidateCsv'
        return [pscustomobject]$result
    }

    $dotnet = Resolve-DotNetCommand
    $validatorProject = Join-Path $PSScriptRoot 'FormulaOfficeMathValidator\FormulaOfficeMathValidator.csproj'
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($file in @($Manifest.files)) {
        $inputName = Split-Path $file.input -Leaf
        $safeName = [string]$file.outputStem
        $sourcePptx = ''
        $sourceKind = 'Normalized'

        $visualArtifacts = @()
        if ($Manifest.PSObject.Properties.Name -contains 'visualAuditArtifacts') {
            $visualArtifacts = @($Manifest.visualAuditArtifacts | Where-Object {
                (Get-ObjectPropertyValue -Object $_ -PropertyName 'file' -DefaultValue '') -eq $file.input
            })
        }
        foreach ($visual in $visualArtifacts) {
            $fixed = Get-ObjectPropertyValue -Object $visual -PropertyName 'fixedPptx' -DefaultValue ''
            if (-not [string]::IsNullOrWhiteSpace($fixed) -and (Test-Path -LiteralPath $fixed)) {
                $sourcePptx = $fixed
                $sourceKind = 'VisualFixed'
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($sourcePptx)) {
            $sourcePptx = [string]$file.normalizedPptx
        }

        $fileCopyDir = Join-Path $copyRoot $safeName
        $fileReviewDir = Join-Path $reviewRoot $safeName
        foreach ($dir in @($fileCopyDir, $fileReviewDir)) {
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        }

        $outputPptx = $null
        $applyReport = Join-Path $fileCopyDir 'formula-omml-apply-report.csv'
        $validatorJson = Join-Path $fileReviewDir 'formula-omml-validator.json'
        $visualAuditCsv = Join-Path $fileReviewDir 'pptx-visual-audit.csv'
        $visualAuditContactSheet = Join-Path $fileReviewDir 'pptx-visual-audit.contact-sheet.png'

        $applyStatus = 'SkippedNoSourcePptx'
        $applyError = ''
        [int]$insertedCount = 0
        [int]$applyIssueCount = 0
        [int]$fileCandidateCount = 0
        $validatorStatus = 'SkippedNoOutput'
        $validatorError = ''
        [int]$openXmlErrorCount = 0
        [int]$a14MathCount = 0
        [int]$officeMathCount = 0
        $visualAuditStatus = 'SkippedNoOutput'
        $visualAuditError = ''
        [int]$visualAuditErrorCount = 0
        [int]$visualAuditWarningCount = 0
        $visualConfirmation = $null
        $baselineVisualAuditDir = ''
        [int]$baselineVisualAuditErrorCount = 0
        [int]$baselineVisualAuditWarningCount = 0
        $automationGateStatus = 'Skipped'
        $automationGateReason = ''

        try {
            $fileCandidateCount = @(
                Import-Csv -LiteralPath $FormulaReviewIndex -Encoding UTF8 |
                    Where-Object {
                        $_.SuggestedAction -eq 'ReviewWhitelistConversion' -and (
                            ($null -ne $_.PSObject.Properties['FilePath'] -and $_.FilePath -eq $file.input) -or
                            ($null -ne $_.PSObject.Properties['FileRelativePath'] -and $_.FileRelativePath -eq $file.inputRelativePath)
                        )
                    }
            ).Count
        } catch {
            $fileCandidateCount = 0
        }

        if ($fileCandidateCount -le 0) {
            $applyStatus = 'SkippedNoFormulaWhitelistCandidate'
        } elseif (-not [string]::IsNullOrWhiteSpace($sourcePptx) -and (Test-Path -LiteralPath $sourcePptx)) {
            $sourceBase = [System.IO.Path]::GetFileNameWithoutExtension($sourcePptx)
            $sourceExt = [System.IO.Path]::GetExtension($sourcePptx)
            $outputPptx = Join-Path $fileCopyDir ($sourceBase + '.formula-omml' + $sourceExt)
            try {
                & (Join-Path $PSScriptRoot 'Apply-FormulaOmmlWhitelist.ps1') -InputPath $sourcePptx -FormulaReviewCsv $FormulaReviewIndex -OmmlCandidateCsv $candidateCsv -OutputPath $outputPptx -ReportPath $applyReport -ReviewFileName $file.input -MaxItems $MaxItems
                $applyStatus = if (Test-Path -LiteralPath $outputPptx) { 'Completed' } else { 'FailedNoOutput' }
            } catch {
                $applyStatus = 'Failed'
                $applyError = $_.Exception.Message
            }

            if (Test-Path -LiteralPath $applyReport) {
                $applyRows = @(Import-Csv -LiteralPath $applyReport -Encoding UTF8)
                $insertedCount = @($applyRows | Where-Object { $_.Issue -eq 'FormulaOmmlInserted' }).Count
                $applyIssueCount = @($applyRows | Where-Object { $_.Issue -notin @('FormulaOmmlInserted', 'SavedAs') }).Count
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($outputPptx) -and (Test-Path -LiteralPath $outputPptx)) {
            if ([string]::IsNullOrWhiteSpace($dotnet)) {
                $validatorStatus = 'SkippedDotNetMissing'
            } elseif (-not (Test-Path -LiteralPath $validatorProject)) {
                $validatorStatus = 'SkippedValidatorMissing'
            } else {
                try {
                    $validatorOutput = & $dotnet run --project $validatorProject -c Release -- $outputPptx --max-errors 20 2>&1
                    $validatorText = ($validatorOutput | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
                    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
                    [System.IO.File]::WriteAllText($validatorJson, $validatorText, $utf8Bom)
                    $jsonStart = $validatorText.IndexOf('{')
                    $jsonEnd = $validatorText.LastIndexOf('}')
                    if ($jsonStart -lt 0 -or $jsonEnd -lt $jsonStart) {
                        throw 'Validator output did not contain JSON.'
                    }
                    $validatorJsonObject = $validatorText.Substring($jsonStart, $jsonEnd - $jsonStart + 1) | ConvertFrom-Json
                    $openXmlErrorCount = [int](Get-ObjectPropertyValue -Object $validatorJsonObject -PropertyName 'OpenXmlErrorCount' -DefaultValue 0)
                    $a14MathCount = [int](Get-ObjectPropertyValue -Object $validatorJsonObject -PropertyName 'A14MathCount' -DefaultValue 0)
                    $officeMathCount = [int](Get-ObjectPropertyValue -Object $validatorJsonObject -PropertyName 'OfficeMathCount' -DefaultValue 0)
                    $validatorStatus = if ($LASTEXITCODE -eq 0) { 'Completed' } else { 'Failed' }
                } catch {
                    $validatorStatus = 'Failed'
                    $validatorError = $_.Exception.Message
                }
            }

            if ($RunVisualAudit) {
                try {
                    & (Join-Path $PSScriptRoot 'Export-PptxVisualAudit.ps1') -InputPath $outputPptx -OutputDir $fileReviewDir -ContactSheet
                    $visualManifest = Read-JsonObject -Path (Join-Path $fileReviewDir 'pptx-visual-audit-manifest.json')
                    $visualAuditErrorCount = [int](Get-ObjectPropertyValue -Object $visualManifest -PropertyName 'errorCount' -DefaultValue 0)
                    $visualAuditWarningCount = [int](Get-ObjectPropertyValue -Object $visualManifest -PropertyName 'warningCount' -DefaultValue 0)
                    $visualAuditStatus = 'Completed'
                } catch {
                    $visualAuditStatus = 'Failed'
                    $visualAuditError = $_.Exception.Message
                }
            } else {
                $visualAuditStatus = 'SkippedFastFormulaProfile'
            }
        }

        foreach ($visual in $visualArtifacts) {
            if ($sourceKind -eq 'VisualFixed') {
                $baselineVisualAuditErrorCount = [int](Get-ObjectPropertyValue -Object $visual -PropertyName 'fixedVisualAuditErrorCount' -DefaultValue 0)
                $baselineVisualAuditWarningCount = [int](Get-ObjectPropertyValue -Object $visual -PropertyName 'fixedVisualAuditWarningCount' -DefaultValue 0)
                $baselineVisualAuditDir = [string](Get-ObjectPropertyValue -Object $visual -PropertyName 'fixedVisualAuditDir' -DefaultValue '')
            } else {
                $baselineVisualAuditErrorCount = [int](Get-ObjectPropertyValue -Object $visual -PropertyName 'visualAuditErrorCount' -DefaultValue 0)
                $baselineVisualAuditWarningCount = [int](Get-ObjectPropertyValue -Object $visual -PropertyName 'visualAuditWarningCount' -DefaultValue 0)
                $baselineVisualAuditDir = [string](Get-ObjectPropertyValue -Object $visual -PropertyName 'visualAuditDir' -DefaultValue '')
            }
            break
        }

        if ($RunVisualAudit -and $visualAuditStatus -eq 'Completed') {
            $visualConfirmation = Invoke-VisualConfirmationArtifact -VisualAuditDir $fileReviewDir -BaselineVisualAuditDir $baselineVisualAuditDir
        }

        if ($fileCandidateCount -le 0) {
            $automationGateStatus = 'Skipped'
            $automationGateReason = 'No formula whitelist candidate for this file.'
        } elseif ($applyStatus -ne 'Completed') {
            $automationGateStatus = 'Failed'
            $automationGateReason = "OMML apply status is $applyStatus."
        } elseif ($insertedCount -le 0) {
            $automationGateStatus = 'Failed'
            $automationGateReason = 'No OMML formula was inserted.'
        } elseif ($applyIssueCount -gt 0) {
            $automationGateStatus = 'Failed'
            $automationGateReason = "OMML apply report has $applyIssueCount issue(s)."
        } elseif ($validatorStatus -ne 'Completed' -or $openXmlErrorCount -gt 0) {
            $automationGateStatus = 'Failed'
            $automationGateReason = "Open XML validator status is $validatorStatus with $openXmlErrorCount error(s)."
        } elseif (-not $RunVisualAudit) {
            $automationGateStatus = 'Passed'
            $automationGateReason = 'Fast formula profile passed: whitelist, OMML write, and Open XML validation passed; rendered visual audit was skipped.'
        } elseif ($visualAuditStatus -ne 'Completed') {
            $automationGateStatus = 'Failed'
            $automationGateReason = "Visual audit status is $visualAuditStatus."
        } elseif ($null -eq $visualConfirmation -or $visualConfirmation.gateStatus -ne 'Passed') {
            $gate = if ($null -eq $visualConfirmation) { 'Skipped' } else { $visualConfirmation.gateStatus }
            $status = if ($null -eq $visualConfirmation) { 'Skipped' } else { $visualConfirmation.status }
            $automationGateStatus = 'Failed'
            $automationGateReason = "Visual confirmation gate is $gate ($status)."
        } elseif ($visualAuditErrorCount -gt $baselineVisualAuditErrorCount) {
            $automationGateStatus = 'Failed'
            $automationGateReason = "Visual audit errors increased from $baselineVisualAuditErrorCount to $visualAuditErrorCount."
        } else {
            $automationGateStatus = 'Passed'
            $automationGateReason = 'Whitelist, Open XML validation, PowerPoint render, visual confirmation, and visual audit baseline gate passed.'
        }

        $items.Add([pscustomobject]@{
            file = $file.input
            fileRelativePath = $file.inputRelativePath
            sourcePptx = $sourcePptx
            sourceKind = $sourceKind
            outputPptx = if (-not [string]::IsNullOrWhiteSpace($outputPptx) -and (Test-Path -LiteralPath $outputPptx)) { $outputPptx } else { $null }
            applyReport = if (Test-Path -LiteralPath $applyReport) { $applyReport } else { $null }
            applyStatus = $applyStatus
            applyError = $applyError
            fileCandidateCount = $fileCandidateCount
            insertedCount = $insertedCount
            applyIssueCount = $applyIssueCount
            validatorJson = if (Test-Path -LiteralPath $validatorJson) { $validatorJson } else { $null }
            validatorStatus = $validatorStatus
            validatorError = $validatorError
            openXmlErrorCount = $openXmlErrorCount
            a14MathCount = $a14MathCount
            officeMathCount = $officeMathCount
            visualAuditDir = $fileReviewDir
            visualAuditCsv = if (Test-Path -LiteralPath $visualAuditCsv) { $visualAuditCsv } else { $null }
            visualAuditContactSheet = if (Test-Path -LiteralPath $visualAuditContactSheet) { $visualAuditContactSheet } else { $null }
            visualAuditStatus = $visualAuditStatus
            visualAuditError = $visualAuditError
            visualAuditErrorCount = $visualAuditErrorCount
            visualAuditWarningCount = $visualAuditWarningCount
            visualConfirmationStatus = if ($null -ne $visualConfirmation) { $visualConfirmation.status } else { 'SkippedNoAudit' }
            visualConfirmationGateStatus = if ($null -ne $visualConfirmation) { $visualConfirmation.gateStatus } else { 'Skipped' }
            visualConfirmationCsv = if ($null -ne $visualConfirmation) { $visualConfirmation.csv } else { $null }
            visualConfirmationManifest = if ($null -ne $visualConfirmation) { $visualConfirmation.manifest } else { $null }
            visualConfirmationContactSheet = if ($null -ne $visualConfirmation) { $visualConfirmation.contactSheet } else { $null }
            visualConfirmationError = if ($null -ne $visualConfirmation) { $visualConfirmation.error } else { '' }
            visualConfirmationNeedsReviewCount = if ($null -ne $visualConfirmation) { $visualConfirmation.needsReviewCount } else { 0 }
            visualConfirmationFailedCount = if ($null -ne $visualConfirmation) { $visualConfirmation.failedCount } else { 0 }
            visualConfirmationNewErrorCount = if ($null -ne $visualConfirmation) { $visualConfirmation.newErrorCount } else { 0 }
            visualConfirmationNewWarningCount = if ($null -ne $visualConfirmation) { $visualConfirmation.newWarningCount } else { 0 }
            baselineVisualAuditDir = $baselineVisualAuditDir
            baselineVisualAuditErrorCount = $baselineVisualAuditErrorCount
            baselineVisualAuditWarningCount = $baselineVisualAuditWarningCount
            automationGateStatus = $automationGateStatus
            automationGateReason = $automationGateReason
        }) | Out-Null
    }

    $result.files = @($items.ToArray())
    return [pscustomobject]$result
}

$root = Split-Path -Parent $PSScriptRoot
$inputItem = Get-Item -LiteralPath $InputPath
$inputFullPath = $inputItem.FullName
if ($inputItem.PSIsContainer) {
    $identityMap = Get-IdentityMap -InputRoot $inputFullPath -Files @(Get-PptFiles -Path $inputFullPath -Pattern $FilePattern -Recurse:$Recurse)
} else {
    $identityMap = @{}
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Get-DefaultOutputRoot -InputItem $inputItem
}
$OutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)

$reportDir = Join-Path $OutputRoot '00_检查报告'
$normalizedDir = Join-Path $OutputRoot '01_规范化PPTX'
$pdfDir = Join-Path $OutputRoot '02_导出PDF'
$backupDir = Join-Path $OutputRoot '03_原始备份'
$imageDir = Join-Path $OutputRoot '04_页面图片'
$contactSheetDir = Join-Path $OutputRoot '05_页面总览'
$reviewSheetDir = Join-Path $OutputRoot '06_重点复核'
$sourceImageDir = Join-Path $OutputRoot '07_原始页面图片'
$beforeAfterDir = Join-Path $OutputRoot '08_前后对比'
$reviewPagePackageDir = Join-Path $OutputRoot '09_重点单页'
$reviewIndexDir = Join-Path $OutputRoot '10_复核索引'
$visualAuditDir = Join-Path $OutputRoot '11_视觉审查'
$visualFixDir = Join-Path $OutputRoot '12_视觉修复'
$visualFixAuditDir = Join-Path $OutputRoot '13_视觉修复审查'
$formulaOmmlDir = Join-Path $OutputRoot '14_公式OMML副本'
$formulaOmmlAuditDir = Join-Path $OutputRoot '15_公式OMML审查'
$exportsPdf = $Mode -notin @('CheckOnly', 'SafeNormalize')
$outputDirs = @($OutputRoot, $reportDir, $normalizedDir, $backupDir)
if ($exportsPdf) {
    $outputDirs += @($pdfDir)
}
if ($IncludeReviewArtifacts) {
    $outputDirs += @($imageDir, $contactSheetDir, $reviewSheetDir, $sourceImageDir, $beforeAfterDir, $reviewPagePackageDir, $reviewIndexDir)
}
if ($IncludeVisualAudit) {
    $outputDirs += @($visualAuditDir)
    if ($ApplyVisualAuditFixes) {
        $outputDirs += @($visualFixDir, $visualFixAuditDir)
    }
}
if ($ApplyFormulaOmmlWhitelist) {
    $outputDirs += @($formulaOmmlDir, $formulaOmmlAuditDir)
}
foreach ($dir in $outputDirs) {
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

$files = @(Get-PptFiles -Path $inputFullPath -Pattern $FilePattern -Recurse:$Recurse)
if ($files.Count -eq 0) { throw "No .pptx/.pptm files found in $inputFullPath" }
if ($identityMap.Count -eq 0) {
    $identityMap = Get-IdentityMap -InputRoot $inputFullPath -Files $files
}

$stepCount = 5
if ($IncludeVisualAudit) { $stepCount++ }
if ($ApplyFormulaOmmlWhitelist) { $stepCount++ }

Write-Host "Step 1/${stepCount}: self-check"
& (Join-Path $PSScriptRoot 'Test-ToolkitFiles.ps1')
& (Join-Path $PSScriptRoot 'Assert-Toolchain.ps1')
Assert-PowerPointAutomationReady

$reportOnlyPath = Join-Path $reportDir 'physics-ppt-report-only.csv'
$normalizeReportPath = Join-Path $normalizedDir 'physics-ppt-normalize-report.csv'
$finalReportPath = Join-Path $reportDir 'physics-ppt-normalize-report.csv'

if ($Mode -eq 'CheckOnly' -or ($Mode -ne 'ForceRebuild' -and -not $SkipPreflightReport)) {
    Write-Host "Step 2/${stepCount}: report-only inspection"
    $reportArgs = @{
        InputPath = $inputFullPath
        OutputDir = $reportDir
        Recurse = [bool]$Recurse
        ReportOnly = $true
        FilePattern = $FilePattern
    }
    & (Join-Path $PSScriptRoot 'Normalize-PhysicsPpt.ps1') @reportArgs
    $generatedReport = Join-Path $reportDir 'physics-ppt-normalize-report.csv'
    if (Test-Path -LiteralPath $generatedReport) {
        Move-Item -LiteralPath $generatedReport -Destination $reportOnlyPath -Force
    }
} elseif ($Mode -eq 'ForceRebuild') {
    Write-Host "Step 2/${stepCount}: report-only inspection skipped for ForceRebuild"
} else {
    Write-Host "Step 2/${stepCount}: report-only inspection skipped for fast profile"
}

if ($Mode -eq 'CheckOnly') {
    Write-Host "Step 3/${stepCount}: normalization skipped"
    $finalReportPath = $reportOnlyPath
} else {
    Write-Host "Step 3/${stepCount}: normalize"
    $normalizeArgs = @{
        InputPath = $inputFullPath
        OutputDir = $normalizedDir
        Recurse = [bool]$Recurse
        UpdateMaster = [bool]$UpdateMaster
        NoPdf = ($Mode -eq 'SafeNormalize')
        Force = ($Mode -eq 'ForceRebuild')
        FilePattern = $FilePattern
    }
    if ($IncludeReviewArtifacts) {
        $normalizeArgs.ImageOutputDir = $imageDir
    }
    & (Join-Path $PSScriptRoot 'Normalize-PhysicsPpt.ps1') @normalizeArgs

    if (Test-Path -LiteralPath $normalizeReportPath) {
        Copy-Item -LiteralPath $normalizeReportPath -Destination $finalReportPath -Force
    }

    Get-ChildItem -LiteralPath $normalizedDir -Filter '*.normalized.pdf' -File -ErrorAction SilentlyContinue |
        Move-Item -Destination $pdfDir -Force

    $sourceBackupDir = Join-Path $normalizedDir '_backup_originals'
    if (Test-Path -LiteralPath $sourceBackupDir) {
        Get-ChildItem -LiteralPath $sourceBackupDir -File | Move-Item -Destination $backupDir -Force
        Remove-Item -LiteralPath $sourceBackupDir -Force
    }
}

Write-Host "Step 4/${stepCount}: build review artifacts and manifest"
$rows = Convert-ReportCsv -Path $finalReportPath
$sourceImageDirs = if ($Mode -eq 'CheckOnly' -or -not $IncludeReviewArtifacts) { @{} } else { Export-SourcePageImages -Files $files -OutputRoot $OutputRoot -IdentityMap $identityMap }
$contactSheets = if ($Mode -eq 'CheckOnly' -or -not $IncludeReviewArtifacts) { @{} } else { New-ContactSheets -Files $files -OutputRoot $OutputRoot -IdentityMap $identityMap }
$reviewSheets = if ($Mode -eq 'CheckOnly' -or -not $IncludeReviewArtifacts) { @{} } else { New-ReviewContactSheets -Files $files -OutputRoot $OutputRoot -Rows $rows -IdentityMap $identityMap }
$beforeAfterSheets = if ($Mode -eq 'CheckOnly' -or -not $IncludeReviewArtifacts) { @{} } else { New-BeforeAfterReviewSheets -Files $files -OutputRoot $OutputRoot -Rows $rows -SourceImageDirs $sourceImageDirs -IdentityMap $identityMap }
$reviewPagePackages = if ($Mode -eq 'CheckOnly' -or -not $IncludeReviewArtifacts) { @{} } else { New-ReviewPagePackages -Files $files -OutputRoot $OutputRoot -Rows $rows -IdentityMap $identityMap }
$reviewIndexes = if ($Mode -eq 'CheckOnly' -or -not $IncludeReviewArtifacts) { @{} } else { New-ReviewIndexes -Files $files -OutputRoot $OutputRoot -Rows $rows -SourceImageDirs $sourceImageDirs -IdentityMap $identityMap }
$formulaReviewIndex = New-FormulaReviewIndex -Rows $rows -OutputRoot $OutputRoot -IdentityMap $identityMap
$manifestPath = Join-Path $OutputRoot 'review-manifest.json'
$summaryPath = Join-Path $OutputRoot 'summary.md'
$manifest = New-Manifest -Files $files -InputFullPath $inputFullPath -OutputRoot $OutputRoot -ReportPath $finalReportPath -Mode $Mode -IncludeReviewArtifacts ([bool]$IncludeReviewArtifacts) -ReportRows $rows -ContactSheets $contactSheets -ReviewSheets $reviewSheets -SourceImageDirs $sourceImageDirs -BeforeAfterSheets $beforeAfterSheets -ReviewPagePackages $reviewPagePackages -ReviewIndexes $reviewIndexes -FormulaReviewIndex $formulaReviewIndex -IdentityMap $identityMap

$formulaWhitelistSuggestionArtifacts = Invoke-FormulaWhitelistSuggestionArtifacts -FormulaReviewIndex $formulaReviewIndex -OutputRoot $OutputRoot -MaxSuggestions 120
$manifest | Add-Member -NotePropertyName formulaWhitelistSuggestionArtifacts -NotePropertyValue $formulaWhitelistSuggestionArtifacts -Force

if ($IncludeVisualAudit) {
    $currentStep = 5
    Write-Host "Step ${currentStep}/${stepCount}: visual audit and low-risk fixes"
    $visualAuditArtifacts = @(Invoke-VisualAuditArtifacts -Manifest $manifest -OutputRoot $OutputRoot -ApplyFixes ([bool]$ApplyVisualAuditFixes) -Mode $Mode)
    $manifest | Add-Member -NotePropertyName visualAuditEnabled -NotePropertyValue ([bool]$IncludeVisualAudit) -Force
    $manifest | Add-Member -NotePropertyName visualAuditFixesEnabled -NotePropertyValue ([bool]$ApplyVisualAuditFixes) -Force
    $manifest | Add-Member -NotePropertyName visualAuditArtifacts -NotePropertyValue $visualAuditArtifacts -Force
    $currentStep++
} else {
    $currentStep = 5
    $manifest | Add-Member -NotePropertyName visualAuditEnabled -NotePropertyValue $false -Force
    $manifest | Add-Member -NotePropertyName visualAuditFixesEnabled -NotePropertyValue $false -Force
    $manifest | Add-Member -NotePropertyName visualAuditArtifacts -NotePropertyValue @() -Force
}

if ($ApplyFormulaOmmlWhitelist) {
    Write-Host "Step ${currentStep}/${stepCount}: formula OMML whitelist copies"
    $formulaOmmlArtifacts = Invoke-FormulaOmmlArtifacts -Manifest $manifest -FormulaReviewIndex $formulaReviewIndex -OutputRoot $OutputRoot -MaxItems $FormulaOmmlMaxItems -RunVisualAudit ([bool]$FormulaOmmlVisualAudit)
    $manifest | Add-Member -NotePropertyName formulaOmmlEnabled -NotePropertyValue $true -Force
    $manifest | Add-Member -NotePropertyName formulaOmmlMaxItems -NotePropertyValue $FormulaOmmlMaxItems -Force
    $manifest | Add-Member -NotePropertyName formulaOmmlVisualAuditEnabled -NotePropertyValue ([bool]$FormulaOmmlVisualAudit) -Force
    $manifest | Add-Member -NotePropertyName formulaOmmlArtifacts -NotePropertyValue $formulaOmmlArtifacts -Force
} else {
    $manifest | Add-Member -NotePropertyName formulaOmmlEnabled -NotePropertyValue $false -Force
    $manifest | Add-Member -NotePropertyName formulaOmmlMaxItems -NotePropertyValue $FormulaOmmlMaxItems -Force
    $manifest | Add-Member -NotePropertyName formulaOmmlVisualAuditEnabled -NotePropertyValue ([bool]$FormulaOmmlVisualAudit) -Force
    $manifest | Add-Member -NotePropertyName formulaOmmlArtifacts -NotePropertyValue $null -Force
}

Write-Host "Step ${stepCount}/${stepCount}: write summary and manifest"
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Summary -Path $summaryPath -Mode $Mode -InputFullPath $inputFullPath -OutputRoot $OutputRoot -ReportPath $finalReportPath -ManifestPath $manifestPath -Files $files -IncludeReviewArtifacts ([bool]$IncludeReviewArtifacts) -Rows $rows -Manifest $manifest

Write-Host "Done"
Write-Host "Output: $OutputRoot"
Write-Host "Summary: $summaryPath"
Write-Host "Manifest: $manifestPath"

if ($OpenGeneratedPptx) {
    Open-GeneratedPptxResult -Manifest $manifest -NormalizedDir $normalizedDir -VisualFixDir $visualFixDir -OutputRoot $OutputRoot
}

if ($OpenOutput) {
    Invoke-Item -LiteralPath $OutputRoot
}
