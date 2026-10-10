<#
.SYNOPSIS
  Review-artifact builders and evidence metrics for Invoke-PhysicsPptWorkflow.

.DESCRIPTION
  Verbatim extraction of the workflow's review-artifact block: input identity
  mapping, page/contact-sheet/review-package/index builders, source page-image
  export, and page/file metric helpers. Like PhysicsPpt.Common.ps1 this file
  must stay function-definitions only (no top-level statements); it depends on
  PhysicsPpt.Common.ps1 being dot-sourced first and resolves everything else
  through the dot-sourcing workflow script's scope.
#>
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

    Initialize-SystemDrawing
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

    Initialize-SystemDrawing
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

function Export-SourcePageImages {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$OutputRoot,
        $IdentityMap
    )

    $sourceImageRoot = Join-Path $OutputRoot '07_原始页面图片'
    $items = @{}
    $pp = $null
    $presentations = $null
    try {
        $pp = New-PowerPointApplication
        try { $presentations = $pp.Presentations } catch { $presentations = $null }
        foreach ($file in $Files) {
            $identity = $IdentityMap[$file.FullName]
            $imageDir = Join-Path $sourceImageRoot $identity.safeStem
            $pres = $null
            try {
                if (-not (Test-Path -LiteralPath $imageDir)) { New-Item -ItemType Directory -Path $imageDir -Force | Out-Null }
                Get-ChildItem -LiteralPath $imageDir -Filter '*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force
                # Read-only open: this pass only renders page images and never saves.
                $pres = Invoke-WithComRetry -Action { $presentations.Open($file.FullName, -1, 0, 0) }
                $expectedSlides = [int]$pres.Slides.Count
                $failedSlides = New-Object System.Collections.Generic.List[int]
                for ($slideNo = 1; $slideNo -le $expectedSlides; $slideNo++) {
                    $target = Join-Path $imageDir ('page-{0:000}.png' -f $slideNo)
                    try {
                        $slide = $pres.Slides.Item($slideNo)
                        Invoke-WithComRetry -Action { $slide.Export($target, 'PNG') } | Out-Null
                        if (-not (Test-Path -LiteralPath $target) -or (Get-Item -LiteralPath $target).Length -le 0) {
                            throw "PowerPoint did not create a non-empty PNG: $target"
                        }
                        $imageInfo = Get-BasicImageInfo -Path $target
                        if ($imageInfo.Width -le 0 -or $imageInfo.Height -le 0) {
                            throw "PowerPoint created an undecodable PNG: $target"
                        }
                    } catch {
                        $failedSlides.Add($slideNo) | Out-Null
                    }
                }
                $actualSlides = @(Get-ChildItem -LiteralPath $imageDir -Filter 'page-*.png' -File | ForEach-Object {
                    if ($_.BaseName -match '^page-(\d+)$') { [int]$Matches[1] }
                } | Sort-Object -Unique)
                $expectedKey = if ($expectedSlides -gt 0) { ((1..$expectedSlides) -join ',') } else { '' }
                $actualKey = (($actualSlides | ForEach-Object { [string]$_ }) -join ',')
                if ($failedSlides.Count -gt 0 -or $actualKey -ne $expectedKey) {
                    throw "Original page image set is incomplete: expected=$expectedKey; actual=$actualKey; failed=$($failedSlides -join ',')."
                }
                $items[$file.FullName] = $imageDir
            } catch {
                Write-Warning "Original page image export failed for $($file.Name): $($_.Exception.Message)"
            } finally {
                if ($null -ne $pres) {
                    try { $pres.Close() | Out-Null } catch { }
                    Release-ComObjectSafe -ComObject $pres
                }
            }
        }
    } catch {
        Write-Warning "PowerPoint original export setup failed: $($_.Exception.Message)"
    } finally {
        if ($null -ne $pp) {
            try { $pp.Quit() | Out-Null } catch { }
        }
        Release-ComObjectSafe -ComObject $presentations
        Release-ComObjectSafe -ComObject $pp
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

    Initialize-SystemDrawing
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
        $indexRows = @(
            $reviewSlides | ForEach-Object {
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
            }
        )
        Write-Utf8BomCsv -InputObject $indexRows -Path $indexPath
        $items[$file.FullName] = $indexPath
    }
    return $items
}

function Get-ExpectedSlideCount {
    param($Rows, [string]$NormalizedPptxPath = '')
    $slides = @($Rows | Where-Object { $_.Issue -eq 'SlideType' -and $_.Slide -match '^\d+$' } | Select-Object -ExpandProperty Slide -Unique)
    if ($slides.Count -gt 0) { return $slides.Count }
    # A cache-hit rerun records only SkippedUpToDate (no per-slide rows). The
    # cache record next to the normalized PPTX carries expectedSlides; without
    # this fallback the rerun would falsely fail the page-count validation.
    if (-not [string]::IsNullOrWhiteSpace($NormalizedPptxPath) -and (Test-Path -LiteralPath $NormalizedPptxPath)) {
        $cachePath = "$NormalizedPptxPath.cache.json"
        if (Test-Path -LiteralPath $cachePath) {
            try {
                $cache = Get-Content -LiteralPath $cachePath -Raw -Encoding UTF8 | ConvertFrom-Json
                $cachedCount = 0
                if ($null -ne $cache -and $null -ne $cache.PSObject.Properties['expectedSlides']) {
                    [void][int]::TryParse([string]$cache.expectedSlides, [ref]$cachedCount)
                }
                if ($cachedCount -gt 0) { return $cachedCount }
            } catch { }
        }
    }
    return $null
}

function Get-PageImageCount {
    param([string]$ImageDir)
    if ([string]::IsNullOrWhiteSpace($ImageDir) -or -not (Test-Path -LiteralPath $ImageDir)) { return 0 }
    return @(Get-ChildItem -LiteralPath $ImageDir -Filter 'page-*.png' -File).Count
}

# Test-UsablePageImage / Test-PageImageSet are shared and come from
# PhysicsPpt.Common.ps1 (dot-sourced above) — do not re-declare them here.

function Get-FileSha256OrNull {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    try { return (Get-FileSha256Hex -Path $Path) } catch { return '' }
}

function Get-FileLength {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return 0 }
    return (Get-Item -LiteralPath $Path).Length
}

function Get-PdfPageCount {
    param([string]$PdfPath)
    if ([string]::IsNullOrWhiteSpace($PdfPath) -or -not (Test-Path -LiteralPath $PdfPath)) { return $null }
    try {
        $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($PdfPath))
        $count = ([regex]::Matches($text, '/Type\s*/Page\b')).Count
        if ($count -le 0) { return $null }
        return $count
    } catch {
        return $null
    }
}

# Get-ImageDeltaPercent / Get-ImageWhitePercent live in PhysicsPpt.Common.ps1
# so the workflow overview and the visual confirmation gate share one metric
# (same thumbnail resolution, same 8% threshold semantics).

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
