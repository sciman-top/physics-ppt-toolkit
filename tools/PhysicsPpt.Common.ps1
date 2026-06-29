<#
.SYNOPSIS
  Shared low-level helpers for PPTX package scripts.

.DESCRIPTION
  Keep this file limited to deterministic helpers with no top-level side
  effects. The production scripts dot-source it to avoid duplicated ZIP and
  Open XML relationship handling.
#>

function Read-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return '' }
    $stream = $null
    $reader = $null
    try {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        return $reader.ReadToEnd()
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Resolve-PackagePath {
    param([string]$PackagePath)

    if ([string]::IsNullOrWhiteSpace($PackagePath)) { return '' }
    $normalized = $PackagePath -replace '\\', '/'
    if ($normalized -match '^[A-Za-z][A-Za-z0-9+.-]*:') { return '' }
    $normalized = $normalized.TrimStart('/')

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($normalized -split '/')) {
        if ([string]::IsNullOrWhiteSpace($part) -or $part -eq '.') { continue }
        if ($part -eq '..') {
            if ($parts.Count -eq 0) { return '' }
            $parts.RemoveAt($parts.Count - 1)
            continue
        }
        $parts.Add($part) | Out-Null
    }
    return ($parts -join '/')
}

function Resolve-PackageTarget {
    param(
        [string]$SourcePart,
        [string]$Target
    )

    if ([string]::IsNullOrWhiteSpace($Target)) { return '' }
    $targetPath = $Target -replace '\\', '/'
    if ($targetPath -match '^[A-Za-z][A-Za-z0-9+.-]*:') { return '' }

    if ($targetPath.StartsWith('/')) {
        $combined = $targetPath.TrimStart('/')
    } else {
        $sourceDir = ''
        $lastSlash = $SourcePart.LastIndexOf('/')
        if ($lastSlash -ge 0) { $sourceDir = $SourcePart.Substring(0, $lastSlash) }
        $combined = if ([string]::IsNullOrWhiteSpace($sourceDir)) { $targetPath } else { "$sourceDir/$targetPath" }
    }

    return Resolve-PackagePath -PackagePath $combined
}

function Test-PathInsideDirectory {
    param(
        [string]$ChildPath,
        [string]$ParentPath
    )

    if ([string]::IsNullOrWhiteSpace($ChildPath) -or [string]::IsNullOrWhiteSpace($ParentPath)) {
        return $false
    }

    $childFull = [System.IO.Path]::GetFullPath($ChildPath).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $parentFull = [System.IO.Path]::GetFullPath($ParentPath).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    if ($childFull.Equals($parentFull, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }

    $separator = [System.IO.Path]::DirectorySeparatorChar
    return $childFull.StartsWith($parentFull + $separator, [System.StringComparison]::OrdinalIgnoreCase)
}

function Resolve-ExtractedPackageFilePath {
    param(
        [string]$ExtractionRoot,
        [string]$PackagePath
    )

    $resolvedPackagePath = Resolve-PackagePath -PackagePath $PackagePath
    if ([string]::IsNullOrWhiteSpace($resolvedPackagePath)) { return '' }

    $relativePath = $resolvedPackagePath -replace '/', [System.IO.Path]::DirectorySeparatorChar
    $candidate = Join-Path $ExtractionRoot $relativePath
    if (-not (Test-PathInsideDirectory -ChildPath $candidate -ParentPath $ExtractionRoot)) { return '' }
    return [System.IO.Path]::GetFullPath($candidate)
}

function Convert-ToSafePathSegment {
    param([string]$Name)

    $safe = [string]$Name
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$ch, '_')
    }
    $safe = $safe -replace '\.', '_'
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'presentation' }
    return $safe
}

function Get-RelativePathSafeStem {
    param(
        [string]$RootPath,
        [string]$TargetPath
    )

    if ([string]::IsNullOrWhiteSpace($TargetPath)) { return 'presentation' }

    $targetItem = Get-Item -LiteralPath $TargetPath
    if ([string]::IsNullOrWhiteSpace($RootPath)) {
        return Convert-ToSafePathSegment -Name ([System.IO.Path]::GetFileNameWithoutExtension($targetItem.Name))
    }

    $rootFull = [System.IO.Path]::GetFullPath($RootPath)
    $targetFull = [System.IO.Path]::GetFullPath($targetItem.FullName)
    $baseDir = if ($targetItem.PSIsContainer) { $targetFull } else { Split-Path -Parent $targetFull }

    if (-not (Test-PathInsideDirectory -ChildPath $baseDir -ParentPath $rootFull)) {
        return Convert-ToSafePathSegment -Name ([System.IO.Path]::GetFileNameWithoutExtension($targetItem.Name))
    }

    $relativeDir = ''
    if (-not $baseDir.Equals($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relativeDir = $baseDir.Substring($rootFull.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    }

    $parts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($relativeDir)) {
        foreach ($segment in ($relativeDir -split '[\\/]')) {
            $safeSegment = Convert-ToSafePathSegment -Name $segment
            if (-not [string]::IsNullOrWhiteSpace($safeSegment)) { $parts.Add($safeSegment) | Out-Null }
        }
    }

    $parts.Add((Convert-ToSafePathSegment -Name ([System.IO.Path]::GetFileNameWithoutExtension($targetItem.Name)))) | Out-Null
    return ($parts -join '__')
}

function Get-VideoPosterImageMap {
    param([System.IO.FileInfo]$PptxFile)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $map = @{}
    $relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($PptxFile.FullName)
        $slideEntries = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide(\d+)\.xml$' })
        foreach ($slideEntry in $slideEntries) {
            if ($slideEntry.FullName -notmatch '^ppt/slides/slide(\d+)\.xml$') { continue }
            $slideNumber = [int]$Matches[1]
            $sourcePart = "ppt/slides/slide$slideNumber.xml"
            $relEntryName = "ppt/slides/_rels/slide$slideNumber.xml.rels"

            $relText = Read-ZipEntryText -Zip $zip -EntryName $relEntryName
            if ([string]::IsNullOrWhiteSpace($relText)) { continue }
            $relXml = New-Object System.Xml.XmlDocument
            $relXml.PreserveWhitespace = $false
            $relXml.LoadXml($relText)

            $relTargets = @{}
            foreach ($rel in $relXml.GetElementsByTagName('Relationship')) {
                $targetMode = [string]$rel.GetAttribute('TargetMode')
                if ($targetMode -eq 'External') { continue }
                $id = [string]$rel.GetAttribute('Id')
                if ([string]::IsNullOrWhiteSpace($id)) { continue }
                $relTargets[$id] = Resolve-PackageTarget -SourcePart $sourcePart -Target ([string]$rel.GetAttribute('Target'))
            }

            $slideText = Read-ZipEntryText -Zip $zip -EntryName $slideEntry.FullName
            if ([string]::IsNullOrWhiteSpace($slideText)) { continue }
            $slideXml = New-Object System.Xml.XmlDocument
            $slideXml.PreserveWhitespace = $false
            $slideXml.LoadXml($slideText)

            foreach ($picNode in $slideXml.SelectNodes('//*[local-name()="pic"]')) {
                $hasVideo = ($null -ne $picNode.SelectSingleNode('.//*[local-name()="videoFile"]')) -or
                    ($null -ne $picNode.SelectSingleNode('.//*[local-name()="media"]'))
                if (-not $hasVideo) { continue }

                foreach ($blipNode in $picNode.SelectNodes('.//*[local-name()="blip"]')) {
                    $relId = [string]$blipNode.GetAttribute('embed', $relationshipNamespace)
                    if ([string]::IsNullOrWhiteSpace($relId)) { $relId = [string]$blipNode.GetAttribute('r:embed') }
                    if ([string]::IsNullOrWhiteSpace($relId) -or -not $relTargets.ContainsKey($relId)) { continue }

                    $resolved = [string]$relTargets[$relId]
                    if (-not $resolved.StartsWith('ppt/media/', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                    if (-not $map.ContainsKey($resolved)) {
                        $map[$resolved] = New-Object System.Collections.Generic.List[int]
                    }
                    if (-not $map[$resolved].Contains($slideNumber)) {
                        $map[$resolved].Add($slideNumber) | Out-Null
                    }
                }
            }
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
    return $map
}

function New-PptxPackageFromDirectory {
    param(
        [string]$SourceDir,
        [string]$DestinationPath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $DestinationPath) { Remove-Item -LiteralPath $DestinationPath -Force }

    $basePath = [System.IO.Path]::GetFullPath($SourceDir).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $archive = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Get-ChildItem -LiteralPath $SourceDir -Recurse -File | ForEach-Object {
            $fullName = [System.IO.Path]::GetFullPath($_.FullName)
            $relative = $fullName.Substring($basePath.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
            $entryName = $relative -replace '\\', '/'
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $fullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
    } finally {
        $archive.Dispose()
    }
}
