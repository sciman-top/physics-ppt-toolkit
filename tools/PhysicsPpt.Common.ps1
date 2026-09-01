<#
.SYNOPSIS
  Shared low-level helpers for PPTX package scripts.

.DESCRIPTION
  Keep this file limited to deterministic helpers with no top-level side
  effects. Production scripts dot-source it to share encoding, file discovery,
  COM cleanup, ZIP, and Open XML relationship handling.
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

function Write-Utf8BomCsv {
    <#
      Export-Csv's UTF8 behavior differs between Windows PowerShell and
      PowerShell 7.  Writing through an explicit BOM keeps Chinese reports
      stable for Excel and downstream tooling on both hosts.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [Alias('Rows')]
        [AllowNull()]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $parent = Split-Path -Parent ([System.IO.Path]::GetFullPath($Path))
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $csvLines = @($InputObject | ConvertTo-Csv -NoTypeInformation)
    [System.IO.File]::WriteAllLines([System.IO.Path]::GetFullPath($Path), $csvLines, $utf8Bom)
}

function Write-Utf8BomText {
    param(
        [AllowEmptyString()]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $parent = Split-Path -Parent ([System.IO.Path]::GetFullPath($Path))
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($Path), $Text, $utf8Bom)
}

function Get-NormalizedFormulaText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return (($Text -replace '\s+', '') -replace '＝', '=').Trim()
}

function Get-FormulaDetailValue {
    param([string]$Details, [string]$Key)
    if ([string]::IsNullOrWhiteSpace($Details)) { return '' }
    $pattern = '(?:^|;\s*)' + [regex]::Escape($Key) + '=(.*?)(?=;\s*\w+=|$)'
    $match = [regex]::Match($Details, $pattern)
    if (-not $match.Success) { return '' }
    return $match.Groups[1].Value.Trim()
}

function Get-FormulaRuleValue {
    param($Rule, [string]$Name, [string]$Default = '')
    if ($null -eq $Rule) { return $Default }
    $property = $Rule.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return [string]$property.Value
}

function Release-ComObjectSafe {
    param($ComObject)
    if ($null -eq $ComObject) { return }
    try {
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        }
    } catch {
        # Cleanup must not hide the processing error that led to this boundary.
    }
}

function Get-PowerShellHostInfo {
    <#
      Resolve the host used for child PowerShell processes.  PowerShell 7 is
      the primary runtime; Windows PowerShell 5.1 is retained only as an
      explicit compatibility fallback for older managed desktops.
    #>
    $pwshCommand = Get-Command -Name 'pwsh.exe' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -eq $pwshCommand) {
        $pwshCommand = Get-Command -Name 'pwsh' -CommandType Application -ErrorAction SilentlyContinue |
            Select-Object -First 1
    }
    if ($null -ne $pwshCommand) {
        $path = [string]$pwshCommand.Source
        if ($pwshCommand.PSObject.Properties['Path'] -and -not [string]::IsNullOrWhiteSpace([string]$pwshCommand.Path)) {
            $path = [string]$pwshCommand.Path
        }
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            return [pscustomobject]@{
                Path = $path
                Name = 'PowerShell 7'
                IsPrimary = $true
            }
        }
    }

    $legacyCommand = Get-Command -Name 'powershell.exe' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -ne $legacyCommand) {
        $path = [string]$legacyCommand.Source
        if ($legacyCommand.PSObject.Properties['Path'] -and -not [string]::IsNullOrWhiteSpace([string]$legacyCommand.Path)) {
            $path = [string]$legacyCommand.Path
        }
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            return [pscustomobject]@{
                Path = $path
                Name = 'Windows PowerShell 5.1 (fallback)'
                IsPrimary = $false
            }
        }
    }

    return $null
}

function Resolve-PowerShellHost {
    <#
      Return an executable path for a child PowerShell process.  The fallback
      is deliberately kept here, at one boundary, so all callers use the same
      PS7-first policy and no worker silently regresses to powershell.exe.
    #>
    param([switch]$RequirePowerShell7)

    $hostInfo = Get-PowerShellHostInfo
    if ($null -eq $hostInfo) {
        throw "No PowerShell host was found. Install PowerShell 7 (pwsh) or enable the Windows PowerShell 5.1 compatibility host."
    }
    if ($RequirePowerShell7 -and -not $hostInfo.IsPrimary) {
        throw "PowerShell 7 (pwsh) is required for this operation, but only the Windows PowerShell 5.1 fallback was found."
    }
    return [string]$hostInfo.Path
}

function New-PowerPointApplication {
    <#
      Start PowerPoint for automation without showing an interactive window.
      Callers can still explicitly open the generated result after the workflow
      completes when a human review is wanted.  DisplayAlerts=1 is the
      PowerPoint ppAlertsNone value; the guarded assignment keeps compatibility
      with installations that expose a reduced COM surface.
    #>
    param([switch]$Visible)

    $application = New-Object -ComObject PowerPoint.Application
    try {
        $application.Visible = $(if ($Visible) { -1 } else { 0 })
    } catch {
        # A COM host may not expose Visible until its first call; keep going and
        # let the caller report any later automation failure.
    }
    if (-not $Visible) {
        try { $application.DisplayAlerts = 1 } catch { }
    }
    return $application
}

function Convert-ToSafeFormulaPathSegment {
    param([string]$Name)
    $safe = [string]$Name
    foreach ($character in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    $safe = $safe -replace '\s+', '_'
    $safe = $safe -replace '[^\p{L}\p{Nd}_-]+', '_'
    $safe = $safe.Trim('_')
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'formula' }
    return $safe
}

function Convert-ToSafeFileNameSegment {
    param([string]$Name)
    $safe = [string]$Name
    foreach ($character in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'presentation' }
    return $safe
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

function Get-PresentationFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$Pattern = '*.ppt*',
        [switch]$Recurse,
        [string[]]$SupportedExtensions = @('.pptx'),
        [string[]]$ExcludedRoots = @()
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "InputPath not found: $Path" }
    $item = Get-Item -LiteralPath $Path
    $extensions = @($SupportedExtensions | ForEach-Object { ([string]$_).ToLowerInvariant() })

    if (-not $item.PSIsContainer) {
        if ($item.Extension.ToLowerInvariant() -notin $extensions) {
            throw "Unsupported presentation extension: $($item.FullName)"
        }
        return @($item)
    }

    $rootFull = [System.IO.Path]::GetFullPath($item.FullName)
    $excludedFull = @(
        $ExcludedRoots |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            ForEach-Object { [System.IO.Path]::GetFullPath([string]$_) }
    )
    $options = @{ LiteralPath = $rootFull; Filter = $Pattern; File = $true }
    if ($Recurse) { $options.Recurse = $true }

    return @(
        Get-ChildItem @options |
            Where-Object {
                $candidate = $_
                if ($candidate.Name -like '~$*' -or $candidate.Extension.ToLowerInvariant() -notin $extensions) {
                    return $false
                }
                foreach ($excludedRoot in $excludedFull) {
                    if (Test-PathInsideDirectory -ChildPath $candidate.FullName -ParentPath $excludedRoot) {
                        return $false
                    }
                }
                $relative = $candidate.FullName.Substring($rootFull.Length).TrimStart(
                    [System.IO.Path]::DirectorySeparatorChar,
                    [System.IO.Path]::AltDirectorySeparatorChar
                )
                foreach ($segment in ($relative -split '[\\/]')) {
                    if ($segment -match '^_physics_ppt_output_\d{8}_\d{6}$') { return $false }
                }
                return $true
            } |
            Sort-Object FullName
    )
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

function Expand-PptxPackageSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PptxPath,
        [Parameter(Mandatory = $true)]
        [string]$DestinationDir
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $sourceFull = [System.IO.Path]::GetFullPath($PptxPath)
    $destinationFull = [System.IO.Path]::GetFullPath($DestinationDir)
    if (-not (Test-Path -LiteralPath $sourceFull -PathType Leaf)) { throw "PPTX package not found: $sourceFull" }
    if (Test-Path -LiteralPath $destinationFull) {
        if (@(Get-ChildItem -LiteralPath $destinationFull -Force).Count -gt 0) {
            throw "Extraction directory must be empty: $destinationFull"
        }
    } else {
        New-Item -ItemType Directory -Path $destinationFull -Force | Out-Null
    }

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($sourceFull)
        foreach ($entry in $zip.Entries) {
            $target = Resolve-ExtractedPackageFilePath -ExtractionRoot $destinationFull -PackagePath $entry.FullName
            if ([string]::IsNullOrWhiteSpace($target)) {
                throw "Unsafe package entry path: $($entry.FullName)"
            }
            if ([string]::IsNullOrWhiteSpace($entry.Name)) {
                if (-not (Test-Path -LiteralPath $target)) { New-Item -ItemType Directory -Path $target -Force | Out-Null }
                continue
            }

            $parent = Split-Path -Parent $target
            if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
            $inputStream = $null
            $outputStream = $null
            try {
                $inputStream = $entry.Open()
                $outputStream = New-Object System.IO.FileStream($target, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                $inputStream.CopyTo($outputStream)
            } finally {
                if ($null -ne $outputStream) { $outputStream.Dispose() }
                if ($null -ne $inputStream) { $inputStream.Dispose() }
            }
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
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
    $basePath = [System.IO.Path]::GetFullPath($SourceDir).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $destinationFull = [System.IO.Path]::GetFullPath($DestinationPath)
    $destinationDir = Split-Path -Parent $destinationFull
    if (-not (Test-Path -LiteralPath $destinationDir)) { New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null }
    $temporaryPath = Join-Path $destinationDir ('.' + [System.IO.Path]::GetFileName($destinationFull) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::Open($temporaryPath, [System.IO.Compression.ZipArchiveMode]::Create)
        Get-ChildItem -LiteralPath $SourceDir -Recurse -File | ForEach-Object {
            $fullName = [System.IO.Path]::GetFullPath($_.FullName)
            $relative = $fullName.Substring($basePath.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
            $entryName = $relative -replace '\\', '/'
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $fullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
        $archive.Dispose()
        $archive = $null
        if (Test-Path -LiteralPath $destinationFull) {
            [System.IO.File]::Replace($temporaryPath, $destinationFull, $null)
        } else {
            [System.IO.File]::Move($temporaryPath, $destinationFull)
        }
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
    }
}
