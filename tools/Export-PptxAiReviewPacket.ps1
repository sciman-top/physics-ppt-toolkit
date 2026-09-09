<##
.SYNOPSIS
  Build a read-only page-pair packet for host-AI visual review.
##>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ManifestPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PageNumber {
    param([System.IO.FileInfo]$File)
    if ($File.BaseName -match '^(?:page|slide)[-_]?(\d+)$') { return [int]$Matches[1] }
    return 0
}

function Get-FileSha256 {
    param([string]$Path)
    $stream = $null
    $sha256 = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        return ($sha256.ComputeHash($stream) | ForEach-Object { $_.ToString('x2') }) -join ''
    } finally {
        if ($null -ne $sha256) { $sha256.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Test-UsablePageImage {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $image = $null
    try {
        if ((Get-Item -LiteralPath $Path).Length -le 0) { return $false }
        Add-Type -AssemblyName System.Drawing
        $image = [System.Drawing.Image]::FromFile($Path)
        return ($image.Width -gt 0 -and $image.Height -gt 0)
    } catch {
        return $false
    } finally {
        if ($null -ne $image) { $image.Dispose() }
    }
}

function Get-PageImageInventory {
    param([string]$Directory)
    $files = @(Get-ChildItem -LiteralPath $Directory -Filter '*.png' -File | Sort-Object Name)
    $byPage = @{}
    $invalid = New-Object System.Collections.Generic.List[string]
    $duplicates = New-Object System.Collections.Generic.List[int]
    foreach ($file in $files) {
        $page = Get-PageNumber $file
        if ($page -le 0) {
            $invalid.Add($file.Name) | Out-Null
            continue
        }
        if ($byPage.ContainsKey([string]$page)) {
            $duplicates.Add($page) | Out-Null
            continue
        }
        $byPage[[string]$page] = $file
    }
    [pscustomobject]@{
        Files = $files
        ByPage = $byPage
        Pages = @($byPage.Keys | ForEach-Object { [int]$_ } | Sort-Object)
        Invalid = @($invalid.ToArray())
        DuplicatePages = @($duplicates.ToArray() | Sort-Object -Unique)
    }
}

function Get-ManifestExpectedSlideCount {
    param($ManifestFile)
    $validationProperty = $ManifestFile.PSObject.Properties['validation']
    if ($null -eq $validationProperty) { return $null }
    $expectedProperty = $validationProperty.Value.PSObject.Properties['expectedSlides']
    if ($null -eq $expectedProperty) { return $null }
    $expected = 0
    if (-not [int]::TryParse([string]$expectedProperty.Value, [ref]$expected) -or $expected -le 0) { return -1 }
    return $expected
}

$manifestFullPath = [System.IO.Path]::GetFullPath($ManifestPath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$manifest = Get-Content -LiteralPath $manifestFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
$manifestSha256 = Get-FileSha256 -Path $manifestFullPath
$filePackets = New-Object System.Collections.Generic.List[object]

foreach ($file in @($manifest.files)) {
    $sourceDir = [string]$file.sourcePageImages
    $normalizedDir = [string]$file.pageImages
    $pages = New-Object System.Collections.Generic.List[object]
    $status = 'Ready'
    $reason = ''
    if ([string]::IsNullOrWhiteSpace($sourceDir) -or [string]::IsNullOrWhiteSpace($normalizedDir) -or -not (Test-Path -LiteralPath $sourceDir) -or -not (Test-Path -LiteralPath $normalizedDir)) {
        $status = 'ReviewUnavailable'
        $reason = 'Original or normalized page-image directory is missing. Re-run with -IncludeReviewArtifacts.'
    } else {
        $sourceInventory = Get-PageImageInventory -Directory $sourceDir
        $normalizedInventory = Get-PageImageInventory -Directory $normalizedDir
        $expectedSlides = Get-ManifestExpectedSlideCount -ManifestFile $file
        if ($null -eq $expectedSlides) { $expectedSlides = $sourceInventory.Pages.Count }
        $expectedKey = if ($expectedSlides -gt 0) { ((1..$expectedSlides) -join ',') } else { '' }
        $sourceKey = (($sourceInventory.Pages | ForEach-Object { [string]$_ }) -join ',')
        $normalizedKey = (($normalizedInventory.Pages | ForEach-Object { [string]$_ }) -join ',')
        if ($expectedSlides -le 0 -or $sourceInventory.Invalid.Count -gt 0 -or $normalizedInventory.Invalid.Count -gt 0 -or
            $sourceInventory.DuplicatePages.Count -gt 0 -or $normalizedInventory.DuplicatePages.Count -gt 0 -or
            $sourceKey -ne $expectedKey -or $normalizedKey -ne $expectedKey) {
            $status = 'ReviewUnavailable'
            $reason = "Page-image set is incomplete or ambiguous: expected=$expectedKey; source=$sourceKey; normalized=$normalizedKey."
        } elseif (@($sourceInventory.Pages | Where-Object { -not (Test-UsablePageImage -Path $sourceInventory.ByPage[[string]$_].FullName) }).Count -gt 0 -or
            @($normalizedInventory.Pages | Where-Object { -not (Test-UsablePageImage -Path $normalizedInventory.ByPage[[string]$_].FullName) }).Count -gt 0) {
            $status = 'ReviewUnavailable'
            $reason = 'One or more page images are empty or cannot be decoded.'
        } else {
            foreach ($slide in $sourceInventory.Pages) {
                $source = $sourceInventory.ByPage[[string]$slide]
                $normalized = $normalizedInventory.ByPage[[string]$slide]
                $pages.Add([pscustomobject]@{
                    slide = $slide
                    sourceImage = $source.FullName
                    normalizedImage = $normalized.FullName
                    sourceSha256 = Get-FileSha256 -Path $source.FullName
                    normalizedSha256 = Get-FileSha256 -Path $normalized.FullName
                }) | Out-Null
            }
        }
    }
    $filePackets.Add([pscustomobject]@{
        input = [string]$file.input
        status = $status
        reason = $reason
        expectedSlideCount = if ($pages.Count -gt 0) { $pages.Count } else { 0 }
        pages = @($pages.ToArray())
    }) | Out-Null
}

$packetArray = @($filePackets.ToArray())
$packet = [pscustomobject]@{
    protocolVersion = 1
    generatedAt = (Get-Date).ToUniversalTime().ToString('o')
    sourceManifest = $manifestFullPath
    skill = 'manual/physics-ppt-visual-review/SKILL.md'
    resultSchema = 'manual/physics-ppt-visual-review/references/review-result.schema.json'
    manifestSha256 = $manifestSha256
    status = if ($packetArray.Count -gt 0 -and @($packetArray | Where-Object { $_.status -ne 'Ready' }).Count -eq 0) { 'Ready' } else { 'ReviewUnavailable' }
    files = $packetArray
}
$parent = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
$json = $packet | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText($outputFullPath, $json, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
Write-Output $outputFullPath
