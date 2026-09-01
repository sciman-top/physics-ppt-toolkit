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
    if ($File.BaseName -match '(\d+)$') { return [int]$Matches[1] }
    return 0
}

$manifestFullPath = [System.IO.Path]::GetFullPath($ManifestPath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$manifest = Get-Content -LiteralPath $manifestFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
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
        $sourceImages = @(Get-ChildItem -LiteralPath $sourceDir -Filter '*.png' -File | Sort-Object { Get-PageNumber $_ })
        $normalizedImages = @(Get-ChildItem -LiteralPath $normalizedDir -Filter '*.png' -File | Sort-Object { Get-PageNumber $_ })
        $normalizedByPage = @{}; foreach ($image in $normalizedImages) { $normalizedByPage[[string](Get-PageNumber $image)] = $image.FullName }
        foreach ($source in $sourceImages) {
            $slide = Get-PageNumber $source
            if ($slide -le 0 -or -not $normalizedByPage.ContainsKey([string]$slide)) {
                $status = 'ReviewUnavailable'
                $reason = "Page pairing failed at source image $($source.Name)."
                continue
            }
            $pages.Add([pscustomobject]@{ slide = $slide; sourceImage = $source.FullName; normalizedImage = $normalizedByPage[[string]$slide] }) | Out-Null
        }
        if ($sourceImages.Count -ne $normalizedImages.Count) {
            $status = 'ReviewUnavailable'
            $reason = "Page-image count mismatch: source=$($sourceImages.Count), normalized=$($normalizedImages.Count)."
        }
    }
    $filePackets.Add([pscustomobject]@{ input = [string]$file.input; status = $status; reason = $reason; pages = @($pages.ToArray()) }) | Out-Null
}

$packetArray = @($filePackets.ToArray())
$packet = [pscustomobject]@{
    protocolVersion = 1
    generatedAt = (Get-Date).ToUniversalTime().ToString('o')
    sourceManifest = $manifestFullPath
    skill = 'manual/physics-ppt-visual-review/SKILL.md'
    resultSchema = 'manual/physics-ppt-visual-review/references/review-result.schema.json'
    status = if (@($packetArray | Where-Object { $_.status -ne 'Ready' }).Count -eq 0) { 'Ready' } else { 'ReviewUnavailable' }
    files = $packetArray
}
$parent = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
$json = $packet | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText($outputFullPath, $json, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
Write-Output $outputFullPath
