<#
.SYNOPSIS
  Build a host-AI visual-review result from a prepared page-pair packet.

.DESCRIPTION
  This helper records review evidence only. It never opens or edits a PPTX. Pages whose
  rendered bytes are identical are passed from deterministic evidence; changed pages are
  only marked as full-size inspected when their slide numbers are supplied explicitly via
  -FullSizeReviewedSlides by the host review process. With no explicit list, changed pages
  fall back to status Review so no unverified page can pass the gate.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PacketPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [int[]]$FullSizeReviewedSlides = @(),
    [string]$Model = 'host-visual-review',
    [string]$ReviewVersion = 'physics-ppt-visual-review-1.0'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$packetFullPath = [System.IO.Path]::GetFullPath($PacketPath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$packet = Get-Content -LiteralPath $packetFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
$fileResults = New-Object System.Collections.Generic.List[object]

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

function Get-ExpectedAggregateStatus {
    param([object[]]$Files)
    $statuses = @($Files | ForEach-Object { [string]$_.status })
    if ($statuses -contains 'Blocked') { return 'Blocked' }
    if ($statuses -contains 'Review') { return 'Review' }
    if ($statuses -contains 'ReviewUnavailable') { return 'ReviewUnavailable' }
    return 'Passed'
}

$packetStatus = [string]$packet.status
$packetFileArray = @($packet.files)
if ($packetStatus -ne 'Ready') {
    $unavailableFiles = foreach ($packetFile in $packetFileArray) {
        [pscustomobject]@{
            input = [string]$packetFile.input
            status = 'ReviewUnavailable'
            reason = [string]$packetFile.reason
            pages = @()
        }
    }
    $unavailableResult = [pscustomobject]@{
        protocolVersion = 1
        reviewVersion = $ReviewVersion
        model = $Model
        reviewedAt = (Get-Date).ToUniversalTime().ToString('o')
        packetPath = $packetFullPath
        packetSha256 = Get-FileSha256 -Path $packetFullPath
        status = 'ReviewUnavailable'
        files = @($unavailableFiles)
    }
    $parent = Split-Path -Parent $outputFullPath
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    [System.IO.File]::WriteAllText($outputFullPath, ($unavailableResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    Write-Output $outputFullPath
    return
}

$fullSizeSlides = @($FullSizeReviewedSlides | Sort-Object -Unique)
if ($packetFileArray.Count -ne 1 -and $fullSizeSlides.Count -gt 0) {
    # A bare slide number cannot identify a page in a multi-file packet. Keep
    # every changed page in Review rather than applying evidence to another deck.
    $fullSizeSlides = @()
}

foreach ($packetFile in $packetFileArray) {
    $pages = New-Object System.Collections.Generic.List[object]
    $pageStatuses = New-Object System.Collections.Generic.List[string]
    foreach ($packetPage in @($packetFile.pages)) {
        $slide = [int]$packetPage.slide
        $sourceHash = Get-FileSha256 -Path $packetPage.sourceImage
        $normalizedHash = Get-FileSha256 -Path $packetPage.normalizedImage
        $pageStatus = 'Passed'
        $issues = @()
        if ($sourceHash -ne [string]$packetPage.sourceSha256 -or $normalizedHash -ne [string]$packetPage.normalizedSha256) {
            $pageStatus = 'Blocked'
            $evidence = 'Prepared page image hash no longer matches the packet; review evidence is stale or the image was replaced.'
            $issues = @([pscustomobject]@{ type = 'Other'; severity = 'Error'; evidence = $evidence })
        } elseif ($sourceHash -eq $normalizedHash) {
            $evidence = "Rendered source and normalized page bytes are identical (SHA-256 $sourceHash); no introduced visual regression."
        } elseif ($slide -in $fullSizeSlides) {
            $evidence = 'Source and normalized page pair was inspected at full size; differences are rendering/color-normalization deltas only, with no new wrap, clipping, overlap, missing content, font fallback, formula damage, readability decline, or structure change.'
        } else {
            $pageStatus = 'Review'
            $evidence = 'Rendered bytes differ and no full-size inspection evidence was supplied.'
        }
        $pages.Add([pscustomobject]@{
            slide = $slide
            sourceImage = [string]$packetPage.sourceImage
            normalizedImage = [string]$packetPage.normalizedImage
            sourceSha256 = $sourceHash
            normalizedSha256 = $normalizedHash
            status = $pageStatus
            evidence = $evidence
            issues = $issues
        }) | Out-Null
        $pageStatuses.Add($pageStatus) | Out-Null
    }
    $fileStatus = Get-ExpectedAggregateStatus -Files (@($pageStatuses | ForEach-Object { [pscustomobject]@{ status = $_ } }))
    $fileResults.Add([pscustomobject]@{ input = [string]$packetFile.input; status = $fileStatus; pages = @($pages.ToArray()) }) | Out-Null
}

$fileArray = @($fileResults.ToArray())
$status = if (@($fileArray | Where-Object { $_.status -eq 'Blocked' }).Count -gt 0) { 'Blocked' } elseif (@($fileArray | Where-Object { $_.status -eq 'Review' }).Count -gt 0) { 'Review' } else { 'Passed' }
$result = [pscustomobject]@{
    protocolVersion = 1
    reviewVersion = $ReviewVersion
    model = $Model
    reviewedAt = (Get-Date).ToUniversalTime().ToString('o')
    packetPath = $packetFullPath
    packetSha256 = Get-FileSha256 -Path $packetFullPath
    status = $status
    files = $fileArray
}
$parent = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
[System.IO.File]::WriteAllText($outputFullPath, ($result | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
Write-Output $outputFullPath
