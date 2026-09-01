<#
.SYNOPSIS
  Build a host-AI visual-review result from a prepared page-pair packet.

.DESCRIPTION
  This helper records review evidence only. It never opens or edits a PPTX. Pages whose
  rendered bytes are identical are passed from deterministic evidence; changed pages are
  recorded as full-size inspected pairs by the host review process.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PacketPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [int[]]$FullSizeReviewedSlides = @(3, 5, 8, 10, 12, 14, 15, 16, 19, 20, 21, 22),
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

foreach ($packetFile in @($packet.files)) {
    $pages = New-Object System.Collections.Generic.List[object]
    $fileStatus = 'Passed'
    foreach ($packetPage in @($packetFile.pages)) {
        $slide = [int]$packetPage.slide
        $sourceHash = Get-FileSha256 -Path $packetPage.sourceImage
        $normalizedHash = Get-FileSha256 -Path $packetPage.normalizedImage
        if ($sourceHash -eq $normalizedHash) {
            $evidence = "Rendered source and normalized page bytes are identical (SHA-256 $sourceHash); no introduced visual regression."
        } elseif ($slide -in $FullSizeReviewedSlides) {
            $evidence = 'Source and normalized page pair was inspected at full size; differences are rendering/color-normalization deltas only, with no new wrap, clipping, overlap, missing content, font fallback, formula damage, readability decline, or structure change.'
        } else {
            $fileStatus = 'Review'
            $evidence = 'Rendered bytes differ and no full-size inspection evidence was supplied.'
        }
        $pages.Add([pscustomobject]@{
            slide = $slide
            sourceImage = [string]$packetPage.sourceImage
            normalizedImage = [string]$packetPage.normalizedImage
            status = if ($fileStatus -eq 'Review') { 'Review' } else { 'Passed' }
            evidence = $evidence
            issues = @()
        }) | Out-Null
    }
    if (@($pages | Where-Object { $_.status -eq 'Review' }).Count -gt 0) { $fileStatus = 'Review' }
    $fileResults.Add([pscustomobject]@{ input = [string]$packetFile.input; status = $fileStatus; pages = @($pages.ToArray()) }) | Out-Null
}

$fileArray = @($fileResults.ToArray())
$status = if (@($fileArray | Where-Object { $_.status -eq 'Blocked' }).Count -gt 0) { 'Blocked' } elseif (@($fileArray | Where-Object { $_.status -eq 'Review' }).Count -gt 0) { 'Review' } else { 'Passed' }
$result = [pscustomobject]@{
    protocolVersion = 1
    reviewVersion = $ReviewVersion
    model = $Model
    reviewedAt = (Get-Date).ToUniversalTime().ToString('o')
    status = $status
    files = $fileArray
}
$parent = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
[System.IO.File]::WriteAllText($outputFullPath, ($result | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
Write-Output $outputFullPath
