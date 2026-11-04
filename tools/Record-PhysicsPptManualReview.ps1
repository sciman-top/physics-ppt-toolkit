<#
.SYNOPSIS
  Record a user's visual acceptance of the final PPTX and PDF without editing them.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,
    [Parameter(Mandatory = $true)]
    [ValidateSet('Passed', 'Rejected')]
    [string]$Decision,
    [Parameter(Mandatory = $true)]
    [string]$ReviewerNote,
    [Parameter(Mandatory = $true)]
    [string]$EvidencePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')
. (Join-Path $PSScriptRoot 'PhysicsPpt.ReviewArtifacts.ps1')

function Get-MetadataValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

$manifestPath = [IO.Path]::GetFullPath($ManifestPath)
$evidencePath = [IO.Path]::GetFullPath($EvidencePath)
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Manifest not found: $manifestPath" }
if ([string]::IsNullOrWhiteSpace($ReviewerNote)) { throw 'ReviewerNote must describe the observed final PPTX/PDF result.' }
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not [bool](Get-MetadataValue $manifest 'fullChainEnabled' $false)) {
    throw 'Final PPTX/PDF manual review requires a FullChain manifest.'
}

$files = @((Get-MetadataValue $manifest 'files' @()))
if ($files.Count -eq 0) { throw 'Manifest contains no final delivery files.' }
    $reviews = New-Object System.Collections.Generic.List[object]
    foreach ($file in $files) {
    $input = [string](Get-MetadataValue $file 'input' '')
    $stem = [IO.Path]::GetFileNameWithoutExtension($input)
    $finalVisualItems = @((Get-MetadataValue $manifest 'finalVisualAuditArtifacts' @()) | Where-Object {
        [string](Get-MetadataValue $_ 'file' '') -eq $input
    })
    $pdfPath = [string](Get-MetadataValue $file 'pdf' '')
    $deliveryDir = if ([string]::IsNullOrWhiteSpace($pdfPath)) { '' } else { Split-Path -Parent $pdfPath }
    if ([string]::IsNullOrWhiteSpace($input) -or [string]::IsNullOrWhiteSpace($deliveryDir)) {
        throw 'Manifest does not identify the original input and final delivery directory.'
    }
    $pptxFullPath = [IO.Path]::GetFullPath((Join-Path $deliveryDir ([IO.Path]::GetFileName($input))))
    $pdfFullPath = [IO.Path]::GetFullPath((Join-Path $deliveryDir ($stem + '.pdf')))
    $pptx = Get-Item -LiteralPath $pptxFullPath -ErrorAction Stop
    $pdf = Get-Item -LiteralPath $pdfFullPath -ErrorAction Stop
    $markerNames = @($pptx.Name, $pdf.Name)
    $deliveryFiles = @(Get-ChildItem -LiteralPath $deliveryDir -File)
    if (@($deliveryFiles | Where-Object { $_.Name -notin $markerNames }).Count -gt 0) {
        throw 'Final delivery directory contains files other than the final same-name PPTX and PDF.'
    }
    if (-not (Test-PptxPackageIntegrity -Path $pptx.FullName)) { throw "Final PPTX is not readable: $($pptx.FullName)" }
    if (-not (Test-PdfPackageIntegrity -Path $pdf.FullName)) { throw "Final PDF is not readable: $($pdf.FullName)" }
    $validation = Get-MetadataValue $file 'validation' $null
    $expectedPdfPages = [int](Get-MetadataValue $validation 'pdfPageCount' 0)
    $actualPdfPages = Get-PdfPageCount -PdfPath $pdf.FullName
    if ($null -eq $actualPdfPages -or $actualPdfPages -ne $expectedPdfPages) {
        throw "Final PDF page count does not match the manifest: $actualPdfPages vs $expectedPdfPages"
    }
    $finalPptxSha256 = Get-FileSha256Hex -Path $pptx.FullName
    $finalPdfSha256 = Get-FileSha256Hex -Path $pdf.FullName
    $brandStage = Get-MetadataValue $manifest 'brandStageArtifacts' $null
    $brandFiles = @((Get-MetadataValue $brandStage 'files' @()))
    $chainRecord = @($brandFiles)
    if ($chainRecord.Count -ne 1) { throw "Final PPTX receipt is missing or ambiguous ($($chainRecord.Count) of $($brandFiles.Count))." }
    if ([string](Get-MetadataValue $chainRecord[0] 'file' '') -ine $input) {
        throw 'The full-chain receipt belongs to a different source presentation.'
    }
    $receiptPptxSha256 = [string](Get-MetadataValue $chainRecord[0] 'finalPptxSha256' '')
    if ($finalPptxSha256 -ine $receiptPptxSha256) { throw 'Final PPTX bytes do not match the full-chain receipt.' }
    if ($finalPdfSha256 -ine [string](Get-MetadataValue $file 'pdfSha256' '')) {
        throw 'Final PDF bytes do not match the manifest receipt.'
    }
    $autoReviewRequired = @((Get-MetadataValue $manifest 'deliveryReviews' @())).Count -gt 0
    $reviews.Add([pscustomobject]@{
        input = $input
        finalPptx = $pptx.FullName
        finalPptxSha256 = $finalPptxSha256
        finalPdf = $pdf.FullName
        finalPdfSha256 = $finalPdfSha256
        finalPdfPageCount = [int]$actualPdfPages
        automaticReviewRequired = $autoReviewRequired
        finalVisualGate = if ($finalVisualItems.Count -eq 1) { [string](Get-MetadataValue $finalVisualItems[0] 'visualConfirmationGateStatus' '') } else { '' }
        decision = $Decision
        reviewerNote = $ReviewerNote
    }) | Out-Null
}

$recording = [ordered]@{
    schemaVersion = 1
    recordedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    source = 'UserConfirmedBothFormats'
    kind = 'HumanFinalDeliveryReview'
    status = $Decision
    reviewerNote = $ReviewerNote
    files = @($reviews.ToArray())
}
$parent = Split-Path -Parent $evidencePath
if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
[IO.File]::WriteAllText($evidencePath, ($recording | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($false)))
$manifest | Add-Member -NotePropertyName humanFinalReview -NotePropertyValue ([pscustomobject]@{ status = $Decision; evidence = $evidencePath; files = @($reviews.ToArray()) }) -Force
$blocked = @((Get-MetadataValue $manifest 'deliveryFailures' @())).Count -gt 0 -or [bool](Get-MetadataValue $manifest 'deliveryBlocked' $false)
if ($Decision -eq 'Rejected') {
    $manifest | Add-Member -NotePropertyName deliveryBlocked -NotePropertyValue $true -Force
    $manifest | Add-Member -NotePropertyName deliveryStatus -NotePropertyValue 'BlockedManualVisualReview' -Force
} elseif (-not $blocked) {
    $manifest | Add-Member -NotePropertyName deliveryBlocked -NotePropertyValue $false -Force
    $manifest | Add-Member -NotePropertyName deliveryStatus -NotePropertyValue 'Ready' -Force
}
[IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 16), (New-Object System.Text.UTF8Encoding($false)))
$summary = Join-Path (Split-Path -Parent $manifestPath) 'summary.md'
if (Test-Path -LiteralPath $summary -PathType Leaf) {
    $text = Get-Content -LiteralPath $summary -Raw -Encoding UTF8
    $text = [regex]::Replace($text, '(?m)^- 交付状态：.*$', "- 交付状态：$($manifest.deliveryStatus)")
    [IO.File]::WriteAllText($summary, $text, (New-Object System.Text.UTF8Encoding($false)))
}
Write-Output $evidencePath
