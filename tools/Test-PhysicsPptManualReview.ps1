<#
.SYNOPSIS
  Fixture test for final-delivery human visual review recording.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')
. (Join-Path $PSScriptRoot 'PhysicsPpt.ReviewArtifacts.ps1')

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('physics-ppt-manual-review-' + [Guid]::NewGuid().ToString('N'))
try {
    $deliveryDir = Join-Path $testRoot '01_交付物'
    New-Item -ItemType Directory -Path $deliveryDir -Force | Out-Null
    $input = Join-Path $testRoot 'lesson.pptx'
    $finalPptx = Join-Path $deliveryDir 'lesson.pptx'
    $finalPdf = Join-Path $deliveryDir 'lesson.pdf'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '..\examples\fixtures\minimal-physics-sample.pptx') -Destination $input
    Copy-Item -LiteralPath $input -Destination $finalPptx
    [IO.File]::WriteAllText($finalPdf, "%PDF-1.7`n1 0 obj<</Type/Page>>endobj`n%%EOF`n")
    $pdfHash = Get-FileSha256Hex -Path $finalPdf
    $pptxHash = Get-FileSha256Hex -Path $finalPptx
    $manifestPath = Join-Path $testRoot 'review-manifest.json'
    $manifest = [ordered]@{
        fullChainEnabled = $true
        deliveryBlocked = $false
        deliveryStatus = 'Pending'
        files = @([ordered]@{
            input = $input
            normalizedPptx = Join-Path $testRoot '02_阶段副本\lesson.normalized.pptx'
            pdf = $finalPdf
            pdfSha256 = $pdfHash
            validation = [ordered]@{ pdfPageCount = 1 }
        })
        brandStageArtifacts = [ordered]@{
            files = @([ordered]@{ file = $input; finalPptx = $finalPptx; finalPptxSha256 = $pptxHash })
        }
        finalVisualAuditArtifacts = @([ordered]@{ file = $input; visualConfirmationGateStatus = 'NeedsReview' })
        deliveryReviews = @('Visual confirmation requires review')
    }
    [IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($false)))
    $evidence = Join-Path $testRoot 'manual-review.json'
    & (Join-Path $PSScriptRoot 'Record-PhysicsPptManualReview.ps1') -ManifestPath $manifestPath -Decision Passed -ReviewerNote 'Fixture final PPTX and PDF are clear.' -EvidencePath $evidence | Out-Null
    $recorded = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($recorded.humanFinalReview.status -ne 'Passed' -or $recorded.deliveryStatus -ne 'Ready' -or $recorded.deliveryBlocked) {
        throw 'Human final review did not produce a Ready, unblocked delivery.'
    }
    if ($recorded.humanFinalReview.files[0].finalPptxSha256 -ne $pptxHash -or $recorded.humanFinalReview.files[0].finalPdfSha256 -ne $pdfHash) {
        throw 'Human final review did not bind both final artifact hashes.'
    }
    [IO.File]::WriteAllText($finalPdf, "%PDF-1.7`n1 0 obj<</Type/Page>>endobj`n%%EOF`nchanged")
    $rejectedStale = $false
    try {
        & (Join-Path $PSScriptRoot 'Record-PhysicsPptManualReview.ps1') -ManifestPath $manifestPath -Decision Passed -ReviewerNote 'Stale review must fail.' -EvidencePath $evidence | Out-Null
    } catch {
        $rejectedStale = $_.Exception.Message -match 'Final PDF bytes do not match'
    }
    if (-not $rejectedStale) { throw 'Human final review accepted a replaced PDF.' }
    Write-Host 'Physics PPT manual final-review fixture passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
