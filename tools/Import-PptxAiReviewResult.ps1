<#
.SYNOPSIS
  Validate and attach a read-only host-AI visual-review result to a workflow manifest.

.DESCRIPTION
  This script never opens, edits, copies, or saves a presentation. It only validates a
  JSON result against the prepared page-pair packet and records the resulting delivery gate
  in review-manifest.json. Blocked results remain visible and make the workflow gate fail.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ManifestPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [string]$PacketPath,
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PropertyValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return $property.Value
}

function Test-PathIdentityEqual {
    param([string]$Left, [string]$Right)

    if ([string]::IsNullOrWhiteSpace($Left) -or [string]::IsNullOrWhiteSpace($Right)) { return $false }
    $comparison = if ([System.IO.Path]::DirectorySeparatorChar -eq '\') {
        [System.StringComparison]::OrdinalIgnoreCase
    } else {
        [System.StringComparison]::Ordinal
    }
    return [string]::Equals($Left, $Right, $comparison)
}

function Assert-RequiredString {
    param($Object, [string]$Name, [string]$Context)
    $value = [string](Get-PropertyValue $Object $Name '')
    if ([string]::IsNullOrWhiteSpace($value)) { throw "$Context is missing required string '$Name'." }
    return $value
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

function Get-ExpectedAggregateStatus {
    param([object[]]$Files)
    $statuses = @($Files | ForEach-Object { [string](Get-PropertyValue $_ 'status' '') })
    if ($statuses -contains 'Blocked') { return 'Blocked' }
    if ($statuses -contains 'Review') { return 'Review' }
    if ($statuses -contains 'ReviewUnavailable') { return 'ReviewUnavailable' }
    return 'Passed'
}

$allowedStatuses = @('Passed', 'Review', 'Blocked', 'ReviewUnavailable')
$allowedPageStatuses = @('Passed', 'Review', 'Blocked')
$allowedIssueTypes = @('NewWrap', 'Clipping', 'Overlap', 'MissingContent', 'NewBlank', 'FontFallback', 'FormulaDamage', 'Readability', 'StructureChange', 'KnownSourceIssue', 'Other')
$allowedSeverities = @('Info', 'Warning', 'Error')

$manifestFullPath = [System.IO.Path]::GetFullPath($ManifestPath)
$resultFullPath = [System.IO.Path]::GetFullPath($ResultPath)
if ([string]::IsNullOrWhiteSpace($PacketPath)) { $PacketPath = Join-Path (Split-Path -Parent $manifestFullPath) 'ai-visual-review-request.json' }
$packetFullPath = [System.IO.Path]::GetFullPath($PacketPath)
foreach ($path in @($manifestFullPath, $resultFullPath, $packetFullPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required file is missing: $path" }
}

$manifest = Get-Content -LiteralPath $manifestFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
$packet = Get-Content -LiteralPath $packetFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
$result = Get-Content -LiteralPath $resultFullPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ([int](Get-PropertyValue $packet 'protocolVersion' 0) -ne 1) { throw 'AI review packet protocolVersion must be 1.' }
if (-not (Test-PathIdentityEqual -Left ([string](Get-PropertyValue $packet 'sourceManifest' '')) -Right $manifestFullPath)) { throw 'AI review packet sourceManifest does not match the current manifest.' }
$packetManifestSha256 = [string](Get-PropertyValue $packet 'manifestSha256' '')
$currentManifestSha256 = Get-FileSha256 -Path $manifestFullPath
$existingReview = Get-PropertyValue $manifest 'aiVisualReview' $null
$preparedManifestSha256 = [string](Get-PropertyValue $existingReview 'preparedManifestSha256' '')
if ([string]::IsNullOrWhiteSpace($packetManifestSha256) -or
    ($packetManifestSha256 -ne $currentManifestSha256 -and $packetManifestSha256 -ne $preparedManifestSha256)) {
    throw 'AI review packet manifest hash does not match the current or previously prepared manifest.'
}
$packetStatus = Assert-RequiredString $packet 'status' 'AI review packet'
if ($packetStatus -notin @('Ready', 'ReviewUnavailable')) { throw "Unsupported AI review packet status: $packetStatus" }

if ([int](Get-PropertyValue $result 'protocolVersion' 0) -ne 1) { throw 'AI review result protocolVersion must be 1.' }
Assert-RequiredString $result 'reviewVersion' 'AI review result' | Out-Null
Assert-RequiredString $result 'model' 'AI review result' | Out-Null
Assert-RequiredString $result 'reviewedAt' 'AI review result' | Out-Null
$resultStatus = Assert-RequiredString $result 'status' 'AI review result'
if ($resultStatus -notin $allowedStatuses) { throw "Unsupported AI review result status: $resultStatus" }
$resultPacketPath = Assert-RequiredString $result 'packetPath' 'AI review result'
if (-not (Test-PathIdentityEqual -Left $resultPacketPath -Right $packetFullPath)) { throw 'AI review result packetPath does not match the prepared packet.' }
$resultPacketSha256 = Assert-RequiredString $result 'packetSha256' 'AI review result'
if ($resultPacketSha256 -ne (Get-FileSha256 -Path $packetFullPath)) { throw 'AI review result packet hash does not match the prepared packet.' }
if ($packetStatus -eq 'Ready' -and $resultStatus -eq 'ReviewUnavailable') { throw 'AI review result cannot claim ReviewUnavailable for a Ready packet.' }
if ($packetStatus -eq 'ReviewUnavailable' -and $resultStatus -ne 'ReviewUnavailable') { throw 'AI review result must remain ReviewUnavailable when the prepared packet is unavailable.' }

$packetFiles = @((Get-PropertyValue $packet 'files' @()))
$resultFiles = @((Get-PropertyValue $result 'files' @()))
if ($packetStatus -eq 'Ready' -and $packetFiles.Count -eq 0) { throw 'A Ready AI review packet must contain at least one prepared file.' }
if ($packetFiles.Count -ne $resultFiles.Count) { throw "AI review file count mismatch: expected $($packetFiles.Count), received $($resultFiles.Count)." }

$packetByInput = @{}
foreach ($packetFile in $packetFiles) {
    $packetInput = Assert-RequiredString $packetFile 'input' 'AI review packet file'
    if ($packetByInput.ContainsKey($packetInput)) { throw "AI review packet contains duplicate input: $packetInput" }
    $packetByInput[$packetInput] = $packetFile
}
$seenInputs = @{}
foreach ($file in $resultFiles) {
    $input = Assert-RequiredString $file 'input' 'AI review file'
    if (-not $packetByInput.ContainsKey($input)) { throw "AI review refers to unprepared input: $input" }
    if ($seenInputs.ContainsKey($input)) { throw "AI review contains duplicate input: $input" }
    $seenInputs[$input] = $true
    $fileStatus = Assert-RequiredString $file 'status' "AI review file '$input'"
    if ($fileStatus -notin $allowedStatuses) { throw "Unsupported AI review file status: $fileStatus" }
    $expectedPages = @((Get-PropertyValue $packetByInput[$input] 'pages' @()))
    $actualPages = @((Get-PropertyValue $file 'pages' @()))
    $packetFileStatus = [string](Get-PropertyValue $packetByInput[$input] 'status' '')
    if ($fileStatus -eq 'ReviewUnavailable') {
        if ($actualPages.Count -ne 0) { throw "ReviewUnavailable file '$input' must not claim partial page coverage." }
        # Cross-check against the packet: unavailability is decided when the
        # packet is prepared, not by the reviewing model afterwards. A packet
        # marked Ready still carries full page evidence and must be reviewed.
        if ($packetFileStatus -ne 'ReviewUnavailable') {
            throw "AI review file '$input' claims ReviewUnavailable but the prepared packet marks it '$packetFileStatus'; only a ReviewUnavailable packet may be reported back as unavailable."
        }
        continue
    }
    if ($packetFileStatus -ne 'Ready') { throw "AI review file '$input' cannot claim $fileStatus for a packet file marked '$packetFileStatus'." }
    if ($actualPages.Count -ne $expectedPages.Count) { throw "AI review page count mismatch for '$input': expected $($expectedPages.Count), received $($actualPages.Count)." }
    $expectedBySlide = @{}
    foreach ($page in $expectedPages) {
        $expectedSlide = [int](Get-PropertyValue $page 'slide' 0)
        if ($expectedSlide -lt 1 -or $expectedBySlide.ContainsKey([string]$expectedSlide)) { throw "AI review packet has duplicate or invalid slide number for '$input': $expectedSlide" }
        $expectedBySlide[[string]$expectedSlide] = $page
    }
    $pageStatuses = New-Object System.Collections.Generic.List[string]
    $seenSlides = @{}
    foreach ($page in $actualPages) {
        $slide = [int](Get-PropertyValue $page 'slide' 0)
        if ($slide -lt 1 -or -not $expectedBySlide.ContainsKey([string]$slide)) { throw "AI review page has an unexpected slide number for '$input': $slide" }
        if ($seenSlides.ContainsKey([string]$slide)) { throw "AI review contains duplicate slide $slide for '$input'." }
        $seenSlides[[string]$slide] = $true
        $expectedPage = $expectedBySlide[[string]$slide]
        foreach ($pathProperty in @('sourceImage', 'normalizedImage')) {
            if ([string](Get-PropertyValue $page $pathProperty '') -ne [string](Get-PropertyValue $expectedPage $pathProperty '')) {
                throw "AI review page $slide in '$input' does not match its prepared $pathProperty."
            }
        }
        foreach ($hashProperty in @('sourceSha256', 'normalizedSha256')) {
            $expectedHash = Assert-RequiredString $expectedPage $hashProperty "AI review packet page $slide in '$input'"
            $actualHash = Assert-RequiredString $page $hashProperty "AI review page $slide in '$input'"
            if ($actualHash -ne $expectedHash) { throw "AI review page $slide in '$input' has a mismatched $hashProperty." }
            $pathProperty = if ($hashProperty -eq 'sourceSha256') { 'sourceImage' } else { 'normalizedImage' }
            $actualPath = [string](Get-PropertyValue $page $pathProperty '')
            if (-not (Test-Path -LiteralPath $actualPath) -or (Get-FileSha256 -Path $actualPath) -ne $expectedHash) { throw "AI review page $slide in '$input' has a stale or replaced image for $hashProperty." }
        }
        $pageStatus = Assert-RequiredString $page 'status' "AI review page $slide in '$input'"
        if ($pageStatus -notin $allowedPageStatuses) { throw "Unsupported AI review page status: $pageStatus" }
        Assert-RequiredString $page 'evidence' "AI review page $slide in '$input'" | Out-Null
        $pageStatuses.Add($pageStatus) | Out-Null
        foreach ($issue in @((Get-PropertyValue $page 'issues' @()))) {
            if ([string](Get-PropertyValue $issue 'type' '') -notin $allowedIssueTypes) { throw "AI review page $slide in '$input' has an unsupported issue type." }
            if ([string](Get-PropertyValue $issue 'severity' '') -notin $allowedSeverities) { throw "AI review page $slide in '$input' has an unsupported severity." }
            Assert-RequiredString $issue 'evidence' "AI review issue on slide $slide in '$input'" | Out-Null
            if ([string](Get-PropertyValue $issue 'severity' '') -eq 'Error' -and $pageStatus -ne 'Blocked') { throw "AI review page $slide in '$input' must be Blocked when it contains an Error issue." }
            if ([string](Get-PropertyValue $issue 'severity' '') -eq 'Warning' -and $pageStatus -eq 'Passed') { throw "AI review page $slide in '$input' must be Review or Blocked when it contains a Warning issue." }
        }
    }
    if ($seenSlides.Count -ne $expectedBySlide.Count) { throw "AI review page set is incomplete for '$input'." }
    $expectedFileStatus = if ($pageStatuses -contains 'Blocked') { 'Blocked' } elseif ($pageStatuses -contains 'Review') { 'Review' } else { 'Passed' }
    if ($fileStatus -ne $expectedFileStatus) { throw "AI review file '$input' aggregate status must be $expectedFileStatus, received $fileStatus." }
}

$expectedResultStatus = Get-ExpectedAggregateStatus -Files $resultFiles
if ($resultStatus -ne $expectedResultStatus) { throw "AI review aggregate status must be $expectedResultStatus, received $resultStatus." }

$gate = [pscustomobject]@{
    status = $resultStatus
    deliveryBlocked = ($resultStatus -eq 'Blocked')
    reviewRequired = ($resultStatus -in @('Review', 'ReviewUnavailable'))
    preparedManifestSha256 = $packetManifestSha256
    result = $resultFullPath
    packet = $packetFullPath
    model = [string]$result.model
    reviewVersion = [string]$result.reviewVersion
    reviewedAt = [string]$result.reviewedAt
}

if (-not $ValidateOnly) {
    $manifest | Add-Member -NotePropertyName aiVisualReview -NotePropertyValue $gate -Force
    $invariantGate = Get-PropertyValue $manifest 'invariantGate' $null
    $invariantBlocked = ([string](Get-PropertyValue $invariantGate 'status' '') -eq 'Blocked')
    $deliveryBlocked = [bool]($invariantBlocked -or $gate.deliveryBlocked)
    $deliveryStatus = if ($invariantBlocked) { 'BlockedInvariantGate' } elseif ($gate.deliveryBlocked) { 'BlockedAiVisualReview' } elseif ($resultStatus -eq 'Passed') { 'Ready' } else { 'Pending' }
    $manifest | Add-Member -NotePropertyName deliveryBlocked -NotePropertyValue $deliveryBlocked -Force
    $manifest | Add-Member -NotePropertyName deliveryStatus -NotePropertyValue $deliveryStatus -Force
    $manifest | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath $manifestFullPath -Encoding UTF8
    $summaryPath = Join-Path (Split-Path -Parent $manifestFullPath) 'summary.md'
    if (Test-Path -LiteralPath $summaryPath) {
        $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8
        $summary = [regex]::Replace($summary, '(?m)^- 宿主 AI 视觉门禁：.*$', "- 宿主 AI 视觉门禁：$resultStatus")
        $summary = [regex]::Replace($summary, '(?m)^- 交付状态：.*$', "- 交付状态：$deliveryStatus")
        [System.IO.File]::WriteAllText($summaryPath, $summary, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    }
}

$gate | ConvertTo-Json -Depth 6
