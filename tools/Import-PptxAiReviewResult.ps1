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

function Assert-RequiredString {
    param($Object, [string]$Name, [string]$Context)
    $value = [string](Get-PropertyValue $Object $Name '')
    if ([string]::IsNullOrWhiteSpace($value)) { throw "$Context is missing required string '$Name'." }
    return $value
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

if ([int](Get-PropertyValue $result 'protocolVersion' 0) -ne 1) { throw 'AI review result protocolVersion must be 1.' }
Assert-RequiredString $result 'reviewVersion' 'AI review result' | Out-Null
Assert-RequiredString $result 'model' 'AI review result' | Out-Null
Assert-RequiredString $result 'reviewedAt' 'AI review result' | Out-Null
$resultStatus = Assert-RequiredString $result 'status' 'AI review result'
if ($resultStatus -notin $allowedStatuses) { throw "Unsupported AI review result status: $resultStatus" }

$packetFiles = @((Get-PropertyValue $packet 'files' @()))
$resultFiles = @((Get-PropertyValue $result 'files' @()))
if ($packetFiles.Count -ne $resultFiles.Count) { throw "AI review file count mismatch: expected $($packetFiles.Count), received $($resultFiles.Count)." }

$packetByInput = @{}
foreach ($packetFile in $packetFiles) { $packetByInput[[string](Get-PropertyValue $packetFile 'input' '')] = $packetFile }
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
    if ($fileStatus -eq 'ReviewUnavailable') {
        if ($actualPages.Count -ne 0) { throw "ReviewUnavailable file '$input' must not claim partial page coverage." }
        continue
    }
    if ($actualPages.Count -ne $expectedPages.Count) { throw "AI review page count mismatch for '$input': expected $($expectedPages.Count), received $($actualPages.Count)." }
    $expectedBySlide = @{}; foreach ($page in $expectedPages) { $expectedBySlide[[string](Get-PropertyValue $page 'slide' 0)] = $page }
    $pageStatuses = New-Object System.Collections.Generic.List[string]
    foreach ($page in $actualPages) {
        $slide = [int](Get-PropertyValue $page 'slide' 0)
        if ($slide -lt 1 -or -not $expectedBySlide.ContainsKey([string]$slide)) { throw "AI review page has an unexpected slide number for '$input': $slide" }
        $expectedPage = $expectedBySlide[[string]$slide]
        foreach ($pathProperty in @('sourceImage', 'normalizedImage')) {
            if ([string](Get-PropertyValue $page $pathProperty '') -ne [string](Get-PropertyValue $expectedPage $pathProperty '')) {
                throw "AI review page $slide in '$input' does not match its prepared $pathProperty."
            }
        }
        $pageStatus = Assert-RequiredString $page 'status' "AI review page $slide in '$input'"
        if ($pageStatus -notin $allowedPageStatuses) { throw "Unsupported AI review page status: $pageStatus" }
        Assert-RequiredString $page 'evidence' "AI review page $slide in '$input'" | Out-Null
        $pageStatuses.Add($pageStatus) | Out-Null
        foreach ($issue in @((Get-PropertyValue $page 'issues' @()))) {
            if ([string](Get-PropertyValue $issue 'type' '') -notin $allowedIssueTypes) { throw "AI review page $slide in '$input' has an unsupported issue type." }
            if ([string](Get-PropertyValue $issue 'severity' '') -notin $allowedSeverities) { throw "AI review page $slide in '$input' has an unsupported severity." }
            Assert-RequiredString $issue 'evidence' "AI review issue on slide $slide in '$input'" | Out-Null
        }
    }
    $expectedFileStatus = if ($pageStatuses -contains 'Blocked') { 'Blocked' } elseif ($pageStatuses -contains 'Review') { 'Review' } else { 'Passed' }
    if ($fileStatus -ne $expectedFileStatus) { throw "AI review file '$input' aggregate status must be $expectedFileStatus, received $fileStatus." }
}

$expectedResultStatus = Get-ExpectedAggregateStatus -Files $resultFiles
if ($resultStatus -ne $expectedResultStatus) { throw "AI review aggregate status must be $expectedResultStatus, received $resultStatus." }

$gate = [pscustomobject]@{
    status = $resultStatus
    deliveryBlocked = ($resultStatus -eq 'Blocked')
    reviewRequired = ($resultStatus -in @('Review', 'ReviewUnavailable'))
    result = $resultFullPath
    packet = $packetFullPath
    model = [string]$result.model
    reviewVersion = [string]$result.reviewVersion
    reviewedAt = [string]$result.reviewedAt
}

if (-not $ValidateOnly) {
    $manifest | Add-Member -NotePropertyName aiVisualReview -NotePropertyValue $gate -Force
    $manifest | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath $manifestFullPath -Encoding UTF8
    $summaryPath = Join-Path (Split-Path -Parent $manifestFullPath) 'summary.md'
    if (Test-Path -LiteralPath $summaryPath) {
        $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8
        $summary = [regex]::Replace($summary, '(?m)^- 宿主 AI 视觉门禁：.*$', "- 宿主 AI 视觉门禁：$resultStatus")
        [System.IO.File]::WriteAllText($summaryPath, $summary, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    }
}

$gate | ConvertTo-Json -Depth 6
