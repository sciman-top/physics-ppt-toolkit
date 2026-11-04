<#
.SYNOPSIS
  Controlled PowerPoint playback smoke test for a generated presentation.

.DESCRIPTION
  Opens the requested PPTX read-only, advances through every slide in
  slide-show mode, captures the first and last slide, and reports the exact
  deck hash. Use only the full-screen switch on a host where taking over the
  visible desktop is acceptable.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [Parameter(Mandatory = $true)]
    [string]$EvidenceDir,
    [switch]$FullScreen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$resolvedInput = [IO.Path]::GetFullPath($InputPath)
$resolvedEvidence = [IO.Path]::GetFullPath($EvidenceDir)
if (-not (Test-Path -LiteralPath $resolvedInput -PathType Leaf)) { throw "Playback input not found: $resolvedInput" }
if (-not (Test-PptxPackageIntegrity -Path $resolvedInput)) { throw "Playback input is not a readable PPTX: $resolvedInput" }
New-Item -ItemType Directory -Path $resolvedEvidence -Force | Out-Null

$result = [ordered]@{
    status = 'Failed'
    inputPath = $resolvedInput
    inputSha256 = Get-FileSha256Hex -Path $resolvedInput
    slideCount = 0
    visitedSlides = 0
    firstSlidePng = $null
    lastSlidePng = $null
    fullScreen = [bool]$FullScreen
    error = ''
}
$slideCount = 0

try {
    # Run the show on its dedicated COM thread. PowerShell's COM apartment
    # rejects a Run() return-value assignment when there is no active host
    # window, even though PowerPoint has started the show.
    $showHost = Resolve-PowerShellHost
    $workerPath = Join-Path $resolvedEvidence 'playback-show.ps1'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Test-PhysicsPptPlaybackShow.ps1') -Destination $workerPath
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1') -Destination (Join-Path $resolvedEvidence 'PhysicsPpt.Common.ps1')
    $showStdout = Join-Path $resolvedEvidence 'playback-show.stdout.log'
    $showStderr = Join-Path $resolvedEvidence 'playback-show.stderr.log'
    $workerInputs = [ordered]@{
        InputPath = $resolvedInput
        FullScreen = [bool]$FullScreen
    }
    [IO.File]::WriteAllText((Join-Path $resolvedEvidence 'playback-show-input.json'), ($workerInputs | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    $showParameters = @{
        FilePath = $showHost
        ArgumentList = @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', $workerPath,
            '-EvidenceDir', $resolvedEvidence
        )
        NoNewWindow = $true
        Wait = $true
        PassThru = $true
    }
    $showParameters.RedirectStandardOutput = $showStdout
    $showParameters.RedirectStandardError = $showStderr
    $showProcess = Start-Process @showParameters
    if ($showProcess.ExitCode -ne 0) {
        $showOutput = if (Test-Path $showStdout) { Get-Content -LiteralPath $showStdout -Raw -Encoding UTF8 } else { '' }
        $showErrorOutput = if (Test-Path $showStderr) { Get-Content -LiteralPath $showStderr -Raw -Encoding UTF8 } else { '' }
        throw "PowerPoint slide-show worker failed with exit code $($showProcess.ExitCode): $showOutput $showErrorOutput"
    }
    $showReceiptPath = Join-Path $resolvedEvidence 'playback-show-receipt.json'
    if (-not (Test-Path -LiteralPath $showReceiptPath -PathType Leaf)) { throw 'Playback worker produced no receipt.' }
    $showReceipt = Get-Content -LiteralPath $showReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $slideCount = [int]$showReceipt.slideCount
    $result.slideCount = $slideCount
    $result.visitedSlides = [int]$showReceipt.visitedSlides
    $result.firstSlidePng = [string]$showReceipt.firstSlidePng
    $result.lastSlidePng = [string]$showReceipt.lastSlidePng
    if ($result.visitedSlides -ne $slideCount) { throw "Visited $($result.visitedSlides) of $slideCount slides." }
    $result.showReceipt = $showReceiptPath
    $result.workerFullScreen = [bool]$showReceipt.fullScreen
    if (-not (Test-UsablePageImage -Path $result.firstSlidePng) -or -not (Test-UsablePageImage -Path $result.lastSlidePng)) {
        throw 'Playback capture did not produce readable first/last slide images.'
    }
    $result.status = 'Passed'
} catch {
    $result.error = $_.Exception.Message
} finally {
    $result.evidenceDir = $resolvedEvidence
    [IO.File]::WriteAllText((Join-Path $resolvedEvidence 'playback-receipt.json'), ($result | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
}

if ($result.status -ne 'Passed') { throw "PowerPoint playback smoke test failed: $($result.error)" }
Write-Host "Playback smoke test passed: $slideCount slides; evidence=$resolvedEvidence"
