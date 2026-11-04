<#
.SYNOPSIS
  Internal PowerPoint slide-show worker for Test-PhysicsPptPlayback.ps1.
#>
param(
    [Parameter(Mandatory = $true)][string]$EvidenceDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$workerInputs = Get-Content -LiteralPath (Join-Path $EvidenceDir 'playback-show-input.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$InputPath = [string]$workerInputs.InputPath
$FullScreen = [bool]$workerInputs.FullScreen
$resolvedInput = [IO.Path]::GetFullPath($InputPath)
$application = $null
$presentation = $null
$window = $null
try {
    $application = New-PowerPointApplication
    $presentation = Invoke-WithComRetry -MaxRetries 3 -DelayMs 800 -Action {
        $application.Presentations.Open($resolvedInput, -1, 0, 0)
    }
    if ($presentation.Slides.Count -eq 0) { throw 'Playback presentation contains no slides.' }
    $SlideCount = [int]$presentation.Slides.Count
    $FirstSlidePng = Join-Path $EvidenceDir 'playback-slide-001.png'
    $LastSlidePng = Join-Path $EvidenceDir ('playback-slide-{0:D3}.png' -f $SlideCount)
    $presentation.Slides.Item(1).Export($FirstSlidePng, 'PNG', 1600, 900) | Out-Null
    $presentation.Slides.Item($SlideCount).Export($LastSlidePng, 'PNG', 1600, 900) | Out-Null
    if (-not (Test-UsablePageImage -Path $FirstSlidePng) -or -not (Test-UsablePageImage -Path $LastSlidePng)) { throw 'PowerPoint did not render readable first/last slide images.' }
    $settings = $presentation.SlideShowSettings
    $settings.StartingSlide = 1
    $settings.EndingSlide = $SlideCount
    $settings.ShowPresenterView = 0
    $settings.ShowScrollbar = 0
    $settings.ShowWithAnimation = -1
    $show = $settings.Run()
    if ($null -eq $show) { throw 'PowerPoint did not start slide-show playback.' }
    if ($FullScreen -and [int]$show.View.Type -ne 3) { throw "PowerPoint did not enter full-screen playback (actual view $($show.View.Type))." }
    for ($slideNumber = 1; $slideNumber -le $SlideCount; $slideNumber++) {
        $current = [int]$show.View.Slide.SlideIndex
        if ($current -ne $slideNumber) { throw "Playback stopped on slide $current; expected $slideNumber." }
        if ($slideNumber -lt $SlideCount) {
            $expectedSlide = $slideNumber + 1
            $show.View.GotoSlide($expectedSlide) | Out-Null
            for ($attempt = 0; $attempt -lt 20; $attempt++) {
                try { $current = [int]$show.View.Slide.SlideIndex } catch { $current = 0 }
                if ($current -eq $expectedSlide) { break }
                Start-Sleep -Milliseconds 250
            }
            if ($current -ne $expectedSlide) { throw "Playback did not reach slide $expectedSlide; actual $current." }
        }
    }
    $show.View.Exit()
    $window = $null
    [IO.File]::WriteAllText((Join-Path (Split-Path -Parent $FirstSlidePng) 'playback-show-receipt.json'), ([ordered]@{ slideCount = $SlideCount; visitedSlides = $SlideCount; firstSlidePng = $FirstSlidePng; lastSlidePng = $LastSlidePng } | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Playback visited $SlideCount slides."
} finally {
    if ($null -ne $window) { try { $window.View.Exit() } catch { } ; Release-ComObjectSafe $window }
    if ($null -ne $presentation) { try { $presentation.Close() } catch { } ; Release-ComObjectSafe $presentation }
    if ($null -ne $application) { try { Stop-PowerPointApplicationIfIdle -Application $application } catch { } ; Release-ComObjectSafe $application }
}
