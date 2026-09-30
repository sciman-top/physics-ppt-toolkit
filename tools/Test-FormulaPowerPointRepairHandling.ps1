<#
.SYNOPSIS
  Probe PowerPoint COM behavior for a malformed temporary PPTX copy.

.DESCRIPTION
  This is a bounded F5-05 evidence probe. It corrupts only a temporary copy,
  opens that copy read-only through the local PowerPoint COM host, and records
  whether COM rejects it or opens it without exposing a repair signal. Hidden
  COM automation cannot prove what a human-facing "Open and Repair" dialog did,
  so a successful open is deliberately recorded as ManualRequired rather than
  Passed. The source fixture is hash-checked before and after the probe.
#>
[CmdletBinding()]
param(
    [string]$FixturePptx = (Join-Path (Split-Path -Parent $PSScriptRoot) 'examples\fixtures\minimal-physics-sample.pptx'),
    [string]$OutputJsonPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$FixturePptx = [IO.Path]::GetFullPath($FixturePptx)
if (-not (Test-Path -LiteralPath $FixturePptx -PathType Leaf)) { throw "Fixture PPTX not found: $FixturePptx" }
$sourceHash = (Get-FileHash -LiteralPath $FixturePptx -Algorithm SHA256).Hash.ToLowerInvariant()
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('formula-powerpoint-repair-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
$pp = $null
$presentation = $null
$comStatus = 'NotRun'
$comReason = ''
$slideCount = 0

try {
    $corruptPptx = Join-Path $testRoot 'malformed-copy.pptx'
    Copy-Item -LiteralPath $FixturePptx -Destination $corruptPptx -Force

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($corruptPptx, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $archive.GetEntry('ppt/slides/slide1.xml')
        if ($null -eq $entry) { throw 'Fault fixture lacks ppt/slides/slide1.xml.' }
        $entry.Delete()
        $replacement = $archive.CreateEntry('ppt/slides/slide1.xml')
        $writer = New-Object IO.StreamWriter($replacement.Open(), (New-Object Text.UTF8Encoding($false)))
        try { $writer.Write('<p:sld') } finally { $writer.Dispose() }
    } finally {
        $archive.Dispose()
    }

    try {
        $pp = New-PowerPointApplication
        # ReadOnly=true, Untitled=false, WithWindow=false. Do not use a visible
        # repair dialog as a test oracle; this probe is intentionally headless.
        $presentation = $pp.Presentations.Open($corruptPptx, -1, 0, 0)
        $slideCount = [int]$presentation.Slides.Count
        $comStatus = 'Opened'
        $comReason = 'PowerPoint opened the malformed copy without exposing a COM error; silent repair cannot be distinguished from tolerant parsing in hidden automation.'
    } catch {
        $comStatus = 'Failed'
        $comReason = $_.Exception.Message
    }

    $sourceHashAfter = (Get-FileHash -LiteralPath $FixturePptx -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($sourceHashAfter -ne $sourceHash) { throw 'PowerPoint repair probe modified the source fixture.' }

    $decision = if ($comStatus -eq 'Failed') { 'OriginalKept' } else { 'ManualRequired' }
    $receipt = [ordered]@{
        schemaVersion = 1
        fixture = [ordered]@{ path = $FixturePptx; sha256 = $sourceHash }
        temporaryCopy = [ordered]@{ path = $corruptPptx; malformed = $true }
        powerPointCom = [ordered]@{
            status = $comStatus
            slideCount = $slideCount
            reason = $comReason
        }
        repairSignal = 'NotObservable'
        decision = $decision
        writeBackAllowed = $false
        sourceModified = $false
        limitation = 'Hidden PowerPoint COM automation cannot prove a human-facing Open and Repair prompt or silent repair outcome.'
    }
    $json = $receipt | ConvertTo-Json -Depth 20
    if (-not [string]::IsNullOrWhiteSpace($OutputJsonPath)) {
        $OutputJsonPath = [IO.Path]::GetFullPath($OutputJsonPath)
        $parent = Split-Path -Parent $OutputJsonPath
        if (-not [string]::IsNullOrWhiteSpace($parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        [IO.File]::WriteAllText($OutputJsonPath, $json, (New-Object Text.UTF8Encoding($false)))
    }
    $json
} finally {
    if ($null -ne $presentation) { try { $presentation.Close() | Out-Null } catch { }; Release-ComObjectSafe -ComObject $presentation }
    if ($null -ne $pp) { try { $pp.Quit() | Out-Null } catch { }; Release-ComObjectSafe -ComObject $pp }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
