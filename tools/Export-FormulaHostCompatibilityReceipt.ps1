<#
.SYNOPSIS
  Record read-only formula host compatibility evidence for one PPTX.

.DESCRIPTION
  Opens and exports a copy through the locally installed PowerPoint COM host.
  WPS and physical projection are probed only for availability and are never
  reported as passed without an actual host/export test. This tool never edits
  the input PPTX and never grants write-back permission.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-Sha256Local {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

$InputPath = [IO.Path]::GetFullPath($InputPath)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { throw "Input PPTX not found: $InputPath" }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$targets = New-Object System.Collections.Generic.List[object]
$pp = $null
$presentation = $null
$pdfPath = Join-Path $OutputDir 'powerpoint-com-export.pdf'
try {
    try {
        $pp = New-PowerPointApplication
        $version = [string]$pp.Version
        $build = ''
        try { $build = [string]$pp.Build } catch { }
        # Keep the same COM argument types and read-only/open-window policy as
        # the proven visual-audit path.  PowerPoint's late-bound COM surface
        # rejects some PowerShell 7 integer dispatches for ExportAsFixedFormat.
        $presentation = $pp.Presentations.Open($InputPath, -1, 0, 0)
        $slideCount = [int]$presentation.Slides.Count
        $exportMethod = ''
        $exportError = ''
        try {
            $presentation.ExportAsFixedFormat($pdfPath, 2) | Out-Null
            $exportMethod = 'ExportAsFixedFormat'
        } catch {
            $exportError = $_.Exception.Message
            # This is the same documented fallback used by the visual audit.
            # It still writes only to the report directory, never to InputPath.
            $presentation.SaveAs($pdfPath, 32) | Out-Null
            $exportMethod = 'SaveAsPdfFallback'
        }
        $exported = Test-Path -LiteralPath $pdfPath -PathType Leaf
        $targets.Add([ordered]@{
            host = 'Microsoft PowerPoint COM'
            status = if ($exported) { 'Passed' } else { 'Failed' }
            version = $version
            build = $build
            slideCount = $slideCount
            opened = $true
            exported = $exported
            exportMethod = $exportMethod
            exportFallbackReason = $exportError
            artifact = if ($exported) { [ordered]@{ path = $pdfPath; sha256 = Get-Sha256Local -Path $pdfPath } } else { $null }
            reason = if ($exported) { 'PowerPoint opened the input read-only and exported a PDF copy.' } else { 'PowerPoint opened the input but the PDF export artifact was not produced.' }
        }) | Out-Null
    } catch {
        $targets.Add([ordered]@{
            host = 'Microsoft PowerPoint COM'
            status = 'Failed'
            version = ''
            build = ''
            slideCount = 0
            opened = $false
            exported = $false
            exportMethod = ''
            exportFallbackReason = ''
            artifact = $null
            reason = $_.Exception.Message
        }) | Out-Null
    }
} finally {
    if ($null -ne $presentation) { try { $presentation.Close() } catch { }; Release-ComObjectSafe -ComObject $presentation }
    if ($null -ne $pp) { try { $pp.Quit() } catch { }; Release-ComObjectSafe -ComObject $pp }
}

$wpsCommand = Get-Command wps,wpp -ErrorAction SilentlyContinue | Select-Object -First 1
$targets.Add([ordered]@{
    host = 'WPS Office'
    status = 'Unavailable'
    version = if ($null -ne $wpsCommand) { [string]$wpsCommand.Version } else { '' }
    build = ''
    slideCount = 0
    opened = $false
    exported = $false
    exportMethod = ''
    exportFallbackReason = ''
    artifact = $null
    reason = if ($null -ne $wpsCommand) { 'WPS executable was discovered, but no safe non-interactive open/export contract is configured; no pass is claimed.' } else { 'No WPS executable was discovered on PATH; compatibility was not evaluated.' }
}) | Out-Null
$targets.Add([ordered]@{
    host = 'Physical projector or extended display'
    status = 'NotEvaluated'
    version = ''
    build = ''
    slideCount = 0
    opened = $false
    exported = $false
    exportMethod = ''
    exportFallbackReason = ''
    artifact = $null
    reason = 'A local automation receipt cannot prove physical display or projector behavior.'
}) | Out-Null

$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ readOnly = $true; writeBackAllowed = $false; inputModified = $false }
    input = [ordered]@{ path = $InputPath; sha256 = Get-Sha256Local -Path $InputPath }
    targets = @($targets.ToArray())
    limitations = @('WPS is not marked passed without an actual open/export test.', 'Physical projection and extended-display behavior are outside this host receipt.')
}
$jsonPath = Join-Path $OutputDir 'formula-host-compatibility-receipt.json'
$manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
Write-Host "Formula host compatibility receipt written: $jsonPath"
