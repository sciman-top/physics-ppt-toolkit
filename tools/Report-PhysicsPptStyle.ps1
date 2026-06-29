<#
.SYNOPSIS
  Report style issues in PowerPoint files without modifying them.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [switch]$Recurse
)

$scriptPath = Join-Path $PSScriptRoot 'Normalize-PhysicsPpt.ps1'
& $scriptPath -InputPath $InputPath -OutputDir $OutputDir -Recurse:$Recurse -ReportOnly
