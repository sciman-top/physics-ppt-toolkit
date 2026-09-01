<##
.SYNOPSIS
  Compare two PPTX invariant snapshots and allow only style-property changes.
##>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$BeforePath,
    [Parameter(Mandatory = $true)][string]$AfterPath,
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-Snapshot { param([string]$Path) return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
function Add-Difference { param($Rows, [string]$Path, [string]$Before, [string]$After, [string]$Kind)
    $Rows.Add([pscustomobject]@{ path = $Path; before = $Before; after = $After; kind = $Kind }) | Out-Null
}
function Compare-Value { param($Rows, [string]$Path, $Before, $After, [string]$Kind = 'Invariant')
    if ("$Before" -ne "$After") { Add-Difference $Rows $Path "$Before" "$After" $Kind }
}

$before = Read-Snapshot $BeforePath
$after = Read-Snapshot $AfterPath
$rows = New-Object System.Collections.Generic.List[object]
Compare-Value $rows 'slideCount' $before.slideCount $after.slideCount 'Blocker'
Compare-Value $rows 'slideWidth' $before.slideWidth $after.slideWidth 'Blocker'
Compare-Value $rows 'slideHeight' $before.slideHeight $after.slideHeight 'Blocker'

$beforeSlides = @($before.slides); $afterSlides = @($after.slides)
$slideCount = [Math]::Min($beforeSlides.Count, $afterSlides.Count)
for ($i = 0; $i -lt $slideCount; $i++) {
    $bs = $beforeSlides[$i]; $as = $afterSlides[$i]; $prefix = "slides[$i]"
    Compare-Value $rows "$prefix.index" $bs.index $as.index 'Blocker'
    Compare-Value $rows "$prefix.slideId" $bs.slideId $as.slideId 'Blocker'
    Compare-Value $rows "$prefix.hidden" $bs.hidden $as.hidden 'Blocker'
    if ("$($bs.transition.advanceOnClick)" -ne "$($as.transition.advanceOnClick)") {
        $advanceKind = if ([int]$bs.transition.advanceOnClick -eq -1 -and [int]$as.transition.advanceOnClick -eq 0) { 'AllowedChange' } else { 'Blocker' }
        Add-Difference $rows "$prefix.transition.advanceOnClick" "$($bs.transition.advanceOnClick)" "$($as.transition.advanceOnClick)" $advanceKind
    }
    foreach ($transitionProperty in @('advanceOnTime','advanceTime','entryEffect','speed')) {
        Compare-Value $rows "$prefix.transition.$transitionProperty" $bs.transition.$transitionProperty $as.transition.$transitionProperty 'Blocker'
    }
    Compare-Value $rows "$prefix.animations" ($bs.animations | ConvertTo-Json -Compress) ($as.animations | ConvertTo-Json -Compress) 'Blocker'

    $beforeShapes = @($bs.shapes); $afterShapes = @($as.shapes)
    $shapeMapBefore = @{}; foreach ($s in $beforeShapes) { $shapeMapBefore[[string]$s.id] = $s }
    $shapeMapAfter = @{}; foreach ($s in $afterShapes) { $shapeMapAfter[[string]$s.id] = $s }
    foreach ($id in @($shapeMapBefore.Keys + $shapeMapAfter.Keys | Sort-Object -Unique)) {
        if (-not $shapeMapBefore.ContainsKey($id)) { Add-Difference $rows "$prefix.shapes[$id]" '' 'present' 'Blocker'; continue }
        if (-not $shapeMapAfter.ContainsKey($id)) { Add-Difference $rows "$prefix.shapes[$id]" 'present' '' 'Blocker'; continue }
        $b = $shapeMapBefore[$id]; $a = $shapeMapAfter[$id]
        foreach ($property in @('name','type','left','top','width','height','rotation','zOrder','autoSize','wordWrap','text')) {
            Compare-Value $rows "$prefix.shapes[$id].$property" $b.$property $a.$property 'Blocker'
        }
    }
}

# Package relationships and embedded media are part of the teaching artifact.
# Normalization is never allowed to rewrite them in this safety-first workflow.
Compare-Value $rows 'package.relationships' ($before.package.relationships | ConvertTo-Json -Compress -Depth 8) ($after.package.relationships | ConvertTo-Json -Compress -Depth 8) 'Blocker'
Compare-Value $rows 'package.media' ($before.package.media | ConvertTo-Json -Compress -Depth 8) ($after.package.media | ConvertTo-Json -Compress -Depth 8) 'Blocker'

$differenceArray = @($rows.ToArray())
$blockerCount = @($differenceArray | Where-Object { $_.kind -eq 'Blocker' }).Count
$result = [pscustomobject]@{
    schemaVersion = 1
    before = [System.IO.Path]::GetFullPath($BeforePath)
    after = [System.IO.Path]::GetFullPath($AfterPath)
    passed = ($blockerCount -eq 0)
    blockerCount = $blockerCount
    allowedChangeCount = @($differenceArray | Where-Object { $_.kind -eq 'AllowedChange' }).Count
    differences = $differenceArray
}
$json = $result | ConvertTo-Json -Depth 20
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $parent = Split-Path -Parent ([System.IO.Path]::GetFullPath($OutputPath))
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($OutputPath), $json, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
}
Write-Output $json
