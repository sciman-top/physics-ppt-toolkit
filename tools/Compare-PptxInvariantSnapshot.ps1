<##
.SYNOPSIS
  Compare two PPTX invariant snapshots and allow only style-property changes.
##>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$BeforePath,
    [Parameter(Mandatory = $true)][string]$AfterPath,
    [string]$OutputPath,
    [switch]$AllowAdvanceOnClickDisable,
    [switch]$AllowReviewedFormulaConversion
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# PowerPoint exposes geometry as Single.  A no-geometry-change SaveAs can
# round-trip a value by a few ten-thousandths of a point, which is below any
# meaningful layout movement but used to become a false invariant blocker.
# Keep this far below one screen pixel and below the smallest allowed manual
# adjustment, so real position, size, and rotation changes still block.
$geometryTolerance = 0.01

function Read-Snapshot { param([string]$Path) return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
function Add-Difference { param($Rows, [string]$Path, [string]$Before, [string]$After, [string]$Kind)
    $Rows.Add([pscustomobject]@{ path = $Path; before = $Before; after = $After; kind = $Kind }) | Out-Null
}
function Compare-Value { param($Rows, [string]$Path, $Before, $After, [string]$Kind = 'Invariant')
    if ("$Before" -ne "$After") { Add-Difference $Rows $Path "$Before" "$After" $Kind }
}
function Compare-GeometryValue { param($Rows, [string]$Path, $Before, $After, [string]$Kind = 'Invariant')
    if ($null -eq $Before -or $null -eq $After) {
        Compare-Value $Rows $Path $Before $After $Kind
        return
    }
    try {
        if ([Math]::Abs(([double]$Before) - ([double]$After)) -le $geometryTolerance) { return }
    } catch {
        Compare-Value $Rows $Path $Before $After $Kind
        return
    }
    Add-Difference $Rows $Path "$Before" "$After" $Kind
}
function Get-SnapshotProperty { param($Row, [string]$Name)
    # Snapshots written by older exporters (and synthetic test fixtures) may not
    # carry newer fields; treat them as absent rather than throwing.
    $property = $Row.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Test-ReviewedFormulaConversion {
    param($BeforeShape, $AfterShape)
    if (-not $AllowReviewedFormulaConversion) { return $false }
    # This is deliberately narrow: only an Equation.* OLE (type 7) replaced
    # by the generated FormulaOmml text shape (type 1) may use the opt-in
    # allowance. Geometry, z-order, group membership, animation and package
    # media are still compared as blockers below.
    return ([int]$BeforeShape.type -eq 7 -and
        [int]$AfterShape.type -eq 1 -and
        [string]$AfterShape.name -like 'FormulaOmml*' -and
        [string]::IsNullOrWhiteSpace([string]$BeforeShape.text) -and
        -not [string]::IsNullOrWhiteSpace([string]$AfterShape.text))
}

$before = Read-Snapshot $BeforePath
$after = Read-Snapshot $AfterPath
$rows = New-Object System.Collections.Generic.List[object]
Compare-Value $rows 'slideCount' $before.slideCount $after.slideCount 'Blocker'
# Slide dimensions come from the same Single COM pipeline as shape geometry,
# so they get the same rounding tolerance (a string compare produced false
# blockers for custom page sizes).
Compare-GeometryValue $rows 'slideWidth' $before.slideWidth $after.slideWidth 'Blocker'
Compare-GeometryValue $rows 'slideHeight' $before.slideHeight $after.slideHeight 'Blocker'

$beforeSlides = @($before.slides); $afterSlides = @($after.slides)
$slideCount = [Math]::Min($beforeSlides.Count, $afterSlides.Count)
for ($i = 0; $i -lt $slideCount; $i++) {
    $bs = $beforeSlides[$i]; $as = $afterSlides[$i]; $prefix = "slides[$i]"
    Compare-Value $rows "$prefix.index" $bs.index $as.index 'Blocker'
    Compare-Value $rows "$prefix.slideId" $bs.slideId $as.slideId 'Blocker'
    Compare-Value $rows "$prefix.hidden" $bs.hidden $as.hidden 'Blocker'
    foreach ($readStatusProperty in @('slideReadStatus', 'animationReadStatus', 'transitionReadStatus')) {
        $beforeReadStatus = [string](Get-SnapshotProperty $bs $readStatusProperty)
        $afterReadStatus = [string](Get-SnapshotProperty $as $readStatusProperty)
        if ($beforeReadStatus -ne $afterReadStatus) {
            Add-Difference $rows "$prefix.$readStatusProperty" $beforeReadStatus $afterReadStatus 'Blocker'
        } elseif ($beforeReadStatus -ne 'Readable') {
            # Stable unreadable in both snapshots: source-inherited, not a change.
            Add-Difference $rows "$prefix.$readStatusProperty" $beforeReadStatus $afterReadStatus 'AllowedChange'
        }
    }
    $beforeUnreadableShapeCount = Get-SnapshotProperty $bs 'unreadableShapeCount'
    $afterUnreadableShapeCount = Get-SnapshotProperty $as 'unreadableShapeCount'
    if ($null -eq $beforeUnreadableShapeCount -or $null -eq $afterUnreadableShapeCount -or
        [int]$beforeUnreadableShapeCount -ne [int]$afterUnreadableShapeCount) {
        # A count CHANGE means normalization touched shapes it could not read.
        Add-Difference $rows "$prefix.unreadableShapeCount" $beforeUnreadableShapeCount $afterUnreadableShapeCount 'Blocker'
    } elseif ([int]$afterUnreadableShapeCount -ne 0) {
        # A source shape that reads as Unreadable both before and after was
        # born unreadable in the source deck (15.2 slide5); it is not a change
        # normalization made. Recorded as allowed, never silently dropped.
        Add-Difference $rows "$prefix.unreadableShapeCount" $beforeUnreadableShapeCount $afterUnreadableShapeCount 'AllowedChange'
    }
    $beforeAdvanceOnClick = Get-SnapshotProperty $bs.transition 'advanceOnClick'
    $afterAdvanceOnClick = Get-SnapshotProperty $as.transition 'advanceOnClick'
    if ("$beforeAdvanceOnClick" -ne "$afterAdvanceOnClick") {
        $advanceKind = if ($AllowAdvanceOnClickDisable -and [int]$beforeAdvanceOnClick -eq -1 -and [int]$afterAdvanceOnClick -eq 0) { 'AllowedChange' } else { 'Blocker' }
        Add-Difference $rows "$prefix.transition.advanceOnClick" "$beforeAdvanceOnClick" "$afterAdvanceOnClick" $advanceKind
    }
    foreach ($transitionProperty in @('advanceOnTime','advanceTime','entryEffect','speed')) {
        Compare-Value $rows "$prefix.transition.$transitionProperty" (Get-SnapshotProperty $bs.transition $transitionProperty) (Get-SnapshotProperty $as.transition $transitionProperty) 'Blocker'
    }
    Compare-Value $rows "$prefix.animations" ($bs.animations | ConvertTo-Json -Compress) ($as.animations | ConvertTo-Json -Compress) 'Blocker'

    $beforeShapes = @($bs.shapes); $afterShapes = @($as.shapes)
    $shapeMapBefore = @{}; foreach ($s in $beforeShapes) { $shapeMapBefore[[string]$s.id] = $s }
    $shapeMapAfter = @{}; foreach ($s in $afterShapes) { $shapeMapAfter[[string]$s.id] = $s }
    foreach ($id in @($shapeMapBefore.Keys + $shapeMapAfter.Keys | Sort-Object -Unique)) {
        if (-not $shapeMapBefore.ContainsKey($id)) { Add-Difference $rows "$prefix.shapes[$id]" '' 'present' 'Blocker'; continue }
        if (-not $shapeMapAfter.ContainsKey($id)) { Add-Difference $rows "$prefix.shapes[$id]" 'present' '' 'Blocker'; continue }
        $b = $shapeMapBefore[$id]; $a = $shapeMapAfter[$id]
        $isReviewedFormulaConversion = Test-ReviewedFormulaConversion -BeforeShape $b -AfterShape $a
        foreach ($property in @('name','type','zOrder','autoSize','wordWrap','text')) {
            if ($isReviewedFormulaConversion -and $property -in @('name', 'type', 'autoSize', 'wordWrap', 'text')) {
                if ("$($b.$property)" -ne "$($a.$property)") {
                    Add-Difference $rows "$prefix.shapes[$id].$property" "$($b.$property)" "$($a.$property)" 'AllowedChange'
                }
            } else {
                Compare-Value $rows "$prefix.shapes[$id].$property" $b.$property $a.$property 'Blocker'
            }
        }
        $beforeShapeReadStatus = [string](Get-SnapshotProperty $b 'readStatus')
        $afterShapeReadStatus = [string](Get-SnapshotProperty $a 'readStatus')
        if ($beforeShapeReadStatus -ne $afterShapeReadStatus) {
            # Readability flipped: normalization touched a shape it could not
            # read before (or vice versa) — a real regression signal.
            Add-Difference $rows "$prefix.shapes[$id].readStatus" $beforeShapeReadStatus $afterShapeReadStatus 'Blocker'
        } elseif ($beforeShapeReadStatus -ne 'Readable') {
            # Stable in both snapshots: the shape was unreadable in the source
            # deck itself. Not a normalization change; keep it visible.
            Add-Difference $rows "$prefix.shapes[$id].readStatus" $beforeShapeReadStatus $afterShapeReadStatus 'AllowedChange'
        }
        foreach ($property in @('left','top','width','height','rotation')) {
            Compare-GeometryValue $rows "$prefix.shapes[$id].$property" $b.$property $a.$property 'Blocker'
        }
        foreach ($property in @('cropLeft','cropRight','cropTop','cropBottom')) {
            Compare-GeometryValue $rows "$prefix.shapes[$id].$property" (Get-SnapshotProperty $b $property) (Get-SnapshotProperty $a $property) 'Blocker'
        }
        # Author emphasis (run colors + bold) is protected semantics: a copy
        # whose red runs flipped to black with unchanged text must fail the
        # gate (18.2 emphasis-color incident). Formula conversions legitimately
        # restyle their own block, so the reviewed-conversion allowance applies.
        $beforeFormats = Get-SnapshotProperty $b 'runFormats'
        $afterFormats = Get-SnapshotProperty $a 'runFormats'
        if ($null -ne $beforeFormats -or $null -ne $afterFormats) {
            if ("$beforeFormats" -ne "$afterFormats") {
                if ($isReviewedFormulaConversion) {
                    Add-Difference $rows "$prefix.shapes[$id].runFormats" "$beforeFormats" "$afterFormats" 'AllowedChange'
                } else {
                    Add-Difference $rows "$prefix.shapes[$id].runFormats" "$beforeFormats" "$afterFormats" 'Blocker'
                }
            }
        }
        # Group-child rows carry the parent id; if it changes, the group was restructured.
        Compare-Value $rows "$prefix.shapes[$id].groupId" (Get-SnapshotProperty $b 'groupId') (Get-SnapshotProperty $a 'groupId') 'Blocker'
    }
}

# Package relationships and embedded media are part of the teaching artifact.
# Normalization is never allowed to rewrite them in this safety-first workflow.
# PowerPoint SaveAs renumbers media part names and relationship ids even when
# no byte of media changes (same save-churn class as the geometry rounding
# false blockers), so compare canonical forms: media as an
# extension+sha256+length multiset, relationships with media targets resolved
# to the referenced content hash (external targets stay literal). A real
# rewrite, deletion, or retarget still changes the canonical form and blocks.
function Get-MediaExtension { param([string]$Part)
    $extension = ''
    if ($Part -match '\.([^.]+)$') { $extension = $Matches[1].ToLowerInvariant() }
    return $extension
}
function Get-CanonicalMediaText { param($MediaEntries)
    (@($MediaEntries | ForEach-Object {
        '{0}|{1}|{2}' -f (Get-MediaExtension $_.part), $_.sha256, $_.length
    } | Sort-Object) -join "`n")
}
function Get-CanonicalRelationshipText { param($RelationshipEntries, $MediaEntries)
    $mediaHashByLeaf = @{}
    foreach ($media in @($MediaEntries)) {
        $leaf = [System.IO.Path]::GetFileName([string]$media.part)
        $mediaHashByLeaf[$leaf] = [string]$media.sha256
    }
    (@($RelationshipEntries | ForEach-Object {
        $target = [string]$_.target
        $targetLeaf = [System.IO.Path]::GetFileName($target)
        if (-not [string]::IsNullOrEmpty($targetLeaf) -and $mediaHashByLeaf.ContainsKey($targetLeaf)) {
            $target = 'media:' + $mediaHashByLeaf[$targetLeaf]
        }
        '{0}|{1}|{2}|{3}' -f $_.part, $_.type, $target, $_.targetMode
    } | Sort-Object) -join "`n")
}
Compare-Value $rows 'package.relationships' (Get-CanonicalRelationshipText $before.package.relationships $before.package.media) (Get-CanonicalRelationshipText $after.package.relationships $after.package.media) 'Blocker'
Compare-Value $rows 'package.media' (Get-CanonicalMediaText $before.package.media) (Get-CanonicalMediaText $after.package.media) 'Blocker'

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
