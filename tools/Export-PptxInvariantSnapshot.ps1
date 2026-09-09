<##
.SYNOPSIS
  Export a read-only PowerPoint snapshot for invariant comparison.

.DESCRIPTION
  Captures slide order, text, object geometry/layering, animation and transition
  summaries. It does not save or modify the presentation.
##>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-PackageSnapshot {
    param([string]$PptxPath)

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
        $relationships = New-Object System.Collections.Generic.List[object]
        $media = New-Object System.Collections.Generic.List[object]
        foreach ($entry in @($archive.Entries)) {
            if ($entry.FullName -like '*.rels') {
                try {
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    try {
                        [xml]$xml = $reader.ReadToEnd()
                        foreach ($relationship in @($xml.Relationships.Relationship)) {
                            $type = [string]$relationship.Type
                            $targetMode = [string]$relationship.GetAttribute('TargetMode')
                            # PowerPoint legitimately regenerates package-internal metadata links when saving.
                            # The invariant gate therefore tracks only externally reachable targets and media links.
                            if ($targetMode -eq 'External' -or $type -match '/(image|audio|video|media)$') {
                                $relationships.Add([pscustomobject]@{
                                    part = $entry.FullName
                                    id = [string]$relationship.Id
                                    type = $type
                                    target = [string]$relationship.Target
                                    targetMode = $targetMode
                                }) | Out-Null
                            }
                        }
                    } finally { $reader.Dispose() }
                } catch { throw "Unable to read relationship part $($entry.FullName): $($_.Exception.Message)" }
            }
            if ($entry.FullName -like 'ppt/media/*') {
                $stream = $null
                try {
                    $stream = $entry.Open()
                    $sha256 = [System.Security.Cryptography.SHA256]::Create()
                    try {
                        $hash = ($sha256.ComputeHash($stream) | ForEach-Object { $_.ToString('x2') }) -join ''
                    } finally { $sha256.Dispose() }
                    $media.Add([pscustomobject]@{ part = $entry.FullName; length = [long]$entry.Length; sha256 = $hash }) | Out-Null
                } finally { if ($null -ne $stream) { $stream.Dispose() } }
            }
        }
        return [pscustomobject]@{
            relationships = @($relationships.ToArray() | Sort-Object part, id, type, target, targetMode)
            media = @($media.ToArray() | Sort-Object part)
        }
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
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

function Get-SnapshotText {
    param($Shape, [ref]$ReadFailed)
    try {
        if ($Shape.TextFrame2.HasText -eq -1) { return [string]$Shape.TextFrame2.TextRange.Text }
    } catch { $ReadFailed.Value = $true }
    return ''
}

function Get-SafeComProperty {
    param($Object, [string]$Name, $Default = $null, [ref]$ReadFailed)
    try {
        if ($null -eq $Object) { throw "COM object is null while reading '$Name'." }
        return $Object.$Name
    } catch {
        if ($PSBoundParameters.ContainsKey('ReadFailed')) { $ReadFailed.Value = $true }
        return $Default
    }
}

function Get-ShapeSnapshot {
    param($Shape)
    $readFailed = $false
    $autoSize = $null
    $wordWrap = $null
    try { $autoSize = [int](Get-SafeComProperty $Shape.TextFrame2 'AutoSize' $null ([ref]$readFailed)) } catch { $readFailed = $true }
    try { $wordWrap = [int](Get-SafeComProperty $Shape.TextFrame2 'WordWrap' $null ([ref]$readFailed)) } catch { $readFailed = $true }
    # Picture crops are protected by the same invariant as geometry; capture
    # them so a future write path that touches srcRect cannot slip the gate.
    $cropLeft = $null
    $cropRight = $null
    $cropTop = $null
    $cropBottom = $null
    try {
        if ([int](Get-SafeComProperty $Shape 'Type' 0 ([ref]$readFailed)) -eq 13) {
            $cropLeft = [double](Get-SafeComProperty $Shape.PictureFormat 'CropLeft' $null ([ref]$readFailed))
            $cropRight = [double](Get-SafeComProperty $Shape.PictureFormat 'CropRight' $null ([ref]$readFailed))
            $cropTop = [double](Get-SafeComProperty $Shape.PictureFormat 'CropTop' $null ([ref]$readFailed))
            $cropBottom = [double](Get-SafeComProperty $Shape.PictureFormat 'CropBottom' $null ([ref]$readFailed))
        }
    } catch { $readFailed = $true }
    [pscustomobject]@{
        id = [int](Get-SafeComProperty $Shape 'Id' 0 ([ref]$readFailed))
        name = [string](Get-SafeComProperty $Shape 'Name' '' ([ref]$readFailed))
        type = [int](Get-SafeComProperty $Shape 'Type' 0 ([ref]$readFailed))
        left = [double](Get-SafeComProperty $Shape 'Left' 0 ([ref]$readFailed))
        top = [double](Get-SafeComProperty $Shape 'Top' 0 ([ref]$readFailed))
        width = [double](Get-SafeComProperty $Shape 'Width' 0 ([ref]$readFailed))
        height = [double](Get-SafeComProperty $Shape 'Height' 0 ([ref]$readFailed))
        rotation = [double](Get-SafeComProperty $Shape 'Rotation' 0 ([ref]$readFailed))
        zOrder = [int](Get-SafeComProperty $Shape 'ZOrderPosition' 0 ([ref]$readFailed))
        autoSize = $autoSize
        wordWrap = $wordWrap
        cropLeft = $cropLeft
        cropRight = $cropRight
        cropTop = $cropTop
        cropBottom = $cropBottom
        text = Get-SnapshotText $Shape ([ref]$readFailed)
        readStatus = if ($readFailed) { 'Unreadable' } else { 'Readable' }
    }
}

function Add-ShapeSnapshotRows {
    param(
        $Shape,
        [int]$ParentGroupId,
        [System.Collections.Generic.List[object]]$Rows,
        [ref]$UnreadableCount
    )

    try {
        $row = Get-ShapeSnapshot $Shape
        if ($ParentGroupId -gt 0) {
            $row | Add-Member -NotePropertyName groupId -NotePropertyValue $ParentGroupId -Force
        }
        $Rows.Add($row) | Out-Null
        if ([string]$row.readStatus -eq 'Unreadable') { $UnreadableCount.Value++ }

        # Grouped shapes are skipped by all write paths, so recurse through the
        # complete group tree to ensure nested child mutations remain visible to
        # the invariant comparator.
        if ([int]$row.type -eq 6) {
            $groupId = [int]$row.id
            foreach ($child in $Shape.GroupItems) {
                Add-ShapeSnapshotRows -Shape $child -ParentGroupId $groupId -Rows $Rows -UnreadableCount $UnreadableCount
            }
        }
    } catch {
        $UnreadableCount.Value++
    }
}

function Get-SlideSnapshot {
    param($Slide, [int]$Index)
    $shapeRows = New-Object System.Collections.Generic.List[object]
    $unreadableShapeCount = 0
    try {
        foreach ($shape in $Slide.Shapes) {
            Add-ShapeSnapshotRows -Shape $shape -ParentGroupId 0 -Rows $shapeRows -UnreadableCount ([ref]$unreadableShapeCount)
        }
    } catch {
        $unreadableShapeCount++
    }

    $animationRows = New-Object System.Collections.Generic.List[object]
    $animationReadFailed = $false
    try {
        $sequence = $Slide.TimeLine.MainSequence
        $sequenceCount = [int]$sequence.Count
        for ($i = 1; $i -le $sequenceCount; $i++) {
            $effect = $sequence.Item($i)
            $timing = Get-SafeComProperty $effect 'Timing' $null ([ref]$animationReadFailed)
            $effectShape = Get-SafeComProperty $effect 'Shape' $null ([ref]$animationReadFailed)
            # TriggerShape is legitimately null for ordinary click/timing
            # animations; its absence is not a read failure.
            $triggerShape = Get-SafeComProperty $timing 'TriggerShape' $null
            $triggerShapeId = 0
            if ($null -ne $triggerShape) {
                $triggerShapeId = [int](Get-SafeComProperty $triggerShape 'Id' 0 ([ref]$animationReadFailed))
            }
            $animationRows.Add([pscustomobject]@{
                index = $i
                shapeId = [int](Get-SafeComProperty $effectShape 'Id' 0 ([ref]$animationReadFailed))
                effectType = [int](Get-SafeComProperty $effect 'EffectType' 0 ([ref]$animationReadFailed))
                triggerType = [int](Get-SafeComProperty $timing 'TriggerType' 0 ([ref]$animationReadFailed))
                triggerShapeId = $triggerShapeId
                # Rounded so SaveAs float jitter in timing fields cannot become a
                # false blocker; millisecond precision is far below perception.
                duration = [Math]::Round([double](Get-SafeComProperty $timing 'Duration' 0 ([ref]$animationReadFailed)), 3)
                triggerDelayTime = [Math]::Round([double](Get-SafeComProperty $timing 'TriggerDelayTime' 0 ([ref]$animationReadFailed)), 3)
            }) | Out-Null
        }
    } catch {
        $animationReadFailed = $true
    }
    if ($animationReadFailed) { $animationRows.Clear() }

    $slideReadFailed = $false
    $transitionReadFailed = $false
    $transition = $null
    try { $transition = $Slide.SlideShowTransition } catch { $transitionReadFailed = $true }
    $transitionSnapshot = [pscustomobject]@{
        advanceOnClick = Get-SafeComProperty $transition 'AdvanceOnClick' $null ([ref]$transitionReadFailed)
        advanceOnTime = Get-SafeComProperty $transition 'AdvanceOnTime' $null ([ref]$transitionReadFailed)
        advanceTime = Get-SafeComProperty $transition 'AdvanceTime' $null ([ref]$transitionReadFailed)
        entryEffect = Get-SafeComProperty $transition 'EntryEffect' $null ([ref]$transitionReadFailed)
        speed = Get-SafeComProperty $transition 'Speed' $null ([ref]$transitionReadFailed)
    }
    [pscustomobject]@{
        index = $Index
        slideId = [int](Get-SafeComProperty $Slide 'SlideID' 0 ([ref]$slideReadFailed))
        # PowerPoint exposes the hidden-slide flag on SlideShowTransition, not
        # consistently as Slide.Hidden across desktop versions.
        hidden = [bool](Get-SafeComProperty $transition 'Hidden' $false ([ref]$transitionReadFailed))
        unreadableShapeCount = $unreadableShapeCount
        slideReadStatus = if ($slideReadFailed) { 'Unreadable' } else { 'Readable' }
        animationReadStatus = if ($animationReadFailed) { 'Unreadable' } else { 'Readable' }
        transitionReadStatus = if ($transitionReadFailed) { 'Unreadable' } else { 'Readable' }
        transition = $transitionSnapshot
        shapes = @($shapeRows.ToArray())
        animations = @($animationRows.ToArray())
    }
}

$inputFullPath = [System.IO.Path]::GetFullPath($InputPath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-PowerPointApplication
    $presentation = $powerPoint.Presentations.Open($inputFullPath, -1, -1, 0)
    $slides = @()
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $slides += Get-SlideSnapshot -Slide $presentation.Slides.Item($i) -Index $i
    }

    $snapshot = [pscustomobject]@{
        schemaVersion = 1
        generatedAt = (Get-Date).ToUniversalTime().ToString('o')
        sourcePath = $inputFullPath
        sourceSha256 = Get-FileSha256 -Path $inputFullPath
        slideWidth = [double]$presentation.PageSetup.SlideWidth
        slideHeight = [double]$presentation.PageSetup.SlideHeight
        slideCount = [int]$presentation.Slides.Count
        slides = @($slides)
        package = Get-PackageSnapshot -PptxPath $inputFullPath
    }
    $parent = Split-Path -Parent $outputFullPath
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $json = $snapshot | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($outputFullPath, $json, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    Write-Output $outputFullPath
} finally {
    if ($null -ne $presentation) { try { $presentation.Close() | Out-Null } catch { }; Release-ComObjectSafe $presentation }
    if ($null -ne $powerPoint) { try { $powerPoint.Quit() } catch { }; Release-ComObjectSafe $powerPoint }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
