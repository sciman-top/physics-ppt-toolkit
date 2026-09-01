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
    param($Shape)
    try {
        if ($Shape.TextFrame2.HasText -eq -1) { return [string]$Shape.TextFrame2.TextRange.Text }
    } catch { }
    return ''
}

function Get-SafeComProperty {
    param($Object, [string]$Name, $Default = $null)
    try { return $Object.$Name } catch { return $Default }
}

function Get-ShapeSnapshot {
    param($Shape)
    $autoSize = $null
    $wordWrap = $null
    try { $autoSize = [int]$Shape.TextFrame2.AutoSize } catch { }
    try { $wordWrap = [int]$Shape.TextFrame2.WordWrap } catch { }
    [pscustomobject]@{
        id = [int](Get-SafeComProperty $Shape 'Id' 0)
        name = [string](Get-SafeComProperty $Shape 'Name' '')
        type = [int](Get-SafeComProperty $Shape 'Type' 0)
        left = [double](Get-SafeComProperty $Shape 'Left' 0)
        top = [double](Get-SafeComProperty $Shape 'Top' 0)
        width = [double](Get-SafeComProperty $Shape 'Width' 0)
        height = [double](Get-SafeComProperty $Shape 'Height' 0)
        rotation = [double](Get-SafeComProperty $Shape 'Rotation' 0)
        zOrder = [int](Get-SafeComProperty $Shape 'ZOrderPosition' 0)
        autoSize = $autoSize
        wordWrap = $wordWrap
        text = Get-SnapshotText $Shape
    }
}

function Get-SlideSnapshot {
    param($Slide, [int]$Index)
    $shapeRows = @()
    foreach ($shape in $Slide.Shapes) {
        try { $shapeRows += Get-ShapeSnapshot $shape } catch { }
    }

    $animationRows = @()
    try {
        $sequence = $Slide.TimeLine.MainSequence
        for ($i = 1; $i -le $sequence.Count; $i++) {
            $effect = $sequence.Item($i)
            $animationRows += [pscustomobject]@{
                index = $i
                shapeId = [int](Get-SafeComProperty $effect.Shape 'Id' 0)
                effectType = [int](Get-SafeComProperty $effect 'EffectType' 0)
                triggerType = [int](Get-SafeComProperty $effect.Timing 'TriggerType' 0)
                triggerShapeId = [int](Get-SafeComProperty (Get-SafeComProperty (Get-SafeComProperty $effect 'Timing' $null) 'TriggerShape' $null) 'Id' 0)
                duration = [double](Get-SafeComProperty $effect.Timing 'Duration' 0)
                triggerDelayTime = [double](Get-SafeComProperty $effect.Timing 'TriggerDelayTime' 0)
            }
        }
    } catch { }

    [pscustomobject]@{
        index = $Index
        slideId = [int](Get-SafeComProperty $Slide 'SlideID' 0)
        hidden = [bool](Get-SafeComProperty $Slide 'Hidden' $false)
        transition = [pscustomobject]@{
            advanceOnClick = Get-SafeComProperty $Slide.SlideShowTransition 'AdvanceOnClick' $null
            advanceOnTime = Get-SafeComProperty $Slide.SlideShowTransition 'AdvanceOnTime' $null
            advanceTime = Get-SafeComProperty $Slide.SlideShowTransition 'AdvanceTime' $null
            entryEffect = Get-SafeComProperty $Slide.SlideShowTransition 'EntryEffect' $null
            speed = Get-SafeComProperty $Slide.SlideShowTransition 'Speed' $null
        }
        shapes = @($shapeRows)
        animations = @($animationRows)
    }
}

$inputFullPath = [System.IO.Path]::GetFullPath($InputPath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
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
