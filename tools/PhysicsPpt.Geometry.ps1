<#
.SYNOPSIS
  Geometry conservation layer (pure ZIP/XML) for Normalize-PhysicsPpt.

.DESCRIPTION
  Verbatim extraction of the deferred geometry-heal family: AutoSize drift
  measurement, intentional-change bookkeeping, slide XML readers, xfrm
  geometry maps, and the post-save source-geometry restore. Works entirely on
  the PPTX package and never touches PowerPoint COM. Host contract for the
  dot-sourcing script: define Add-ReportRow and initialize
  $script:IntentionalGeometryShapes plus $script:PresentationSlidePartMap
  before these functions run; both state bags stay owned by
  Normalize-PhysicsPpt.ps1. No top-level statements in this file.
#>
function Get-AutoSizeGeometryDrift {
    param($Shape, [double]$Left, [double]$Top, [double]$Width, [double]$Height)
    return [Math]::Max(
        [Math]::Max([Math]::Abs(([double]$Shape.Left) - $Left), [Math]::Abs(([double]$Shape.Top) - $Top)),
        [Math]::Max([Math]::Abs(([double]$Shape.Width) - $Width), [Math]::Abs(([double]$Shape.Height) - $Height)))
}

function Add-IntentionalGeometryChange {
    # Keyed by PHYSICAL slide part name + shapeId so the XML geometry
    # restore pass (which iterates physical part names) agrees on the key
    # domain. $SlideNumber arrives as the COM presentation order and is
    # translated through the sldIdLst map; the creation-order slideN.xml
    # number drifts from the presentation order after reorders/deletes.
    param([int]$SlideNumber, [int]$ShapeId, [string]$Reason)
    if ($SlideNumber -le 0 -or $ShapeId -le 0) { return }
    $partName = $null
    if ($null -ne $script:PresentationSlidePartMap -and $script:PresentationSlidePartMap.ContainsKey([int]$SlideNumber)) {
        $partName = [string]$script:PresentationSlidePartMap[[int]$SlideNumber]
    }
    if ([string]::IsNullOrWhiteSpace($partName)) {
        # The restore pass only recognizes physical-part-name keys; a
        # creation-order fallback key would never match, so the intentional
        # change would be silently reverted by geometry conservation. Fail
        # before the caller moves the shape instead of writing a dead key.
        throw "Add-IntentionalGeometryChange: no slide part map entry for presentation slide $SlideNumber; refusing to book an unverifiable intentional geometry change (shape $ShapeId, reason '$Reason')."
    }
    $script:IntentionalGeometryShapes["$partName|$ShapeId"] = $Reason
}

function Get-GeometrySlideXmlNames {
    # Enumerates slide part names from an already-open archive only; the
    # geometry and restore passes each hold their package open, and a
    # path-based call here reopened the same zip a second time on every file.
    param($Zip)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($entry in @($Zip.Entries)) {
        if ($entry.FullName -match '^ppt/slides/slide(\d+)\.xml$') { $names.Add($entry.FullName) | Out-Null }
    }
    return $names
}

function Read-GeometrySlideXmlDocument {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $doc = New-Object System.Xml.XmlDocument
    # Slide XML is re-serialized verbatim when a geometry heal lands, so
    # whitespace-only <a:t> runs must survive Load; without this flag .NET
    # drops them and the normalized copy silently loses formula spacing.
    $doc.PreserveWhitespace = $true
    $stream = $entry.Open()
    try {
        $doc.Load($stream)
    } finally {
        $stream.Dispose()
    }
    return $doc
}

function Get-ShapeGeometryFromXfrm {
    # Enumerates text-box families of transforms (sp/pic/cxnSp spPr and grpSp
    # grpSpPr). graphicFrame p:xfrm is intentionally excluded: normalize never
    # moves OLE/table frames and the OLE AlternateContent branches share ids.
    param([System.Xml.XmlDocument]$Document)
    $shapes = New-Object System.Collections.Generic.List[object]
    foreach ($xfrm in @($Document.SelectNodes('//*[local-name()="xfrm" and namespace-uri()="http://schemas.openxmlformats.org/drawingml/2006/main"]'))) {
        $owner = $xfrm.ParentNode.ParentNode
        if ($null -eq $owner) { continue }
        $ownerLocal = [string]$owner.LocalName
        if ($ownerLocal -notin @('sp', 'pic', 'cxnSp', 'grpSp')) { continue }
        $cNvPr = $null
        foreach ($candidate in @($owner.SelectNodes('.//*[local-name()="cNvPr"]'))) {
            $cNvPr = $candidate
            break
        }
        if ($null -eq $cNvPr) { continue }
        $shapeId = 0
        if (-not [int]::TryParse([string]$cNvPr.GetAttribute('id'), [ref]$shapeId)) { continue }
        $off = $xfrm.SelectSingleNode('./*[local-name()="off"]')
        $ext = $xfrm.SelectSingleNode('./*[local-name()="ext"]')
        if ($null -eq $off -or $null -eq $ext) { continue }
        $shapes.Add([pscustomobject]@{
            ShapeId = $shapeId
            Xfrm = $xfrm
            X = [string]$off.GetAttribute('x')
            Y = [string]$off.GetAttribute('y')
            Cx = [string]$ext.GetAttribute('cx')
            Cy = [string]$ext.GetAttribute('cy')
        }) | Out-Null
    }
    return $shapes
}

function Get-SourceShapeGeometryMap {
    # Captures the exact source xfrm attribute strings per slide part. Values
    # are kept as strings and written back verbatim so the restored geometry is
    # byte-identical to the source (no EMU conversion rounding at all).
    param([string]$SourcePath)
    $map = @{}
    if (-not (Test-Path -LiteralPath $SourcePath)) { return $map }
    $zip = $null
    $zipStream = $null
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zipStream = [System.IO.File]::Open($SourcePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)
        foreach ($entryName in @(Get-GeometrySlideXmlNames -Zip $zip)) {
            if ($entryName -notmatch '^ppt/slides/slide\d+\.xml$') { continue }
            $doc = Read-GeometrySlideXmlDocument -Zip $zip -EntryName $entryName
            if ($null -eq $doc) { continue }
            foreach ($shape in @(Get-ShapeGeometryFromXfrm -Document $doc)) {
                # Keyed by physical part name so the restore pass (which also
                # iterates part names) shares one key domain; normalize never
                # renumbers parts between source and output.
                $key = "$entryName|$($shape.ShapeId)"
                if ($map.ContainsKey($key)) {
                    # Duplicate transform for one id (AlternateContent branches):
                    # conservation cannot be attributed unambiguously, so the
                    # shape is excluded from the restore instead of guessed.
                    $map[$key] = $null
                    continue
                }
                $map[$key] = [pscustomobject]@{ X = $shape.X; Y = $shape.Y; Cx = $shape.Cx; Cy = $shape.Cy }
            }
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
        if ($null -ne $zipStream) { $zipStream.Dispose() }
    }
    return $map
}

function Test-GeometryRestoreLockConflict {
    # The SaveAs target can stay locked for minutes after PowerPoint quits.
    # Lock conflicts are classified by HRESULT (0x80070020 sharing violation /
    # 0x80070021 lock violation) because the IO error message is localized on
    # the Windows PowerShell 5.1 fallback host (zh-CN throws 正由另一进程使用);
    # the known phrases stay as a net for wrapper exceptions without an HResult.
    param($Exception)
    $current = $Exception
    $visited = 0
    while ($null -ne $current -and $visited -le 5) {
        $hresult = 0
        try { $hresult = [int]$current.HResult } catch { }
        if ($hresult -eq -2147024864 -or $hresult -eq -2147024863) { return $true }
        if ([string]$current.Message -match 'being used by another process|正由另一进程使用|正在使用') { return $true }
        $current = $current.InnerException
        $visited++
    }
    return $false
}

function Restore-SourceShapeGeometry {
    # Deterministic geometry conservation. COM restores of AutoSize bounds are
    # advisory: PowerPoint recomputes them from font metrics during SaveAs
    # layout, so a re-grown box can reach the saved package even though COM
    # reported the restored value. This pass rewrites the saved slide XML back
    # to the source xfrm attribute strings unless the run recorded an explicit
    # intentional geometry change for that shape.
    # PowerPoint keeps a handle on the SaveAs target until the automation
    # process actually exits, and Quit() completion is observed to lag one to
    # a few minutes behind the call (async teardown of large decks). The pass
    # is idempotent, so the lock is polled with front-loaded one-second probes
    # for the common short-lag case and a five-second probe afterwards, with a
    # six-plus minute budget before the failure is allowed to surface.
    # The source map is read once before the loop: the source package never
    # changes during a run, and re-parsing it on every lock retry wasted a
    # full package parse per five-second wait.
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [string]$FileName
    )
    $sourceMap = Get-SourceShapeGeometryMap -SourcePath $SourcePath
    if ($sourceMap.Count -eq 0) { return 0 }
    $attempts = 84
    for ($attempt = 1; $attempt -le $attempts; $attempt++) {
        try {
            return Restore-SourceShapeGeometryOnce -OutputPath $OutputPath -FileName $FileName -SourceMap $sourceMap
        } catch {
            if ($attempt -ge $attempts) { throw }
            if (-not (Test-GeometryRestoreLockConflict -Exception $_.Exception)) { throw }
            Start-Sleep -Seconds $(if ($attempt -le 6) { 1 } else { 5 })
        }
    }
    return 0
}

function Restore-SourceShapeGeometryOnce {
    param(
        [string]$OutputPath,
        [string]$FileName,
        $SourceMap
    )
    $sourceMap = $SourceMap
    if ($sourceMap.Count -eq 0) { return 0 }
    # Report rows use presentation order everywhere else in the evidence CSV;
    # translate physical part names back through the host-owned map so the
    # page numbers stay comparable for reordered decks.
    $partToPresentationOrder = @{}
    if ($null -ne $script:PresentationSlidePartMap) {
        foreach ($order in @($script:PresentationSlidePartMap.Keys)) {
            $partToPresentationOrder[[string]$script:PresentationSlidePartMap[$order]] = [int]$order
        }
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = $null
    $zipStream = $null
    $healedTotal = 0
    try {
        # Update mode on a shared stream: allows co-existing read/write-share
        # handles (antivirus/watcher) that would reject an exclusive open.
        $zipStream = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Update)
        foreach ($entryName in @(Get-GeometrySlideXmlNames -Zip $zip)) {
            $slideNumber = 0
            if ($entryName -notmatch '^ppt/slides/slide(\d+)\.xml$') { continue }
            $slideNumber = [int]$Matches[1]
            $doc = Read-GeometrySlideXmlDocument -Zip $zip -EntryName $entryName
            if ($null -eq $doc) { continue }
            $changed = $false
            foreach ($shape in @(Get-ShapeGeometryFromXfrm -Document $doc)) {
                $key = "$entryName|$($shape.ShapeId)"
                if (-not $sourceMap.ContainsKey($key)) { continue }
                $sourceGeometry = $sourceMap[$key]
                if ($null -eq $sourceGeometry) { continue }
                if ($script:IntentionalGeometryShapes.ContainsKey($key)) { continue }
                $off = $shape.Xfrm.SelectSingleNode('./*[local-name()="off"]')
                $ext = $shape.Xfrm.SelectSingleNode('./*[local-name()="ext"]')
                if ($null -eq $off -or $null -eq $ext) { continue }
                if ([string]$off.GetAttribute('x') -eq $sourceGeometry.X -and
                    [string]$off.GetAttribute('y') -eq $sourceGeometry.Y -and
                    [string]$ext.GetAttribute('cx') -eq $sourceGeometry.Cx -and
                    [string]$ext.GetAttribute('cy') -eq $sourceGeometry.Cy) { continue }
                $before = '{0}|{1}|{2}|{3}' -f $off.GetAttribute('x'), $off.GetAttribute('y'), $ext.GetAttribute('cx'), $ext.GetAttribute('cy')
                $off.SetAttribute('x', $sourceGeometry.X)
                $off.SetAttribute('y', $sourceGeometry.Y)
                $ext.SetAttribute('cx', $sourceGeometry.Cx)
                $ext.SetAttribute('cy', $sourceGeometry.Cy)
                $after = '{0}|{1}|{2}|{3}' -f $sourceGeometry.X, $sourceGeometry.Y, $sourceGeometry.Cx, $sourceGeometry.Cy
                $reportSlideNumber = if ($partToPresentationOrder.ContainsKey($entryName)) { $partToPresentationOrder[$entryName] } else { $slideNumber }
                Add-ReportRow -File $FileName -SlideNumber $reportSlideNumber -ShapeName "shapeId=$($shape.ShapeId)" `
                    -Issue 'AutoSizeDeferredReflowHealed' `
                    -Details 'PowerPoint re-grew the AutoSize bounds during SaveAs layout; the saved slide XML was restored to the exact source xfrm values.' `
                    -RuleId 'SAFETY.GEOMETRY.AUTOSIZE' -Property 'xfrm/off/ext' -Before $before -After $after `
                    -RiskLevel 'R1' -Result 'Applied'
                $healedTotal++
                $changed = $true
            }
            if ($changed) {
                $entry = $zip.GetEntry($entryName)
                if ($null -eq $entry) { continue }
                $settings = New-Object System.Xml.XmlWriterSettings
                $settings.Indent = $false
                $settings.OmitXmlDeclaration = $false
                $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
                $stream = New-Object System.IO.MemoryStream
                try {
                    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
                    try {
                        $doc.Save($writer)
                        $writer.Flush()
                    } finally {
                        $writer.Dispose()
                    }
                    $xmlBytes = $stream.ToArray()
                } finally {
                    $stream.Dispose()
                }
                $entry.Delete()
                $newEntry = $zip.CreateEntry($entryName)
                $entryStream = $newEntry.Open()
                try {
                    $entryStream.Write($xmlBytes, 0, $xmlBytes.Length)
                } finally {
                    $entryStream.Dispose()
                }
            }
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
        if ($null -ne $zipStream) { $zipStream.Dispose() }
    }
    if ($healedTotal -gt 0) {
        # Verify the patch landed: reopen the package and compare every
        # recorded shape against the source xfrm values.
        $verifyDoc = $null
        $map2 = Get-SourceShapeGeometryMap -SourcePath $OutputPath
        foreach ($key in @($sourceMap.Keys)) {
            $expected = $sourceMap[$key]
            if ($null -eq $expected -or -not $map2.ContainsKey($key) -or $null -eq $map2[$key]) { continue }
            $actual = $map2[$key]
            if ([string]$actual.X -ne [string]$expected.X -or [string]$actual.Y -ne [string]$expected.Y -or
                [string]$actual.Cx -ne [string]$expected.Cx -or [string]$actual.Cy -ne [string]$expected.Cy) {
                throw "Geometry restore verification failed for slide shape '$key': saved '$($actual.X)|$($actual.Y)|$($actual.Cx)|$($actual.Cy)' expected '$($expected.X)|$($expected.Y)|$($expected.Cx)|$($expected.Cy)'."
            }
        }
    }
    return $healedTotal
}
