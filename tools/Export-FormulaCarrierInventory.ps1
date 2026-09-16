<#
.SYNOPSIS
  Inventory formula carriers in one PPTX without modifying the source file.

.DESCRIPTION
  Combines Open XML package facts, PowerPoint COM shape geometry, the current
  formula whitelist, and the deterministic image-candidate scanners into a
  FormulaIR-backed inventory.  The tool identifies native OfficeMath,
  Equation/MathType OLE, formula-like text, formula-image candidates, mixed
  diagram/text images, grouped formula carriers, and unknown OLE carriers.

  Image and OCR/model signals are discovery evidence only.  This script never
  writes a PPTX and never treats an image candidate as canonical formula text.

.PARAMETER InputPath
  A single .pptx file to inspect.

.PARAMETER OutputDir
  Directory for formula-carrier-inventory.csv/json and the summary report.

.PARAMETER FormulaImageCandidateCsv
  Optional formula-image-candidates.csv from Export-FormulaImageCandidates.ps1.
  When omitted, the deterministic image scanners run into a private subfolder
  of OutputDir.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$FormulaImageCandidateCsv
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$script:NsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:MsoGroup = 6
$script:MsoPicture = 13
$script:MsoTextBox = 17
$script:MsoMedia = 16
$script:EmuPerPoint = 12700.0

function Get-Sha256Text {
    param([AllowEmptyString()][string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Get-ZipEntryEvidence {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = $entry.Open()
        $hash = $sha.ComputeHash($stream)
        return [pscustomobject]@{
            Path = $EntryName
            Bytes = [int64]$entry.Length
            Sha256 = ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
        }
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
        $sha.Dispose()
    }
}

function Convert-ToBoolean {
    param([object]$Value)
    if ($Value -is [bool]) { return [bool]$Value }
    $parsed = $false
    if ([bool]::TryParse([string]$Value, [ref]$parsed)) { return $parsed }
    return $false
}

function Convert-ToDouble {
    param([object]$Value)
    $parsed = 0.0
    if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return 0.0
}

function Get-RowValue {
    param(
        [object]$Row,
        [string]$Name,
        [object]$Default = ''
    )
    if ($null -eq $Row) { return $Default }
    $property = $Row.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-ImageEvidenceMap {
    param([string]$CsvPath)
    $map = @{}
    if ([string]::IsNullOrWhiteSpace($CsvPath) -or -not (Test-Path -LiteralPath $CsvPath)) { return $map }
    foreach ($row in @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8)) {
        $mediaPath = ([string](Get-RowValue -Row $row -Name 'MediaPath')).Replace('\', '/').ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($mediaPath)) { continue }
        $map[$mediaPath] = $row
    }
    return $map
}

function Test-NodeInFallback {
    param([System.Xml.XmlNode]$Node)
    $ancestor = $Node.ParentNode
    while ($null -ne $ancestor) {
        if ($ancestor.NodeType -eq [System.Xml.XmlNodeType]::Element -and $ancestor.LocalName -eq 'Fallback') { return $true }
        $ancestor = $ancestor.ParentNode
    }
    return $false
}

function Get-PreferredDescendant {
    param(
        [System.Xml.XmlNode]$Node,
        [string]$LocalName
    )
    foreach ($candidate in @($Node.SelectNodes('.//*[local-name()="' + $LocalName + '"]'))) {
        if (-not (Test-NodeInFallback -Node $candidate)) { return $candidate }
    }
    foreach ($candidate in @($Node.SelectNodes('.//*[local-name()="' + $LocalName + '"]'))) {
        return $candidate
    }
    return $null
}

function Get-NodeText {
    param([System.Xml.XmlNode]$Node)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($textNode in @($Node.SelectNodes('.//*[local-name()="t"]'))) {
        if (-not (Test-NodeInFallback -Node $textNode)) {
            $parts.Add([string]$textNode.InnerText) | Out-Null
        }
    }
    return ($parts -join '')
}

function Get-NodeMediaPaths {
    param(
        [System.Xml.XmlNode]$Node,
        [hashtable]$Relationships
    )
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($blip in @($Node.SelectNodes('.//*[local-name()="blip"]'))) {
        $relId = [string]$blip.GetAttribute('embed', $script:NsR)
        if ([string]::IsNullOrWhiteSpace($relId)) { $relId = [string]$blip.GetAttribute('r:embed') }
        if ([string]::IsNullOrWhiteSpace($relId) -or -not $Relationships.ContainsKey($relId)) { continue }
        $path = [string]$Relationships[$relId]
        if (-not $path.StartsWith('ppt/media/', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if (-not $paths.Contains($path)) { $paths.Add($path) | Out-Null }
    }
    return @($paths.ToArray())
}

function Get-NodeOleInfo {
    param([System.Xml.XmlNode]$Node)
    $progIds = New-Object System.Collections.Generic.List[string]
    $oleNodes = @($Node.SelectNodes('.//*[local-name()="oleObj"]'))
    foreach ($ole in $oleNodes) {
        $progId = [string]$ole.GetAttribute('progId')
        if (-not [string]::IsNullOrWhiteSpace($progId) -and -not $progIds.Contains($progId)) {
            $progIds.Add($progId) | Out-Null
        }
    }
    $equationLike = @($progIds | Where-Object { [string]$_ -match '(?i)(equation|mathtype|math type|mtef)' }).Count -gt 0
    return [pscustomobject]@{
        Count = $oleNodes.Count
        ProgIds = @($progIds.ToArray())
        EquationLike = $equationLike
    }
}

function Get-NodeMathCount {
    param([System.Xml.XmlNode]$Node)
    $count = 0
    foreach ($mathNode in @($Node.SelectNodes('.//*[local-name()="oMath" or local-name()="oMathPara"]'))) {
        if (-not (Test-NodeInFallback -Node $mathNode)) { $count++ }
    }
    return $count
}

function Get-SlideRelationshipMap {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$SourcePart
    )
    $map = @{}
    $relEntryName = $SourcePart -replace '^ppt/slides/(slide\d+\.xml)$', 'ppt/slides/_rels/$1.rels'
    $relText = Read-ZipEntryText -Zip $Zip -EntryName $relEntryName
    if ([string]::IsNullOrWhiteSpace($relText)) { return $map }
    $relDoc = New-Object System.Xml.XmlDocument
    $relDoc.PreserveWhitespace = $false
    $relDoc.LoadXml($relText)
    foreach ($rel in @($relDoc.GetElementsByTagName('Relationship'))) {
        if ([string]$rel.GetAttribute('TargetMode') -eq 'External') { continue }
        $id = [string]$rel.GetAttribute('Id')
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $target = Resolve-PackageTarget -SourcePart $SourcePart -Target ([string]$rel.GetAttribute('Target'))
        if (-not [string]::IsNullOrWhiteSpace($target)) { $map[$id] = $target }
    }
    return $map
}

function Test-IsFormulaCandidateText {
    param([string]$Text)
    $normalized = Get-NormalizedFormulaText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized) -or $normalized.Length -gt 80) { return $false }
    # Hyperlinks and local paths contain slash characters but are not formula
    # carriers.  Keep them out of the formula inventory deterministically.
    if ($normalized -match '(?i)(https?://|www\.|[A-Za-z]:[\\/])') { return $false }
    if ($normalized -match '[=ηΩρΔ]') { return $true }
    if ($normalized -match '([PWUIRFSη]|W有|W总|W额|G物|G动|Q吸|Q放)[=＝]') { return $true }
    if ($normalized -match '[÷×∙·√]|(W有|W总|W额|G物|G动|Q吸|Q放)') { return $true }
    return $false
}

function Get-FormulaProfile {
    param([string]$Text)
    $normalized = Get-NormalizedFormulaText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return [pscustomobject]@{ IsCandidate = $false; Normalized = ''; Kind = 'Empty'; Risk = 'ReviewOnly'; Reason = 'empty' }
    }
    if ($normalized -match '(?i)(https?://|www\.|[A-Za-z]:[\\/])') {
        return [pscustomobject]@{ IsCandidate = $false; Normalized = $normalized; Kind = 'LinkOrPath'; Risk = 'ReviewOnly'; Reason = 'url-or-path' }
    }
    $hasEquation = $normalized -match '='
    $hasDivision = $normalized -match '[/÷]'
    $hasKnownSubscript = $normalized -match '(W有|W总|W额|G物|G动|Q吸|Q放)'
    $hasSentencePunctuation = $Text -match '[，。；：、？！,;:?!]'
    $withoutKnown = $normalized -replace '(W有|W总|W额|G物|G动|Q吸|Q放)', ''
    $hasUnexpectedChinese = $withoutKnown -match '[一-龥]'
    $kind = if ($hasEquation -and $hasDivision) { 'LinearFractionEquation' } elseif ($hasEquation) { 'EquationText' } elseif ($hasDivision) { 'FractionLikeText' } elseif ($hasKnownSubscript) { 'SubscriptLikeText' } else { 'ShortMathText' }
    $reasons = New-Object System.Collections.Generic.List[string]
    if ($normalized.Length -gt 48) { $reasons.Add('too-long-for-auto-style') | Out-Null }
    if ($hasSentencePunctuation) { $reasons.Add('sentence-punctuation') | Out-Null }
    if ($hasUnexpectedChinese) { $reasons.Add('unexpected-chinese-text') | Out-Null }
    $risk = if ($reasons.Count -eq 0 -and ($hasEquation -or $hasDivision -or $hasKnownSubscript)) { 'LowRiskStandaloneText' } else { 'ReviewOnly' }
    return [pscustomobject]@{
        IsCandidate = (Test-IsFormulaCandidateText -Text $Text)
        Normalized = $normalized
        Kind = $kind
        Risk = $risk
        Reason = if ($reasons.Count -gt 0) { $reasons -join ';' } else { 'low-risk-standalone-text' }
    }
}

function Get-ComShapeTypeName {
    param([int]$Type)
    switch ($Type) {
        6 { return 'Group' }
        13 { return 'Picture' }
        16 { return 'Media' }
        17 { return 'TextBox' }
        19 { return 'Table' }
        default { return "Type$Type" }
    }
}

function Get-ComShapeText {
    param($Shape)
    try {
        if ($Shape.TextFrame2.HasText -eq -1) { return [string]$Shape.TextFrame2.TextRange.Text }
    } catch { }
    return ''
}

function Get-ComShapeMap {
    param($Presentation)
    $map = @{}
    for ($slideNo = 1; $slideNo -le [int]$Presentation.Slides.Count; $slideNo++) {
        $slide = $Presentation.Slides.Item($slideNo)
        foreach ($shape in $slide.Shapes) {
            try {
                $id = [int]$shape.Id
                $text = Get-ComShapeText -Shape $shape
                $map["$slideNo|$id"] = [pscustomobject]@{
                    Slide = $slideNo
                    Id = $id
                    Name = [string]$shape.Name
                    Type = [int]$shape.Type
                    TypeName = Get-ComShapeTypeName -Type ([int]$shape.Type)
                    Text = $text
                    Left = [double]$shape.Left
                    Top = [double]$shape.Top
                    Width = [double]$shape.Width
                    Height = [double]$shape.Height
                    Rotation = [double]$shape.Rotation
                }
            } catch {
                # Open XML remains the identity source when COM cannot read one
                # damaged or unsupported shape; the missing geometry is explicit.
            }
        }
    }
    return $map
}

function New-RawCandidate {
    param(
        [string]$Value,
        [string]$Source,
        [AllowNull()][double]$Score,
        [string[]]$Warnings = @()
    )
    return [ordered]@{
        value = $Value
        source = $Source
        score = $Score
        warnings = @($Warnings)
    }
}

function Get-CanonicalForText {
    param(
        [string]$Text,
        [object[]]$Whitelist
    )
    $match = Test-FormulaWhitelistMatch -Text $Text -Whitelist $Whitelist
    if ($null -eq $match) {
        return [pscustomobject]@{
            Match = $null
            Canonical = [ordered]@{
                status = 'Unresolved'
                source = [ordered]@{ kind = 'None'; id = ''; sha256 = $null }
                unicodeMath = ''
                tex = ''
                mathml = ''
                tokens = @()
            }
        }
    }
    $name = Get-FormulaRuleValue -Rule $match -Name 'name' -Default 'formula'
    $unicodeMath = Get-FormulaRuleValue -Rule $match -Name 'targetUnicodeMath'
    $tex = Get-FormulaRuleValue -Rule $match -Name 'targetTex'
    $canonicalHash = Get-Sha256Text -Text ("$name|$unicodeMath|$tex")
    return [pscustomobject]@{
        Match = $match
        Canonical = [ordered]@{
            # A whitelist hit is a trusted lookup key, but it is still
            # unresolved until the generated MathML/tokens pass their gates.
            status = 'Unresolved'
            source = [ordered]@{ kind = 'Whitelist'; id = $name; sha256 = $canonicalHash }
            unicodeMath = $unicodeMath
            tex = $tex
            mathml = ''
            tokens = @()
        }
    }
}

function Get-ShapeBoundsFromSignal {
    param($Xfrm)
    if ($null -eq $Xfrm) {
        return [pscustomobject]@{ Left = $null; Top = $null; Width = $null; Height = $null }
    }
    $off = $Xfrm.SelectSingleNode('./*[local-name()="off"]')
    $ext = $Xfrm.SelectSingleNode('./*[local-name()="ext"]')
    if ($null -eq $off -or $null -eq $ext) {
        return [pscustomobject]@{ Left = $null; Top = $null; Width = $null; Height = $null }
    }
    return [pscustomobject]@{
        Left = [double]$off.GetAttribute('x') / $script:EmuPerPoint
        Top = [double]$off.GetAttribute('y') / $script:EmuPerPoint
        Width = [double]$ext.GetAttribute('cx') / $script:EmuPerPoint
        Height = [double]$ext.GetAttribute('cy') / $script:EmuPerPoint
    }
}

$inputItem = Get-Item -LiteralPath $InputPath
if ($inputItem.PSIsContainer -or $inputItem.Extension.ToLowerInvariant() -ne '.pptx') {
    throw "InputPath must be a single .pptx file: $InputPath"
}
$inputFullPath = [System.IO.Path]::GetFullPath($inputItem.FullName)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $outputFullPath)) { New-Item -ItemType Directory -Path $outputFullPath -Force | Out-Null }

$resolvedFormulaImageCsv = $FormulaImageCandidateCsv
if ([string]::IsNullOrWhiteSpace($resolvedFormulaImageCsv)) {
    $imageScanDir = Join-Path $outputFullPath '_image-candidates'
    $formulaScanDir = Join-Path $outputFullPath '_formula-image-candidates'
    & (Join-Path $PSScriptRoot 'Export-PptxImageCandidates.ps1') -InputPath $inputFullPath -OutputDir $imageScanDir | Out-Null
    $imageCandidateCsv = Join-Path $imageScanDir 'pptx-image-candidates.csv'
    & (Join-Path $PSScriptRoot 'Export-FormulaImageCandidates.ps1') -ImageCandidateCsv $imageCandidateCsv -OutputDir $formulaScanDir | Out-Null
    $resolvedFormulaImageCsv = Join-Path $formulaScanDir 'formula-image-candidates.csv'
}
if (-not [string]::IsNullOrWhiteSpace($resolvedFormulaImageCsv)) {
    $resolvedFormulaImageCsv = [System.IO.Path]::GetFullPath($resolvedFormulaImageCsv)
}
$imageEvidenceMap = Get-ImageEvidenceMap -CsvPath $resolvedFormulaImageCsv
$whitelist = @(Get-PhysicsPptFormulaWhitelist)
$inputSha256 = (Get-FileHash -LiteralPath $inputFullPath -Algorithm SHA256).Hash.ToLowerInvariant()
$inputBytes = [int64](Get-Item -LiteralPath $inputFullPath).Length

$packageByKey = @{}
$mediaEvidenceCache = @{}
$zip = $null
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($inputFullPath)
    $slidePartMap = Get-PresentationOrderSlidePartMap -Zip $zip
    foreach ($slidePair in @($slidePartMap.GetEnumerator() | Sort-Object { [int]$_.Key })) {
        $slideNo = [int]$slidePair.Key
        $sourcePart = [string]$slidePair.Value
        $slideText = Read-ZipEntryText -Zip $zip -EntryName $sourcePart
        if ([string]::IsNullOrWhiteSpace($slideText)) { continue }
        $doc = New-Object System.Xml.XmlDocument
        $doc.PreserveWhitespace = $true
        $doc.LoadXml($slideText)
        $relationships = Get-SlideRelationshipMap -Zip $zip -SourcePart $sourcePart
        $spTree = $doc.SelectSingleNode('//*[local-name()="spTree"]')
        if ($null -eq $spTree) { continue }
        foreach ($child in @($spTree.ChildNodes)) {
            if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            if ($child.LocalName -notin @('sp', 'pic', 'graphicFrame', 'grpSp', 'AlternateContent')) { continue }
            $cNvPr = Get-PreferredDescendant -Node $child -LocalName 'cNvPr'
            if ($null -eq $cNvPr) { continue }
            $shapeId = 0
            if (-not [int]::TryParse([string]$cNvPr.GetAttribute('id'), [ref]$shapeId)) { continue }
            $xfrm = Get-PreferredDescendant -Node $child -LocalName 'xfrm'
            $bounds = Get-ShapeBoundsFromSignal -Xfrm $xfrm
            $ole = Get-NodeOleInfo -Node $child
            $mediaPaths = @(Get-NodeMediaPaths -Node $child -Relationships $relationships)
            $mediaEvidence = New-Object System.Collections.Generic.List[object]
            foreach ($mediaPath in $mediaPaths) {
                if (-not $mediaEvidenceCache.ContainsKey($mediaPath)) {
                    $mediaEvidenceCache[$mediaPath] = Get-ZipEntryEvidence -Zip $zip -EntryName $mediaPath
                }
                if ($null -ne $mediaEvidenceCache[$mediaPath]) { $mediaEvidence.Add($mediaEvidenceCache[$mediaPath]) | Out-Null }
            }
            $xmlHash = Get-Sha256Text -Text ([string]$child.OuterXml)
            $packageByKey["$slideNo|$shapeId"] = [pscustomobject]@{
                Slide = $slideNo
                ShapeId = $shapeId
                ShapeName = [string]$cNvPr.GetAttribute('name')
                NodeKind = [string]$child.LocalName
                PackagePart = $sourcePart
                Text = Get-NodeText -Node $child
                MathCount = Get-NodeMathCount -Node $child
                OleCount = [int]$ole.Count
                OleProgIds = @($ole.ProgIds)
                EquationOle = [bool]$ole.EquationLike
                MediaPaths = $mediaPaths
                MediaEvidence = @($mediaEvidence.ToArray())
                Bounds = $bounds
                XmlSha256 = $xmlHash
            }
        }
    }
}
finally {
    if ($null -ne $zip) { $zip.Dispose() }
}

$ppt = $null
$pres = $null
$comMap = @{}
try {
    $ppt = New-PowerPointApplication
    $pres = $ppt.Presentations.Open($inputFullPath, $true, $false, $false)
    $comMap = Get-ComShapeMap -Presentation $pres
}
finally {
    if ($null -ne $pres) { $pres.Close(); Release-ComObjectSafe $pres }
    if ($null -ne $ppt) { $ppt.Quit(); Release-ComObjectSafe $ppt }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$records = New-Object System.Collections.Generic.List[object]
$csvRows = New-Object System.Collections.Generic.List[object]
foreach ($signal in @($packageByKey.Values | Sort-Object Slide, @{ Expression = { $_.Bounds.Top } }, @{ Expression = { $_.Bounds.Left } }, ShapeId)) {
    $key = "$($signal.Slide)|$($signal.ShapeId)"
    $com = if ($comMap.ContainsKey($key)) { $comMap[$key] } else { $null }
    $text = if ($null -ne $com -and -not [string]::IsNullOrWhiteSpace([string]$com.Text)) { [string]$com.Text } else { [string]$signal.Text }
    $profile = Get-FormulaProfile -Text $text
    $whitelistResult = Get-CanonicalForText -Text $text -Whitelist $whitelist
    $mediaRows = @(
        $signal.MediaPaths |
            ForEach-Object {
                $mediaKey = ([string]$_).Replace('\', '/').ToLowerInvariant()
                if ($imageEvidenceMap.ContainsKey($mediaKey)) { $imageEvidenceMap[$mediaKey] }
            }
    )
    $formulaImageRows = @($mediaRows | Where-Object { Convert-ToBoolean (Get-RowValue -Row $_ -Name 'FormulaImageCandidate') })
    $hasFormulaImageCandidate = $formulaImageRows.Count -gt 0
    $hasMixedImageEvidence = @($formulaImageRows | Where-Object {
            [string](Get-RowValue -Row $_ -Name 'SourceCandidateLevel') -eq 'LikelyDiagramOrText'
        }).Count -gt 0

    $carrier = $null
    $fallbackCarrier = ''
    $detectionMethod = 'OpenXml'
    $detectionStatus = 'Candidate'
    $confidence = 0.0
    $decisionStatus = 'ManualRequired'
    $decisionTarget = 'None'
    $decisionReason = ''
    $rawCandidates = New-Object System.Collections.Generic.List[object]

    # A group is a protected structural boundary.  Even if it contains native
    # OMML or an OLE fallback, inventory it as GroupFormula so a future writer
    # cannot accidentally treat a child formula as a top-level replace target.
    if ($signal.NodeKind -eq 'grpSp' -and ($signal.MathCount -gt 0 -or $profile.IsCandidate -or $hasFormulaImageCandidate -or $signal.OleCount -gt 0)) {
        $carrier = 'GroupFormula'
        $confidence = 0.8
        $decisionStatus = 'ManualRequired'
        $decisionTarget = 'None'
        $decisionReason = 'Formula-like content is inside a group; grouped objects remain protected from automatic replacement.'
        $rawCandidates.Add((New-RawCandidate -Value 'grouped formula signal' -Source 'OpenXml.grpSp' -Score 0.8 -Warnings @('group-preserved'))) | Out-Null
    } elseif ($signal.MathCount -gt 0) {
        $carrier = 'OfficeMath'
        $detectionStatus = 'Observed'
        $confidence = 1.0
        $decisionStatus = 'NativeKept'
        $decisionTarget = 'OfficeMath'
        $decisionReason = 'Existing OfficeMath/OMML observed in the Open XML package; no conversion is required.'
        if ($signal.EquationOle) { $fallbackCarrier = 'MathTypeOle' }
        $rawCandidates.Add((New-RawCandidate -Value 'OfficeMath' -Source 'OpenXml.oMath' -Score 1.0)) | Out-Null
    } elseif ($signal.EquationOle) {
        $carrier = 'MathTypeOle'
        $detectionStatus = 'Observed'
        $confidence = 1.0
        $decisionStatus = 'OriginalKept'
        $decisionTarget = 'Original'
        $decisionReason = 'Equation/MathType OLE observed; MTEF content is preserved until an independently reviewed canonical export exists.'
        $rawCandidates.Add((New-RawCandidate -Value (($signal.OleProgIds -join '|')) -Source 'OpenXml.p:oleObj@progId' -Score 1.0)) | Out-Null
    } elseif ($profile.IsCandidate) {
        $carrier = 'TextFormula'
        $detectionMethod = if ($null -ne $com) { 'PowerPointCom' } else { 'OpenXml' }
        $confidence = if ($null -ne $whitelistResult.Match) { 0.9 } else { 0.65 }
        $decisionTarget = 'OfficeMath'
        if ($null -ne $whitelistResult.Match) {
            $decisionStatus = 'Skipped'
            $decisionReason = 'Formula-like text matched the current whitelist lookup, but no canonical FormulaIR/OMML write-back is authorized by inventory.'
        } else {
            $decisionStatus = 'ManualRequired'
            $decisionTarget = 'None'
            $decisionReason = 'Formula-like text has no exact current whitelist match; canonical content requires review.'
        }
        $warnings = @()
        if ($profile.Risk -ne 'LowRiskStandaloneText') { $warnings += $profile.Reason }
        $rawCandidates.Add((New-RawCandidate -Value $text -Source "$detectionMethod.TextRange" -Score $confidence -Warnings $warnings)) | Out-Null
    } elseif ($hasFormulaImageCandidate) {
        $carrier = if ($hasMixedImageEvidence) { 'MixedImage' } else { 'FormulaImage' }
        $detectionMethod = 'OpenXml'
        $bestScore = (@($formulaImageRows | ForEach-Object { Convert-ToDouble (Get-RowValue -Row $_ -Name 'FormulaImageScore') } | Measure-Object -Maximum)).Maximum
        if ($bestScore -le 0) { $bestScore = 45 }
        $confidence = [Math]::Min(1.0, [Math]::Max(0.0, $bestScore / 100.0))
        $decisionStatus = 'ManualRequired'
        $decisionTarget = 'None'
        $decisionReason = if ($carrier -eq 'MixedImage') {
            'Image heuristic found a diagram/text candidate that may mix labels, apparatus, tables, or formulas; OCR/model output is not canonical content.'
        } else {
            'Image heuristic found a possible formula image; an independently reviewed FormulaIR/GoldSet record is required before conversion.'
        }
        foreach ($imageRow in $formulaImageRows) {
            $score = Convert-ToDouble (Get-RowValue -Row $imageRow -Name 'FormulaImageScore')
            $rawCandidates.Add((New-RawCandidate -Value ([string](Get-RowValue -Row $imageRow -Name 'MediaPath')) -Source 'FormulaImageHeuristic' -Score ([Math]::Min(1.0, $score / 100.0)) -Warnings @([string](Get-RowValue -Row $imageRow -Name 'FormulaCandidateReason')))) | Out-Null
        }
    } elseif ($signal.OleCount -gt 0) {
        $carrier = 'Unknown'
        $confidence = 0.4
        $decisionStatus = 'ManualRequired'
        $decisionTarget = 'None'
        $decisionReason = 'Non-equation OLE carrier observed; inspect its ProgId and source application before any migration.'
        $rawCandidates.Add((New-RawCandidate -Value (($signal.OleProgIds -join '|')) -Source 'OpenXml.p:oleObj@progId' -Score 0.4 -Warnings @('unknown-ole'))) | Out-Null
    } else {
        continue
    }

    $bounds = if ($null -ne $com) {
        [ordered]@{ left = [double]$com.Left; top = [double]$com.Top; width = [double]$com.Width; height = [double]$com.Height }
    } else {
        [ordered]@{ left = $signal.Bounds.Left; top = $signal.Bounds.Top; width = $signal.Bounds.Width; height = $signal.Bounds.Height }
    }
    $mediaPaths = @($signal.MediaPaths)
    $mediaHashes = @($signal.MediaEvidence | ForEach-Object { [string]$_.Sha256 })
    $sourceHash = if ($mediaHashes.Count -eq 1 -and $carrier -in @('FormulaImage', 'MixedImage')) {
        $mediaHashes[0]
    } elseif ($mediaHashes.Count -gt 0 -and $carrier -in @('FormulaImage', 'MixedImage')) {
        Get-Sha256Text -Text (($mediaHashes | Sort-Object) -join '|')
    } elseif ($carrier -eq 'TextFormula') {
        Get-Sha256Text -Text (Get-NormalizedFormulaText -Text $text)
    } else {
        $signal.XmlSha256
    }
    $recordId = "s$($signal.Slide)-sh$($signal.ShapeId)-$carrier"
    $canonical = $whitelistResult.Canonical
    $source = [ordered]@{
        filePath = $inputFullPath
        fileSha256 = $inputSha256
        sourceSha256 = $sourceHash
        carrier = $carrier
        slide = [int]$signal.Slide
        shapeId = if ($null -ne $com) { [int]$com.Id } else { [int]$signal.ShapeId }
        shapeName = if ($null -ne $com -and -not [string]::IsNullOrWhiteSpace([string]$com.Name)) { [string]$com.Name } else { $signal.ShapeName }
        bbox = $bounds
        previewPath = $null
        shapeType = if ($null -ne $com) { $com.TypeName } else { [string]$signal.NodeKind }
        rotation = if ($null -ne $com) { [double]$com.Rotation } else { 0.0 }
        packagePart = [string]$signal.PackagePart
        mediaPaths = $mediaPaths
        mediaSha256 = $mediaHashes
        oleProgId = ($signal.OleProgIds -join '|')
        fallbackCarrier = $fallbackCarrier
        originalText = $text
    }
    $detection = [ordered]@{
        method = $detectionMethod
        status = $detectionStatus
        confidence = [Math]::Round($confidence, 4)
        rawCandidates = @($rawCandidates.ToArray())
    }
    $decision = [ordered]@{
        mode = 'CandidateOnly'
        status = $decisionStatus
        targetCarrier = $decisionTarget
        reason = $decisionReason
    }
    $evidencePaths = New-Object System.Collections.Generic.List[string]
    $evidencePaths.Add($inputFullPath) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($resolvedFormulaImageCsv) -and $hasFormulaImageCandidate) { $evidencePaths.Add($resolvedFormulaImageCsv) | Out-Null }
    $evidence = [ordered]@{
        paths = @($evidencePaths.ToArray())
        rollbackPath = $inputFullPath
        validatorPaths = @()
        sourcePreviewSha256 = $null
        targetPreviewSha256 = $null
    }
    $record = [ordered]@{
        schemaVersion = 1
        recordId = $recordId
        source = $source
        detection = $detection
        canonical = $canonical
        decision = $decision
        evidence = $evidence
    }
    $records.Add($record) | Out-Null
    $csvRows.Add([pscustomobject]@{
        RecordId = $recordId
        Carrier = $carrier
        Slide = [int]$signal.Slide
        ShapeId = [int]$source.shapeId
        ShapeName = [string]$source.shapeName
        ShapeType = [string]$source.shapeType
        Left = if ($null -eq $bounds.left) { '' } else { [Math]::Round([double]$bounds.left, 3) }
        Top = if ($null -eq $bounds.top) { '' } else { [Math]::Round([double]$bounds.top, 3) }
        Width = if ($null -eq $bounds.width) { '' } else { [Math]::Round([double]$bounds.width, 3) }
        Height = if ($null -eq $bounds.height) { '' } else { [Math]::Round([double]$bounds.height, 3) }
        Rotation = if ($null -eq $source.rotation) { '' } else { [Math]::Round([double]$source.rotation, 3) }
        OriginalText = $text
        WhitelistName = [string]$canonical.source.id
        UnicodeMath = [string]$canonical.unicodeMath
        Tex = [string]$canonical.tex
        DetectionMethod = $detectionMethod
        DetectionStatus = $detectionStatus
        Confidence = [Math]::Round($confidence, 4)
        DecisionMode = 'CandidateOnly'
        DecisionStatus = $decisionStatus
        TargetCarrier = $decisionTarget
        DecisionReason = $decisionReason
        SourceSha256 = $sourceHash
        FileSha256 = $inputSha256
        PackagePart = [string]$signal.PackagePart
        MediaPaths = ($mediaPaths -join ';')
        MediaSha256 = ($mediaHashes -join ';')
        OleProgId = ($signal.OleProgIds -join ';')
        FallbackCarrier = $fallbackCarrier
        EvidenceCsv = if ($hasFormulaImageCandidate) { $resolvedFormulaImageCsv } else { '' }
    }) | Out-Null
}

$csvPath = Join-Path $outputFullPath 'formula-carrier-inventory.csv'
$jsonPath = Join-Path $outputFullPath 'formula-carrier-inventory.json'
$summaryPath = Join-Path $outputFullPath 'formula-carrier-inventory-summary.md'
$csvColumns = @(
    'RecordId', 'Carrier', 'Slide', 'ShapeId', 'ShapeName', 'ShapeType',
    'Left', 'Top', 'Width', 'Height', 'Rotation', 'OriginalText',
    'WhitelistName', 'UnicodeMath', 'Tex', 'DetectionMethod',
    'DetectionStatus', 'Confidence', 'DecisionMode', 'DecisionStatus',
    'TargetCarrier', 'DecisionReason', 'SourceSha256', 'FileSha256',
    'PackagePart', 'MediaPaths', 'MediaSha256', 'OleProgId',
    'FallbackCarrier', 'EvidenceCsv'
)
if ($csvRows.Count -gt 0) {
    Write-Utf8BomCsv -InputObject @($csvRows.ToArray() | Select-Object $csvColumns) -Path $csvPath
} else {
    Write-Utf8BomText -Text (($csvColumns -join ',') + "`r`n") -Path $csvPath
}

$counts = [ordered]@{}
foreach ($carrierName in @('OfficeMath', 'MathTypeOle', 'TextFormula', 'FormulaImage', 'MixedImage', 'GroupFormula', 'Unknown')) {
    $counts[$carrierName] = @($records | Where-Object { $_.source.carrier -eq $carrierName }).Count
}
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    input = [ordered]@{ path = $inputFullPath; sha256 = $inputSha256; bytes = $inputBytes }
    imageEvidence = [ordered]@{
        formulaImageCandidateCsv = if ([string]::IsNullOrWhiteSpace($resolvedFormulaImageCsv)) { $null } else { $resolvedFormulaImageCsv }
        candidateCount = $imageEvidenceMap.Count
        source = 'DeterministicImageHeuristic'
        canonicalWriteBackAllowed = $false
    }
    oleMapping = [ordered]@{
        expectedCsv = 'formula-ole-mapping.csv'
        approvalRequirement = 'ReviewStatus=Approved'
        note = 'MathTypeOle 写回必须先经人工审定的 GoldSet（ReviewStatus=Approved），由 Export-FormulaOleMapping 校验后产出映射；本清单只是候选盘点。'
    }
    counts = $counts
    recordCount = $records.Count
    records = @($records.ToArray())
    artifacts = [ordered]@{ csv = $csvPath; json = $jsonPath; summary = $summaryPath }
}
Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 24) -Path $jsonPath

$summaryLines = New-Object System.Collections.Generic.List[string]
$summaryLines.Add('# 公式载体盘点') | Out-Null
$summaryLines.Add('') | Out-Null
$summaryLines.Add("- 输入：$inputFullPath") | Out-Null
$summaryLines.Add("- 输入 SHA-256：$inputSha256") | Out-Null
$summaryLines.Add("- 盘点记录：$($records.Count)") | Out-Null
$summaryLines.Add('') | Out-Null
$summaryLines.Add('| 载体 | 数量 | 默认决定 |') | Out-Null
$summaryLines.Add('|---|---:|---|') | Out-Null
$summaryLines.Add("| OfficeMath | $($counts.OfficeMath) | 保留原生公式 |") | Out-Null
$summaryLines.Add("| MathTypeOle | $($counts.MathTypeOle) | 保留原 OLE，待审定导出 |") | Out-Null
$summaryLines.Add("| TextFormula | $($counts.TextFormula) | 白名单仅作候选，未写回 |") | Out-Null
$summaryLines.Add("| FormulaImage | $($counts.FormulaImage) | 进入识别/白名单审查 |") | Out-Null
$summaryLines.Add("| MixedImage | $($counts.MixedImage) | 保留图片，禁止自由 OCR 写回 |") | Out-Null
$summaryLines.Add("| GroupFormula | $($counts.GroupFormula) | 组合对象保护，人工/显式迁移 |") | Out-Null
$summaryLines.Add("| Unknown | $($counts.Unknown) | 人工识别来源后再决定 |") | Out-Null
$summaryLines.Add('') | Out-Null
$summaryLines.Add('本报告是只读盘点。图片启发式、OCR 或视觉模型只产生候选，不构成 canonical 公式内容；任何 OfficeMath 写回仍需 FormulaIR、结构校验、PowerPoint 保存/导出和视觉门禁。') | Out-Null
Write-Utf8BomText -Text ($summaryLines -join "`r`n") -Path $summaryPath

Write-Output ("Formula carrier inventory done: {0}`nRecords: {1}; OfficeMath={2}; MathTypeOle={3}; TextFormula={4}; FormulaImage={5}; MixedImage={6}; GroupFormula={7}; Unknown={8}" -f $outputFullPath, $records.Count, $counts.OfficeMath, $counts.MathTypeOle, $counts.TextFormula, $counts.FormulaImage, $counts.MixedImage, $counts.GroupFormula, $counts.Unknown)
