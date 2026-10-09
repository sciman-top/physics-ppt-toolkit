<#
.SYNOPSIS
  Create a hash-bound, read-only visual adjudication proposal for current MathType/OLE crops.

.DESCRIPTION
  Joins the current carrier inventory, OLE crop manifest, and a human/AI-prepared
  adjudication CSV. Every current MathTypeOle record must appear exactly once.
  The proposal records visual evidence and a possible canonical FormulaIR input,
  but never writes a PPTX and never enables write-back. CandidateOnly means that
  a later explicit migration step may consider the row; it is not an approval.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CarrierInventoryJson,
    [Parameter(Mandatory = $true)][string]$CropManifestJson,
    [Parameter(Mandatory = $true)][string]$AdjudicationCsv,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Require-Column { param($Rows, [string]$Name) if ($Rows.Count -eq 0 -or $null -eq $Rows[0].PSObject.Properties[$Name]) { throw "Adjudication CSV is missing required column: $Name" } }

$inventoryPath = [IO.Path]::GetFullPath($CarrierInventoryJson)
$cropManifestPath = [IO.Path]::GetFullPath($CropManifestJson)
$adjudicationPath = [IO.Path]::GetFullPath($AdjudicationCsv)
$outputPath = [IO.Path]::GetFullPath($OutputDir)
foreach ($path in @($inventoryPath, $cropManifestPath, $adjudicationPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required input not found: $path" }
}
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

$inventory = Get-Content -LiteralPath $inventoryPath -Raw -Encoding UTF8 | ConvertFrom-Json
$cropManifest = Get-Content -LiteralPath $cropManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$inventory.schemaVersion -ne 1 -or [int]$cropManifest.schemaVersion -ne 1) { throw 'Unsupported inventory or crop manifest schemaVersion.' }
if ([bool]$cropManifest.writeBackAllowed) { throw 'Crop manifest illegally enables write-back.' }

$sourcePath = [IO.Path]::GetFullPath([string]$inventory.input.path)
if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "Inventory source PPTX not found: $sourcePath" }
$sourceSha256 = Get-FileSha256Hex -Path $sourcePath
if ($sourceSha256 -ne ([string]$inventory.input.sha256).ToLowerInvariant()) { throw 'Inventory source PPTX hash does not match current input.' }
if ($sourceSha256 -ne ([string]$cropManifest.input.sha256).ToLowerInvariant()) { throw 'Crop manifest is bound to a different source PPTX.' }
$inventorySha256 = Get-FileSha256Hex -Path $inventoryPath
if ($inventorySha256 -ne ([string]$cropManifest.inventory.sha256).ToLowerInvariant()) { throw 'Crop manifest inventory hash does not match current inventory.' }

$cropCsvPath = Join-Path (Split-Path -Parent $cropManifestPath) 'ole-crops.csv'
if (-not (Test-Path -LiteralPath $cropCsvPath -PathType Leaf)) { throw "OLE crop CSV not found beside manifest: $cropCsvPath" }
$cropRows = @(Import-Csv -LiteralPath $cropCsvPath -Encoding UTF8)
$cropById = @{}
foreach ($row in $cropRows) {
    $id = [string]$row.RecordId
    if ([string]::IsNullOrWhiteSpace($id) -or $cropById.ContainsKey($id)) { throw "OLE crop CSV has an empty or duplicate record: $id" }
    $cropById[$id] = $row
}

$rows = @(Import-Csv -LiteralPath $adjudicationPath -Encoding UTF8)
foreach ($column in @('RecordId', 'SourceSha256', 'CropPng', 'CropSha256', 'PageSha256', 'Status', 'TargetCarrier', 'UnicodeMath', 'TeX', 'Basis', 'ContextAssessment', 'Reviewer', 'Confidence')) { Require-Column -Rows $rows -Name $column }
$inventoryOle = @($inventory.records | Where-Object { [string]$_.source.carrier -eq 'MathTypeOle' })
$inventoryById = @{}
foreach ($record in $inventoryOle) {
    $id = [string]$record.recordId
    if ($inventoryById.ContainsKey($id)) { throw "Inventory has duplicate MathTypeOle record: $id" }
    $inventoryById[$id] = $record
}
if ($rows.Count -ne $inventoryOle.Count) { throw "Adjudication coverage mismatch: rows=$($rows.Count), current MathTypeOle records=$($inventoryOle.Count)" }

$seen = @{}
$outputRecords = New-Object System.Collections.Generic.List[object]
foreach ($row in $rows) {
    $id = [string]$row.RecordId
    if ($seen.ContainsKey($id)) { throw "Adjudication CSV has duplicate record: $id" }
    $seen[$id] = $true
    if (-not $inventoryById.ContainsKey($id)) { throw "Adjudication record is not a current MathTypeOle inventory record: $id" }
    if (-not $cropById.ContainsKey($id)) { throw "Missing crop evidence for current record: $id" }
    $record = $inventoryById[$id]
    $crop = $cropById[$id]
    if (([string]$row.SourceSha256).ToLowerInvariant() -ne ([string]$record.source.sourceSha256).ToLowerInvariant()) { throw "Source object hash mismatch: $id" }
    if (([string]$row.CropPng) -ne [string]$crop.CropPng -or ([string]$row.CropSha256).ToLowerInvariant() -ne ([string]$crop.CropSha256).ToLowerInvariant()) { throw "Crop path or hash mismatch: $id" }
    if (([string]$row.PageSha256).ToLowerInvariant() -ne ([string]$crop.PageSha256).ToLowerInvariant()) { throw "Page hash mismatch: $id" }
    if (-not (Test-Path -LiteralPath ([string]$row.CropPng) -PathType Leaf)) { throw "Crop image not found: $id" }
    if ((Get-FileSha256Hex -Path ([string]$row.CropPng)) -ne ([string]$row.CropSha256).ToLowerInvariant()) { throw "Crop image has drifted since export: $id" }
    $status = [string]$row.Status
    $target = [string]$row.TargetCarrier
    if ($status -notin @('CandidateOnly', 'ManualRequired', 'OriginalKept')) { throw "Unsupported visual adjudication status '$status' for $id" }
    if ($status -eq 'CandidateOnly' -and ($target -ne 'OfficeMath' -or [string]::IsNullOrWhiteSpace([string]$row.UnicodeMath) -or [string]::IsNullOrWhiteSpace([string]$row.TeX))) { throw "CandidateOnly requires OfficeMath plus UnicodeMath and TeX: $id" }
    if ($status -ne 'CandidateOnly' -and $target -ne 'Original') { throw "Non-candidate visual adjudication must keep Original: $id" }
    $confidence = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$row.Confidence)) { $confidence = [double]$row.Confidence; if ($confidence -lt 0 -or $confidence -gt 1) { throw "Confidence must be between 0 and 1: $id" } }
    $outputRecords.Add([ordered]@{
        recordId = $id
        source = [ordered]@{
            fileSha256 = [string]$record.source.fileSha256
            sourceSha256 = [string]$record.source.sourceSha256
            carrier = [string]$record.source.carrier
            slide = [int]$record.source.slide
            shapeId = [int]$record.source.shapeId
            mediaSha256 = @($record.source.mediaSha256)
        }
        crop = [ordered]@{ path = [string]$row.CropPng; sha256 = ([string]$row.CropSha256).ToLowerInvariant(); pageSha256 = ([string]$row.PageSha256).ToLowerInvariant() }
        proposal = [ordered]@{
            status = $status
            targetCarrier = $target
            unicodeMath = [string]$row.UnicodeMath
            tex = [string]$row.TeX
            basis = [string]$row.Basis
            contextAssessment = [string]$row.ContextAssessment
            reviewer = if ([string]::IsNullOrWhiteSpace([string]$row.Reviewer)) { $null } else { [string]$row.Reviewer }
            confidence = $confidence
            writeBackAllowed = $false
        }
    }) | Out-Null
}
if ($seen.Count -ne $inventoryById.Count) { throw 'Adjudication CSV did not cover every current MathTypeOle record.' }

$recordsSorted = @($outputRecords.ToArray() | Sort-Object recordId)
$evidenceLines = New-Object System.Collections.Generic.List[string]
$evidenceLines.Add("source|$sourceSha256") | Out-Null
$evidenceLines.Add("inventory|$inventorySha256") | Out-Null
$evidenceLines.Add("cropManifest|$(Get-FileSha256Hex -Path $cropManifestPath)") | Out-Null
$evidenceLines.Add("adjudicationCsv|$(Get-FileSha256Hex -Path $adjudicationPath)") | Out-Null
foreach ($r in $recordsSorted) { $evidenceLines.Add("$($r.recordId)|$($r.source.sourceSha256)|$($r.crop.sha256)|$($r.crop.pageSha256)|$($r.proposal.status)|$($r.proposal.unicodeMath)|$($r.proposal.tex)") | Out-Null }
$counts = [ordered]@{ total = $recordsSorted.Count; candidateOnly = @($recordsSorted | Where-Object { $_.proposal.status -eq 'CandidateOnly' }).Count; manualRequired = @($recordsSorted | Where-Object { $_.proposal.status -eq 'ManualRequired' }).Count; originalKept = @($recordsSorted | Where-Object { $_.proposal.status -eq 'OriginalKept' }).Count }
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ kind = 'CurrentOleVisualAdjudicationProposal'; readOnly = $true; planOnly = $true; writeBackAllowed = $false }
    input = [ordered]@{
        sourcePptx = [ordered]@{ path = $sourcePath; sha256 = $sourceSha256 }
        inventory = [ordered]@{ path = $inventoryPath; sha256 = $inventorySha256 }
        cropManifest = [ordered]@{ path = $cropManifestPath; sha256 = Get-FileSha256Hex -Path $cropManifestPath }
        adjudicationCsv = [ordered]@{ path = $adjudicationPath; sha256 = Get-FileSha256Hex -Path $adjudicationPath }
    }
    counts = $counts
    records = $recordsSorted
    evidenceSetSha256 = Get-TextSha256Hex -Text (($evidenceLines.ToArray() | Sort-Object) -join "`n")
}
$jsonPath = Join-Path $outputPath 'formula-ole-visual-adjudication-proposal.json'
$csvPath = Join-Path $outputPath 'formula-ole-visual-adjudication-proposal.csv'
Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 20) -Path $jsonPath
$csvOutput = @($recordsSorted | ForEach-Object { [pscustomobject]@{ RecordId = $_.recordId; Slide = $_.source.slide; ShapeId = $_.source.shapeId; SourceSha256 = $_.source.sourceSha256; CropSha256 = $_.crop.sha256; Status = $_.proposal.status; TargetCarrier = $_.proposal.targetCarrier; UnicodeMath = $_.proposal.unicodeMath; TeX = $_.proposal.tex; ContextAssessment = $_.proposal.contextAssessment; WriteBackAllowed = $_.proposal.writeBackAllowed } })
Write-Utf8BomCsv -InputObject $csvOutput -Path $csvPath
Write-Output "Current OLE visual adjudication proposal complete: $outputPath`nRecords: $($recordsSorted.Count); CandidateOnly: $($counts.candidateOnly); ManualRequired: $($counts.manualRequired); OriginalKept: $($counts.originalKept); writeBackAllowed: false"
