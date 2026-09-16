<#
.SYNOPSIS
  Convert an approved OLE GoldSet into safe OMML mapping/review inputs.

.DESCRIPTION
  Validates a curated GoldSet against formula-carrier-inventory.json, the
  current whitelist, the source PPTX hash, and the visual evidence paths.  It
  then emits the MappingCsv consumed by Apply-FormulaOmmlForOle and the
  formula-review.csv consumed by Export-FormulaOmmlCandidates.

  This adapter is candidate generation only.  It never edits a PPTX.  A row
  must identify Equation/MathType OLE shape IDs and use ReviewStatus=Approved;
  source formula text and canonical values are resolved from the current
  whitelist rather than from OCR or a visual model response.

.PARAMETER CarrierInventoryJson
  formula-carrier-inventory.json produced by Export-FormulaCarrierInventory.

.PARAMETER GoldSetCsv
  Approved CSV with columns:
  ReviewStatus,Slide,ShapeIds,WhitelistName,SourceFormulaText,SizePt,
  MainColorHex,SubColorHex,Note,EvidencePath.

.PARAMETER OutputDir
  Directory for formula-ole-mapping.csv, formula-ole-review.csv, and the
  validation manifest.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CarrierInventoryJson,

    [Parameter(Mandatory = $true)]
    [string]$GoldSetCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [ValidateRange(1, 100)]
    [int]$MaxItems = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

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

function Write-CsvWithHeader {
    param(
        [object[]]$Rows,
        [string[]]$Columns,
        [string]$Path
    )
    if (@($Rows).Count -gt 0) {
        Write-Utf8BomCsv -InputObject @($Rows | Select-Object $Columns) -Path $Path
    } else {
        Write-Utf8BomText -Text (($Columns -join ',') + "`r`n") -Path $Path
    }
}

function Convert-ToHex {
    param([string]$Value)
    $hex = ([string]$Value).Trim().TrimStart('#').ToUpperInvariant()
    if ($hex -notmatch '^[0-9A-F]{6}$') { return '' }
    return $hex
}

function Get-InventoryRecordKey {
    param([int]$Slide, [int]$ShapeId)
    return "$Slide|$ShapeId"
}

$inventoryPath = [System.IO.Path]::GetFullPath($CarrierInventoryJson)
$goldSetPath = [System.IO.Path]::GetFullPath($GoldSetCsv)
$outputDir = [System.IO.Path]::GetFullPath($OutputDir)
foreach ($path in @($inventoryPath, $goldSetPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required input not found: $path" }
}
if (-not (Test-Path -LiteralPath $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

$inventory = Get-Content -LiteralPath $inventoryPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$inventory.schemaVersion -ne 1) { throw "Unsupported carrier inventory schemaVersion: $($inventory.schemaVersion)" }
$inputPath = [System.IO.Path]::GetFullPath([string]$inventory.input.path)
if (-not (Test-Path -LiteralPath $inputPath)) { throw "Inventory source PPTX not found: $inputPath" }
$actualInputSha256 = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualInputSha256 -ne ([string]$inventory.input.sha256).ToLowerInvariant()) {
    throw "Inventory source hash does not match the current PPTX: expected $($inventory.input.sha256), actual $actualInputSha256"
}

$oleRecords = @(
    $inventory.records |
        Where-Object { [string]$_.source.carrier -eq 'MathTypeOle' } |
        Sort-Object @{ Expression = { [int]$_.source.slide } }, @{ Expression = { [double]$_.source.bbox.top } }, @{ Expression = { [double]$_.source.bbox.left } }, @{ Expression = { [int]$_.source.shapeId } }
)
$indexByKey = @{}
for ($i = 0; $i -lt $oleRecords.Count; $i++) {
    $record = $oleRecords[$i]
    $key = Get-InventoryRecordKey -Slide ([int]$record.source.slide) -ShapeId ([int]$record.source.shapeId)
    if ($indexByKey.ContainsKey($key)) { throw "Duplicate OLE inventory identity: $key" }
    $indexByKey[$key] = [pscustomobject]@{ Index = $i; Record = $record }
}

$whitelist = @(Get-PhysicsPptFormulaWhitelist)
$goldRows = @(Import-Csv -LiteralPath $goldSetPath -Encoding UTF8 | Select-Object -First $MaxItems)
$mappingRows = New-Object System.Collections.Generic.List[object]
$reviewRows = New-Object System.Collections.Generic.List[object]
$validatedRows = New-Object System.Collections.Generic.List[object]
$usedIndexes = @{}
$errors = New-Object System.Collections.Generic.List[string]

foreach ($row in $goldRows) {
    $rowLabel = "slide=$([string](Get-RowValue -Row $row -Name 'Slide')); shapes=$([string](Get-RowValue -Row $row -Name 'ShapeIds')); whitelist=$([string](Get-RowValue -Row $row -Name 'WhitelistName'))"
    if ([string](Get-RowValue -Row $row -Name 'ReviewStatus') -ne 'Approved') {
        $errors.Add("${rowLabel}: ReviewStatus must be Approved") | Out-Null
        continue
    }
    $slideNo = 0
    if (-not [int]::TryParse([string](Get-RowValue -Row $row -Name 'Slide'), [ref]$slideNo) -or $slideNo -lt 1) {
        $errors.Add("${rowLabel}: invalid Slide") | Out-Null
        continue
    }
    $shapeIdTexts = @([string](Get-RowValue -Row $row -Name 'ShapeIds') -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    if ($shapeIdTexts.Count -eq 0) {
        $errors.Add("${rowLabel}: ShapeIds is empty") | Out-Null
        continue
    }
    $shapeIndexes = New-Object System.Collections.Generic.List[int]
    $shapeRecords = New-Object System.Collections.Generic.List[object]
    $rowValid = $true
    foreach ($shapeIdText in $shapeIdTexts) {
        $shapeId = 0
        if (-not [int]::TryParse($shapeIdText, [ref]$shapeId)) {
            $errors.Add("${rowLabel}: invalid shape ID '$shapeIdText'") | Out-Null
            $rowValid = $false
            continue
        }
        $key = Get-InventoryRecordKey -Slide $slideNo -ShapeId $shapeId
        if (-not $indexByKey.ContainsKey($key)) {
            $errors.Add("${rowLabel}: OLE shape not found in inventory: $key") | Out-Null
            $rowValid = $false
            continue
        }
        $entry = $indexByKey[$key]
        if ($usedIndexes.ContainsKey([int]$entry.Index)) {
            $errors.Add("${rowLabel}: OLE index $($entry.Index) is already used") | Out-Null
            $rowValid = $false
            continue
        }
        if ([string]$entry.Record.source.carrier -ne 'MathTypeOle') {
            $errors.Add("${rowLabel}: source carrier is not MathTypeOle") | Out-Null
            $rowValid = $false
            continue
        }
        $shapeIndexes.Add([int]$entry.Index) | Out-Null
        $shapeRecords.Add($entry.Record) | Out-Null
    }

    $whitelistName = [string](Get-RowValue -Row $row -Name 'WhitelistName')
    $sourceFormulaText = [string](Get-RowValue -Row $row -Name 'SourceFormulaText')
    $match = Test-FormulaWhitelistMatch -Text $sourceFormulaText -Whitelist $whitelist
    if ($null -eq $match -or [string]$match.name -ne $whitelistName) {
        $errors.Add("${rowLabel}: SourceFormulaText does not resolve to the named current whitelist rule") | Out-Null
        $rowValid = $false
    }
    $evidencePathTexts = @([string](Get-RowValue -Row $row -Name 'EvidencePath') -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $fullEvidencePaths = @($evidencePathTexts | ForEach-Object { [System.IO.Path]::GetFullPath($_) })
    $missingEvidencePaths = @($fullEvidencePaths | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($fullEvidencePaths.Count -eq 0 -or $missingEvidencePaths.Count -gt 0) {
        $errors.Add("${rowLabel}: EvidencePath is missing or not found: $($missingEvidencePaths -join ',')") | Out-Null
        $rowValid = $false
    }
    $sizePt = 0
    if (-not [int]::TryParse([string](Get-RowValue -Row $row -Name 'SizePt'), [ref]$sizePt) -or $sizePt -lt 8 -or $sizePt -gt 96) {
        $errors.Add("${rowLabel}: SizePt must be an integer from 8 to 96") | Out-Null
        $rowValid = $false
    }
    $mainHex = Convert-ToHex -Value ([string](Get-RowValue -Row $row -Name 'MainColorHex'))
    $subHex = Convert-ToHex -Value ([string](Get-RowValue -Row $row -Name 'SubColorHex'))
    $accentHex = Convert-ToHex -Value ([string](Get-RowValue -Row $row -Name 'AccentColorHex'))
    $accentTokens = [string](Get-RowValue -Row $row -Name 'AccentTokens')
    if ([string]::IsNullOrWhiteSpace($mainHex) -or [string]$mainHex -notmatch '^[0-9A-F]{6}$') {
        $errors.Add("${rowLabel}: MainColorHex is invalid") | Out-Null
        $rowValid = $false
    }
    if (-not [string]::IsNullOrWhiteSpace([string](Get-RowValue -Row $row -Name 'SubColorHex')) -and [string]::IsNullOrWhiteSpace($subHex)) {
        $errors.Add("${rowLabel}: SubColorHex is invalid") | Out-Null
        $rowValid = $false
    }
    if (-not [string]::IsNullOrWhiteSpace([string](Get-RowValue -Row $row -Name 'AccentColorHex')) -and [string]::IsNullOrWhiteSpace($accentHex)) {
        $errors.Add("${rowLabel}: AccentColorHex is invalid") | Out-Null
        $rowValid = $false
    }
    if (-not [string]::IsNullOrWhiteSpace($accentTokens) -and [string]::IsNullOrWhiteSpace($accentHex)) {
        $errors.Add("${rowLabel}: AccentTokens requires AccentColorHex") | Out-Null
        $rowValid = $false
    }
    if (-not $rowValid) { continue }

    foreach ($index in $shapeIndexes) { $usedIndexes[$index] = $true }
    $targetUnicodeMath = Get-FormulaRuleValue -Rule $match -Name 'targetUnicodeMath'
    $targetTex = Get-FormulaRuleValue -Rule $match -Name 'targetTex'
    $note = [string](Get-RowValue -Row $row -Name 'Note')
    $recordIds = @($shapeRecords | ForEach-Object { [string]$_.recordId })
    $shapeNames = @($shapeRecords | ForEach-Object { [string]$_.source.shapeName })
    $mappingRows.Add([pscustomobject]@{
        Slide = $slideNo
        OleIndex = (($shapeIndexes | ForEach-Object { [string]$_ }) -join ';')
        WhitelistName = $whitelistName
        SizePt = $sizePt
        MainColorHex = $mainHex
        SubColorHex = $subHex
        AccentColorHex = $accentHex
        AccentTokens = $accentTokens
        Note = $note
        InventoryRecordIds = ($recordIds -join ';')
        ShapeIds = ($shapeIdTexts -join ';')
        EvidencePath = ($fullEvidencePaths -join ';')
    }) | Out-Null
    $whitelistCandidate = "name=$whitelistName; targetUnicodeMath=$targetUnicodeMath; targetTex=$targetTex; note=$note"
    $reviewRows.Add([pscustomobject]@{
        File = [System.IO.Path]::GetFileName($inputPath)
        FilePath = $inputPath
        FileRelativePath = ''
        Slide = $slideNo
        Shape = ($shapeNames -join ';')
        FormulaText = $sourceFormulaText
        WhitelistCandidate = $whitelistCandidate
        SuggestedAction = 'ReviewWhitelistConversion'
        SourceKind = 'GoldSet'
        EvidencePath = ($fullEvidencePaths -join ';')
        TargetUnicodeMath = $targetUnicodeMath
        TargetTex = $targetTex
        AccentColorHex = $accentHex
        AccentTokens = $accentTokens
    }) | Out-Null
    $validatedRows.Add([pscustomobject]@{
        Slide = $slideNo
        ShapeIds = ($shapeIdTexts -join ';')
        OleIndex = (($shapeIndexes | ForEach-Object { [string]$_ }) -join ';')
        WhitelistName = $whitelistName
        SourceFormulaText = $sourceFormulaText
        TargetUnicodeMath = $targetUnicodeMath
        TargetTex = $targetTex
        AccentColorHex = $accentHex
        AccentTokens = $accentTokens
        EvidencePath = ($fullEvidencePaths -join ';')
        InventoryRecordIds = ($recordIds -join ';')
    }) | Out-Null
}

$mappingPath = Join-Path $outputDir 'formula-ole-mapping.csv'
$reviewPath = Join-Path $outputDir 'formula-ole-review.csv'
$manifestPath = Join-Path $outputDir 'formula-ole-goldset-manifest.json'
$mappingColumns = @('Slide', 'OleIndex', 'WhitelistName', 'SizePt', 'MainColorHex', 'SubColorHex', 'AccentColorHex', 'AccentTokens', 'Note', 'InventoryRecordIds', 'ShapeIds', 'EvidencePath')
$reviewColumns = @('File', 'FilePath', 'FileRelativePath', 'Slide', 'Shape', 'FormulaText', 'WhitelistCandidate', 'SuggestedAction', 'SourceKind', 'EvidencePath', 'TargetUnicodeMath', 'TargetTex')
Write-CsvWithHeader -Rows @($mappingRows.ToArray()) -Columns $mappingColumns -Path $mappingPath
Write-CsvWithHeader -Rows @($reviewRows.ToArray()) -Columns $reviewColumns -Path $reviewPath

$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    input = [ordered]@{ path = $inputPath; sha256 = $actualInputSha256 }
    inventory = [ordered]@{ path = $inventoryPath; sha256 = (Get-FileHash -LiteralPath $inventoryPath -Algorithm SHA256).Hash.ToLowerInvariant(); oleCount = $oleRecords.Count }
    goldSet = [ordered]@{ path = $goldSetPath; sha256 = (Get-FileHash -LiteralPath $goldSetPath -Algorithm SHA256).Hash.ToLowerInvariant(); requestedCount = $goldRows.Count; approvedCount = $validatedRows.Count }
    status = if ($errors.Count -gt 0) { 'Failed' } else { 'Passed' }
    errors = @($errors.ToArray())
    mappingCsv = $mappingPath
    formulaReviewCsv = $reviewPath
    rows = @($validatedRows.ToArray())
    writeBackAllowed = $false
    note = 'GoldSet validation and candidate generation only; Apply-FormulaOmmlForOle remains a separate explicit write-back step.'
}
Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 12) -Path $manifestPath

if ($errors.Count -gt 0) {
    throw ("Formula OLE GoldSet validation failed: {0}" -f ($errors -join '; '))
}
Write-Output ("Formula OLE mapping done: {0}`nApproved rows: {1}; OLE inventory: {2}; mapping: {3}" -f $outputDir, $validatedRows.Count, $oleRecords.Count, $mappingPath)
