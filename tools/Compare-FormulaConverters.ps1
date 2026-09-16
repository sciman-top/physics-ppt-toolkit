<#
.SYNOPSIS
  Compare the controlled OMML candidate with a Pandoc reference conversion.

.DESCRIPTION
  This is a converter comparison harness, not a PPTX writer.  It consumes
  formula-omml-candidates.csv, sends each TargetTex through the vendored
  Pandoc markdown -> docx -> OMML reference path, and records structural
  evidence beside the repository's controlled OMML/MathML fragments.
  A Pandoc result is never copied into a slide and a failed reference does not
  make the controlled candidate successful.

.PARAMETER OmmlCandidateCsv
  CSV from Export-FormulaOmmlCandidates.ps1.

.PARAMETER OutputDir
  Directory for reference docx files and comparison reports.

.PARAMETER PandocPath
  Optional Pandoc executable. Defaults to the vendored repository copy.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OmmlCandidateCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$PandocPath = '',

    [ValidateRange(1, 1000)]
    [int]$MaxItems = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

function Get-RowValue {
    param([object]$Row, [string]$Name, [object]$Default = '')
    if ($null -eq $Row) { return $Default }
    $property = $Row.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Resolve-Pandoc {
    param([string]$Requested)
    if (-not [string]::IsNullOrWhiteSpace($Requested)) {
        $path = [System.IO.Path]::GetFullPath($Requested)
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "PandocPath not found: $path" }
        return $path
    }
    $vendored = Join-Path (Split-Path -Parent $PSScriptRoot) 'tools\vendor\pandoc\pandoc-3.9.0.2\pandoc.exe'
    if (Test-Path -LiteralPath $vendored -PathType Leaf) { return $vendored }
    $command = Get-Command pandoc -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $command) { return [string]$command.Source }
    return ''
}

function Invoke-PandocVersion {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return 'Unavailable' }
    $output = & $Path --version 2>&1
    if ($LASTEXITCODE -ne 0) { return 'ProbeFailed' }
    return (($output | Select-Object -First 1 | ForEach-Object { [string]$_ }) -join ' ').Trim()
}

function Get-XmlDigest {
    param([string]$XmlText)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($XmlText)
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    } finally { $sha.Dispose() }
}

function Get-PandocOmml {
    param([string]$Pandoc, [string]$Tex, [string]$WorkDir, [string]$Stem)
    if ([string]::IsNullOrWhiteSpace($Pandoc)) { return [pscustomobject]@{ Status = 'Unavailable'; Message = 'Pandoc executable is not available.'; Docx = ''; Xml = '' } }
    $mdPath = Join-Path $WorkDir ($Stem + '.md')
    $docxPath = Join-Path $WorkDir ($Stem + '.docx')
    # One backslash must reach Pandoc.  TargetTex is already a canonical TeX
    # value from the current whitelist/candidate exporter.
    [System.IO.File]::WriteAllText($mdPath, ('$$' + $Tex + '$$'), [System.Text.UTF8Encoding]::new($false))
    $output = & $Pandoc $mdPath -o $docxPath --from markdown --to docx 2>&1
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $docxPath -PathType Leaf)) {
        return [pscustomobject]@{ Status = 'Failed'; Message = (($output | ForEach-Object { [string]$_ }) -join ' ').Trim(); Docx = $docxPath; Xml = '' }
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = $null; $reader = $null; $stream = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
        $entry = $zip.GetEntry('word/document.xml')
        if ($null -eq $entry) { return [pscustomobject]@{ Status = 'Failed'; Message = 'Pandoc docx has no word/document.xml.'; Docx = $docxPath; Xml = '' } }
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        $xml = $reader.ReadToEnd()
        return [pscustomobject]@{ Status = if ($xml -match 'oMath') { 'Generated' } else { 'NoOfficeMath' }; Message = ''; Docx = $docxPath; Xml = $xml }
    } catch {
        return [pscustomobject]@{ Status = 'Failed'; Message = $_.Exception.Message; Docx = $docxPath; Xml = '' }
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

$csvPath = [System.IO.Path]::GetFullPath($OmmlCandidateCsv)
$outputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) { throw "OMML candidate CSV not found: $csvPath" }
if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
$referenceDir = Join-Path $outputDir 'pandoc-reference'
if (-not (Test-Path -LiteralPath $referenceDir -PathType Container)) { New-Item -ItemType Directory -Path $referenceDir -Force | Out-Null }
$pandoc = Resolve-Pandoc -Requested $PandocPath
$version = Invoke-PandocVersion -Path $pandoc
$rows = @(Import-Csv -LiteralPath $csvPath -Encoding UTF8 | Where-Object { [string](Get-RowValue -Row $_ -Name 'Status') -eq 'Generated' } | Select-Object -First $MaxItems)
$results = New-Object System.Collections.Generic.List[object]
for ($index = 0; $index -lt $rows.Count; $index++) {
    $row = $rows[$index]
    $directPath = [string](Get-RowValue -Row $row -Name 'OmmlFragment')
    $mathMlPath = [string](Get-RowValue -Row $row -Name 'MathMlFragment')
    $tex = [string](Get-RowValue -Row $row -Name 'TargetTex')
    $stem = 'reference-{0:0000}' -f ($index + 1)
    $reference = Get-PandocOmml -Pandoc $pandoc -Tex $tex -WorkDir $referenceDir -Stem $stem
    $directStatus = if ($directPath -and (Test-Path -LiteralPath $directPath -PathType Leaf)) { 'Present' } else { 'Missing' }
    $mathMlStatus = if ($mathMlPath -and (Test-Path -LiteralPath $mathMlPath -PathType Leaf)) { 'Present' } else { 'Missing' }
    $directXml = if ($directStatus -eq 'Present') { Get-Content -LiteralPath $directPath -Raw -Encoding UTF8 } else { '' }
    $directOmmlCount = if ($directXml) { ([regex]::Matches($directXml, '<m:oMath(?:\s|>)')).Count } else { 0 }
    $referenceOmmlCount = if ($reference.Xml) { ([regex]::Matches($reference.Xml, '<m:oMath(?:\s|>)')).Count } else { 0 }
    $directFractionCount = if ($directXml) { ([regex]::Matches($directXml, '<m:f(?:\s|>)')).Count } else { 0 }
    $referenceFractionCount = if ($reference.Xml) { ([regex]::Matches($reference.Xml, '<m:f(?:\s|>)')).Count } else { 0 }
    $comparisonStatus = if ($directStatus -ne 'Present') { 'DirectCandidateMissing' } elseif ($reference.Status -ne 'Generated') { 'ReferenceFailed' } elseif ($directOmmlCount -gt 0 -and $referenceOmmlCount -gt 0) { 'Compared' } else { 'StructureUnsupported' }
    $results.Add([pscustomobject]@{
        Index = $index + 1
        File = Get-RowValue -Row $row -Name 'File'
        Slide = Get-RowValue -Row $row -Name 'Slide'
        Shape = Get-RowValue -Row $row -Name 'Shape'
        Name = Get-RowValue -Row $row -Name 'Name'
        TargetTex = $tex
        DirectOmml = $directPath
        DirectMathMl = $mathMlPath
        DirectStatus = $directStatus
        DirectOmmlCount = $directOmmlCount
        DirectFractionCount = $directFractionCount
        PandocStatus = $reference.Status
        PandocMessage = $reference.Message
        PandocDocx = $reference.Docx
        PandocOmmlCount = $referenceOmmlCount
        PandocFractionCount = $referenceFractionCount
        DirectXmlSha256 = if ($directXml) { Get-XmlDigest -XmlText $directXml } else { '' }
        PandocXmlSha256 = if ($reference.Xml) { Get-XmlDigest -XmlText $reference.Xml } else { '' }
        ComparisonStatus = $comparisonStatus
    }) | Out-Null
}

$resultRows = @($results.ToArray())
$resultCsv = Join-Path $outputDir 'formula-converter-comparison.csv'
$resultJson = Join-Path $outputDir 'formula-converter-comparison.json'
$resultManifest = Join-Path $outputDir 'formula-converter-comparison-manifest.json'
Write-Utf8BomCsv -InputObject $resultRows -Path $resultCsv
Write-Utf8BomText -Path $resultJson -Text ($resultRows | ConvertTo-Json -Depth 10)
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    inputCsv = $csvPath
    pandoc = [ordered]@{ path = $pandoc; version = $version; role = 'ReferenceOnly'; writesPptx = $false }
    candidateCount = $resultRows.Count
    comparedCount = @($resultRows | Where-Object { $_.ComparisonStatus -eq 'Compared' }).Count
    referenceFailedCount = @($resultRows | Where-Object { $_.ComparisonStatus -eq 'ReferenceFailed' }).Count
    unsupportedCount = @($resultRows | Where-Object { $_.ComparisonStatus -eq 'StructureUnsupported' }).Count
    csv = $resultCsv
    json = $resultJson
    rows = $resultRows
}
Write-Utf8BomText -Path $resultManifest -Text ($manifest | ConvertTo-Json -Depth 12)
Write-Output ("Formula converter comparison done: {0}`nCandidates: {1}; Compared: {2}; ReferenceFailed: {3}" -f $outputDir, $manifest.candidateCount, $manifest.comparedCount, $manifest.referenceFailedCount)
