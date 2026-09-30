<#
.SYNOPSIS
  Prove that FormulaOfficeMathValidator rejects a malformed PPTX copy.

.DESCRIPTION
  The test corrupts only a temporary copy of the minimal fixture. It verifies
  both the validator's machine-readable failure contract and preservation of
  the input fixture. It does not start PowerPoint and does not write a deck.
#>
[CmdletBinding()]
param(
    [string]$FixturePptx = (Join-Path (Split-Path -Parent $PSScriptRoot) 'examples\fixtures\minimal-physics-sample.pptx')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$FixturePptx = [IO.Path]::GetFullPath($FixturePptx)
if (-not (Test-Path -LiteralPath $FixturePptx -PathType Leaf)) { throw "Fixture PPTX not found: $FixturePptx" }
$sourceHash = (Get-FileHash -LiteralPath $FixturePptx -Algorithm SHA256).Hash.ToLowerInvariant()
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root 'tools\FormulaOfficeMathValidator\FormulaOfficeMathValidator.csproj'
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) { throw 'dotnet is required to run the OfficeMath validator fault-injection test.' }

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('formula-omml-validator-fault-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $corruptPptx = Join-Path $testRoot 'malformed-slide.pptx'
    $resultJson = Join-Path $testRoot 'validator-result.json'
    Copy-Item -LiteralPath $FixturePptx -Destination $corruptPptx -Force

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($corruptPptx, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $archive.GetEntry('ppt/slides/slide1.xml')
        if ($null -eq $entry) { throw 'Fault fixture lacks ppt/slides/slide1.xml.' }
        $entry.Delete()
        $replacement = $archive.CreateEntry('ppt/slides/slide1.xml')
        $writer = New-Object IO.StreamWriter($replacement.Open(), (New-Object Text.UTF8Encoding($false)))
        try { $writer.Write('<p:sld') } finally { $writer.Dispose() }
    } finally {
        $archive.Dispose()
    }

    & dotnet run --project $project -- $corruptPptx --json $resultJson
    $validatorExit = $LASTEXITCODE
    if ($validatorExit -eq 0) { throw 'Fault-injected malformed PPTX was accepted by FormulaOfficeMathValidator.' }
    if (-not (Test-Path -LiteralPath $resultJson -PathType Leaf)) { throw 'Validator did not emit its JSON failure receipt.' }
    $result = Get-Content -LiteralPath $resultJson -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int]$result.OpenXmlErrorCount -lt 1) { throw 'Validator failure receipt does not record an Open XML/package error.' }
    $firstDescription = if (@($result.OpenXmlErrors).Count -gt 0) { [string]$result.OpenXmlErrors[0].Description } else { '' }
    if ([string]::IsNullOrWhiteSpace($firstDescription)) { throw 'Validator failure receipt has no error description.' }
    $sourceHashAfter = (Get-FileHash -LiteralPath $FixturePptx -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($sourceHashAfter -ne $sourceHash) { throw 'Validator fault injection modified the source fixture.' }
    Write-Host "FormulaOfficeMathValidator fault injection passed: exit=$validatorExit; errors=$($result.OpenXmlErrorCount)"
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
