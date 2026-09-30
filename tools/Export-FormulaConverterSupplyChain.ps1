<#
.SYNOPSIS
  Produce a versioned dependency and license receipt for formula converters.

.DESCRIPTION
  The manifest is evidence only. It does not install packages and explicitly
  distinguishes adopted dependencies from unreviewed candidate projects.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$PandocPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tools\vendor\pandoc\pandoc-3.9.0.2\pandoc.exe')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-Sha256FileLocal { param([string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }

$root = Split-Path -Parent $PSScriptRoot
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$packageLockPath = Join-Path $root 'package-lock.json'
$packageJsonPath = Join-Path $root 'package.json'
$mathJaxPackagePath = Join-Path $root 'node_modules\@mathjax\src\package.json'
$mathJaxLicensePath = Join-Path $root 'node_modules\@mathjax\src\LICENSE'
$tokenizerPath = Join-Path $root 'tools\Export-FormulaOmmlCandidates.ps1'
$pandocCopyrightPath = Join-Path (Split-Path -Parent $PandocPath) 'COPYRIGHT.txt'
foreach ($path in @($packageLockPath, $packageJsonPath, $mathJaxPackagePath, $mathJaxLicensePath, $tokenizerPath, $PandocPath, $pandocCopyrightPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Supply-chain evidence is missing: $path" }
}
$mathJaxPackage = Get-Content -LiteralPath $mathJaxPackagePath -Raw -Encoding UTF8 | ConvertFrom-Json
$pandocVersion = ((& $PandocPath --version | Select-Object -First 1) -replace '^pandoc\s+', '').Trim()
if ([string]::IsNullOrWhiteSpace($pandocVersion)) { throw 'Could not read Pandoc version.' }
$manifest = [ordered]@{
    schemaVersion = 1
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    policy = [ordered]@{ writesPptx = $false; defaultConverter = 'InternalTokenOmml'; referenceConverter = 'PandocReferenceOnly' }
    lockEvidence = [ordered]@{ packageJson = $packageJsonPath; packageJsonSha256 = Get-Sha256FileLocal -Path $packageJsonPath; packageLock = $packageLockPath; packageLockSha256 = Get-Sha256FileLocal -Path $packageLockPath }
    components = @(
        [ordered]@{ id = 'InternalTokenOmml'; status = 'Adopted'; role = 'Controlled FormulaIR to OMML candidate generation'; version = 'repository-source'; sha256 = Get-Sha256FileLocal -Path $tokenizerPath; license = 'Repository license'; licenseEvidence = 'Repository source control and project license'; source = $tokenizerPath; consumer = 'Export-FormulaOmmlCandidates.ps1'; removal = 'Remove only after replacement converter passes FormulaIR, validator, PowerPoint, and visual gates.' }
        [ordered]@{ id = 'MathJax'; status = 'FallbackOnly'; role = 'Canonical TeX to SVG/MathML fallback rendering'; version = [string]$mathJaxPackage.version; sha256 = Get-Sha256FileLocal -Path $mathJaxPackagePath; license = [string]$mathJaxPackage.license; licenseEvidence = $mathJaxLicensePath; source = 'package-lock.json node_modules/@mathjax/src'; consumer = 'Render-FormulaSvg.mjs'; removal = 'npm remove @mathjax/src and remove the SVG fallback path only after a tested replacement exists.' }
        [ordered]@{ id = 'Pandoc'; status = 'ReferenceOnly'; role = 'LaTeX to DOCX/OMML comparison reference'; version = $pandocVersion; sha256 = Get-Sha256FileLocal -Path $PandocPath; license = 'GPL-2.0-or-later'; licenseEvidence = $pandocCopyrightPath; source = $PandocPath; consumer = 'Compare-FormulaConverters.ps1'; removal = 'Delete tools/vendor/pandoc/pandoc-3.9.0.2 only after removing the comparison harness or selecting another verified reference.' }
        [ordered]@{ id = 'mathml2omml'; status = 'NotAdopted'; role = 'Potential MathML to OMML adapter'; version = 'not-installed'; sha256 = $null; license = 'Unverified for this checkout'; licenseEvidence = 'No package installation or lock entry'; source = 'https://github.com/fiduswriter/mathml2omml'; consumer = 'None'; removal = 'No removal action; do not install without license, lock, and host validation.' }
        [ordered]@{ id = 'node-latex-to-omml'; status = 'NotAdopted'; role = 'Potential LaTeX to OMML adapter'; version = 'not-installed'; sha256 = $null; license = 'Unverified for this checkout'; licenseEvidence = 'No package installation or lock entry'; source = 'https://github.com/JaredYe04/node-latex-to-omml'; consumer = 'None'; removal = 'No removal action; do not install without license, lock, and host validation.' }
    )
}
$directory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Host "Formula converter supply-chain receipt written: $OutputPath"
