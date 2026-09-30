<#
.SYNOPSIS
  Verify FormulaRecognitionAdapter status and write-back safety contracts.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('physics-formula-adapter-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $imagePath = Join-Path $testRoot 'fixture.png'
    [IO.File]::WriteAllBytes($imagePath, [byte[]](137,80,78,71,13,10,26,10))
    $tool = Join-Path $PSScriptRoot 'Run-FormulaRecognitionAdapter.ps1'
    $fixtures = Join-Path (Split-Path -Parent $PSScriptRoot) 'examples\fixtures'
    $pwshPath = (Get-Command pwsh -ErrorAction Stop).Source

    $cases = @(
        [ordered]@{ name = 'missing-runner'; runner = (Join-Path $testRoot 'missing-runner.exe'); arguments = @(); timeout = 1; expected = 'Unavailable' }
        [ordered]@{ name = 'valid-runner'; runner = $pwshPath; arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', (Join-Path $fixtures 'formula-adapter-valid-runner.ps1')); timeout = 5; expected = 'Passed' }
        [ordered]@{ name = 'invalid-output'; runner = $pwshPath; arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', (Join-Path $fixtures 'formula-adapter-invalid-runner.ps1')); timeout = 5; expected = 'InvalidOutput' }
        [ordered]@{ name = 'timeout'; runner = $pwshPath; arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', (Join-Path $fixtures 'formula-adapter-timeout-runner.ps1')); timeout = 1; expected = 'Timeout' }
    )
    foreach ($case in $cases) {
        $output = Join-Path $testRoot ($case.name + '.json')
        $invokeArgs = @{
            Adapter = 'pix2tex'
            ImagePath = $imagePath
            OutputPath = $output
            RunnerPath = [string]$case.runner
            TimeoutSeconds = [int]$case.timeout
        }
        if (@($case.arguments).Count -gt 0) { $invokeArgs.RunnerArgumentList = @($case.arguments) }
        & $tool @invokeArgs
        if (-not $?) { throw "Adapter case '$($case.name)' failed to run." }
        $result = Get-Content -LiteralPath $output -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]$result.status -ne $case.expected) { throw "Adapter case '$($case.name)' returned '$($result.status)', expected '$($case.expected)'." }
        if ([bool]$result.writeBackAllowed) { throw "Adapter case '$($case.name)' enabled write-back." }
        if ([string]$result.input.sha256 -notmatch '^[a-f0-9]{64}$') { throw "Adapter case '$($case.name)' lost input hash." }
        if ([string]::IsNullOrWhiteSpace([string]$result.diagnostics.reason)) { throw "Adapter case '$($case.name)' has no diagnostic reason." }
    }
    Write-Host 'Formula recognition adapter tests passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
