<#
.SYNOPSIS
  Verify that an unexplained recognition false acceptance closes the unattended circuit breaker.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-TestJson {
    param([string]$Path, [object]$Value)
    $Value | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $Path -Encoding UTF8
}

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('physics-closed-world-circuit-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $sourceHash = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
    $recordHash = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
    $inventoryPath = Join-Path $testRoot 'inventory.json'
    $goldSetPath = Join-Path $testRoot 'goldset.json'
    $contextPath = Join-Path $testRoot 'context.json'
    $recognitionPath = Join-Path $testRoot 'recognition.json'
    $outputDir = Join-Path $testRoot 'output'

    Write-TestJson -Path $inventoryPath -Value ([ordered]@{
        schemaVersion = 1
        input = [ordered]@{ sha256 = $sourceHash }
        records = @([ordered]@{
            recordId = 's1-sh1-OfficeMath'
            source = [ordered]@{ carrier = 'OfficeMath'; slide = 1; shapeId = 1; sourceSha256 = $recordHash }
        })
    })
    Write-TestJson -Path $goldSetPath -Value ([ordered]@{
        schemaVersion = 1
        input = [ordered]@{ sourcePptxSha256 = $sourceHash }
        policy = [ordered]@{ writeBackAllowed = $false }
        records = @()
    })
    $goldSetHash = (Get-FileHash -LiteralPath $goldSetPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Write-TestJson -Path $contextPath -Value ([ordered]@{
        schemaVersion = 1
        policy = [ordered]@{ writeBackAllowed = $false }
        results = @()
    })
    Write-TestJson -Path $recognitionPath -Value ([ordered]@{
        schemaVersion = 1
        inputs = [ordered]@{ goldSetManifestSha256 = $goldSetHash }
        policy = [ordered]@{ writeBackAllowed = $false; releaseDecision = 'CandidateOnlyNoAvailableAdapter' }
        adapters = @([ordered]@{ adapter = 'Fixture'; falseAcceptCount = 1; decision = 'FalseAcceptObserved' })
    })

    & (Join-Path $PSScriptRoot 'Plan-ClosedWorldUnattended.ps1') `
        -CarrierInventoryJson $inventoryPath `
        -GoldSetManifestJson $goldSetPath `
        -ContextResolutionJson $contextPath `
        -RecognitionEvaluationJson $recognitionPath `
        -OutputDir $outputDir | Out-Null
    if (-not $?) { throw 'Closed-world planner failed during circuit-breaker fixture.' }

    $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'closed-world-unattended-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]$manifest.circuitBreaker.state -ne 'Closed' -or -not [bool]$manifest.circuitBreaker.triggered) {
        throw 'Unexplained false acceptance did not close the circuit breaker.'
    }
    if ([string]$manifest.circuitBreaker.reason -notmatch '(?i)false acceptance') {
        throw 'Circuit-breaker reason does not record the false-acceptance cause.'
    }
    if ([int]$manifest.counts.Converted -ne 0 -or @($manifest.results | Where-Object { $_.decision.writeBackAllowed }).Count -ne 0) {
        throw 'Circuit-breaker fixture produced an unauthorized conversion or write-back.'
    }
    Write-Host 'Closed-world false-acceptance circuit-breaker test passed.'
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
