<#
.SYNOPSIS
  Run or safely probe a formula-recognition adapter.

.DESCRIPTION
  This is a transport boundary, not a content writer. Without an explicitly
  supplied trusted runner it returns Unavailable, so missing model packages do
  not become implicit downloads or false recognition evidence.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('PP-FormulaNet_plus-M', 'PP-FormulaNet_plus-L', 'UniMERNet', 'pix2tex')]
    [string]$Adapter,

    [Parameter(Mandatory = $true)]
    [string]$ImagePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$RunnerPath,

    [string[]]$RunnerArgumentList = @(),

    [string]$ModelPath,

    [ValidateRange(1, 3600)]
    [int]$TimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-Sha256FileLocal { param([string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Write-JsonUtf8 { param([string]$Path, $Value) $dir = Split-Path -Parent $Path; if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }; $Value | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $Path -Encoding UTF8 }

$ImagePath = [IO.Path]::GetFullPath($ImagePath)
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $ImagePath -PathType Leaf)) { throw "Formula image not found: $ImagePath" }
$inputFile = Get-Item -LiteralPath $ImagePath
$started = [Diagnostics.Stopwatch]::StartNew()
$result = [ordered]@{
    schemaVersion = 1
    adapter = $Adapter
    status = 'Unavailable'
    writeBackAllowed = $false
    input = [ordered]@{ path = $ImagePath; sha256 = Get-Sha256FileLocal -Path $ImagePath; bytes = [int64]$inputFile.Length }
    candidates = @()
    diagnostics = [ordered]@{ runtime = 'PowerShell transport boundary'; reason = ''; elapsedMs = $null; modelPath = if ([string]::IsNullOrWhiteSpace($ModelPath)) { $null } else { [IO.Path]::GetFullPath($ModelPath) }; runnerPath = if ([string]::IsNullOrWhiteSpace($RunnerPath)) { $null } else { [IO.Path]::GetFullPath($RunnerPath) } }
}
try {
    if ([string]::IsNullOrWhiteSpace($RunnerPath)) {
        $result.diagnostics.reason = "Adapter '$Adapter' is not locally provisioned; no runner was supplied."
    } elseif (-not (Test-Path -LiteralPath $RunnerPath -PathType Leaf)) {
        $result.status = 'Unavailable'
        $result.diagnostics.reason = "Configured runner does not exist: $RunnerPath"
    } elseif (-not [string]::IsNullOrWhiteSpace($ModelPath) -and -not (Test-Path -LiteralPath $ModelPath -PathType Leaf)) {
        $result.status = 'Unavailable'
        $result.diagnostics.reason = "Configured model does not exist: $ModelPath"
    } else {
        $psi = New-Object Diagnostics.ProcessStartInfo
        $psi.FileName = [IO.Path]::GetFullPath($RunnerPath)
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardInput = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        # ProcessStartInfo.ArgumentList is .NET Core only; build a quoted
        # argument string so the Windows PowerShell 5.1 fallback still works.
        $argumentTexts = @($RunnerArgumentList) + @()
        $psi.Arguments = ($argumentTexts | ForEach-Object { '"' + ([string]$_).Replace('"', '\"') + '"' }) -join ' '
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $psi
        if (-not $process.Start()) { throw "Could not start adapter runner: $RunnerPath" }
        $request = [ordered]@{
            schemaVersion = 1
            adapter = $Adapter
            imagePath = $ImagePath
            imageSha256 = [string]$result.input.sha256
            modelPath = if ([string]::IsNullOrWhiteSpace($ModelPath)) { $null } else { [IO.Path]::GetFullPath($ModelPath) }
        } | ConvertTo-Json -Depth 8 -Compress
        $process.StandardInput.Write($request)
        $process.StandardInput.Close()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            try { $process.Kill($true) } catch { try { $process.Kill() } catch {} }
            $result.status = 'Timeout'
            $result.diagnostics.reason = "Adapter runner exceeded ${TimeoutSeconds}s."
        } else {
            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()
            if ($process.ExitCode -ne 0) {
                $result.status = 'Failed'
                $result.diagnostics.reason = "Adapter runner exited $($process.ExitCode): $stderr"
            } else {
                try {
                    $runnerResult = $stdout | ConvertFrom-Json
                    if ($null -eq $runnerResult -or $null -eq $runnerResult.PSObject.Properties['status'] -or $null -eq $runnerResult.PSObject.Properties['candidates']) {
                        throw 'Runner output must contain status and candidates.'
                    }
                    $runnerWriteBack = $runnerResult.PSObject.Properties['writeBackAllowed']
                    if ($null -ne $runnerWriteBack -and [bool]$runnerWriteBack.Value) { throw 'Runner attempted to enable write-back.' }
                    $allowed = @('Passed', 'Unavailable', 'Failed', 'InvalidOutput', 'Timeout')
                    if ([string]$runnerResult.status -notin $allowed) { throw "Runner returned unsupported status: $($runnerResult.status)" }
                    $result.status = [string]$runnerResult.status
                    $result.candidates = @($runnerResult.candidates)
                    $result.diagnostics.reason = if ($null -ne $runnerResult.diagnostics -and -not [string]::IsNullOrWhiteSpace([string]$runnerResult.diagnostics.reason)) { [string]$runnerResult.diagnostics.reason } else { 'Runner completed.' }
                } catch {
                    $result.status = 'InvalidOutput'
                    $result.diagnostics.reason = "Runner returned invalid JSON contract: $($_.Exception.Message)"
                }
            }
        }
        $process.Dispose()
    }
} catch {
    $result.status = 'Failed'
    $result.diagnostics.reason = $_.Exception.Message
} finally {
    $started.Stop()
    $result.diagnostics.elapsedMs = [int][Math]::Max(0, $started.Elapsed.TotalMilliseconds)
    Write-JsonUtf8 -Path $OutputPath -Value $result
}
Write-Host "Formula recognition adapter result: $OutputPath ($($result.status))"
