<#
.SYNOPSIS
  Reports whether the local toolkit dependencies are installed and callable.

.DESCRIPTION
  Checks the production, review, and optional helper tools used by this PPT
  workflow. The default run avoids opening PowerPoint. Use -LaunchPowerPoint
  when you need to verify live COM activation as well as registration.

.EXAMPLE
  .\tools\Assert-Toolchain.ps1

.EXAMPLE
  .\tools\Assert-Toolchain.ps1 -Deep -LaunchPowerPoint
#>
[CmdletBinding()]
param(
    [switch]$Deep,
    [switch]$LaunchPowerPoint,
    [switch]$RequireFormulaValidator,
    [switch]$RequireMediaOptimization
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$root = Split-Path -Parent $PSScriptRoot
$checks = New-Object System.Collections.Generic.List[object]
$nodeTier = if ($RequireMediaOptimization) { 'Required' } else { 'Recommended' }
$sharpTier = if ($RequireMediaOptimization) { 'Required' } else { 'Recommended' }
$dotNetTier = if ($RequireFormulaValidator) { 'Required' } else { 'Recommended' }

function Add-ToolchainCheck {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet('Required', 'Recommended', 'Optional')][string]$Tier,
        [Parameter(Mandatory = $true)][ValidateSet('OK', 'WARN', 'MISSING', 'FAIL', 'SKIP')][string]$Status,
        [string]$Version = '',
        [string]$Path = '',
        [string]$Details = ''
    )

    $order = switch ($Tier) {
        'Required' { 1 }
        'Recommended' { 2 }
        'Optional' { 3 }
    }

    $checks.Add([pscustomobject]@{
        Order = $order
        Tier = $Tier
        Name = $Name
        Status = $Status
        Version = $Version
        Path = $Path
        Details = $Details
    }) | Out-Null
}

function Resolve-CommandPath {
    param([Parameter(Mandatory = $true)][string]$Name)

    $cmd = Get-Command $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $cmd) { return '' }

    if (-not [string]::IsNullOrWhiteSpace([string]$cmd.Source)) { return [string]$cmd.Source }
    if ($cmd.PSObject.Properties['Path'] -and -not [string]::IsNullOrWhiteSpace([string]$cmd.Path)) {
        return [string]$cmd.Path
    }
    return [string]$cmd.Name
}

function Invoke-VersionProbe {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @('--version'),
        [int]$MaxLines = 1
    )

    try {
        # Lower EAP around the native call: on Windows PowerShell 5.1, stderr
        # from `2>&1` would otherwise become a terminating NativeCommandError
        # (or, inside this try/catch, a false probe failure) before
        # $LASTEXITCODE could be judged.
        $previousEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $output = & $FilePath @Arguments 2>&1
        } finally {
            $ErrorActionPreference = $previousEap
        }
        $text = ($output | Select-Object -First $MaxLines | ForEach-Object { [string]$_ }) -join ' | '
        return [pscustomobject]@{
            ExitCode = $LASTEXITCODE
            Text = $text.Trim()
        }
    } catch {
        return [pscustomobject]@{
            ExitCode = 999
            Text = $_.Exception.Message
        }
    }
}

# Version probes each spawn a full interpreter (npm alone can take seconds), so
# every startup paid 2-3s before any real work. Successful probes are cached in
# the user temp dir keyed by the executable's identity (path + size + UTC write
# time): upgrading or replacing a tool yields a new file identity and
# invalidates its own entry, so a cached hit reports exactly what a live probe
# would. Failed probes are never cached (transient breakage stays visible on
# the next run) and entries expire after 7 days to bound stale-hit risk.
$script:ToolchainProbeCachePath = Join-Path ([System.IO.Path]::GetTempPath()) 'physics-ppt-toolkit.toolchain-probe-cache.json'
$script:ToolchainProbeCacheRoot = $null
$script:ToolchainProbeCacheTtl = (New-TimeSpan -Days 7)

function Get-ToolchainProbeCacheRoot {
    if ($null -ne $script:ToolchainProbeCacheRoot) { return $script:ToolchainProbeCacheRoot }
    try {
        if (Test-Path -LiteralPath $script:ToolchainProbeCachePath) {
            $raw = [System.IO.File]::ReadAllText($script:ToolchainProbeCachePath)
            if (-not [string]::IsNullOrWhiteSpace($raw)) { $script:ToolchainProbeCacheRoot = $raw | ConvertFrom-Json }
        }
    } catch {
        $script:ToolchainProbeCacheRoot = $null
    }
    return $script:ToolchainProbeCacheRoot
}

function Get-ToolchainProbeIdentityKey {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @(),
        [int]$MaxLines = 0
    )
    try {
        $exe = Get-Item -LiteralPath $FilePath -ErrorAction Stop
    } catch {
        return $null
    }
    $argKey = if ($null -eq $Arguments) { '' } else { ($Arguments -join ' ') }
    return ('{0}|{1}|{2}|{3}|{4}' -f $exe.FullName.ToLowerInvariant(), $exe.Length, $exe.LastWriteTimeUtc.Ticks, $argKey, $MaxLines)
}

function Save-ToolchainProbeCacheEntry {
    param([string]$Key, $Payload)
    try {
        $root = Get-ToolchainProbeCacheRoot
        if ($null -eq $root) { $root = [pscustomobject]@{} }
        $payloadWithStamp = $Payload | Add-Member -MemberType NoteProperty -Name 'createdUtc' -Value ((Get-Date).ToUniversalTime().ToString('o')) -Force -PassThru
        $existing = $root.PSObject.Properties[$Key]
        if ($null -ne $existing) { $existing.Value = $payloadWithStamp } else { $root | Add-Member -MemberType NoteProperty -Name $Key -Value $payloadWithStamp }
        $script:ToolchainProbeCacheRoot = $root
        [System.IO.File]::WriteAllText($script:ToolchainProbeCachePath, ($root | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    } catch {
        # An unwritable or racing temp cache must never fail the toolchain check.
    }
}

function Get-ToolchainProbeCacheEntry {
    param([string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return $null }
    $root = Get-ToolchainProbeCacheRoot
    if ($null -eq $root) { return $null }
    $prop = $root.PSObject.Properties[$Key]
    if ($null -eq $prop) { return $null }
    try {
        $createdUtc = [datetime]::Parse([string]$prop.Value.createdUtc, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
        if (([datetime]::UtcNow - $createdUtc) -gt $script:ToolchainProbeCacheTtl) { return $null }
    } catch {
        return $null
    }
    return $prop.Value
}

function Invoke-VersionProbeCached {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @('--version'),
        [int]$MaxLines = 1
    )
    $key = Get-ToolchainProbeIdentityKey -FilePath $FilePath -Arguments $Arguments -MaxLines $MaxLines
    $entry = Get-ToolchainProbeCacheEntry -Key $key
    if ($null -ne $entry) {
        return [pscustomobject]@{ ExitCode = [int]$entry.exitCode; Text = [string]$entry.text }
    }
    $probe = Invoke-VersionProbe -FilePath $FilePath -Arguments $Arguments -MaxLines $MaxLines
    if ($probe.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($key)) {
        Save-ToolchainProbeCacheEntry -Key $key -Payload ([pscustomobject]@{ exitCode = [int]$probe.ExitCode; text = [string]$probe.Text })
    }
    return $probe
}

function Invoke-DotNetSdkProbeCached {
    param([Parameter(Mandatory = $true)][string]$DotNetPath)
    $key = Get-ToolchainProbeIdentityKey -FilePath $DotNetPath -Arguments @('--list-sdks')
    $entry = Get-ToolchainProbeCacheEntry -Key $key
    if ($null -ne $entry) {
        return [pscustomobject]@{ ExitCode = [int]$entry.exitCode; FirstSdk = [string]$entry.firstSdk }
    }

    # Lower EAP around the native call: on Windows PowerShell 5.1, stderr from
    # `2>&1` would otherwise become a terminating NativeCommandError before
    # $LASTEXITCODE could be judged (same guard as Invoke-VersionProbe).
    $previousEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $sdks = & $DotNetPath --list-sdks 2>&1
    } finally {
        $ErrorActionPreference = $previousEap
    }
    $sdkLines = @($sdks | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $exitCode = $LASTEXITCODE
    if ($exitCode -eq 0 -and $sdkLines.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($key)) {
        Save-ToolchainProbeCacheEntry -Key $key -Payload ([pscustomobject]@{ exitCode = [int]$exitCode; firstSdk = [string]$sdkLines[0] })
    }
    return [pscustomobject]@{ ExitCode = $exitCode; FirstSdk = if ($sdkLines.Count -gt 0) { [string]$sdkLines[0] } else { '' } }
}

function Invoke-NodeRepositoryProbe {
    param(
        [Parameter(Mandatory = $true)][string]$NodePath,
        [Parameter(Mandatory = $true)][string]$Script
    )

    Push-Location -LiteralPath $root
    try {
        # Same EAP guard as Invoke-VersionProbe: node/npm banners on stderr must
        # not terminate the probe before $LASTEXITCODE is read.
        $previousEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $output = & $NodePath -e $Script 2>&1
        } finally {
            $ErrorActionPreference = $previousEap
        }
        return [pscustomobject]@{
            ExitCode = $LASTEXITCODE
            Text = (($output | ForEach-Object { [string]$_ }) -join ' ').Trim()
        }
    } catch {
        return [pscustomobject]@{
            ExitCode = 999
            Text = $_.Exception.Message
        }
    } finally {
        Pop-Location
    }
}

function Test-NodePackage {
    param(
        [Parameter(Mandatory = $true)][string]$PackageName,
        [Parameter(Mandatory = $true)][string]$RelativePackageJson,
        [Parameter(Mandatory = $true)][string]$Tier
    )

    $packagePath = Join-Path $root $RelativePackageJson
    if (-not (Test-Path -LiteralPath $packagePath)) {
        Add-ToolchainCheck -Name $PackageName -Tier $Tier -Status 'MISSING' -Path $packagePath -Details 'node_modules package is missing; run npm install if package-lock.json is trusted.'
        return
    }

    try {
        $packageJson = Get-Content -LiteralPath $packagePath -Raw -Encoding UTF8 | ConvertFrom-Json
        Add-ToolchainCheck -Name $PackageName -Tier $Tier -Status 'OK' -Version ([string]$packageJson.version) -Path $packagePath
    } catch {
        Add-ToolchainCheck -Name $PackageName -Tier $Tier -Status 'FAIL' -Path $packagePath -Details $_.Exception.Message
    }
}

function Test-DotNetSdk {
    param([Parameter(Mandatory = $true)][string]$Tier)
    $candidates = New-Object System.Collections.Generic.List[string]
    $userDotnet = Join-Path $env:USERPROFILE '.dotnet\dotnet.exe'
    if (Test-Path -LiteralPath $userDotnet) { $candidates.Add($userDotnet) | Out-Null }

    $pathDotnet = Resolve-CommandPath 'dotnet'
    if (-not [string]::IsNullOrWhiteSpace($pathDotnet) -and $pathDotnet -notin $candidates) {
        $candidates.Add($pathDotnet) | Out-Null
    }

    if ($candidates.Count -eq 0) {
        Add-ToolchainCheck -Name '.NET SDK' -Tier $Tier -Status 'MISSING' -Details 'FormulaOfficeMathValidator requires a dotnet SDK.'
        return
    }

    foreach ($candidate in $candidates) {
        $sdkProbe = Invoke-DotNetSdkProbeCached -DotNetPath $candidate
        if ($sdkProbe.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($sdkProbe.FirstSdk)) {
            Add-ToolchainCheck -Name '.NET SDK' -Tier $Tier -Status 'OK' -Version $sdkProbe.FirstSdk -Path $candidate
            return
        }
    }

    Add-ToolchainCheck -Name '.NET SDK' -Tier $Tier -Status 'FAIL' -Path ($candidates -join '; ') -Details 'dotnet exists, but no SDK is available.'
}

function Test-PowerPointCom {
    $type = [type]::GetTypeFromProgID('PowerPoint.Application')
    if ($null -eq $type) {
        Add-ToolchainCheck -Name 'PowerPoint COM registration' -Tier 'Required' -Status 'MISSING' -Details 'PowerPoint.Application ProgID is not registered.'
        return
    }

    Add-ToolchainCheck -Name 'PowerPoint COM registration' -Tier 'Required' -Status 'OK' -Details 'PowerPoint.Application ProgID is registered.'

    if (-not $LaunchPowerPoint) {
        Add-ToolchainCheck -Name 'PowerPoint COM live activation' -Tier 'Recommended' -Status 'SKIP' -Details 'Use -LaunchPowerPoint to start PowerPoint and read Application.Version.'
        return
    }

    $existing = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
    $pp = $null
    try {
        $pp = New-PowerPointApplication
        $version = [string]$pp.Version
        Add-ToolchainCheck -Name 'PowerPoint COM live activation' -Tier 'Recommended' -Status 'OK' -Version $version
    } catch {
        Add-ToolchainCheck -Name 'PowerPoint COM live activation' -Tier 'Recommended' -Status 'FAIL' -Details $_.Exception.Message
    } finally {
        if ($null -ne $pp) {
            try {
                if ($existing.Count -eq 0) { $pp.Quit() | Out-Null }
            } catch {
                # Release below is still useful even if PowerPoint refuses Quit.
            }
            Release-ComObjectSafe -ComObject $pp
        }
    }
}

function Test-VendoredExecutable {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][string]$Tier,
        [string[]]$VersionArguments = @(),
        [int]$MaxLines = 1
    )

    $path = Join-Path $root $RelativePath
    if (-not (Test-Path -LiteralPath $path)) {
        Add-ToolchainCheck -Name $Name -Tier $Tier -Status 'MISSING' -Path $path
        return
    }

    if ($VersionArguments.Count -eq 0) {
        Add-ToolchainCheck -Name $Name -Tier $Tier -Status 'OK' -Path $path
        return
    }

    $version = Invoke-VersionProbeCached -FilePath $path -Arguments $VersionArguments -MaxLines $MaxLines
    # A version-looking banner from a broken executable or a WindowsApps
    # placeholder is not proof that the tool is callable. Exit code is the
    # authoritative probe result.
    if ($version.ExitCode -eq 0) {
        Add-ToolchainCheck -Name $Name -Tier $Tier -Status 'OK' -Version $version.Text -Path $path
    } else {
        Add-ToolchainCheck -Name $Name -Tier $Tier -Status 'FAIL' -Path $path -Details 'Version probe failed.'
    }
}

# Required local runtime checks.  PowerShell 7 is the primary host; Windows
# PowerShell 5.1 remains a compatibility fallback for legacy callers.
$currentPsVersion = $PSVersionTable.PSVersion.ToString()
$currentPsIsSeven = ($PSVersionTable.PSEdition -eq 'Core' -and $PSVersionTable.PSVersion.Major -ge 7)
$currentPsStatus = if ($currentPsIsSeven) { 'OK' } else { 'WARN' }
$currentPsDetails = if ($currentPsIsSeven) { '' } else { 'Run the default entrypoints with pwsh; this invocation is using the legacy host.' }
Add-ToolchainCheck -Name 'Current PowerShell runtime' -Tier 'Recommended' -Status $currentPsStatus -Version $currentPsVersion -Path $PSHOME -Details $currentPsDetails

$primaryPowerShellPath = Resolve-CommandPath 'pwsh'
if ([string]::IsNullOrWhiteSpace($primaryPowerShellPath)) {
    Add-ToolchainCheck -Name 'PowerShell 7 (primary host)' -Tier 'Required' -Status 'MISSING' -Details 'The default entrypoints require pwsh. Windows PowerShell 5.1 is supported only as a compatibility fallback.'
} else {
    $primaryVersion = Invoke-VersionProbeCached -FilePath $primaryPowerShellPath -Arguments @('-NoLogo', '-NoProfile', '-NonInteractive', '-Command', '$PSVersionTable.PSVersion.ToString()')
    $primaryStatus = if ($primaryVersion.ExitCode -eq 0) { 'OK' } else { 'FAIL' }
    Add-ToolchainCheck -Name 'PowerShell 7 (primary host)' -Tier 'Required' -Status $primaryStatus -Version $primaryVersion.Text -Path $primaryPowerShellPath
}

$legacyPowerShellPath = Resolve-CommandPath 'powershell.exe'
if ([string]::IsNullOrWhiteSpace($legacyPowerShellPath)) {
    Add-ToolchainCheck -Name 'Windows PowerShell 5.1 (compatibility fallback)' -Tier 'Optional' -Status 'MISSING' -Details 'Not required when PowerShell 7 is available.'
} else {
    $legacyVersion = Invoke-VersionProbeCached -FilePath $legacyPowerShellPath -Arguments @('-NoLogo', '-NoProfile', '-NonInteractive', '-Command', '$PSVersionTable.PSVersion.ToString()')
    $legacyStatus = if ($legacyVersion.ExitCode -eq 0) { 'OK' } else { 'WARN' }
    Add-ToolchainCheck -Name 'Windows PowerShell 5.1 (compatibility fallback)' -Tier 'Optional' -Status $legacyStatus -Version $legacyVersion.Text -Path $legacyPowerShellPath
}

Test-PowerPointCom

$nodePath = Resolve-CommandPath 'node'
if ([string]::IsNullOrWhiteSpace($nodePath)) {
    Add-ToolchainCheck -Name 'Node.js' -Tier $nodeTier -Status 'MISSING' -Details 'sharp/libvips media optimization requires Node.js.'
} else {
    $nodeVersion = Invoke-VersionProbeCached -FilePath $nodePath -Arguments @('--version')
    $nodeStatus = if ($nodeVersion.ExitCode -eq 0) { 'OK' } else { 'FAIL' }
    Add-ToolchainCheck -Name 'Node.js' -Tier $nodeTier -Status $nodeStatus -Version $nodeVersion.Text -Path $nodePath
}

$npmPath = Resolve-CommandPath 'npm'
if ([string]::IsNullOrWhiteSpace($npmPath)) {
    Add-ToolchainCheck -Name 'npm' -Tier 'Recommended' -Status 'MISSING' -Details 'Needed only when restoring node_modules.'
} else {
    $npmVersion = Invoke-VersionProbeCached -FilePath $npmPath -Arguments @('--version')
    $npmStatus = if ($npmVersion.ExitCode -eq 0) { 'OK' } else { 'FAIL' }
    Add-ToolchainCheck -Name 'npm' -Tier 'Recommended' -Status $npmStatus -Version $npmVersion.Text -Path $npmPath
}

Test-NodePackage -PackageName 'sharp' -RelativePackageJson 'node_modules\sharp\package.json' -Tier $sharpTier

$runNodeSmoke = $Deep -or $RequireMediaOptimization
if ($runNodeSmoke -and -not [string]::IsNullOrWhiteSpace($nodePath)) {
    # sharp does not export package.json in current releases; requiring the
    # module itself is the callable probe. Keep the version optional so an
    # exports-map change cannot turn a healthy install into a false failure.
    $sharpProbe = Invoke-NodeRepositoryProbe -NodePath $nodePath -Script "const sharp=require('sharp'); process.stdout.write(String(sharp.versions?.sharp || 'loaded'))"
    if ($sharpProbe.ExitCode -eq 0) {
        Add-ToolchainCheck -Name 'sharp require call' -Tier $sharpTier -Status 'OK' -Version $sharpProbe.Text
    } else {
        Add-ToolchainCheck -Name 'sharp require call' -Tier $sharpTier -Status 'FAIL' -Details $sharpProbe.Text
    }
}

Test-DotNetSdk -Tier $dotNetTier

# Recommended and optional portable tools.
Test-VendoredExecutable -Name 'oxipng portable' -Tier 'Recommended' -RelativePath 'tools\vendor\oxipng-10.1.1\oxipng-10.1.1-x86_64-pc-windows-msvc\oxipng.exe' -VersionArguments @('--version')
Test-VendoredExecutable -Name 'Real-ESRGAN ncnn Vulkan portable' -Tier 'Recommended' -RelativePath 'tools\vendor\realesrgan-ncnn-vulkan-20220424\realesrgan-ncnn-vulkan.exe'
Test-VendoredExecutable -Name 'Pandoc portable' -Tier 'Optional' -RelativePath 'tools\vendor\pandoc\pandoc-3.9.0.2\pandoc.exe' -VersionArguments @('--version') -MaxLines 1

foreach ($tool in @('magick', 'ffmpeg', 'pngquant', 'cjpeg', 'jpegtran')) {
    $path = Resolve-CommandPath $tool
    if ([string]::IsNullOrWhiteSpace($path)) {
        Add-ToolchainCheck -Name "optional command $tool" -Tier 'Optional' -Status 'MISSING'
    } else {
        Add-ToolchainCheck -Name "optional command $tool" -Tier 'Optional' -Status 'OK' -Path $path
    }
}

$ordered = @($checks | Sort-Object Order, Name)
$requiredFailures = @($ordered | Where-Object { $_.Tier -eq 'Required' -and $_.Status -in @('MISSING', 'FAIL') })

$ordered | Select-Object Tier, Name, Status, Version, Path, Details | Format-Table -AutoSize
Write-Host ("Required failures: {0}" -f $requiredFailures.Count)

if ($requiredFailures.Count -gt 0) {
    throw "Toolchain check failed: required=$($requiredFailures.Count)."
}
