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
    [switch]$RequireFormulaSvg,
    [switch]$RequireMediaOptimization,
    [switch]$Strict,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PhysicsPpt.Common.ps1')

$root = Split-Path -Parent $PSScriptRoot
$checks = New-Object System.Collections.Generic.List[object]
$nodeTier = if ($RequireFormulaSvg -or $RequireMediaOptimization) { 'Required' } else { 'Recommended' }
$mathJaxTier = if ($RequireFormulaSvg) { 'Required' } else { 'Recommended' }
$sharpTier = if ($RequireMediaOptimization) { 'Required' } else { 'Recommended' }
$dotNetTier = if ($RequireFormulaValidator) { 'Required' } else { 'Recommended' }

function Add-ToolchainCheck {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet('Required', 'Recommended', 'Optional', 'Experimental')][string]$Tier,
        [Parameter(Mandatory = $true)][ValidateSet('OK', 'WARN', 'MISSING', 'FAIL', 'SKIP')][string]$Status,
        [string]$Version = '',
        [string]$Path = '',
        [string]$Details = ''
    )

    $order = switch ($Tier) {
        'Required' { 1 }
        'Recommended' { 2 }
        'Optional' { 3 }
        'Experimental' { 4 }
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
        $output = & $FilePath @Arguments 2>&1
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

function Invoke-NodeRepositoryProbe {
    param(
        [Parameter(Mandatory = $true)][string]$NodePath,
        [Parameter(Mandatory = $true)][string]$Script
    )

    Push-Location -LiteralPath $root
    try {
        $output = & $NodePath -e $Script 2>&1
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
        return ''
    }

    try {
        $packageJson = Get-Content -LiteralPath $packagePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $version = [string]$packageJson.version
        Add-ToolchainCheck -Name $PackageName -Tier $Tier -Status 'OK' -Version $version -Path $packagePath
        return $version
    } catch {
        Add-ToolchainCheck -Name $PackageName -Tier $Tier -Status 'FAIL' -Path $packagePath -Details $_.Exception.Message
        return ''
    }
}

function Test-PythonModule {
    param(
        [Parameter(Mandatory = $true)][string]$PythonExe,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $true)][string]$Tier
    )

    $code = "import importlib, sys; mod=importlib.import_module(sys.argv[1]); print(getattr(mod, '__version__', 'OK'))"
    try {
        $output = & $PythonExe -c $code $ModuleName 2>&1
        if ($LASTEXITCODE -eq 0) {
            $version = (($output | Select-Object -First 1) -as [string]).Trim()
            Add-ToolchainCheck -Name "Python module $ModuleName" -Tier $Tier -Status 'OK' -Version $version -Path $PythonExe
        } else {
            $details = (($output | ForEach-Object { [string]$_ }) -join ' ').Trim()
            Add-ToolchainCheck -Name "Python module $ModuleName" -Tier $Tier -Status 'MISSING' -Path $PythonExe -Details $details
        }
    } catch {
        $details = $_.Exception.Message
        if ($_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
            $details = $_.ErrorDetails.Message
        }
        Add-ToolchainCheck -Name "Python module $ModuleName" -Tier $Tier -Status 'MISSING' -Path $PythonExe -Details $details
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
        $sdks = & $candidate --list-sdks 2>&1
        $sdkLines = @($sdks | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($LASTEXITCODE -eq 0 -and $sdkLines.Count -gt 0) {
            Add-ToolchainCheck -Name '.NET SDK' -Tier $Tier -Status 'OK' -Version ([string]$sdkLines[0]) -Path $candidate
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

    $version = Invoke-VersionProbe -FilePath $path -Arguments $VersionArguments -MaxLines $MaxLines
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
    $primaryVersion = Invoke-VersionProbe -FilePath $primaryPowerShellPath -Arguments @('-NoLogo', '-NoProfile', '-NonInteractive', '-Command', '$PSVersionTable.PSVersion.ToString()')
    $primaryStatus = if ($primaryVersion.ExitCode -eq 0) { 'OK' } else { 'FAIL' }
    Add-ToolchainCheck -Name 'PowerShell 7 (primary host)' -Tier 'Required' -Status $primaryStatus -Version $primaryVersion.Text -Path $primaryPowerShellPath
}

$legacyPowerShellPath = Resolve-CommandPath 'powershell.exe'
if ([string]::IsNullOrWhiteSpace($legacyPowerShellPath)) {
    Add-ToolchainCheck -Name 'Windows PowerShell 5.1 (compatibility fallback)' -Tier 'Optional' -Status 'MISSING' -Details 'Not required when PowerShell 7 is available.'
} else {
    $legacyVersion = Invoke-VersionProbe -FilePath $legacyPowerShellPath -Arguments @('-NoLogo', '-NoProfile', '-NonInteractive', '-Command', '$PSVersionTable.PSVersion.ToString()')
    $legacyStatus = if ($legacyVersion.ExitCode -eq 0) { 'OK' } else { 'WARN' }
    Add-ToolchainCheck -Name 'Windows PowerShell 5.1 (compatibility fallback)' -Tier 'Optional' -Status $legacyStatus -Version $legacyVersion.Text -Path $legacyPowerShellPath
}

Test-PowerPointCom

$nodePath = Resolve-CommandPath 'node'
if ([string]::IsNullOrWhiteSpace($nodePath)) {
    Add-ToolchainCheck -Name 'Node.js' -Tier $nodeTier -Status 'MISSING' -Details 'MathJax SVG rendering and sharp/libvips media optimization require Node.js.'
} else {
    $nodeVersion = Invoke-VersionProbe -FilePath $nodePath -Arguments @('--version')
    $nodeStatus = if ($nodeVersion.ExitCode -eq 0) { 'OK' } else { 'FAIL' }
    Add-ToolchainCheck -Name 'Node.js' -Tier $nodeTier -Status $nodeStatus -Version $nodeVersion.Text -Path $nodePath
}

$npmPath = Resolve-CommandPath 'npm'
if ([string]::IsNullOrWhiteSpace($npmPath)) {
    Add-ToolchainCheck -Name 'npm' -Tier 'Recommended' -Status 'MISSING' -Details 'Needed only when restoring node_modules.'
} else {
    $npmVersion = Invoke-VersionProbe -FilePath $npmPath -Arguments @('--version')
    $npmStatus = if ($npmVersion.ExitCode -eq 0) { 'OK' } else { 'FAIL' }
    Add-ToolchainCheck -Name 'npm' -Tier 'Recommended' -Status $npmStatus -Version $npmVersion.Text -Path $npmPath
}

Test-NodePackage -PackageName '@mathjax/src' -RelativePackageJson 'node_modules\@mathjax\src\package.json' -Tier $mathJaxTier | Out-Null
Test-NodePackage -PackageName 'sharp' -RelativePackageJson 'node_modules\sharp\package.json' -Tier $sharpTier | Out-Null

$runNodeSmoke = $Deep -or $RequireFormulaSvg -or $RequireMediaOptimization
if ($runNodeSmoke -and -not [string]::IsNullOrWhiteSpace($nodePath)) {
    $svgOut = Join-Path ([System.IO.Path]::GetTempPath()) ("physics-ppt-toolchain-" + [guid]::NewGuid().ToString('N') + ".svg")
    try {
        $renderScript = Join-Path $root 'tools\Render-FormulaSvg.mjs'
        $renderOutput = & $nodePath $renderScript --tex 'P=\frac{W}{t}' --out $svgOut 2>&1
        if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $svgOut) -and (Get-Item -LiteralPath $svgOut).Length -gt 0) {
            Add-ToolchainCheck -Name 'MathJax SVG render call' -Tier $mathJaxTier -Status 'OK' -Path $renderScript
        } else {
            Add-ToolchainCheck -Name 'MathJax SVG render call' -Tier $mathJaxTier -Status 'FAIL' -Path $renderScript -Details (($renderOutput | ForEach-Object { [string]$_ }) -join ' ')
        }
    } finally {
        if (Test-Path -LiteralPath $svgOut) { Remove-Item -LiteralPath $svgOut -Force }
    }

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

# Resolve python from PATH only: machine-specific absolute paths leak personal
# environment layout and make the check result host-dependent.
$pythonPath = Resolve-CommandPath 'python'

if ([string]::IsNullOrWhiteSpace($pythonPath) -or -not (Test-Path -LiteralPath $pythonPath)) {
    Add-ToolchainCheck -Name 'Python for OCR probes' -Tier 'Recommended' -Status 'MISSING' -Details 'RapidOCR review probes need local Python.'
} else {
    $pythonVersion = Invoke-VersionProbe -FilePath $pythonPath -Arguments @('--version')
    # Judge by exit code like the node/npm probes: the WindowsApps store
    # placeholder also prints a version-ish banner but exits non-zero.
    if ($pythonVersion.ExitCode -eq 0) {
        Add-ToolchainCheck -Name 'Python for OCR probes' -Tier 'Recommended' -Status 'OK' -Version $pythonVersion.Text -Path $pythonPath
        foreach ($module in @('PIL', 'cv2', 'numpy', 'onnxruntime', 'rapidocr_onnxruntime')) {
            Test-PythonModule -PythonExe $pythonPath -ModuleName $module -Tier 'Recommended'
        }
    } else {
        Add-ToolchainCheck -Name 'Python for OCR probes' -Tier 'Recommended' -Status 'MISSING' -Details ('python --version exited with code ' + $pythonVersion.ExitCode + ' (store placeholder or broken install).')
    }
}

foreach ($tool in @('magick', 'tesseract', 'ffmpeg', 'pngquant', 'cjpeg', 'jpegtran')) {
    $path = Resolve-CommandPath $tool
    if ([string]::IsNullOrWhiteSpace($path)) {
        Add-ToolchainCheck -Name "optional command $tool" -Tier 'Optional' -Status 'MISSING'
    } else {
        Add-ToolchainCheck -Name "optional command $tool" -Tier 'Optional' -Status 'OK' -Path $path
    }
}

foreach ($tool in @('pix2tex', 'pix2text', 'texteller')) {
    $path = Resolve-CommandPath $tool
    if ([string]::IsNullOrWhiteSpace($path)) {
        Add-ToolchainCheck -Name "experimental formula OCR $tool" -Tier 'Experimental' -Status 'MISSING' -Details 'Not part of the default reviewed workflow.'
    } else {
        Add-ToolchainCheck -Name "experimental formula OCR $tool" -Tier 'Experimental' -Status 'OK' -Path $path
    }
}

$ordered = @($checks | Sort-Object Order, Name)
$requiredFailures = @($ordered | Where-Object { $_.Tier -eq 'Required' -and $_.Status -in @('MISSING', 'FAIL') })
$strictFailures = @()
if ($Strict) {
    $strictFailures = @($ordered | Where-Object { $_.Tier -in @('Required', 'Recommended') -and $_.Status -in @('MISSING', 'FAIL') })
}

if ($AsJson) {
    [pscustomobject]@{
        generatedAt = (Get-Date).ToString('s')
        root = $root
        deep = [bool]$Deep
        launchPowerPoint = [bool]$LaunchPowerPoint
        strict = [bool]$Strict
        requireFormulaValidator = [bool]$RequireFormulaValidator
        requireFormulaSvg = [bool]$RequireFormulaSvg
        requireMediaOptimization = [bool]$RequireMediaOptimization
        requiredFailureCount = $requiredFailures.Count
        strictFailureCount = $strictFailures.Count
        checks = $ordered
    } | ConvertTo-Json -Depth 6
} else {
    $ordered | Select-Object Tier, Name, Status, Version, Path, Details | Format-Table -AutoSize
    Write-Host ("Required failures: {0}" -f $requiredFailures.Count)
    if ($Strict) { Write-Host ("Strict failures: {0}" -f $strictFailures.Count) }
}

if ($requiredFailures.Count -gt 0 -or ($Strict -and $strictFailures.Count -gt 0)) {
    throw "Toolchain check failed: required=$($requiredFailures.Count), strict=$($strictFailures.Count)."
}
