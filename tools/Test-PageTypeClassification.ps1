<#
.SYNOPSIS
  Behavior probe for the brand chain's page-type classification.

.DESCRIPTION
  Exercises Test-DividerSlide and Get-SlideRole from
  Apply-PptxBrandVisualRefresh.ps1 against synthetic slides built from plain
  PSObjects (the functions only duck-type Name/Type/HasTextFrame/
  TextFrame.HasText/TextRange.Text, so no PowerPoint COM is needed).
  Cases pin the real-sample regression decisions recorded in the source
  comments: 18.2 slide11 "kW·h" and slide14 "P=UI" formula dividers, 18.2
  slide28 16-char CJK divider, the 15.1 slide34 symbol-diagram exclusion,
  the brand-icon skip, and the 20-char / 6-char boundaries.
  It never starts PowerPoint and never creates or edits a PPTX.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$brandToolPath = Join-Path $PSScriptRoot 'Apply-PptxBrandVisualRefresh.ps1'
$brandAst = [System.Management.Automation.Language.Parser]::ParseFile($brandToolPath, [ref]$null, [ref]$null)
$probeFunctions = @{}
foreach ($functionName in @('Test-DividerSlide', 'Get-SlideRole')) {
    $functionAst = $brandAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) |
        Where-Object { $_.Name -eq $functionName } |
        Select-Object -First 1
    if ($null -eq $functionAst) { throw "Function not found in brand tool: $functionName" }
    # Invoke-Expression of the full definition text is the reliable way to load
    # a function body outside its script (same pattern as Test-ToolkitFiles 2b).
    Invoke-Expression $functionAst.Extent.Text
    $probeFunctions[$functionName] = $true
}

function New-PageShape {
    param([string]$Text = '', [int]$Type = 17, [string]$Name = 'TextBox')
    return [pscustomobject]@{
        Name         = $Name
        Type         = $Type
        HasTextFrame = -1
        TextFrame    = [pscustomobject]@{
            HasText   = -1
            TextRange = [pscustomobject]@{ Text = $Text }
        }
    }
}

function New-PageSlide {
    param([object[]]$Shapes)
    return [pscustomobject]@{ Shapes = $Shapes }
}

# Mirrors Get-SlidePlainText: all shape texts joined with newlines, including
# empty ones, so boundary lengths match production input exactly.
function Get-PageText {
    param([object[]]$Shapes)
    $parts = @()
    foreach ($shape in @($Shapes)) { $parts += [string]$shape.TextFrame.TextRange.Text }
    return ($parts -join "`n")
}

$script:caseCount = 0
function Assert-PageClass {
    param(
        [string]$Case,
        [object[]]$Shapes,
        [string]$ExpectedRole,
        [bool]$ExpectedDivider
    )
    $slide = New-PageSlide -Shapes $Shapes
    $slideText = Get-PageText -Shapes $Shapes
    $dividerResult = [bool](Test-DividerSlide -Slide $slide -SlideText $slideText)
    $roleResult = [string](Get-SlideRole -Slide $slide -SlideText $slideText)
    if ($dividerResult -ne $ExpectedDivider) { throw ("Divider case '{0}: expected {1}, got {2}." -f $Case, $ExpectedDivider, $dividerResult) }
    if ($roleResult -ne $ExpectedRole) { throw ("Role case '{0}: expected {1}, got {2}." -f $Case, $ExpectedRole, $roleResult) }
    $script:caseCount++
}

# --- Divider positives ---
# Bare CJK headline divider (13.3/14.x series style: one text shape).
Assert-PageClass -Case 'CJK headline divider' -Shapes @(, (New-PageShape -Text '比热容')) -ExpectedRole 'Divider' -ExpectedDivider $true
# 18.2 slide28 regression: multi-word CJK divider headline (<=20 chars).
$longDivider = '额定功率、实际功率/伏安法测电功率'
if ($longDivider.Length -gt 20 -or $longDivider.Length -lt 2) { throw 'Fixture drift: divider headline outside the CJK boundary.' }
Assert-PageClass -Case 'multi-word CJK divider (18.2 s28)' -Shapes @(, (New-PageShape -Text $longDivider)) -ExpectedRole 'Divider' -ExpectedDivider $true
# 18.2 slide14 / slide11 regressions: formula and unit dividers without CJK.
Assert-PageClass -Case 'formula divider P=UI (18.2 s14)' -Shapes @(, (New-PageShape -Text 'P=UI')) -ExpectedRole 'Divider' -ExpectedDivider $true
Assert-PageClass -Case 'unit divider kW·h (18.2 s11)' -Shapes @(, (New-PageShape -Text 'kW·h')) -ExpectedRole 'Divider' -ExpectedDivider $true
# Brand icon pictures are skipped before the disqualifying type check.
Assert-PageClass -Case 'brand icon + title' -Shapes @(
    (New-PageShape -Text '内能'),
    (New-PageShape -Type 13 -Name 'sciman-brand-icon')) -ExpectedRole 'Divider' -ExpectedDivider $true
# Text-less autoshapes (decorative rectangles) do not disqualify.
Assert-PageClass -Case 'empty autoshape + title' -Shapes @(
    (New-PageShape -Text '第十四章'),
    (New-PageShape -Type 1)) -ExpectedRole 'Divider' -ExpectedDivider $true
# Boundary pin: exactly 20 chars still divider, 21 chars never.
$twentyChars = '比热容' * 6 + '比热'
if ($twentyChars.Length -ne 20) { throw 'Fixture drift: expected 20-char divider headline.' }
Assert-PageClass -Case '20-char CJK boundary' -Shapes @(, (New-PageShape -Text $twentyChars)) -ExpectedRole 'Divider' -ExpectedDivider $true
Assert-PageClass -Case '21-char CJK over boundary' -Shapes @(, (New-PageShape -Text ($twentyChars + '内'))) -ExpectedRole 'Content' -ExpectedDivider $false
# Formula divider boundary: 6 chars with '=' divider, 7 chars content.
Assert-PageClass -Case '6-char formula boundary' -Shapes @(, (New-PageShape -Text 'P=UI+1')) -ExpectedRole 'Divider' -ExpectedDivider $true
Assert-PageClass -Case '7-char formula over boundary' -Shapes @(, (New-PageShape -Text 'P=UI+12')) -ExpectedRole 'Content' -ExpectedDivider $false

# --- Divider negatives ---
Assert-PageClass -Case 'two text shapes' -Shapes @(
    (New-PageShape -Text '比热容'),
    (New-PageShape -Text '第十三章')) -ExpectedRole 'Content' -ExpectedDivider $false
foreach ($disqualifier in @(
    @{ Case = 'picture'; Type = 13 },
    @{ Case = 'group'; Type = 6 },
    @{ Case = 'ole'; Type = 7 },
    @{ Case = 'chart'; Type = 3 },
    @{ Case = 'media'; Type = 16 },
    @{ Case = 'table'; Type = 19 })) {
    Assert-PageClass -Case "content shape: $($disqualifier.Case)" -Shapes @(
        (New-PageShape -Text '比热容'),
        (New-PageShape -Type $disqualifier.Type)) -ExpectedRole 'Content' -ExpectedDivider $false
}
# 15.1 slide34 regression: symbol-only text (no CJK, no [=·•×]) is content.
Assert-PageClass -Case 'symbol diagram text (15.1 s34)' -Shapes @(, (New-PageShape -Text '+ - +3 -3')) -ExpectedRole 'Content' -ExpectedDivider $false
Assert-PageClass -Case 'latin-only title' -Shapes @(, (New-PageShape -Text 'Gravitation')) -ExpectedRole 'Content' -ExpectedDivider $false
Assert-PageClass -Case 'whitespace-only slide' -Shapes @(
    (New-PageShape -Text ''),
    (New-PageShape -Text ' ')) -ExpectedRole 'Content' -ExpectedDivider $false

# --- Get-SlideRole direct roles ---
# Role-order contract: the resource/link-hub marker pages are single-text CJK
# pages that DO satisfy the divider heuristic on their own; only the regex
# checks running BEFORE the Divider branch keep them from being restyled.
# ExpectedDivider=$true below pins that precedence — reordering Get-SlideRole
# to check Divider first would flip these to Divider and the probe must fail.
Assert-PageClass -Case 'cover marker' -Shapes @(
    (New-PageShape -Text '初中物理 课后巩固'),
    (New-PageShape -Text '知乎主页：sciman')) -ExpectedRole 'Cover' -ExpectedDivider $false
Assert-PageClass -Case 'end marker' -Shapes @(, (New-PageShape -Text 'END')) -ExpectedRole 'End' -ExpectedDivider $false
Assert-PageClass -Case 'resource marker' -Shapes @(, (New-PageShape -Text '课件下载地址：见评论区')) -ExpectedRole 'Resource' -ExpectedDivider $true
Assert-PageClass -Case 'link hub marker' -Shapes @(, (New-PageShape -Text '欢迎加入网盘群')) -ExpectedRole 'LinkHub' -ExpectedDivider $true
Assert-PageClass -Case 'blank slide' -Shapes @() -ExpectedRole 'Blank' -ExpectedDivider $false
# Word-boundary discipline: END only matches as a whole word. A single short
# CJK text shape is divider-shaped by design, so the substring page lands on
# Divider — the assertion proves the End regex did not fire on 'TREND'.
Assert-PageClass -Case 'END substring not End role' -Shapes @(, (New-PageShape -Text 'TREND 观察趋势')) -ExpectedRole 'Divider' -ExpectedDivider $true
Assert-PageClass -Case 'plain content' -Shapes @(
    (New-PageShape -Text '第三节 电流的测量'),
    (New-PageShape -Text '1. 认识电流表；2. 学会使用电流表测量电流。')) -ExpectedRole 'Content' -ExpectedDivider $false

Write-Host "Page-type classification probe passed: $script:caseCount cases."
