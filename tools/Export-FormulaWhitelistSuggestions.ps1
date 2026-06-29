<#
.SYNOPSIS
  Export reviewable formula whitelist suggestions from formula-review.csv.

.DESCRIPTION
  Aggregates formula-review.csv, filters out incomplete labels, numeric-only
  values, long Chinese explanations, placeholders, and existing whitelist
  entries, then proposes controlled sourcePattern / UnicodeMath / LaTeX rows.
  The script is report-only and never writes config or PPTX files.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FormulaReviewCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$ConfigPath,

    [ValidateRange(1, 500)]
    [int]$MaxSuggestions = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Utf8BomCsv {
    param([object[]]$Rows, [string]$Path)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $csvLines = $Rows | ConvertTo-Csv -NoTypeInformation
    [System.IO.File]::WriteAllLines($Path, $csvLines, $utf8Bom)
}

function Write-Utf8BomText {
    param([string]$Text, [string]$Path)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Text, $utf8Bom)
}

function Normalize-FormulaSource {
    param([string]$Text)
    $t = [string]$Text
    $t = $t -replace '\s+', ''
    $t = $t.Replace('＝', '=').Replace('＋', '+').Replace('－', '-').Replace('（', '(').Replace('）', ')')
    $t = $t.Replace('×', '*').Replace('∙', '·').Replace('·', '·').Replace('÷', '/')
    return $t.Trim()
}

function Convert-ToTex {
    param([string]$Text)

    $t = [string]$Text
    $suffixChars = @('物', '总', '有', '额', '动', '排', '液')
    foreach ($ch in $suffixChars) {
        $t = $t -replace "([A-Za-zρτηΩμ])$([regex]::Escape($ch))", ('$1_{\text{' + $ch + '}}')
    }
    $t = $t -replace '([A-Za-zρτηΩμ])([0-9]+)', '$1_{$2}'
    $t = $t.Replace('＝', '=')
    $t = $t.Replace('×', '\times ')
    $t = $t.Replace('*', '\times ')
    $t = $t.Replace('∙', '\cdot ')
    $t = $t.Replace('·', '\cdot ')
    $t = $t.Replace('ρ', '\rho ')
    $t = $t.Replace('η', '\eta ')
    $t = $t.Replace('Ω', '\Omega ')
    $t = $t.Replace('％', '\%')
    return ($t -replace '\s+', ' ').Trim()
}

function Convert-ToUnicodeMath {
    param([string]$Text)

    $t = [string]$Text
    $t = $t.Replace('＝', '=').Replace('×', '×').Replace('*', '×').Replace('∙', '·')
    foreach ($ch in @('物', '总', '有', '额', '动', '排', '液')) {
        $t = $t -replace "([A-Za-zρτηΩμ])$([regex]::Escape($ch))", ('$1_' + $ch)
    }
    $t = $t -replace '([A-Za-zρτηΩμ])([0-9]+)', '$1_$2'
    return ($t -replace '\s+', ' ').Trim()
}

function Test-ExistingWhitelistMatch {
    param([string]$Normalized, [object[]]$Rules)
    foreach ($rule in @($Rules)) {
        try {
            if ($Normalized -match ([string]$rule.sourcePattern)) { return $true }
        } catch { }
    }
    return $false
}

function Get-SuggestionDecision {
    param([string]$Normalized, [int]$Count, [int]$FileCount)

    if ([string]::IsNullOrWhiteSpace($Normalized)) {
        return [pscustomobject]@{ Status = 'Rejected'; Reason = 'empty' }
    }
    if ($Normalized.Length -gt 48) {
        return [pscustomobject]@{ Status = 'Rejected'; Reason = 'too-long' }
    }
    if ($Normalized -match '[_]{2,}|…|\.{2,}|意义|方法|问题|分析|有关|无关|说明|使用|测量|观察') {
        return [pscustomobject]@{ Status = 'Rejected'; Reason = 'placeholder-or-explanation' }
    }
    if ($Normalized -match '^[=+\-*/·×÷]+$') {
        return [pscustomobject]@{ Status = 'Rejected'; Reason = 'operator-only' }
    }
    if ($Normalized -match '^[=+\-*/·×÷<>≈]' -or $Normalized -match '[=+\-*/·×÷<>≈]$') {
        return [pscustomobject]@{ Status = 'Rejected'; Reason = 'incomplete-expression' }
    }
    if ($Normalized -match '^[≈=]?\d+(\.\d+)?(%|％|Ω|V|A|W|J|N|m|cm|s|min)?$') {
        return [pscustomobject]@{ Status = 'Rejected'; Reason = 'numeric-value-only' }
    }
    if ($Normalized -match '[\p{IsCJKUnifiedIdeographs}]') {
        $allowed = '物总有额动排液'
        $chars = @([regex]::Matches($Normalized, '[\p{IsCJKUnifiedIdeographs}]') | ForEach-Object { $_.Value } | Sort-Object -Unique)
        foreach ($ch in $chars) {
            if (-not $allowed.Contains($ch)) {
                return [pscustomobject]@{ Status = 'Rejected'; Reason = 'contains-non-formula-chinese' }
            }
        }
    }
    if ($Normalized -notmatch '[=<>≈]|[/·×*÷]') {
        return [pscustomobject]@{ Status = 'ReviewSubscriptOnly'; Reason = 'no-equation-operator' }
    }
    if ($Normalized -match '^[A-Za-zρτηΩμ0-9物总有额动排液_+\-*/=<>≈.·×÷()%％]+$') {
        $priority = if ($Count -ge 3 -or $FileCount -ge 2) { 'PromoteCandidate' } else { 'ReviewCandidate' }
        return [pscustomobject]@{ Status = $priority; Reason = 'low-risk-formula-pattern' }
    }
    return [pscustomobject]@{ Status = 'ReviewCandidate'; Reason = 'needs-human-formula-review' }
}

$FormulaReviewCsv = [System.IO.Path]::GetFullPath($FormulaReviewCsv)
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path -LiteralPath $FormulaReviewCsv)) { throw "FormulaReviewCsv not found: $FormulaReviewCsv" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'config\physics-ppt-style.config.json'
}
$ConfigPath = [System.IO.Path]::GetFullPath($ConfigPath)
$whitelistRules = @()
if (Test-Path -LiteralPath $ConfigPath) {
    $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $whitelistRules = @($config.formulaWhitelist)
}

$rows = @(Import-Csv -LiteralPath $FormulaReviewCsv -Encoding UTF8)
$groups = @(
    $rows |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.FormulaText) } |
        Group-Object { Normalize-FormulaSource -Text $_.FormulaText }
)

$items = New-Object System.Collections.Generic.List[object]
foreach ($group in $groups) {
    $normalized = [string]$group.Name
    $fileCount = @($group.Group | Select-Object -ExpandProperty File -Unique).Count
    $slideCount = @($group.Group | ForEach-Object { "$($_.File)|$($_.Slide)" } | Sort-Object -Unique).Count
    $decision = Get-SuggestionDecision -Normalized $normalized -Count $group.Count -FileCount $fileCount
    $exists = Test-ExistingWhitelistMatch -Normalized $normalized -Rules $whitelistRules
    $status = if ($exists) { 'ExistingWhitelist' } else { $decision.Status }

    $firstText = [string]$group.Group[0].FormulaText
    $items.Add([pscustomobject]@{
        SuggestionStatus = $status
        Reason = if ($exists) { 'already-covered-by-config' } else { $decision.Reason }
        Frequency = $group.Count
        FileCount = $fileCount
        SlideCount = $slideCount
        FormulaText = $firstText
        NormalizedSource = $normalized
        ProposedName = ('公式_' + ($normalized -replace '[^\p{L}\p{Nd}]+', '_')).Trim('_')
        ProposedSourcePattern = ('^' + [regex]::Escape($normalized) + '$')
        ProposedUnicodeMath = Convert-ToUnicodeMath -Text $firstText
        ProposedTex = Convert-ToTex -Text $firstText
        SourceFiles = (@($group.Group | ForEach-Object { if (-not [string]::IsNullOrWhiteSpace($_.FilePath)) { $_.FilePath } else { $_.File } } | Sort-Object -Unique) -join ';')
        SourceSlides = (@($group.Group | ForEach-Object { "$($(if (-not [string]::IsNullOrWhiteSpace($_.FilePath)) { $_.FilePath } else { $_.File })):$($_.Slide)" } | Sort-Object -Unique | Select-Object -First 12) -join ';')
        SuggestedActions = (@($group.Group | Select-Object -ExpandProperty SuggestedAction -Unique | Sort-Object) -join ';')
    }) | Out-Null
}

$ordered = @(
    $items.ToArray() |
        Sort-Object `
            @{ Expression = { switch ($_.SuggestionStatus) { 'PromoteCandidate' { 1 } 'ReviewCandidate' { 2 } 'ReviewSubscriptOnly' { 3 } 'ExistingWhitelist' { 4 } default { 9 } } }; Ascending = $true },
            @{ Expression = { [int]$_.Frequency }; Descending = $true },
            @{ Expression = { [int]$_.FileCount }; Descending = $true } |
        Select-Object -First $MaxSuggestions
)

$csvPath = Join-Path $OutputDir 'formula-whitelist-suggestions.csv'
$manifestPath = Join-Path $OutputDir 'formula-whitelist-suggestions-manifest.json'
Write-Utf8BomCsv -Rows $ordered -Path $csvPath

$manifest = [ordered]@{
    generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    formulaReviewCsv = $FormulaReviewCsv
    outputDir = $OutputDir
    suggestionCsv = $csvPath
    totalGroups = $groups.Count
    exportedSuggestions = $ordered.Count
    promoteCandidateCount = @($ordered | Where-Object { $_.SuggestionStatus -eq 'PromoteCandidate' }).Count
    reviewCandidateCount = @($ordered | Where-Object { $_.SuggestionStatus -eq 'ReviewCandidate' }).Count
    reviewSubscriptOnlyCount = @($ordered | Where-Object { $_.SuggestionStatus -eq 'ReviewSubscriptOnly' }).Count
    existingWhitelistCount = @($ordered | Where-Object { $_.SuggestionStatus -eq 'ExistingWhitelist' }).Count
    rule = 'Report-only whitelist suggestions. Do not write config automatically; promote only after source review, OMML generation, and Open XML validation; use rendered visual gate for broad rollout or risky formulas.'
}
Write-Utf8BomText -Text ($manifest | ConvertTo-Json -Depth 5) -Path $manifestPath

Write-Host "Formula whitelist suggestions done: $OutputDir"
Write-Host "Promote: $($manifest.promoteCandidateCount); Review: $($manifest.reviewCandidateCount); Subscript-only: $($manifest.reviewSubscriptOnlyCount); Existing: $($manifest.existingWhitelistCount)"
