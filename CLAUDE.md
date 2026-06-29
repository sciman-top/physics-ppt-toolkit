# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Junior-middle-school physics PPT style normalization toolkit. Batch-applies consistent fonts, sizes, colors, and backgrounds to existing PPT files without changing text content, object positions, sizes, animations, or picture crops.

**Dual-implementation architecture:**
- **PowerShell** (`tools/`): Batch COM automation via PowerPoint.Application. Reads `config/physics-ppt-style.config.json`.
- **VBA** (`vba/`): Standalone offline macros running inside PowerPoint. Uses hardcoded constants — no JSON dependency.

## Commands

```powershell
# Self-check (no PowerPoint required)
powershell -File tools/Test-ToolkitFiles.ps1

# Report-only (inspect without modifying)
.\tools\Report-PhysicsPptStyle.ps1 -InputPath "D:\课件" -OutputDir "D:\报告" -Recurse

# Normalize (creates new files, never overwrites originals)
.\tools\Normalize-PhysicsPpt.ps1 -InputPath "D:\课件" -OutputDir "D:\输出" -Recurse

# Normalize with master style update
.\tools\Normalize-PhysicsPpt.ps1 -InputPath "D:\课件" -OutputDir "D:\输出" -Recurse -UpdateMaster

# PowerShell syntax validation (all .ps1 files)
$files = Get-ChildItem tools\*.ps1; foreach ($f in $files) { $t=$null; $e=$null; [System.Management.Automation.Language.Parser]::ParseFile($f.FullName,[ref]$t,[ref]$e) | Out-Null; if($e.Count -gt 0){Write-Host "ERROR: $($f.Name)"}else{Write-Host "OK: $($f.Name)"} }
```

## Architecture

### PowerShell Pipeline
`config JSON → Normalize-PhysicsPpt.ps1 → COM (PowerPoint.Application) → normalized .pptx + CSV report`

- `Normalize-PhysicsPpt.ps1` is the core; `Report-PhysicsPptStyle.ps1` is a thin wrapper passing `-ReportOnly`.
- COM calls use retry logic (`Invoke-WithComRetry`) and explicit `ReleaseComObject` + GC cleanup in `finally` blocks.
- All constants use `$script:` scope prefix.

### VBA Module Dependencies
```
PhysicsPptCommon.bas  ←  shared constants + utilities (MUST be imported first)
       ↑
       ├── PhysicsPptNormalize.bas    (full normalization + report)
       ├── PhysicsPptReportOnly.bas   (read-only inspection)
       └── ApplyPhysicsPptMasterStyle.bas  (master slide style)
```

`PhysicsPptCommon.bas` provides: `Public Const` (fonts, sizes, Office enums), `GuardActivePresentation`, `GetReportPath`, `ShapeHasText`, `GetShapeText`, `GetTextSize`, `IsFormulaCandidateText`, `CsvLine`, `CsvEscape`, `SaveReport`.

### Critical: JSON ↔ VBA Config Sync

VBA macros do NOT read `config/physics-ppt-style.config.json`. When updating the JSON config, you MUST manually sync the corresponding VBA constants in `PhysicsPptCommon.bas`. The mapping is documented in comments at the top of each VBA file:
- `FONT_CN` ↔ `fonts.chinese`
- `FONT_LATIN` ↔ `fonts.latin`
- `FONT_MATH` ↔ `fonts.math`
- `SIZE_*` ↔ `fontSizes.*`

## Key Constraints

- **Zero content modification**: Only style properties (font, size, color, background, decorative effects).
- **Grouped shapes are skipped**: Never ungroup or modify group internals to prevent layout damage.
- **Highlight box normalization**: Only targets yellow-ish fills (R>200, G>200, B<180) — does not alter arbitrary colored shapes.
- **Video slide detection**: Based on `msoMedia` type or keywords from config (`videoSlideKeywords`).
- **Formula detection**: Pattern-based heuristics (`=`, `η`, `Ω`, `W有`, `R1`, etc.). Report-only by default; no auto-conversion.
- **CSV output**: Must include UTF-8 BOM for Excel Chinese character compatibility.
- **Backup**: Originals always copied to `_backup_originals/` subfolder (unless `-NoBackup`).

## Testing

No unit test framework. Validation relies on:
1. `Test-ToolkitFiles.ps1` — file existence, JSON schema, PS syntax, VBA `Option Explicit`, color hex format.
2. Manual review of output PPTX files.
3. PowerShell parser syntax check (command above).
