# LumberTools

A collection of PowerShell GUI utilities for managing client endpoints, distributed by Lumberstack MSP.

## Project Structure
```
Deploy.ps1                      # RMM script: clones/pulls repo, runs Setup.ps1
Setup.ps1                       # Creates Start Menu shortcuts from tool manifests
tools/
  <ToolName>/
    tool.json                   # Manifest: displayName, description, launcher
    <ToolName>.bat              # Launcher (double-click or shortcut target)
    <ToolName>.ps1              # Main script
    (helper scripts, lib/, etc.)
```

## Adding a New Tool
1. Create `tools/NewTool/` with `NewTool.bat`, `NewTool.ps1`, and `tool.json`
2. The `.bat` launcher pattern: `powershell.exe -ExecutionPolicy Bypass -File "%~dp0NewTool.ps1"`
3. `tool.json` format:
```json
{
    "name": "NewTool",
    "displayName": "Human-Readable Name",
    "description": "Shows in the shortcut tooltip",
    "launcher": "NewTool.bat"
}
```
4. Commit and push. Next sync + Setup.ps1 run creates the Start Menu shortcut.

## Conventions
- Each tool is self-contained in its own `tools/<Name>/` folder
- Use `$PSScriptRoot` for resolving sibling file paths (not hardcoded paths)
- Auto-downloaded dependencies go in a `lib/` subfolder (gitignored)
- Target Windows PowerShell 5.1 for maximum compatibility
- Hide the console window in GUI tools (Win32 ShowWindow)

## Deployment
`Deploy.ps1` is designed to be run via Intune or RMM (as SYSTEM). It handles both fresh installs and updates:
- If `C:\LumberTools` doesn't exist, clones the repo and runs `Setup.ps1`
- If it already exists, pulls latest changes and re-runs `Setup.ps1`
- Requires Git to be pre-installed on the endpoint

Usage: `powershell.exe -ExecutionPolicy Bypass -File "Deploy.ps1"`

## Tools

### MergeDocs (`tools/MergeDocs/`)
Merges PDFs and Word documents (.doc/.docx) into a single PDF.
- `MergeDocs.ps1` - GUI app (WinForms, drag-and-drop)
- `Word2PDF.ps1` - subprocess for Word-to-PDF conversion via COM automation
- PdfSharp 1.50.x for PDF merging (auto-downloaded; 1.50 targets .NET Framework 4.x)
- Word COM via `Documents.Open()` + `ExportAsFixedFormat()` in a clean subprocess
