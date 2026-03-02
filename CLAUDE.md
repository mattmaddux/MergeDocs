# MergeDocs

A PowerShell GUI tool for merging PDFs and Word documents (.doc/.docx) into a single PDF.

## Tech Stack
- PowerShell 5.1+ with Windows Forms for the GUI
- PdfSharp 1.50 (.NET Framework) for PDF merging — auto-downloaded from NuGet on first run
- Microsoft Word COM automation for .doc/.docx to PDF conversion

## Project Structure
- `MergeDocs.ps1` — the entire tool is a single self-contained script
- `lib/PdfSharp.dll` — auto-created on the user's machine at first run (not committed)

## Key Design Decisions
- Targets Windows PowerShell 5.1 (ships with Windows 10/11) for maximum compatibility
- PdfSharp 1.50.x specifically because it targets .NET Framework 4.x (PowerShell 5.1's runtime); PdfSharp 6.x requires .NET 6+ and won't work
- Word COM automation handles both .doc and .docx formats
- Temp files from Word conversion are cleaned up in a finally block

## Deployment
Users create a shortcut with target:
```
powershell.exe -ExecutionPolicy Bypass -File "C:\path\to\MergeDocs.ps1"
```
