# TestWord2PDF.ps1 - Simple standalone Word-to-PDF converter for testing.
# Uses InvokeMember to bypass broken Office PIAs.
# Usage: powershell -ExecutionPolicy Bypass -File TestWord2PDF.ps1 "C:\path\to\file.docx"
param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile
)

$pdfPath = [System.IO.Path]::ChangeExtension($InputFile, '.pdf')

$word = New-Object -ComObject Word.Application
Write-Host "Word created OK"

$docs = $word.Documents
Write-Host "Documents collection OK (Count: $($docs.Count))"

$doc = $docs.GetType().InvokeMember("Open", "InvokeMethod", $null, $docs, @($InputFile))
Write-Host "Document opened OK: $doc"

$doc.GetType().InvokeMember("SaveAs", "InvokeMethod", $null, $doc, @($pdfPath, 17))
Write-Host "SaveAs OK"

$doc.GetType().InvokeMember("Close", "InvokeMethod", $null, $doc, @(0))
Write-Host "Close OK"

$word.GetType().InvokeMember("Quit", "InvokeMethod", $null, $word, $null)
Write-Host "Quit OK"

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Host "Saved: $pdfPath"
