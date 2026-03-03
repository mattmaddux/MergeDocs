<#
.SYNOPSIS
    PDF Merger - Merges PDFs and Word documents into a single PDF.
.DESCRIPTION
    Add any combination of PDF and Word (.doc/.docx) files, reorder them,
    then merge into a single PDF. Word documents are converted to PDF via
    Microsoft Word automation before merging. PdfSharp library is
    auto-downloaded on first run.
#>

#Requires -Version 5.1
param([switch]$SkipWordCheck)

# ============================================================
# SETUP
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$libDir    = Join-Path $scriptDir "lib"
$dllPath   = Join-Path $libDir "PdfSharp.dll"

# ============================================================
# WORD COM COMPATIBILITY CHECK
# Detects bitness mismatch (e.g. 64-bit PowerShell + 32-bit Office)
# and auto-relaunches in the matching PowerShell if needed.
# ============================================================
if (-not $SkipWordCheck) {
    $testWord = $null
    try {
        $testWord = New-Object -ComObject Word.Application
        $testWord.Visible = $false  # triggers TYPE_E_CANTLOADLIBRARY on mismatch
        $testWord.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($testWord)
        $testWord = $null
        [System.GC]::Collect()
    }
    catch {
        # Clean up the test Word instance so it doesn't leave a zombie process
        if ($testWord) {
            try { $testWord.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($testWord) } catch {}
            [System.GC]::Collect()
        }

        if ($_.Exception.Message -match 'TYPE_E_CANTLOADLIBRARY|80029C4A') {
            if ([Environment]::Is64BitProcess) {
                $altPs = Join-Path $env:SystemRoot "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
            } else {
                $altPs = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
            }

            if (Test-Path $altPs) {
                Start-Process $altPs -ArgumentList @(
                    "-ExecutionPolicy", "Bypass",
                    "-File", "`"$PSCommandPath`"",
                    "-SkipWordCheck"
                )
                exit
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Word is installed as a different architecture than PowerShell.`n`n" +
                    "Word features will be unavailable. PDF-only merges will still work.",
                    "Word Compatibility", "OK", "Warning")
            }
        }
        else {
            # Word not installed or other issue — continue, PDF merges still work
        }
    }
}

# Auto-download PdfSharp on first run
if (-not (Test-Path $dllPath)) {
    if (-not (Test-Path $libDir)) {
        New-Item -ItemType Directory -Path $libDir -Force | Out-Null
    }

    $nugetUrl    = "https://www.nuget.org/api/v2/package/PdfSharp/1.50.5147"
    $tempZip     = Join-Path $env:TEMP "PdfSharp_$(Get-Random).zip"
    $tempExtract = Join-Path $env:TEMP "PdfSharp_extract_$(Get-Random)"

    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $progressPref = $ProgressPreference
        $ProgressPreference = "SilentlyContinue"
        Invoke-WebRequest -Uri $nugetUrl -OutFile $tempZip -UseBasicParsing
        $ProgressPreference = $progressPref

        Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force

        $dllSource = Get-ChildItem -Path $tempExtract -Filter "PdfSharp.dll" -Recurse |
            Where-Object { $_.FullName -match 'net' } |
            Select-Object -First 1

        if (-not $dllSource) { throw "PdfSharp.dll not found in NuGet package." }

        Copy-Item $dllSource.FullName $dllPath -Force
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to download PdfSharp library:`n$_`n`nCheck your internet connection and try again.",
            "Setup Error", "OK", "Error")
        exit 1
    }
    finally {
        Remove-Item $tempZip     -Force   -ErrorAction SilentlyContinue
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Add-Type -Path $dllPath

# ============================================================
# HELPER: Convert Word doc to PDF via COM automation
# Returns the path to a temporary PDF, or $null on failure.
# ============================================================
function Convert-WordToPdf {
    param([string]$WordPath)

    $tempPdf = Join-Path $env:TEMP "merge_$(Get-Random).pdf"
    $word = $null
    $doc  = $null

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0              # wdAlertsNone
        $word.AutomationSecurity = 3         # msoAutomationSecurityForceDisable — prevent macro prompts

        # ConfirmConversions=$false, ReadOnly=$true, AddToRecentFiles=$false
        $doc = $word.Documents.Open($WordPath, $false, $true, $false)

        if ($null -eq $doc) {
            throw "Word could not open the file. It may be corrupted or in an unsupported format."
        }

        $doc.SaveAs($tempPdf, 17)  # 17 = wdFormatPDF
        $doc.Close(0)  # wdDoNotSaveChanges
        $doc = $null

        return $tempPdf
    }
    catch {
        throw "Failed to convert '$([System.IO.Path]::GetFileName($WordPath))' to PDF:`n$_"
    }
    finally {
        if ($doc)  { try { $doc.Close(0) } catch {} }
        if ($word) { try { $word.Quit()        } catch {} }
        if ($word) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [System.GC]::Collect()
    }
}

# ============================================================
# BUILD THE GUI
# ============================================================
$form = New-Object System.Windows.Forms.Form -Property @{
    Text            = "PDF Merger"
    Size            = New-Object System.Drawing.Size(560, 420)
    StartPosition   = "CenterScreen"
    MinimumSize     = New-Object System.Drawing.Size(480, 350)
    Font            = New-Object System.Drawing.Font("Segoe UI", 9)
}

# --- File list ---
$listBox = New-Object System.Windows.Forms.ListBox -Property @{
    Location      = New-Object System.Drawing.Point(20, 20)
    Size          = New-Object System.Drawing.Size(390, 260)
    Anchor        = "Top,Left,Right,Bottom"
    SelectionMode = "One"
}
$form.Controls.Add($listBox)

# We store full paths separately since the listbox shows display names
$filePaths = [System.Collections.ArrayList]::new()

# --- Buttons panel (right side) ---
$btnX     = 420
$btnW     = 100
$btnFont  = New-Object System.Drawing.Font("Segoe UI", 9)

$addBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Add Files..."
    Location = New-Object System.Drawing.Point($btnX, 20)
    Size     = New-Object System.Drawing.Size($btnW, 30)
    Anchor   = "Top,Right"
    Font     = $btnFont
}

$removeBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Remove"
    Location = New-Object System.Drawing.Point($btnX, 58)
    Size     = New-Object System.Drawing.Size($btnW, 30)
    Anchor   = "Top,Right"
    Font     = $btnFont
}

$upBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Move Up"
    Location = New-Object System.Drawing.Point($btnX, 108)
    Size     = New-Object System.Drawing.Size($btnW, 30)
    Anchor   = "Top,Right"
    Font     = $btnFont
}

$downBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Move Down"
    Location = New-Object System.Drawing.Point($btnX, 146)
    Size     = New-Object System.Drawing.Size($btnW, 30)
    Anchor   = "Top,Right"
    Font     = $btnFont
}

$clearBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Clear All"
    Location = New-Object System.Drawing.Point($btnX, 196)
    Size     = New-Object System.Drawing.Size($btnW, 30)
    Anchor   = "Top,Right"
    Font     = $btnFont
}

$form.Controls.Add($addBtn)
$form.Controls.Add($removeBtn)
$form.Controls.Add($upBtn)
$form.Controls.Add($downBtn)
$form.Controls.Add($clearBtn)

# --- Merge button (bottom) ---
$mergeBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Merge to PDF..."
    Size     = New-Object System.Drawing.Size(160, 40)
    Anchor   = "Bottom,Right"
    Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
}
# Position it relative to form bottom-right
$mergeBtn.Location = New-Object System.Drawing.Point(
    ($form.ClientSize.Width  - $mergeBtn.Width  - 20),
    ($form.ClientSize.Height - $mergeBtn.Height - 15))
$form.Controls.Add($mergeBtn)

# --- Status label (bottom left) ---
$statusLabel = New-Object System.Windows.Forms.Label -Property @{
    Text      = "Add files to begin."
    Location  = New-Object System.Drawing.Point(20, ($form.ClientSize.Height - 42))
    Size      = New-Object System.Drawing.Size(300, 20)
    Anchor    = "Bottom,Left"
    ForeColor = [System.Drawing.SystemColors]::GrayText
}
$form.Controls.Add($statusLabel)

# ============================================================
# HELPER: refresh the listbox display from $filePaths
# ============================================================
function Update-ListDisplay {
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    for ($i = 0; $i -lt $filePaths.Count; $i++) {
        $name = [System.IO.Path]::GetFileName($filePaths[$i])
        $ext  = [System.IO.Path]::GetExtension($filePaths[$i]).ToLower()
        $tag  = if ($ext -eq ".pdf") { "PDF" } else { "Word" }
        $listBox.Items.Add("$($i + 1). [$tag] $name") | Out-Null
    }
    $listBox.EndUpdate()

    $count = $filePaths.Count
    $statusLabel.Text = if ($count -eq 0) { "Add files to begin." }
                        elseif ($count -eq 1) { "1 file added." }
                        else { "$count files added." }
}

# ============================================================
# BUTTON HANDLERS
# ============================================================

# --- Add Files ---
$addBtn.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title       = "Select PDFs or Word documents"
        Filter      = "Supported files (*.pdf;*.doc;*.docx)|*.pdf;*.doc;*.docx|PDF files (*.pdf)|*.pdf|Word documents (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*"
        Multiselect = $true
    }

    if ($dlg.ShowDialog() -eq "OK") {
        foreach ($f in ($dlg.FileNames | Sort-Object)) {
            $filePaths.Add($f) | Out-Null
        }
        Update-ListDisplay
        $listBox.SelectedIndex = $listBox.Items.Count - 1
    }
})

# --- Remove ---
$removeBtn.Add_Click({
    $idx = $listBox.SelectedIndex
    if ($idx -ge 0) {
        $filePaths.RemoveAt($idx)
        Update-ListDisplay
        if ($filePaths.Count -gt 0) {
            $listBox.SelectedIndex = [Math]::Min($idx, $filePaths.Count - 1)
        }
    }
})

# --- Move Up ---
$upBtn.Add_Click({
    $idx = $listBox.SelectedIndex
    if ($idx -gt 0) {
        $temp = $filePaths[$idx]
        $filePaths[$idx]     = $filePaths[$idx - 1]
        $filePaths[$idx - 1] = $temp
        Update-ListDisplay
        $listBox.SelectedIndex = $idx - 1
    }
})

# --- Move Down ---
$downBtn.Add_Click({
    $idx = $listBox.SelectedIndex
    if ($idx -ge 0 -and $idx -lt ($filePaths.Count - 1)) {
        $temp = $filePaths[$idx]
        $filePaths[$idx]     = $filePaths[$idx + 1]
        $filePaths[$idx + 1] = $temp
        Update-ListDisplay
        $listBox.SelectedIndex = $idx + 1
    }
})

# --- Clear All ---
$clearBtn.Add_Click({
    if ($filePaths.Count -gt 0) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Remove all files from the list?",
            "Clear All", "YesNo", "Question")
        if ($answer -eq "Yes") {
            $filePaths.Clear()
            Update-ListDisplay
        }
    }
})

# --- Merge ---
$mergeBtn.Add_Click({
    if ($filePaths.Count -lt 2) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please add at least 2 files to merge.",
            "Not Enough Files", "OK", "Warning")
        return
    }

    # Ask where to save
    $saveDlg = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        Title      = "Save merged PDF as"
        Filter     = "PDF files (*.pdf)|*.pdf"
        DefaultExt = "pdf"
        FileName   = "Merged_$(Get-Date -Format 'yyyy-MM-dd')"
    }

    if ($saveDlg.ShowDialog() -ne "OK") { return }

    $outputFile = $saveDlg.FileName
    $tempFiles  = [System.Collections.ArrayList]::new()

    # Change cursor to wait
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $mergeBtn.Enabled = $false
    $statusLabel.Text = "Merging..."
    $form.Refresh()

    try {
        # Step 1: Convert any Word docs to temp PDFs
        $pdfFiles = @()
        foreach ($file in $filePaths) {
            $ext = [System.IO.Path]::GetExtension($file).ToLower()
            if ($ext -eq ".doc" -or $ext -eq ".docx") {
                $statusLabel.Text = "Converting $([System.IO.Path]::GetFileName($file))..."
                $form.Refresh()
                $tempPdf = Convert-WordToPdf $file
                $tempFiles.Add($tempPdf) | Out-Null
                $pdfFiles += $tempPdf
            }
            else {
                $pdfFiles += $file
            }
        }

        # Step 2: Merge all PDFs with PdfSharp
        $statusLabel.Text = "Merging PDFs..."
        $form.Refresh()

        $outputDoc  = New-Object PdfSharp.Pdf.PdfDocument
        $totalPages = 0

        foreach ($pdf in $pdfFiles) {
            $inputDoc = [PdfSharp.Pdf.IO.PdfReader]::Open(
                $pdf, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)

            for ($p = 0; $p -lt $inputDoc.PageCount; $p++) {
                $outputDoc.AddPage($inputDoc.Pages[$p]) | Out-Null
                $totalPages++
            }
        }

        $outputDoc.Save($outputFile)
        $outputDoc.Close()

        [System.Windows.Forms.MessageBox]::Show(
            "Merged PDF saved to:`n$outputFile`n`n$($filePaths.Count) files merged, $totalPages total pages.",
            "Success", "OK", "Information")

        $statusLabel.Text = "Done - $totalPages pages saved."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error during merge:`n$_",
            "Merge Error", "OK", "Error")
        $statusLabel.Text = "Merge failed."
    }
    finally {
        # Clean up temp files from Word conversion
        foreach ($t in $tempFiles) {
            Remove-Item $t -Force -ErrorAction SilentlyContinue
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $mergeBtn.Enabled = $true
    }
})

# ============================================================
# SHOW
# ============================================================
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
$form.Dispose()
