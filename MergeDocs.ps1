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

# Hide the console window so only the GUI is visible
Add-Type -Name Win32 -Namespace Native -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
$consoleHwnd = [Native.Win32]::GetConsoleWindow()
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [void][Native.Win32]::ShowWindow($consoleHwnd, 0)  # 0 = SW_HIDE
}

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
# HELPER: Convert Word docs to PDF via Word2PDF.ps1 subprocess
# ============================================================
$word2pdfScript = Join-Path $scriptDir "Word2PDF.ps1"

function Convert-WordFilesToPdf {
    param([string[]]$WordFiles)

    $tempDir = Join-Path $env:TEMP "mergedocs_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    # Write file list to a temp file to avoid argument parsing issues
    $listFile = Join-Path $env:TEMP "mergedocs_list_$(Get-Random).txt"
    $WordFiles | Set-Content -LiteralPath $listFile -Encoding UTF8

    $result = & powershell.exe -ExecutionPolicy Bypass -File $word2pdfScript `
        -InputList $listFile -OutputDir $tempDir 2>&1

    Remove-Item $listFile -Force -ErrorAction SilentlyContinue

    $pdfPaths = @()
    $errors = @()
    foreach ($line in $result) {
        if ($line -is [System.Management.Automation.ErrorRecord]) {
            $errors += $line.ToString()
        }
        elseif ($line -match '\.pdf$') {
            $pdfPaths += $line.ToString()
        }
    }

    if ($errors.Count -gt 0) {
        throw ($errors -join "`n")
    }

    return $pdfPaths
}

# ============================================================
# BUILD THE GUI
# ============================================================
$form = New-Object System.Windows.Forms.Form -Property @{
    Text            = "PDF Merger"
    Size            = New-Object System.Drawing.Size(500, 420)
    StartPosition   = "CenterScreen"
    MinimumSize     = New-Object System.Drawing.Size(380, 320)
    Font            = New-Object System.Drawing.Font("Segoe UI", 9)
    AllowDrop       = $true
}

# --- File list (owner-drawn, full width) ---
$listBox = New-Object System.Windows.Forms.ListBox -Property @{
    Location      = New-Object System.Drawing.Point(20, 20)
    Size          = New-Object System.Drawing.Size(($form.ClientSize.Width - 40), 260)
    Anchor        = "Top,Left,Right,Bottom"
    SelectionMode = "One"
    DrawMode      = "OwnerDrawFixed"
    ItemHeight    = 26
    AllowDrop     = $true
}
$form.Controls.Add($listBox)

# We store full paths separately since the listbox shows display names
$filePaths = [System.Collections.ArrayList]::new()

# --- Bottom buttons ---
$clearBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Clear All"
    Size     = New-Object System.Drawing.Size(80, 36)
    Anchor   = "Bottom,Left"
    Font     = New-Object System.Drawing.Font("Segoe UI", 9)
}
$clearBtn.Location = New-Object System.Drawing.Point(20, ($form.ClientSize.Height - 52))
$form.Controls.Add($clearBtn)

$mergeBtn = New-Object System.Windows.Forms.Button -Property @{
    Text     = "Merge to PDF..."
    Size     = New-Object System.Drawing.Size(160, 40)
    Anchor   = "Bottom,Right"
    Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
}
$mergeBtn.Location = New-Object System.Drawing.Point(
    ($form.ClientSize.Width  - $mergeBtn.Width  - 20),
    ($form.ClientSize.Height - $mergeBtn.Height - 15))
$form.Controls.Add($mergeBtn)

# --- Status label (bottom center) ---
$statusLabel = New-Object System.Windows.Forms.Label -Property @{
    Text      = "Drag files here to begin."
    Location  = New-Object System.Drawing.Point(110, ($form.ClientSize.Height - 44))
    Size      = New-Object System.Drawing.Size(200, 20)
    Anchor    = "Bottom,Left"
    ForeColor = [System.Drawing.SystemColors]::GrayText
}
$form.Controls.Add($statusLabel)

# ============================================================
# DRAG STATE for internal reorder
# ============================================================
$script:dragIndex = -1
$script:dragStart = [System.Drawing.Point]::Empty
$script:isInternalDrag = $false
$script:xBtnWidth = 26

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
    $listBox.Invalidate()

    $count = $filePaths.Count
    $statusLabel.Text = if ($count -eq 0) { "Drag files here to begin." }
                        elseif ($count -eq 1) { "1 file added." }
                        else { "$count files added." }
}

# ============================================================
# OWNER-DRAW: draw items with X remove button
# ============================================================
$listBox.Add_DrawItem({
    param($sender, $e)
    if ($e.Index -lt 0) { return }

    $e.DrawBackground()

    $text = $sender.Items[$e.Index]
    $isSelected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected)
    $textColor = if ($isSelected) {
        [System.Drawing.SystemColors]::HighlightText
    } else {
        [System.Drawing.SystemColors]::WindowText
    }

    # Draw item text (leave room for X button)
    $textRect = New-Object System.Drawing.RectangleF(
        ($e.Bounds.X + 6), $e.Bounds.Y,
        ($e.Bounds.Width - $script:xBtnWidth - 10), $e.Bounds.Height)
    $sf = New-Object System.Drawing.StringFormat
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $brush = New-Object System.Drawing.SolidBrush $textColor
    $e.Graphics.DrawString($text, $e.Font, $brush, $textRect, $sf)
    $brush.Dispose()

    # Draw X button on the right
    $xRect = New-Object System.Drawing.RectangleF(
        ($e.Bounds.Right - $script:xBtnWidth), $e.Bounds.Y,
        $script:xBtnWidth, $e.Bounds.Height)
    $xBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(160, 160, 160))
    $xFont = New-Object System.Drawing.Font("Segoe UI", 8)
    $xSf = New-Object System.Drawing.StringFormat
    $xSf.Alignment = [System.Drawing.StringAlignment]::Center
    $xSf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $e.Graphics.DrawString([char]0x2715, $xFont, $xBrush, $xRect, $xSf)
    $xBrush.Dispose()
    $xFont.Dispose()

    $e.DrawFocusRectangle()
})

# Paint hint text when list is empty
$listBox.Add_Paint({
    param($sender, $e)
    if ($filePaths.Count -eq 0) {
        $hint = "Drag PDF or Word files here"
        $hintFont = New-Object System.Drawing.Font("Segoe UI", 11)
        $hintBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(170, 170, 170))
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $rect = New-Object System.Drawing.RectangleF(0, 0, $sender.ClientSize.Width, $sender.ClientSize.Height)
        $e.Graphics.DrawString($hint, $hintFont, $hintBrush, $rect, $sf)
        $hintFont.Dispose()
        $hintBrush.Dispose()
    }
})

# ============================================================
# CLICK: X button removes item
# ============================================================
$listBox.Add_MouseClick({
    param($sender, $e)
    $idx = $sender.IndexFromPoint($e.Location)
    if ($idx -lt 0) { return }

    $itemRect = $sender.GetItemRectangle($idx)
    if ($e.X -ge ($itemRect.Right - $script:xBtnWidth)) {
        $filePaths.RemoveAt($idx)
        Update-ListDisplay
        if ($filePaths.Count -gt 0) {
            $sender.SelectedIndex = [Math]::Min($idx, $filePaths.Count - 1)
        }
    }
})

# ============================================================
# DRAG-TO-REORDER: mouse handlers on listbox
# ============================================================
$listBox.Add_MouseDown({
    param($sender, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $idx = $sender.IndexFromPoint($e.Location)
    if ($idx -ge 0) {
        # Don't start drag if clicking the X button
        $itemRect = $sender.GetItemRectangle($idx)
        if ($e.X -lt ($itemRect.Right - $script:xBtnWidth)) {
            $script:dragIndex = $idx
            $script:dragStart = $e.Location
        }
    }
})

$listBox.Add_MouseMove({
    param($sender, $e)
    if ($script:dragIndex -lt 0) { return }
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) {
        $script:dragIndex = -1
        return
    }

    $dx = [Math]::Abs($e.X - $script:dragStart.X)
    $dy = [Math]::Abs($e.Y - $script:dragStart.Y)
    if ($dx -gt [System.Windows.Forms.SystemInformation]::DragSize.Width -or
        $dy -gt [System.Windows.Forms.SystemInformation]::DragSize.Height) {
        $script:isInternalDrag = $true
        $sender.DoDragDrop($script:dragIndex.ToString(), [System.Windows.Forms.DragDropEffects]::Move)
        $script:isInternalDrag = $false
        $script:dragIndex = -1
    }
})

$listBox.Add_MouseUp({
    $script:dragIndex = -1
})

# ============================================================
# DRAG-DROP: accept files from Explorer + internal reorder
# ============================================================
$listBox.Add_DragEnter({
    param($sender, $e)
    if ($script:isInternalDrag) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
    else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

$listBox.Add_DragOver({
    param($sender, $e)
    if ($script:isInternalDrag) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})

$listBox.Add_DragDrop({
    param($sender, $e)

    if ($script:isInternalDrag) {
        # Internal reorder
        $fromIdx = [int]$e.Data.GetData([System.Windows.Forms.DataFormats]::Text)
        $pt = $sender.PointToClient((New-Object System.Drawing.Point($e.X, $e.Y)))
        $toIdx = $sender.IndexFromPoint($pt)
        if ($toIdx -lt 0) { $toIdx = $filePaths.Count - 1 }

        if ($fromIdx -ne $toIdx) {
            $item = $filePaths[$fromIdx]
            $filePaths.RemoveAt($fromIdx)
            $filePaths.Insert($toIdx, $item)
            Update-ListDisplay
            $sender.SelectedIndex = $toIdx
        }
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        # External file drop
        $droppedFiles = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        $supported = @('.pdf', '.doc', '.docx')
        foreach ($f in ($droppedFiles | Sort-Object)) {
            $ext = [System.IO.Path]::GetExtension($f).ToLower()
            if ($ext -in $supported) {
                $filePaths.Add($f) | Out-Null
            }
        }
        Update-ListDisplay
        if ($filePaths.Count -gt 0) {
            $sender.SelectedIndex = $filePaths.Count - 1
        }
    }
})

# Also accept drops on the form itself (in case they miss the listbox)
$form.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})

$form.Add_DragDrop({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $droppedFiles = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        $supported = @('.pdf', '.doc', '.docx')
        foreach ($f in ($droppedFiles | Sort-Object)) {
            $ext = [System.IO.Path]::GetExtension($f).ToLower()
            if ($ext -in $supported) {
                $filePaths.Add($f) | Out-Null
            }
        }
        Update-ListDisplay
        if ($filePaths.Count -gt 0) {
            $listBox.SelectedIndex = $filePaths.Count - 1
        }
    }
})

# ============================================================
# BUTTON HANDLERS
# ============================================================

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

    # Disable UI during processing
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $mergeBtn.Enabled = $false
    $clearBtn.Enabled = $false
    $listBox.Enabled  = $false
    $statusLabel.Text = "Merging..."
    $form.Refresh()

    try {
        # Step 1: Convert any Word docs to temp PDFs via subprocess
        $wordFiles = @($filePaths | Where-Object {
            $ext = [System.IO.Path]::GetExtension($_).ToLower()
            $ext -eq ".doc" -or $ext -eq ".docx"
        })

        $convertedPdfs = @{}
        if ($wordFiles.Count -gt 0) {
            $statusLabel.Text = "Converting Word documents..."
            $form.Refresh()
            $pdfResults = @(Convert-WordFilesToPdf $wordFiles)
            # Map each input Word file to its output PDF by matching filenames
            for ($i = 0; $i -lt $wordFiles.Count; $i++) {
                $convertedPdfs[$wordFiles[$i]] = $pdfResults[$i]
                $tempFiles.Add($pdfResults[$i]) | Out-Null
            }
        }

        $pdfFiles = @()
        foreach ($file in $filePaths) {
            $ext = [System.IO.Path]::GetExtension($file).ToLower()
            if ($ext -eq ".doc" -or $ext -eq ".docx") {
                $pdfFiles += $convertedPdfs[$file]
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

        # Clean up temp files before closing
        foreach ($t in $tempFiles) {
            Remove-Item $t -Force -ErrorAction SilentlyContinue
        }
        $form.Close()
        return
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
        $clearBtn.Enabled = $true
        $listBox.Enabled  = $true
    }
})

# ============================================================
# SHOW
# ============================================================
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
$form.Dispose()
