<#
.SYNOPSIS
    LumberTools Setup - Creates Start Menu shortcuts for all tools.
.DESCRIPTION
    Scans tools/*/tool.json for tool manifests and creates Start Menu
    shortcuts under a "LumberTools" folder. Idempotent — safe to re-run
    after adding or removing tools.
#>

#Requires -Version 5.1

$lumberRoot = $PSScriptRoot
$toolsDir   = Join-Path $lumberRoot "tools"

# Use All Users Start Menu if running as admin, otherwise Current User
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    $startMenuBase = [Environment]::GetFolderPath("CommonPrograms")
} else {
    $startMenuBase = [Environment]::GetFolderPath("Programs")
    Write-Host "Note: Running without admin — shortcuts will be for current user only."
}

$lumberMenu = Join-Path $startMenuBase "LumberTools"
if (-not (Test-Path $lumberMenu)) {
    New-Item -ItemType Directory -Path $lumberMenu -Force | Out-Null
}

$shell = New-Object -ComObject WScript.Shell
$createdShortcuts = @()

# Scan for tool manifests and create shortcuts
$manifests = Get-ChildItem -Path $toolsDir -Filter "tool.json" -Recurse -ErrorAction SilentlyContinue
foreach ($manifestFile in $manifests) {
    $manifest = Get-Content $manifestFile.FullName -Raw | ConvertFrom-Json
    $toolDir  = $manifestFile.DirectoryName
    $launcherPath = Join-Path $toolDir $manifest.launcher

    if (-not (Test-Path $launcherPath)) {
        Write-Warning "Launcher not found for '$($manifest.displayName)': $launcherPath"
        continue
    }

    $lnkName  = "$($manifest.displayName).lnk"
    $lnkPath  = Join-Path $lumberMenu $lnkName
    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath       = $launcherPath
    $shortcut.WorkingDirectory = $toolDir
    $shortcut.Description      = $manifest.description
    if ($manifest.icon) {
        $iconPath = Join-Path $toolDir $manifest.icon
        if (Test-Path $iconPath) {
            $shortcut.IconLocation = $iconPath
        }
    }
    $shortcut.Save()

    $createdShortcuts += $lnkName
    Write-Host "  + $($manifest.displayName)"
}

# Clean up shortcuts for tools that no longer exist
$existingLinks = Get-ChildItem -Path $lumberMenu -Filter "*.lnk" -ErrorAction SilentlyContinue
foreach ($link in $existingLinks) {
    if ($link.Name -notin $createdShortcuts) {
        Remove-Item $link.FullName -Force
        Write-Host "  - Removed: $($link.BaseName)"
    }
}

# Remove the LumberTools folder if it's now empty
if ((Get-ChildItem -Path $lumberMenu -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
    Remove-Item $lumberMenu -Force -ErrorAction SilentlyContinue
}

Write-Host "`nSetup complete. Tools are available in Start Menu > LumberTools."
