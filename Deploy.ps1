<#
.SYNOPSIS
    LumberTools installer/updater for RMM deployment.
.DESCRIPTION
    If C:\LumberTools doesn't exist, clones the repo and runs Setup.ps1.
    If it already exists, pulls latest changes and re-runs Setup.ps1.
    Designed to run as SYSTEM via Intune or RMM.
#>

#Requires -Version 5.1

$installDir = "C:\LumberTools"
$repoUrl    = "https://github.com/mattmaddux/LumberTools.git"

# Ensure git is available
$git = Get-Command git -ErrorAction SilentlyContinue
if (-not $git) {
    Write-Host "Git is not installed yet. Please deploy Git to this endpoint before running LumberTools Install."
    Write-Host "Install Git via Intune/RMM first, then re-run this script."
    exit 1
}

if (Test-Path (Join-Path $installDir ".git")) {
    # Already installed - pull latest
    Write-Host "Updating LumberTools..."
    git -C $installDir pull --ff-only
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Git pull failed."
        exit 1
    }
} else {
    # Fresh install
    Write-Host "Installing LumberTools..."
    if (Test-Path $installDir) {
        Remove-Item $installDir -Recurse -Force
    }
    git clone $repoUrl $installDir
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Git clone failed."
        exit 1
    }
}

# Run Setup.ps1 to create/update Start Menu shortcuts
$setupScript = Join-Path $installDir "Setup.ps1"
if (Test-Path $setupScript) {
    Write-Host "Running Setup.ps1..."
    & powershell.exe -ExecutionPolicy Bypass -File $setupScript
} else {
    Write-Warning "Setup.ps1 not found - shortcuts were not created."
}

Write-Host "Done."
