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

# Resolve winget (not in PATH when running as SYSTEM)
$winget = Get-Command winget -ErrorAction SilentlyContinue
if (-not $winget) {
    $wingetPath = Resolve-Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1
    if ($wingetPath) { $winget = $wingetPath.Path }
}
if (-not $winget) {
    Write-Error "WinGet is not installed. Cannot install Git."
    exit 1
}

# Ensure git is available, install via winget if missing
$git = Get-Command git -ErrorAction SilentlyContinue
if (-not $git) {
    Write-Host "Git not found. Installing via winget..."
    & $winget install --id Git.Git --source winget --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Git installation failed."
        exit 1
    }
    # Refresh PATH so git is available in this session
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath    = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path    = "$machinePath;$userPath"
    $git = Get-Command git -ErrorAction SilentlyContinue
    if (-not $git) {
        Write-Error "Git was installed but could not be found in PATH."
        exit 1
    }
    Write-Host "Git installed successfully."
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
