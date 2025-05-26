
<#
.SYNOPSIS
    Automated installation script for Windows Terminal dependencies, Winget, and Chocolatey
.DESCRIPTION
    This script downloads and installs Windows Terminal dependencies, Winget, and Chocolatey
    with proper verification at each step.
.NOTES
    File Name      : Install-WingetDependencies.ps1
    Prerequisite   : PowerShell 5.1 or later (run as Administrator)
#>

#Requires -RunAsAdministrator

# Clear the screen
Clear-Host

# Set execution policy to allow script execution
Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# Create Windows Terminal directory if it doesn't exist
$terminalPath = "C:\Windows\Windows Terminal"
if (-not (Test-Path -Path $terminalPath)) {
    try {
        New-Item -Path $terminalPath -ItemType Directory -Force | Out-Null
        Write-Host "Created directory: $terminalPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create directory: $terminalPath" -ForegroundColor Red
        exit 1
    }
}

# Define the files to download
$filesToDownload = @(
    @{
        Url = "https://scripts.kalixhosting.com/winget22/required/Microsoft.UI.Xaml.2.8_8.2501.31001.0_x64.appx"
        FileName = "Microsoft.UI.Xaml.2.8_8.2501.31001.0_x64.appx"
    },
    @{
        Url = "https://scripts.kalixhosting.com/winget22/required/Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64.appx"
        FileName = "Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64.appx"
    },
    @{
        Url = "https://scripts.kalixhosting.com/winget22/required/e53e159d00e04f729cc2180cffd1c02e_License1.xml"
        FileName = "e53e159d00e04f729cc2180cffd1c02e_License1.xml"
    },
    @{
        Url = "https://scripts.kalixhosting.com/winget22/required/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
        FileName = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    }
)

# Download files with verification
foreach ($file in $filesToDownload) {
    $outputPath = Join-Path -Path $terminalPath -ChildPath $file.FileName
    
    # Skip if file already exists
    if (Test-Path -Path $outputPath) {
        Write-Host "File already exists: $($file.FileName)" -ForegroundColor Yellow
        continue
    }

    try {
        Write-Host "Downloading $($file.FileName)..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $file.Url -OutFile $outputPath -UseBasicParsing
        Write-Host "Download completed: $($file.FileName)" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to download $($file.FileName)" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        exit 1
    }
}

# Verify all files were downloaded
$missingFiles = @()
foreach ($file in $filesToDownload) {
    $filePath = Join-Path -Path $terminalPath -ChildPath $file.FileName
    if (-not (Test-Path -Path $filePath)) {
        $missingFiles += $file.FileName
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "Missing files:" -ForegroundColor Red
    $missingFiles | ForEach-Object { Write-Host "- $_" -ForegroundColor Red }
    exit 1
}

# Define paths for installation
$destinationFolderPath = "C:\Windows\Windows Terminal"
$appx1Filename = "Microsoft.UI.Xaml.2.8_8.2501.31001.0_x64.appx"
$appx2Filename = "Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64.appx"
$msixwingetFilename = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$wingetLic = "e53e159d00e04f729cc2180cffd1c02e_License1.xml"

$appx1Path = Join-Path -Path $destinationFolderPath -ChildPath $appx1Filename
$appx2Path = Join-Path -Path $destinationFolderPath -ChildPath $appx2Filename
$msixwingetPath = Join-Path -Path $destinationFolderPath -ChildPath $msixwingetFilename
$wingetLicPath = Join-Path -Path $destinationFolderPath -ChildPath $wingetLic

# Install Appx packages
try {
    Write-Host "Installing Microsoft.UI.Xaml package..." -ForegroundColor Cyan
    Add-AppxPackage -Path $appx1Path -ErrorAction Stop
    Write-Host "Microsoft.UI.Xaml package installed successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to install Microsoft.UI.Xaml package" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

try {
    Write-Host "Installing Microsoft.VCLibs package..." -ForegroundColor Cyan
    Add-AppxPackage -Path $appx2Path -ErrorAction Stop
    Write-Host "Microsoft.VCLibs package installed successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to install Microsoft.VCLibs package" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

# Provision Winget package
try {
    Write-Host "Provisioning Winget package..." -ForegroundColor Cyan
    Add-AppxProvisionedPackage -Online -PackagePath $msixwingetPath -LicensePath $wingetLicPath -ErrorAction Stop
    Write-Host "Winget provisioned successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to provision Winget package" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

# Install Chocolatey
try {
    if (-not (Test-Path "$env:ProgramData\chocolatey\choco.exe")) {
        Write-Host "Installing Chocolatey package manager..." -ForegroundColor Cyan
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        Write-Host "Chocolatey installed successfully" -ForegroundColor Green
    } else {
        Write-Host "Chocolatey is already installed" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Failed to install Chocolatey" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

# Verify Winget installation
try {
    Write-Host "Verifying Winget installation..." -ForegroundColor Cyan
    $wingetCheck = Get-Command winget -ErrorAction Stop
    Write-Host "Winget is installed successfully!" -ForegroundColor Green
    
    # Run winget upgrade if Winget is available
    try {
        Write-Host "Running winget upgrade --all..." -ForegroundColor Cyan
        winget upgrade --all
        Write-Host "Winget upgrade completed successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Winget upgrade failed" -ForegroundColor Yellow
        Write-Host "Error: $_" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Winget is not installed correctly" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

Write-Host "All installations completed successfully!" -ForegroundColor Green


