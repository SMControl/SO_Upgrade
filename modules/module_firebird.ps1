# InstallFirebirdWithDownload.ps1
# This script checks for administrative privileges, verifies the existence of a directory, downloads the Firebird installer, installs Firebird if not already installed, modifies the firebird.conf file, adjusts permissions, starts the Firebird service, and cleans up temporary files.
# ---
# version 1.06
# Summary of Changes and fixes since last version
# - Added timer to measure and display the duration of the script execution

# Function to write output in green
function Write-Success {
    param (
        [string]$message
    )
    Write-Host $message -ForegroundColor Green
}

# Function to write output in red
function Write-ErrorOutput {
    param (
        [string]$message
    )
    Write-Host $message -ForegroundColor Red
}

# Start timer
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Part 1 - Pre Install Check
# ------------------------------------------------

# Check if running as admin
Write-Output "Checking if running as administrator..."
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-ErrorOutput "Please run this script as an administrator."
    exit 1
}

# Check if the directory exists
Write-Output "Checking if the directory 'C:\Program Files (x86)\Firebird' exists..."
if (!(Test-Path "C:\Program Files (x86)\Firebird")) {
    Write-Output "The directory 'C:\Program Files (x86)\Firebird' does NOT exist. Proceeding with the installation..."

    # Part 2 - Download Firebird Installer
    # ------------------------------------------------

    Write-Output "Downloading Firebird Installer..."
    $installerUrl = "https://github.com/SMControl/SM_Firebird_Installer/raw/main/Firebird-4.0.1.exe"
    $installerPath = "$env:TEMP\Firebird-4.0.1.exe"
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    Write-Success "Firebird Installer downloaded successfully."

    # Part 3 - Installation of Firebird with scripted parameters
    # ------------------------------------------------

    Write-Output "Installing Firebird..."
    Start-Process -FilePath $installerPath -ArgumentList "/LANG=en", "/NORESTART", "/VERYSILENT", "/MERGETASKS=UseClassicServerTask,UseServiceTask,CopyFbClientAsGds32Task" -Wait
    Write-Success "Firebird installed successfully."
} else {
    Write-ErrorOutput "Firebird is already installed. Continuing with other tasks..."
}

# Part 4 - Modify firebird.conf
# ------------------------------------------------

Write-Output "Modifying firebird.conf..."
(Get-Content "C:\Program Files (x86)\Firebird\Firebird_4_0\firebird.conf") -replace '#DataTypeCompatibility.*', 'DataTypeCompatibility = 3.0' | Set-Content "C:\Program Files (x86)\Firebird\Firebird_4_0\firebird.conf"
Write-Success "firebird.conf modified successfully."

# Part 5 - Adjusting permissions
# ------------------------------------------------

Write-Output "Adjusting permissions..."
icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C | Out-Null
Write-Success "Permissions adjusted successfully."

# Part 6 - Start Firebird service
# ------------------------------------------------

Write-Output "Starting Firebird service..."
Start-Service -Name "FirebirdServerDefaultInstance"
Write-Success "Firebird service started successfully."

# Part 7 - Cleanup
# ------------------------------------------------

Write-Output "Cleaning up temporary files..."
Remove-Item $installerPath -ErrorAction SilentlyContinue
Write-Success "Temporary files cleaned up successfully."

# Part 8 - Installation Successful
# ------------------------------------------------

Write-Success "Firebird installation completed successfully."

# Stop timer
$stopwatch.Stop()
$elapsedTime = $stopwatch.Elapsed
Write-Output ("Script execution time: {0} minutes and {1} seconds." -f $elapsedTime.Minutes, $elapsedTime.Seconds)
