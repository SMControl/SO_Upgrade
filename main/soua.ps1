Write-Host "SOUpgradeAssistant.ps1 - Version 1.211" # previous working to this was 1.203
# This script automates the upgrade process for Smart Office (SO) software.
#
# Recent Changes:
# - Updated script to version 1.209.
# - Incorporated the latest Part 6 (stopping service) and Part 13 (starting service).
# - Set LOCK=ON for all parts.
# - Fixed the script duration calculation.
# - Improved comments and formatting.
# - Added a message to the user about potential delays in Part 11.
# - Added SO_UC.exe download
# - Changed SO_UC execution
# - Added check for SO_UC scheduled task and execute if not exists
# - Added logic to handle SO Live Sales service startup more robustly.
# - Changed Part 9 to prevent window closure on cancel.
# - Initialized $wasRunning variable for Part 6.
# - Uncommented feedback message in Part 15 for scheduled task check.

# Initialize script start time
$startTime = Get-Date

# Function to display the script's introduction
function Show-Intro {
    Write-Host "SO Upgrade Assistant - Version 1.209" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------------------------"
    Write-Host ""
}

# Set the working directory
$workingDir = "C:\winsm"
if (-not (Test-Path $workingDir -PathType Container)) {
    try {
        New-Item -Path $workingDir -ItemType Directory -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Error= Unable to create directory $workingDir" -ForegroundColor Red
        exit
    }
}
Set-Location -Path $workingDir

# ==================================
# Part 1 - Check for Admin Rights
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 1/15] System Pre-Checks" -ForegroundColor Cyan
Write-Host "[______________________________]" -ForegroundColor Cyan
Write-Host ""

# Function to test for administrator privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Check if the script is running with administrator rights
if (-not (Test-Admin)) {
    Write-Host "Error= Administrator rights required to run this script. Exiting." -ForegroundColor Red
    pause
    exit
}

#outdated-so disabled
# Download the SO_UC.exe and save it to C:\winsm
#$soucExeUrl = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe?raw=true"
#$soucExeDestinationPath = "C:\winsm\SO_UC.exe"
#if (-Not (Test-Path $soucExeDestinationPath)) {
#    Write-Host "Downloading SO_UC.exe..." -ForegroundColor Yellow
#    try {
#        Invoke-WebRequest -Uri $soucExeUrl -OutFile $soucExeDestinationPath -ErrorAction Stop
#        Write-Host "SO_UC.exe downloaded successfully to $soucExeDestinationPath." -ForegroundColor Green
#    } catch {
#        Write-Host "Error downloading SO_UC.exe $($_.Exception.Message)" -ForegroundColor Red
#        exit
#    }
#} else {
#    Write-Host "SO_UC.exe already exists at $soucExeDestinationPath. Skipping download." -ForegroundColor Yellow
#}

# ==================================
# Part 3 - SO_UC.exe // calling module_soget
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 3/15] Checking for Setup Files. Please Wait." -ForegroundColor Cyan
Write-Host "[██████________________________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
# Define the URL for the SO Get module
$sogetScriptURL = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_soget.ps1"
# Run module_soget directly in the current shell
try {
    Write-Host "Invoking module_soget.ps1..." -ForegroundColor Yellow
    Invoke-Expression (Invoke-RestMethod -Uri $sogetScriptURL -ErrorAction Stop)
    Write-Host "module_soget.ps1 executed successfully." -ForegroundColor Green
} catch {
    Write-Host "Error executing module_soget.ps1 $($_.Exception.Message)" -ForegroundColor Red
    exit
}


# ==================================
# Part 4 - Firebird Installation // calling module_firebird
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 4/15] Checking for Firebird installation" -ForegroundColor Cyan
Write-Host "[████████______________________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
# Define the URL for the Firebird installation script
$firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_firebird.ps1"
# Check if Firebird is installed
$firebirdDir = "C:\Program Files (x86)\Firebird"
if (-not (Test-Path $firebirdDir)) {
    Write-Host "Firebird not found. Installing Firebird..." -ForegroundColor Yellow
    try {
        Invoke-Expression (Invoke-RestMethod -Uri $firebirdInstallerURL -ErrorAction Stop)
        Write-Host "Firebird installation script executed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error executing Firebird installation script: $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
} else {
    Write-Host "Firebird is already installed." -ForegroundColor Green
}

# ==================================
# Part 5 - Stop SMUpdates if Running
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 5/15] Monitoring and Stopping SMUpdates" -ForegroundColor Cyan
Write-Host "[██████████____________________]" -ForegroundColor Cyan
Write-Host ""

$monitorJob = Start-Job -ScriptBlock {
    function Monitor-SmUpdates {
        while ($true) {
            $smUpdatesProcess = Get-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
            if ($smUpdatesProcess) {
                Write-Host "SMUpdates process detected. Attempting to stop..." -ForegroundColor Yellow
                Stop-Process -Name "SMUpdates" -Force -ErrorAction SilentlyContinue
                Write-Host "SMUpdates process stopped." -ForegroundColor Green
            }
            Start-Sleep -Seconds 2
        }
    }
    
    Monitor-SmUpdates
}
Write-Host "Monitoring for SMUpdates process in the background..." -ForegroundColor Yellow

# ==================================
# Part 6 - Manage SO Live Sales Service
# PartVersion-1.01
# - Initialized $wasRunning variable.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 6/15] Managing SO Live Sales Service" -ForegroundColor Cyan
Write-Host "[████████████__________________]" -ForegroundColor Cyan
Write-Host ""

$ServiceName = "srvSOLiveSales"
$wasRunning = $false # Initialize the variable

try {
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        if ($service.Status -eq 'Running') {
            $wasRunning = $true
            Write-Host "Service '$ServiceName' is running. Stopping..." -ForegroundColor Yellow
            Stop-Service -Name $ServiceName -Force -ErrorAction Stop
            Write-Host "Service '$ServiceName' stopped successfully." -ForegroundColor Green
        } else {
            Write-Host "Service '$ServiceName' is not running." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Service '$ServiceName' not found." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error managing service '$ServiceName': $($_.Exception.Message)" -ForegroundColor Red
}

# ==================================
# Part 7 - Manage PDTWiFi Processes
# PartVersion-1.02
# - Total redo
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 7/15] Managing PDTWiFi processes" -ForegroundColor Cyan
Write-Host "[██████████████________________]" -ForegroundColor Cyan
Write-Host ""

# Initialize the process states
$PDTWiFiStates = @{}
# Section A - PDTWiFi
$PDTWiFi = "PDTWiFi"
$pdtWiFiProcess = Get-Process -Name $PDTWiFi -ErrorAction SilentlyContinue
if ($pdtWiFiProcess) {
    $PDTWiFiStates[$PDTWiFi] = "Running"
    Stop-Process -Name $PDTWiFi -Force -ErrorAction SilentlyContinue
    Write-Host "$PDTWiFi stopped." -ForegroundColor Green
} else {
    $PDTWiFiStates[$PDTWiFi] = "Not running"
    Write-Host "$PDTWiFi is not running." -ForegroundColor Yellow
}
# Section B - PDTWiFi64
$PDTWiFi64 = "PDTWiFi64"
$pdtWiFi64Process = Get-Process -Name $PDTWiFi64 -ErrorAction SilentlyContinue
if ($pdtWiFi64Process) {
    $PDTWiFiStates[$PDTWiFi64] = "Running"
    Stop-Process -Name $PDTWiFi64 -Force -ErrorAction SilentlyContinue
    Write-Host "$PDTWiFi64 stopped." -ForegroundColor Green
} else {
    $PDTWiFiStates[$PDTWiFi64] = "Not running"
    Write-Host "$PDTWiFi64 is not running." -ForegroundColor Yellow
}

# ==================================
# Part 8 - Make Sure SO is closed & Wait for Single Instance of Firebird.exe
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 8/15] Wait SO Closed & Single Firebird process..." -ForegroundColor Cyan
Write-Host "[████████████████______________]" -ForegroundColor Cyan
Write-Host ""

# Make sure SO is closed
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    # Check if the process is running
    $processRunning = Get-Process -Name $process -ErrorAction SilentlyContinue
    if ($processRunning) {
        Write-Host "Smart Office is open. Please close it to continue." -ForegroundColor Red
        while (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            # Wait without spamming the terminal
            Start-Sleep -Seconds 3  # Check every 3 seconds
        }
        Write-Host "Smart Office is now closed." -ForegroundColor Green
    } else {
        Write-Host "Smart Office ($process) is not running." -ForegroundColor Yellow
    }
}

# Wait for single firebird instance
$setupDir = "$workingDir\SmartOffice_Installer"
if (-not (Test-Path $setupDir -PathType Container)) {
    Write-Host "Error Setup directory '$setupDir' does not exist." -ForegroundColor Red
    exit
}
function WaitForSingleFirebirdInstance {
    $firebirdProcesses = Get-Process -Name "firebird" -ErrorAction SilentlyContinue
    while ($firebirdProcesses.Count -gt 1) {
        Write-Host "`rWarning= Multiple instances of 'firebird.exe' are running. Currently: $($firebirdProcesses.Count) " -ForegroundColor Yellow -NoNewline
        Start-Sleep -Seconds 3
        $firebirdProcesses = Get-Process -Name "firebird" -ErrorAction SilentlyContinue
    }
    Write-Host "Only one instance of 'firebird.exe' is running." -ForegroundColor Green
}
WaitForSingleFirebirdInstance

# ==================================
# Part 9 - Launch Setup
# PartVersion 1.06
# - Improved terminal selection menu with colors and table formatting
# - Changed logic to prevent script termination on cancel
# - Script now stops, but does not exit, on cancel.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 9/15] Launching SO setup..." -ForegroundColor Cyan
Write-Host "[██████████████████____________]" -ForegroundColor Cyan
Write-Host ""

# Get all setup executables in the SmartOffice_Installer directory
$setupExes = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "*.exe"
if ($setupExes.Count -eq 0) {
    Write-Host "Error No executable (.exe) found in 'C:\winsm\SmartOffice_Installer'." -ForegroundColor Red
    exit
} elseif ($setupExes.Count -eq 1) {
    # Only one file found, proceed without asking the user
    $selectedExe = $setupExes[0]
    Write-Host "Found setup $($selectedExe.Name)" -ForegroundColor Green
} else {
    # Sort executables by version (numeric part in the name) ascending
    $setupExes = $setupExes | Sort-Object {
        [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int]
    }
    # Multiple setup files found, present a terminal selection menu
    Write-Host "`nPlease select the setup to run`n" -ForegroundColor Yellow
    Write-Host ("{0,-5} {1,-30} {2,-20} {3,-10}" -f "No.", "Executable Name", "Date Modified", "Version") -ForegroundColor White
    Write-Host ("{0,-5} {1,-30} {2,-20} {3,-10}" -f "---", "------------------------------", "--------------------", "-----------") -ForegroundColor Gray
    for ($i = 0; $i -lt $setupExes.Count; $i++) {
        $exe = $setupExes[$i]
        $dateModified = $exe.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        $versionType = if ($i -eq 0) { "Current" } else { "Next" }
        $color = if ($i -eq 0) { "Green" } else { "Yellow" }
        # Present the setup info with adjusted spacing for a more compact look
        Write-Host ("{0,-5} {1,-30} {2,-20} {3,-10}" -f ($i + 1), $exe.Name, $dateModified, $versionType) -ForegroundColor $color
    }
    Write-Host "`nEnter the number of your selection (or press Enter to cancel):" -ForegroundColor Cyan
    # Get user input
    $selection = Read-Host "Selection"
    # Check if the user wants to cancel
    if ([string]::IsNullOrWhiteSpace($selection)) {
        Write-Host "Operation cancelled. Script execution stopped for this part." -ForegroundColor Red
        # Note: As per the comment in the original script, this 'return' stops the current part
        # but allows the script to continue to subsequent parts if they are not dependent.
        # If the intention is to exit the entire script, 'exit' should be used instead.
        return 
    }
    # Validate the selection
    if ($selection -match '^\d+$' -and $selection -ge 1 -and $selection -le $setupExes.Count) {
        $selectedExe = $setupExes[$selection - 1]  # Convert to 0-based index
        Write-Host "Selected setup executable $($selectedExe.Name)" -ForegroundColor Green
    } else {
        Write-Host "Invalid selection. Exiting." -ForegroundColor Red
        exit # Exits the entire script
    }
}
# Launch the selected setup executable
try {
    Write-Host "Starting executable: $($selectedExe.Name) ..." -ForegroundColor Cyan
    Start-Process -FilePath $selectedExe.FullName -Wait
    Write-Host "Setup executable finished." -ForegroundColor Green
} catch {
    Write-Host "Error starting setup executable $_" -ForegroundColor Red
    exit
}

# ==================================
# Part 10 - Wait for User Confirmation
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 10/15] Post Upgrade" -ForegroundColor Cyan
Write-Host "[████████████████████__________]" -ForegroundColor Cyan
Write-Host ""

# Stop monitoring SMUpdates process
Stop-Job -Job $monitorJob
Remove-Job -Job $monitorJob
Write-Host "SMUpdates monitoring stopped." -ForegroundColor Green
Write-Host "Waiting for confirmation Upgrade is Complete..." -ForegroundColor Yellow
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("Please ensure the upgrade is complete and Smart Office is closed before clicking OK.", "SO Post Upgrade Confirmation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
# Check for Running SO Processes Again
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Smart Office is still running. Please close it and press Enter to continue..." -ForegroundColor Red
        Read-Host
    }
}
Write-Host "Smart Office processes confirmed closed." -ForegroundColor Green

# ==================================
# Part 11 - Set Permissions for SM Folder
# PartVersion-1.10
# - Reverted to the original icacls command for setting permissions.
# - Added a message to the user about potential delays.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 11/15] Setting permissions for Stationmaster folder. Please Wait..." -ForegroundColor Cyan
Write-Host "[██████████████████████________]" -ForegroundColor Cyan
Write-Host ""

Write-Host "Please wait, this task may take ~1-30+ minutes to complete depending on PC speed. Do not interrupt." -ForegroundColor Yellow
try {
    & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
    Write-Host "Permissions for StationMaster folder set successfully." -ForegroundColor Green
} catch {
    Write-Host "Error setting permissions for SM folder: $_" -ForegroundColor Red
}


# ==================================
# Part 12 - Set Permissions for Firebird Folder
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 12/15] Setting permissions for Firebird folder. Please Wait..." -ForegroundColor Cyan
Write-Host "[████████████████████████______]" -ForegroundColor Cyan
Write-Host ""

try {
    & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
    Write-Host "Permissions for Firebird folder set successfully." -ForegroundColor Green
} catch {
    Write-Host "Error setting permissions for Firebird folder: $_" -ForegroundColor Red
}

# ==================================
# Part 13 - Revert SO Live Sales Service
# PartVersion-1.06
# - Modified to only start the service if it was previously running.
# - Removed logic for setting service startup type.
# - Changed maxRetries to 3.
# - Added instruction for user to manually start live sales if needed.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 13/15] SO Live Sales" -ForegroundColor Cyan
Write-Host "[██████████████████████████____]" -ForegroundColor Cyan
Write-Host ""

if ($wasRunning) {
    Write-Host "Service '$ServiceName' was running before. Ensuring it restarts..." -ForegroundColor Yellow
    try {
        # Removed: Set-Service -Name $ServiceName -StartupType Automatic
        # As per new requirement, we are only starting the service if it was running.

        $retryCount = 0
        $maxRetries = 3 # Changed from 5 to 3
        $retryIntervalSeconds = 5

        while ($retryCount -lt $maxRetries) {
            Write-Host "Attempting to start service '$ServiceName' (Attempt $($retryCount + 1) of $maxRetries)..." -ForegroundColor Yellow
            Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
            if ((Get-Service -Name $ServiceName).Status -eq "Running") {
                Write-Host "Service '$ServiceName' is now running." -ForegroundColor Green
                break
            } else {
                Write-Host "Service '$ServiceName' is not running. Waiting $retryIntervalSeconds seconds before retrying..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryIntervalSeconds
                $retryCount++
            }
        }

        if ((Get-Service -Name $ServiceName).Status -ne "Running") {
            Write-Warning "Failed to automatically start service '$ServiceName' after $maxRetries attempts."
            Write-Host ""
            Write-Host "Please manually start the '$ServiceName' service now." -ForegroundColor Red
            Write-Host "The script will wait until the service is running..." -ForegroundColor Yellow

            while ((Get-Service -Name $ServiceName).Status -ne "Running") {
                Write-Host "Waiting for '$ServiceName' service to be running. Please manually start the service. Checking again in 3 seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
            }
            Write-Host "Service '$ServiceName' is now running. Continuing..." -ForegroundColor Green

        } else {
            Write-Host "'$ServiceName' service confirmed to be running." -ForegroundColor Green
        }

    } catch {
        Write-Host "Error encountered while starting service '$ServiceName' $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "Service '$ServiceName' was not running before, so no action needed." -ForegroundColor Yellow
}

# ==================================
# Part 14 - Revert PDTWiFi Processes
# PartVersion 1.02
# - Total re-do
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 14/15] Reverting PDTWiFi processes" -ForegroundColor Cyan
Write-Host "[████████████████████████████__]" -ForegroundColor Cyan
Write-Host ""

# Section A - Recall and Revert PDTWiFi
if ($PDTWiFiStates[$PDTWiFi] -eq "Running") {
    try {
        Start-Process "C:\Program Files (x86)\StationMaster\PDTWiFi.exe" -ErrorAction Stop
        Write-Host "$PDTWiFi started." -ForegroundColor Green
    } catch {
        Write-Host "Error starting $PDTWiFi $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "$PDTWiFi was not running, no action taken." -ForegroundColor Yellow
}
# Section B - Recall and Revert PDTWiFi64
if ($PDTWiFiStates[$PDTWiFi64] -eq "Running") {
    try {
        Start-Process "C:\Program Files (x86)\StationMaster\PDTWiFi64.exe" -ErrorAction Stop
        Write-Host "$PDTWiFi64 started." -ForegroundColor Green
    } catch {
        Write-Host "Error starting $PDTWiFi64 $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "$PDTWiFi64 was not running, no action taken." -ForegroundColor Yellow
}

# ==================================
# Part 15 - Clean up and Finish Script
# PartVersion-1.07
# - Fixed the script duration calculation to correctly capture start time.
# - Improved comments and formatting.
# - Uncommented feedback message for scheduled task check.
#LOCK=ON
# ==================================
Clear-Host
Show-Intro
Write-Host "[Part 15/15] Clean up and finish" -ForegroundColor Cyan
Write-Host "[██████████████████████████████]" -ForegroundColor Cyan

# Get status of services and processes
$liveSalesService = Get-Service -Name "srvSOLiveSales" -ErrorAction SilentlyContinue
if ($liveSalesService) {
    $liveSalesServiceStatus = $liveSalesService.Status
} else {
    $liveSalesServiceStatus = "Not Installed"
}
$pdtWifiStatus = if (Get-Process -Name "PDTWiFi" -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }
$pdtWifi64Status = if (Get-Process -Name "PDTWiFi64" -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }

# Output the status table
Write-Host " "
Write-Host "Process Status:" -ForegroundColor Yellow
Write-Host "------------------------------------------------" -ForegroundColor Yellow
Write-Host ("{0,-25} {1,-15}" -f "Item", "Status") -ForegroundColor Yellow
Write-Host ("{0,-25} {1,-15}" -f "-------------------------", "---------------") -ForegroundColor Yellow

# Function to determine status color
function Get-StatusColor ($status) {
    if ($status -eq "Running") {
        return "Green"
    } elseif ($status -eq "Not Installed") {
        return "Gray" # Use gray for "Not Installed" to distinguish from "Stopped"
    } else {
        return "Yellow"
    }
}

Write-Host ("{0,-25} {1,-15}" -f "SO Live Sales Service", $liveSalesServiceStatus) -ForegroundColor (Get-StatusColor $liveSalesServiceStatus)
Write-Host ("{0,-25} {1,-15}" -f "PDTWiFi.exe", $pdtWifiStatus) -ForegroundColor (Get-StatusColor $pdtWifiStatus)
Write-Host ("{0,-25} {1,-15}" -f "PDTWiFi64.exe", $pdtWifi64Status) -ForegroundColor (Get-StatusColor $pdtWifi64Status)

Write-Host "------------------------------------------------" -ForegroundColor Yellow

# outdated so disabled
# Run SO_UC.exe if its Task doesn't exist.
#$taskExists = Get-ScheduledTask -TaskName "SO InstallerUpdates" -ErrorAction SilentlyContinue
#if (-not $taskExists) {
#    Write-Host "Scheduled Task 'SO InstallerUpdates' does not exist. Running SO_UC.exe..." -ForegroundColor Yellow
#    try {
#        Start-Process -FilePath "C:\winsm\SO_UC.exe" -Wait -ErrorAction Stop
#        Write-Host "SO_UC.exe executed successfully." -ForegroundColor Green
#    } catch {
#        Write-Host "Error executing SO_UC.exe: $($_.Exception.Message)" -ForegroundColor Red
#    }
#} else {
#    Write-Host "Scheduled Task 'SO InstallerUpdates' already exists. Skipping SO_UC.exe execution." -ForegroundColor Yellow
#}

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds

Write-Host " "
Write-Host "Completed in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
Write-Host " "
Write-Host "Consider if you need to Reboot at this stage." -ForegroundColor Yellow
Write-Host " "
Write-Host "Press Enter to start Smart Office, '9' to reboot now, or any other key to exit."
$key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

if ($key.VirtualKeyCode -eq 13) { # Enter key
    Write-Host "Starting Smart Office..." -ForegroundColor Cyan
    try {
        Start-Process "C:\Program Files (x86)\StationMaster\Sm32.exe" -ErrorAction Stop
        Write-Host "Smart Office started successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error starting Smart Office: $($_.Exception.Message)" -ForegroundColor Red
    }
} elseif ($key.VirtualKeyCode -eq 57) { # 57 is the VirtualKeyCode for '9'
    Write-Host "Rebooting..." -ForegroundColor Cyan
    Restart-Computer -Force
} else {
    Write-Host "Exiting..." -ForegroundColor Cyan
}
