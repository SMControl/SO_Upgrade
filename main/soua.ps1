Write-Host "SOUpgradeAssistant.ps1 - Version 1.218" # previous working to this was 1.217
# This script automates the upgrade process for Smart Office (SO) software.

# Global Script Configuration
$TOTAL_PARTS = 15
$startTime = Get-Date
$workingDir = "C:\winsm"
# Stores initial running status of services/processes for later reversion (Parts 6, 7, 13, 14)
$processStates = @{}
$SO_SERVICE = "srvSOLiveSales"
$SO_PROCESSES = @("Sm32Main", "Sm32")

# ==================================
# Utility Functions
# ==================================

# Function to display part intro and progress
function Start-Part {
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][int]$Current,
        [Parameter(Mandatory=$true)][string]$PartVersion
    )
    Clear-Host
    Write-Host "SO Upgrade Assistant - Version 1.218" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------------------------"
    #Write-Host ""
    
    $progressBarWidth = 30
    $filled = [int]($progressBarWidth * $Current / $TOTAL_PARTS)
    $empty = $progressBarWidth - $filled
    $progress = "[" + ("█" * $filled) + ("_" * $empty) + "]"
    
    Write-Host "[Part $Current/$TOTAL_PARTS] $Title" -ForegroundColor Cyan
    Write-Host "$progress" -ForegroundColor Cyan
    Write-Host ""
    # Removed: Write-Host "# PartVersion-$PartVersion"
}

# Function to test for administrator privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to stop a process by name and record its original state
# Only outputs when a process is detected and being stopped.
function Stop-ProcessIfRunning {
    param(
        [Parameter(Mandatory=$true)][string]$ProcessName
    )
    $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    $wasRunning = $false
    
    if ($process) {
        $wasRunning = $true
        Write-Host "Process '$ProcessName' detected. Stopping..." -ForegroundColor Yellow
        Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
        # Success message removed for conciseness
    } 
    # "is not running" message removed for conciseness
    
    $processStates[$ProcessName] = $wasRunning
    return $wasRunning
}

# ==================================
# Pre-Script Setup
# ==================================

# Setup working directory
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
# Part 1 - Check for Admin Rights & Install SO Setup Get
# PartVersion-1.02
# - Consolidated module installation with admin check, removed pre-action Write-Host.
#LOCK=ON
# ==================================
Start-Part -Title "System Pre-Checks & Setup Tool Installation" -Current 1 -PartVersion "1.02"

if (-not (Test-Admin)) {
    Write-Host "Error= Administrator rights required to run this script. Exiting." -ForegroundColor Red
    pause
    exit
}

# Make sure SO Setup Get is installed. (NOTE: Uses potentially risky 'iex' pattern)
irm https://raw.githubusercontent.com/SMControl/SM_Tasks/refs/heads/main/tasks/task_SO%20Setup%20Get.ps1 | iex
Write-Host "SO Setup Get module checked/installed." -ForegroundColor Green

# ==================================
# Part 3 - SO_UC.exe // calling module_soget
# PartVersion-1.02
# - Removed pre-action Write-Host.
#LOCK=ON
# ==================================
Start-Part -Title "Checking for Setup Files (module_soget)" -Current 3 -PartVersion "1.02"

# Define the URL for the SO Get module
$sogetScriptURL = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_soget.ps1"
try {
    Invoke-Expression (Invoke-RestMethod -Uri $sogetScriptURL -ErrorAction Stop)
    Write-Host "module_soget.ps1 executed successfully." -ForegroundColor Green
} catch {
    Write-Host "Error executing module_soget.ps1 $($_.Exception.Message)" -ForegroundColor Red
    exit
}


# ==================================
# Part 4 - Firebird Installation // calling module_firebird
# PartVersion-1.03
# - Shortened pre-action message, removed "already installed" message.
#LOCK=ON
# ==================================
Start-Part -Title "Checking for Firebird installation" -Current 4 -PartVersion "1.03"

$firebirdDir = "C:\Program Files (x86)\Firebird"
if (-not (Test-Path $firebirdDir)) {
    Write-Host "Installing Firebird..." -ForegroundColor Yellow
    $firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_firebird.ps1"
    try {
        Invoke-Expression (Invoke-RestMethod -Uri $firebirdInstallerURL -ErrorAction Stop)
        Write-Host "Firebird installation script executed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error executing Firebird installation script: $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
} 
# Removed: Firebird is already installed message.

# ==================================
# Part 5 - Stop SMUpdates if Running (Background Monitor)
# PartVersion-1.03
# - Removed redundant starting message.
#LOCK=ON
# ==================================
Start-Part -Title "Monitoring and Stopping SMUpdates (Background)" -Current 5 -PartVersion "1.03"

$monitorJob = Start-Job -ScriptBlock {
    function Monitor-SmUpdates {
        while ($true) {
            $smUpdatesProcess = Get-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
            if ($smUpdatesProcess) {
                Stop-Process -Name "SMUpdates" -Force -ErrorAction SilentlyContinue
                # Note: Console output from background jobs is not immediately visible.
            }
            Start-Sleep -Seconds 2
        }
    }
    Monitor-SmUpdates
}

# ==================================
# Part 6 - Manage SO Live Sales Service
# PartVersion-1.03
# - Removed "stopped successfully," "not running," and "not found" messages.
#LOCK=ON
# ==================================
Start-Part -Title "Stopping SO Live Sales Service" -Current 6 -PartVersion "1.03"

try {
    $service = Get-Service -Name $SO_SERVICE -ErrorAction SilentlyContinue
    if ($service) {
        $processStates[$SO_SERVICE] = ($service.Status -eq 'Running')
        if ($processStates[$SO_SERVICE]) {
            Write-Host "Service '$SO_SERVICE' is running. Stopping..." -ForegroundColor Yellow
            Stop-Service -Name $SO_SERVICE -Force -ErrorAction Stop
        } 
        # Removed: "stopped successfully," "not running," and "not found" messages.
    } 
} catch {
    Write-Host "Error managing service '$SO_SERVICE': $($_.Exception.Message)" -ForegroundColor Red
}

# ==================================
# Part 7 - Manage PDTWiFi Processes
# PartVersion-1.04
# - Messages handled by modified Stop-ProcessIfRunning function (only output on detected/stopping).
#LOCK=ON
# ==================================
Start-Part -Title "Managing PDTWiFi processes" -Current 7 -PartVersion "1.04"

Stop-ProcessIfRunning -ProcessName "PDTWiFi"
Stop-ProcessIfRunning -ProcessName "PDTWiFi64"

# ==================================
# Part 8 - Make Sure SO is closed & Wait for Single Instance of Firebird.exe
# PartVersion-1.02
# - Removed "SO is not running" and "Firebird is running" success messages.
#LOCK=ON
# ==================================
Start-Part -Title "Wait SO Closed & Single Firebird process..." -Current 8 -PartVersion "1.02"

# Make sure SO is closed
foreach ($process in $SO_PROCESSES) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Smart Office is open. Please close it to continue." -ForegroundColor Red
        while (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            Start-Sleep -Seconds 3
        }
        Write-Host "Smart Office is now closed." -ForegroundColor Green
    } 
    # Removed: Smart Office is not running message.
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
    # Removed: Only one instance of 'firebird.exe' is running success message.
}
WaitForSingleFirebirdInstance

# ==================================
# Part 9 - Launch Setup
# PartVersion 1.09
# - Updated column headers in the setup selection table based on user request.
#LOCK=ON
# ==================================
Start-Part -Title "Launching SO setup..." -Current 9 -PartVersion "1.09"

$setupExes = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "*.exe"
if ($setupExes.Count -eq 0) {
    Write-Host "Error No executable (.exe) found in 'C:\winsm\SmartOffice_Installer'." -ForegroundColor Red
    exit
} elseif ($setupExes.Count -eq 1) {
    $selectedExe = $setupExes[0]
    Write-Host "Found setup $($selectedExe.Name)" -ForegroundColor Green
} else {
    $setupExes = $setupExes | Sort-Object { [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int] }
    Write-Host "`nPlease select the setup to run`n" -ForegroundColor Yellow
    
    # Updated headers: No. -> #, Executable Name -> Name, Date Modified -> Downloaded
    # Using widths: # (5), Name (25), Downloaded (20), Version (10)
    Write-Host ("{0,-5} {1,-25} {2,-20} {3,-10}" -f "#", "Name", "Downloaded", "Version") -ForegroundColor White
    Write-Host ("{0,-5} {1,-25} {2,-20} {3,-10}" -f "---", "-------------------------", "--------------------", "----------") -ForegroundColor Gray
    
    for ($i = 0; $i -lt $setupExes.Count; $i++) {
        $exe = $setupExes[$i]
        $dateModified = $exe.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        $versionType = if ($i -eq 0) { "Current" } else { "Next" }
        $color = if ($i -eq 0) { "Green" } else { "Yellow" }
        
        Write-Host ("{0,-5} {1,-25} {2,-20} {3,-10}" -f ($i + 1), $exe.Name, $dateModified, $versionType) -ForegroundColor $color
    }
    Write-Host "`nEnter the number of your selection (or press Enter to cancel):" -ForegroundColor Cyan
    $selection = Read-Host "Selection"
    if ([string]::IsNullOrWhiteSpace($selection)) {
        Write-Host "Operation cancelled. Script execution stopped for this part." -ForegroundColor Red
        return
    }
    if ($selection -match '^\d+$' -and $selection -ge 1 -and $selection -le $setupExes.Count) {
        $selectedExe = $setupExes[$selection - 1]
        Write-Host "Selected setup executable $($selectedExe.Name)" -ForegroundColor Green
    } else {
        Write-Host "Invalid selection. Exiting." -ForegroundColor Red
        exit
    }
}

try {
    Start-Process -FilePath $selectedExe.FullName -Wait
    Write-Host "Setup executable finished." -ForegroundColor Green
} catch {
    Write-Host "Error starting setup executable $_" -ForegroundColor Red
    exit
}

# ==================================
# Part 10 - Wait for User Confirmation
# PartVersion-1.03
# - Removed SMUpdates stopped message outside of Part 5.
#LOCK=ON
# ==================================
Start-Part -Title "Post Upgrade Confirmation" -Current 10 -PartVersion "1.03"

Stop-Job -Job $monitorJob
Remove-Job -Job $monitorJob
# Removed: Write-Host "SMUpdates monitoring stopped." -ForegroundColor Green

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("Please ensure the upgrade is complete and Smart Office is closed before clicking OK.", "SO Post Upgrade Confirmation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Re-check for Running SO Processes
foreach ($process in $SO_PROCESSES) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Smart Office is still running. Please close it and press Enter to continue..." -ForegroundColor Red
        Read-Host | Out-Null
    }
}
Write-Host "Smart Office processes confirmed closed." -ForegroundColor Green

# ==================================
# Part 11 - Set Permissions for SM Folder
# PartVersion-1.13
# - No changes to messaging.
#LOCK=ON
# ==================================
Start-Part -Title "Setting permissions for Stationmaster folder" -Current 11 -PartVersion "1.13"

Write-Host "Executing permissions update (may take time)..." -ForegroundColor Yellow
try {
    # Suppress all output (*>$null) and check exit code ($LASTEXITCODE)
    & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C *>$null
    if ($LASTEXITCODE -ne 0) { throw "icacls failed with exit code $LASTEXITCODE." }
    Write-Host "Permissions for StationMaster folder set successfully." -ForegroundColor Green
} catch {
    Write-Host "Error setting permissions for SM folder: $($_.Exception.Message)" -ForegroundColor Red
}


# ==================================
# Part 12 - Set Permissions for Firebird Folder
# PartVersion-1.03
# - No changes to messaging.
#LOCK=ON
# ==================================
Start-Part -Title "Setting permissions for Firebird folder" -Current 12 -PartVersion "1.03"

Write-Host "Executing permissions update (may take time)..." -ForegroundColor Yellow
try {
    # Suppress all output (*>$null) and check exit code ($LASTEXITCODE)
    & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C *>$null
    if ($LASTEXITCODE -ne 0) { throw "icacls failed with exit code $LASTEXITCODE." }
    Write-Host "Permissions for Firebird folder set successfully." -ForegroundColor Green
} catch {
    Write-Host "Error setting permissions for Firebird folder: $($_.Exception.Message)" -ForegroundColor Red
}

# ==================================
# Part 13 - Revert SO Live Sales Service
# PartVersion-1.08
# - No changes to messaging (messages here are crucial for retries/manual intervention).
#LOCK=ON
# ==================================
Start-Part -Title "Restarting SO Live Sales Service" -Current 13 -PartVersion "1.08"

if ($processStates[$SO_SERVICE]) {
    Write-Host "Service '$SO_SERVICE' was running. Restarting..." -ForegroundColor Yellow
    $retryCount = 0
    $maxRetries = 3
    $retryIntervalSeconds = 5

    while ($retryCount -lt $maxRetries) {
        Write-Host "Attempting to start service '$SO_SERVICE' (Attempt $($retryCount + 1) of $maxRetries)..." -ForegroundColor Yellow
        Start-Service -Name $SO_SERVICE -ErrorAction SilentlyContinue
        if ((Get-Service -Name $SO_SERVICE).Status -eq "Running") {
            Write-Host "Service '$SO_SERVICE' is now running." -ForegroundColor Green
            break
        }
        Start-Sleep -Seconds $retryIntervalSeconds
        $retryCount++
    }
    
    # Final check and user manual intervention prompt
    if ((Get-Service -Name $SO_SERVICE).Status -ne "Running") {
        Write-Host "Failed to automatically start service '$SO_SERVICE'. Please manually start the service now." -ForegroundColor Red
        while ((Get-Service -Name $SO_SERVICE).Status -ne "Running") {
            Write-Host "Waiting for service to run. Checking again in 3 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
        }
        Write-Host "Service '$SO_SERVICE' is now running. Continuing..." -ForegroundColor Green
    } else {
        Write-Host "'$SO_SERVICE' service confirmed to be running." -ForegroundColor Green
    }
} else {
    Write-Host "Service '$SO_SERVICE' was not running before, no action taken." -ForegroundColor Yellow
}

# ==================================
# Part 14 - Revert PDTWiFi Processes
# PartVersion 1.03
# - No changes to messaging.
#LOCK=ON
# ==================================
Start-Part -Title "Reverting PDTWiFi processes" -Current 14 -PartVersion "1.03"

$PDTWiFi_PATH = "C:\Program Files (x86)\StationMaster"

if ($processStates["PDTWiFi"]) {
    try {
        Start-Process "$PDTWiFi_PATH\PDTWiFi.exe" -ErrorAction Stop
        Write-Host "PDTWiFi started." -ForegroundColor Green
    } catch {
        Write-Host "Error starting PDTWiFi $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "PDTWiFi was not running, no action taken." -ForegroundColor Yellow
}

if ($processStates["PDTWiFi64"]) {
    try {
        Start-Process "$PDTWiFi_PATH\PDTWiFi64.exe" -ErrorAction Stop
        Write-Host "PDTWiFi64 started." -ForegroundColor Green
    } catch {
        Write-Host "Error starting PDTWiFi64 $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "PDTWiFi64 was not running, no action taken." -ForegroundColor Yellow
}

# ==================================
# Part 15 - Clean up and Finish Script
# PartVersion-1.10
# - Reverted status table to only show Item and Current Status for conciseness.
#LOCK=ON
# ==================================
Start-Part -Title "Clean up and finish" -Current 15 -PartVersion "1.10"

# Get Current Status of services and processes
$liveSalesService = Get-Service -Name $SO_SERVICE -ErrorAction SilentlyContinue
$liveSalesServiceStatus = if ($liveSalesService) { $liveSalesService.Status } else { "Not Installed" }
$pdtWifiStatus = if (Get-Process -Name "PDTWiFi" -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }
$pdtWifi64Status = if (Get-Process -Name "PDTWiFi64" -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }

# Function to determine status color (kept for conciseness in the final output)
function Get-StatusColor ($status) {
    if ($status -eq "Running") { return "Green" }
    if ($status -eq "Not Installed") { return "Gray" }
    return "Yellow"
}

Write-Host " "
Write-Host "Process Status:" -ForegroundColor Yellow
Write-Host "------------------------------------------------" -ForegroundColor Yellow
Write-Host ("{0,-25} {1,-15}" -f "Item", "Current Status") -ForegroundColor White
Write-Host ("{0,-25} {1,-15}" -f "-------------------------", "---------------") -ForegroundColor Gray

Write-Host ("{0,-25} {1,-15}" -f "SO Live Sales Service", $liveSalesServiceStatus) -ForegroundColor (Get-StatusColor $liveSalesServiceStatus)
Write-Host ("{0,-25} {1,-15}" -f "PDTWiFi.exe", $pdtWifiStatus) -ForegroundColor (Get-StatusColor $pdtWifiStatus)
Write-Host ("{0,-25} {1,-15}" -f "PDTWiFi64.exe", $pdtWifi64Status) -ForegroundColor (Get-StatusColor $pdtWifi64Status)

Write-Host "------------------------------------------------" -ForegroundColor Yellow

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds

Write-Host " "
Write-Host "Completed in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
Write-Host "Consider if you need to Reboot at this stage." -ForegroundColor Yellow
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
