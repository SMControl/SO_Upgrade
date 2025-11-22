# ==================================================================================================
# Script: SOUpgradeAssistant_REBOOT.ps1
# Version: 2.140
# Description: Automates the upgrade process for Smart Office (SO) software with reboot-resume capability.
# ==================================================================================================
# Recent Changes:
# - Version 2.140: TRIGGER CHANGE TO ONLOGON
#   - Changed scheduled task trigger from ONSTART to ONLOGON
#   - More reliable - triggers when user logs in, not at system startup
#   - Uses domain\username for the logon trigger
# - Version 2.130: RUN DIRECTLY FROM GITHUB
#   - Scheduled task now runs script directly from GitHub (no local file)
#   - Uses Invoke-Expression with ScriptBlock to pass -Resume parameter
#   - Simpler, always uses latest version, no file management needed
# - Version 2.120: (deprecated approach)
# - Version 2.110: SCHEDULED TASK SIMPLIFICATION
#   - Replaced PowerShell cmdlets with schtasks.exe one-liner for better compatibility
#   - Added /RU parameter with domain\username for reliability across different PCs
# - Version 2.100: REBOOT-RESUME FIXES
#   - Simplified reboot logic: reboot only happens at Part 8
#   - Removed exit code checking (assume reboot might happen regardless)
#   - Improved scheduled task creation with verification
#   - Added independent scheduled task test script (TEST_ScheduledTask.ps1)
# - Version 2.000: REBOOT-RESUME EDITION
#   - Added state management (JSON file) to preserve progress across reboots
#   - Added scheduled task for automatic resume after reboot
#   - Implements resume logic with -Resume parameter
#   - Saves service/process states before reboot
#   - Cleans up state file and scheduled task on completion
# ==================================================================================================

# Script parameters
param(
    [switch]$Resume
)

# Initialize script start time
$startTime = Get-Date

# ==================================================================================================
# CONFIGURATION
# ==================================================================================================

$Global:Config = @{
    ScriptVersion  = "2.140"
    WorkingDir     = "C:\winsm"
    LogDir         = "C:\winsm\SmartOffice_Installer\soua_logs"
    StateFile      = "C:\winsm\soua_state.json"
    ResumeTaskName = "SOUpgradeAssistant_Resume"
    Services       = @{
        LiveSales = "srvSOLiveSales"
        Firebird  = "FirebirdServerDefaultInstance"
    }
    Processes      = @{
        SMUpdates   = "SMUpdates"
        PDTWiFi     = "PDTWiFi"
        PDTWiFi64   = "PDTWiFi64"
        SmartOffice = @("Sm32Main", "Sm32")
        Firebird    = "firebird"
    }
    Paths          = @{
        StationMaster = "C:\Program Files (x86)\StationMaster"
        Firebird      = "C:\Program Files (x86)\Firebird"
        SetupDir      = "C:\winsm\SmartOffice_Installer"
    }
    URLs           = @{
        ModuleSOGets   = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_soget.ps1"
        ModuleFirebird = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_firebird.ps1"
    }
    Timeouts       = @{
        ProcessCheckInterval = 3
        MonitorInterval      = 2
        ServiceRetryInterval = 5
        ServiceMaxRetries    = 3
    }
}

# VirtualKeyCode constants for keyboard input
$VK_ENTER = 13
$VK_9 = 57

# Session ID for this run (used for logging and state tracking)
$Global:SessionId = Get-Date -Format 'yyyy-MM-dd_HHmm'

# ==================================================================================================
# FUNCTIONS
# ==================================================================================================

function Write-Log {
    <#
    .SYNOPSIS
        Writes messages to both console and log file with color support.
    #>
    param(
        [Parameter(Mandatory = $false)][string]$Message = "",
        [string]$ForegroundColor = "White",
        [switch]$NoNewline
    )

    # Ensure Log Directory Exists
    if (-not (Test-Path $Global:Config.LogDir)) {
        New-Item -Path $Global:Config.LogDir -ItemType Directory -Force | Out-Null
    }

    # Use session-specific log file if resuming, otherwise create new one
    if ($Global:ResumeLogFile) {
        $logFile = $Global:ResumeLogFile
    }
    else {
        $logFile = Join-Path $Global:Config.LogDir "soua_log_$($Global:SessionId).log"
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"

    # Write to Console
    if ($NoNewline) {
        Write-Host $Message -ForegroundColor $ForegroundColor -NoNewline
    }
    else {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }

    # Write to File (Strip colors/formatting for file)
    Add-Content -Path $logFile -Value $logEntry
}

function Show-Intro {
    <#
    .SYNOPSIS
        Displays the script's introduction banner.
    #>
    Write-Log "SO Upgrade Assistant - Version $($Global:Config.ScriptVersion) [REBOOT-RESUME EDITION]" -ForegroundColor Green
    Write-Log "--------------------------------------------------------------------------------"
    Write-Log ""
}

function Test-Admin {
    <#
    .SYNOPSIS
        Tests if the script is running with administrator privileges.
    #>
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Save-ScriptState {
    <#
    .SYNOPSIS
        Saves the current script state to a JSON file.
    #>
    param(
        [int]$CurrentPart,
        [bool]$RebootRequired = $false,
        [hashtable]$AdditionalData = @{}
    )

    try {
        $state = @{
            ScriptVersion  = $Global:Config.ScriptVersion
            SessionId      = $Global:SessionId
            StartTime      = $startTime.ToString("o")
            CurrentPart    = $CurrentPart
            RebootRequired = $RebootRequired
            LogFile        = (Join-Path $Global:Config.LogDir "soua_log_$($Global:SessionId).log")
            SavedAt        = (Get-Date).ToString("o")
        }

        # Merge additional data
        foreach ($key in $AdditionalData.Keys) {
            $state[$key] = $AdditionalData[$key]
        }

        # Convert to JSON and save
        $state | ConvertTo-Json -Depth 10 | Set-Content -Path $Global:Config.StateFile -Force
        Write-Log "State saved to $($Global:Config.StateFile)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Log "Error saving state: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Get-ScriptState {
    <#
    .SYNOPSIS
        Loads the script state from JSON file.
    #>
    try {
        if (Test-Path $Global:Config.StateFile) {
            $stateJson = Get-Content -Path $Global:Config.StateFile -Raw
            $state = $stateJson | ConvertFrom-Json
            Write-Log "State loaded from $($Global:Config.StateFile)" -ForegroundColor Green
            return $state
        }
        else {
            Write-Log "No state file found at $($Global:Config.StateFile)" -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        Write-Log "Error loading state: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "State file may be corrupted. Starting fresh..." -ForegroundColor Yellow
        return $null
    }
}

function Clear-ScriptState {
    <#
    .SYNOPSIS
        Clears the state file and removes the resume scheduled task.
    #>
    try {
        # Remove state file
        if (Test-Path $Global:Config.StateFile) {
            Remove-Item -Path $Global:Config.StateFile -Force
            Write-Log "State file removed." -ForegroundColor Green
        }

        # Remove scheduled task
        Remove-ResumeTask

        return $true
    }
    catch {
        Write-Log "Error clearing state: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Register-ResumeTask {
    <#
    .SYNOPSIS
        Creates a scheduled task to resume the script after reboot.
    #>
    try {
        $scriptUrl = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/main/souaTEST_REBOOT.ps1"
        
        Write-Log "Creating scheduled task..." -ForegroundColor Yellow
        Write-Log "  Script URL: $scriptUrl" -ForegroundColor Gray
        Write-Log "  Task Name: $($Global:Config.ResumeTaskName)" -ForegroundColor Gray
        
        # Create a command that downloads and runs the script with -Resume parameter
        # We use a wrapper command because the script only exists online
        $downloadAndRun = "Invoke-Expression `"& ([ScriptBlock]::Create((Invoke-RestMethod -Uri '$scriptUrl'))) -Resume`""
        $taskCmd = "PowerShell.exe -ExecutionPolicy Bypass -WindowStyle Normal -Command `"$downloadAndRun`""
        $currentUser = "$env:USERDOMAIN\$env:USERNAME"
        Write-Log "  Task will download and run script from GitHub with -Resume parameter" -ForegroundColor Gray
        Write-Log "  Trigger: On user logon ($currentUser)" -ForegroundColor Gray
        
        $result = & schtasks /Create /TN $Global:Config.ResumeTaskName /TR $taskCmd /SC ONLOGON /RU $currentUser /RL HIGHEST /F 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  Running as: $currentUser" -ForegroundColor Gray
            
            # Verify task was created
            $verifyTask = Get-ScheduledTask -TaskName $Global:Config.ResumeTaskName -ErrorAction SilentlyContinue
            if ($verifyTask) {
                Write-Log "Resume task '$($Global:Config.ResumeTaskName)' registered and verified successfully." -ForegroundColor Green
                return $true
            }
            else {
                Write-Log "Task registration completed but verification failed." -ForegroundColor Red
                return $false
            }
        }
        else {
            Write-Log "schtasks returned exit code $LASTEXITCODE : $result" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Log "Error registering resume task: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Remove-ResumeTask {
    <#
    .SYNOPSIS
        Removes the resume scheduled task.
    #>
    try {
        $task = Get-ScheduledTask -TaskName $Global:Config.ResumeTaskName -ErrorAction SilentlyContinue
        if ($task) {
            Unregister-ScheduledTask -TaskName $Global:Config.ResumeTaskName -Confirm:$false
            Write-Log "Resume task '$($Global:Config.ResumeTaskName)' removed." -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Log "Error removing resume task: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Wait-SmartOfficeClosed {
    <#
    .SYNOPSIS
        Waits for Smart Office processes to be closed.
    #>
    param(
        [string]$Message = "Smart Office is open. Please close it to continue."
    )
    
    foreach ($process in $Global:Config.Processes.SmartOffice) {
        $processRunning = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($processRunning) {
            Write-Log $Message -ForegroundColor Red
            while (Get-Process -Name $process -ErrorAction SilentlyContinue) {
                Start-Sleep -Seconds $Global:Config.Timeouts.ProcessCheckInterval
            }
            Write-Log "Smart Office is now closed." -ForegroundColor Green
        }
        else {
            Write-Log "Smart Office ($process) is not running." -ForegroundColor Yellow
        }
    }
}

function Wait-SingleFirebirdInstance {
    <#
    .SYNOPSIS
        Waits until only one instance of Firebird is running.
    #>
    $firebirdProcesses = Get-Process -Name $Global:Config.Processes.Firebird -ErrorAction SilentlyContinue
    while ($firebirdProcesses.Count -gt 1) {
        Write-Log "`rWarning= Multiple instances of 'firebird.exe' are running. Currently: $($firebirdProcesses.Count) " -ForegroundColor Yellow -NoNewline
        Start-Sleep -Seconds $Global:Config.Timeouts.ProcessCheckInterval
        $firebirdProcesses = Get-Process -Name $Global:Config.Processes.Firebird -ErrorAction SilentlyContinue
    }
    Write-Log "Only one instance of 'firebird.exe' is running." -ForegroundColor Green
}

function Get-StatusColor {
    <#
    .SYNOPSIS
        Returns appropriate color for service/process status.
    #>
    param([string]$Status)
    
    if ($Status -eq "Running") {
        return "Green"
    }
    elseif ($Status -eq "Not Installed") {
        return "Gray"
    }
    else {
        return "Yellow"
    }
}

# ==================================================================================================
# INITIALIZATION
# ==================================================================================================

# Display initial version
Write-Host "SOUpgradeAssistant_REBOOT.ps1 - Version $($Global:Config.ScriptVersion)"

# Create working directory if it doesn't exist
if (-not (Test-Path $Global:Config.WorkingDir -PathType Container)) {
    try {
        New-Item -Path $Global:Config.WorkingDir -ItemType Directory -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Log "Error= Unable to create directory $($Global:Config.WorkingDir)" -ForegroundColor Red
        exit
    }
}

Set-Location -Path $Global:Config.WorkingDir

# ==================================================================================================
# RESUME LOGIC
# ==================================================================================================

if ($Resume) {
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "RESUMING AFTER REBOOT" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""

    $state = Get-ScriptState

    if ($state) {
        # Set session ID to match the original session
        $Global:SessionId = $state.SessionId
        $Global:ResumeLogFile = $state.LogFile
        
        Write-Log "=================================================================================" -ForegroundColor Cyan
        Write-Log "RESUMING FROM PART $($state.CurrentPart) AFTER REBOOT" -ForegroundColor Cyan
        Write-Log "Original session started: $($state.StartTime)" -ForegroundColor Cyan
        Write-Log "=================================================================================" -ForegroundColor Cyan
        Write-Log ""

        # Restore variables from state
        $wasRunning = $state.ServicesState.srvSOLiveSales.WasRunning
        $PDTWiFiStates = @{}
        $PDTWiFiStates[$Global:Config.Processes.PDTWiFi] = $state.ProcessesState.PDTWiFi
        $PDTWiFiStates[$Global:Config.Processes.PDTWiFi64] = $state.ProcessesState.PDTWiFi64
        
        if ($state.SetupExecutable) {
            $selectedExe = Get-Item $state.SetupExecutable -ErrorAction SilentlyContinue
        }

        # Jump to the appropriate part based on saved state
        if ($state.CurrentPart -eq 8) {
            Write-Log "Setup has completed. Continuing to post-installation tasks..." -ForegroundColor Green
            # Skip directly to Part 9 (Post Upgrade)
            goto Part9
        }
        else {
            Write-Log "Unexpected resume point: Part $($state.CurrentPart)" -ForegroundColor Yellow
            Write-Log "Continuing with normal execution..." -ForegroundColor Yellow
        }
    }
    else {
        Write-Log "Resume requested but no valid state file found. Starting fresh..." -ForegroundColor Yellow
        # Continue with normal execution
    }
}

# ==================================================================================================
# PART 1 - Check for Admin Rights
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 1/14] System Pre-Checks" -ForegroundColor Cyan
Write-Log "[______________________________]" -ForegroundColor Cyan
Write-Log ""

# Check if the script is running with administrator rights
if (-not (Test-Admin)) {
    Write-Log "Error= Administrator rights required to run this script. Exiting." -ForegroundColor Red
    pause
    exit
}

# ==================================================================================================
# PART 2 - Download Setup Files (module_soget)
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 2/14] Checking for Setup Files. Please Wait." -ForegroundColor Cyan
Write-Log "[██________________________]" -ForegroundColor Cyan
Write-Log ""
Start-Sleep -Seconds 1

# Define the URL for the SO Get module
$sogetScriptURL = $Global:Config.URLs.ModuleSOGets

# Run module_soget directly in the current shell
try {
    Write-Log "Invoking module_soget.ps1..." -ForegroundColor Yellow
    Invoke-Expression (Invoke-RestMethod -Uri $sogetScriptURL -ErrorAction Stop)
    Write-Log "module_soget.ps1 executed successfully." -ForegroundColor Green
}
catch {
    Write-Log "Error executing module_soget.ps1 $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# ==================================================================================================
# PART 3 - Firebird Installation (module_firebird)
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 3/14] Checking for Firebird installation" -ForegroundColor Cyan
Write-Log "[████______________________]" -ForegroundColor Cyan
Write-Log ""
Start-Sleep -Seconds 1

# Define the URL for the Firebird installation script
$firebirdInstallerURL = $Global:Config.URLs.ModuleFirebird

# Check if Firebird is installed
$firebirdDir = $Global:Config.Paths.Firebird
if (-not (Test-Path $firebirdDir)) {
    Write-Log "Firebird not found. Installing Firebird..." -ForegroundColor Yellow
    try {
        Invoke-Expression (Invoke-RestMethod -Uri $firebirdInstallerURL -ErrorAction Stop)
        Write-Log "Firebird installation script executed successfully." -ForegroundColor Green
    }
    catch {
        Write-Log "Error executing Firebird installation script: $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
}
else {
    Write-Log "Firebird is already installed." -ForegroundColor Green
}

# ==================================================================================================
# PART 4 - Stop SMUpdates if Running
# PartVersion-1.01
# - Fixed background job scope issue (pass process name as parameter)
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 4/14] Monitoring and Stopping SMUpdates" -ForegroundColor Cyan
Write-Log "[██████____________________]" -ForegroundColor Cyan
Write-Log ""

# Start background job to monitor SMUpdates (pass process name as parameter)
$monitorJob = Start-Job -ArgumentList $Global:Config.Processes.SMUpdates -ScriptBlock {
    param($ProcessName)
    
    while ($true) {
        $smUpdatesProcess = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($smUpdatesProcess) {
            Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 2
    }
}

Write-Log "Monitoring for SMUpdates process in the background..." -ForegroundColor Yellow

# ==================================================================================================
# PART 5 - Manage SO Live Sales Service
# PartVersion-1.01
# - Initialized $wasRunning variable.
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 5/14] Managing SO Live Sales Service" -ForegroundColor Cyan
Write-Log "[████████__________________]" -ForegroundColor Cyan
Write-Log ""

$ServiceName = $Global:Config.Services.LiveSales
$wasRunning = $false # Initialize the variable

try {
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        if ($service.Status -eq 'Running') {
            $wasRunning = $true
            Write-Log "Service '$ServiceName' is running. Stopping..." -ForegroundColor Yellow
            Stop-Service -Name $ServiceName -Force -ErrorAction Stop
            Write-Log "Service '$ServiceName' stopped successfully." -ForegroundColor Green
        }
        else {
            Write-Log "Service '$ServiceName' is not running." -ForegroundColor Yellow
        }
    }
    else {
        Write-Log "Service '$ServiceName' not found." -ForegroundColor Yellow
    }
}
catch {
    Write-Log "Error managing service '$ServiceName': $($_.Exception.Message)" -ForegroundColor Red
}

# ==================================================================================================
# PART 6 - Manage PDTWiFi Processes
# PartVersion-1.02
# - Total redo
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 6/14] Managing PDTWiFi processes" -ForegroundColor Cyan
Write-Log "[██████████________________]" -ForegroundColor Cyan
Write-Log ""

# Initialize the process states
$PDTWiFiStates = @{}

# Section A - PDTWiFi
$PDTWiFi = $Global:Config.Processes.PDTWiFi
$pdtWiFiProcess = Get-Process -Name $PDTWiFi -ErrorAction SilentlyContinue
if ($pdtWiFiProcess) {
    $PDTWiFiStates[$PDTWiFi] = "Running"
    Stop-Process -Name $PDTWiFi -Force -ErrorAction SilentlyContinue
    Write-Log "$PDTWiFi stopped." -ForegroundColor Green
}
else {
    $PDTWiFiStates[$PDTWiFi] = "Not running"
    Write-Log "$PDTWiFi is not running." -ForegroundColor Yellow
}

# Section B - PDTWiFi64
$PDTWiFi64 = $Global:Config.Processes.PDTWiFi64
$pdtWiFi64Process = Get-Process -Name $PDTWiFi64 -ErrorAction SilentlyContinue
if ($pdtWiFi64Process) {
    $PDTWiFiStates[$PDTWiFi64] = "Running"
    Stop-Process -Name $PDTWiFi64 -Force -ErrorAction SilentlyContinue
    Write-Log "$PDTWiFi64 stopped." -ForegroundColor Green
}
else {
    $PDTWiFiStates[$PDTWiFi64] = "Not running"
    Write-Log "$PDTWiFi64 is not running." -ForegroundColor Yellow
}

# ==================================================================================================
# PART 7 - Make Sure SO is Closed & Wait for Single Instance of Firebird.exe
# PartVersion-1.01
# - Using consolidated Wait-SmartOfficeClosed function
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 7/14] Wait SO Closed & Single Firebird process..." -ForegroundColor Cyan
Write-Log "[████████████______________]" -ForegroundColor Cyan
Write-Log ""

# Make sure SO is closed (using consolidated function)
Wait-SmartOfficeClosed

# Wait for single firebird instance
if (-not (Test-Path $Global:Config.Paths.SetupDir -PathType Container)) {
    Write-Log "Error Setup directory '$($Global:Config.Paths.SetupDir)' does not exist." -ForegroundColor Red
    exit
}

Wait-SingleFirebirdInstance

# ==================================================================================================
# PART 8 - Launch Setup
# PartVersion 2.10
# - Simplified: Assume reboot might happen (don't check exit code)
# - Reboot ONLY happens at Part 8 (nowhere else)
# - Improved scheduled task creation with verification
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 8/14] Launching SO setup..." -ForegroundColor Cyan
Write-Log "[██████████████____________]" -ForegroundColor Cyan
Write-Log ""

# Get all setup executables in the SmartOffice_Installer directory
$setupExes = Get-ChildItem -Path $Global:Config.Paths.SetupDir -Filter "*.exe"

if ($setupExes.Count -eq 0) {
    Write-Log "Error No executable (.exe) found in '$($Global:Config.Paths.SetupDir)'." -ForegroundColor Red
    exit
}
elseif ($setupExes.Count -eq 1) {
    # Only one file found, proceed without asking the user
    $selectedExe = $setupExes[0]
    Write-Log "Found setup $($selectedExe.Name)" -ForegroundColor Green
}
else {
    # Sort executables by version (numeric part in the name) ascending
    $setupExes = $setupExes | Sort-Object {
        [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int]
    }
    
    # Multiple setup files found, present a terminal selection menu
    Write-Log "`nPlease select the setup to run`n" -ForegroundColor Yellow
    Write-Log ("{0,-5} {1,-30} {2,-20} {3,-10}" -f "No.", "Executable Name", "Date Modified", "Version") -ForegroundColor White
    Write-Log ("{0,-5} {1,-30} {2,-20} {3,-10}" -f "---", "------------------------------", "--------------------", "-----------") -ForegroundColor Gray
    
    for ($i = 0; $i -lt $setupExes.Count; $i++) {
        $exe = $setupExes[$i]
        $dateModified = $exe.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        $versionType = if ($i -eq 0) { "Current" } else { "Next" }
        $color = if ($i -eq 0) { "Green" } else { "Yellow" }
        Write-Log ("{0,-5} {1,-30} {2,-20} {3,-10}" -f ($i + 1), $exe.Name, $dateModified, $versionType) -ForegroundColor $color
    }
    
    Write-Log "`nEnter the number of your selection (or press Enter to cancel):" -ForegroundColor Cyan
    $selection = Read-Host "Selection"
    
    # Check if the user wants to cancel
    if ([string]::IsNullOrWhiteSpace($selection)) {
        Write-Log "Operation cancelled. Setup is required to continue. Exiting." -ForegroundColor Red
        exit
    }
    
    # Validate the selection
    if ($selection -match '^\d+$' -and $selection -ge 1 -and $selection -le $setupExes.Count) {
        $selectedExe = $setupExes[$selection - 1]
        Write-Log "Selected setup executable $($selectedExe.Name)" -ForegroundColor Green
    }
    else {
        Write-Log "Invalid selection. Exiting." -ForegroundColor Red
        exit
    }
}

# Save state BEFORE launching setup (in case reboot is needed)
Write-Log "Saving state before launching setup..." -ForegroundColor Yellow
Save-ScriptState -CurrentPart 8 -AdditionalData @{
    ServicesState   = @{
        srvSOLiveSales = @{
            WasRunning = $wasRunning
            Status     = "Stopped"
        }
    }
    ProcessesState  = @{
        PDTWiFi   = $PDTWiFiStates[$PDTWiFi]
        PDTWiFi64 = $PDTWiFiStates[$PDTWiFi64]
    }
    SetupExecutable = $selectedExe.FullName
}

# Register scheduled task BEFORE launching setup (in case reboot happens)
Write-Log "Registering resume task (in case reboot is required)..." -ForegroundColor Yellow
if (-not (Register-ResumeTask)) {
    Write-Log "Warning: Failed to register resume task." -ForegroundColor Red
    Write-Log "If setup requires a reboot, you will need to manually re-run this script with -Resume parameter." -ForegroundColor Yellow
    Write-Log ""
    Write-Log "Do you want to continue anyway? (Y/N)" -ForegroundColor Cyan
    $response = Read-Host
    if ($response -ne "Y" -and $response -ne "y") {
        Write-Log "Operation cancelled." -ForegroundColor Red
        Clear-ScriptState
        exit
    }
}

# Launch the selected setup executable
try {
    Write-Log "Starting executable: $($selectedExe.Name) ..." -ForegroundColor Cyan
    Write-Log ""
    Write-Log "NOTE: If setup requires a reboot, the system will restart and this script will automatically resume." -ForegroundColor Yellow
    Write-Log ""
    
    $setupProcess = Start-Process -FilePath $selectedExe.FullName -Wait -PassThru
    
    Write-Log ""
    Write-Log "Setup executable finished with exit code: $($setupProcess.ExitCode)" -ForegroundColor Cyan
    
    # Ask user if reboot is needed (we don't rely on exit codes)
    Write-Log ""
    Write-Log "Did the setup indicate that a REBOOT is required? (Y/N)" -ForegroundColor Yellow
    $rebootNeeded = Read-Host
    
    if ($rebootNeeded -eq "Y" -or $rebootNeeded -eq "y") {
        Write-Log "" -ForegroundColor Yellow
        Write-Log "================================================================================" -ForegroundColor Yellow
        Write-Log "REBOOT REQUIRED" -ForegroundColor Yellow
        Write-Log "================================================================================" -ForegroundColor Yellow
        Write-Log "Setup requires a reboot to complete installation." -ForegroundColor Yellow
        Write-Log "The script will automatically resume after reboot." -ForegroundColor Cyan
        Write-Log "All progress has been saved." -ForegroundColor Green
        Write-Log ""
        Write-Log "================================================================================" -ForegroundColor Cyan
        Write-Log "Press any key to reboot now, or close this window to reboot manually later..." -ForegroundColor Cyan
        Write-Log "================================================================================" -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        # Reboot
        Write-Log "Rebooting system..." -ForegroundColor Yellow
        Restart-Computer -Force
        exit
    }
    else {
        Write-Log "No reboot required. Continuing with post-installation tasks..." -ForegroundColor Green
        # Remove the scheduled task since we don't need it
        Remove-ResumeTask
    }
}
catch {
    Write-Log "Error starting setup executable $_" -ForegroundColor Red
    Clear-ScriptState  # Clean up on error
    exit
}

# ==================================================================================================
# PART 9 - Wait for User Confirmation
# PartVersion-2.00
# - Added label for resume jump point
#LOCK=ON
# ==================================================================================================
:Part9
Clear-Host
Show-Intro
Write-Log "[Part 9/14] Post Upgrade" -ForegroundColor Cyan
Write-Log "[████████████████__________]" -ForegroundColor Cyan
Write-Log ""

# Stop monitoring SMUpdates process (if job exists)
if ($monitorJob) {
    Stop-Job -Job $monitorJob -ErrorAction SilentlyContinue
    Remove-Job -Job $monitorJob -ErrorAction SilentlyContinue
    Write-Log "SMUpdates monitoring stopped." -ForegroundColor Green
}

Write-Log "Waiting for confirmation Upgrade is Complete..." -ForegroundColor Yellow
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("Please ensure the upgrade is complete and Smart Office is closed before clicking OK.", "SO Post Upgrade Confirmation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Check for Running SO Processes Again (using consolidated function)
Wait-SmartOfficeClosed -Message "Smart Office is still running. Please close it to continue..."
Write-Log "Smart Office processes confirmed closed." -ForegroundColor Green

# ==================================================================================================
# PART 10 - Set Permissions for SM Folder
# PartVersion-1.10
# - Reverted to the original icacls command for setting permissions.
# - Added a message to the user about potential delays.
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 10/14] Setting permissions for Stationmaster folder. Please Wait..." -ForegroundColor Cyan
Write-Log "[██████████████████________]" -ForegroundColor Cyan
Write-Log ""

Write-Log "Please wait, this task may take ~1-30+ minutes to complete depending on PC speed. Do not interrupt." -ForegroundColor Yellow
try {
    & icacls $Global:Config.Paths.StationMaster /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
    Write-Log "Permissions for StationMaster folder set successfully." -ForegroundColor Green
}
catch {
    Write-Log "Error setting permissions for SM folder: $_" -ForegroundColor Red
}

# ==================================================================================================
# PART 11 - Set Permissions for Firebird Folder
# PartVersion-1.00
# - Initial version.
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 11/14] Setting permissions for Firebird folder. Please Wait..." -ForegroundColor Cyan
Write-Log "[████████████████████______]" -ForegroundColor Cyan
Write-Log ""

try {
    & icacls $Global:Config.Paths.Firebird /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
    Write-Log "Permissions for Firebird folder set successfully." -ForegroundColor Green
}
catch {
    Write-Log "Error setting permissions for Firebird folder: $_" -ForegroundColor Red
}

# ==================================================================================================
# PART 12 - Revert SO Live Sales Service
# PartVersion-1.07
# - Modified to only start the service if it was previously running.
# - Using timeout constants from config
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 12/14] SO Live Sales" -ForegroundColor Cyan
Write-Log "[██████████████████████____]" -ForegroundColor Cyan
Write-Log ""

if ($wasRunning) {
    Write-Log "Service '$ServiceName' was running before. Ensuring it restarts..." -ForegroundColor Yellow
    try {
        $retryCount = 0
        $maxRetries = $Global:Config.Timeouts.ServiceMaxRetries
        $retryIntervalSeconds = $Global:Config.Timeouts.ServiceRetryInterval

        while ($retryCount -lt $maxRetries) {
            Write-Log "Attempting to start service '$ServiceName' (Attempt $($retryCount + 1) of $maxRetries)..." -ForegroundColor Yellow
            Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
            
            if ((Get-Service -Name $ServiceName).Status -eq "Running") {
                Write-Log "Service '$ServiceName' is now running." -ForegroundColor Green
                break
            }
            else {
                Write-Log "Service '$ServiceName' is not running. Waiting $retryIntervalSeconds seconds before retrying..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryIntervalSeconds
                $retryCount++
            }
        }

        if ((Get-Service -Name $ServiceName).Status -ne "Running") {
            Write-Warning "Failed to automatically start service '$ServiceName' after $maxRetries attempts."
            Write-Log ""
            Write-Log "Please manually start the '$ServiceName' service now." -ForegroundColor Red
            Write-Log "The script will wait until the service is running..." -ForegroundColor Yellow

            while ((Get-Service -Name $ServiceName).Status -ne "Running") {
                Write-Log "Waiting for '$ServiceName' service to be running. Please manually start the service. Checking again in $($Global:Config.Timeouts.ProcessCheckInterval) seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $Global:Config.Timeouts.ProcessCheckInterval
            }
            Write-Log "Service '$ServiceName' is now running. Continuing..." -ForegroundColor Green
        }
        else {
            Write-Log "'$ServiceName' service confirmed to be running." -ForegroundColor Green
        }
    }
    catch {
        Write-Log "Error encountered while starting service '$ServiceName' $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Log "Service '$ServiceName' was not running before, so no action needed." -ForegroundColor Yellow
}

# ==================================================================================================
# PART 13 - Revert PDTWiFi Processes
# PartVersion 1.02
# - Total re-do
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 13/14] Reverting PDTWiFi processes" -ForegroundColor Cyan
Write-Log "[████████████████████████__]" -ForegroundColor Cyan
Write-Log ""

# Section A - Recall and Revert PDTWiFi
if ($PDTWiFiStates[$PDTWiFi] -eq "Running") {
    try {
        Start-Process (Join-Path $Global:Config.Paths.StationMaster "PDTWiFi.exe") -ErrorAction Stop
        Write-Log "$PDTWiFi started." -ForegroundColor Green
    }
    catch {
        Write-Log "Error starting $PDTWiFi $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Log "$PDTWiFi was not running, no action taken." -ForegroundColor Yellow
}

# Section B - Recall and Revert PDTWiFi64
if ($PDTWiFiStates[$PDTWiFi64] -eq "Running") {
    try {
        Start-Process (Join-Path $Global:Config.Paths.StationMaster "PDTWiFi64.exe") -ErrorAction Stop
        Write-Log "$PDTWiFi64 started." -ForegroundColor Green
    }
    catch {
        Write-Log "Error starting $PDTWiFi64 $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Log "$PDTWiFi64 was not running, no action taken." -ForegroundColor Yellow
}

# ==================================================================================================
# PART 14 - Clean up and Finish Script
# PartVersion-2.00
# - Added state file and scheduled task cleanup
#LOCK=ON
# ==================================================================================================
Clear-Host
Show-Intro
Write-Log "[Part 14/14] Clean up and finish" -ForegroundColor Cyan
Write-Log "[██████████████████████████]" -ForegroundColor Cyan

# Get status of services and processes
$liveSalesService = Get-Service -Name $Global:Config.Services.LiveSales -ErrorAction SilentlyContinue
if ($liveSalesService) {
    $liveSalesServiceStatus = $liveSalesService.Status
}
else {
    $liveSalesServiceStatus = "Not Installed"
}

$pdtWifiStatus = if (Get-Process -Name $Global:Config.Processes.PDTWiFi -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }
$pdtWifi64Status = if (Get-Process -Name $Global:Config.Processes.PDTWiFi64 -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }

# Output the status table
Write-Log " "
Write-Log "Process Status:" -ForegroundColor Yellow
Write-Log "------------------------------------------------" -ForegroundColor Yellow
Write-Log ("{0,-25} {1,-15}" -f "Item", "Status") -ForegroundColor Yellow
Write-Log ("{0,-25} {1,-15}" -f "-------------------------", "---------------") -ForegroundColor Yellow

Write-Log ("{0,-25} {1,-15}" -f "SO Live Sales Service", $liveSalesServiceStatus) -ForegroundColor (Get-StatusColor $liveSalesServiceStatus)
Write-Log ("{0,-25} {1,-15}" -f "PDTWiFi.exe", $pdtWifiStatus) -ForegroundColor (Get-StatusColor $pdtWifiStatus)
Write-Log ("{0,-25} {1,-15}" -f "PDTWiFi64.exe", $pdtWifi64Status) -ForegroundColor (Get-StatusColor $pdtWifi64Status)

Write-Log "------------------------------------------------" -ForegroundColor Yellow

# Clean up state file and scheduled task
Write-Log ""
Write-Log "Cleaning up state file and resume task..." -ForegroundColor Yellow
Clear-ScriptState
Write-Log "Cleanup completed successfully." -ForegroundColor Green

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds

Write-Log " "
Write-Log "Completed in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
Write-Log " "
Write-Log "Consider if you need to Reboot at this stage." -ForegroundColor Yellow
Write-Log " "
Write-Log "Press Enter to start Smart Office, '9' to reboot now, or any other key to exit."
$key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

if ($key.VirtualKeyCode -eq $VK_ENTER) {
    # Enter key
    Write-Log "Starting Smart Office..." -ForegroundColor Cyan
    try {
        Start-Process (Join-Path $Global:Config.Paths.StationMaster "Sm32.exe") -ErrorAction Stop
        Write-Log "Smart Office started successfully." -ForegroundColor Green
    }
    catch {
        Write-Log "Error starting Smart Office: $($_.Exception.Message)" -ForegroundColor Red
    }
}
elseif ($key.VirtualKeyCode -eq $VK_9) {
    # '9' key
    Write-Log "Rebooting..." -ForegroundColor Cyan
    Restart-Computer -Force
}
else {
    Write-Log "Exiting..." -ForegroundColor Cyan
}
