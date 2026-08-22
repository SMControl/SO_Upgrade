# soua.ps1 - Script v1.04 - PartVersion-1.04
Write-Host "SOUpgradeAssistant.ps1 - Version 1.218" # previous working to this was 1.217
# This script automates the upgrade process for Smart Office (SO) software.

# Global Script Configuration
$TOTAL_PARTS = 14
$startTime = Get-Date
$workingDir = "C:\winsm"
# Stores initial running status of services/processes for later reversion (Parts 5, 6, 12, 13)
$processStates = @{}
$SO_SERVICE = "srvSOLiveSales"
$SO_PROCESSES = @("Sm32Main", "Sm32")

# Configure SSL/TLS Security Protocols and Certificate Validation Bypass
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

# ==================================
# Utility Functions
# ==================================

# Function to display part intro and progress
function Start-Part {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][int]$Current,
        [Parameter(Mandatory = $true)][string]$PartVersion
    )
    Clear-Host
    Write-Host "SO Upgrade Assistant - Version 1.218" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------------------------"
    
    $progressBarWidth = 30
    $filled = [int]($progressBarWidth * $Current / $TOTAL_PARTS)
    $empty = $progressBarWidth - $filled
    $progress = "[" + ("█" * $filled) + ("_" * $empty) + "]"
    
    Write-Host "[Part $Current/$TOTAL_PARTS] $Title" -ForegroundColor Cyan
    Write-Host "$progress" -ForegroundColor Cyan
    Write-Host ""
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
        [Parameter(Mandatory = $true)][string]$ProcessName
    )
    $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    $wasRunning = $false
    
    if ($process) {
        $wasRunning = $true
        Write-Host "Process '$ProcessName' detected. Stopping..." -ForegroundColor Yellow
        Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
    } 
    
    $processStates[$ProcessName] = $wasRunning
    return $wasRunning
}

# Function to get latest Smart Office Setup files from website (formerly module_soget)
function Get-SOSetup {
    Write-Host "Checking Website for latest..." -ForegroundColor Yellow
    try {
        $exeLinks = (Invoke-WebRequest -Uri "https://www.stationmaster.com/downloads/" -ErrorAction Stop).Links | Where-Object { $_.href -match "\.exe$" } | ForEach-Object { $_.href }
        $setupLinks = $exeLinks | Where-Object { $_ -match "^https://www\.stationmaster\.com/Download/Setup\d+\.exe$" }
        $sortedLinks = $setupLinks | Sort-Object { [regex]::Match($_, "Setup(\d+)\.exe").Groups[1].Value -as [int] } -Descending
        $highestTwoLinks = $sortedLinks | Select-Object -First 2
        
        $downloadDirectory = "C:\winsm\SmartOffice_Installer"
        if (-not (Test-Path $downloadDirectory)) {
            Write-Host "Creating Folder C:\winsm\SmartOffice_Installer" -ForegroundColor Yellow
            New-Item -ItemType Directory -Path $downloadDirectory -ErrorAction Stop | Out-Null
        }
        foreach ($downloadLink in $highestTwoLinks) {
            $originalFilename = $downloadLink.Split('/')[-1]
            $destinationPath = Join-Path -Path $downloadDirectory -ChildPath $originalFilename
            
            $request = [System.Net.HttpWebRequest]::Create($downloadLink)
            $request.Method = "HEAD"
            $request.UserAgent = "Mozilla/5.0"
            try {
                $response = $request.GetResponse()
                $contentLength = $response.ContentLength
                $response.Close()
            }
            catch {
                continue
            }
            
            # Compare size with existing file
            $matchingFile = Get-ChildItem -Path $downloadDirectory -Filter $originalFilename -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -eq $contentLength }

            if (-not $matchingFile) {
                Write-Host "Downloading new version: $originalFilename" -ForegroundColor Green
                Invoke-WebRequest -Uri $downloadLink -OutFile $destinationPath -ErrorAction Stop
            }
        }
        
        # Delete older downloads, keeping the latest two
        $downloadedFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe" | Sort-Object LastWriteTime -Descending
        if ($downloadedFiles.Count -gt 2) {
            $filesToDelete = $downloadedFiles | Select-Object -Skip 2
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        throw "Failed downloading setup files: $($_.Exception.Message)"
    }
}

# Function to register scheduled task for daily setup check (formerly task_SO Setup Get.ps1)
function Register-SOSetupTask {
    Write-Host "Checking for scheduled task 'SO Setup Get'..." -ForegroundColor Yellow
    $taskExists = Get-ScheduledTask -TaskName "SO Setup Get" -ErrorAction SilentlyContinue
    if (-not $taskExists) {
        Write-Host "Task not found. Creating a new scheduled task..." -ForegroundColor Green
        $TaskName = "SO Setup Get"
        $Description = "Gets the latest SO Installer, if new."
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -NoProfile -NonInteractive -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; iwr -Uri https://raw.githubusercontent.com/SMControl/SM_Tasks/refs/heads/main/bin/SO_Setup_Get.ps1 -UseBasicParsing | Select-Object -ExpandProperty Content | iex`""
        $randomHour = Get-Random -Minimum 1 -Maximum 6
        $randomMinute = Get-Random -Minimum 0 -Maximum 59
        $randomTime = "{0:D2}:{1:D2}" -f $randomHour, $randomMinute
        Write-Host "Scheduled time set to a random time between 01:00 and 06:00 daily: $randomTime"
        $trigger = New-ScheduledTaskTrigger -Daily -At $randomTime
        $settings = New-ScheduledTaskSettingsSet -Hidden:$true -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
        $PrincipalUser = "$env:USERDOMAIN\$env:USERNAME"
        $PrincipalLogonType = "Interactive"
        $PrincipalRunLevel = "Highest"
        $principal = New-ScheduledTaskPrincipal -UserId $PrincipalUser -LogonType $PrincipalLogonType -RunLevel $PrincipalRunLevel
        try {
            Register-ScheduledTask -TaskName $TaskName -Description $Description -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
            Write-Host "Scheduled task 'SO Setup Get' registered successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error registering scheduled task: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Scheduled task 'SO Setup Get' already exists. No action needed." -ForegroundColor Green
    }
}

# ==================================
# Pre-Script Setup
# ==================================

# Setup working directory
if (-not (Test-Path $workingDir -PathType Container)) {
    try {
        New-Item -Path $workingDir -ItemType Directory -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Error= Unable to create directory $workingDir" -ForegroundColor Red
        exit
    }
}
Set-Location -Path $workingDir

# ==================================
# Part 1 - Check for Admin Rights & Register Setup Task
# PartVersion-1.03
# - Use local Register-SOSetupTask instead of remote script fetch.
#LOCK=ON
# ==================================
Start-Part -Title "System Pre-Checks & Setup Tool Installation" -Current 1 -PartVersion "1.03"

if (-not (Test-Admin)) {
    Write-Host "Error= Administrator rights required to run this script. Exiting." -ForegroundColor Red
    pause
    exit
}

Register-SOSetupTask

# ==================================
# Part 2 - SO_UC.exe // calling module_soget
# PartVersion-1.05
# - Re-numbered part to 2, incorporated module_soget.
#LOCK=ON
# ==================================
Start-Part -Title "Checking for Setup Files (module_soget)" -Current 2 -PartVersion "1.05"

try {
    Get-SOSetup
    Write-Host "Smart Office setup files checked/updated successfully." -ForegroundColor Green
}
catch {
    Write-Host "Warning: Could not check for latest setup files online ($($_.Exception.Message)). Using existing local files." -ForegroundColor Yellow
}

# ==================================
# Part 3 - Check for Firebird Folder
# PartVersion-1.06
# - Check for Firebird installation folder; stop script if missing.
#LOCK=ON
# ==================================
Start-Part -Title "Checking for Firebird installation" -Current 3 -PartVersion "1.06"

$firebirdPaths = @(
    "C:\Program Files (x86)\Firebird",
    "C:\Program Files\Firebird"
)

$firebirdFound = $false
foreach ($path in $firebirdPaths) {
    if (Test-Path $path) {
        $firebirdFound = $true
        break
    }
}

if ($firebirdFound) {
    Write-Host "Firebird folder found. Continuing..." -ForegroundColor Green
}
else {
    Write-Host "Firebird is not installed. Please install Firebird before running this script." -ForegroundColor Red
    exit
}

# ==================================
# Part 4 - Stop SMUpdates if Running (Background Monitor)
# PartVersion-1.03
# - Removed redundant starting message.
#LOCK=ON
# ==================================
Start-Part -Title "Monitoring and Stopping SMUpdates (Background)" -Current 4 -PartVersion "1.03"

$monitorJob = Start-Job -ScriptBlock {
    function Monitor-SmUpdates {
        while ($true) {
            $smUpdatesProcess = Get-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
            if ($smUpdatesProcess) {
                Stop-Process -Name "SMUpdates" -Force -ErrorAction SilentlyContinue
            }
            Start-Sleep -Seconds 2
        }
    }
    Monitor-SmUpdates
}

try {
    # ==================================
    # Part 5 - Manage SO Live Sales Service
    # PartVersion-1.03
    # - Removed "stopped successfully," "not running," and "not found" messages.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Stopping SO Live Sales Service" -Current 5 -PartVersion "1.03"

    try {
        $service = Get-Service -Name $SO_SERVICE -ErrorAction SilentlyContinue
        if ($service) {
            $processStates[$SO_SERVICE] = ($service.Status -eq 'Running')
            if ($processStates[$SO_SERVICE]) {
                Write-Host "Service '$SO_SERVICE' is running. Stopping..." -ForegroundColor Yellow
                Stop-Service -Name $SO_SERVICE -Force -ErrorAction Stop
            } 
        } 
    }
    catch {
        Write-Host "Error managing service '$SO_SERVICE': $($_.Exception.Message)" -ForegroundColor Red
    }

    # ==================================
    # Part 6 - Manage PDTWiFi Processes
    # PartVersion-1.04
    # - Messages handled by modified Stop-ProcessIfRunning function (only output on detected/stopping).
    #LOCK=ON
    # ==================================
    Start-Part -Title "Managing PDTWiFi processes" -Current 6 -PartVersion "1.04"

    Stop-ProcessIfRunning -ProcessName "PDTWiFi"
    Stop-ProcessIfRunning -ProcessName "PDTWiFi64"

    # ==================================
    # Part 7 - Make Sure SO is closed & Wait for Single Instance of Firebird.exe
    # PartVersion-1.03
    # - Wait for single instance of firebird and SO closed.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Wait SO Closed & Single Firebird process..." -Current 7 -PartVersion "1.03"

    # Make sure SO is closed
    foreach ($process in $SO_PROCESSES) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            Write-Host "Smart Office is open. Please close it to continue." -ForegroundColor Red
            while (Get-Process -Name $process -ErrorAction SilentlyContinue) {
                Start-Sleep -Seconds 3
            }
            Write-Host "Smart Office is now closed." -ForegroundColor Green
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
    }
    WaitForSingleFirebirdInstance

    # ==================================
    # Part 8 - Launch Setup
    # PartVersion-1.09
    # - Updated column headers in the setup selection table based on user request.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Launching SO setup..." -Current 8 -PartVersion "1.09"

    $setupExes = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "*.exe"
    if ($setupExes.Count -eq 0) {
        Write-Host "Error No executable (.exe) found in 'C:\winsm\SmartOffice_Installer'." -ForegroundColor Red
        exit
    }
    elseif ($setupExes.Count -eq 1) {
        $selectedExe = $setupExes[0]
        Write-Host "Found setup $($selectedExe.Name)" -ForegroundColor Green
    }
    else {
        $setupExes = $setupExes | Sort-Object { [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int] }
        Write-Host "`nPlease select the setup to run`n" -ForegroundColor Yellow
        
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
        }
        else {
            Write-Host "Invalid selection. Exiting." -ForegroundColor Red
            exit
        }
    }

    try {
        Start-Process -FilePath $selectedExe.FullName -Wait
        Write-Host "Setup executable finished." -ForegroundColor Green
    }
    catch {
        Write-Host "Error starting setup executable $_" -ForegroundColor Red
        exit
    }

    # ==================================
    # Part 9 - Wait for User Confirmation
    # PartVersion-1.03
    # - Removed SMUpdates stopped message outside of Part 5.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Post Upgrade Confirmation" -Current 9 -PartVersion "1.03"

    Stop-Job -Job $monitorJob
    Remove-Job -Job $monitorJob

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
    # Part 10 - Set Permissions for SM Folder
    # PartVersion-1.13
    # - No changes to messaging.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Setting permissions for Stationmaster folder" -Current 10 -PartVersion "1.13"

    Write-Host "Executing permissions update (may take time)..." -ForegroundColor Yellow
    try {
        & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C *>$null
        if ($LASTEXITCODE -ne 0) { throw "icacls failed with exit code $LASTEXITCODE." }
        Write-Host "Permissions for StationMaster folder set successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Error setting permissions for SM folder: $($_.Exception.Message)" -ForegroundColor Red
    }

    # ==================================
    # Part 11 - Set Permissions for Firebird Folder
    # PartVersion-1.03
    # - No changes to messaging.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Setting permissions for Firebird folder" -Current 11 -PartVersion "1.03"

    Write-Host "Executing permissions update (may take time)..." -ForegroundColor Yellow
    try {
        & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C *>$null
        if ($LASTEXITCODE -ne 0) { throw "icacls failed with exit code $LASTEXITCODE." }
        Write-Host "Permissions for Firebird folder set successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Error setting permissions for Firebird folder: $($_.Exception.Message)" -ForegroundColor Red
    }

    # ==================================
    # Part 12 - Revert SO Live Sales Service
    # PartVersion-1.08
    # - No changes to messaging (messages here are crucial for retries/manual intervention).
    #LOCK=ON
    # ==================================
    Start-Part -Title "Restarting SO Live Sales Service" -Current 12 -PartVersion "1.08"

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
        
        if ((Get-Service -Name $SO_SERVICE).Status -ne "Running") {
            Write-Host "Failed to automatically start service '$SO_SERVICE'. Please manually start the service now." -ForegroundColor Red
            while ((Get-Service -Name $SO_SERVICE).Status -ne "Running") {
                Write-Host "Waiting for service to run. Checking again in 3 seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
            }
            Write-Host "Service '$SO_SERVICE' is now running. Continuing..." -ForegroundColor Green
        }
        else {
            Write-Host "'$SO_SERVICE' service confirmed to be running." -ForegroundColor Green
        }
    }
    else {
        Write-Host "Service '$SO_SERVICE' was not running before, no action taken." -ForegroundColor Yellow
    }

    # ==================================
    # Part 13 - Revert PDTWiFi Processes
    # PartVersion-1.03
    # - No changes to messaging.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Reverting PDTWiFi processes" -Current 13 -PartVersion "1.03"

    $PDTWiFi_PATH = "C:\Program Files (x86)\StationMaster"

    if ($processStates["PDTWiFi"]) {
        try {
            Start-Process "$PDTWiFi_PATH\PDTWiFi.exe" -ErrorAction Stop
            Write-Host "PDTWiFi started." -ForegroundColor Green
        }
        catch {
            Write-Host "Error starting PDTWiFi $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "PDTWiFi was not running, no action taken." -ForegroundColor Yellow
    }

    if ($processStates["PDTWiFi64"]) {
        try {
            Start-Process "$PDTWiFi_PATH\PDTWiFi64.exe" -ErrorAction Stop
            Write-Host "PDTWiFi64 started." -ForegroundColor Green
        }
        catch {
            Write-Host "Error starting PDTWiFi64 $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "PDTWiFi64 was not running, no action taken." -ForegroundColor Yellow
    }

    # ==================================
    # Part 14 - Clean up and Finish Script
    # PartVersion-1.12
    # - Re-numbered part to 14.
    #LOCK=ON
    # ==================================
    Start-Part -Title "Clean up and finish" -Current 14 -PartVersion "1.12"

    # Get Current Status of services and processes
    $liveSalesService = Get-Service -Name $SO_SERVICE -ErrorAction SilentlyContinue
    $liveSalesServiceStatus = if ($liveSalesService) { $liveSalesService.Status } else { "Not Installed" }
    $pdtWifiStatus = if (Get-Process -Name "PDTWiFi" -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }
    $pdtWifi64Status = if (Get-Process -Name "PDTWiFi64" -ErrorAction SilentlyContinue) { "Running" } else { "Stopped" }

    # Function to determine status color
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

    if ($key.VirtualKeyCode -eq 13) {
        Write-Host "Starting Smart Office..." -ForegroundColor Cyan
        try {
            Start-Process "C:\Program Files (x86)\StationMaster\Sm32.exe" -ErrorAction Stop
            Write-Host "Smart Office started successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error starting Smart Office: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    elseif ($key.VirtualKeyCode -eq 57) {
        Write-Host "Rebooting..." -ForegroundColor Cyan
        Restart-Computer -Force
    }
    else {
        Write-Host "Exiting..." -ForegroundColor Cyan
    }
}
finally {
    if ($monitorJob) {
        Stop-Job -Job $monitorJob -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $monitorJob -ErrorAction SilentlyContinue | Out-Null
    }
}
