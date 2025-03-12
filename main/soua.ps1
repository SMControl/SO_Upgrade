# Initialize script start time
$startTime = Get-Date
function Show-Intro {
    Write-Host "Smart Office - Upgrade Assistant - Version 1.143" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------------------------"
    Write-Host ""
}
# Changes
# - Part 7&14 total re-do

# Set the working directory
$workingDir = "C:\winsm"
if (-not (Test-Path $workingDir -PathType Container)) {
    try {
        New-Item -Path $workingDir -ItemType Directory -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Error: Unable to create directory $workingDir" -ForegroundColor Red
        exit
    }
}

Set-Location -Path $workingDir


# Part 1 - Check for Admin Rights
# -----
Clear-Host
Show-Intro
Write-Host "[Part 1/15] System Pre-Checks" -ForegroundColor Cyan
Write-Host "[#______________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "Error: Administrator rights required to run this script. Exiting." -ForegroundColor Red
    pause
    exit
}

# Download the SmartOffice_Upgrade_Assistant.exe and save it to C:\winsm
$assistantExeUrl = "https://github.com/SMControl/SO_UC/blob/main/SmartOffice_Upgrade_Assistant.exe?raw=true"
$assistantExeDestinationPath = "C:\winsm\SmartOffice_Upgrade_Assistant.exe"
if (-Not (Test-Path $assistantExeDestinationPath)) {
    Invoke-WebRequest -Uri $assistantExeUrl -OutFile $assistantExeDestinationPath
} else {
    #Write-Host "$assistantExeDestinationPath already exists. Skipping download."
}

# Download the SO_UC.exe and save it to C:\winsm
$soucExeUrl = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe?raw=true"
$soucExeDestinationPath = "C:\winsm\SO_UC.exe"
if (-Not (Test-Path $soucExeDestinationPath)) {
    Invoke-WebRequest -Uri $soucExeUrl -OutFile $soucExeDestinationPath
} else {
    #Write-Host "$soucExeDestinationPath already exists. Skipping download."
}


# Part 2 - Check for Running SO Processes
# -----
Clear-Host
Show-Intro
Write-Host "[Part 2/15] Checking processes" -ForegroundColor Cyan
Write-Host "[##_____________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
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
    }
}

# Part 3 - SO_UC.exe
# -----
Clear-Host
Show-Intro
Write-Host "[Part 3/15] Checking for Setup Files. Please Wait." -ForegroundColor Cyan
Write-Host "[###____________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
# Section A - Retrieve .exe links from the webpage
# PartVersion 1.00
# -----
$downloadDirectory = "C:\winsm\SmartOffice_Installer"
$webpageUrl = "https://www.stationmaster.com/downloads/"

if (-not (Test-Path $downloadDirectory)) {
    # Ensure download directory exists
    New-Item -ItemType Directory -Path $downloadDirectory | Out-Null
}

# Retrieve Setup.exe links
$setupLinks = (Invoke-WebRequest -Uri $webpageUrl).Links |
              Where-Object { $_.href -match "^https://www\.stationmaster\.com/Download/Setup\d+\.exe$" } |
              ForEach-Object { $_.href }

# -----
# Section B - Filter for the highest two versions of Setup.exe
# PartVersion 1.01
# -----
$setupLinksWithVersions = $setupLinks | ForEach-Object {
    [PSCustomObject]@{
        Link = $_
        Version = [regex]::Match($_, "Setup(\d+)\.exe").Groups[1].Value -as [int]
    }
}
$highestTwoLinks = $setupLinksWithVersions |
                   Sort-Object -Property Version -Descending |
                   Select-Object -First 2 -ExpandProperty Link

# -----
# Section C - Download the highest two versions if not already present
# PartVersion 1.03
# -----
foreach ($downloadLink in $highestTwoLinks) {
    $originalFilename = $downloadLink.Split('/')[-1]
    $destinationPath = Join-Path -Path $downloadDirectory -ChildPath $originalFilename

    # Check the size of the file on the server
    $request = [System.Net.HttpWebRequest]::Create($downloadLink)
    $request.Method = "HEAD"
    $request.UserAgent = "Mozilla/5.0"

    try {
        $response = $request.GetResponse()
        $contentLength = $response.ContentLength
        $response.Close()
    } catch {
        Write-Error "Error fetching file information for $originalFilename $($_.Exception.Message)"
        continue
    }

    # Compare size with existing file
    $matchingFile = Get-ChildItem -Path $downloadDirectory -Filter $originalFilename |
                    Where-Object { $_.Length -eq $contentLength }

    if (-not $matchingFile) {
        Write-Host "Downloading new version: $originalFilename..."
        Invoke-WebRequest -Uri $downloadLink -OutFile $destinationPath -UseBasicParsing
    }
}

# -----
# Section D - Delete older downloads, keeping the latest two
# PartVersion 1.01
# -----
$downloadedFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe" |
                   Sort-Object LastWriteTime -Descending
if ($downloadedFiles.Count -gt 2) {
    $filesToDelete = $downloadedFiles | Select-Object -Skip 2
    Remove-Item -Path ($filesToDelete | ForEach-Object { $_.FullName }) -Force
}


# Part 4 - Check for Firebird Installation
# -----
Clear-Host
Show-Intro
Write-Host "[Part 4/15] Checking for Firebird installation" -ForegroundColor Cyan
Write-Host "[####___________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"

if (-not (Test-Path $firebirdDir)) {
    Write-Host "Firebird not found. Installing Firebird..." -ForegroundColor Yellow
    try {
        # Start a new PowerShell process to run the installer
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"irm $firebirdInstallerURL | iex`"" -Wait
        Write-Host "Firebird installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error installing Firebird: $_" -ForegroundColor Red
        exit
    }
} else {
    #Write-Host "Firebird is already installed." -ForegroundColor Green
}


# Part 5 - Stop SMUpdates if Running
# -----
Clear-Host
Show-Intro
Write-Host "[Part 5/15] Stopping SMUpdates if running" -ForegroundColor Cyan
Write-Host "[#####__________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
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

# Part 6 - Manage SO Live Sales Service
# -----
Clear-Host
Show-Intro
Write-Host "[Part 6/15] Managing SO Live Sales service" -ForegroundColor Cyan
Write-Host "[######_________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
$ServiceName = "srvSOLiveSales"
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
$wasRunning = $false

if ($service) {
    if ($service.Status -eq 'Running') {
        $wasRunning = $true
        Write-Host "Stopping $ServiceName service..." -ForegroundColor Yellow
        try {
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            Set-Service -Name $ServiceName -StartupType Disabled
            Write-Host "$ServiceName service stopped and set to Disabled." -ForegroundColor Green
        } catch {
            Write-Host "Error stopping service '$ServiceName': $_" -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "$ServiceName service is not running." -ForegroundColor Yellow
    }
} else {
    #Write-Host "$ServiceName service not found." -ForegroundColor Yellow
}

# Part 7 - Manage PDTWiFi Processes
# -----
# PartVersion_1.02
# - total redo
Clear-Host
Show-Intro
Write-Host "[Part 7/15] Managing PDTWiFi processes" -ForegroundColor Cyan
Write-Host "[#######________]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

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

# Part 8 - Wait for Single Instance of Firebird.exe
# -----
Clear-Host
Show-Intro
Write-Host "[Part 8/15] Waiting for a single instance of Firebird" -ForegroundColor Cyan
Write-Host "[########_______]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

$setupDir = "$workingDir\SmartOffice_Installer"
if (-not (Test-Path $setupDir -PathType Container)) {
    Write-Host "Error: Setup directory '$setupDir' does not exist." -ForegroundColor Red
    exit
}

function WaitForSingleFirebirdInstance {
    $firebirdProcesses = Get-Process -Name "firebird" -ErrorAction SilentlyContinue
    while ($firebirdProcesses.Count -gt 1) {
        Write-Host "Warning: Multiple instances of 'firebird.exe' are running. Currently: $($firebirdProcesses.Count)" -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        $firebirdProcesses = Get-Process -Name "firebird" -ErrorAction SilentlyContinue
    }
}

WaitForSingleFirebirdInstance

# Part 9 - Launch Setup
# PartVersion 1.04
# -----
# Improved terminal selection menu with colors and table formatting
Clear-Host
Show-Intro
Write-Host "[Part 9/15] Launching SO setup..." -ForegroundColor Cyan
Write-Host "[#########______]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

# Get all setup executables in the SmartOffice_Installer directory
$setupExes = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "*.exe"

if ($setupExes.Count -eq 0) {
    Write-Host "Error: No executable (.exe) found in 'C:\winsm\SmartOffice_Installer'." -ForegroundColor Red
    exit
} elseif ($setupExes.Count -eq 1) {
    # Only one file found, proceed without asking the user
    $selectedExe = $setupExes[0]
    Write-Host "Found setup: $($selectedExe.Name)" -ForegroundColor Green
} else {
    # Sort executables by version (numeric part in the name) ascending
    $setupExes = $setupExes | Sort-Object {
        [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int]
    }

    # Multiple setup files found, present a terminal selection menu
    Write-Host "`nPlease select the setup to run:`n" -ForegroundColor Yellow
    Write-Host ("{0,-5} {1,-30} {2,-20} {3,-10}" -f "No.", "Executable Name", "Date Modified", "Version") -ForegroundColor White
    Write-Host ("{0,-5} {1,-30} {2,-20} {3,-10}" -f "---", "------------------------------", "------------", "-------") -ForegroundColor Gray

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
        Write-Host "Operation cancelled. Exiting." -ForegroundColor Red
        exit
    }

    # Validate the selection
    if ($selection -match '^\d+$' -and $selection -ge 1 -and $selection -le $setupExes.Count) {
        $selectedExe = $setupExes[$selection - 1]  # Convert to 0-based index
        Write-Host "Selected setup executable: $($selectedExe.Name)" -ForegroundColor Green
    } else {
        Write-Host "Invalid selection. Exiting." -ForegroundColor Red
        exit
    }
}

# Launch the selected setup executable
try {
    Write-Host "Starting executable: $($selectedExe.Name) ..." -ForegroundColor Cyan
    Start-Process -FilePath $selectedExe.FullName -Wait
} catch {
    Write-Host "Error starting setup executable: $_" -ForegroundColor Red
    exit
}


# Part 10 - Wait for User Confirmation
# -----
Clear-Host
Show-Intro
Write-Host "[Part 10/15] Post Upgrade" -ForegroundColor Cyan
Write-Host "[##########_____]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
# Stop monitoring SMUpdates process
Stop-Job -Job $monitorJob
Remove-Job -Job $monitorJob
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

# Part 11 - Set Permissions for SM Folder
# -----
Clear-Host
Show-Intro
Write-Host "[Part 11/15] Setting permissions for Stationmaster folder. Please Wait..." -ForegroundColor Cyan
Write-Host "[###########____]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
try {
    & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for SM folder: $_" -ForegroundColor Red
}

# Part 12 - Set Permissions for Firebird Folder
# -----
Clear-Host
Show-Intro
Write-Host "[Part 12/15] Setting permissions for Firebird folder. Please Wait..." -ForegroundColor Cyan
Write-Host "[############___]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
try {
    & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for Firebird folder." -ForegroundColor Red
}

# Part 13 - Revert SO Live Sales Service
# -----
Clear-Host
Show-Intro
Write-Host "[Part 13/15] Reverting SO Live Sales service" -ForegroundColor Cyan
Write-Host "[#############__]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1
if ($wasRunning) {
    try {
        Write-Host "Setting $ServiceName service back to Automatic startup..." -ForegroundColor Yellow
        Set-Service -Name $ServiceName -StartupType Automatic
        Start-Service -Name $ServiceName
        Write-Host "$ServiceName service reverted to its previous state." -ForegroundColor Green
    } catch {
        Write-Host "Error reverting service '$ServiceName': $_" -ForegroundColor Red
    }
} else {
    Write-Host "$ServiceName service was not running before, so no action needed." -ForegroundColor Yellow
}

# Part 14 - Revert PDTWiFi Processes
# -----
# PartVersion 1.02
# - total re-do
Clear-Host
Show-Intro
Write-Host "[Part 14/15] Reverting PDTWiFi processes" -ForegroundColor Cyan
Write-Host "[##############_]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

# Section A - Recall and Revert PDTWiFi
if ($PDTWiFiStates[$PDTWiFi] -eq "Running") {
    Start-Process "C:\Program Files (x86)\StationMaster\PDTWiFi.exe"
    Write-Host "$PDTWiFi started." -ForegroundColor Green
} else {
    Write-Host "$PDTWiFi was not running, no action taken." -ForegroundColor Yellow
}

# Section B - Recall and Revert PDTWiFi64
if ($PDTWiFiStates[$PDTWiFi64] -eq "Running") {
    Start-Process "C:\Program Files (x86)\StationMaster\PDTWiFi64.exe"
    Write-Host "$PDTWiFi64 started." -ForegroundColor Green
} else {
    Write-Host "$PDTWiFi64 was not running, no action taken." -ForegroundColor Yellow
}


# Part 15 - Clean up and Finish Script
# -----
Clear-Host
Show-Intro
Write-Host "[Part 15/15] Clean up and finish" -ForegroundColor Cyan
Write-Host "[###############]" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

# Run SO_UC.exe if it's Task doesn't exist.
$taskExists = Get-ScheduledTask -TaskName "SO InstallerUpdates" -ErrorAction SilentlyContinue

if (-not $taskExists) {
    Start-Process -FilePath "C:\winsm\SO_UC.exe" -Wait
} else {
    #Write-Output "Scheduled Task 'SO InstallerUpdates' already exists. Skipping execution."
}

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host " "
Write-Host "Smart Office $(Split-Path -Leaf $selectedExe.Name) Installed"
Write-Host " "
Write-Host "Completed in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
Write-Host " "

Write-Host "Press Enter to start Smart Office, or any other key to exit."
$key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
if ($key.VirtualKeyCode -eq 13) {
    Start-Process "C:\Program Files (x86)\StationMaster\Sm32.exe"
} else {
    Write-Host "Exiting..."
}
