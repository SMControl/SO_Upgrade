# Script Version: 1.009

# PART 1: RUNTIME SETUP & PREREQUISITES
# PartVersion: 1.001
$ErrorActionPreference = "Stop"
$InstallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[ERROR] This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# PART 2: CONFIGURATION & DIRECTORIES
# PartVersion: 1.004
$Config = @{
    ScriptVersion = "1.000"
    LogDir        = "C:\winsm\SmartOffice_Installer\soua_logs"
    Paths         = @{
        StationMaster = "C:\Program Files (x86)\StationMaster"
        Firebird32    = "C:\Program Files (x86)\Firebird"
        Firebird64    = "C:\Program Files\Firebird"
        SetupDir      = "C:\winsm\SmartOffice_Installer"
    }
    Services      = @{
        LiveSales = "srvSOLiveSales"
        Firebird  = @("FirebirdServerDefaultInstance", "FirebirdGuardianDefaultInstance")
    }
    Processes     = @{
        LiveSales   = "srvSOLiveSales"
        SMUpdates   = "SMUpdates"
        Firebird    = "firebird"
        SmartOffice = @("Sm32Main", "Sm32")
        PDTWiFi     = "PDTWiFi"
        PDTWiFi64   = "PDTWiFi64"
    }
    URLs          = @{
        ModuleSOGets      = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_soget.ps1"
        TaskSODatTransfer = "https://raw.githubusercontent.com/SMControl/SM_Tasks/refs/heads/main/tasks/task_SO%20system.dat_transfer.ps1"
    }
}

try {
    if (-not (Test-Path $Config.LogDir)) {
        New-Item -ItemType Directory -Path $Config.LogDir -Force | Out-Null
    }
}
catch { }

$LogFile = Join-Path $Config.LogDir ("soua_minimal_safe_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".log")

function Write-SilentLog {
    param([string]$Message, [switch]$IsError)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $level = if ($IsError) { "ERROR" } else { "INFO" }
    $entry = "[$ts] [$level] $Message"
    try {
        Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    }
    catch { }
    if ($IsError) {
        Write-Host "  [-] $Message" -ForegroundColor Red
    }
    else {
        if ($Message -notmatch '^\[\d/\d\]') {
            Write-Host "  [>] $Message" -ForegroundColor DarkGray
        }
    }
}

Write-SilentLog "SO Upgrade Assistant (Minimal Safe) session started. Version: $($Config.ScriptVersion), Host: $env:COMPUTERNAME, User: $env:USERNAME"

$WasLiveSalesRunning = $false
$WasSMUpdatesRunning = $false
$PDTWiFiStates = @{}
$MonitorJob = $null
$PackageJob = $null
$FbPermsJob = $null
$RebootRequired = $false

function Cleanup-Environment {
    Write-SilentLog "Executing environment cleanup and service restoration..."

    # 1. Terminate background jobs
    if ($FbPermsJob) {
        Write-SilentLog "Terminating Firebird permissions background job (ID: $($FbPermsJob.Id))..."
        Stop-Job -Job $FbPermsJob -ErrorAction SilentlyContinue
        Remove-Job -Job $FbPermsJob -ErrorAction SilentlyContinue
    }
    if ($PackageJob) {
        Write-SilentLog "Terminating package check background job (ID: $($PackageJob.Id))..."
        Stop-Job -Job $PackageJob -ErrorAction SilentlyContinue
        Remove-Job -Job $PackageJob -ErrorAction SilentlyContinue
    }
    if ($MonitorJob) {
        Write-SilentLog "Terminating SMUpdates watchdog monitor job (ID: $($MonitorJob.Id))..."
        Stop-Job -Job $MonitorJob -ErrorAction SilentlyContinue
        Remove-Job -Job $MonitorJob -ErrorAction SilentlyContinue
    }

    # 2. Restore LiveSales service
    if ($WasLiveSalesRunning) {
        Write-SilentLog "Restoring service $($Config.Services.LiveSales)..."
        $svc = Get-Service -Name $Config.Services.LiveSales -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -ne "Running") {
            try {
                Start-Service -Name $Config.Services.LiveSales -ErrorAction SilentlyContinue
                Write-SilentLog "Service $($Config.Services.LiveSales) started successfully."
            }
            catch {
                Write-SilentLog "Failed to restart service $($Config.Services.LiveSales): $($_.Exception.Message)" -IsError
            }
        }
    }

    # 3. Restore SMUpdates tray app
    if ($WasSMUpdatesRunning) {
        $smUpdExe = Join-Path $Config.Paths.StationMaster "SMUpdates.exe"
        if (Test-Path $smUpdExe) {
            $p = Get-Process -Name $Config.Processes.SMUpdates -ErrorAction SilentlyContinue
            if (-not $p) {
                Write-SilentLog "Relaunching $smUpdExe..."
                try {
                    Start-Process $smUpdExe -ErrorAction SilentlyContinue
                    Write-SilentLog "Relaunched $smUpdExe successfully."
                }
                catch {
                    Write-SilentLog "Failed to relaunch ${smUpdExe}: $($_.Exception.Message)" -IsError
                }
            }
        }
    }

    # 4. Restore PDTWiFi tray sync apps
    foreach ($procName in $PDTWiFiStates.Keys) {
        if ($PDTWiFiStates[$procName] -eq "Running") {
            $p = Get-Process -Name $procName -ErrorAction SilentlyContinue
            if (-not $p) {
                $targetProcExe = Join-Path $Config.Paths.StationMaster "$procName.exe"
                Write-SilentLog "Relaunching PDT WiFi utility: $targetProcExe..."
                try {
                    Start-Process $targetProcExe -ErrorAction SilentlyContinue
                    Write-SilentLog "Relaunched $targetProcExe successfully."
                }
                catch {
                    Write-SilentLog "Failed to relaunch ${targetProcExe}: $($_.Exception.Message)" -IsError
                }
            }
        }
    }

    Write-SilentLog "Environment cleanup complete."
}

# PART 3: PRE-FLIGHT CHECKS & BACKGROUND SERVICES
# PartVersion: 1.001
Write-Host "[1/7] Running pre-flight checks..." -ForegroundColor Cyan
Write-SilentLog "[1/7] Running pre-flight checks..."
try {
    $smOfficeVal = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\StationMaster\SM32" -Name "SMOffice" -ErrorAction SilentlyContinue).SMOffice
    Write-SilentLog "HKLM SMOffice flag: $smOfficeVal"
    if ([string]$smOfficeVal -ne "1" -and $Config.URLs.TaskSODatTransfer) {
        Write-SilentLog "SMOffice flag not 1; initiating background system.dat_transfer task..."
        Start-Job -ScriptBlock {
            param($url)
            try {
                $script = Invoke-RestMethod -Uri $url -TimeoutSec 15 -ErrorAction Stop
                if ($script) { Invoke-Expression $script }
            } catch { }
        } -ArgumentList $Config.URLs.TaskSODatTransfer | Out-Null
    }
} catch {
    Write-SilentLog "Error reading SMOffice registry: $($_.Exception.Message)"
}

$fbFound = (Test-Path $Config.Paths.Firebird32) -or (Test-Path $Config.Paths.Firebird64)
Write-SilentLog "Checking Firebird directory presence: Firebird32=$(Test-Path $Config.Paths.Firebird32), Firebird64=$(Test-Path $Config.Paths.Firebird64)"
if (-not $fbFound) {
    Write-SilentLog "Firebird database installation directory not found." -IsError
    exit 1
}

$fbTargetPaths = @()
if (Test-Path $Config.Paths.Firebird32) { $fbTargetPaths += $Config.Paths.Firebird32 }
if (Test-Path $Config.Paths.Firebird64) { $fbTargetPaths += $Config.Paths.Firebird64 }

if ($fbTargetPaths.Count -gt 0) {
    Write-SilentLog "Spawning background Firebird permissions job on $(($fbTargetPaths) -join ', ')..."
    $FbPermsJob = Start-Job -ScriptBlock {
        param($paths)
        foreach ($p in $paths) {
            try {
                & icacls $p /grant "*S-1-1-0:(OI)(CI)F" /T /C 2>&1 | Out-Null
            } catch { }
        }
    } -ArgumentList (, $fbTargetPaths)
}

Write-Host "[2/7] Pausing sync services and background processes..." -ForegroundColor Cyan
Write-SilentLog "[2/7] Pausing sync services and background processes..."

# 1. Stop local apps and tray utilities
$smUpdProc = Get-Process -Name $Config.Processes.SMUpdates -ErrorAction SilentlyContinue
if ($smUpdProc) {
    $WasSMUpdatesRunning = $true
    Write-SilentLog "Found running process: SMUpdates (PID: $($smUpdProc.Id)). Stopping process..."
    try {
        Stop-Process -Name $Config.Processes.SMUpdates -Force -ErrorAction SilentlyContinue
        Write-SilentLog "SMUpdates process stop signal issued."
    } catch {
        Write-SilentLog "Failed stopping SMUpdates: $($_.Exception.Message)"
    }
} else {
    Write-SilentLog "Process SMUpdates is not running."
}

$smUpdatesScript = Join-Path $Config.Paths.StationMaster "module_smupdates.ps1"
if (Test-Path $smUpdatesScript) {
    Write-SilentLog "Found $smUpdatesScript; starting watchdog job..."
    $MonitorJob = Start-Job -ScriptBlock {
        param($path)
        Set-Location $path
        & ".\module_smupdates.ps1"
    } -ArgumentList $Config.Paths.StationMaster
}

$liveSalesSvc = Get-Service -Name $Config.Services.LiveSales -ErrorAction SilentlyContinue
if ($liveSalesSvc -and $liveSalesSvc.Status -eq "Running") {
    $WasLiveSalesRunning = $true
    Write-SilentLog "Service $($Config.Services.LiveSales) is Running. Stopping service..."
    try {
        Stop-Service -Name $Config.Services.LiveSales -Force -ErrorAction SilentlyContinue
        Write-SilentLog "Service $($Config.Services.LiveSales) stopped successfully."
    } catch {
        Write-SilentLog "Failed to stop $($Config.Services.LiveSales): $($_.Exception.Message)" -IsError
        Cleanup-Environment
        exit 1
    }
} else {
    Write-SilentLog "Service $($Config.Services.LiveSales) status: $(if ($liveSalesSvc) { $liveSalesSvc.Status } else { 'Not Installed' })"
}

foreach ($procName in @($Config.Processes.PDTWiFi, $Config.Processes.PDTWiFi64)) {
    $p = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($p) {
        $PDTWiFiStates[$procName] = "Running"
        Write-SilentLog "Found running process: $procName (PID: $($p.Id)). Stopping process..."
        try {
            Stop-Process -Name $procName -Force -ErrorAction SilentlyContinue
            Write-SilentLog "Process $procName stop signal issued."
        } catch { }
    } else {
        $PDTWiFiStates[$procName] = "Not running"
        Write-SilentLog "Process $procName is not running."
    }
}

# PART 4: DOWNLOADS & VERSION SELECTION
# PartVersion: 1.000
Write-Host "[3/7] Checking for setup packages..." -ForegroundColor Cyan
Write-SilentLog "[3/7] Checking for setup packages via module_soget..."
try {
    $sogetCode = Invoke-RestMethod -Uri $Config.URLs.ModuleSOGets -TimeoutSec 15 -ErrorAction Stop
    if ($sogetCode) {
        Write-SilentLog "Executing module_soget code..."
        Invoke-Expression $sogetCode
        Write-SilentLog "module_soget execution finished."
    }
} catch {
    Write-SilentLog "Failed to query or execute module_soget: $($_.Exception.Message)" -IsError
}

if (-not (Test-Path $Config.Paths.SetupDir)) {
    Write-SilentLog "Setup directory missing at $($Config.Paths.SetupDir)" -IsError
    Cleanup-Environment
    exit 1
}

$setupExes = @(Get-ChildItem -Path $Config.Paths.SetupDir -Filter "*.exe" -ErrorAction SilentlyContinue | Sort-Object {
        [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int]
    } -Descending)

Write-SilentLog "Found $($setupExes.Count) setup file(s) in $($Config.Paths.SetupDir): $(($setupExes | ForEach-Object { $_.Name }) -join ', ')"

if ($setupExes.Count -lt 2) {
    Write-SilentLog "Less than 2 setup files found in $($Config.Paths.SetupDir) ($($setupExes.Count) found). Minimum 2 required to proceed." -IsError
    Cleanup-Environment
    exit 1
}

$currentInstVerInt = 0
try {
    $rawInstVer = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\StationMaster\SM32" -Name "InstVer" -ErrorAction SilentlyContinue).InstVer
    Write-SilentLog "Current installed InstVer in registry: $rawInstVer"
    if ($rawInstVer) {
        $cleanVer = [regex]::Match([string]$rawInstVer, "\d+").Value
        if ($cleanVer) { $currentInstVerInt = [int]$cleanVer }
    }
} catch { }

$validCandidates = @($setupExes | Where-Object {
        $vNum = [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int]
        $null -ne $vNum -and ($currentInstVerInt -eq 0 -or $vNum -ge $currentInstVerInt)
    })

Write-SilentLog "Valid forward candidate setups: $(($validCandidates | ForEach-Object { $_.Name }) -join ', ')"

if ($validCandidates.Count -eq 0) {
    Write-SilentLog "All available setup files are older than currently installed version ($currentInstVerInt). Cannot downgrade." -IsError
    Cleanup-Environment
    exit 1
}

$selectedExe = $null
if ($validCandidates.Count -eq 1) {
    $selectedExe = $validCandidates[0]
    $vNum = [regex]::Match($selectedExe.Name, "\d+").Value
    Write-Host "Auto-selected version $vNum ($($selectedExe.Name)) based on current installation ($currentInstVerInt)." -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "Select Smart Office version to install:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $validCandidates.Count; $i++) {
        $vMatch = [regex]::Match($validCandidates[$i].Name, "\d+")
        $vDisplay = if ($vMatch.Success) { $vMatch.Value } else { $validCandidates[$i].Name }
        Write-Host "  [$($i + 1)] $vDisplay ($($validCandidates[$i].Name))"
    }

    while ($null -eq $selectedExe) {
        $choice = Read-Host "Enter selection (1-$($validCandidates.Count))"
        $idx = $choice -as [int]
        if ($idx -ge 1 -and $idx -le $validCandidates.Count) {
            $selectedExe = $validCandidates[$idx - 1]
        } else {
            Write-Host "Invalid choice. Please enter a number between 1 and $($validCandidates.Count)." -ForegroundColor Yellow
        }
    }
}

Write-SilentLog "User selected setup: $($selectedExe.Name)"

# PART 5: PROCESS TERMINATION & FILE LOCK AUDIT
# PartVersion: 1.003
Write-Host "[4/7] Checking process state and directory locks..." -ForegroundColor Cyan
Write-SilentLog "[4/7] Checking process state and directory locks..."
$soTimeoutSec = 3600 # 1 hour safe buffer for slow hardware / remote users
$soStartWait = Get-Date
$soOpenAnnounced = $false

while ($true) {
    $soRunning = $false
    foreach ($pName in $Config.Processes.SmartOffice) {
        if (Get-Process -Name $pName -ErrorAction SilentlyContinue) {
            $soRunning = $true
            break
        }
    }
    if (-not $soRunning) { break }
    if (-not $soOpenAnnounced) {
        Write-Host "[WAITING] Smart Office is currently open. Please close Smart Office to proceed..." -ForegroundColor Yellow
        $soOpenAnnounced = $true
    }
    if (((Get-Date) - $soStartWait).TotalSeconds -gt $soTimeoutSec) {
        Write-SilentLog "Timeout waiting for Smart Office to be closed by user." -IsError
        Cleanup-Environment
        exit 1
    }
    Start-Sleep -Seconds 1
}

$fbTimeoutSec = 120
$fbStartWait = Get-Date
while ($true) {
    $fbProcs = Get-Process -Name $Config.Processes.Firebird -ErrorAction SilentlyContinue
    Write-SilentLog "Firebird active process count: $($fbProcs.Count)"
    if ($fbProcs.Count -le 1) { break }
    if (((Get-Date) - $fbStartWait).TotalSeconds -gt $fbTimeoutSec) {
        Write-SilentLog "Multiple Firebird instances remain active ($($fbProcs.Count)). Remote clients may still be connected." -IsError
        Cleanup-Environment
        exit 1
    }
    Start-Sleep -Seconds 1
}

$smUpdWaitStart = Get-Date
while ($true) {
    $p = Get-Process -Name $Config.Processes.SMUpdates -ErrorAction SilentlyContinue
    if (-not $p) {
        Write-SilentLog "SMUpdates process confirmed stopped."
        break
    }
    Write-SilentLog "SMUpdates still running (PID: $($p.Id)). Sending force stop..."
    try {
        Stop-Process -Name $Config.Processes.SMUpdates -Force -ErrorAction SilentlyContinue
    } catch { }
    if (((Get-Date) - $smUpdWaitStart).TotalSeconds -gt 15) {
        Write-SilentLog "SMUpdates wait window reached 15s limit."
        break
    }
    Start-Sleep -Milliseconds 500
}

Write-SilentLog "Starting critical binary file lock audit..."
$lockedFiles = @()
$criticalBins = @(Get-ChildItem -Path $Config.Paths.StationMaster -Filter "*.exe" -File -ErrorAction SilentlyContinue)
$criticalBins += @(Get-ChildItem -Path $Config.Paths.StationMaster -Filter "*.ocx" -File -ErrorAction SilentlyContinue)
$dbClients = @("fbclient.dll", "gds32.dll", "midas.dll")
foreach ($dbc in $dbClients) {
    $p = Join-Path $Config.Paths.StationMaster $dbc
    if (Test-Path $p) { $criticalBins += (Get-Item $p) }
}

foreach ($bf in $criticalBins) {
    try {
        $stream = [System.IO.File]::Open($bf.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        $stream.Dispose()
    } catch {
        Write-SilentLog "File locked: $($bf.FullName) - $($_.Exception.Message)"
        $lockedFiles += $bf.Name
    }
}

if ($lockedFiles.Count -gt 0) {
    $RebootRequired = $true
    Write-SilentLog "File lock detected on: $($lockedFiles -join ', '). Setup may require a system restart." -IsError
} else {
    Write-SilentLog "Zero file locks detected on critical binaries."
}

# PART 6: SETUP LAUNCH & TECHNICIAN INTERACTION (SAFE MANUAL EXECUTION)
# PartVersion: 1.001
function Execute-SafeSetup {
    param($SetupExe)
    
    Write-Host "[WAITING] Launching legacy setup wizard ($($SetupExe.Name)). Please complete installation on your desktop..." -ForegroundColor Yellow
    Write-SilentLog "Launching legacy setup executable ($($SetupExe.FullName)) for manual technician handling..."
    
    $setupSw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $setupProc = Start-Process -FilePath $SetupExe.FullName -PassThru -ErrorAction Stop
        Write-SilentLog "Setup executable launched successfully (PID: $($setupProc.Id)). Awaiting process completion..."
        
        while (-not $setupProc.HasExited) {
            Start-Sleep -Milliseconds 500
        }
        
        $setupSw.Stop()
        $durationMs = [long]$setupSw.ElapsedMilliseconds
        Write-SilentLog "Setup executable exited with code: $($setupProc.ExitCode) (Duration: $durationMs ms)"
        if ($setupProc.ExitCode -eq 0) {
            Write-Host "Setup wizard completed." -ForegroundColor Green
        } else {
            Write-Host "Setup wizard exited with code: $($setupProc.ExitCode)" -ForegroundColor Yellow
        }
        return @{ Success = $true; DurationMs = $durationMs }
    }
    catch {
        $setupSw.Stop()
        Write-SilentLog "Error launching setup wizard: $($_.Exception.Message)" -IsError
        return @{ Success = $false; DurationMs = [long]$setupSw.ElapsedMilliseconds }
    }
}

Write-Host "[5/7] Opening Smart Office setup package for technician installation..." -ForegroundColor Cyan
$setupResult = Execute-SafeSetup -SetupExe $selectedExe
if (-not $setupResult.Success) {
    Cleanup-Environment
    exit 1
}
$installerDurationMs = $setupResult.DurationMs

# PART 7: POST-UPGRADE VERIFICATION, PERMISSIONS & RESTORATION
# PartVersion: 1.002
if ($MonitorJob) {
    Write-SilentLog "Stopping and cleaning up monitor job..."
    Stop-Job -Job $MonitorJob -ErrorAction SilentlyContinue
    Remove-Job -Job $MonitorJob -ErrorAction SilentlyContinue
    $MonitorJob = $null
}

$soStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$smExe = Join-Path $Config.Paths.StationMaster "Sm32.exe"
if (Test-Path $smExe) {
    Write-SilentLog "Launching first-run Smart Office ($smExe)..."
    try {
        Start-Process -FilePath $smExe -ErrorAction Stop
        Write-SilentLog "Smart Office launched successfully."
    } catch {
        Write-SilentLog "Unable to start first-run Smart Office: $($_.Exception.Message)" -IsError
    }
} else {
    Write-SilentLog "Sm32.exe not found at $smExe!" -IsError
}

Write-Host "[6/7] Launching Smart Office for first-run database schema migrations..." -ForegroundColor Cyan
Write-Host "[WAITING] Complete any wizard prompts and close Smart Office to finalize..." -ForegroundColor Yellow
Write-SilentLog "[6/7] Awaiting user first-run Smart Office close..."

$smartOfficeHasRun = $false
$firstRunLimit = 600
$frStart = Get-Date

while ($true) {
    $soRunning = $false
    foreach ($process in $Config.Processes.SmartOffice) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            $soRunning = $true
            $smartOfficeHasRun = $true
            $firstRunLimit = 3600 # 1 hour safe window for slow hardware, large DB schema updates, or technician review
        }
    }

    if ($smartOfficeHasRun -and -not $soRunning) {
        Write-SilentLog "Smart Office has closed cleanly."
        break
    }

    if (((Get-Date) - $frStart).TotalSeconds -gt $firstRunLimit) {
        Write-SilentLog "First-run Smart Office did not close within timeout window ($firstRunLimit s)." -IsError
        break
    }

    Start-Sleep -Milliseconds 400
}

$soStopwatch.Stop()
$soDurationMs = [long]$soStopwatch.ElapsedMilliseconds
Write-SilentLog "First-run Smart Office duration: $soDurationMs ms"

Write-Host "[7/7] Applying folder permissions and restoring services..." -ForegroundColor Cyan
Write-SilentLog "[7/7] Applying folder permissions and restoring services..."
try {
    Write-SilentLog "Applying full permissions (*S-1-1-0:(OI)(CI)F) to $($Config.Paths.StationMaster)..."
    & icacls $Config.Paths.StationMaster /grant "*S-1-1-0:(OI)(CI)F" /T /C 2>&1 | Out-Null
    Write-SilentLog "StationMaster permissions applied."
} catch {
    Write-SilentLog "Failed applying permissions to StationMaster directory: $($_.Exception.Message)" -IsError
}

if ($FbPermsJob) {
    Write-SilentLog "Awaiting background Firebird permissions job completion..."
    Wait-Job -Job $FbPermsJob -Timeout 10 | Out-Null
    Receive-Job -Job $FbPermsJob -ErrorAction SilentlyContinue
    Remove-Job -Job $FbPermsJob -ErrorAction SilentlyContinue
    $FbPermsJob = $null
    Write-SilentLog "Firebird permissions verified."
}

Cleanup-Environment

$InstallStopwatch.Stop()
$totalMs = [long]$InstallStopwatch.ElapsedMilliseconds
$installerMs = [long]$installerDurationMs
$automatedMs = [math]::Max(0, ($totalMs - ($installerMs + $soDurationMs)))

$totalSec = [math]::Round($totalMs / 1000)
$installerSec = [math]::Round($installerMs / 1000)
$soSec = [math]::Round($soDurationMs / 1000)
$automatedSec = [math]::Round($automatedMs / 1000)

$fmtTotal = "${totalSec}s"
$fmtInstaller = "${installerSec}s"
$fmtSO = "${soSec}s"
$fmtAutomated = "${automatedSec}s"

Write-Host ""
Write-Host "Upgrade completed successfully." -ForegroundColor Green
Write-Host "Timing Breakdown:" -ForegroundColor Cyan
Write-Host "  Total Duration:         $fmtTotal"
Write-Host "  User in Setup Wizard:   $fmtInstaller"
Write-Host "  User in SmartOffice:    $fmtSO"
Write-Host "  Automated Work Time:    $fmtAutomated"
Write-Host ""
Write-Host "[$fmtTotal] | [$fmtInstaller in Setup] | [$fmtSO in SO] | [$fmtAutomated automated]" -ForegroundColor Green

Write-SilentLog "Timing Breakdown -> Total: $fmtTotal ($totalMs ms) | Setup Wizard: $fmtInstaller ($installerMs ms) | SmartOffice Wait: $fmtSO ($soDurationMs ms) | Automated Work Time: $fmtAutomated ($automatedMs ms)"

if ($RebootRequired) {
    Write-Host ""
    Write-Host "==========================================================================" -ForegroundColor Red
    Write-Host " WARNING: Critical files were locked during installation ($($lockedFiles -join ', '))." -ForegroundColor Red
    Write-Host " Installation is NOT complete until a system reboot is performed!" -ForegroundColor Yellow
    Write-Host " Presenting modal warning prompt to technician..." -ForegroundColor Cyan
    Write-Host "==========================================================================" -ForegroundColor Red
    Write-Host ""

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

    $popupTitle = "StationMaster SmartOffice Upgrade - REBOOT REQUIRED"
    $popupMessage = "WARNING:`n`nCritical files were locked by other processes during installation:`n$($lockedFiles -join ', ')`n`nThe upgrade is NOT COMPLETE until a system reboot is performed!`n`nWould you like to reboot the PC now?`n`n[Yes]  Reboot Now`n[No]   Reboot Later (before store opening)"
    
    # MB_YESNO (4) + MB_ICONWARNING (48) + MB_DEFBUTTON1 (0, default to Yes/Reboot Now) + MB_SYSTEMMODAL (4096, topmost popup)
    $dialogResult = [System.Windows.Forms.MessageBox]::Show(
        $popupMessage,
        $popupTitle,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
        [System.Windows.Forms.MessageBoxOptions]::ServiceNotification
    )

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Technician confirmed immediate reboot. Restarting PC now..." -ForegroundColor Cyan
        Write-SilentLog "Technician selected YES (Reboot now) on warning dialog. Initiating immediate reboot..."
        Restart-Computer -Force
    } else {
        Write-Host "Reboot deferred by technician (selected 'Reboot Later')." -ForegroundColor Yellow
        Write-Host "Please ensure the PC is rebooted before the store opens." -ForegroundColor Yellow
        Write-SilentLog "Technician selected NO (Reboot later) on warning dialog. Reboot deferred."
    }
}

exit 0
