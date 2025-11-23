# ==================================================================================================
# Script: SOUpgradeAssistant_GUI.ps1
# Version: 3.168
# Description: GUI version of the Smart Office Upgrade Assistant using Windows Forms
# ==================================================================================================
# Recent Changes:
# - Version 3.168: PROFESSIONAL UI POLISH
#   - Added white header panel to seamlessly integrate the logo
#   - Updated title styling (Dark Blue on White)
#   - Removed redundant step label and centered status text
#   - Refined vertical spacing for a cleaner look
# - Version 3.167: RESTORED 3.164 LAYOUT
#   - Restored file integrity after manual edit error
#   - Reverted to 3.164 layout (Large logo, 18pt progress text)
# - Version 3.166: REVERT LAYOUT
#   - Reverted UI layout to match version 3.164 (fixed "warped" appearance)
# ==================================================================================================

# Requires -RunAsAdministrator
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==================================================================================================
# GLOBAL CONFIGURATION
# ==================================================================================================

$Global:Config = @{
    ScriptVersion = "3.168"
    WorkingDir    = "C:\winsm"
    LogDir        = "C:\winsm\SmartOffice_Installer\soua_logs"
    Services      = @{
        LiveSales = "srvSOLiveSales"
        Firebird  = "FirebirdServerDefaultInstance"
    }
    Processes     = @{
        SmartOffice = @("Sm32Main", "Sm32")
        Firebird    = "firebird"
        PDTWiFi     = "PDTWiFi"
        PDTWiFi64   = "PDTWiFi64"
    }
    Paths         = @{
        StationMaster = "C:\Program Files (x86)\StationMaster"
        Firebird      = "C:\Program Files (x86)\Firebird"
        SetupDir      = "C:\winsm\SmartOffice_Installer"
    }
    URLs          = @{
        ModuleSOGets   = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_soget.ps1"
        ModuleFirebird = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/modules/module_firebird.ps1"
    }
    Timeouts      = @{
        ProcessCheckInterval = 2
        ServiceRetryInterval = 5
        ServiceMaxRetries    = 3
    }
}

# Global variables
$Global:TotalSteps = 14
$Global:StartTime = Get-Date
$Global:UpgradeInProgress = $false
$Global:WasRunning = $false
$Global:MonitorJob = $null
$Global:SelectedExe = $null
$Global:PDTWiFiStates = @{}
$Global:UserCancelled = $false

# ==================================================================================================
# LOGGING FUNCTIONS
# ==================================================================================================

function Write-GuiLog {
    param(
        [string]$Message,
        [string]$Color = "Black"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    
    # Add to log textbox
    $logTextBox.SelectionStart = $logTextBox.TextLength
    $logTextBox.SelectionLength = 0
    
    $colorMap = @{
        "Red"    = [System.Drawing.Color]::FromArgb(239, 68, 68)    # Bright red for errors
        "Yellow" = [System.Drawing.Color]::FromArgb(251, 191, 36)   # Yellow for warnings
        "White"  = [System.Drawing.Color]::White                     # White for everything else
    }
    
    # Default to white if color not specified or not in map
    $displayColor = if ($colorMap.ContainsKey($Color)) { $colorMap[$Color] } else { $colorMap["White"] }
    $logTextBox.SelectionColor = $displayColor
    $logTextBox.AppendText("$logEntry`r`n")
    $logTextBox.SelectionColor = $logTextBox.ForeColor
    $logTextBox.ScrollToCaret()
    
    # Also write to file
    $logFile = Join-Path $Global:Config.LogDir ("soua_log_" + (Get-Date -Format "yyyy-MM-dd_HHmm") + ".log")
    if (-not (Test-Path (Split-Path $logFile))) {
        New-Item -ItemType Directory -Path (Split-Path $logFile) -Force | Out-Null
    }
    Add-Content -Path $logFile -Value $logEntry
}

function Update-Progress {
    param(
        [int]$Step,
        [string]$Status
    )
    
    $percentage = [math]::Round(($Step / $Global:TotalSteps) * 100)
    $progressBar.Value = $percentage
    $statusLabel.Text = "Step $Step/$($Global:TotalSteps): $Status"
    
    [System.Windows.Forms.Application]::DoEvents()
}

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}



function Show-ActionButtons {
    param(
        [string]$Message,
        [hashtable]$Buttons  # @{ "ButtonText" = { ScriptBlock } }
    )
    
    # Clear existing buttons
    $actionPanel.Controls.Clear()
    
    # Add message label
    $messageLabel = New-Object System.Windows.Forms.Label
    $messageLabel.Location = New-Object System.Drawing.Point(10, 10)
    $messageLabel.Size = New-Object System.Drawing.Size(740, 40)
    $messageLabel.Text = $Message
    $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $messageLabel.ForeColor = [System.Drawing.Color]::White
    $actionPanel.Controls.Add($messageLabel)
    
    # Add buttons
    $buttonX = 10
    foreach ($buttonText in $Buttons.Keys) {
        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point($buttonX, 55)
        $button.Size = New-Object System.Drawing.Size(120, 35)
        $button.Text = $buttonText
        $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $button.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)  # #007BFF
        $button.ForeColor = [System.Drawing.Color]::White
        $button.FlatStyle = "Flat"
        $button.Add_Click($Buttons[$buttonText])
        $actionPanel.Controls.Add($button)
        $buttonX += 130
    }
}

function Hide-ActionButtons {
    $actionPanel.Controls.Clear()
}

# ==================================================================================================
# CLEANUP FUNCTION
# ==================================================================================================

function Invoke-Cleanup {
    Write-GuiLog "Performing cleanup..." "Yellow"
    
    # Stop monitoring job if running
    if ($Global:MonitorJob) {
        Stop-Job -Job $Global:MonitorJob -ErrorAction SilentlyContinue
        Remove-Job -Job $Global:MonitorJob -ErrorAction SilentlyContinue
        $Global:MonitorJob = $null
    }
    
    Write-GuiLog "Cleanup complete." "Green"
}

# ==================================================================================================
# UPGRADE PROCESS STEPS
# ==================================================================================================

function Step1-CheckAdmin {
    Update-Progress 1 "Checking administrator rights..."
    Write-GuiLog "[Step 1/14] Checking Administrator Rights" "Cyan"
    
    if (-not (Test-Admin)) {
        Write-GuiLog "ERROR: This script must be run as Administrator!" "Red"
        [System.Windows.Forms.MessageBox]::Show("This script requires Administrator privileges. Please run as Administrator.", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    
    Write-GuiLog "Administrator rights confirmed." "Green"
    return $true
}
function Step2-DownloadSetup {
    Update-Progress 2 "Checking for setup files..."
    Write-GuiLog "[Step 2/14] Checking for Setup Files" "Cyan"
    
    # Define the URL for the SO Get module
    $sogetScriptURL = $Global:Config.URLs.ModuleSOGets
    
    try {
        Write-GuiLog "Downloading and checking setup files. This may take a moment. Please wait..." "Yellow"
        Write-GuiLog "Invoking module_soget.ps1..." "Yellow"
        
        # Download the module content
        $moduleContent = Invoke-RestMethod -Uri $sogetScriptURL -ErrorAction Stop
        
        # Execute and capture output
        Invoke-Expression $moduleContent
        
        Write-GuiLog "module_soget.ps1 executed successfully." "Green"
        
        # Verify that we now have setup files
        if (Test-Path $Global:Config.Paths.SetupDir) {
            $setupFiles = Get-ChildItem -Path $Global:Config.Paths.SetupDir -Filter "*.exe"
            if ($setupFiles.Count -gt 0) {
                Write-GuiLog "Found $($setupFiles.Count) setup file(s)." "Green"
                return $true
            }
            else {
                Write-GuiLog "No setup files found after running module_soget." "Red"
                return $false
            }
        }
        else {
            Write-GuiLog "Setup directory not found after running module_soget." "Red"
            return $false
        }
    }
    catch {
        Write-GuiLog "Error executing module_soget.ps1: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Step3-CheckFirebird {
    Update-Progress 3 "Checking Firebird installation..."
    Write-GuiLog "[Step 3/14] Checking Firebird" "Cyan"
    
    if (Test-Path $Global:Config.Paths.Firebird) {
        Write-GuiLog "Firebird is already installed." "Green"
    }
    else {
        Write-GuiLog "Firebird is not installed. Installing..." "Yellow"
        Write-GuiLog "This process runs in the background and may take a few minutes..." "Yellow"
        
        # Define the URL for the Firebird installation script
        $firebirdInstallerURL = $Global:Config.URLs.ModuleFirebird
        
        try {
            # Run installation in background job to prevent UI freeze
            $job = Start-Job -ScriptBlock {
                param($url)
                try {
                    # Download and execute the module
                    $moduleContent = Invoke-RestMethod -Uri $url -ErrorAction Stop
                    Invoke-Expression $moduleContent
                }
                catch {
                    throw $_
                }
            } -ArgumentList $firebirdInstallerURL
            
            # Poll job status and keep UI responsive
            while ($job.State -eq 'Running') {
                Start-Sleep -Milliseconds 500
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $results = Receive-Job -Job $job
            if ($job.State -eq 'Failed') {
                throw "Job failed"
            }
            Remove-Job -Job $job
            
            Write-GuiLog "Firebird installation script executed." "Green"
            
            # Verify installation
            if (Test-Path $Global:Config.Paths.Firebird) {
                Write-GuiLog "Firebird installation verified." "Green"
            }
            else {
                Write-GuiLog "Warning: Firebird directory not found after installation script." "Yellow"
            }
        }
        catch {
            Write-GuiLog "Error installing Firebird: $($_.Exception.Message)" "Red"
            return $false
        }
    }
    
    return $true
}

function Step4-MonitorSMUpdates {
    Update-Progress 4 "Monitoring SMUpdates..."
    Write-GuiLog "[Step 4/14] Monitoring SMUpdates" "Cyan"
    
    # Start background job to monitor SMUpdates
    $Global:MonitorJob = Start-Job -ScriptBlock {
        param($modulePath)
        Set-Location $modulePath
        & ".\module_smupdates.ps1"
    } -ArgumentList $Global:Config.Paths.StationMaster
    
    Write-GuiLog "SMUpdates monitoring started." "Green"
    return $true
}

function Step5-ManageLiveSales {
    Update-Progress 5 "Managing SO Live Sales service..."
    Write-GuiLog "[Step 5/14] Managing SO Live Sales Service" "Cyan"
    
    $serviceName = $Global:Config.Services.LiveSales
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    
    if ($service) {
        if ($service.Status -eq "Running") {
            $Global:WasRunning = $true
            Write-GuiLog "Service '$serviceName' is running. Stopping..." "Yellow"
            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
            Write-GuiLog "Service stopped." "Green"
        }
        else {
            $Global:WasRunning = $false
            Write-GuiLog "Service '$serviceName' is not running." "Gray"
        }
    }
    else {
        Write-GuiLog "Service '$serviceName' not found." "Yellow"
    }
    
    return $true
}

function Step6-ManagePDTWiFi {
    Update-Progress 6 "Managing PDTWiFi processes..."
    Write-GuiLog "[Step 6/14] Managing PDTWiFi Processes" "Cyan"
    
    $Global:PDTWiFiStates = @{}
    
    # PDTWiFi
    $pdtWiFi = $Global:Config.Processes.PDTWiFi
    $proc = Get-Process -Name $pdtWiFi -ErrorAction SilentlyContinue
    if ($proc) {
        $Global:PDTWiFiStates[$pdtWiFi] = "Running"
        Stop-Process -Name $pdtWiFi -Force -ErrorAction SilentlyContinue
        Write-GuiLog "$pdtWiFi stopped." "Green"
    }
    else {
        $Global:PDTWiFiStates[$pdtWiFi] = "Not running"
        Write-GuiLog "$pdtWiFi is not running." "Yellow"
    }
    
    # PDTWiFi64
    $pdtWiFi64 = $Global:Config.Processes.PDTWiFi64
    $proc = Get-Process -Name $pdtWiFi64 -ErrorAction SilentlyContinue
    if ($proc) {
        $Global:PDTWiFiStates[$pdtWiFi64] = "Running"
        Stop-Process -Name $pdtWiFi64 -Force -ErrorAction SilentlyContinue
        Write-GuiLog "$pdtWiFi64 stopped." "Green"
    }
    else {
        $Global:PDTWiFiStates[$pdtWiFi64] = "Not running"
        Write-GuiLog "$pdtWiFi64 is not running." "Yellow"
    }
    
    return $true
}

function Step7-WaitForClose {
    Update-Progress 7 "Waiting for Smart Office to close..."
    Write-GuiLog "[Step 7/14] Waiting for Smart Office to Close" "Cyan"
    
    # Check if Smart Office is running
    $soRunning = $false
    foreach ($process in $Global:Config.Processes.SmartOffice) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            $soRunning = $true
            break
        }
    }
    
    if ($soRunning) {
        Write-GuiLog "Smart Office is running. Please close it to continue." "Red"
        
        # Show action buttons instead of popup
        Show-ActionButtons -Message "Smart Office is currently running. Please close it to continue." -Buttons @{
            "Continue" = {
                Hide-ActionButtons
            }
            "Cancel"   = {
                Hide-ActionButtons
                Write-GuiLog "User cancelled." "Yellow"
                $Global:UserCancelled = $true
            }
        }
        
        # Wait for user to click a button
        while ($actionPanel.Controls.Count -gt 0 -and !$Global:UserCancelled) {
            Start-Sleep -Milliseconds 100
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        if ($Global:UserCancelled) {
            $Global:UserCancelled = $false
            return $false
        }
        
        # Wait for processes to close
        foreach ($process in $Global:Config.Processes.SmartOffice) {
            while (Get-Process -Name $process -ErrorAction SilentlyContinue) {
                Start-Sleep -Seconds $Global:Config.Timeouts.ProcessCheckInterval
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }
    
    Write-GuiLog "Smart Office is closed." "Green"
    
    # Wait for single Firebird instance
    $fbProcesses = Get-Process -Name $Global:Config.Processes.Firebird -ErrorAction SilentlyContinue
    while ($fbProcesses.Count -gt 1) {
        Write-GuiLog "Warning: Multiple Firebird instances running ($($fbProcesses.Count)). Waiting..." "Yellow"
        Start-Sleep -Seconds $Global:Config.Timeouts.ProcessCheckInterval
        [System.Windows.Forms.Application]::DoEvents()
        $fbProcesses = Get-Process -Name $Global:Config.Processes.Firebird -ErrorAction SilentlyContinue
    }
    
    if ($fbProcesses.Count -eq 1) {
        Write-GuiLog "Only one Firebird instance running." "Green"
    }
    
    return $true
}

function Step8-LaunchSetup {
    Update-Progress 8 "Launching Smart Office setup..."
    Write-GuiLog "[Step 8/14] Launching Setup" "Cyan"
    
    # Validate setup directory exists
    if (-not (Test-Path $Global:Config.Paths.SetupDir -PathType Container)) {
        Write-GuiLog "ERROR: Setup directory does not exist: $($Global:Config.Paths.SetupDir)" "Red"
        [System.Windows.Forms.MessageBox]::Show(
            "Setup directory not found: $($Global:Config.Paths.SetupDir)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    
    # Get setup files
    $setupExes = Get-ChildItem -Path $Global:Config.Paths.SetupDir -Filter "*.exe" -ErrorAction SilentlyContinue
    
    if ($setupExes.Count -eq 0) {
        Write-GuiLog "ERROR: No setup executable found!" "Red"
        [System.Windows.Forms.MessageBox]::Show("No setup executable found in $($Global:Config.Paths.SetupDir)", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    elseif ($setupExes.Count -eq 1) {
        $Global:SelectedExe = $setupExes[0]
        Write-GuiLog "Found setup: $($Global:SelectedExe.Name)" "Green"
    }
    else {
        # Multiple setups - show selection in Action Panel with colorful buttons
        $setupExes = $setupExes | Sort-Object {
            [regex]::Match($_.Name, "Setup(\d+)\.exe").Groups[1].Value -as [int]
        }
        
        Write-GuiLog "Multiple setup files found. Please select one:" "Yellow"
        
        # Define colors for buttons
        $buttonColors = @(
            @{ BG = [System.Drawing.Color]::FromArgb(34, 197, 94); FG = [System.Drawing.Color]::White },   # Green
            @{ BG = [System.Drawing.Color]::FromArgb(249, 115, 22); FG = [System.Drawing.Color]::White },  # Orange
            @{ BG = [System.Drawing.Color]::FromArgb(168, 85, 247); FG = [System.Drawing.Color]::White },  # Purple
            @{ BG = [System.Drawing.Color]::FromArgb(20, 184, 166); FG = [System.Drawing.Color]::White },  # Teal
            @{ BG = [System.Drawing.Color]::FromArgb(239, 68, 68); FG = [System.Drawing.Color]::White },   # Red
            @{ BG = [System.Drawing.Color]::FromArgb(59, 130, 246); FG = [System.Drawing.Color]::White }   # Blue
        )
        
        # Clear action panel and add custom colored buttons
        $actionPanel.Controls.Clear()
        
        # Add message label
        $messageLabel = New-Object System.Windows.Forms.Label
        $messageLabel.Location = New-Object System.Drawing.Point(10, 10)
        $messageLabel.Size = New-Object System.Drawing.Size(740, 30)
        $messageLabel.Text = "Select Smart Office version to install:"
        $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $messageLabel.ForeColor = [System.Drawing.Color]::White
        $actionPanel.Controls.Add($messageLabel)
        
        # Add buttons
        $buttonX = 10
        $buttonWidth = 230
        for ($i = 0; $i -lt $setupExes.Count; $i++) {
            $exe = $setupExes[$i]
            $colorIndex = $i % $buttonColors.Count
            
            $button = New-Object System.Windows.Forms.Button
            $button.Location = New-Object System.Drawing.Point($buttonX, 50)
            $button.Size = New-Object System.Drawing.Size($buttonWidth, 45)
            $button.Text = $exe.Name
            $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $button.BackColor = $buttonColors[$colorIndex].BG
            $button.ForeColor = $buttonColors[$colorIndex].FG
            $button.FlatStyle = "Flat"
            $button.Tag = $exe
            $button.Add_Click({
                    $Global:SelectedExe = $this.Tag
                    Write-GuiLog "Selected: $($this.Tag.Name)" "Green"
                    Hide-ActionButtons
                })
            $actionPanel.Controls.Add($button)
            $buttonX += $buttonWidth + 10
        }
        
        # Wait for user to select a setup
        while ($actionPanel.Controls.Count -gt 0 -and $null -eq $Global:SelectedExe -and -not $form.IsDisposed -and -not $Global:UserCancelled) {
            Start-Sleep -Milliseconds 100
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        if ($Global:UserCancelled -or $form.IsDisposed) {
            Write-GuiLog "Setup selection cancelled." "Yellow"
            return $false
        }
        
        if ($null -eq $Global:SelectedExe) {
            Write-GuiLog "Setup selection cancelled." "Yellow"
            return $false
        }
    }
    
    # Launch setup
    try {
        Write-GuiLog "Starting setup: $($Global:SelectedExe.Name)..." "Cyan"
        $setupProcess = Start-Process -FilePath $Global:SelectedExe.FullName -Wait -PassThru
        
        if ($setupProcess.ExitCode -eq 0) {
            Write-GuiLog "Setup completed successfully." "Green"
        }
        else {
            Write-GuiLog "Setup finished with exit code: $($setupProcess.ExitCode)" "Yellow"
        }
    }
    catch {
        Write-GuiLog "Error launching setup: $($_.Exception.Message)" "Red"
        return $false
    }
    
    return $true
}

function Step9-PostUpgrade {
    Update-Progress 9 "Post-upgrade tasks..."
    Write-GuiLog "[Step 9/14] Post Upgrade" "Cyan"
    
    # Stop monitoring job
    if ($Global:MonitorJob) {
        Stop-Job -Job $Global:MonitorJob -ErrorAction SilentlyContinue
        Remove-Job -Job $Global:MonitorJob -ErrorAction SilentlyContinue
        $Global:MonitorJob = $null
        Write-GuiLog "SMUpdates monitoring stopped." "Green"
    }
    
    # Confirm upgrade complete
    Show-ActionButtons -Message "Please open and close Smart Office before continuing.`nA restart of the Firebird Service may be required." -Buttons @{
        "Continue" = {
            Hide-ActionButtons
        }
    }
    
    # Wait for user to click Continue
    while ($actionPanel.Controls.Count -gt 0) {
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Wait for Smart Office to close
    foreach ($process in $Global:Config.Processes.SmartOffice) {
        while (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            Start-Sleep -Seconds $Global:Config.Timeouts.ProcessCheckInterval
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    
    Write-GuiLog "Smart Office confirmed closed." "Green"
    return $true
}

function Step10-SetPermissionsSM {
    Update-Progress 10 "Setting StationMaster permissions..."
    Write-GuiLog "[Step 10/14] Setting StationMaster Permissions" "Cyan"
    Write-GuiLog "This may take 1-30+ minutes. Please wait..." "Yellow"
    
    try {
        # Run icacls in background job to prevent UI freeze
        $job = Start-Job -ScriptBlock {
            param($path)
            & icacls $path /grant "*S-1-1-0:(OI)(CI)F" /T /C 2>&1
        } -ArgumentList $Global:Config.Paths.StationMaster
        
        # Poll job status and keep UI responsive
        while ($job.State -eq 'Running') {
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        Receive-Job -Job $job | Out-Null
        Remove-Job -Job $job
        
        Write-GuiLog "Permissions set successfully." "Green"
    }
    catch {
        Write-GuiLog "Error setting permissions: $($_.Exception.Message)" "Red"
    }
    
    return $true
}

function Step11-SetPermissionsFB {
    Update-Progress 11 "Setting Firebird permissions..."
    Write-GuiLog "[Step 11/14] Setting Firebird Permissions" "Cyan"
    
    try {
        # Run icacls in background job to prevent UI freeze
        $job = Start-Job -ScriptBlock {
            param($path)
            & icacls $path /grant "*S-1-1-0:(OI)(CI)F" /T /C 2>&1
        } -ArgumentList $Global:Config.Paths.Firebird
        
        # Poll job status and keep UI responsive
        while ($job.State -eq 'Running') {
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        Receive-Job -Job $job | Out-Null
        Remove-Job -Job $job
        
        Write-GuiLog "Permissions set successfully." "Green"
    }
    catch {
        Write-GuiLog "Error setting permissions: $($_.Exception.Message)" "Red"
    }
    
    return $true
}

function Step12-RevertLiveSales {
    Update-Progress 12 "Reverting SO Live Sales service..."
    Write-GuiLog "[Step 12/14] Reverting SO Live Sales Service" "Cyan"
    
    if ($Global:WasRunning) {
        $serviceName = $Global:Config.Services.LiveSales
        Write-GuiLog "Service was running before. Restarting..." "Yellow"
        
        try {
            $retryCount = 0
            $maxRetries = $Global:Config.Timeouts.ServiceMaxRetries
            $retryIntervalSeconds = $Global:Config.Timeouts.ServiceRetryInterval
            
            while ($retryCount -lt $maxRetries) {
                Write-GuiLog "Attempting to start service '$serviceName' (Attempt $($retryCount + 1) of $maxRetries)..." "Yellow"
                Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                
                if ((Get-Service -Name $serviceName).Status -eq "Running") {
                    Write-GuiLog "Service '$serviceName' is now running." "Green"
                    break
                }
                else {
                    Write-GuiLog "Service '$serviceName' is not running. Waiting $retryIntervalSeconds seconds before retrying..." "Yellow"
                    Start-Sleep -Seconds $retryIntervalSeconds
                    [System.Windows.Forms.Application]::DoEvents()
                    $retryCount++
                }
            }
            
            # If still not running, ask user to start manually
            if ((Get-Service -Name $serviceName).Status -ne "Running") {
                Write-GuiLog "Failed to automatically start service '$serviceName' after $maxRetries attempts." "Red"
                
                Show-ActionButtons -Message "The service '$serviceName' could not be started automatically. Please start it manually." -Buttons @{
                    "Continue" = {
                        Hide-ActionButtons
                    }
                    "Cancel"   = {
                        Hide-ActionButtons
                        $Global:UserCancelled = $true
                    }
                }
                
                # Wait for user to click a button
                while ($actionPanel.Controls.Count -gt 0 -and !$Global:UserCancelled) {
                    Start-Sleep -Milliseconds 100
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                if ($Global:UserCancelled) {
                    $Global:UserCancelled = $false
                    Write-GuiLog "User cancelled manual service start." "Yellow"
                    return $false
                }
                
                # Wait for service to be running
                while ((Get-Service -Name $serviceName).Status -ne "Running") {
                    Write-GuiLog "Waiting for '$serviceName' service to be running..." "Yellow"
                    Start-Sleep -Seconds $Global:Config.Timeouts.ProcessCheckInterval
                    [System.Windows.Forms.Application]::DoEvents()
                }
                Write-GuiLog "Service '$serviceName' is now running. Continuing..." "Green"
            }
        }
        catch {
            Write-GuiLog "Error starting service: $($_.Exception.Message)" "Red"
        }
    }
    else {
        Write-GuiLog "Service was not running before. No action needed." "Yellow"
    }
    
    return $true
}

function Step13-RevertPDTWiFi {
    Update-Progress 13 "Reverting PDTWiFi processes..."
    Write-GuiLog "[Step 13/14] Reverting PDTWiFi Processes" "Cyan"
    
    # PDTWiFi
    $pdtWiFi = $Global:Config.Processes.PDTWiFi
    if ($Global:PDTWiFiStates[$pdtWiFi] -eq "Running") {
        try {
            Start-Process (Join-Path $Global:Config.Paths.StationMaster "PDTWiFi.exe") -ErrorAction Stop
            Write-GuiLog "$pdtWiFi started." "Green"
        }
        catch {
            Write-GuiLog "Error starting $pdtWiFi : $($_.Exception.Message)" "Red"
        }
    }
    else {
        Write-GuiLog "$pdtWiFi was not running. No action taken." "Yellow"
    }
    
    # PDTWiFi64
    $pdtWiFi64 = $Global:Config.Processes.PDTWiFi64
    if ($Global:PDTWiFiStates[$pdtWiFi64] -eq "Running") {
        try {
            Start-Process (Join-Path $Global:Config.Paths.StationMaster "PDTWiFi64.exe") -ErrorAction Stop
            Write-GuiLog "$pdtWiFi64 started." "Green"
        }
        catch {
            Write-GuiLog "Error starting $pdtWiFi64 : $($_.Exception.Message)" "Red"
        }
    }
    else {
        Write-GuiLog "$pdtWiFi64 was not running. No action taken." "Yellow"
    }
    
    return $true
}

function Step14-Finish {
    Update-Progress 14 "Finalizing..."
    Write-GuiLog "[Step 14/14] Finishing Up" "Cyan"
    
    # Calculate execution time
    $endTime = Get-Date
    $executionTime = $endTime - $Global:StartTime
    $totalMinutes = [math]::Floor($executionTime.TotalMinutes)
    $totalSeconds = $executionTime.Seconds
    
    Write-GuiLog "" "Green"
    Write-GuiLog "========================================" "Green"
    Write-GuiLog "UPGRADE COMPLETE!" "Green"
    Write-GuiLog "========================================" "Green"
    Write-GuiLog "Completed in $($totalMinutes)m $($totalSeconds)s." "Green"
    Write-GuiLog "" "Green"
    
    # Show completion options
    # Show completion options with custom buttons
    $actionPanel.Controls.Clear()
    
    # Message
    $messageLabel = New-Object System.Windows.Forms.Label
    $messageLabel.Location = New-Object System.Drawing.Point(10, 10)
    $messageLabel.Size = New-Object System.Drawing.Size(740, 30)
    $messageLabel.Text = "Upgrade completed successfully! Choose an action:"
    $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $messageLabel.ForeColor = [System.Drawing.Color]::White
    $actionPanel.Controls.Add($messageLabel)
    
    # Button 1: Start Smart Office (Blue)
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Location = New-Object System.Drawing.Point(10, 50)
    $btnStart.Size = New-Object System.Drawing.Size(160, 40)
    $btnStart.Text = "Start Smart Office"
    $btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnStart.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255) # Blue
    $btnStart.ForeColor = [System.Drawing.Color]::White
    $btnStart.FlatStyle = "Flat"
    $btnStart.Add_Click({
            Hide-ActionButtons
            try {
                Start-Process (Join-Path $Global:Config.Paths.StationMaster "Sm32.exe") -ErrorAction Stop
                Write-GuiLog "Smart Office started." "Green"
            }
            catch {
                Write-GuiLog "Error starting Smart Office: $($_.Exception.Message)" "Red"
            }
        })
    $actionPanel.Controls.Add($btnStart)
    
    # Button 2: Reboot (Red)
    $btnReboot = New-Object System.Windows.Forms.Button
    $btnReboot.Location = New-Object System.Drawing.Point(180, 50)
    $btnReboot.Size = New-Object System.Drawing.Size(120, 40)
    $btnReboot.Text = "Reboot PC"
    $btnReboot.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnReboot.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69) # Red
    $btnReboot.ForeColor = [System.Drawing.Color]::White
    $btnReboot.FlatStyle = "Flat"
    $btnReboot.Add_Click({
            $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to reboot the computer now?", "Confirm Reboot", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Hide-ActionButtons
                Write-GuiLog "Rebooting system..." "Red"
                Restart-Computer -Force
            }
        })
    $actionPanel.Controls.Add($btnReboot)
    
    # Button 3: Close (Gray)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(310, 50)
    $btnClose.Size = New-Object System.Drawing.Size(100, 40)
    $btnClose.Text = "Close"
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125) # Gray
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = "Flat"
    $btnClose.Add_Click({
            Hide-ActionButtons
            $form.Close()
        })
    $actionPanel.Controls.Add($btnClose)
}

# ==================================================================================================
# MAIN UPGRADE PROCESS
# ==================================================================================================

function Start-UpgradeProcess {
    $Global:UpgradeInProgress = $true
    $Global:StartTime = Get-Date
    
    $steps = @(
        { Step1-CheckAdmin },
        { Step2-DownloadSetup },
        { Step3-CheckFirebird },
        { Step4-MonitorSMUpdates },
        { Step5-ManageLiveSales },
        { Step6-ManagePDTWiFi },
        { Step7-WaitForClose },
        { Step8-LaunchSetup },
        { Step9-PostUpgrade },
        { Step10-SetPermissionsSM },
        { Step11-SetPermissionsFB },
        { Step12-RevertLiveSales },
        { Step13-RevertPDTWiFi },
        { Step14-Finish }
    )
    
    foreach ($step in $steps) {
        $result = & $step
        if (-not $result) {
            Write-GuiLog "Upgrade process stopped." "Red"
            $Global:UpgradeInProgress = $false
            Invoke-Cleanup
            return
        }
    }
    
    $Global:UpgradeInProgress = $false
}

# ==================================================================================================
# GUI CREATION
# ==================================================================================================

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Smart Office Upgrade"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "Manual"
$form.Location = New-Object System.Drawing.Point(0, 0)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)  # #003366 - StationMaster Primary

# Add form closing event for cleanup
$form.Add_FormClosing({
        param($sender, $e)
    
        # If upgrade is in progress, confirm cancellation
        if ($Global:UpgradeInProgress) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Upgrade is in progress. Are you sure you want to cancel?",
                "Confirm Cancel",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        
            if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                $e.Cancel = $true
                return
            }
            
            # User confirmed cancel
            $Global:UserCancelled = $true
        }
    
        # Cleanup
        Invoke-Cleanup
    })



# Header Panel (White background for Logo and Title)
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = "Top"
$headerPanel.Height = 100
$headerPanel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($headerPanel)

# Logo (Inside Header)
$logoBox = New-Object System.Windows.Forms.PictureBox
$logoBox.Location = New-Object System.Drawing.Point(20, 12)
$logoBox.Size = New-Object System.Drawing.Size(180, 75)
$logoBox.SizeMode = "Zoom"
$logoBox.ImageLocation = "https://stationmaster.info/logo-station-master.png"
$headerPanel.Controls.Add($logoBox)

# Title Label (Inside Header)
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(220, 30)
$titleLabel.Size = New-Object System.Drawing.Size(550, 40)
$titleLabel.Text = "Smart Office Upgrade"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102) # StationMaster Blue
$headerPanel.Controls.Add($titleLabel)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 120)
$statusLabel.Size = New-Object System.Drawing.Size(760, 30)
$statusLabel.Text = "Ready to start upgrade process"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$statusLabel.TextAlign = "MiddleCenter"
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(191, 219, 254)  # Light blue
$form.Controls.Add($statusLabel)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 160)
$progressBar.Size = New-Object System.Drawing.Size(760, 25)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

# Log TextBox
$logTextBox = New-Object System.Windows.Forms.RichTextBox
$logTextBox.Location = New-Object System.Drawing.Point(20, 200)
$logTextBox.Size = New-Object System.Drawing.Size(760, 220)
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logTextBox.ReadOnly = $true
$logTextBox.BackColor = [System.Drawing.Color]::FromArgb(31, 41, 55)  # Dark gray
$logTextBox.ForeColor = [System.Drawing.Color]::FromArgb(229, 231, 235)  # Light gray text
$logTextBox.BorderStyle = "FixedSingle"
$form.Controls.Add($logTextBox)

# Action Panel
$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Location = New-Object System.Drawing.Point(20, 430)
$actionPanel.Size = New-Object System.Drawing.Size(760, 100)
$actionPanel.BorderStyle = "FixedSingle"
$actionPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 86, 179)  # #0056b3 - StationMaster Accent
$form.Controls.Add($actionPanel)

# Close Button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(690, 540)
$closeButton.Size = New-Object System.Drawing.Size(90, 30)
$closeButton.Text = "Close"
$closeButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$closeButton.BackColor = [System.Drawing.Color]::FromArgb(107, 114, 128)  # Gray
$closeButton.ForeColor = [System.Drawing.Color]::White
$closeButton.FlatStyle = "Flat"
$closeButton.Add_Click({
        $form.Close()
    })
$form.Controls.Add($closeButton)

# Display initial version in log
Write-GuiLog "souaGUI.ps1 - Version $($Global:Config.ScriptVersion)" "Cyan"
Write-GuiLog "Starting upgrade process..." "Gray"
Write-GuiLog ""

# Show form
$form.Add_Shown({
        # Auto-start upgrade when form is shown
        Start-UpgradeProcess
    })

[void]$form.ShowDialog()

# Cleanup on exit
$form.Dispose()
