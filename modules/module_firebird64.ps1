# Script Version: 1.1.0
# Module Name: Firebird 64-bit Installer
# Description: Installs the Firebird SQL Server (64-bit).
# ----------------------------------------------------------------------------------

# Output the script name and version
Write-Host "Running module_firebird64.ps1 - Version 1.1.0"

$installerUrl = "https://firebirdsql.org/download-file?file=https://github.com/FirebirdSQL/firebird/releases/download/v4.0.6/Firebird-4.0.6.3221-0-x64.exe"

# Configuration content declared as a variable for easy adjustment
$firebirdConfigContent = @"
DataTypeCompatibility = 3.0
DefaultDBCachePages = 16384
TempCacheLimit = 8000M
LockHashSlots = 65519
LockMemSize = 30M
InlineSortThreshold = 16384
MaxParallelWorkers = 15
"@

write-host "Checking if Firebird is Installed"
if (!(Test-Path "C:\Program Files\Firebird")) { 
    # Updated installer filename to reflect the 4.0.6 64-bit version
    $installerPath = "$env:TEMP\Firebird-4.0.6-x64.exe"
    write-host "Firebird is not installed" -ForegroundColor Red
    
    # --- Start Download Logic with Retry ---
    write-host "Obtaining Installer (Max 3 attempts)..."
    $maxRetries = 3
    $downloadSuccessful = $false

    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            if ($i -gt 1) {
                # Fix: Use ${i} and ${maxRetries} to prevent parsing error
                Write-Host "Download failed. Retrying in 5 seconds (Attempt ${i}/${maxRetries})..." -ForegroundColor Yellow
                Start-Sleep -Seconds 5
            }
            # Use -ErrorAction Stop to ensure failure is caught by the try block
            Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop
            $downloadSuccessful = $true
            Write-Host "Download successful." -ForegroundColor Green
            break
        } catch {
            # Fix: Use ${i} to prevent parsing error near the colon
            Write-Host "Error during download attempt ${i}: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if (-not $downloadSuccessful) {
        # Fix: Use ${maxRetries}
        Write-Host "Installation aborted due to critical download failure after ${maxRetries} attempts." -ForegroundColor Red
        exit 1
    }
    # --- End Download Logic with Retry ---
    
    write-host "Installing Firebird"
    Start-Process -FilePath $installerPath -ArgumentList "/LANG=en", "/NORESTART", "/VERYSILENT", "/MERGETASKS=UseClassicServerTask,UseServiceTask,CopyFbClientAsGds32Task" -Wait
    
    write-host "Editing firebird.conf"
    
    # Overwrite the entire firebird.conf file with the new content
    Set-Content "C:\Program Files\Firebird\Firebird_4_0\firebird.conf" -Value $firebirdConfigContent -Force
    
    write-host "Starting Firebird Service"
    Start-Service -Name "FirebirdServerDefaultInstance"
}
