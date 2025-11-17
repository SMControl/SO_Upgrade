# Script Version: 1.0.7
# Module Name: Firebird 64-bit Installer
# Description: Installs the Firebird SQL Server (64-bit).
# ----------------------------------------------------------------------------------

# Output the script name and version
Write-Host "Running module_firebird64.ps1 - Version 1.0.7"

$installerUrl = "https://firebirdsql.org/download-file?file=https://github.com/FirebirdSQL/firebird/releases/download/v4.0.6/Firebird-4.0.6.3221-0-x64-pdb.exe"

# Configuration content declared as a variable for easy adjustment
$firebirdConfigContent = @"
DataTypeCompatibility = 3.0
DefaultDBCachePages = 16384
TempCacheLimit = 8000M
LockHashSlots = 65519
LockMemSize = 30M
InlineSortThreshold = 16384
MaxParallelWorkers = 0
"@

write-host "Checking if Firebird is Installed"
if (!(Test-Path "C:\Program Files\Firebird")) { 
    # Updated installer filename to reflect the 4.0.6 64-bit version
    $installerPath = "$env:TEMP\Firebird-4.0.6-x64.exe"
    write-host "Firebird is not installed" -ForegroundColor Red
    write-host "Obtaining Installer"
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    write-host "Installing Firebird"
    Start-Process -FilePath $installerPath -ArgumentList "/LANG=en", "/NORESTART", "/VERYSILENT", "/MERGETASKS=UseClassicServerTask,UseServiceTask,CopyFbClientAsGds32Task" -Wait
    
    write-host "Editing firebird.conf"
    
    # Overwrite the entire firebird.conf file with the new content
    Set-Content "C:\Program Files\Firebird\Firebird_4_0\firebird.conf" -Value $firebirdConfigContent -Force
    
    write-host "Starting Firebird Service"
    Start-Service -Name "FirebirdServerDefaultInstance"
}
