# module_firebird_v5_32bit.ps1 - Script v1.01 - PartVersion-1.01
# Write-Host "Running module_firebird_v5_32bit.ps1"

# Download URL for Firebird 5 32-bit setup binary from GitHub raw storage
$installerUrl = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/8e101759465cbfc70a5e1f2232a28c5eb00d62f8/bin/Firebird-5.0.4.1812-0-windows-x86.exe"

# Configuration content required by Smart Office compatibility mode
$firebirdConfigContent = @"
# Essential
DataTypeCompatibility = 3.0
#Configuration for Firebird 5 (vanilla) Classic (64 bit)
DefaultDBCachePages =2048
LockHashSlots = 65519
LockMemSize  = 30M
ParallelWorkers = 1
MaxParallelWorkers = 64
MaxStatementCacheSize=32M
UseFileSystemCache = true
TempCacheLimit = 256M
InlineSortThreshold = 16384
"@

# Check if Firebird directory exists under 32-bit Program Files
if (!(Test-Path "C:\Program Files (x86)\Firebird")) { 
    $installerPath = "$env:TEMP\Firebird-5.0.4.1812-0-windows-x86.exe"
    write-host "Firebird is not installed" -ForegroundColor Red
    
    try {
        # Fetch the installer binary to local TEMP directory
        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop
        write-host "Installing Firebird 32bit V5.0.4 with DataTypeCompatibility = 3.0"
        
        # Run silent installation with classic server tasks and GDS32 client copy
        Start-Process -FilePath $installerPath -ArgumentList "/LANG=en", "/NORESTART", "/VERYSILENT", "/MERGETASKS=UseClassicServerTask,UseServiceTask,CopyFbClientAsGds32Task" -Wait
        
        # Configure firebird.conf in the Firebird 5.0 installation folder
        write-host "Editing firebird.conf"
        $configPath = "C:\Program Files (x86)\Firebird\Firebird_5_0\firebird.conf"
        
        # If Firebird_5_0 path is not yet present, fallback dynamically to any subfolder containing firebird.conf
        if (!(Test-Path $configPath)) {
            $detectedConfig = Get-ChildItem -Path "C:\Program Files (x86)\Firebird" -Filter "firebird.conf" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($detectedConfig) {
                $configPath = $detectedConfig.FullName
            }
        }
        
        if (Test-Path (Split-Path -Path $configPath -Parent)) {
            Set-Content $configPath -Value $firebirdConfigContent -Force
        }
        
        # Start default Firebird server service
        write-host "Starting Firebird Service"
        Start-Service -Name "FirebirdServerDefaultInstance" -ErrorAction SilentlyContinue
    }
    catch {
        write-host "Error during Firebird 5 installation: $_" -ForegroundColor Red
    }
}
