#Write-Host "Running module_firebird_4.0.6.ps1 - Version 1.3"

$installerUrl = "https://files.stationmaster.info/Firebird-4.0.6.3221-0-Win32.exe"

$firebirdConfigContent = @"
# Essential
DataTypeCompatibility = 3.0
# Safe performance gains
DefaultDBCachePages = 12288
TempCacheLimit = 512
InlineSortThreshold = 134217728
​MaxParallelWorkers = 0 
​# Prevent Concurrency Errors 
LockHashSlots = 65519
LockMemSize = 30M
"@

if (!(Test-Path "C:\Program Files (x86)\Firebird")) { 
    $installerPath = "$env:TEMP\Firebird-4.0.6-Win32.exe"
    write-host "Firebird is not installed" -ForegroundColor Red
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    write-host "Installing Firebird 32bit V4.0.6 with DataTypeCompatibility = 3.0"
    Start-Process -FilePath $installerPath -ArgumentList "/LANG=en", "/NORESTART", "/VERYSILENT", "/MERGETASKS=UseClassicServerTask,UseServiceTask,CopyFbClientAsGds32Task" -Wait
    write-host "Editing firebird.conf"
    $configPath = "C:\Program Files (x86)\Firebird\Firebird_4_0\firebird.conf"
    Set-Content $configPath -Value $firebirdConfigContent -Force
    write-host "Starting Firebird Service"
    Start-Service -Name "FirebirdServerDefaultInstance"
}
