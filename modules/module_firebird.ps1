#Write-Host "Running module_firebird_sm_default.ps1"

$installerUrl = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/main/bin/Firebird-4.0.1.exe"

$firebirdConfigContent = @"
# Essential
DataTypeCompatibility = 3.0
"@

if (!(Test-Path "C:\Program Files (x86)\Firebird")) { 
    $installerPath = "$env:TEMP\Firebird-4.0.1.exe"
    write-host "Firebird is not installed" -ForegroundColor Red
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    write-host "Installing Firebird 32bit V4.0.1 with DataTypeCompatibility = 3.0"
    Start-Process -FilePath $installerPath -ArgumentList "/LANG=en", "/NORESTART", "/VERYSILENT", "/MERGETASKS=UseClassicServerTask,UseServiceTask,CopyFbClientAsGds32Task" -Wait
    write-host "Editing firebird.conf"
    $configPath = "C:\Program Files (x86)\Firebird\Firebird_4_0\firebird.conf"
    Set-Content $configPath -Value $firebirdConfigContent -Force
    write-host "Starting Firebird Service"
    Start-Service -Name "FirebirdServerDefaultInstance"
}
