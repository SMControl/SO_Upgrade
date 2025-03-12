# module_firebird.ps1
$installerUrl = "https://raw.githubusercontent.com/SMControl/SO_Upgrade/main/bin/Firebird-4.0.1.exe"
write-host "Checking if Firebird is Installed"
if (!(Test-Path "C:\Program Files (x86)\Firebird")) { 
    $installerPath = "$env:TEMP\Firebird-4.0.1.exe"
    write-host "Firebird is not installed" -ForegroundColor Red
    write-host "Obtaining Installer"
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    write-host "Installing Firebird"
    Start-Process -FilePath $installerPath -ArgumentList "/LANG=en", "/NORESTART", "/VERYSILENT", "/MERGETASKS=UseClassicServerTask,UseServiceTask,CopyFbClientAsGds32Task" -Wait
    write-host "Editing firebird.conf"
    (Get-Content "C:\Program Files (x86)\Firebird\Firebird_4_0\firebird.conf") -replace '#DataTypeCompatibility.*', 'DataTypeCompatibility = 3.0' | Set-Content "C:\Program Files (x86)\Firebird\Firebird_4_0\firebird.conf"
    write-host "Starting Firebird Service"
    Start-Service -Name "FirebirdServerDefaultInstance"
}

