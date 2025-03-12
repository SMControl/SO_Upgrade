Write-Host "SO_UC.ps1 - Version 1.08"
# -----
# - Cleaned up messaging
# - version bump
# - rmeoved timestamp renaming stuff

# Part 1 - Check if scheduled task exists and create if it doesn't
# PartVersion 1.00
# -----
Write-Host "Checking Task..."
$taskExists = Get-ScheduledTask -TaskName "SO InstallerUpdates" -ErrorAction SilentlyContinue

if (-not $taskExists) {
    $action = New-ScheduledTaskAction -Execute "C:\winsm\SO_UC.exe"
    $randomHour = Get-Random -Minimum 0 -Maximum 5
    $randomMinute = Get-Random -Minimum 0 -Maximum 59
    $trigger = New-ScheduledTaskTrigger -Daily -At "${randomHour}:${randomMinute}"
    $settings = New-ScheduledTaskSettingsSet -Hidden:$true
    Register-ScheduledTask -TaskName "SO InstallerUpdates" -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest
}

# Part 2 - Retrieve .exe links from the webpage
# PartVersion 1.00
# -----
Write-Host "Checking Versions..."
$exeLinks = (Invoke-WebRequest -Uri "https://www.stationmaster.com/downloads/").Links | Where-Object { $_.href -match "\.exe$" } | ForEach-Object { $_.href }

# Part 3 - Filter for the highest two versions of Setup.exe
# PartVersion 1.01
# -----
$setupLinks = $exeLinks | Where-Object { $_ -match "^https://www\.stationmaster\.com/Download/Setup\d+\.exe$" }
$sortedLinks = $setupLinks | Sort-Object { [regex]::Match($_, "Setup(\d+)\.exe").Groups[1].Value -as [int] } -Descending

$highestTwoLinks = $sortedLinks | Select-Object -First 2

# Part 4 - Download the highest two versions if not already present
# PartVersion 1.03
# -----
$downloadDirectory = "C:\winsm\SmartOffice_Installer"
if (-not (Test-Path $downloadDirectory)) {
    New-Item -ItemType Directory -Path $downloadDirectory
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
    } catch {
        Write-Host "Error fetching response from server: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    $existingFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe"
    $fileExists = $existingFiles | Where-Object { $_.Length -eq $contentLength }

    if (-not $fileExists) {
        Write-Host "Downloading new version: $originalFilename" -ForegroundColor Green
        Invoke-WebRequest -Uri $downloadLink -OutFile $destinationPath
    } else {
        Write-Host "File $originalFilename already exists and is up to date." -ForegroundColor Yellow
    }
}


# Part 5 - Delete older downloads, keeping the latest two
# PartVersion 1.01
# -----
Write-Host "Cleaning Up..."
$downloadedFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe" | Sort-Object LastWriteTime -Descending
if ($downloadedFiles.Count -gt 2) {
    $filesToDelete = $downloadedFiles | Select-Object -Skip 2
    foreach ($file in $filesToDelete) {
        Write-Host "Deleting old file: $($file.Name)" -ForegroundColor Red
        Remove-Item -Path $file.FullName -Force
    }
}
Write-Host "Finished"
