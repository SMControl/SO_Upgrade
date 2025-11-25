# ==================================================================================================
# Script: module_soget_NEW.ps1
# Version: 2.00
# Description: Downloads the latest two Smart Office setup files from stationmaster.com
# ==================================================================================================

Write-Host "Checking for latest Smart Office setup files..." -ForegroundColor Cyan

# ==================================================================================================
# PART 1: Get setup file links from website
# ==================================================================================================

try {
    $webpageUrl = "https://www.stationmaster.com/downloads/"
    $allLinks = (Invoke-WebRequest -Uri $webpageUrl -UseBasicParsing).Links
    
    # Filter for Setup*.exe files
    $setupLinks = $allLinks | 
    Where-Object { $_.href -match "^https://www\.stationmaster\.com/Download/Setup\d+\.exe$" } | 
    ForEach-Object { $_.href }
    
    if ($setupLinks.Count -eq 0) {
        Write-Host "Error: No setup files found on website." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Found $($setupLinks.Count) setup file(s) on website." -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving links from website: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ==================================================================================================
# PART 2: Sort and select the highest two versions
# ==================================================================================================

# Extract version numbers and sort descending
$setupLinksWithVersions = $setupLinks | ForEach-Object {
    $versionMatch = [regex]::Match($_, "Setup(\d+)\.exe")
    [PSCustomObject]@{
        Link     = $_
        Version  = $versionMatch.Groups[1].Value -as [int]
        Filename = $_.Split('/')[-1]
    }
}

$highestTwo = $setupLinksWithVersions | 
Sort-Object -Property Version -Descending | 
Select-Object -First 2

Write-Host "Latest versions: $($highestTwo[0].Filename), $($highestTwo[1].Filename)" -ForegroundColor Cyan

# ==================================================================================================
# PART 3: Download files if needed
# ==================================================================================================

$downloadDirectory = "C:\winsm\SmartOffice_Installer"

# Ensure download directory exists
if (-not (Test-Path $downloadDirectory)) {
    Write-Host "Creating download directory: $downloadDirectory" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $downloadDirectory | Out-Null
}

foreach ($setup in $highestTwo) {
    $destinationPath = Join-Path -Path $downloadDirectory -ChildPath $setup.Filename
    
    # Get file size from server
    try {
        $request = [System.Net.HttpWebRequest]::Create($setup.Link)
        $request.Method = "HEAD"
        $request.UserAgent = "Mozilla/5.0"
        $response = $request.GetResponse()
        $serverFileSize = $response.ContentLength
        $response.Close()
    }
    catch {
        Write-Host "Error checking server for $($setup.Filename): $($_.Exception.Message)" -ForegroundColor Red
        continue
    }
    
    # Check if this specific file exists locally with the same size
    $localFile = Get-ChildItem -Path $downloadDirectory -Filter $setup.Filename -ErrorAction SilentlyContinue
    
    if ($localFile -and $localFile.Length -eq $serverFileSize) {
        Write-Host "$($setup.Filename) is already up to date ($(([math]::Round($serverFileSize / 1MB, 1))) MB)" -ForegroundColor Green
    }
    else {
        # File doesn't exist or has different size - download it
        $sizeMB = [math]::Round($serverFileSize / 1MB, 1)
        Write-Host "Downloading $($setup.Filename) ($sizeMB MB)..." -ForegroundColor Yellow
        
        try {
            Invoke-WebRequest -Uri $setup.Link -OutFile $destinationPath -UseBasicParsing
            Write-Host "$($setup.Filename) downloaded successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error downloading $($setup.Filename): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ==================================================================================================
# PART 4: Clean up old files (keep only the 2 newest)
# ==================================================================================================

$allDownloadedFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe" | 
Sort-Object LastWriteTime -Descending

if ($allDownloadedFiles.Count -gt 2) {
    $filesToDelete = $allDownloadedFiles | Select-Object -Skip 2
    
    foreach ($file in $filesToDelete) {
        Write-Host "Deleting old file: $($file.Name)" -ForegroundColor Red
        Remove-Item -Path $file.FullName -Force
    }
    
    Write-Host "Cleanup complete. Kept 2 newest files." -ForegroundColor Green
}

Write-Host "Setup file check complete." -ForegroundColor Cyan
