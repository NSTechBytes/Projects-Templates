param (
    [string]$skinName  # Skin name must be provided when running the script
)

# Check if the skin name is provided
if (-not $skinName) {
    Write-Host "Error: Skin name must be provided. Use -skinName parameter to specify it."
    exit
}

# Function to get Rainmeter skin folder path
function Get-RainmeterSkinsFolder {
    # Default paths for Rainmeter installation
    $defaultProgramFilesPath = "$env:ProgramFiles\Rainmeter\Skins"
    $defaultProgramFilesX86Path = "$env:ProgramFiles(x86)\Rainmeter\Skins"
    $defaultDocumentsPath = "$env:UserProfile\Documents\Rainmeter\Skins"

    # Check if the path is in Program Files (64-bit)
    if (Test-Path $defaultProgramFilesPath) {
        return $defaultProgramFilesPath
    }

    # Check if the path is in Program Files (x86)
    if (Test-Path $defaultProgramFilesX86Path) {
        return $defaultProgramFilesX86Path
    }

    # Check if the path is in Documents
    if (Test-Path $defaultDocumentsPath) {
        return $defaultDocumentsPath
    }

    # Try reading the Rainmeter.ini file for custom skin folder path
    $rainmeterConfigPath = "$env:AppData\Rainmeter\Rainmeter.ini"
    if (Test-Path $rainmeterConfigPath) {
        $config = Get-Content $rainmeterConfigPath
        foreach ($line in $config) {
            if ($line -like "SkinPath=*") {
                $customSkinPath = $line -replace "SkinPath=", ""
                if (Test-Path $customSkinPath) {
                    return $customSkinPath
                }
            }
        }
    }

    # Return a message if no path is found
    return "Rainmeter Skins folder not found!"
}

# Function to check if Rainmeter is running and stop it
function Stop-Rainmeter {
    $rainmeterProcess = Get-Process -Name "Rainmeter" -ErrorAction SilentlyContinue

    if ($rainmeterProcess) {
        Write-Host "Rainmeter is running. Stopping it..."
        Stop-Process -Name "Rainmeter" -Force
        Write-Host "Rainmeter has been stopped."
    } else {
        Write-Host "Rainmeter is not running."
    }
}

# Stop Rainmeter if it's running
Stop-Rainmeter

# GitHub raw URL to the Version.nek file for the given skin name
$versionUrl = "https://raw.githubusercontent.com/NSTechBytes/$skinName/main/%40Resources/Version.nek"

# Destination folder to save the downloaded skin and extraction folder
$destinationFolder = "C:\nstechbytes"
$extractionFolder = "$destinationFolder\$skinName"

# Retry logic and error handling for Version.nek file download
$retryCount = 3
$timeout = 30
$tempVersionFile = "$env:TEMP\Version.nek"
$success = $false

for ($i = 0; $i -lt $retryCount; $i++) {
    try {
        Write-Host "Attempting to download the version file... (Attempt $($i+1) of $retryCount)"
        Invoke-WebRequest -Uri $versionUrl -OutFile $tempVersionFile -TimeoutSec $timeout -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
        $success = $true
        break
    } catch {
        Write-Host "Download failed on attempt $($i+1). Retrying..."
    }
}

if (-not $success) {
    Write-Host "Failed to download the version file after $retryCount attempts."
    exit
}

# Read the Version.nek file and extract the version number
$versionData = Get-Content $tempVersionFile
$versionLine = $versionData | Where-Object { $_ -match "Version=" }
$version = $versionLine -replace "Version=", ""

# GitHub release URL for the skin download based on the extracted version
$url = "https://github.com/NSTechBytes/$skinName/releases/download/$skinName/${skinName}_$version.rmskin"

# Check if the destination folder exists, if not, create it
if (-Not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
}

# Destination file path
$destination = "$destinationFolder\${skinName}_$version.rmskin"

# Retry logic and error handling for the skin download
$success = $false

for ($i = 0; $i -lt $retryCount; $i++) {
    try {
        Write-Host "Attempting to download the skin... (Attempt $($i+1) of $retryCount)"
        Invoke-WebRequest -Uri $url -OutFile $destination -TimeoutSec $timeout -Headers @{ 'User-Agent' = 'Mozilla/5.0' }
        $success = $true
        break
    } catch {
        Write-Host "Download failed on attempt $($i+1). Retrying..."
    }
}

if (-not $success) {
    Write-Host "Failed to download the skin after $retryCount attempts."
    exit
} else {
    Write-Host "Downloaded $skinName version $version to $destination"
}

# Create extraction folder if it doesn't exist
if (-Not (Test-Path -Path $extractionFolder)) {
    New-Item -ItemType Directory -Path $extractionFolder
}

# Extract the .rmskin file (as it is essentially a ZIP archive)
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
[System.IO.Compression.ZipFile]::ExtractToDirectory($destination, $extractionFolder)

Write-Host "Extracted $skinName to $extractionFolder"

# Clean up the temporary version file
Remove-Item $tempVersionFile

# Find the skin path using the function
$skinsDirectory = Get-RainmeterSkinsFolder

# Remove the existing skin folder if it exists
if (Test-Path -Path $skinsDirectory) {
    # Specify the exact path for the skin to remove
    $skinToRemove = Join-Path -Path $skinsDirectory -ChildPath $skinName

    if (Test-Path -Path $skinToRemove) {
        Remove-Item -Path $skinToRemove -Recurse -Force
        Write-Host "Removed the existing skin folder: $skinToRemove"
    } else {
        Write-Host "No existing skin folder found to remove: $skinToRemove"
    }
} else {
    Write-Host "No Rainmeter skins folder found."
}

# Copy the extracted skin folder to the Rainmeter skins path
$extractedSkinFolder = Join-Path -Path $extractionFolder -ChildPath "Skins\$skinName"

if (Test-Path -Path $extractedSkinFolder) {
    $destinationSkinPath = Join-Path -Path $skinsDirectory -ChildPath $skinName
    Copy-Item -Path $extractedSkinFolder -Destination $destinationSkinPath -Recurse -Force
    Write-Host "Copied the extracted folder to: $destinationSkinPath"
} else {
    Write-Host "Extracted skin folder does not exist: $extractedSkinFolder"
}

# Determine the OS architecture (32-bit or 64-bit)
$is64Bit = [Environment]::Is64BitOperatingSystem

# Set the appropriate plugins folder based on OS architecture
if ($is64Bit) {
    $pluginsFolder = Join-Path -Path $extractionFolder -ChildPath "Plugins\64bit"
} else {
    $pluginsFolder = Join-Path -Path $extractionFolder -ChildPath "Plugins\32bit"
}

# Define the Rainmeter Plugins path using environment variable for user profile
$rainmeterPluginsPath = Join-Path -Path $env:UserProfile -ChildPath "AppData\Roaming\Rainmeter\Plugins"

# Check if the plugins folder exists and copy DLL files
if (Test-Path -Path $pluginsFolder) {
    if (Test-Path -Path $rainmeterPluginsPath) {
        # Copy the DLL files to the Rainmeter Plugins folder
        Get-ChildItem -Path $pluginsFolder -Filter "*.dll" | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination $rainmeterPluginsPath -Force
            Write-Host "Copied DLL: $($_.Name) to $rainmeterPluginsPath"
        }
    } else {
        Write-Host "Rainmeter Plugins folder not found!"
    }
} else {
    Write-Host "No DLL files found in: $pluginsFolder"
}

# Start Rainmeter application
$rainmeterPath = Join-Path -Path $env:ProgramFiles -ChildPath "Rainmeter\Rainmeter.exe"
if (-Not (Test-Path -Path $rainmeterPath)) {
    $rainmeterPath = Join-Path -Path $env:ProgramFiles(x86) -ChildPath "Rainmeter\Rainmeter.exe"
}

if (Test-Path -Path $rainmeterPath) {
    Start-Process -FilePath $rainmeterPath
} else {
    Write-Host "Rainmeter executable not found!"
}

# Remove the nstechbytes folder
if (Test-Path -Path $destinationFolder) {
    Remove-Item -Path $destinationFolder -Recurse -Force
    Write-Host "Cleaned up the nstechbytes temporary folder."
}

Write-Host "Update process completed!"
