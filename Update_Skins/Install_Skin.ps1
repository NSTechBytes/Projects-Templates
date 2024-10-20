param (
    [string]$skinName 
)

#=================================================================================================================================#
#                                                      Download Logic of Rainmeter                                                #
#=================================================================================================================================# 

$repoApiUrl = "https://api.github.com/repos/Rainmeter/Rainmeter/releases/latest"
$installerPath = "$env:TEMP\Rainmeter-Installer.exe"
$rainmeterExePath = "$env:ProgramFiles\Rainmeter\Rainmeter.exe"

function Is-RainmeterInstalled {
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Rainmeter",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Rainmeter"
    )

    foreach ($path in $paths) {
        try {
            $rainmeter = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            if ($rainmeter) {
                return $true
            }
        } catch {
            # Handle error silently
        }
    }
    return $false
}

function Get-LatestRainmeterInstallerUrl {
    try {
        $releaseData = Invoke-RestMethod -Uri $repoApiUrl -UseBasicParsing
        foreach ($asset in $releaseData.assets) {
            if ($asset.name -like "*.exe") {
                return $asset.browser_download_url
            }
        }
    } catch {
        Write-Host "Error: Unable to fetch the latest Rainmeter release from GitHub."
        exit 1
    }
}

if (Is-RainmeterInstalled) {
    Write-Host "Rainmeter is already installed."
} else {
    Write-Host "Rainmeter is not installed. Downloading the latest version..."

    $rainmeterUrl = Get-LatestRainmeterInstallerUrl

    if ($rainmeterUrl) {
        Write-Host "Downloading Rainmeter from $rainmeterUrl..."

        # Download the latest Rainmeter installer
        $installerPath = "$env:TEMP\RainmeterInstaller.exe"  # Define the path for the installer
        Invoke-WebRequest -Uri $rainmeterUrl -OutFile $installerPath

        Write-Host "Download complete. Prompting for admin rights..."

        # Run the installer with admin privileges
        $result = Start-Process -FilePath $installerPath -ArgumentList "/S" -Verb RunAs -Wait -PassThru

        if ($result.ExitCode -eq 0) {
            Write-Host "Rainmeter installation completed successfully."
        } else {
            Write-Host "Installation was canceled or failed. Exiting script."
            Remove-Item -Path $installerPath -Force  # Clean up the installer if canceled or failed
            exit
        }

        # Clean up the installer file after installation
        Remove-Item -Path $installerPath -Force
    } else {
        Write-Host "Failed to retrieve the latest Rainmeter installer URL."
    }
}

#=================================================================================================================================#
#                                                      Run The Rainmeter after Installation                                       #
#=================================================================================================================================#

# Start Rainmeter application
$rainmeterPath = Join-Path -Path $env:ProgramFiles -ChildPath "Rainmeter\Rainmeter.exe"
if (-Not (Test-Path -Path $rainmeterPath)) {
    $rainmeterPath = Join-Path -Path $env:ProgramFiles(x86) -ChildPath "Rainmeter\Rainmeter.exe"
}

if (Test-Path -Path $rainmeterPath) {
    Start-Process -FilePath $rainmeterPath 
    Start-Sleep -Seconds 1

    $activationCommands = @(
        "!DeactivateConfig illustro\Clock",
        "!DeactivateConfig illustro\Disk",
        "!DeactivateConfig illustro\System",
        "!DeactivateConfig illustro\Welcome"
    )

    foreach ($command in $activationCommands) {
        Start-Process -FilePath $rainmeterPath -ArgumentList $command
        Write-Host "Deactivated skin using command: $command"
    }
}

#=================================================================================================================================#
#                                                      Installation of Skin                                                       #
#=================================================================================================================================#

if (-not $skinName) {
    Write-Host "Error: Skin name must be provided. Use -skinName parameter to specify it."
    exit
}

function Get-RainmeterSkinsFolder {
    # Default paths for Rainmeter installation
    $defaultProgramFilesPath = "$env:ProgramFiles\Rainmeter\Skins"
    $defaultProgramFilesX86Path = "$env:ProgramFiles(x86)\Rainmeter\Skins"
    $defaultDocumentsPath = "$env:UserProfile\Documents\Rainmeter\Skins"

    if (Test-Path $defaultProgramFilesPath) {
        return $defaultProgramFilesPath
    }

    if (Test-Path $defaultProgramFilesX86Path) {
        return $defaultProgramFilesX86Path
    }

    if (Test-Path $defaultDocumentsPath) {
        return $defaultDocumentsPath
    }

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

    return "Rainmeter Skins folder not found!"
}

$versionUrl = "https://raw.githubusercontent.com/NSTechBytes/$skinName/main/%40Resources/Version.nek"
$destinationFolder = "C:\nstechbytes"
$extractionFolder = "$destinationFolder\$skinName"
$tempVersionFile = "$env:TEMP\Version.nek"
Invoke-WebRequest -Uri $versionUrl -OutFile $tempVersionFile

# Read the Version.nek file and extract the version number
$versionData = Get-Content $tempVersionFile
$versionLine = $versionData | Where-Object { $_ -match "Version=" }
$version = $versionLine -replace "Version=", ""

# GitHub release URL for the skin download based on the extracted version
$url = "https://github.com/NSTechBytes/$skinName/releases/download/v$version/${skinName}_$version.rmskin"

# Check if the destination folder exists, if not, create it
if (-Not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
}

# Destination file path
$destination = "$destinationFolder\${skinName}_$version.rmskin"

# Download the skin using the retrieved version
Invoke-WebRequest -Uri $url -OutFile $destination

Write-Host "Downloaded $skinName version $version to $destination"

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

function Take-OwnershipAndRemoveItem {
    param (
        [string]$path
    )

    if (Test-Path -Path $path) {
        try {
            # Take ownership of the file/directory
            Write-Host "Taking ownership of $path"
            Start-Process "cmd.exe" -ArgumentList "/c takeown /f `"$path`" /r /d y && icacls `"$path`" /grant administrators:F /t" -Verb RunAs -Wait -NoNewWindow

            # Now attempt to remove it
            Write-Host "Removing $path"
            Remove-Item -Path $path -Recurse -Force
            Write-Host "Successfully removed $path"
        } catch {
            Write-Host "Error: Failed to remove $path. Error details: $_"
        }
    } else {
        Write-Host "Path $path does not exist."
    }
}

# Example usage:
Take-OwnershipAndRemoveItem -path $skinToRemove

# Remove the existing skin folder if it exists
if (Test-Path -Path $skinsDirectory) {
    # Specify the exact path for the skin to remove
    $skinToRemove = Join-Path -Path $skinsDirectory -ChildPath $skinName

    if (Test-Path -Path $skinToRemove) {
        Take-OwnershipAndRemoveItem -path $skinToRemove
    } else {
        Write-Host "No existing skin folder found to remove: $skinToRemove"
    }
} else {
    Write-Host "No Rainmeter skins folder found."
}
#=================================================================================================================================#
#                                                      Stop Rainmeter                                                             #                          
#=================================================================================================================================#

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

#=================================================================================================================================#
#                                                      Copy Skin To Skins Folder                                                  #                          
#=================================================================================================================================#

Start-Sleep -Seconds 1

# Copy the extracted skin folder to the Rainmeter skins path
$extractedSkinFolder = Join-Path -Path $extractionFolder -ChildPath "Skins\$skinName"

if (Test-Path -Path $extractedSkinFolder) {
    $destinationSkinPath = Join-Path -Path $skinsDirectory -ChildPath $skinName
    Copy-Item -Path $extractedSkinFolder -Destination $destinationSkinPath -Recurse -Force
    Write-Host "Copied the extracted folder to: $destinationSkinPath"
} else {
    Write-Host "Extracted skin folder does not exist: $extractedSkinFolder"
}

#=================================================================================================================================#
#                                                      Plugins Logic                                                              #                          
#=================================================================================================================================#

$is64Bit = [Environment]::Is64BitOperatingSystem

if ($is64Bit) {
    $pluginsFolder = Join-Path -Path $extractionFolder -ChildPath "Plugins\64bit"
} else {
    $pluginsFolder = Join-Path -Path $extractionFolder -ChildPath "Plugins\32bit"
}

$rainmeterPluginsPath = Join-Path -Path $env:UserProfile -ChildPath "AppData\Roaming\Rainmeter\Plugins"

if (-not (Test-Path -Path $rainmeterPluginsPath)) {
    # Create the plugins folder if it doesn't exist
    New-Item -Path $rainmeterPluginsPath -ItemType Directory
    Write-Host "Created Plugins folder at: $rainmeterPluginsPath"
}

Start-Sleep -Seconds 1

if (Test-Path -Path $rainmeterPluginsPath) {
    Get-ChildItem -Path $pluginsFolder -Filter "*.dll" | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $rainmeterPluginsPath -Force
        Write-Host "Copied DLL: $($_.Name) to $rainmeterPluginsPath"
    }
} else {
    Write-Host "Rainmeter Plugins folder not found!"
}

#=================================================================================================================================#
#                                                      Start Rainmeter and Load Skin                                              #                          
#=================================================================================================================================#

$rainmeterPath = Join-Path -Path $env:ProgramFiles -ChildPath "Rainmeter\Rainmeter.exe"
if (-Not (Test-Path -Path $rainmeterPath)) {
    $rainmeterPath = Join-Path -Path $env:ProgramFiles(x86) -ChildPath "Rainmeter\Rainmeter.exe"
}

if (Test-Path -Path $rainmeterPath) {
    Start-Process -FilePath $rainmeterPath
    Start-Sleep -Seconds 1
    $activationCommand = "!ActivateConfig $skinName\Startup Main.ini"
    Start-Process -FilePath $rainmeterPath -ArgumentList $activationCommand
    Write-Host "Activated skin using command: $activationCommand"
} else {
    Write-Host "Rainmeter executable not found!"
}

# Remove the nstechbytes folder
if (Test-Path -Path $destinationFolder) {
    Remove-Item -Path $destinationFolder -Recurse -Force
    Write-Host "Removed the nstechbytes folder: $destinationFolder"
} else {
    Write-Host "nstechbytes folder not found!"
}

Write-Host "$skinName is Installed Successfully."
