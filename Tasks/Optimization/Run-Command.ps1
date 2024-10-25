<#
.SYNOPSIS
This script will enable Teams VDI Optimization for customer to improve their Teams experience on Dev Box.

.PARAMETER UninstallClassicTeams
Uninstall the classic Teams if the parameter is set to true, false by default.

.DESCRIPTION
Seeing more in [doc](https://microsoft.sharepoint.com/:w:/t/Fidalgo/Ea5PK_IvfqtNv8gWtedntfIBOBZvj-DVkng8NWu1pSGJsQ?e=R5kyIe)

Copyright (c) Microsoft. All rights reserved.
#>

param (
    [bool] $UninstallClassicTeams
)

function EnsureLatestVCInstalled {
    $MinimumVersion = "14.40.33810.0"
    # Check if the latest Visual C++ Redistributable is already installed
    $InstalledVcList = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,*' -Name Version
    if (Compare-VersionsToTarget -VersionList $InstalledVcList -TargetVersion $MinimumVersion) {
        Write-Output "Latest Visual C++ Redistributable is already installed - $InstalledVcList"
        return
    }
    
    # Define the URLs for the latest Visual C++ Redistributable
    $vcRedistUrlX64 = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    $vcRedistUrlX86 = "https://aka.ms/vs/17/release/vc_redist.x86.exe"

    # Define the paths where the installers will be saved
    $vcRedistPathX64 = "C:\Temp\vc_redist.x64.exe"
    $vcRedistPathX86 = "C:\Temp\vc_redist.x86.exe"
    # Install x64 version
    Install-VCRedist -Url $vcRedistUrlX64 -Path $vcRedistPathX64
    # Install x86 version
    Install-VCRedist -Url $vcRedistUrlX86 -Path $vcRedistPathX86

    # Check if installation was successful
    $VcList = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,*' -Name Version
    if (Compare-VersionsToTarget -VersionList $VcList -TargetVersion $MinimumVersion) {
        Write-Output "Visual C++ Redistributable Installation successful - $VcList"
    } else{
        Write-Error "Visual C++ Redistributable Installation failed"
    }
}

function Compare-VersionsToTarget {
    param (
        [string[]]$VersionList,
        [string]$TargetVersion
    )

    $trueResults = 0

    try {
        $target = [version]$TargetVersion

        foreach ($version in $VersionList) {
            $v = [version]$version

            if ($v -ge $target) {
                $trueResults++
            }
        }
    } catch {
        Write-Error "Error: Invalid version format. Please ensure the versions are in the correct format. $VersionList"
        return $false
    }

    # Return true if there are at least 2 versions greater or equal than the target version
    return $trueResults -ge 2
}

function Install-VCRedist {
    param (
        [string]$Url,
        [string]$Path
    )

    try {
        Invoke-WebRequest -Uri $Url -OutFile $Path -ErrorAction Stop
        Write-Output "Visual C++ Redistributable Download completed successfully for $Path."
    } catch {
        Write-Output "Visual C++ Redistributable Download failed for ${Path} $_"
        exit 1
    }

    # Check if the file exists and is not empty
    if (Test-Path $Path) {
        # Run the installer silently
        Write-Output "Installing latest Visual C++ Redistributable for $Path"
        Start-Process -FilePath $Path -ArgumentList '/install', '/quiet', '/norestart' -NoNewWindow -Wait
        Write-Output "Installation completed successfully for $Path."
    } else {
        Write-Output "Downloaded file is missing or empty for $Path."
    }

    # Clean up
    Remove-Item -Path $Path -Force
}

function EnsureLatestRDWebRTCRedirectorSerivceInstalled {
    # Download the latest MSI
    $Url = "https://aka.ms/msrdcwebrtcsvc/msi"
    $Package = "C:\Temp\Latest_WebRTCRedirectorService.MSI"
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Package -ErrorAction Stop
        Write-Output "Remote Desktop Web RTC Redirector Service Download completed successfully."
    } catch {
        Write-Output "Remote Desktop Web RTC Redirector Service Download failed: $_"
        return
    }

    # Check if the file exists and is not empty
    if (Test-Path $Package) {
        # Kick off installation
        Write-Output "Installing $Package"
        Start-Process -FilePath $Package -ArgumentList '/quiet' -Wait
    } else {
        Write-Output "Downloaded file is missing or empty."
        return
    }
    
    # Check if installation was successful
    $Product = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq "Remote Desktop WebRTC Redirector Service" }
    if ($Product -ne $null) {
        $CurrentVersion = $Product.Version
        Write-Output "Remote Desktop WebRTC Redirector Service Installation successful on version $CurrentVersion"
    } else {
        Write-Error "Remote Desktop WebRTC Redirector Service Installation failed"
    }
    
    # Clean up
    Remove-Item -Path $Package -Force
}

function EnsureNewTeamsInstalled {
    # Check if New Teams is already installed
    # There might be multiple versions of Teams installed, we will always use the latest one
    $TeamsPackage = (Get-AppxPackage -name MsTeams -AllUsers)[-1]
    if ($TeamsPackage -ne $null) {
        $Version = $TeamsPackage.Version
        Write-Output "New Teams is already installed with version - $Version"
    } else {
        $TeamsPackage = Install-NewTeams
    }

    # Restart Teams to make sure the latest version has been installed
    Write-Output "Restarting New Teams"
    Restart-NewTeams -TeamsPackage $TeamsPackage
    
    # Unintall the classic Teams
    if ($UninstallClassicTeams) {
        Write-Output "Uninstalling Classic Teams"
        & "$TeamsBootstrapperPath -u"
    }
}

function Install-NewTeams(){
    # Download the Teams Bootstrapper
    $TeamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $TeamsBootstrapperPath = "C:\Temp\TeamsBootstrapper.exe"
    try {
        Invoke-WebRequest -Uri $TeamsBootstrapperUrl -OutFile $TeamsBootstrapperPath -ErrorAction Stop
        Write-Output "TeamsBootstrapper Download completed successfully."
    } catch {
        Write-Output "TeamsBootstrapper Download TeamsBootstrapper failed: $_"
        return $null
    }

    # Check if the file exists and is not empty
    if (Test-Path $TeamsBootstrapperPath) {
        # Install the new Teams
        Write-Output "Installing New Teams"
        Start-Process -FilePath $TeamsBootstrapperPath -ArgumentList '-p' -NoNewWindow -Wait
        Write-Output "Installation completed successfully."
    } else {
        Write-Output "TeamsBootstrapper Downloaded file is missing or empty."
        return $null
    }

    # Check if installation was successful
    $CurrentTeams = (Get-AppxPackage -name MsTeams -AllUsers)[-1]
    if($CurrentTeams -ne $null){
        $CurrentVersion = $CurrentTeams.Version
        Write-Output "New Teams installed successfully on version $CurrentVersion"
        return $CurrentTeams
    } else {
        Write-Error "New Teams installation failed"
        Write-Output "Please run new Teams intsallation pre-check script to see more details: https://aka.ms/NewTeamsReadinessCheck"
        return $null
    }

    # Clean up
    Remove-Item -Path $TeamsBootstrapperPath -Force
}

function Restart-NewTeams(){
    param (
        [array]$TeamsPackage
    )
    if ($TeamsPackage -ne $null) {
        # Restart Teams to get the latest version
        $InstallLocation = $TeamsPackage.InstallLocation
        & "$InstallLocation\ms-teams.exe"
        Write-Output "New Teams process has been started."

        # Wait for a specified amount of time (e.g., 10 seconds)
        Start-Sleep -Seconds 30

        # Get the process ID of the new Teams application
        $TeamsProcess = Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue

        # Check if the process is running
        if ($TeamsProcess) {
            # Kill the process
            Stop-Process -Id $TeamsProcess.Id -Force
            Write-Output "New Teams process has been terminated."
        } else {
            Write-Output "New Teams process is not running."
        }

        $UpdatedVersion = (Get-AppxPackage -name MsTeams -AllUsers)[-1].Version
        Write-Output "The version of Teams is $UpdatedVersion"
    }
}

# -----------------------------------------------
#                    Main 
# -----------------------------------------------
New-Item -Path 'C:\Temp' -ItemType "directory" -Force
EnsureLatestVCInstalled
EnsureLatestRDWebRTCRedirectorSerivceInstalled
EnsureNewTeamsInstalled