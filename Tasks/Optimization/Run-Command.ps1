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
    $Url = "https://aka.ms/msrdcwebrtcsvc/msi"
    $Package = "C:\Temp\Latest_WebRTCRedirectorService.MSI"
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Package -ErrorAction Stop
        Write-Output "Remote Desktop Web RTC Redirector Service Download completed successfully."
    } catch {
        Write-Output "Remote Desktop Web RTC Redirector Service Download failed: $_"
        return
    }

    if (!(Test-Path $Package) -or ((Get-Item $Package).Length -lt 0)) {
        Write-Output "Downloaded file is missing or empty."
        return
    }

    $MaxRetryCount = 3
    for($i = 0; $i -lt $MaxRetryCount; $i++){
        Write-Output "Trying to install Remote Desktop WebRTC Redirector Service for the $i time."
        Install-RDWebRTCRedirectorServiceWithLocalPacakge -Package $Package
        # Check if installation was successful
        $Product = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq "Remote Desktop WebRTC Redirector Service" }
        if ($Product -ne $null) {
            $CurrentVersion = $Product.Version
            Write-Output "Remote Desktop WebRTC Redirector Service Installation successful on version $CurrentVersion"
            break
        } else {
            Write-Output "Remote Desktop WebRTC Redirector Service Installation failed"
        }
    }
    
    # Clean up
    Remove-Item -Path $Package -Force
}

function Install-RDWebRTCRedirectorServiceWithLocalPacakge{
    param (
        [string]$Package
    )
    # Check if the file exists and is not empty
    if ((Test-Path $Package) -and ((Get-Item $Package).Length -gt 0)) {
        # Kick off installation
        Write-Output "Installing $Package"
        try {
            Start-Process -FilePath $Package -ArgumentList '/quiet' -Wait -ErrorAction Stop
            Write-Output "Installation process started successfully."
        } catch {
            Write-Error "Installation process failed: $_"
            return
        }
    } else {
        Write-Output "Downloaded file is missing or empty."
        return
    }
}

function EnsureNewTeamsInstalled {
    # Check if New Teams is already installed
    # There might be multiple versions of Teams installed, we will always use the latest one
    $TeamsPackages = Get-AppxPackage -name MsTeams -AllUsers
    if ($TeamsPackages -ne $null) {
        $TeamsPackage = $TeamsPackages[-1]
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
        $UpdatePath = "$InstallLocation\ms-teamsupdate.exe"
        if(Test-Path $UpdatePath){
            Write-Output "Trying to update new Teams to the latest version, it may not working though."
            & "$UpdatePath"
        }
        
        Write-Output "Trying to start new Teams."
        & "$InstallLocation\ms-teams.exe"

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