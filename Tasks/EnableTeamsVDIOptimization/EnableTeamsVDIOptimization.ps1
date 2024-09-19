param (
    [bool] $UninstallClassicTeams
)

function EnsureLatestVCInstalled {
    $InstalledVcList = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,*' -Name Version
    if ($InstalledVcList -contains '14.40.33810.0') {
        Write-Output "Latest Visual C++ Redistributable is already installed"
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
    if ($VcList -contains '14.40.33810.0') {
        Write-Output "Visual C++ Redistributable Installation successful $VcList"
    } else{
        Write-Error "Visual C++ Redistributable Installation failed"
    }
}

function Install-VCRedist {
    param (
        [string]$Url,
        [string]$Path
    )

    try {
        Invoke-WebRequest -Uri $Url -OutFile $Path -ErrorAction Stop
        Write-Output "Download completed successfully for $Path."
    } catch {
        Write-Output "Download failed for ${Path} $_"
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
        Write-Output "Download completed successfully."
    } catch {
        Write-Output "Download failed: $_"
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
    $TeamsPackage = Get-AppxPackage -name MsTeams -AllUsers
    if ($TeamsPackage -ne $null) {
        Write-Output "New Teams is already installed"
        return
    }

    # Download the Teams Bootstrapper
    $TeamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $TeamsBootstrapperPath = "C:\Temp\TeamsBootstrapper.exe"
    try {
        Invoke-WebRequest -Uri $TeamsBootstrapperUrl -OutFile $TeamsBootstrapperPath -ErrorAction Stop
        Write-Output "Download completed successfully."
    } catch {
        Write-Output "Download failed: $_"
        return
    }

    # Check if the file exists and is not empty
    if (Test-Path $TeamsBootstrapperPath) {
        # Install the new Teams
        Write-Output "Installing New Teams"
        Start-Process -FilePath $TeamsBootstrapperPath -ArgumentList '-p' -NoNewWindow -Wait
        Write-Output "Execution completed successfully."
    } else {
        Write-Output "Downloaded file is missing or empty."
        return
    }

    # Check if installation was successful
    if((Get-AppxPackage -name MsTeams -AllUsers) -ne $null){
        $CurrentTeams = Get-AppxPackage -name MsTeams -AllUsers
        $CurrentVersion = $CurrentTeams.Version
        Write-Output "New Teams installed successfully on version $CurrentVersion"

        # Auto-start New Teams
        Write-Output "Auto-starting New Teams"
        $InstallLocation = $CurrentTeams.InstallLocation
        & "$InstallLocation\ms-teams.exe"
    } else {
        Write-Error "New Teams installation failed"
        Write-Output "Please run new Teams intsallation pre-check script to see more details: https://aka.ms/NewTeamsReadinessCheck"
    }
    
    # Unintall the classic Teams
    if ($UninstallClassicTeams) {
        Write-Output "Uninstalling Classic Teams"
        & "$TeamsBootstrapperPath -u"
    }

    # Clean up
    Remove-Item -Path $TeamsBootstrapperPath -Force
}

# -----------------------------------------------
#                    Main 
# -----------------------------------------------
New-Item -Path 'C:\Temp' -ItemType "directory" -Force
EnsureLatestVCInstalled
EnsureLatestRDWebRTCRedirectorSerivceInstalled
EnsureNewTeamsInstalled
