param (
    [bool] $UninstallClassicTeams,
)

function EnsureLatestVCInstalled {
    # Get the installed Visual C++ Redistributable version list
    $InstalledVcList = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,*' -Name Version

    # Install the VcRedist Module if not already installed
    if ($(Get-Module VcRedist).Length -gt 0) {
        Write-Output "VcRedist Module already installed"
    } else {
        Write-Output "Installing VcRedist Module"
        Install-Module -Name VcRedist -Force
        Import-Module VcRedist
    }
    
    # Install the latest Visual C++ Redistributable if it was not Installed 
    $TempPath = "C:\Temp\VcRedist"
    $VcList = Get-VcList | Where-Object { $_.Version -notin $InstalledVcList } | Get-VcRedist -Path $TempPath
    if ($VcList.Length -gt 0) {
        Write-Output "Installing latest Visual C++ Redistributable"
        $VcList | Install-VcRedist -Path $TempPath

        # Check if installation was successful
        $VcVersions = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,*' -Name Version
        if ((Get-VcList | Where-Object { $_.Version -notin $VcVersions }).Length -eq 0) {
            Write-Output "Visual C++ Redistributable Installation successful"
            Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\VC,*' -Name Version
        } else{
            Write-Error "Visual C++ Redistributable Installation failed"
        }
    } else {
        Write-Output "Latest Visual C++ Redistributable already installed"
    }

    # Clean up
    Remove-Item -Path $TempPath -Force
}

function EnsureLatestRDWebRTCRedirectorSerivceInstalled {
    # Download the latest MSI
    $Url = "https://aka.ms/msrdcwebrtcsvc/msi"
    $Package = "C:\Temp\Latest_WebRTCRedirectorService.MSI"
    Invoke-WebRequest -Uri $Url -OutFile $Package

    # Kick off installation
    Write-Output "Installing $Package"
    Start-Process -FilePath $Package -ArgumentList '/quiet' -Wait
    
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
    Invoke-WebRequest -Uri $TeamsBootstrapperUrl -OutFile $TeamsBootstrapperPath

    # Install the new Teams
    Write-Output "Installing New Teams"
    & "$TeamsBootstrapperPath -p"

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
