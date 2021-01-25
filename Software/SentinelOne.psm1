#!ps
<#  
.SYNOPSIS  
    Downloads and installs SentinelOne.
.DESCRIPTION  
    
.NOTES  
    File Name  : Install-SentinelOne.ps1
    Author     : jack@netlinkinc.net
    Requires   : 

.LINK 
#>

function Install-SentinelOne() {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $SiteToken
    )

    $ErrorActionPreference = "stop"

    # Parameters
    $DownloadUrl = "https://s3.amazonaws.com/netlink-software-downloads/AutomateDependencies/SentinelOne/SentinelOneAgent_Windows.exe"
    $DestinationPath = "C:\Covi"
    $Date = Get-Date -Format "MM-yyyy"
    $DestinationFile = Join-Path $DestinationPath -ChildPath "SentinelOneAgent_Windows-$Date.exe"
    New-Item -ItemType Directory -Force -Path $DestinationPath

    <#
    Functions
#>
    function Remove-Webroot() {
        $WebrootService = Get-Service | Where-Object Name -eq "WRSVC" | Select-Object -Property Name, Status, StartType

        if ($WebrootService.Status -eq "Stopped" -and $WebrootService.StartType -eq "Disabled") {
            Write-Output "Webroot is disabled so we can continue the install."
        }
        else {
            Start-BitsTransfer `
                -Source "http://anywhere.webrootcloudav.com/zerol/wsasme.msi" `
                -Destination "C:\Covi\Webroot.msi"
    
            msiexec /uninstall "C:\Covi\Webroot.msi" /qn

            if (Get-Service | Where-Object Name -eq "WRSVC") {
                throw "Webroot is still installed. Please manually remove it and try, again."
            }

            Write-Output "Webroot was successfully removed!"
        }
    }

    function Remove-Cylance() {
        Start-Process  `
            -FilePath "msiexec.exe" `
            -ArgumentList "/x {2E64FC5C-9286-4A31-916B-0D8AE4B22954} /qn MSIRESTARTMANAGERCONTROL=Disable" `
            -Wait `
            -NoNewWindow

        if (Get-Service | Where-Object Name -eq "CylanceSvc") {
            throw "CylanceSvc is still installed. Please manually remove it and try, again."
        }

        Write-Output "Cylance was successfully removed!"
    }

    <#
    Start
#>

    # Making sure the computer is X64
    if (![Environment]::Is64BitOperatingSystem) {
        throw "This script is only designed to run on X64 platforms."
    }

    # Check to make sure it's not already installed.
    if ((Get-Service | Where-Object Name -eq "SentinelAgent")) {
        Write-Output "Sentinel One is already installed."
        exit
    }

    # Making sure other AVs aren't installed.
    if ((Get-Service | Where-Object Name -eq "WRSVC")) {
        Write-Output "Webroot is installed. Attemtping to uninstall..."
        Remove-Webroot
    }

    if ((Get-Service | Where-Object Name -eq "CylanceSvc")) {
        Write-Output "Cylance is installed. Attempting to uninstall..."
        Remove-Cylance
    }

    if (!(Test-Path -Path $DestinationFile)) {
        Write-Output "Downloading SentinelOne..."
        Start-BitsTransfer -Source $DownloadUrl -Destination $DestinationFile
    }

    # The -Wait parameter ensures this never finishes for some reason.
    Write-Output "Installing SentinelOne..."
    Start-Process `
        -FilePath $DestinationFile `
        -ArgumentList "/SITE_TOKEN=$SiteToken /NORESTART /SILENT"

    # Waiting to see if it's installed.
    $Attempts = 0
    $Installed = (Get-Service | Where-Object Name -eq "SentinelAgent")
    do {
        $Installed = (Get-Service | Where-Object Name -eq "SentinelAgent")
        $Attempts++
        Start-Sleep -Seconds 5
    } until ($Attempts -ge 12 -or $Installed)

    if (!$Installed) {
        throw "SentinelOne failed to install."
    }

    Write-Output "SentinelOne installed successfully!"
}

# Taken from LTPosh: https://github.com/LabtechConsulting/LabTech-Powershell-Module/blob/master/LabTech.psm1
If (($MyInvocation.Line -match 'Import-Module' -or $MyInvocation.MyCommand -match 'Import-Module') -and -not ($MyInvocation.Line -match $ModuleGuid -or $MyInvocation.MyCommand -match $ModuleGuid)) {
    # Only export module members when being loaded as a module
    Export-ModuleMember -Function $PublicFunctions
}