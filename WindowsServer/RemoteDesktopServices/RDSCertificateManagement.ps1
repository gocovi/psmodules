<#  
.SYNOPSIS  
    Pulls the most recent certificate from IIS for a specified domain and sets it for all RDS services.
.DESCRIPTION 
    1. Install your certificate in IIS.
    2. Import the module:
    
        (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gocovi/psmodules/master/WindowsServer/RemoteDesktopServices/RDSCertificateManagement.ps1') | iex

    3. Run 'Set-RDCertificatesFromIIS -Domain remote.yourdomain.com

    
.NOTES  
    File Name  : RDSCertificateManagement.ps1
    Author     : jack@gocovi.com
    Requires   : 

.LINK 


#>

function Set-RDCertificatesFromIIS {
    param (
        [Parameter(Mandatory = $True, Position = 1)]
        [string]$Domain
    )

    # Session settings so it doesn't stay static.
    $SessionGuid = (New-Guid).Guid
    Start-Sleep -Seconds 2
    $TemporaryPassword = (New-Guid).Guid

    $Certificates = Get-ChildItem `
        -Path cert:\localMachine\my | `
        Where-Object Subject -like "*$Domain*" | `
        Sort-Object -Property NotAfter -Descending

    # Only set if the date is greater than today.
    if ($Certificates[0] -and $Certificates[0].NotAfter -gt (Get-Date)) {
        $TempDirectory = "C:\TemporaryCertificates"
        $TempCertificate = "$TempDirectory\$SessionGuid.pfx"

        New-Item -ItemType Directory -Force -Path $TempDirectory

        $Password = ConvertTo-SecureString -String $TemporaryPassword -AsPlainText -Force

        # Exporting the certificate from IIS
        $Certificates[0] | Export-PfxCertificate -FilePath $TempCertificate -Password $Password

        # Importing the certificate into RDS
        Set-RDCertificate -Role RDGateway -ImportPath $TempCertificate -Password $Password -Force
        Set-RDCertificate -Role RDWebAccess -ImportPath $TempCertificate -Password $Password -Force
        Set-RDCertificate -Role RDRedirector -ImportPath $TempCertificate -Password $Password -Force
        Set-RDCertificate -Role RDPublishing -ImportPath $TempCertificate -Password $Password -Force

        # Cleanup
        Remove-Item -Path $TempCertificate -Force
    }
}