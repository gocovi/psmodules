<#  
.SYNOPSIS  

.DESCRIPTION  

.NOTES  
    File Name  : AutomateControlManagement.ps1
    Author     : jack@gocovi.com
    Requires   : 

.LINK 
#>

$CoviApiKey = $null

function Get-AutomateLatestVersion() {
    $LTServiceInfo = Get-LTServiceInfo
    $Server = $LTServiceInfo.'Server Address'.Split("|")[0]
    $Response = Invoke-RestMethod -Uri "$Server/LabTech/Agent.aspx"
    return $Response.Replace("||||||", "")
}

function Write-CoviLog($Status, $Message) {
    if ($script:CoviApiKey) {
        $LTServiceInfo = Get-LTServiceInfo
        $LatestVersion = Get-AutomateLatestVersion

        if (!$Message) {
            $Message = "None"
        }

        $Body = [PSCustomObject]@{
            "service_name"        = "ConnectWise Automate";
            "service_description" = "Agent Version Check";
            "sent_from"           = "ConnectWise Control";
            "short_message"       = "Current Version: $StartVersion   Latest Version: $LatestVersion";
            "long_message"        = $Message;
            "trigger_name"        = "Manually ran by Jack Musick";
            "company_id"          = "$($LTServiceInfo.ClientID)";
            "computer_name"       = $env:COMPUTERNAME;
            "computer_id"         = "$($LTServiceInfo.ID)";
            "location_name"       = "Unknown";
            "location_id"         = "$LocationID";
            "status"              = "$Status";
            "company"             = "Unknown"
        } | ConvertTo-Json


        Invoke-WebRequest `
            -Uri "https://api.gocovi.com/standard/covi/logs" `
            -Headers @{ "x-api-key" = $script:CoviApiKey } `
            -UseBasicParsing `
            -Method Post `
            -Body $Body `
            -ContentType "application/json" | Out-Null
    }
}

# Uses LTPosh to compare the server version to the version on the agent.
function Confirm-AutomateLatestVersion() {
    param (
        [string]$CoviApiKey,

        [Switch]$Update
    )

    if ($CoviApiKey) {
        $script:CoviApiKey = $CoviApiKey
    }

    $LTService = Get-Service | Where-Object { $_.Name -eq "LTService" }

    if (!$LTService) {
        Write-Output "LTService is not installed."
    }
    else {
        Write-Host "Importing Labtech Powershell Module..."
        (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/LabtechConsulting/LabTech-Powershell-Module/master/LabTech.psm1') | Invoke-Expression

        $LatestVersion = (Get-AutomateLatestVersion)

        if (!$LatestVersion) {
            Write-Error "Unable to get the latest version from Automate."
            exit
        }

        if ((Get-LTServiceInfo).Version -eq $LatestVersion) {
            Write-Output "This Automate agent is running the latest version."
        }
        else {
            Write-Output "Current Version: $((Get-LTServiceInfo).Version)`nLatest Version: $LatestVersion"

            Write-CoviLog -Status "Warning" -Message "An update is available for this agent."

            if ($Update) {
                Update-LTService -Confirm:$False

                # Rechecking to make sure the agent got updated.
                if ((Get-LTServiceInfo).Version -ne $LatestVersion) {
                    Redo-LTService -Confirm:$False
                }

                # Checking again after a full reinstall
                if ((Get-LTServiceInfo).Version -ne $LatestVersion) {
                    Write-Output "Agent failed to update to the latest version by both updating and reinstalling."
                    Write-CoviLog -Status "Failed"
                }
                else {
                    Write-Output "Agent updated successfully."
                    Write-CoviLog -Status "Success"
                }
            }
        }
    }
}

$PublicFunctions = @(
    "Confirm-AutomateLatestVersion"
)

# Taken from LTPosh: https://github.com/LabtechConsulting/LabTech-Powershell-Module/blob/master/LabTech.psm1
If (($MyInvocation.Line -match 'Import-Module' -or $MyInvocation.MyCommand -match 'Import-Module') -and -not ($MyInvocation.Line -match $ModuleGuid -or $MyInvocation.MyCommand -match $ModuleGuid)) {
    # Only export module members when being loaded as a module
    Export-ModuleMember -Function $PublicFunctions
}