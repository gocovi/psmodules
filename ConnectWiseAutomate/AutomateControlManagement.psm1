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

function Start-UpdateCheckLoop($LatestVersion) {
    $Count = 0
    $Success = $false

    do {
        if ((Get-LTServiceInfo).Version -ne $LatestVersion) {
            Start-Sleep -Seconds 5
        }
        else {
            $Success = $true
        }

        $Count++
    } until ($Count -eq 3 -or $Success)

    return $Success
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
            "short_message"       = "Checking for Automate updates via Control.";
            "long_message"        = $Message;
            "trigger_name"        = "Manually ran by Jack Musick";
            "company_id"          = "$($LTServiceInfo.ClientID)";
            "computer_name"       = $env:COMPUTERNAME;
            "computer_id"         = "$($LTServiceInfo.ID)";
            "location_name"       = "Unknown";
            "location_id"         = "$LocationID";
            "current_version"     = $LTServiceInfo.Version;
            "latest_version"      = $LatestVersion;
            "status"              = "$Status";
            "company"             = "Unknown"
        } | ConvertTo-Json
        
        Invoke-WebRequest `
            -Uri "https://logconnector.gocovi.com" `
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

        [Switch]$Update,

        [Switch]$Force,

        [Switch]$Verbose,

        [string]$Server
    )

    if ($CoviApiKey) {
        $script:CoviApiKey = $CoviApiKey
    }

    try {
        $LTService = Get-Service | Where-Object { $_.Name -eq "LTService" }

        if (!$LTService) {
            Write-Output "LTService is not installed."
        }
        else {
            # Making sure service is running
            if ($LTService.Status -ne "Running") {
                try {
                    # Killing the process if it's hung
                    $LTProcessID = Get-WmiObject -Class Win32_Service -Filter "Name LIKE 'LTService'" | Select-Object -ExpandProperty ProcessId
                    
                    if ($LTProcessID) {
                        taskkill /f /pid $LTProcessID
                    }

                    $LTService | Start-Service -Confirm:$False
                }
                catch {
                    Write-Output "LTService couldn't be started. We will try the update anyways, but this device could have an issue."
                    Write-CoviLog -Status "Error" -Message "LTService couldn't be started. We will try the update anyways, but this device could have an issue."
                }
            }

            Write-Host "Importing Labtech Powershell Module..."

            (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/LabtechConsulting/LabTech-Powershell-Module/master/LabTech.psm1') | Invoke-Expression

            $LatestVersion = (Get-AutomateLatestVersion)

            if (!$LatestVersion) {
                Write-Error "Unable to get the latest version from Automate."
                exit
            }

            $LTServiceInfo = Get-LTServiceInfo

            if ($Server) {
                if ($LTServiceInfo.'Server Address' -notlike "*$Server*") {
                    Write-Output "Mismatched server found: $($LTServiceInfo.Server)"
                    Write-CoviLog -Status "Error" -Message "Mismatched server found: $($LTServiceInfo.Server)"
                    exit
                }
            }

            if ($LTServiceInfo.Version -eq $LatestVersion) {
                Write-Output "This Automate agent is running the latest version."

                # Mostly used to confirm logging works
                if ($Verbose) {
                    Write-CoviLog -Status "Verbose" -Message "This Automate agent is running the latest version."
                }
            }
            else {
                Write-Output "Current Version: $((Get-LTServiceInfo).Version)`nLatest Version: $LatestVersion"

                Write-CoviLog -Status "Information" -Message "An update is available for this agent."

                # Flag to do the actual update.
                if ($Update) {
                    Update-LTService -Confirm:$False

                    $Success = Start-UpdateCheckLoop -LatestVersion $LatestVersion

                    # Rechecking to make sure the agent got updated.
                    if (!$Success -and $Force) {
                        Redo-LTService -Confirm:$False
                    }

                    $Success = Start-UpdateCheckLoop -LatestVersion $LatestVersion

                    # Checking again after a full reinstall
                    if (!$Success) {
                        Write-Output "Agent failed to update to the latest version."
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
    catch {
        Write-Output "Unknown Error: $($_.Exception.Message)"
        Write-CoviLog -Status "Failed" -Message "Unknown Error: $($_.Exception.Message)"
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