<#
.SYNOPSIS
A script to export Microsoft Teams users and policy configurations into CSV files.

.DESCRIPTION
This script connects to Microsoft Teams using provided credentials and extracts specified data about users and policy configurations.
The data is then exported to CSV files at the specified file path. 

.PARAMETER Account
Optional parameter. The user account credential used to connect to Microsoft Teams. If not provided, the script will prompt for it during runtime.

.PARAMETER Password
    Optional parameter. The password for the user account. If not provided, the script will prompt for it during runtime.

.PARAMETER FilePath
    Mandatory parameter. The location where the exported CSV files will be stored.

.PARAMETER IncludeTeamsUsers
    Optional switch parameter. If included, the script will export Microsoft Teams users data.

.PARAMETER IncludeTeamsCallingPolicy
    Optional switch parameter. If included, the script will export Microsoft Teams calling policy configurations.

.PARAMETER IncludeTeamsMeetingPolicy
    Optional switch parameter. If included, the script will export Microsoft Teams meeting policy configurations.

.PARAMETER IncludeTeamsMessagingPolicy
    Optional switch parameter. If included, the script will export Microsoft Teams messaging policy configurations.

.EXAMPLE
    .\Export-TeamsData.ps1 -Account "admin@contoso.com" -Password "Pa$$w0rd" -FilePath "C:\Export" -IncludeTeamsUsers -IncludeTeamsMeetingPolicy
    This example exports the Teams users and Teams meeting policy data to the "C:\Export" directory.

    .\Export-TeamsData.ps1 -Account "admin@contoso.com" -Password "Pa$$w0rd" -FilePath "C:\Export" -IncludeTeamsUsers -IncludeTeamsMeetingPolicy -IncludeTeamsCallingPolicy -IncludeTeamsMessagingPolicy
    This example exports the Teams users, Teams meeting policy, Teams calling policy and Teams messaging policy data to the "C:\Export" directory.

.NOTES
    1. Make sure you have the MicrosoftTeams module installed before running this script. (Install-Module -Name MicrosoftTeams)
    2. Ensure that log4net.dll is located in the assemblies folder.
    3. The 'Required' column of each property has to be set to 'YES' for it to be included in the exported CSV file (see the Attributes*.csv configuration file).
    4. Please note that filtering for nested objects is not implemented in this version of the script. Therefore, properties of nested objects will not be included correctly in the exported CSV file.0
#>

[CmdletBinding()]
PARAM (
    [Parameter(Mandatory = $false)]
    [string]$Account = [System.String]::Empty,
    [Parameter(Mandatory = $false)]
    [string]$Password = [System.String]::Empty,
    [Parameter(Mandatory = $true)]
    [string]$FilePath = [System.String]::Empty,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeTeamsUsers,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeTeamsCallingPolicy,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeTeamsMeetingPolicy,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeTeamsMessagingPolicy
) 
BEGIN {
        
    #Region: variables

    Set-Location $PSScriptRoot
    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = "Stop"

    $scriptPath = (Get-Item $PSScriptRoot).Parent.FullName
    
    $Culture = (Get-Culture).DateTimeFormat 
    $DateTimeShortFormat = $Culture.ShortDatePattern + " " + $Culture.ShortTimePattern
    $CurrentDateTime = Get-Date
    $date = $CurrentDateTime.ToString($DateTimeShortFormat) -replace("/", ".") -replace(":", ".")

    #endregion

    #Region: modules and assemblies

    #Import MicrosoftTeams module
    try {
        if (!(Get-Module -Name MicrosoftTeams)) {
            Import-Module -Name MicrosoftTeams -NoClobber
        }
    }
    catch {
        throw "Couldn't load MicrosoftTeams module. PLease make sure the MicrosoftTeams module is installed [Install-Module -Name MicrosoftTeams]"
    }
    
    #Load assemblies
    try {
        Add-Type -Path $([System.IO.Path]::Combine($scriptPath, "assemblies\log4net.dll"))
    }
    catch {
        throw "Couldn't load log4net assembly. Please make sure the log4net.dll is present in the files folder"
    }

    #endregion

    #Region: functions

    #Get credentials
    function GetCredentials {
        param (
            [Parameter(Mandatory = $false)]
            [string]$Account,
            [Parameter(Mandatory = $false)]
            [string]$Password
        )
        PROCESS {
            if (![string]::IsNullOrEmpty(($Account)) -or ![string]::IsNullOrEmpty(($Password))) {
                $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Account, ($Password | ConvertTo-SecureString -AsPlainText -Force)
            }
            else {
                Write-Host $("Account or Password is empty") -ForegroundColor Red
                $Credentials = Get-Credential -Message "Enter admin credentials for SharePoint Online"
            }
            if ($null -eq $Credentials) {
                throw "Credentials are empty"
            }
            else {
                Return $Credentials
            }
        }
    }

    #Get all Microsoft Teams users
    function Get-TeamsUsers {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {
            $logger.Info("Getting Teams Users")
            Get-CsOnlineUser |
            ForEach-Object {
                #Apply the passed array to each user
                $user = $_ | Select-Object $SelectedAttributes
    
                return $user
            } |
            #Export users to CSV
            Export-Csv $ExportPath -NoTypeInformation
            $logger.Info("Teams Users exported to CSV")
        }
    }
    
    #Get all Microsoft Teams Calling Policies
    function Get-TeamsCallingPolicy {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {
            $logger.Info("Getting Teams Calling Policies")
            Get-CsTeamsCallingPolicy |
            ForEach-Object {
                #Apply the passed array to each policy
                $policy = $_ | Select-Object $SelectedAttributes
    
                return $policy
            } |
            #Export policies to CSV
            Export-Csv $ExportPath -NoTypeInformation
            $logger.Info("Teams Calling Policies exported to CSV")
        }
    }
    
    #Get all Microsoft Teams Meeting Policies
    function Get-TeamsMeetingPolicy {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {
            $logger.Info("Getting Teams Meeting Policies")
            Get-CsTeamsMeetingPolicy |
            ForEach-Object {
                #Apply the passed array to each policy
                $policy = $_ | Select-Object $SelectedAttributes
    
                return $policy
            } |
            #Export policies to CSV
            Export-Csv $ExportPath -NoTypeInformation
            $logger.Info("Teams Meeting Policies exported to CSV")
        }
    }
    
    #Get all Microsoft Teams Messaging Policies
    function Get-TeamsMessagingPolicy {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {
            $logger.Info("Getting Teams Messaging Policies")
            Get-CsTeamsMessagingPolicy |
            ForEach-Object {
                #Apply the passed array to each policy
                $policy = $_ | Select-Object $SelectedAttributes

                return $policy
            } |
            #Export policies to CSV
            Export-Csv $ExportPath -NoTypeInformation
            $logger.Info("Teams Messaging Policies exported to CSV")
        }
    }
    
    #endregion

    #Region: logging configuration initialization and password

    #Get credentials
    $Credentials = GetCredentials -Account $Account -Password $Password
    #Initialize logging configuration
    $configPath = $([System.IO.Path]::Combine($scriptPath, "Configurations\log4net_MicrosoftTeams_Policies.config"))
    $fileinfo = New-Object System.IO.FileInfo($configPath)
    [log4net.Config.XmlConfigurator]::Configure($fileinfo)
    $logger = [log4net.LogManager]::GetLogger([System.Management.Automation.PowerShell])

    #endregion
}
PROCESS {
    #Connect to Microsoft Teams
    try {
        Connect-MicrosoftTeams -Credential $Credentials
        $logger.Info("Connected to Microsoft Teams")
    }
    catch {
        $logger.Error("Couldn't connect to Microsoft Teams")
        throw $_
    }

    #Export the selected reports
    try {
        if ($IncludeTeamsUsers) {
            $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_MicrosoftTeams_User.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
        
            $logger.Info("Required attributes for Teams Users imported")

            Get-TeamsUsers -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\MicrosoftTeams_Users ${date}.csv") -logger $logger
        }
        if ($IncludeTeamsCallingPolicy) {
            $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_MicrosoftTeams_CallingPolicy.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
        
            $logger.Info("Required attributes for Teams Calling Policies imported")

            Get-TeamsCallingPolicy -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\MicrosoftTeams_CallingPolicy ${date}.csv") -logger $logger
        }
        if ($IncludeTeamsMeetingPolicy) {
            $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_MicrosoftTeams_MeetingPolicy.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
        
            $logger.Info("Required attributes for Teams Meeting Policies imported")

            Get-TeamsMeetingPolicy -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\MicrosoftTeams_MeetingPolicy ${date}.csv") -logger $logger
        }
        if ($IncludeTeamsMessagingPolicy) {
            $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_Microsofteams_MessagingPolicy.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
        
            $logger.Info("Required attributes for Teams Messaging Policies imported")

            Get-TeamsMessagingPolicy -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\MicrosoftTeams_MessagingPolicy ${date}.csv") -logger $logger
        }
        if ($IncludeTeamsUsers -eq $false -and $IncludeTeamsCallingPolicy -eq $false -and $IncludeTeamsMeetingPolicy -eq $false -and $IncludeTeamsMessagingPolicy -eq $false) {
            $logger.Warn("No policies or user selected! Please select at least one policy or the user.")
        }
    }
    catch {
        $logger.Error("Couldn't export policies")
        throw $_     
    }

    #Disconnect from Microsoft Teams
    try {
        Disconnect-MicrosoftTeams 
        $logger.Info("Disconnected from Microsoft Teams")
    }
    catch {
        $logger.Error("Couldn't disconnect from Microsoft Teams")
        throw $_ 
    }
}   
END {
    #Reset ErrorActionPreference
    $ErrorActionPreference = $oldEAP
}