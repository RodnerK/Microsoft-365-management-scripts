<#
.SYNOPSIS
    This script connects to Exchange Online and exports the properties of Office 365 Unified Groups to a CSV file.

.DESCRIPTION
    The script will first load the required ExchangeOnlineManagement module and the log4net assembly. It then retrieves the user's credentials and initializes the logging configuration.
    It proceeds to import required attributes from a CSV file and creates a script block for selection.
    Then, the script connects to Exchange Online using the credentials provided, retrieves the properties of Office 365 Unified Groups as defined in the attributes CSV file, and exports them to a specified CSV file.
    Finally, the script disconnects from Exchange Online and resets the error action preference.

.PARAMETER Account
    The account username to connect to Exchange Online. If not provided, the script will prompt for credentials.

.PARAMETER Password
    The account password to connect to Exchange Online. If not provided, the script will prompt for credentials.

.PARAMETER FilePath
    The file path to export the CSV file with Office 365 Unified Groups properties. This parameter is mandatory.

.EXAMPLE
    .\Office365_UnifiedGroups.ps1 -Account "username" -Password "password" -FilePath "C:\Export"

.NOTES
    1. Make sure you have the ExchangeOnlineManagement module installed before running this script. (Install-Module -Name ExchangeOnlineManagement)
    2. Ensure that log4net.dll is located in the assemblies folder.
    3. The 'Required' column of each property has to be set to 'YES' for it to be included in the exported CSV file (see the Attributes*.csv configuration file).
    4. Please note that filtering for nested objects is not implemented in this version of the script. Therefore, properties of nested objects will not be included correctly in the exported CSV file.
#>

[CmdletBinding()]
PARAM (
    [Parameter(Mandatory = $false)]
    [string]$Account = [System.String]::Empty,
    [Parameter(Mandatory = $false)]
    [string]$Password = [System.String]::Empty,
    [Parameter(Mandatory = $true)]
    [string]$FilePath = [System.String]::Empty
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

    #Import required ExchangeOnlineManagement module
    try {
        if (!(Get-Module -Name ExchangeOnlineManagement)) {
            Import-Module -Name ExchangeOnlineManagement -NoClobber
        }
    }
    catch {
        throw "Couldn't load ExchangeOnlineManagement module. PLease make sure the ExchangeOnlineManagement module is installed [Install-Module -Name ExchangeOnlineManagement]"
    }

    #Load assemblies
    try {
        Add-Type -Path $([System.IO.Path]::Combine($scriptPath, "assemblies\log4net.dll"))
    }
    catch {
        throw "Couldn't load log4net assembly. Please make sure the log4net.dll is present in the assemblies folder"
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
                $Credentials = Get-Credential -Message "Enter credentials for Azure Active Directory"
            }
            if ($null -eq $Credentials) {
                throw "Credentials are empty"
            }
            else {
                Return $Credentials
            }
        }
    }

    #Export Office 365 Unified Groups to CSV with selected attributes
    function Get-Office365UnifiedGroup {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
    
        process {
            $logger.Info("Getting Office 365 Unified Groups")
            Get-UnifiedGroup -ResultSize Unlimited | ForEach-Object {
                #Apply the passed array to each group
                $group = $_ | Select-Object $SelectedAttributes
                #Add the date property
                                
                # Output the group object to the pipeline
                return $group
            } | 
            #Export groups to CSV
            Export-Csv $ExportPath -NoTypeInformation

            $logger.Info("Office 365 Unified Groups exported to CSV")
        }
    }
    
    #endregion

    #Region: logging configuration initialization and password

    #Get credentials
    $Credentials = GetCredentials -Account $Account -Password $Password
    #Initialize logging configuration
    $configPath = $([System.IO.Path]::Combine($scriptPath, "Configurations\log4net_Office365_UnifiedGroups.config"))
    $fileinfo = New-Object System.IO.FileInfo($configPath)
    [log4net.Config.XmlConfigurator]::Configure($fileinfo)
    $logger = [log4net.LogManager]::GetLogger([System.Management.Automation.PowerShell])

    #endregion
}
PROCESS {
    #Import the required attributes from a CSV file and create an array
    try {
        $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_Office365_UnifiedGroups.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
    }
    catch {
        $logger.Error("Couldn't import attributes from CSV file")
        throw $_        
    }

    #Connect to Exchange Online
    try {
        Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false
        $logger.Info("Connected to Exchange Online")
    }
    catch {
        $logger.Error("Couldn't connect to Exchange Online")
        throw $_
    }

    #Export Office 365 Unified Groups to CSV with selected attributes
    try {
        Get-Office365UnifiedGroup -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\Office365_UnifiedGroups ${date}.csv") -logger $logger
    }
    catch {
        $logger.Error("Couldn't export Office 365 Unified Groups to CSV")
        throw $_
    }

    #Disconnect from Exchange Online
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        $logger.Info("Disconnected from Exchange Online")
    }
    catch {
        $logger.Error("Couldn't disconnect from Exchange Online")
        throw $_
    }
}
END {
    #Reset error action preference
    $ErrorActionPreference = $oldEAP
}