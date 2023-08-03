<#
.SYNOPSIS
    This script exports all SharePoint Online sites from all specified Multi-Geo locations.

.DESCRIPTION
    The script connects to each Admin Center URL through the specified Multi-Geo locations, retrieves all SharePoint Online sites, and exports the data to a CSV file.
    It makes use of the SharePointPnPPowerShellOnline module and log4net for logging.

.PARAMETER Account
    The account used to connect to the SharePoint Admin Center. If not provided, script will prompt for it.

.PARAMETER Password
    The password for the account used to connect to the SharePoint Admin Center. If not provided, script will prompt for it.

.PARAMETER FilePath
    The path to the directory where the CSV file will be exported. This parameter is mandatory.

.EXAMPLE
    PS> .\SharePointOnline_Sites.ps1 -Account "admin@contoso.com" -Password "Password123" -FilePath "C:\Exports"
    This example shows how to use this script to export all SharePoint Online sites to the C:\Exports directory.
    
.NOTES
    1. Make sure you have the SharePointPnPPowerShellOnline module installed before running this script. (Install-Module -Name SharePointPnPPowerShellOnline)
    2. Ensure that log4net.dll is located in the assemblies folder.
    3. The 'Required' column of each property has to be set to 'YES' for it to be included in the exported CSV file (see the Attributes*.csv configuration file).
    4. Please note that filtering for nested objects is not implemented in this version of the script. Therefore, properties of nested objects will not be included correctly in the exported CSV file.
#>

[CmdletBinding()]
param (
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

    #Import SharePointPnPPowerShellOnline module
    try {
        if (!(Get-Module -Name SharePointPnPPowerShellOnline)) {
            Import-Module -Name SharePointPnPPowerShellOnline -NoClobber
        }
    }
    catch {
        throw "Couldn't load SharePointPnPPowerShellOnline module. PLease make sure the SharePointPnPPowerShellOnline module is installed [Install-Module -Name SharePointPnPPowerShellOnline]"
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

    #Get all SharePoint Online sites from all specified Multi-Geo locations
    function Export-AllSharePointOnlineSites {
        param (
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.PSCredential]$Credentials,
            [Parameter(Mandatory = $true)]
            [string]$AdmincenterListPath,
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        PROCESS {
            #Import SharePointOnline_AdminCenters_list.csv
            try {
                $AdminCenters = Import-Csv -Path $AdmincenterListPath -Delimiter ";"
            }
            catch {
                $logger.Error("Couldn't import SharePointOnline_AdminCenters_list.csv")
                throw $_
            }
    
            foreach ($AdminCenter in $AdminCenters) {
                try {
                    #Connect to SharePoint Online Admin Center
                    Connect-PnPOnline -Url ($AdminCenter.AdminCenterUrl)  -Credentials $Credentials
                    $AdminCenter.AdminCenterUrl
                    $logger.Info("Connected to `"$($AdminCenter.AdminCenterUrl)`"")

                    #Get all SharePoint Online sites
                    $logger.Info("Getting all SharePoint Online sites from `"$($AdminCenter.AdminCenterUrl)`"")
                    Get-PnPTenantSite | 
                    ForEach-Object {
                        
                        $properties = $_ | Select-Object $SelectedAttributes
                        $properties | Add-Member -MemberType NoteProperty -Name "Multi Geo Location" -Value $AdminCenter.MultiGeoLocation
                        
                        return $properties
    
                    } |
                    Export-Csv -Path $ExportPath -NoTypeInformation -Append
                    $logger.Info("Exported all SharePoint Online sites from `"$($AdminCenter.AdminCenterUrl)`"")

                    Disconnect-PnPOnline
                    $logger.Info("Disconnected from `"$($AdminCenter.AdminCenterUrl)`"")
                }
                catch {
                    $logger.Error("Couldn't get all SharePoint Online sites from `"$($AdminCenter.AdminCenterUrl)`"")
                    throw $_
                }
            }
        }       
    }    

    #endregion

    #Region: logging configuration initialization and password

    #Get credentials
    $Credentials = GetCredentials -Account $Account -Password $Password
    #Initialize logging configuration
    $configPath = $([System.IO.Path]::Combine($scriptPath, "Configurations\log4net_SharePointOnline_Sites.config"))
    $fileinfo = New-Object System.IO.FileInfo($configPath)
    [log4net.Config.XmlConfigurator]::Configure($fileinfo)
    $logger = [log4net.LogManager]::GetLogger([System.Management.Automation.PowerShell])

    #endregion
}
PROCESS {
    #Import the required attributes from a CSV file and create a script block
    try {
        $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_SharePointOnline_Sites.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
        
        $logger.Info("Imported attributes from CSV file")
    }
    catch {
        $logger.Error("Couldn't import attributes from CSV file")
        throw $_        
    }

    #Export all SharePoint Online sites from all specified Multi-Geo locations
    try {
        Export-AllSharePointOnlineSites -SelectedAttributes $Attributes -Credentials $Credentials -AdmincenterListPath $([System.IO.Path]::Combine($scriptPath, "Configurations\SharePointOnline_AdminCenters_list.csv")) -ExportPath (Join-Path $FilePath "\SharePointOnline_Sites ${date}.csv") -logger $logger
    }
    catch {
        $logger.Error("Couldn't export all SharePoint Online sites")
        throw $_
    }
}
END {
    #Reset error action preference
    $ErrorActionPreference = $oldEAP
}