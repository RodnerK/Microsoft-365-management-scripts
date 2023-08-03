<#
.SYNOPSIS
    This script connects to Azure Active Directory (Azure AD), retrieves all users, and exports the user data to a specified CSV file.

.DESCRIPTION
    The script authenticates to Azure AD using provided credentials (username and password) and pulls all user information. 
    It then creates an object for each user, capturing specified properties in the attributes csv file.
    The date when the information was fetched is also included. This information is then written to a CSV file at the provided file path.
    
    The script also utilizes the log4net library to generate logs, which can help track its process and debug in case of errors. The 
    log4net configuration is set up at the beginning of the script and used throughout the process.

.PARAMETER Account
    The username to connect to Azure Active Directory. If not provided, script will prompt for it.

.PARAMETER Password
    The password to connect to Azure Active Directory. If not provided, script will prompt for it.

.PARAMETER FilePath
    The path to the directory where the CSV file will be exported. This parameter is mandatory.

.EXAMPLE
    PS> .\AzureActiveDirectory_Users.ps1 -Account "admin@contoso.com" -Password "password123" -FilePath "C:\Exports"
    This example shows how to use this script to export all users to the C:\Exports directory.

.OUTPUTS
    CSV file with Azure AD user data.
    
.NOTES
    1. Make sure you have the AzureAD module installed before running this script. (Install-Module -Name AzureAD)
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

    # Import required AzureAD module
    try {
        if (!(Get-Module -Name AzureAD)) {
            Import-Module -Name AzureAD -NoClobber
        }
    }
    catch {
        throw "Couldn't load AzureAD module. PLease make sure the AzureAD module is installed [Install-Module -Name AzureAD]"
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

    #Export all users
    function Export-AllUsers {
        param (
            [Parameter(Mandatory = $true)]
            [system.array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        PROCESS {    
            $logger.Info("Querying all users and creating results")
            Get-AzureADUser -All $true | ForEach-Object {
    
                $user = $_ | Select-Object $SelectedAttributes
      
                return $user
            } | 
            Export-Csv -Path $ExportPath -NoTypeInformation -Append
    
            $logger.Info("Results export successfully to $ExportPath")
        }
    }    

    #endregion
    
    #Region: logging configuration initialization and password

    #Get credentials
    $Credentials = GetCredentials -Account $Account -Password $Password
    #Initialize logging configuration
    $configPath = $([System.IO.Path]::Combine($scriptPath, "Configurations\log4net_AzureActiveDirectory_Users.config"))
    $fileinfo = New-Object System.IO.FileInfo($configPath)
    [log4net.Config.XmlConfigurator]::Configure($fileinfo)
    $logger = [log4net.LogManager]::GetLogger([System.Management.Automation.PowerShell])

    #endregion
}
PROCESS {
    #Import the required attributes from a CSV file and create an array
    try {
        $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_AzureActiveDirectory_Users.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
    }
    catch {
        $logger.Error("Couldn't import attributes from CSV file")
        throw $_        
    }

    #Connect to Azure Active Directory
    try {
        Connect-AzureAD -Credential $Credentials
        $logger.Info("Connected to Azure Active Directory")
    }
    catch {
        $logger.Error("Couldn't connect to Azure Active Directory")
        throw $_
    }
    
    #Export all users from Azure Active Directory
    try {
        Export-AllUsers -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\AzureActiveDirectory_Users ${date}.csv") -logger $logger
    }
    catch {
        $logger.Error("Couldn't export users from Azure Active Directory")
        throw $_
    }
    
    #Disconnect from Azure Active Directory
    try {
        Disconnect-AzureAD
        $logger.Info("Disconnected from Azure Active Directory")
    }
    catch {
        $logger.Error("Couldn't disconnect from Azure Active Directory")
        throw $_
    }
}
END {
    #Reset error action preference
    $ErrorActionPreference = $oldEAP
}