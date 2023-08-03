<#
.SYNOPSIS
    This script connects to Azure Active Directory and exports guest users to a CSV file.

.DESCRIPTION
    The script uses the MSOnline PowerShell module to connect to Azure AD using provided credentials.
    It then queries all guest users and exports their specific attributes to a CSV file.
    The script requires the log4net assembly for logging purposes.

.PARAMETER Account
    The username to connect to Azure Active Directory. If not provided, script will prompt for it.

.PARAMETER Password
    The password to connect to Azure Active Directory. If not provided, script will prompt for it.

.PARAMETER FilePath
    The path to the directory where the CSV file will be exported. This parameter is mandatory.

.EXAMPLE
    .\AzureActiveDirectory_GuestUsers.ps1 -Account "admin@contoso.com" -Password "password123" -FilePath "C:\Exports"
    This example shows how to use this script to export guest users to the C:\Exports directory.

.OUTPUTS
    CSV file with Azure AD user data.
    
.NOTES
    1. Make sure you have the MSOnline module installed before running this script. (Install-Module -Name MSOnline)
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

    #Import MSOnline module
    try {
        if (!(Get-Module -Name MSOnline)) {
            Import-Module -Name MSOnline -NoClobber
        }
    }
    catch {
        throw "Couldn't load MSOnline module. PLease make sure the MSOnline module is installed [Install-Module -Name MSOnline]"
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
            [Parameter(Mandatory = $true)]
            [string]$Account,
            [Parameter(Mandatory = $true)]
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

    #Export guest users from Azure Active Directory
    function Export-GuestUsers {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $false)]
            [string]$ExportPath,
            [Parameter(Mandatory = $false)]
            [log4net.Core.LogImpl]$logger
        )
        PROCESS {    
            $logger.Info("Querying all guest users and creating results")
            Get-MsolUser -All | Where-Object { $_.UserType -like "Guest" } | ForEach-Object {
                
                $user = $_ | Select-Object $SelectedAttributes
                return $user
            } | 
            Export-Csv -NoTypeInformation -Path $ExportPath
    
            $logger.Info("Export completed to $ExportPath")
        }
    }    

    #endregion

    #Region: logging configuration initialization and password

    #Get credentials
    $Credentials = GetCredentials -Account $Account -Password $Password
    #Initialize logging configuration
    $configPath = $([System.IO.Path]::Combine($scriptPath, "Configurations\log4net_AzureActiveDirectory_GuestUsers.config"))
    $fileinfo = New-Object System.IO.FileInfo($configPath)
    [log4net.Config.XmlConfigurator]::Configure($fileinfo)
    $logger = [log4net.LogManager]::GetLogger([System.Management.Automation.PowerShell])

    #endregion

}
PROCESS {
    #Import the required attributes from a CSV file and create an array
    try {
        $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_AzureActiveDirectory_GuestUsers.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
    }
    catch {
        $logger.Error("Couldn't import attributes from CSV file")
        throw $_        
    }

    #Connect to Azure Active Directory
    try {
        Connect-MsolService -Credential $Credentials
        $logger.Info("Connected to Azure Active Directory")
    }
    catch {
        $logger.Error("Couldn't connect to Azure Active Directory", $_.Exception)
        throw $_
    }
    
    #Export guest users from Azure Active Directory
    try {
        Export-GuestUsers -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\AzureActiveDirectory_GuestUsers ${date}.csv") -logger $logger
    }
    catch {
        $logger.Error("Couldn't export guest users from Azure Active Directory", $_.Exception)
        throw $_
    }
}
END {
    #Reset error action preference
    $ErrorActionPreference = $oldEAP
}