<#
.SYNOPSIS
    This script exports selected mailbox types (Active, Disabled, Soft Deleted, and Shared) from Exchange Online to a specified CSV file.

.DESCRIPTION
    The script connects to Exchange Online using provided credentials and exports the requested types of mailboxes to a CSV file.
    It allows you to specify the type of mailboxes to export (Active, Disabled, Soft Deleted, Shared), and exports all mailbox data or specified attributes (depending on the CSV configuration).
    The script requires the ExchangeOnlineManagement module and log4net assembly.

.PARAMETER Account
    Optional. Specifies the account name to connect to Exchange Online. If not provided, you will be prompted to enter the credentials manually.

.PARAMETER Password
    Optional. Specifies the password for the account. If not provided, you will be prompted to enter the credentials manually.

.PARAMETER FilePath
    Mandatory. Specifies the path to the CSV file where the mailbox data will be exported.

.PARAMETER IncludeActiveMailboxes
    Optional. A switch to indicate whether to export Active mailboxes. By default, this is false.

.PARAMETER IncludeDisabledMailboxes
    Optional. A switch to indicate whether to export Disabled mailboxes. By default, this is false.

.PARAMETER IncludeSoftDeletedMailboxes
    Optional. A switch to indicate whether to export Soft Deleted mailboxes. By default, this is false.

.PARAMETER IncludeSharedMailboxes
    Optional. A switch to indicate whether to export Shared mailboxes. By default, this is false.

.EXAMPLE
    .\Exchange_Online_Mailboxes.ps1 -Account "admin@contoso.com" -Password "P@ssword" -FilePath "C:\Users\Administrator\Desktop\" -IncludeActiveMailboxes -IncludeDisabledMailboxes -IncludeSoftDeletedMailboxes -IncludeSharedMailboxes

    This command will export all active, disabled, soft deleted, and shared mailboxes to the specified CSV file.

    .\Exchange_Online_Mailboxes.ps1 -Account "admin@contoso.com" -Password "P@ssword" -FilePath "C:\Users\Administrator\Desktop\" -IncludeSharedMailboxes

    This command will export all shared mailboxes to the specified CSV file.

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
    [string]$FilePath = [System.String]::Empty,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeActiveMailboxes = $false,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabledMailboxes = $false,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSoftDeletedMailboxes = $false,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSharedMailboxes = $false
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

    # Import required ExchangeOnlineManagement module
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

    #Export all active mailboxes to a CSV file
    function Get-ActiveMailboxes {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {    
            $logger.Info("Exporting active mailboxes to $ExportPath")
            Get-EXOMailbox -ResultSize unlimited -PropertySets All |
            ForEach-Object {
                $mailboxInfo = $_ | Select-Object $SelectedAttributes
                
                return $mailboxInfo
            } | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            $logger.Info("Exported active mailboxes to $ExportPath")
        }
    }    

    #Export all disabled mailboxes to a CSV file
    function Get-DisabledMailboxes {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {
            $logger.Info("Exporting disabled mailboxes to $ExportPath")
            #Get all the Exchange online mailboxes statistics
            $MailboxStatistics = Get-EXOMailbox -ResultSize unlimited -PropertySets All | Get-EXOMailboxStatistics
    
            #Filter for disabled mailboxes
            $DisabledMailboxes = $MailboxStatistics | Where-Object { $true -eq $_.AccountDisabled }
    
            #Select the specified attributes and directly export the result to a CSV file
            $DisabledMailboxes | ForEach-Object {
                $mailboxInfo = $_ | Select-Object $SelectedAttributes
                
                return $mailboxInfo
            } | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            $logger.Info("Exported disabled mailboxes to $ExportPath")
        }
    }    

    #Export all soft deleted mailboxes to a CSV file
    function Get-SoftDeletedMailboxes {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {
            $logger.Info("Exporting soft deleted mailboxes to $ExportPath")
            Get-EXOMailbox -SoftDeletedMailbox -ResultSize unlimited -PropertySets All | ForEach-Object {
                $mailboxInfo = $_ | Select-Object $SelectedAttributes

                return $mailboxInfo
            } | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            $logger.Info("Exported soft deleted mailboxes to $ExportPath")
        }
    }    

    #Export all shared mailboxes to a CSV file
    function Get-SharedMailboxes {
        param (
            [Parameter(Mandatory = $true)]
            [System.Array]$SelectedAttributes,
            [Parameter(Mandatory = $true)]
            [string]$ExportPath,            
            [Parameter(Mandatory = $true)]
            [log4net.Core.LogImpl]$logger
        )
        process {
            $logger.Info("Exporting shared mailboxes to $ExportPath")
            Get-EXOMailbox -ResultSize unlimited -PropertySets All |
            Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' } |
            ForEach-Object {
                $properties = $_ | Select-Object $SelectedAttributes
                
                $properties
            } |
            Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            $logger.Info("Exported shared mailboxes to $ExportPath")
        }
    }
    
    #endregion

    #Region: logging configuration initialization and password

    #Get credentials
    $Credentials = GetCredentials -Account $Account -Password $Password
    #Initialize logging configuration
    $configPath = $([System.IO.Path]::Combine($scriptPath, "Configurations\log4net_Exchange_Online_Mailboxes.config"))
    $fileinfo = New-Object System.IO.FileInfo($configPath)
    [log4net.Config.XmlConfigurator]::Configure($fileinfo)
    $logger = [log4net.LogManager]::GetLogger([System.Management.Automation.PowerShell])

    #endregion
}
PROCESS {
    #Import the required attributes from a CSV file and create an array
    try {
        $Attributes = Import-Csv -Path $([System.IO.Path]::Combine($scriptPath, "Configurations\Attributes_Exchange_Online_Mailboxes.csv")) | Where-Object { $_.Required -eq "YES" } | ForEach-Object { $_.Attributes }
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

    #Get selected mailboxes
    $logger.Info("Getting mailboxes")
    try {
        if ($IncludeActiveMailboxes) {
            Get-ActiveMailboxes -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\ExchangeOnline_ActiveMailboxes ${date}.csv") -logger $logger
        }
        if ($IncludeDisabledMailboxes) {
            Get-DisabledMailboxes -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\ExchangeOnline_DisabledMailboxes ${date}.csv") -logger $logger
        }
        if ($IncludeSoftDeletedMailboxes) {
            Get-SoftDeletedMailboxes -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\ExchangeOnline_SoftDeletedMailboxes ${date}.csv") -logger $logger
        }
        if ($IncludeSharedMailboxes) {
            Get-SharedMailboxes -SelectedAttributes $Attributes -ExportPath (Join-Path $FilePath "\ExchangeOnline_SharedMailboxes ${date}.csv") -logger $logger
        }
        if ($IncludeActiveMailboxes -eq $false -and $IncludeDisabledMailboxes -eq $false -and $IncludeSoftDeletedMailboxes -eq $false -and $IncludeSharedMailboxes -eq $false) {
            $logger.Warn("No mailboxes selected! Please select at least one mailbox type what you want to export")
        }
    }
    catch {
        $logger.Error("Couldn't export all selected mailboxes types")
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