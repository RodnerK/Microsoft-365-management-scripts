<#
.SYNOPSIS
    This function executes a PowerShell 'Get-' command, retrieves the attributes of the returned object(s), and exports them to a CSV file.

.DESCRIPTION
    This function is designed to help identify the attributes returned by a 'Get-' command, allowing for better script writing and data manipulation.
    The function works by invoking the specified command and selecting the first returned object. It then extracts the attribute names, their data types, and sets a 'Required' column to 'NO'. These data points are then exported to a CSV file at the specified path.

.PARAMETER Command
    A 'Get-' command for which to retrieve attributes. Only 'Get-' commands are allowed. The command must be passed as a string.

.PARAMETER FilePath
    The full file path where the CSV file with the attributes will be created. The path must be passed as a string.

.EXAMPLE
    Get-AttributesOfReturnedObject -Command "Get-Process" -FilePath "C:\Temp\ProcessAttributes.csv"

    This example will retrieve the attributes of the objects returned by the 'Get-Process' command and export them to a CSV file at 'C:\Temp\ProcessAttributes.csv'.

.NOTES
    Please ensure that the command passed does not have any destructive potential as it will be invoked as is. Also, make sure the specified file path is valid and writable.
#>

function Get-AttributesOfReturnedObject {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$Command,
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    Process {
        # Validate the command, that it is only a Get- command
        if ($Command -notmatch '^Get-') {
            throw 'Invalid command. Only Get- commands are allowed.'
        }
        
        # Get the first returned object
        $attr = Invoke-Expression $Command | Select-Object -First 1

        # Get the list of property names, their types, and set 'Required' column to 'NO'
        $attr.PSObject.Properties | 
        Select-Object @{Name = 'Attributes'; Expression = { $_.Name } }, @{Name = 'Attribute Type'; Expression = { $_.TypeNameOfValue } }, @{Name = 'Required'; Expression = { 'NO' } } |
        # Export the properties to a CSV file
        Export-Csv -Path $FilePath -NoTypeInformation
    }
}