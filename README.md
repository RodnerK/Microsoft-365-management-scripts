# Microsoft-365-management-scripts

This repository houses a collection of PowerShell scripts designed to export various types of data from Microsoft 365 services. Each script is crafted to efficiently extract specific data elements from services including OneDrive, SharePoint, Azure Active Directory, Exchange, and Microsoft Teams.

Scripts:

OneDrive for Business - Personal Sites:
This script connects to SharePoint Online and exports all personal sites (OneDrive for Business sites) along with selected site properties. The export is done per multi geo location based on a provided list of SharePoint Admin Centers (OneDrive for Business personal sites). This script uses the SharePointPnPPowerShellOnline PowerShell module to authenticate and extract the data.

SharePoint Site Collections:
This script is designed to connect to SharePoint Online and export all site collections and their specific properties for each specified SharePoint Admin Center. This script uses the SharePointPnPPowerShellOnline PowerShell module to authenticate and extract the data.

Azure Active Directory Guest Users:
This PowerShell script is used to export guest users from Azure Active Directory along with specific attributes. This script uses the MSOnline PowerShell module to authenticate and extract the data.

Exchange Online Mailboxes:
This script exports mailbox details from Exchange Online. It allows you to extract specific attributes and properties for each user mailbox. You can specify to export Active-, Disabled-, Shared- and Soft deleted mailboxes.

Microsoft Teams Data:
This script exports detailed information about each Microsoft Teams users and policies.

Additional Resources:
The repository also includes the following resources:

log4net assembly: A powerful logging framework used for extensive logging of the operations performed by the scripts.
Configuration Files: These files allow you to specify log4net settings for each script.
Attributes CSV Files: These files list the attributes to be exported by each script. You can easily modify these files to change the set of attributes that get exported.
Admin Centers CSV File: This file provides a list of SharePoint Admin Centers for use in the OneDrive and SharePoint export scripts.
Utilities Folder: This folder contains script(s) to help you discover and export attributes for various commands.

Each script in this repository is designed with robust error handling and logging mechanisms to ensure that the data extraction process is as smooth as possible. They are also equipped with easy-to-modify parameters, making these scripts highly adaptable to your specific needs.

All scripts have been thoroughly tested and confirmed to be working as expected. However, please note that these scripts are being provided as-is, and they may not be maintained or updated in the future. They are intended to serve as starting points, and you're encouraged to modify and adapt them to suit your specific requirements.

While these scripts were crafted with care, it's always recommended to execute them in a test environment before running in a production scenario. This will help you understand their operation and adjust any parameters as necessary.

I hope you find these scripts and the additional resources helpful in your data export and reporting tasks within Microsoft 365.