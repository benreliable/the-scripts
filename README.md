# The Scripts

This repository contains a collection of PowerShell and batch scripts for IT administration.

## Structure

The repository is organized into the following directories:

*   `docs`: Contains detailed documentation for the scripts.
*   `scripts`: Contains the scripts, organized by function.
*   `tools`: Contains tools and utilities used by the scripts.

## Scripts

### Active Directory

*   [ComputersNotLoggedIn.ps1](scripts/ad/Reports/ComputersNotLoggedIn.ps1)
    *   **Description:** This script retrieves a list of computers that have not logged into the domain in the last 90 days.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Modify the `$OUpath` and `$ExportPath` variables in the script to match your environment.
        3.  Run the script.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [DistributionGroupAudit.ps1](scripts/ad/Reports/DistributionGroupAudit.ps1)
    *   **Description:** This script retrieves all distribution groups from Active Directory, including their name, email address, and the number of members in each group. The data is displayed as a formatted table and exported to a CSV file for auditing purposes.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Run the script. The script will create a CSV file in `C:\Data` with the current date as a timestamp.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [UserDistributionGroupAudit.ps1](scripts/ad/Reports/UserDistributionGroupAudit.ps1)
    *   **Description:** This script retrieves a list of users and contacts from Active Directory, including their distribution group memberships. It combines user and contact objects, retrieves their properties, and filters for distribution groups. The script creates a custom object for each user or contact, including their entry type, first name, last name, email address, and group memberships. The resulting information is displayed as a formatted table and exported to a CSV file. This script is useful for auditing and documenting user and contact information, particularly their distribution group memberships, in Active Directory.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Run the script. The script will create a CSV file in `C:\Data` with the current date as a timestamp.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [UserProxyAddressCheck.ps1](scripts/ad/Reports/UserProxyAddressCheck.ps1)
    *   **Description:** This script checks all users in a specified Organizational Unit (OU) in Active Directory to verify if their SMTP addresses include any of the specified domains.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Modify the `$ou` and `$acceptedDomains` variables in the script to match your environment.
        3.  Run the script. The script will output a table of users without matching SMTP addresses and export the results to a CSV file in `C:\Export`.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [UserReport.ps1](scripts/ad/Reports/UserReport.ps1)
    *   **Description:** This script exports user information from an Active Directory domain, including the user's name, username, email address, status, and groups they belong to. It retrieves all user objects and iterates over them, collecting group membership details for each user. The script then creates custom objects for each user, adding the relevant properties, and appends them to an array. Finally, it sorts and displays the report on the screen and exports it to a CSV file. This script is useful for generating a comprehensive user report for analysis or documentation purposes.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Run the script. The script will create a CSV file in `C:\Data` with the current date as a timestamp.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [UsersNotLoggedIn.ps1](scripts/ad/Reports/UsersNotLoggedIn.ps1)
    *   **Description:** This script defines a function that retrieves user information from an Active Directory domain based on their login activity. It filters and selects enabled users who haven't logged in over the past 90 days within a specified organizational unit. The resulting user information, including their username, name, and last login date, is sorted and exported to a CSV file. This function is useful for identifying inactive user accounts within a specific organizational unit.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Modify the `$OUpath` and `$ExportPath` variables in the script to match your environment.
        3.  Run the script.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [CheckUserPasswordNeverExpire.ps1](scripts/ad/Tools/CheckUserPasswordNeverExpire.ps1)
    *   **Description:** This script searches for users whose passwords are set to never expire and exports a report to a CSV file.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Run the script. The script will create a CSV file named `pw_never_expires.csv` in `C:\Data`.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [Counts.ps1](scripts/ad/Tools/Counts.ps1)
    *   **Description:** This script provides a summary of the counts for users, groups, and computers in an Active Directory domain.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Run the script. The script will output the number of users, groups, and computers in the domain.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [GetAllSecurityGroups.ps1](scripts/ad/Tools/GetAllSecurityGroups.ps1)
    *   **Description:** This script exports the names of security groups in an Active Directory domain to a CSV file.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Modify the `$ExportPath` variable in the script to match your environment.
        3.  Run the script.
    *   **Dependencies:** Active Directory module for PowerShell.
*   [ExportUsersToExcel.ps1](scripts/ad/Reports/ExportUsersToExcel.ps1)
    *   **Description:** This script exports user information from a specified Organizational Unit (OU) in Active Directory to an Excel file. The Excel file will have separate worksheets for each OU.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Modify the `$StartingOU` variable in the script to match your environment.
        3.  Run the script. The script will create an Excel file named `UserExport.xlsx` in `C:\Users\Public`.
    *   **Dependencies:** Active Directory module for PowerShell, ImportExcel module.

### Azure

*   [ListAllAssignedLicenses.ps1](scripts/azure/Reports/ListAllAssignedLicenses.ps1)
    *   **Description:** This script installs the "AzureADPreview" module, connects to Azure Active Directory, retrieves user data including their licenses, and presents the information in a formatted table.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Run the script. The script will prompt you to log in to Azure.
    *   **Dependencies:** AzureADPreview module.
*   [ListAllTenantUsers.ps1](scripts/azure/Reports/ListAllTenantUsers.ps1)
    *   **Description:** This script installs the "AzureADPreview" module, connects to Azure Active Directory, retrieves user data including their first names, last names, emails, and presents the information in a formatted table. It also exports the data to a CSV file.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Modify the `$csvFilePath` variable in the script to match your environment.
        3.  Run the script. The script will prompt you to log in to Azure.
    *   **Dependencies:** AzureADPreview module.
*   [ListAllUsersInSecurityGroups.ps1](scripts/azure/Reports/ListAllUsersInSecurityGroups.ps1)
    *   **Description:** This script connects to Azure AD, retrieves all users, and exports their security group memberships to a CSV file. The script allows filtering by a keyword in the group name.
    *   **Usage:**
        1.  Open PowerShell as an administrator.
        2.  Modify the `$useKeyword` and `$keyword` variables in the script to match your environment.
        3.  Run the script. The script will prompt you to log in to Azure.
    *   **Dependencies:** AzureAD module.
*   [UserDistributionGroupReport.ps1](scripts/azure/Reports/UserDistributionGroupReport.ps1)
*   [Reset-SPOFolderAndFilesPermissions.ps1](scripts/azure/SharePoint/Reset-SPOFolderAndFilesPermissions.ps1)

### Exchange

*   [DeletedEmails.ps1](scripts/exchange/DeletedEmails.ps1)
*   [HideUsersFromGAL.ps1](scripts/exchange/HideUsersFromGAL.ps1)
*   [MailboxReport.py](scripts/exchange/MailboxReport.py)
    *   **Description:** A Tkinter desktop app that connects to Exchange Online, lists user and shared mailboxes with their size, quotas and Full Access / Send As / Send on Behalf permissions, and exports the report to CSV. The required `ExchangeOnlineManagement` PowerShell module is installed automatically on first run. See [its README](scripts/exchange/MailboxReport.md) to run it or build a standalone `.exe`.

### Windows

*   [ApplicationSetup.bat](scripts/windows/ApplicationSetup.bat)
*   [Dev-TechSetup.ps1](scripts/windows/Dev-TechSetup.ps1)
*   [DownloadApplicationPackage.bat](scripts/windows/DownloadApplicationPackage.bat)
*   [PersonaliseComputer.bat](scripts/windows/PersonaliseComputer.bat)

### Tools

*   [DirectoryAuditor.ps1](tools/DirectoryAuditor.ps1)
*   [GeneratePasswordGUI.ps1](tools/GeneratePasswordGUI.ps1)
*   [ListAllPrograms.bat](tools/ListAllPrograms.bat)
*   [RestartPrintSpoolerService.bat](tools/RestartPrintSpoolerService.bat)
