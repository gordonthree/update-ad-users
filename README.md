# update-ad-users

Powershell script to read changes from an Excel spreadsheet and create, update or remove users from on-prem Active Directory.

Requirements:
- ActiveDirectory module from RSAT or the Active Directory Management Tools role on Windows Server
- ADSync module from Azure Connect or similar Microsoft cloud sync tool package
- Import-Excel module found on the [PowerShell Gallery](https://www.powershellgallery.com/packages/ImportExcel/7.8.6)
- Run the script in an administrator powershell with domain admin rights
