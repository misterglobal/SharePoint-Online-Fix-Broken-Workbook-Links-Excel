# SharePoint-Online-Fix-Broken-Workbook-Links-Excel
This short script will repair broken workbook links in SharePoint Online document libary

```markdown
# SharePoint Excel Link Updater

This PowerShell script connects to a SharePoint Online site, retrieves all Excel files from a specified document library, updates broken internal links in those files, and re-uploads them to the site. The script is designed to be run locally and requires user authentication via a browser window.

## Prerequisites

Before running this script, ensure that the following PowerShell modules are installed:

- **PnP.PowerShell**: This module provides the necessary cmdlets to interact with SharePoint Online.
- **ImportExcel**: This module is used for opening and manipulating Excel files.

To install these modules, use the following commands:

```powershell
Install-Module -Name PnP.PowerShell
Install-Module -Name ImportExcel
```

## Script Overview

1. **SharePoint Connection**: 
   - The script connects to a specified SharePoint site using the `Connect-PnPOnline` cmdlet. It uses the `-UseWebLogin` parameter to open a browser window for user authentication.
   
2. **File Retrieval**: 
   - The script retrieves all Excel files (`.xlsx` files) from the "Documents" library of the SharePoint site.
   
3. **Link Updating**: 
   - It checks for any broken internal links in the Excel files that point to non-HTTP(S) paths and attempts to resolve them by finding the corresponding file in the SharePoint site.
   
4. **File Re-upload**: 
   - After updating the links, the Excel files are saved and re-uploaded to their original location in SharePoint. The temporary files are then deleted from the local system.

5. **Cleanup**: 
   - The script disconnects from SharePoint and cleans up any temporary files.

## Usage

1. Open PowerShell and navigate to the directory where the script is located.
   
2. Run the script by typing the following command:

   ```powershell
   .\Update-ExcelLinks.ps1
   ```

3. A browser window will open, prompting you to authenticate your SharePoint account.

4. The script will connect to the specified SharePoint site, process the Excel files, and update any internal links.

5. Once the script completes, it will disconnect from SharePoint.

## Configuration

Update the following variables in the script to match your environment:

- `$siteUrl`: The URL of your SharePoint site.
- The document library name (`"Documents"`) in the `Get-PnPListItem` and `Get-PnPFile` cmdlets should be updated if you are using a different library name.

## Error Handling

- If there are any issues connecting to SharePoint, the script will exit and display an error message.
- If the script cannot find a corresponding file for a broken link in SharePoint, it will notify you but will not terminate execution.

## License

This project is licensed under the MIT License. Feel free to use, modify, and distribute it.

## Contributions

Contributions and improvements are welcome! Please submit a pull request or open an issue for any suggestions or bugs.
```

This README file provides instructions on how to install the necessary modules, configure the script, and run it. It also includes a brief overview of what the script does, error handling, and licensing information. Let me know if you'd like any additional changes!
