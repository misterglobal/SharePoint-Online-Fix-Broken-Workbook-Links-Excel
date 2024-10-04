# Requires installation of PnP.PowerShell and ImportExcel modules
# Install-Module -Name PnP.PowerShell
# Install-Module -Name ImportExcel

# SharePoint site URL
$siteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"

# Connect to SharePoint using WebLogin
try {
    Write-Host "Connecting to SharePoint. A browser window will open for authentication." -ForegroundColor Yellow
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Write-Host "Successfully connected to SharePoint." -ForegroundColor Green
}
catch {
    Write-Host "Error connecting to SharePoint: $_" -ForegroundColor Red
    exit
}

function Get-SharePointFileUrl($fileName) {
    $results = Get-PnPListItem -List "Documents" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$fileName</Value></Eq></Where></Query></View>"
    if ($results) {
        return $results[0].FieldValues.FileRef
    }
    return $null
}

function Update-ExcelLinks($filePath) {
    $excel = Open-ExcelPackage -Path $filePath
    $workbook = $excel.Workbook
    $linksUpdated = $false

    foreach ($worksheet in $workbook.Worksheets) {
        $hyperlinks = $worksheet.Hyperlinks
        foreach ($link in $hyperlinks) {
            if ($link.AbsoluteUri -and -not $link.AbsoluteUri.StartsWith("http")) {
                $linkedFileName = Split-Path $link.AbsoluteUri -Leaf
                $newUrl = Get-SharePointFileUrl $linkedFileName
                if ($newUrl) {
                    $link.AbsoluteUri = $newUrl
                    $linksUpdated = $true
                    Write-Host "Updated link to $linkedFileName in $($worksheet.Name)"
                }
                else {
                    Write-Host "Could not find $linkedFileName in SharePoint" -ForegroundColor Yellow
                }
            }
        }
    }

    if ($linksUpdated) {
        Close-ExcelPackage $excel -Save
        Write-Host "Updated and saved $filePath" -ForegroundColor Green
    }
    else {
        Close-ExcelPackage $excel
        Write-Host "No updates needed for $filePath" -ForegroundColor Cyan
    }
}

# Get all Excel files in the SharePoint site
$excelFiles = Get-PnPListItem -List "Documents" -Query "<View><Query><Where><Eq><FieldRef Name='File_x0020_Type'/><Value Type='Text'>xlsx</Value></Eq></Where></Query></View>"

foreach ($file in $excelFiles) {
    $filePath = Join-Path $env:TEMP $file.FieldValues.FileLeafRef
    Get-PnPFile -Url $file.FieldValues.FileRef -Path $env:TEMP -Filename $file.FieldValues.FileLeafRef -AsFile
    Update-ExcelLinks $filePath
    Add-PnPFile -Path $filePath -Folder $file.FieldValues.FileDirRef -Overwrite
    Remove-Item $filePath
}

Disconnect-PnPOnline
Write-Host "Disconnected from SharePoint. Script execution completed." -ForegroundColor Green
