# Get-InfoPathType.ps1

Scans all site collections across all web applications in an on-premises SharePoint farm and returns lists that use InfoPath, either as a Form Library (base template 115) or as a list with a content type that has InfoPath enabled.

## Requirements

- Windows PowerShell 5.1
- Microsoft.SharePoint.PowerShell snap-in (available on any SharePoint server with the SharePoint binaries installed)
- Run from a SharePoint server with an account that has Farm Admin or sufficient read access across all web applications

## Output

Exports a tab-delimited CSV to the path specified by `-OutPath` with the following columns:

| Column | Description |
|---|---|
| Site | URL of the web where the list resides |
| List | Title of the list or library |
| ListId | GUID of the list |
| ItemCount | Combined count of items and folders |
| LastItemModified | Date the list last had an item modified |
| DetectionType | Either `FormLibrary` or `CustomizedListForm` |

> The output file is tab-delimited. When opening in Excel, use Data > From Text/CSV and set the delimiter to Tab.

## Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| OutPath | String | Yes | Full path to the output CSV file. The file will be created if it does not exist. |
| ExcludedWebApps | String[] | No | One or more web application names to skip. Matches against the web application Name property as shown in Central Administration. |

## Usage
```powershell
# Scan all web applications
.\Get-InfoPathType.ps1 -OutPath "$env:TEMP\InfoPath.csv"

# Exclude specific web applications
.\Get-InfoPathType.ps1 -OutPath "$env:TEMP\InfoPath.csv" -ExcludedWebApps "Central Admin", "MySites"
```

> `ExcludedWebApps` values must match the web application **Name** field exactly as it appears in Central Administration, not the URL.
