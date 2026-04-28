<#
.SYNOPSIS
    Scans SharePoint on-premises web application for InfoPath usage.

.DESCRIPTION
    Iterates through all site collections and webs across SharePoint on-premises web applications to identify lists using InfoPath forms.

    Differentiates between Form Library (BaseTemplate 115) and CustomizedListForm (content types with _ipfs_infopathenabled property. 
    Outputs a tab-delimited CSV with metadata for each match.

.PARAMETER OutPath
    Required. File path for the output CSV.

.PARAMETER ExcludedWebApps
    Optional. One or more web application display names to skip during the scan.

.EXAMPLE
    .\Get-InfoPathType.ps1 -OutPath "$env:TEMP\InfoPath.csv"
    
    Run in SharePoint Management Shell. Scans all web applications.

.EXAMPLE
    .\Get-InfoPathType.ps1 -OutPath "$env:TEMP\InfoPath.csv" -ExcludedWebApps "Central Admin", "MySites"

    Run in SharePoint Management Shell. Scans all web applications except Central Admin and MySites
    
.NOTES
    Author   :  Clouds Connected Ltd
    Requires :  SharePoint Management Shell (Microsoft.SharePoint.PowerShell snap-in)
                Run from a SharePoint server with Farm Administrator permissions.
.OUTPUTS
    CSV file with columns: Site, ListTitle, ListId, ItemCount, LastItemModified, DetectionType
#>


[CmdletBinding()]
param(
    [string]$OutPath,
    [string[]]$ExcludedWebApps = $null
) 

try{Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue}
catch[System.ArgumentException]{}


function Get-InfoPathType {
    param($list)

    if ($list.BaseTemplate -eq 115){
        return "FormLibrary"
    }

    if($list.ContentTypesEnabled -and $list.ContentTypes -ne $null -and $list.ContentTypes.Count -gt 0){
        foreach($ct in $list.ContentTypes){
            if($null -ne $ct.ResourceFolder -and $ct.ResourceFolder.Properties["_ipfs_infopathenabled"] -eq "True"){
                return "CustomizedListForm"
            }
        }
    }
    return $null
}



$webapplications = Get-SPWebApplication | Where-Object { -not $ExcludedWebApps -or $_.Name -notin $ExcludedWebApps }

foreach ($webapp in $webapplications)
{
    foreach ($site in $webapp.Sites)
    {
     try{         
        foreach ($web in $site.AllWebs)
        {
        try{
            foreach ($list in $web.Lists)
            {
            try{
                $detectionType = Get-InfoPathType -list $list
                if($detectionType){
                    Write-Host "$detectionType : $($list.Title) @ $($web.Url)"
                    [pscustomobject]@{
                        Site = $web.Url
                        ListTitle = $list.Title
                        ListId = $list.Id
                        ItemCount = $list.Items.Count+$list.Folders.Count
                        LastItemModified = $list.LastItemModifiedDate
                        DetectionType = $detectionType
                    } | Export-Csv -Path $OutPath -Append -NoTypeInformation -Delimiter "`t" -Encoding UTF8
                }
                }
                catch{Write-Host "Error processing list '$($list.Title)' in $($web.Url) : $_"}
            }
            }
        catch{Write-Host "Error processing web '$($web.Url)' : $_"}
        finally{$web.Dispose()}
        }
    }
    catch{Write-Host "Error processing site '$($site.Url)' : $_"}
    finally{$site.Dispose()}
    }
}


Write-Host "Output file in $($OutPath)"
