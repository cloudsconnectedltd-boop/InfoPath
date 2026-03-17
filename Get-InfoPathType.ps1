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


$OutPath = "C:\SiteEnumeration\resultsInfoPath.csv" #Add your output path for the result csv

$webapplications = Get-SPWebApplication | Where-Object { $_.Name -notin ("Add your excluded Web Apps here") }

foreach ($webapp in $webapplications)
{
    foreach ($site in $webapp.Sites)
    {
        foreach ($web in $site.AllWebs)
        {
            foreach ($list in $web.Lists)
            {
                $detectionType = Get-InfoPathType -list $list

                if($detectionType){
                    Write-Host "$detectionType : $($list.Title) @ $($web.Url)"
                    [pscustomobject]@{
                        Site = $web.Url
                        List = $list.Title
                        ListId = $list.Id
                        ItemCount = $($list.Items.Count)+$($list.Folders.Count)
                        LastItemModified = $list.LastItemModifiedDate
                        DetectionType = $detectionType
                    } | Export-Csv -Path $OutPath -Append -NoTypeInformation -Delimiter "`t"
                }
            }
            $web.Dispose()
        }
        $site.Dispose()
    }
}
