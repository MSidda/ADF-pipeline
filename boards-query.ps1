Param(
   [string]$OrganizationName = "siddamoolam",
   [string]$collectionurl = "https://dev.azure.com/$($OrganizationName)/",
   [string]$project = "Azure-boards",
   [string]$token = "y7jqzo72rxah5dy3thndwwsvfjzuixemjge7tta6ak4r4tdsmmya"
)

# Base64-encodes the Personal Access Token (PAT) appropriately
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user,$token)))



#$baseUrl = $collectionurl + $project + "/_apis/wit/wiql?api-version=5.0"

#$body = "
#{
#`"query`": `"SELECT [System.ID], `[System.WorkItemType]`, `[System.Title]`,
#    `[System.AssignedTo]` FROM workitems WHERE `[System.AssignedTo]` = 'sidda.moolam@imaginea.com'

#`"}"

#echo $baseUrl



#$response = Invoke-RestMethod -Uri $baseUrl -Method POST -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}  -Body $body -ContentType 'application/json'

$Url =  $collectionurl + $project + "/_apis/wit/reporting/workitemrevisions?fields=System.Id,System.Title,System.WorkItemType,System.AssignedTo&includeLatestOnly=true&api-version=5.0"

#$resTest = Invoke-RestMethod -Uri $Url -Method POST -Body $body -ContentType 'application/json' -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}.values  


$respo = (Invoke-RestMethod -Uri $Url -Method Get -UseDefaultCredential -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}).values


#$wits = $respo | where({$_.fields.'System.AssignedTo' -eq 'sidda.moolam@imaginea.com'})



$respo.fields | Export-Excel -Path '/home/siddar/Downloads/outputazboards.xlsx'







