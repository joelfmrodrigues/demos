using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."


#Interact with query parameters or the body of the request.
$userToken = $Request.Query.UserToken
if (-not $userToken) {
    $userToken = $Request.Body.UserToken
}
if (-not $userToken) {
    throw "UserToken body param is empty or missing"
}

# hardcoded variables, for demo only...
$siteUrl = 'https://jfmr365.sharepoint.com/sites/Hubtemplate'


# connect using user access token passed from SPFx context
$userConnection = Connect-PnPOnline -Url $siteUrl -ReturnConnection -AccessToken $userToken
Write-Host "Connected to site using access token from client side"

# get items from the list that the user has access to
$items = Get-PnPListItem -List "ToDo" -Connection $userConnection
Write-Host "$($items.Count) items found"

$listArray = New-Object System.Collections.Generic.List[System.Object]
foreach ($item in $items)
{   
    Write-Host "Item title: $($item["Title"])"
    $listArray.Add([hashtable]@{
        Title=$item["Title"]; 
        }
    )
}
$body = $listArray | ConvertTo-Json


# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
