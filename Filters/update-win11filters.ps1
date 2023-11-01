[CmdletBinding()]
param (
    [Parameter()]
    [String]$ApplicationId,
    [Parameter()]
    [String]$ApplicationSecret,
    [Parameter()]
    [String]$TenantId
)

$body = @{
    grant_Type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_Id     = $ApplicationId
    client_Secret = $ApplicationSecret
}

$connectParams = @{
    uri    = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token" -f $TenantId
    method = "POST"
    body   = $body
}
$connect = Invoke-RestMethod @connectParams
$authHeader = @{
    'Content-Type' = 'application/json'
    Authorization  = 'Bearer ' + $connect.access_token
}

$filterUrl = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?`$filter=platform eq 'windows10AndLater'"
$filterReqParams = @{
    uri     = $filterUrl
    method  = "GET"
    headers = $authHeader
}
$filterResults = Invoke-RestMethod @filterReqParams

$filterResults.value.ForEach({
    $filterInfo = $_
    $foundMatches = Select-String -InputObject $filterInfo.rule  -Pattern "\b10.0.22000\b|\b10.0.22621\b" -AllMatches
    $foundMatches.Matches.Count
    $createNewFilter = $false
    $foundNewBuildAlready = Select-String -InputObject $filterInfo.rule  -Pattern "\b10.0.22631\b" -AllMatches
    switch ($foundMatches.Matches.Count) {
        0 {
            Write-Host "No matches found or allready added for filter $($filterInfo.displayName), do nothing."
        }
        1 {
            "One match found for filter $($filterInfo.displayName), it seems this is a specific Windows 11 filter, if there is no Windows 23H2 filter already, I create one later"
            if ($foundNewBuildAlready.Matches.Count -eq 0) {
                $createNewFilter = $true
            }
        }
        2 {
            if ($null -eq $foundNewBuildAlready) {
            Write-Host "Two matches found for filter $($filterInfo.displayName), it seems this is a full Windows 11 filter, adding new number to the filter"
            $newRule = "{0} -or (device.osVersion -contains `"10.0.22631`")" -f $filterInfo.rule
            $filterUrl = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters('{0}')" -f $filterInfo.id
            $filterBody = @{
                rule = $newRule
            } | ConvertTo-Json
            $filterReqParams = @{
                uri     = $filterUrl
                method  = "PATCH"
                headers = $authHeader
                body = $filterBody
            }
            Invoke-RestMethod @filterReqParams
        } else {
            Write-Host "Found new build number filter $($filterInfo.displayName) already, do nothing."
        }
        }
    }
    if ($createNewFilter) {
        Write-Host "No filter found with new Windows 11 build number found, creating one for the new build 23H2"
        $newFilterUrl = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters"
        $filterBody = @{
            displayName = "Windows 11 23H2 filter"
            description = "Filter for Windows 11 23H2"
            platform    = "windows10AndLater"
            rule = "(device.osVersion -contains `"10.0.22631`")"
        } | ConvertTo-Json
        $newFilterReqParams = @{
            uri     = $newFilterUrl
            method  = "POST"
            headers = $authHeader
            body = $filterBody
        }
        Invoke-RestMethod @newFilterReqParams
    }
})