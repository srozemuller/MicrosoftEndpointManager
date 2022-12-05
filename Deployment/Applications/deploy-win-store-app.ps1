
$appId = ""
$appSecret = ""
$tenantId = ""

$body = @{    
    grant_Type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_Id     = $appId
    client_Secret = $appSecret 
} 
$connectParams = @{
    uri    = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token" -f $tenantId
    method = "POST" 
    body   = $body
}
$connect = Invoke-RestMethod @connectParams
$authHeader = @{
    'Content-Type' = 'application/json'
    Authorization  = 'Bearer ' + $connect.access_token
}

$appName = "WhatsApp"
$storeSearchUrl = "https://storeedgefd.dsx.mp.microsoft.com/v9.0/manifestSearch"
$body = @{
    Query = @{
        KeyWord   = $appName
        MatchType = "Substring"
    }
} | ConvertTo-Json
$appSearch = Invoke-RestMethod -Uri $storeSearchUrl -Method POST -ContentType 'application/json' -body $body
$exactApp = $appSearch.Data | Where-Object { $_.PackageName -eq $appName }

$appUrl = "https://storeedgefd.dsx.mp.microsoft.com/v9.0/packageManifests/{0}" -f $exactApp.PackageIdentifier
$app = Invoke-RestMethod -Uri $appUrl -Method GET 
$appId = $app.Data.PackageIdentifier
$appInfo = $app.Data.Versions[-1].DefaultLocale
$appInstaller = $app.Data.Versions[-1].Installers


$imageUrl = "https://apps.microsoft.com/store/api/ProductsDetails/GetProductDetailsById/{0}?hl=en-US&gl=US" -f $exactApp.PackageIdentifier
$test = Invoke-RestMethod -Uri $imageUrl -Method GET 
$test
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($test.IconUrl, "./est.jpg")
$base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes('./est.jpg'))


$deployUrl = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps"
$appBody = @{
    '@odata.type'         = "#microsoft.graph.winGetApp"
    description           = $appInfo.ShortDescription
    developer             = $appInfo.Publisher
    displayName           = $appInfo.packageName
    informationUrl        = $appInfo.PublisherSupportUrl
    largeIcon             = @{
        "@odata.type"= "#microsoft.graph.mimeContent"
        "type" ="String"
        "value" = $base64string 
    }
    installExperience     = @{
        runAsAccount = $appInstaller.scope
    }
    isFeatured            = $false
    packageIdentifier     = $appId
    privacyInformationUrl = $appInfo.PrivacyUrl
    publisher             = $appInfo.publisher
    repositoryType        = "microsoftStore"
    roleScopeTagIds       = @()
} | ConvertTo-Json 
$appDeploy = Invoke-RestMethod -uri $deployUrl -Method POST -Headers $authHeader -Body $appBody


$assignUrl = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{0}/assign" -f $appDeploy.Id
$assignBody = 
@{
    mobileAppAssignments = @(
        @{
            "@odata.type" = "#microsoft.graph.mobileAppAssignment"
            target        = @{
                "@odata.type" = "#microsoft.graph.allDevicesAssignmentTarget"
            }
            intent        = "Required"
            settings      = @{
                "@odata.type"       = "#microsoft.graph.winGetAppAssignmentSettings"
                notifications       = "showAll"
                installTimeSettings = $null
                restartSettings     = $null
            }
        }
    )
} | ConvertTo-Json -Depth 8
Invoke-RestMethod -uri $assignUrl -Method POST -Headers $authHeader -ContentType 'application/json' -body $assignBody