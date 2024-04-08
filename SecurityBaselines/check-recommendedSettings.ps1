$clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e" # Microsoft Graph Powershell enterprise application
$token = Get-MsalToken -ClientId $clientId -Scopes "https://graph.microsoft.com/DeviceManagementConfiguration.Read.All"
$header = @{
    Authorization = "Bearer {0}" -f $token.AccessToken
}

# Search for all Edge security baselines
$currentConfiguredBaselinesUrl = "https://graph.microsoft.com/beta//deviceManagement/configurationPolicies?`$filter=(templateReference/TemplateId%20eq%20%27c66347b7-8325-4954-a235-3bf2233dfbfd_1%27%20or%20templateReference/TemplateId%20eq%20%27c66347b7-8325-4954-a235-3bf2233dfbfd_2%27)%20and%20(templateReference/TemplateFamily%20eq%20%27Baseline%27)"
$getCurrentBaselines = Invoke-RestMethod -Uri $currentConfiguredBaselinesUrl -Headers $header -Method Get

# Then for each baseline, get the settings and check the insights against the insights API.
# Settings are needed to get nice names for the settings.
$getCurrentBaselines.value.ForEach
({
    $baseline = $_
    # Get current baseline information
    $baselineSettingsUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('{0}')/settings?`$expand=settingDefinitions&top=1000" -f $baseline.Id
    $getBaselineSettings = Invoke-RestMethod -Uri $baselineSettingsUrl -Headers $header -Method Get

    # Get setting insights for the baseline
    $getSettingInsightsUrl = "https://graph.microsoft.com/beta/deviceManagement/templateInsights('{0}')/settingInsights" -f $baseline.templateReference.templateid
    $getSettingInsights = Invoke-RestMethod -Uri $getSettingInsightsUrl -Headers $header -Method Get

    # When there are insights available, loop through them and compare the current setting with the recommended setting.
    if ($getSettingInsights.count -gt 0) {
        $getSettingInsights.value.ForEach({
                $getSettingInsightsId = $_.settingDefinitionId
                $recommendedSettingId = $_.settingInsight.value
                # Get Setting information from the definition based on the insights setting definition Id
                $currentBaselineDefinition = $getBaselineSettings.value.settingDefinitions.Where({ $_.id -eq $getSettingInsightsId })
                # Search for the current setting value in the policy
                $currentBaselineSetting = ($getBaselineSettings.value.settingInstance.Where({ $_.settingDefinitionId -eq $getSettingInsightsId })).choiceSettingValue.value

                # Make the value human readable by searching in the setting definition for the specific setting Id
                $currentBaselineSettingReadable = $currentBaselineDefinition.options.Where({ $_.itemId -eq $currentBaselineSetting }).displayName
                # At last, search in the same definition for the setting value that Microsoft suggests.
                $shouldBeValue = $currentBaselineDefinition.options.Where({ $_.itemId -eq $recommendedSettingId }).displayName

                # Send message to the user based on the findings.
                if ($currentBaselineSettingReadable -ne $shouldBeValue) {
                    Write-Warning "Baseline: $($baseline.Name) has setting: $($currentBaselineDefinition.displayName) with value: $($currentBaselineSettingReadable) but should be: $shouldBeValue"
                }
                else {
                    Write-Host "Baseline: $($baseline.Name) has setting: $($currentBaselineDefinition.displayName) with value: $($currentBaselineSettingReadable) and is correct."
                }
            })
    }
    else {
        Write-Host "Baseline: $($baseline.Name) has no insights."
    }
})

