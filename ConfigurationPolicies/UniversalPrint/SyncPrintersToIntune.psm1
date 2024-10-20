function Sync-PrintersToIntune {
    <#
       .SYNOPSIS
       Adds an AVD session host to the AVD Insights workbook.
       .DESCRIPTION
       The function will install the needed extensions on the AVD session host.
       .PARAMETER AccessToken
       If you have an access token, you can pass it to the function.
       .PARAMETER Mode
        Create mode will create new policies (leaves existing policies), Update mode will update existing policies (does not create), FullSync will DELETE all universal print polices and create new policies.
        .EXAMPLE
       Sync-PrintersToIntune -AccessToken $AccessToken -Mode Create
       .EXAMPLE
        Sync-PrintersToIntune -Mode Update
       .EXAMPLE
        Sync-PrintersToIntune -Mode FullSync -Force
       #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [SecureString]$AccessToken,
        [Parameter(parametersetname = 'Mode', Mandatory = $true)]
        [validateSet('Create', 'Update', 'FullSync')]
        [string]$Mode = 'Create',
        [Parameter(parametersetname = 'Mode', Mandatory = $false)]
        [switch]$Force
    )
   
    if (Get-Module -Name 'Microsoft.Graph.Authentication' -ListAvailable) {
        Import-Module -Name 'Microsoft.Graph.Authentication' -Force
    }
    else {
        Install-Module -Name 'Microsoft.Graph.Authentication' -Force
        Import-Module -Name 'Microsoft.Graph.Authentication' -Force
    }
   
    function AssignPolicy {
        param (
            [Parameter(Mandatory = $true)]
            [string]$PrinterShareId,
            [Parameter(Mandatory = $true)]
            [string]$IntunePrintPolicyId
        )
        try {
            # Get group information from Entra for assignment
            $groupUrl = "https://graph.microsoft.com/beta/groups?`$filter=startswith(displayName,'UniversalPrint_{0}')&`$select=id,displayName" -f $PrinterShareId
            $groupInfo = Invoke-MgGraphRequest -Uri $groupUrl -Method GET -OutputType JSON
            $groupObject = $groupInfo | ConvertFrom-Json
            Write-Information "Group Info found for group: $($groupObject.value.displayName)" -InformationAction $informationPreference
        }
        catch {
            Write-Error "Unable to get group information from Entra. $_"
        }
        $assignPolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{0}/assign" -f $IntunePrintPolicyId
        $groupId = $groupObject.value.id
        Write-Information "Assigning policy $($IntunePrintPolicyId) to group id $groupId" -InformationAction $informationPreference
        $assignment = @{
            "assignments" = @(
                @{
                    "target" = @{
                        "@odata.type" = '#microsoft.graph.groupAssignmentTarget'
                        "groupId"     = "{0}" -f $groupId
                    }
                }
            )
        } | ConvertTo-Json -Depth 6
        try {
            Invoke-MgGraphRequest -Uri $assignPolicyUrl -Method POST -Body $assignment
        }
        catch {
            Write-Error "Unable to assign policy to group. $_"
        }
    }
   
    function GeneratePolicyTemplate {
        param (
            [Parameter(Mandatory = $true)]
            [object]$printer,
            [Parameter(Mandatory = $true)]
            [string]$policyName
        )
        # Path to the JSON file with the print policy template
        # The JSON file should contain the policy template with tokens that will be replaced with the actual values
        $jsonString = 
        @"
        {
    "@odata.type": "#microsoft.graph.deviceManagementConfigurationPolicy",
    "name": "<!--policyName-->",
    "description": "<!--description-->",
    "platforms": "windows10",
    "technologies": "mdm",
    "roleScopeTagIds": [
        "0"
    ],
    "settings": [
        {
            "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
            "settingInstance": {
                "@odata.type": "#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance",
                "settingDefinitionId": "user_vendor_msft_printerprovisioning_upprinterinstalls_{printersharedid}",
                "groupSettingCollectionValue": [
                    {
                        "children": [
                            {
                                "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance",
                                "settingDefinitionId": "user_vendor_msft_printerprovisioning_upprinterinstalls_{printersharedid}_clouddeviceid",
                                "simpleSettingValue": {
                                    "@odata.type": "#microsoft.graph.deviceManagementConfigurationStringSettingValue",
                                    "value": "<!--printerId-->"
                                }
                            },
                            {
                                "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
                                "settingDefinitionId": "user_vendor_msft_printerprovisioning_upprinterinstalls_{printersharedid}_install",
                                "choiceSettingValue": {
                                    "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue",
                                    "value": "user_vendor_msft_printerprovisioning_upprinterinstalls_{printersharedid}_install_true",
                                    "children": []
                                }
                            },
                            {
                                "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance",
                                "settingDefinitionId": "user_vendor_msft_printerprovisioning_upprinterinstalls_{printersharedid}_printersharedname",
                                "simpleSettingValue": {
                                    "@odata.type": "#microsoft.graph.deviceManagementConfigurationStringSettingValue",
                                    "value": "<!--printShareName-->"
                                }
                            },
                            {
                                "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance",
                                "settingDefinitionId": "user_vendor_msft_printerprovisioning_upprinterinstalls_{printersharedid}_printersharedid",
                                "simpleSettingValue": {
                                    "@odata.type": "#microsoft.graph.deviceManagementConfigurationStringSettingValue",
                                    "value": "<!--printShareId-->"
                                }
                            }
                        ]
                    }
                ]
            }
        }
    ]
}
"@
   
        # Define the tokens to replace and the replacement value
        $policyToken = "<!--policyName-->"
   
        $printerToken = "<!--printerId-->"
        $printerId = $printer.id
   
        $printerNameToken = "<!--printShareName-->"
        $printerName = $printer.shares[0].name
   
        $printerShareIdToken = "<!--printShareId-->"
        $printerShareId = $printer.shares[0].id
   
        $descriptionToken = "<!--description-->"
        $description = "Printer configruation policy that assignes printer {0} to Entra ID group {1}.\nPrinter information:\n- Model: {2}\n- Street: {3}\n-Building: {4}\n-Room: {5}" -f $printer.displayName, $groupInfo.displayName, $printer.model, $printer.location.streetAddress, $printer.location.building, $printer.location.roomName
   

        # Replace the token with the replacement value
        $updatedJsonContent = $jsonString -replace [regex]::Escape($policyToken), $policyName
        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($descriptionToken), $description
        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($printerToken), $printerId
        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($printerNameToken), $printerName
        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($printerShareIdToken), $printerShareId
        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($groupIdToken), $groupId
        return $updatedJsonContent
    }
    # Set the infomraiton preference to continue, otherwise the script will not output any information
    $informationPreference = "Continue"
   
    # Set the client ID Graph Commandline Tools and scopes for the Microsoft Graph connection
    $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    $scopes = "Printer.Read.All PrinterShare.Read.All DeviceManagementConfiguration.ReadWrite.All"
    $universalPrintPrefix = "BL-WIN-UNIVERSAL-PRINT"
    $createPolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
    # Connect to the Microsoft Graph with access token, otherwise interactive login will be used
    If ($null -ne $AccessToken) {
        Connect-MgGraph -AccessToken $AccessToken
    }
    else {
        Connect-MgGraph -Scopes $Scopes -TenantId $tenantId -ClientId $clientId
    }
   
    try {
        # Get all printers and their shares using the $expand query parameter
        $apiUrl = 'https://graph.microsoft.com/beta/print/printers?$expand=shares'
        $results = Invoke-MgGraphRequest -Uri $apiUrl -Method GET -OutputType JSON
        $printers = $results | ConvertFrom-Json
        Write-Information "Found $($printers.value.length) printers" -InformationAction $informationPreference
    }
    catch {
        Write-Error "Unable to get printer information from Universal Print. $_"
    }
    # Check if any printers were found
    if ($printers.value.length -gt 0) {
        # First of all, we need all the printer policies from Intune
        Write-Information "Getting all printer policies from Intune" -InformationAction $informationPreference
   
        $printerResults = [System.Collections.ArrayList]@() 
        $fetchExistingPolicyUrl = '{0}?$filter=startswith(name%2c''{1}'')' -f $createPolicyUrl, $universalPrintPrefix
           
        # Because Graph filtering is done at the front instead of the backend we need to loop till the next link is null
        # Results are added to the printerResults array
        while ($null -ne $fetchExistingPolicyUrl) {
            try {
                $policyTestResult = Invoke-MgGraphRequest -Uri $fetchExistingPolicyUrl -Method GET
                $printerResults.AddRange($policyTestResult.value) >> $null
                $fetchExistingPolicyUrl = $policyTestResult.'@odata.nextLink'
            }
            catch {
                Write-Error "Unable to get policy information from Intune. $_"
                $fetchExistingPolicyUrl = $null
            }
        } 

        if ($Mode -eq 'FullSync') {
            try {
                Write-Information "Deleting all existing printer policies" -InformationAction $informationPreference
                $printerResults | ForEach-Object {
                    $deletePolicyUrl = "{0}/{1}" -f $createPolicyUrl, $_.id
                    if ($Force.IsPresent) {
                        Invoke-MgGraphRequest -Uri $deletePolicyUrl -Method DELETE
                    }
                    else {
                        $confirmation = Read-Host "Force switch not provided. Do you want to proceed with FullSync? (Yes/No)"
                        if ($confirmation -ne 'Yes') {
                            Write-Output "Operation cancelled by the user."
                            return
                        }
                        Invoke-MgGraphRequest -Uri $deletePolicyUrl -Method DELETE
                    }
                }
            }
            catch {
                Write-Error "Unable to delete print policies in Intune. $_"
            }
        }
        # Loop through each printer and create a policy for it if needed
        $printers.value | ForEach-Object {
            $printer = $_
            $policyName = "{0}-{1}" -f $universalPrintPrefix, $printer.displayName
            if ($Mode -eq 'Create') {
                try {
                    # Testing for existing policy first
                    if ($policyName -in $printerResults.Name) {
                        $intunePrintPolicy = $printerResults | Where-Object { $_.name -eq $policyName }
                        Write-Warning "Policy already exists, skipping"
                    }
                    else {
                        Write-Information "Policy does not exist, creating new policy"
                        # Create the policy first
                        $jsonBody = GeneratePolicyTemplate -Printer $printer -policyName $policyName
                        $intunePrintPolicy = Invoke-MgGraphRequest -Uri $createPolicyUrl -Method POST -Body $jsonBody
                        # Assign the Entra ID group to the policy
                        Write-Information "Assign policy to group id $($printer.shares[0].id)" -InformationAction $informationPreference
                        AssignPolicy -IntunePrintPolicyId $intunePrintPolicy.id -PrinterShareId $printer.shares[0].id
                    }
                }
                catch {
                    Write-Error "Unable to create print policy in Intune. $_"
                }
            }
            if ($Mode -eq 'FullSync') {
                try {
                    $jsonBody = GeneratePolicyTemplate -Printer $printer -policyName $policyName
                    # Create the policy first
                    $intunePrintPolicy = Invoke-MgGraphRequest -Uri $createPolicyUrl -Method POST -Body $jsonBody
                    Write-Information "Policy created for printer $($printer.displayName), has id $($intunePrintPolicy.id)" -InformationAction $informationPreference
                    # Assign the Entra ID group to the policy
                    Write-Information "Assign policy to group id $($printer.shares[0].id)" -InformationAction $informationPreference
                    AssignPolicy -IntunePrintPolicyId $intunePrintPolicy.id -PrinterShareId $printer.shares[0].id
                }
                catch {
                    Write-Error "Unable to create print policy in Intune. $_"
                }
            }
            if ($Mode -eq 'Update') {
                try {
                    $intunePrintPolicy = $printerResults | Where-Object { $_.name -eq $policyName }
                    if ($null -eq $intunePrintPolicy) {
                        Write-Error "Policy ($($policyName)) does not exist use Create of FullSync mode to create the policy"
                        Continue
                    }
                    $jsonBody = GeneratePolicyTemplate -Printer $printer -policyName $intunePrintPolicy.Name
                    $policyUpdateUrl = "{0}/{1}" -f $createPolicyUrl, $intunePrintPolicy.id
                    Write-Information "Updating policy $($policyUpdateUrl)" -InformationAction $informationPreference
                    Invoke-MgGraphRequest -Uri $policyUpdateUrl -Method PUT -Body $jsonBody
                    # Assign the Entra ID group to the policy
                    AssignPolicy -IntunePrintPolicyId $intunePrintPolicy.id -PrinterShareId $printer.shares[0].id
                }
                catch {
                    Write-Error "Unable to update print policy in Intune. $_"
                }
            }
        }
    } 
    else {
        Write-Warning "No printers found in Universal Print"
    }
}
Export-ModuleMember -Function Sync-PrintersToIntune