function SyncPrintersToIntune {
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
    SyncPrintersToIntune -AccessToken $AccessToken -Mode Create
    .EXAMPLE
    SyncPrintersToIntune -Mode Update
    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [SecureString]$AccessToken,
        [Parameter(parametersetname='Mode', Mandatory=$true)]
        [validateSet('Create','Update','FullSync')]
        [string]$Mode = 'Create',
        [Parameter(parametersetname='Mode', Mandatory=$false)]
        [switch]$Force
    )

    if (Get-Module -Name 'Microsoft.Graph.Authentication' -ListAvailable) {
        Import-Module -Name 'Microsoft.Graph.Authentication' -Force
    } else {
        Install-Module -Name 'Microsoft.Graph.Authentication' -Force
        Import-Module -Name 'Microsoft.Graph.Authentication' -Force
    }

    function AssignPolicy {
    param (
        [Parameter(Mandatory=$true)]
        [string]$groupId,
        [Parameter(Mandatory=$true)]
        [string]$intunePrintPolicyId
    )
    $assignPolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{0}/assign" -f $intunePrintPolicyId
    $assignment = @{
        "assignments" = @(
            @{
                "target" = @{
                    "@odata.type" = '#microsoft.graph.groupAssignmentTarget'
                    "groupId" = "{0}" -f $groupId
                }
            }
        )
    } | ConvertTo-Json -Depth 6
    Invoke-MgGraphRequest -Uri $assignPolicyUrl -Method POST -Body $assignment
    }

    # Set the infomraiton preference to continue, otherwise the script will not output any information
    $informationPreference = "Continue"

    # Set the client ID Graph Commandline Tools and scopes for the Microsoft Graph connection
    $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    $scopes = "Printer.Read.All PrinterShare.Read.All DeviceManagementConfiguration.ReadWrite.All"
    $universalPrintPrefix = "BL-WIN-UNIVERSAL-PRINT"
    # Connect to the Microsoft Graph with access token, otherwise interactive login will be used
    If ($null -ne $AccessToken){
        Connect-MgGraph -AccessToken $AccessToken
    } else {
        Connect-MgGraph -Scopes $Scopes -TenantId $tenantId -ClientId $clientId
    }

    # Get all printers and their shares using the $expand query parameter
    $apiUrl = 'https://graph.microsoft.com/beta/print/printers?$expand=shares'
    $results = Invoke-MgGraphRequest -Uri $apiUrl -Method GET -OutputType JSON
    $printers = $results | ConvertFrom-Json


    if ($Mode -eq 'FullSync') {
        Write-Information "Deleting all existing printer policies" -InformationAction $informationPreference
        $printerResults | ForEach-Object {
            if ($Force.IsPresent){
            $deletePolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{0}" -f $_.id
            #Invoke-MgGraphRequest -Uri $deletePolicyUrl -Method DELETE
            } else {
                $confirmation = Read-Host "Force switch not provided. Do you want to proceed with FullSync? (Yes/No)"
                if ($confirmation -ne 'Yes') {
                    Write-Output "Operation cancelled by the user."
                    return
                }
                $deletePolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/{0}" -f $_.id
                #Invoke-MgGraphRequest -Uri $deletePolicyUrl -Method DELETE
            }
        }
    }
    # Check if any printers were found
    if ($printers.value.length -gt 0) {
        Write-Information "Found $($printers.value.length) printers" -InformationAction $informationPreference
        # First of all, we need all the printer policies from Intune
        Write-Information "Getting all printer policies from Intune" -InformationAction $informationPreference

        $printerResults = [System.Collections.ArrayList]@() 
        $testPolicyUrl = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$filter=startswith(name%2c''{0}'')' -f $universalPrintPrefix
        
        # Because Graph filtering is done at the front instead of the backend we need to loop till the next link is null
        # Results are added to the printerResults array
        while ($null -ne $testPolicyUrl) {
            try {
                $policyTestResult = Invoke-MgGraphRequest -Uri $testPolicyUrl -Method GET
                $printerResults.AddRange($policyTestResult.value) >> $null
                $testPolicyUrl = $policyTestResult.'@odata.nextLink'
            } catch {
                Write-Error "Unable to get policy information from Intune. $_"
                $testPolicyUrl = $null
            }
        } 
        
        # Loop through each printer and create a policy for it if needed
            $printers.value[0] | ForEach-Object {
                if ($Mode -eq 'Create' -or $Mode -eq 'FullSync') {
                    try {
                        $printer = $_
                        # Path to the JSON file with the print policy template
                        # The JSON file should contain the policy template with tokens that will be replaced with the actual values
                        $jsonFilePath = ".\printpolicytemplate.json"
                        # Read the JSON file content as text
                        $jsonContent = Get-Content -Path $jsonFilePath -Raw

                        # Define the tokens to replace and the replacement value
                        $policyToken = "<!--policyName-->"
                        $policyName = "{0}-{1}" -f $universalPrintPrefix, $printer.displayName

                        $printerToken = "<!--printerId-->"
                        $printerId = $printer.id

                        $printerNameToken = "<!--printShareName-->"
                        $printerName = $printer.shares[0].name

                        $printerShareIdToken = "<!--printShareId-->"
                        $printerShareId = $printer.shares[0].id

                        $descriptionToken = "<!--description-->"
                        $description = "Printer configruation policy that assignes printer {0} to Entra ID group {1}.\nPrinter information:\n- Model: {2}\n- Street: {3}\n-Building: {4}\n-Room: {5}" -f $printer.displayName, $groupInfo.displayName, $printer.model, $printer.location.streetAddress, $printer.location.building, $printer.location.roomName

                        $groupIdToken = "<!--groupId-->"
                        $groupId = $printerShareId

                        # Get group information from Entra for assignment
                        $groupUrl = "https://graph.microsoft.com/beta/groups?`$filter=startswith(displayName,'UniversalPrint_{0}')&`$select=id,displayName" -f $printerShareId
                        $groupInfo = Invoke-MgGraphRequest -Uri $groupUrl -Method GET -OutputType JSON
                        $groupObject = $groupInfo | ConvertFrom-Json
                        Write-Information "Group Info found for group: $($groupObject.value.displayName)" -InformationAction $informationPreference

                                                
                        # Replace the token with the replacement value
                        $updatedJsonContent = $jsonContent -replace [regex]::Escape($policyToken), $policyName
                        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($descriptionToken), $description
                        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($printerToken), $printerId
                        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($printerNameToken), $printerName
                        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($printerShareIdToken), $printerShareId
                        $updatedJsonContent = $updatedJsonContent -replace [regex]::Escape($groupIdToken), $groupId

                        # Testing for existing policy first
                        if ($policyName -in $printerResults.Name) {
                            $intunePrintPolicy = $printerResults | Where-Object { $_.name -eq $policyName }
                            Write-Warning "Policy already exists, only assigning group"
                        } else {
                            Write-Information "Policy does not exist, creating new policy"
                            # Create the policy first
                            $createPolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
                            $intunePrintPolicy = Invoke-MgGraphRequest -Uri $createPolicyUrl -Method POST -Body $updatedJsonContent
                        }
                        # Assign the Entra ID group to the policy
                        AssignPolicy -IntunePrintPolicyId $intunePrintPolicy.id -GroupId $groupId
                    }
                    catch {
                        Write-Error "Unable to create print policy in Intune. $_"
                    }
                }
                if ($Mode -eq 'Update'){
                    try {
                        $intunePrintPolicy = $printerResults | Where-Object { $_.name -eq $policyName }
                        if ($null -eq $intunePrintPolicy) {
                            Write-Error "Policy ($($policyName)) does not exist, creating new policy"
                            Continue
                        }
                        AssignPolicy -IntunePrintPolicyId $intunePrintPolicy.id -GroupId $groupId
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
