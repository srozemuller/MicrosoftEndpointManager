$GRAPH_URL = "https://graph.microsoft.com/beta"
$managedDevicesUrl = "{0}/deviceManagement/managedDevices" -f $GRAPH_URL

$mangedDevices = Invoke-GraphRequest -Method GET $managedDevicesUrl
$secondHighestVersion = $mangedDevices.value | select deviceName, osVersion  |
    Sort-Object osVersion -Descending |
    Select-Object -Unique -Skip 1 -First 1

$secondHighestVersion
