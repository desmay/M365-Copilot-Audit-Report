# Check if MG Online module is already installed
$module = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'Microsoft.Graph' }

if ($module -eq $null) {
    try {
        Write-Host "Installing module..."
        Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser
    } 
    catch {
        Write-Host "Failed to install module: $_"
        exit
    }
}
# Connect to Microsoft Graph
try {
    Connect-mggraph -Scopes "User.Read.All" -NoWelcome
    Write-Host "Connected to Microsoft Graph."
}
catch {
    Write-Host "Failed to connect to Microsoft Graph: $_"
}

# CSV File path  
$csvUserspath = "/Users/dereksmay/code/M365-Copilot-Audit-Report/scripts/Copilot_Users.csv"


# Replace with actual Copilot SKU ID(s) from your tenant
$copilotSkuIds = "be936ece-5b91-4517-b61d-d87a525bbd9f"
$copilotSkuIds = 'b529eafb-839c-4fd3-adf5-dfe02fa5dcb2'

# Get users with manager details in a single call
$users = Get-MgUser -ConsistencyLevel eventual -CountVariable CopilotLicensedUserCount  -All -Property Id, DisplayName,  
UserPrincipalName, JobTitle, Department, City, Country, UsageLocation, AssignedLicenses, manager -ExpandProperty manager | 
Where-Object { $_.JobTitle -ne $null }

# Build enriched objects with manager info and license check (only users with a JobTitle)
$results = foreach ($user in $users | Where-Object { $_.JobTitle }) {

    $managerName = ""
    $managerUPN = ""
    $hasCopilot = $false

    # Get manager info from expanded property (no additional Graph calls)
    if ($user.Manager) {
        $managerName = $user.Manager.AdditionalProperties['displayName']
        $managerUPN  = $user.Manager.AdditionalProperties['userPrincipalName']
    }

     # Check for Copilot license
    if ($user.AssignedLicenses) {
        foreach ($lic in $user.AssignedLicenses) {
            $skuIdString = "$($lic.SkuId)"
            if ($copilotSkuIds -contains $skuIdString) {
                $hasCopilot = $true
                break
            }
        }
    }

    [PSCustomObject]@{
        EntraID           = $user.Id
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        JobTitle          = $user.JobTitle
        Department        = $user.Department
        City              = $user.City
        Country           = $user.Country
        UsageLocation     = $user.UsageLocation
        ManagerName       = $managerName
        ManagerUPN        = $managerUPN
        HasCopilotLicense = $hasCopilot
    }
}

# Export to CSV
$results | Export-Csv $csvUserspath -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to $csvUserspath"
