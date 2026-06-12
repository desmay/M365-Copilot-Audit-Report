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
$csvUserspath = "C:\M365CopilotReport\Copilot_Users.csv"


# Replace with actual Copilot SKU ID(s) from your tenant
$copilotSkuIds = "be936ece-5b91-4517-b61d-d87a525bbd9f"

# Get users with manager details in a single call
try {
    $users = Get-MgUser -ConsistencyLevel eventual -CountVariable TenantUserCount -All -Property Id, DisplayName,  
    UserPrincipalName, JobTitle, Department, City, Country, UsageLocation, AssignedLicenses, manager -ExpandProperty manager -ErrorAction Stop
}
catch {
    Write-Host "Failed to retrieve users from Microsoft Graph. Check authentication, Graph permissions (User.Read.All), and network connectivity. Error: $($_.Exception.Message)"
    exit 1
}

# Build enriched objects with manager info and license check (only users with a JobTitle)
$usersWithJobTitle = $users | Where-Object { $_.JobTitle }
$results = foreach ($user in $usersWithJobTitle) {

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
# Fall back to local collection count if Graph count variable isn't returned.
if (-not $TenantUserCount) {
    $TenantUserCount = if ($null -ne $users) { $users.Count } else { 0 }
}
$copilotLicensedUsers = ($results | Where-Object { $_.HasCopilotLicense }).Count
Write-Host "Total users in tenant: $TenantUserCount"
Write-Host "Users with JobTitle exported: $($results.Count)"
Write-Host "Users with Copilot license: $copilotLicensedUsers"
Write-Host "Report exported to $csvUserspath"
