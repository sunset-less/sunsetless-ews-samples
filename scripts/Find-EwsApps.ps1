<#PSScriptInfo

.VERSION 1.0.0

.GUID 93bd44b9-a886-4025-a1c1-49e59102898d

.AUTHOR Sunsetless

.COMPANYNAME Sunsetless

.COPYRIGHT (c) Sunsetless. MIT License.

.TAGS EWS ExchangeOnline Exchange EntraID MicrosoftGraph EwsAllowedAppIDs

.LICENSEURI https://github.com/sunset-less/sunsetless-ews-samples/blob/main/LICENSE

.PROJECTURI https://sunsetless.com/guides/find-apps-using-ews

.RELEASENOTES
First release.

#>

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns, ExchangeOnlineManagement

<#
.SYNOPSIS
Lists the applications that use EWS in Exchange Online or hold access to it.

.DESCRIPTION
Collects application IDs from four places: the EWS usage report (optional), the EwsAllowedAppIDs list, EWS
permissions in Microsoft Entra ID (full_access_as_app and EWS.AccessAsUser.All) and RBAC for Applications in
Exchange Online. For each application it adds the name, whether it is Microsoft's, a vendor's or your own, the
verified publisher and the owners. The result goes to the screen and to ews-apps.csv in the current folder.

The script changes nothing in the tenant. It signs in to Microsoft Graph with Directory.Read.All and to Exchange Online PowerShell.

.PARAMETER UsageReport
The CSV exported from the EWS usage report in the Microsoft 365 admin center. Leave it out where the report is
not available.

.EXAMPLE
Find-EwsApps.ps1

.EXAMPLE
Find-EwsApps.ps1 -UsageReport .\EWSUsage.csv

.LINK
https://sunsetless.com/guides/find-apps-using-ews
#>
param(
    # The CSV exported from the EWS usage report. Leave it out where the report is not available.
    [string]$UsageReport
)
$ErrorActionPreference = "Stop"

Connect-MgGraph -Scopes "Directory.Read.All" -NoWelcome
Connect-ExchangeOnline -ShowBanner:$false

$exchangeOnlineAppId = "00000002-0000-0ff1-ce00-000000000000"
$microsoftGraphAppId = "00000003-0000-0000-c000-000000000000"
$microsoftTenantIds = "f8cdef31-a31e-4b4a-93e4-5f571e91255a", "72f988bf-86f1-41af-91ab-2d7cd011db47"
$tenantId = (Get-MgContext).TenantId
$guid = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}"
$apps = @{}

function Add-App($AppId, $FoundIn) {
    $id = "$AppId".Trim().ToLower()
    if (-not $apps[$id]) {
        $apps[$id] = [pscustomobject]@{
            AppId = $id; Name = ""; Kind = ""; Publisher = ""; Owners = ""
            FoundIn = @(); Calls = 0; SoapActions = @()
        }
    }
    if ($apps[$id].FoundIn -notcontains $FoundIn) { $apps[$id].FoundIn += $FoundIn }
    $apps[$id]
}

# The usage report
if ($UsageReport) {
    $report = Import-Csv $UsageReport
    $columns = $report[0].PSObject.Properties.Name
    $idColumn = $columns -match "app.*id" | Select-Object -First 1
    $actionColumn = $columns -match "action" | Select-Object -First 1
    $callsColumn = $columns -match "call" | Select-Object -First 1
    foreach ($row in $report) {
        if ($row.$idColumn -notmatch $guid) { continue }
        $app = Add-App $row.$idColumn "usage report"
        $app.Calls += [long]($row.$callsColumn -replace "\D", "")
        $action = $row.$actionColumn -replace ".*/", ""
        if ($app.SoapActions -notcontains $action) { $app.SoapActions += $action }
    }
}

# The allow list
$config = Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy
"EwsEnabled: $($config.EwsEnabled)"
foreach ($id in $config.EwsAllowedAppIDs -split ",") {
    if ($id -match $guid) { $null = Add-App $id "allow list" }
}

# EWS permissions in Microsoft Entra ID
$exchange = Get-MgServicePrincipal -Filter "appId eq '$exchangeOnlineAppId'"
$graph = Get-MgServicePrincipal -Filter "appId eq '$microsoftGraphAppId'"
$fullAccess = ($exchange.AppRoles | Where-Object Value -eq "full_access_as_app").Id

foreach ($assignment in Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $exchange.Id -All) {
    if ($assignment.AppRoleId -eq $fullAccess) {
        $null = Add-App (Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId).AppId "full_access_as_app"
    }
}
foreach ($resource in $exchange, $graph) {
    foreach ($grant in Get-MgOauth2PermissionGrant -Filter "resourceId eq '$($resource.Id)'" -All) {
        if (($grant.Scope -split " ") -contains "EWS.AccessAsUser.All") {
            $null = Add-App (Get-MgServicePrincipal -ServicePrincipalId $grant.ClientId).AppId "EWS.AccessAsUser.All"
        }
    }
}

# RBAC for Applications in Exchange Online
foreach ($principal in Get-ServicePrincipal) {
    $roles = Get-ManagementRoleAssignment -RoleAssignee $principal.Identity | Where-Object Role -like "*EWS*"
    if ($roles) { $null = Add-App $principal.AppId "Exchange RBAC" }
}

# A name, a kind and an owner for each application
foreach ($app in $apps.Values) {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'"
    if (-not $sp) { $app.Kind = "Not in your directory"; continue }
    $app.Name = $sp.DisplayName
    if ($sp.ServicePrincipalType -eq "ManagedIdentity" -or $sp.AppOwnerOrganizationId -eq $tenantId) {
        $app.Kind = "Your own"
        $registration = Get-MgApplication -Filter "appId eq '$($app.AppId)'"
        if ($registration) {
            $app.Owners = (Get-MgApplicationOwnerAsUser -ApplicationId $registration.Id).UserPrincipalName -join "; "
        }
    } elseif ($microsoftTenantIds -contains $sp.AppOwnerOrganizationId) {
        $app.Kind = "Microsoft"
    } else {
        $app.Kind = "Vendor"
        $app.Publisher = $sp.VerifiedPublisher.DisplayName
    }
}

$result = $apps.Values | Sort-Object Kind, Name | Select-Object AppId, Name, Kind, Publisher, Owners,
    @{ Name = "FoundIn"; Expression = { $_.FoundIn -join "; " } }, Calls,
    @{ Name = "SoapActions"; Expression = { $_.SoapActions -join " " } }
$result | Export-Csv .\ews-apps.csv -NoTypeInformation -Encoding UTF8
$result | Format-Table AppId, Name, Kind, FoundIn -AutoSize
