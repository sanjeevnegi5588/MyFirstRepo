<#
.SYNOPSIS
    Exports every enterprise application (service principal) that holds critical or high-risk
    Microsoft Graph APPLICATION permissions to a single CSV: owners, permissions, the auth methods
    it uses with the age and expiry of each credential, and its last sign-in date.

.DESCRIPTION
    One row per enterprise app with these columns:

        RiskRating           Critical = critical permission + active client secret
                             High     = critical permission, or high permission + client secret
                             Medium   = high permission only
        AppName, AppId
        Owners               UPNs (or display names) of the app registration / enterprise app owners
        CriticalPermissions  Critical Graph application permissions held
        HighPermissions      High-risk Graph application permissions held
        AuthMethods          ClientSecret, Certificate, FederatedCredential, ManagedIdentity, ...
        CredentialAges       Every credential with its age and expiry, e.g.
                             "ClientSecret 'prod-2024': 742 days old, expires 2026-12-01 (17 days left) | Certificate ..."
        LastSignIn           Last service principal sign-in (needs Entra ID P1/P2; blank if unavailable)

    Read-only Graph permissions used: Application.Read.All, Directory.Read.All, AuditLog.Read.All.

.PARAMETER OutputPath
    CSV file to write. Defaults to .\CriticalGraphApps_<timestamp>.csv

.PARAMETER IncludeNonCritical
    Export every app with any Graph application permission, not only critical/high ones.

.PARAMETER AdditionalCriticalPermissions
    Extra Graph permission names to treat as Critical, e.g. 'AuditLog.Read.All'.

.PARAMETER TenantId
    Tenant to connect to. Optional.

.PARAMETER SkipConnect
    Use an existing Connect-MgGraph session (e.g. app-only with a certificate).

.EXAMPLE
    .\Export-CriticalGraphApps.ps1

.EXAMPLE
    Connect-MgGraph -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumb> -NoWelcome
    .\Export-CriticalGraphApps.ps1 -SkipConnect -OutputPath C:\Reports\GraphApps.csv

.NOTES
    Requires PowerShell 7+ and: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("CriticalGraphApps_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),
    [switch]$IncludeNonCritical,
    [string[]]$AdditionalCriticalPermissions = @(),
    [string]$TenantId,
    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Permission catalogue ---------------------------------------------------------------------
# Critical = can lead to tenant takeover / privilege escalation.  High = broad access to tenant data.
# Edit to match your organisation's risk model.
$PermissionTiers = @{
    # --- Critical ---
    'RoleManagement.ReadWrite.Directory'                      = 'Critical'
    'RoleAssignmentSchedule.ReadWrite.Directory'              = 'Critical'
    'RoleEligibilitySchedule.ReadWrite.Directory'             = 'Critical'
    'RoleManagementPolicy.ReadWrite.Directory'                = 'Critical'
    'PrivilegedAccess.ReadWrite.AzureAD'                      = 'Critical'
    'PrivilegedAccess.ReadWrite.AzureADGroup'                 = 'Critical'
    'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup'     = 'Critical'
    'PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup'    = 'Critical'
    'AppRoleAssignment.ReadWrite.All'                         = 'Critical'
    'Application.ReadWrite.All'                               = 'Critical'
    'Directory.ReadWrite.All'                                 = 'Critical'
    'Domain.ReadWrite.All'                                    = 'Critical'
    'Organization.ReadWrite.All'                              = 'Critical'
    'IdentityProvider.ReadWrite.All'                          = 'Critical'
    'Policy.ReadWrite.ConditionalAccess'                      = 'Critical'
    'Policy.ReadWrite.AuthenticationMethod'                   = 'Critical'
    'Policy.ReadWrite.PermissionGrant'                        = 'Critical'
    'Policy.ReadWrite.CrossTenantAccess'                      = 'Critical'
    'Policy.ReadWrite.ApplicationConfiguration'               = 'Critical'
    'UserAuthenticationMethod.ReadWrite.All'                  = 'Critical'
    'User-PasswordProfile.ReadWrite.All'                      = 'Critical'
    'User.ReadWrite.All'                                      = 'Critical'
    'Group.ReadWrite.All'                                     = 'Critical'
    'GroupMember.ReadWrite.All'                               = 'Critical'
    'ServicePrincipalEndpoint.ReadWrite.All'                  = 'Critical'
    'EntitlementManagement.ReadWrite.All'                     = 'Critical'
    'DeviceManagementRBAC.ReadWrite.All'                      = 'Critical'
    'DeviceManagementConfiguration.ReadWrite.All'             = 'Critical'
    'DeviceManagementManagedDevices.ReadWrite.All'            = 'Critical'
    'DeviceManagementManagedDevices.PrivilegedOperations.All' = 'Critical'
    'DeviceManagementApps.ReadWrite.All'                      = 'Critical'

    # --- High ---
    'Application.ReadWrite.OwnedBy'                           = 'High'
    'User.DeleteRestore.All'                                  = 'High'
    'User.EnableDisableAccount.All'                           = 'High'
    'User.ManageIdentities.All'                               = 'High'
    'Mail.ReadWrite'                                          = 'High'
    'Mail.Read'                                               = 'High'
    'Mail.Send'                                               = 'High'
    'MailboxSettings.ReadWrite'                               = 'High'
    'Calendars.ReadWrite'                                     = 'High'
    'Contacts.ReadWrite'                                      = 'High'
    'Files.ReadWrite.All'                                     = 'High'
    'Files.Read.All'                                          = 'High'
    'Sites.FullControl.All'                                   = 'High'
    'Sites.Manage.All'                                        = 'High'
    'Sites.ReadWrite.All'                                     = 'High'
    'Sites.Read.All'                                          = 'High'
    'Notes.ReadWrite.All'                                     = 'High'
    'Chat.Read.All'                                           = 'High'
    'Chat.ReadWrite.All'                                      = 'High'
    'ChannelMessage.Read.All'                                 = 'High'
    'TeamMember.ReadWrite.All'                                = 'High'
    'eDiscovery.ReadWrite.All'                                = 'High'
    'SecurityEvents.ReadWrite.All'                            = 'High'
    'IdentityRiskyUser.ReadWrite.All'                         = 'High'
}
foreach ($p in $AdditionalCriticalPermissions) { if ($p) { $PermissionTiers[$p.Trim()] = 'Critical' } }
#endregion

#region Helpers ----------------------------------------------------------------------------------
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$Now        = [datetime]::UtcNow

function Invoke-Graph {
    param([Parameter(Mandatory)][string]$Uri)
    # The Microsoft Graph SDK retries 429/503 with back-off on its own.
    Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
}

function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    while ($next) {
        $resp = Invoke-Graph -Uri $next
        if ($resp.PSObject.Properties['value'] -and $resp.value) {
            foreach ($i in $resp.value) { $items.Add($i) }
        }
        $next = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    }
    return , $items
}

function Get-Prop {
    # Safe property read under StrictMode
    param($Object, [string]$Name)
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $null
}

function ConvertTo-UtcDate {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        return $Value.ToUniversalTime()
    }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime }
    return [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
}

function Format-Date { param($d) if ($d) { $d.ToString('yyyy-MM-dd') } else { '' } }

function Test-NotFound {
    param($ErrorRecord)
    $msg = "$($ErrorRecord.Exception.Message) $($ErrorRecord.ErrorDetails)"
    return ($msg -match 'NotFound|404|Request_ResourceNotFound|does not exist')
}

function Format-Credential {
    # Builds e.g. "ClientSecret 'prod-2024': 742 days old, expires 2026-12-01 (17 days left)"
    param([string]$Type, [string]$Name, $Start, $End, [string]$Source)
    $s = ConvertTo-UtcDate $Start
    $e = ConvertTo-UtcDate $End
    $label = if ($Name) { "$Type '$Name'" } else { $Type }
    if ($Source -eq 'ServicePrincipal') { $label += ' [on enterprise app]' }
    $age = if ($s) { '{0} days old' -f [int][math]::Floor(($Now - $s).TotalDays) } else { 'age unknown' }
    $exp = if (-not $e) { 'no expiry' }
           elseif ($e -lt $Now) { 'EXPIRED {0}' -f (Format-Date $e) }
           else { 'expires {0} ({1} days left)' -f (Format-Date $e), [int][math]::Floor(($e - $Now).TotalDays) }
    [pscustomobject]@{
        Text    = "${label}: $age, $exp"
        Active  = -not ($e -and $e -lt $Now)
        AgeDays = if ($s) { ($Now - $s).TotalDays } else { -1 }
    }
}
#endregion

#region Connect ----------------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Module 'Microsoft.Graph.Authentication' not found. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication

if (-not $SkipConnect) {
    $connectParams = @{ Scopes = @('Application.Read.All', 'Directory.Read.All', 'AuditLog.Read.All'); NoWelcome = $true }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams
}
$ctx = Get-MgContext
if (-not $ctx) { throw 'Not connected to Microsoft Graph. Run Connect-MgGraph or drop -SkipConnect.' }
Write-Host "Connected to tenant $($ctx.TenantId)" -ForegroundColor Cyan
#endregion

#region 1. Application permissions granted on Microsoft Graph ------------------------------------
Write-Host 'Reading Microsoft Graph application permissions...' -ForegroundColor Cyan
$graphSp = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$GraphAppId')?`$select=id,appRoles"
$appRoleById = @{}
foreach ($r in $graphSp.appRoles) { $appRoleById[[string]$r.id] = $r.value }

$permsBySp = @{}   # spId -> HashSet of permission names
$assignments = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($graphSp.id)/appRoleAssignedTo?`$top=999"
foreach ($a in $assignments) {
    if ($a.principalType -ne 'ServicePrincipal') { continue }
    $name = $appRoleById[[string]$a.appRoleId]
    if (-not $name) { continue }
    if (-not $permsBySp.ContainsKey($a.principalId)) { $permsBySp[$a.principalId] = [System.Collections.Generic.HashSet[string]]::new() }
    [void]$permsBySp[$a.principalId].Add($name)
}

$targetSpIds = @($permsBySp.Keys | Where-Object {
    $IncludeNonCritical -or ($permsBySp[$_] | Where-Object { $PermissionTiers.ContainsKey($_) })
})
Write-Host ("{0} enterprise apps with critical/high Graph application permissions" -f $targetSpIds.Count) -ForegroundColor Cyan
#endregion

#region 2. Last sign-in dates (needs Entra ID P1/P2) ---------------------------------------------
$signInByAppId = @{}
$signInAvailable = $true
try {
    Write-Host 'Reading service principal sign-in activity...' -ForegroundColor Cyan
    foreach ($s in (Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities')) {
        $signInByAppId[[string]$s.appId] = $s
    }
} catch {
    $signInAvailable = $false
    Write-Warning "Last sign-in dates unavailable (needs AuditLog.Read.All and Entra ID P1/P2): $($_.Exception.Message)"
}
#endregion

#region 3. Inspect each app ----------------------------------------------------------------------
$rows = [System.Collections.Generic.List[object]]::new()
$spSelect  = 'id,appId,displayName,servicePrincipalType,passwordCredentials,keyCredentials'
$appSelect = 'id,appId,passwordCredentials,keyCredentials'

$i = 0
foreach ($spId in $targetSpIds) {
    $i++
    Write-Progress -Activity 'Inspecting enterprise apps' -Status "$i of $($targetSpIds.Count)" -PercentComplete (($i / [math]::Max($targetSpIds.Count, 1)) * 100)

    try {
        $sp = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($spId)?`$select=$spSelect"
    } catch {
        if (Test-NotFound $_) { Write-Warning "Service principal $spId no longer exists - skipped."; continue }
        throw
    }

    # The app registration exists only in the publisher's tenant; 404 = third-party / multi-tenant app.
    $app = $null
    if ($sp.servicePrincipalType -eq 'Application') {
        try {
            $app = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/applications(appId='$($sp.appId)')?`$select=$appSelect"
        } catch {
            if (-not (Test-NotFound $_)) { Write-Warning "Could not read app registration for $($sp.displayName): $($_.Exception.Message)" }
        }
    }

    # ---- Owners ----
    $owners = @()
    if ($app) { try { $owners = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/owners?`$select=id,displayName,userPrincipalName" } catch { Write-Verbose $_ } }
    if (-not $owners) { try { $owners = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/owners?`$select=id,displayName,userPrincipalName" } catch { Write-Verbose $_ } }
    $ownerList = @($owners | ForEach-Object { if (Get-Prop $_ 'userPrincipalName') { $_.userPrincipalName } else { Get-Prop $_ 'displayName' } }) -join '; '

    # ---- Credentials ----
    # SAML token-signing certs sit on the enterprise app as a 'Sign' key plus paired password/'Verify' key
    # with the same customKeyIdentifier. They are not used to authenticate the app, so skip them.
    $samlKeyIds = @{}
    foreach ($kc in @(Get-Prop $sp 'keyCredentials')) {
        if ($kc -and (Get-Prop $kc 'usage') -eq 'Sign' -and (Get-Prop $kc 'customKeyIdentifier')) { $samlKeyIds[[string]$kc.customKeyIdentifier] = $true }
    }
    $isSaml = { param($c) $k = Get-Prop $c 'customKeyIdentifier'; [bool]($k -and $samlKeyIds.ContainsKey([string]$k)) }

    $creds = [System.Collections.Generic.List[object]]::new()   # Format-Credential results, tagged with Type
    $sources = @()
    if ($app) { $sources += , @('AppRegistration', $app) }
    if ($sp.servicePrincipalType -ne 'ManagedIdentity') { $sources += , @('ServicePrincipal', $sp) }
    foreach ($src in $sources) {
        foreach ($pc in @(Get-Prop $src[1] 'passwordCredentials')) {
            if ($null -eq $pc -or ($src[0] -eq 'ServicePrincipal' -and (& $isSaml $pc))) { continue }
            $c = Format-Credential -Type 'ClientSecret' -Name (Get-Prop $pc 'displayName') -Start (Get-Prop $pc 'startDateTime') -End (Get-Prop $pc 'endDateTime') -Source $src[0]
            $c | Add-Member -NotePropertyName Type -NotePropertyValue 'ClientSecret'
            $creds.Add($c)
        }
        foreach ($kc in @(Get-Prop $src[1] 'keyCredentials')) {
            if ($null -eq $kc -or (Get-Prop $kc 'usage') -ne 'Verify' -or ($src[0] -eq 'ServicePrincipal' -and (& $isSaml $kc))) { continue }
            $c = Format-Credential -Type 'Certificate' -Name (Get-Prop $kc 'displayName') -Start (Get-Prop $kc 'startDateTime') -End (Get-Prop $kc 'endDateTime') -Source $src[0]
            $c | Add-Member -NotePropertyName Type -NotePropertyValue 'Certificate'
            $creds.Add($c)
        }
    }
    if ($app) {
        try {
            foreach ($fc in (Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/federatedIdentityCredentials")) {
                $creds.Add([pscustomobject]@{
                    Text = "FederatedCredential '$($fc.name)': issuer $($fc.issuer), no expiry"; Active = $true; AgeDays = -1; Type = 'FederatedCredential'
                })
            }
        } catch { Write-Verbose $_ }
    }

    # ---- Auth methods in use ----
    $active  = @($creds | Where-Object Active)
    $methods = @($active | ForEach-Object Type | Select-Object -Unique)
    if ($sp.servicePrincipalType -eq 'ManagedIdentity') {
        $methods = @('ManagedIdentity')
        $creds.Add([pscustomobject]@{ Text = 'ManagedIdentity: credentials managed and rotated by Azure'; Active = $true; AgeDays = -1; Type = 'ManagedIdentity' })
    } elseif (-not $methods) {
        $methods = @(if (-not $app -and $sp.servicePrincipalType -eq 'Application') { 'Unknown (app registered in another tenant)' } else { 'None (no valid credential)' })
    }
    $hasSecret = [bool]($active | Where-Object Type -eq 'ClientSecret')

    # ---- Permissions & risk ----
    $perms     = @($permsBySp[$spId])
    $critPerms = @($perms | Where-Object { $PermissionTiers[$_] -eq 'Critical' } | Sort-Object)
    $highPerms = @($perms | Where-Object { $PermissionTiers[$_] -eq 'High' } | Sort-Object)
    $otherPerms = @($perms | Where-Object { -not $PermissionTiers.ContainsKey($_) } | Sort-Object)

    $risk = if ($critPerms.Count -and $hasSecret) { 'Critical' }
            elseif ($critPerms.Count -or ($highPerms.Count -and $hasSecret)) { 'High' }
            elseif ($highPerms.Count) { 'Medium' }
            else { 'Low' }

    # ---- Last sign-in ----
    $lastSignIn = ''
    if ($signInByAppId.ContainsKey([string]$sp.appId)) {
        $lastSignIn = Format-Date (ConvertTo-UtcDate (Get-Prop (Get-Prop $signInByAppId[[string]$sp.appId] 'lastSignInActivity') 'lastSignInDateTime'))
    }
    if (-not $lastSignIn) { $lastSignIn = if ($signInAvailable) { 'No sign-in recorded' } else { 'Unavailable' } }

    $row = [ordered]@{
        RiskRating          = $risk
        AppName             = $sp.displayName
        AppId               = $sp.appId
        Owners              = if ($ownerList) { $ownerList } else { 'No owners' }
        CriticalPermissions = $critPerms -join '; '
        HighPermissions     = $highPerms -join '; '
        AuthMethods         = $methods -join ', '
        CredentialAges      = if ($creds.Count) { (@($creds | Sort-Object AgeDays -Descending) | ForEach-Object Text) -join ' | ' } else { 'No credentials visible in this tenant' }
        LastSignIn          = $lastSignIn
    }
    if ($IncludeNonCritical) { $row['OtherPermissions'] = $otherPerms -join '; ' }
    $rows.Add([pscustomobject]$row)
}
Write-Progress -Activity 'Inspecting enterprise apps' -Completed
#endregion

#region 4. Export --------------------------------------------------------------------------------
$outDir = Split-Path -Path $OutputPath -Parent
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$riskOrder = @{ Critical = 0; High = 1; Medium = 2; Low = 3 }
$rows | Sort-Object @{ Expression = { $riskOrder[$_.RiskRating] } }, AppName |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8

Write-Host ''
Write-Host '================ Summary ================' -ForegroundColor Green
foreach ($level in 'Critical', 'High', 'Medium', 'Low') {
    $n = @($rows | Where-Object RiskRating -eq $level).Count
    if ($n) { Write-Host ("{0,-9}: {1}" -f $level, $n) }
}
Write-Host ("Apps using client secrets: {0}" -f @($rows | Where-Object { $_.AuthMethods -match 'ClientSecret' }).Count)
Write-Host "Report: $OutputPath" -ForegroundColor Green
#endregion
