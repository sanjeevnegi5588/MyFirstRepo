<#
.SYNOPSIS
    Exports every enterprise application (service principal) that holds critical
    Microsoft Graph API permissions, shows whether it authenticates with a client
    secret, and reports the age and expiry of each secret.

.DESCRIPTION
    1. Reads the Microsoft Graph service principal in your tenant to resolve permission IDs to names.
    2. Collects every APPLICATION permission (app role) granted on Microsoft Graph, and optionally
       every DELEGATED permission grant (-IncludeDelegated).
    3. Keeps the service principals that hold at least one permission from the critical list
       (or all of them with -IncludeNonCritical).
    4. For each one, reads the credentials on the app registration and on the service principal:
       client secrets, certificates and federated identity credentials.
    5. Writes three CSV files:
         *_Summary.csv      one row per enterprise app (permissions, auth method, oldest secret age, findings)
         *_Secrets.csv      one row per client secret (created, expires, age in days, status)
         *_Permissions.csv  one row per app + Graph permission

    Only read permissions are needed:
        Application.Read.All, Directory.Read.All
        (+ AuditLog.Read.All with -IncludeSignInActivity, which needs Entra ID P1/P2)

.PARAMETER OutputFolder
    Folder for the CSV files. Created if missing. Defaults to .\GraphAppAudit_<timestamp>.

.PARAMETER IncludeDelegated
    Also evaluate delegated (OAuth2) permission grants on Microsoft Graph, not only application permissions.

.PARAMETER IncludeNonCritical
    Export every app that has any Microsoft Graph permission, not only the ones holding critical permissions.

.PARAMETER IncludeSignInActivity
    Add last sign-in dates per app (beta reports API, requires AuditLog.Read.All and Entra ID P1/P2).

.PARAMETER SecretAgeThresholdDays
    Secrets older than this are flagged in the Findings column. Default 180.

.PARAMETER MaxSecretLifetimeDays
    Secrets whose total lifetime (start -> expiry) exceeds this are flagged as long-lived. Default 365.

.PARAMETER AdditionalCriticalPermissions
    Extra Graph permission names to treat as Critical, e.g. 'AuditLog.Read.All','Reports.Read.All'.

.PARAMETER TenantId
    Tenant to connect to. Optional.

.PARAMETER SkipConnect
    Use an existing Connect-MgGraph session (e.g. app-only with a certificate) instead of connecting interactively.

.EXAMPLE
    .\Export-CriticalGraphApps.ps1

.EXAMPLE
    .\Export-CriticalGraphApps.ps1 -IncludeDelegated -IncludeSignInActivity -SecretAgeThresholdDays 90

.EXAMPLE
    # App-only / unattended
    Connect-MgGraph -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumb> -NoWelcome
    .\Export-CriticalGraphApps.ps1 -SkipConnect -OutputFolder C:\Reports

.NOTES
    Requires PowerShell 7+ and the Microsoft.Graph.Authentication module:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [string]$OutputFolder = (Join-Path -Path (Get-Location) -ChildPath ("GraphAppAudit_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),
    [switch]$IncludeDelegated,
    [switch]$IncludeNonCritical,
    [switch]$IncludeSignInActivity,
    [ValidateRange(1, 36500)][int]$SecretAgeThresholdDays = 180,
    [ValidateRange(1, 36500)][int]$MaxSecretLifetimeDays = 365,
    [string[]]$AdditionalCriticalPermissions = @(),
    [string]$TenantId,
    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Critical permission catalogue ------------------------------------------------------------
# Critical = can lead to tenant takeover / privilege escalation to Global Admin.
# High     = broad access to tenant data (mail, files, chats, users, groups).
# Edit this list to match your organisation's risk model.
$PermissionTiers = @{
    # --- Critical: privilege escalation / tenant control ---
    'RoleManagement.ReadWrite.Directory'                     = 'Critical'
    'RoleAssignmentSchedule.ReadWrite.Directory'             = 'Critical'
    'RoleEligibilitySchedule.ReadWrite.Directory'            = 'Critical'
    'RoleManagementPolicy.ReadWrite.Directory'               = 'Critical'
    'PrivilegedAccess.ReadWrite.AzureAD'                     = 'Critical'
    'PrivilegedAccess.ReadWrite.AzureADGroup'                = 'Critical'
    'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup'    = 'Critical'
    'PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup'   = 'Critical'
    'AppRoleAssignment.ReadWrite.All'                        = 'Critical'
    'Application.ReadWrite.All'                              = 'Critical'
    'Directory.ReadWrite.All'                                = 'Critical'
    'Directory.AccessAsUser.All'                             = 'Critical'
    'Domain.ReadWrite.All'                                   = 'Critical'
    'Organization.ReadWrite.All'                             = 'Critical'
    'IdentityProvider.ReadWrite.All'                         = 'Critical'
    'Policy.ReadWrite.ConditionalAccess'                     = 'Critical'
    'Policy.ReadWrite.AuthenticationMethod'                  = 'Critical'
    'Policy.ReadWrite.PermissionGrant'                       = 'Critical'
    'Policy.ReadWrite.CrossTenantAccess'                     = 'Critical'
    'Policy.ReadWrite.ApplicationConfiguration'              = 'Critical'
    'UserAuthenticationMethod.ReadWrite.All'                 = 'Critical'
    'User-PasswordProfile.ReadWrite.All'                     = 'Critical'
    'User.ReadWrite.All'                                     = 'Critical'
    'Group.ReadWrite.All'                                    = 'Critical'
    'GroupMember.ReadWrite.All'                              = 'Critical'
    'ServicePrincipalEndpoint.ReadWrite.All'                 = 'Critical'
    'EntitlementManagement.ReadWrite.All'                    = 'Critical'
    'DeviceManagementRBAC.ReadWrite.All'                     = 'Critical'
    'DeviceManagementConfiguration.ReadWrite.All'            = 'Critical'
    'DeviceManagementManagedDevices.ReadWrite.All'           = 'Critical'
    'DeviceManagementManagedDevices.PrivilegedOperations.All' = 'Critical'
    'DeviceManagementApps.ReadWrite.All'                     = 'Critical'

    # --- High: broad data access / limited escalation ---
    'Application.ReadWrite.OwnedBy'                          = 'High'
    'User.DeleteRestore.All'                                 = 'High'
    'User.EnableDisableAccount.All'                          = 'High'
    'User.ManageIdentities.All'                              = 'High'
    'Mail.ReadWrite'                                         = 'High'
    'Mail.Read'                                              = 'High'
    'Mail.Send'                                              = 'High'
    'MailboxSettings.ReadWrite'                              = 'High'
    'Calendars.ReadWrite'                                    = 'High'
    'Contacts.ReadWrite'                                     = 'High'
    'Files.ReadWrite.All'                                    = 'High'
    'Files.Read.All'                                         = 'High'
    'Sites.FullControl.All'                                  = 'High'
    'Sites.Manage.All'                                       = 'High'
    'Sites.ReadWrite.All'                                    = 'High'
    'Sites.Read.All'                                         = 'High'
    'Notes.ReadWrite.All'                                    = 'High'
    'Chat.Read.All'                                          = 'High'
    'Chat.ReadWrite.All'                                     = 'High'
    'ChannelMessage.Read.All'                                = 'High'
    'TeamMember.ReadWrite.All'                               = 'High'
    'eDiscovery.ReadWrite.All'                               = 'High'
    'InformationProtectionPolicy.Read.All'                   = 'High'
    'SecurityEvents.ReadWrite.All'                           = 'High'
    'IdentityRiskyUser.ReadWrite.All'                        = 'High'
}
foreach ($p in $AdditionalCriticalPermissions) { if ($p) { $PermissionTiers[$p.Trim()] = 'Critical' } }
#endregion

#region Helpers ----------------------------------------------------------------------------------
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$Now        = [datetime]::UtcNow

function Invoke-Graph {
    param([Parameter(Mandatory)][string]$Uri)
    # Microsoft.Graph SDK already retries 429/503 with back-off.
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

function Test-NotFound {
    param($ErrorRecord)
    $msg = "$($ErrorRecord.Exception.Message) $($ErrorRecord.ErrorDetails)"
    return ($msg -match 'NotFound|404|Request_ResourceNotFound|does not exist')
}

function Format-Date { param($d) if ($d) { $d.ToString('yyyy-MM-dd') } else { '' } }
#endregion

#region Connect ----------------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Module 'Microsoft.Graph.Authentication' not found. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication

if (-not $SkipConnect) {
    $scopes = @('Application.Read.All', 'Directory.Read.All')
    if ($IncludeSignInActivity) { $scopes += 'AuditLog.Read.All' }
    $connectParams = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams
}
$ctx = Get-MgContext
if (-not $ctx) { throw 'Not connected to Microsoft Graph. Run Connect-MgGraph or drop -SkipConnect.' }
Write-Host "Connected to tenant $($ctx.TenantId) as $(if ($ctx.Account) { $ctx.Account } else { $ctx.ClientId })" -ForegroundColor Cyan
#endregion

#region 1. Resolve Microsoft Graph permissions ---------------------------------------------------
Write-Host 'Reading Microsoft Graph service principal...' -ForegroundColor Cyan
$graphSp = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$GraphAppId')?`$select=id,appRoles"
$appRoleById = @{}
foreach ($r in $graphSp.appRoles) { $appRoleById[[string]$r.id] = $r.value }

# Per-service-principal permission bag: spId -> list of permission records
$permsBySp = @{}
function Add-Perm {
    param([string]$SpId, [string]$Name, [string]$Type, [string]$ConsentType, $GrantedOn)
    if (-not $Name) { return }
    if (-not $permsBySp.ContainsKey($SpId)) { $permsBySp[$SpId] = [System.Collections.Generic.List[object]]::new() }
    $tier = if ($PermissionTiers.ContainsKey($Name)) { $PermissionTiers[$Name] } else { 'Other' }
    $permsBySp[$SpId].Add([pscustomobject]@{
        Permission  = $Name
        Type        = $Type
        ConsentType = $ConsentType
        Tier        = $tier
        GrantedOn   = Format-Date (ConvertTo-UtcDate $GrantedOn)
    })
}
#endregion

#region 2. Application permissions (app roles) granted on Graph ---------------------------------
Write-Host 'Reading application permissions granted on Microsoft Graph...' -ForegroundColor Cyan
$assignments = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($graphSp.id)/appRoleAssignedTo?`$top=999"
foreach ($a in $assignments) {
    if ($a.principalType -ne 'ServicePrincipal') { continue }
    $name = $appRoleById[[string]$a.appRoleId]
    Add-Perm -SpId $a.principalId -Name $name -Type 'Application' -ConsentType 'Admin' -GrantedOn (Get-Prop $a 'createdDateTime')
}
Write-Host ("  {0} application permission assignments" -f $assignments.Count)
#endregion

#region 3. Delegated permissions (optional) ------------------------------------------------------
if ($IncludeDelegated) {
    Write-Host 'Reading delegated permission grants on Microsoft Graph...' -ForegroundColor Cyan
    $grants = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=resourceId eq '$($graphSp.id)'&`$top=999"
    $seen = @{}
    foreach ($g in $grants) {
        $consent = if ($g.consentType -eq 'AllPrincipals') { 'Admin (all users)' } else { 'User' }
        foreach ($scope in ($g.scope -split '\s+' | Where-Object { $_ })) {
            $key = "$($g.clientId)|$scope|$consent"
            if ($seen.ContainsKey($key)) { continue }   # one row per app/scope/consent type, not per user
            $seen[$key] = $true
            Add-Perm -SpId $g.clientId -Name $scope -Type 'Delegated' -ConsentType $consent -GrantedOn $null
        }
    }
    Write-Host ("  {0} delegated permission grants" -f $grants.Count)
}
#endregion

#region 4. Select apps in scope ------------------------------------------------------------------
$targetSpIds = @($permsBySp.Keys | Where-Object {
    $IncludeNonCritical -or ($permsBySp[$_] | Where-Object { $_.Tier -in 'Critical', 'High' })
})
Write-Host ("{0} enterprise apps in scope" -f $targetSpIds.Count) -ForegroundColor Cyan
#endregion

#region 4b. Optional sign-in activity
$signInByAppId = @{}
if ($IncludeSignInActivity) {
    try {
        Write-Host 'Reading service principal sign-in activity (beta)...' -ForegroundColor Cyan
        $acts = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities'
        foreach ($s in $acts) { $signInByAppId[[string]$s.appId] = $s }
    } catch {
        Write-Warning "Could not read sign-in activity (needs AuditLog.Read.All and Entra ID P1/P2): $($_.Exception.Message)"
    }
}
#endregion

#region 5. Inspect each app ----------------------------------------------------------------------
$summaryRows    = [System.Collections.Generic.List[object]]::new()
$secretRows     = [System.Collections.Generic.List[object]]::new()
$permissionRows = [System.Collections.Generic.List[object]]::new()

$spSelect  = 'id,appId,displayName,servicePrincipalType,accountEnabled,appOwnerOrganizationId,publisherName,passwordCredentials,keyCredentials'
$appSelect = 'id,appId,displayName,createdDateTime,signInAudience,passwordCredentials,keyCredentials'

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

    # App registration lives only in the publisher's tenant. 404 => multi-tenant / Microsoft / external app.
    $app = $null
    if ($sp.servicePrincipalType -eq 'Application') {
        try {
            $app = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/applications(appId='$($sp.appId)')?`$select=$appSelect"
        } catch {
            if (-not (Test-NotFound $_)) { Write-Warning "Could not read app registration for $($sp.displayName): $($_.Exception.Message)" }
        }
    }

    $fedCreds = @()
    $owners   = @()
    if ($app) {
        try { $fedCreds = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/federatedIdentityCredentials" } catch { Write-Verbose $_ }
        try { $owners   = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/owners?`$select=id,displayName,userPrincipalName" } catch { Write-Verbose $_ }
    }
    if (-not $owners) {
        try { $owners = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/owners?`$select=id,displayName,userPrincipalName" } catch { Write-Verbose $_ }
    }

    # SAML token-signing certificates live on the service principal as a 'Sign' key credential plus a
    # paired password credential (same customKeyIdentifier). Those are not client secrets - exclude them.
    $samlKeyIds = @{}
    foreach ($kc in @(Get-Prop $sp 'keyCredentials')) {
        if ($kc -and (Get-Prop $kc 'usage') -eq 'Sign' -and (Get-Prop $kc 'customKeyIdentifier')) { $samlKeyIds[[string]$kc.customKeyIdentifier] = $true }
    }
    function Test-SamlPair { param($Cred) $cki = Get-Prop $Cred 'customKeyIdentifier'; return ($cki -and $samlKeyIds.ContainsKey([string]$cki)) }

    # ---- Client secrets (app registration + service principal) ----
    $secrets = [System.Collections.Generic.List[object]]::new()
    $sources = @()
    if ($app) { $sources += , @('AppRegistration', @(Get-Prop $app 'passwordCredentials')) }
    $sources += , @('ServicePrincipal', @(Get-Prop $sp 'passwordCredentials'))
    foreach ($src in $sources) {
        foreach ($pc in $src[1]) {
            if ($null -eq $pc) { continue }
            if ($src[0] -eq 'ServicePrincipal' -and (Test-SamlPair $pc)) { continue }
            $start = ConvertTo-UtcDate (Get-Prop $pc 'startDateTime')
            $end   = ConvertTo-UtcDate (Get-Prop $pc 'endDateTime')
            $status = if ($end -and $end -lt $Now) { 'Expired' } elseif ($start -and $start -gt $Now) { 'NotYetValid' } else { 'Active' }
            $secret = [pscustomobject]@{
                AppDisplayName  = $sp.displayName
                AppId           = $sp.appId
                ServicePrincipalId = $sp.id
                CredentialSource = $src[0]
                SecretName      = Get-Prop $pc 'displayName'
                SecretHint      = Get-Prop $pc 'hint'
                KeyId           = Get-Prop $pc 'keyId'
                Created         = Format-Date $start
                Expires         = Format-Date $end
                AgeDays         = if ($start) { [int][math]::Floor(($Now - $start).TotalDays) } else { $null }
                DaysToExpiry    = if ($end)   { [int][math]::Floor(($end - $Now).TotalDays) } else { $null }
                LifetimeDays    = if ($start -and $end) { [int][math]::Round(($end - $start).TotalDays) } else { $null }
                Status          = $status
            }
            $secrets.Add($secret)
            $secretRows.Add($secret)
        }
    }
    $activeSecrets = @($secrets | Where-Object Status -eq 'Active')

    # ---- Certificates used for client auth (usage = Verify) ----
    $certs = @()
    if ($sp.servicePrincipalType -ne 'ManagedIdentity') {
        if ($app) { $certs += @(@(Get-Prop $app 'keyCredentials') | Where-Object { $_ -and (Get-Prop $_ 'usage') -eq 'Verify' }) }
        $certs += @(@(Get-Prop $sp 'keyCredentials') | Where-Object { $_ -and (Get-Prop $_ 'usage') -eq 'Verify' -and -not (Test-SamlPair $_) })
    }
    $activeCerts = @($certs | Where-Object { $e = ConvertTo-UtcDate (Get-Prop $_ 'endDateTime'); -not $e -or $e -ge $Now })

    # ---- Auth method ----
    $methods = @()
    if ($sp.servicePrincipalType -eq 'ManagedIdentity') { $methods += 'ManagedIdentity' }
    if ($activeSecrets.Count) { $methods += 'ClientSecret' }
    if ($activeCerts.Count)   { $methods += 'Certificate' }
    if ($fedCreds.Count)      { $methods += 'FederatedCredential' }
    if (-not $methods) {
        if ($sp.servicePrincipalType -eq 'Application' -and -not $app) { $methods += 'Unknown (app registered in another tenant)' }
        else { $methods += 'None (no valid credential)' }
    }

    # ---- Permissions ----
    $perms     = $permsBySp[$spId]
    $critPerms = @($perms | Where-Object Tier -eq 'Critical' | ForEach-Object { "$($_.Permission) ($($_.Type))" } | Sort-Object -Unique)
    $highPerms = @($perms | Where-Object Tier -eq 'High'     | ForEach-Object { "$($_.Permission) ($($_.Type))" } | Sort-Object -Unique)
    $allPerms  = @($perms | ForEach-Object { "$($_.Permission) ($($_.Type))" } | Sort-Object -Unique)
    foreach ($p in $perms) {
        $permissionRows.Add([pscustomobject]@{
            AppDisplayName     = $sp.displayName
            AppId              = $sp.appId
            ServicePrincipalId = $sp.id
            Permission         = $p.Permission
            PermissionType     = $p.Type
            ConsentType        = $p.ConsentType
            Tier               = $p.Tier
            GrantedOn          = $p.GrantedOn
        })
    }

    # ---- Secret ages ----
    $oldestActive = $activeSecrets | Where-Object { $null -ne $_.AgeDays } | Sort-Object AgeDays -Descending | Select-Object -First 1
    $newestActive = $activeSecrets | Where-Object { $null -ne $_.AgeDays } | Sort-Object AgeDays | Select-Object -First 1
    $nextExpiry   = $activeSecrets | Where-Object { $null -ne $_.DaysToExpiry } | Sort-Object DaysToExpiry | Select-Object -First 1

    # ---- Findings ----
    $findings = @()
    if ($activeSecrets.Count -and $critPerms.Count) { $findings += 'Critical Graph permission + client secret' }
    elseif ($activeSecrets.Count)                    { $findings += 'Uses client secret' }
    if ($activeSecrets | Where-Object { $_.AgeDays -gt $SecretAgeThresholdDays }) { $findings += "Secret older than $SecretAgeThresholdDays days" }
    if ($activeSecrets | Where-Object { $_.LifetimeDays -gt $MaxSecretLifetimeDays }) { $findings += "Secret lifetime > $MaxSecretLifetimeDays days" }
    if ($activeSecrets.Count -gt 1) { $findings += "Multiple active secrets ($($activeSecrets.Count))" }
    if ($secrets | Where-Object Status -eq 'Expired') { $findings += 'Expired secrets not cleaned up' }
    if ($nextExpiry -and $nextExpiry.DaysToExpiry -le 30) { $findings += 'Secret expires within 30 days' }
    if ($activeSecrets | Where-Object CredentialSource -eq 'ServicePrincipal') { $findings += 'Secret on service principal (hidden from App registrations blade)' }
    if (-not $owners.Count -and $sp.servicePrincipalType -eq 'Application') { $findings += 'No owners' }
    if (-not $sp.accountEnabled) { $findings += 'Disabled' }

    $riskRating = if ($critPerms.Count -and $activeSecrets.Count) { 'Critical' }
                  elseif ($critPerms.Count -or ($highPerms.Count -and $activeSecrets.Count)) { 'High' }
                  elseif ($highPerms.Count) { 'Medium' }
                  else { 'Low' }

    $signIn = $signInByAppId[[string]$sp.appId]
    $lastSignIn       = if ($signIn) { Format-Date (ConvertTo-UtcDate (Get-Prop (Get-Prop $signIn 'lastSignInActivity') 'lastSignInDateTime')) } else { '' }
    $lastAppOnlyLogin = if ($signIn) { Format-Date (ConvertTo-UtcDate (Get-Prop (Get-Prop $signIn 'applicationAuthenticationClientSignInActivity') 'lastSignInDateTime')) } else { '' }

    $summaryRows.Add([pscustomobject]@{
        RiskRating                 = $riskRating
        AppDisplayName             = $sp.displayName
        AppId                      = $sp.appId
        ServicePrincipalId         = $sp.id
        ServicePrincipalType       = $sp.servicePrincipalType
        Enabled                    = $sp.accountEnabled
        Publisher                  = Get-Prop $sp 'publisherName'
        RegisteredInThisTenant     = [bool]$app
        AppCreated                 = if ($app) { Format-Date (ConvertTo-UtcDate (Get-Prop $app 'createdDateTime')) } else { '' }
        Owners                     = ($owners | ForEach-Object { if (Get-Prop $_ 'userPrincipalName') { $_.userPrincipalName } else { Get-Prop $_ 'displayName' } }) -join '; '
        CriticalPermissions        = $critPerms -join '; '
        HighPermissions            = $highPerms -join '; '
        CriticalPermissionCount    = $critPerms.Count
        AllGraphPermissions        = $allPerms -join '; '
        AuthMethods                = $methods -join ', '
        UsesClientSecret           = [bool]$activeSecrets.Count
        ActiveSecretCount          = $activeSecrets.Count
        ExpiredSecretCount         = @($secrets | Where-Object Status -eq 'Expired').Count
        OldestActiveSecretAgeDays  = if ($oldestActive) { $oldestActive.AgeDays } else { $null }
        OldestActiveSecretCreated  = if ($oldestActive) { $oldestActive.Created } else { '' }
        NewestActiveSecretAgeDays  = if ($newestActive) { $newestActive.AgeDays } else { $null }
        NextSecretExpiry           = if ($nextExpiry) { $nextExpiry.Expires } else { '' }
        NextSecretDaysToExpiry     = if ($nextExpiry) { $nextExpiry.DaysToExpiry } else { $null }
        ActiveCertificateCount     = $activeCerts.Count
        FederatedCredentialCount   = $fedCreds.Count
        LastSignIn                 = $lastSignIn
        LastAppOnlySignIn          = $lastAppOnlyLogin
        Findings                   = $findings -join '; '
    })
}
Write-Progress -Activity 'Inspecting enterprise apps' -Completed
#endregion

#region 6. Export --------------------------------------------------------------------------------
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$stamp       = Get-Date -Format 'yyyyMMdd_HHmmss'
$summaryPath = Join-Path $OutputFolder "CriticalGraphApps_${stamp}_Summary.csv"
$secretsPath = Join-Path $OutputFolder "CriticalGraphApps_${stamp}_Secrets.csv"
$permsPath   = Join-Path $OutputFolder "CriticalGraphApps_${stamp}_Permissions.csv"

$riskOrder = @{ Critical = 0; High = 1; Medium = 2; Low = 3 }
$summaryRows |
    Sort-Object @{ Expression = { $riskOrder[$_.RiskRating] } }, @{ Expression = 'OldestActiveSecretAgeDays'; Descending = $true }, AppDisplayName |
    Export-Csv -Path $summaryPath -NoTypeInformation -Encoding utf8

$inScope = @{}; foreach ($r in $summaryRows) { $inScope[$r.ServicePrincipalId] = $true }
$secretRows | Where-Object { $inScope.ContainsKey($_.ServicePrincipalId) } |
    Sort-Object AppDisplayName, @{ Expression = 'AgeDays'; Descending = $true } |
    Export-Csv -Path $secretsPath -NoTypeInformation -Encoding utf8

$permissionRows |
    Sort-Object AppDisplayName, @{ Expression = { switch ($_.Tier) { 'Critical' { 0 } 'High' { 1 } default { 2 } } } }, Permission |
    Export-Csv -Path $permsPath -NoTypeInformation -Encoding utf8
#endregion

#region 7. Console summary -----------------------------------------------------------------------
$withSecret = @($summaryRows | Where-Object UsesClientSecret)
$critSecret = @($summaryRows | Where-Object RiskRating -eq 'Critical')
$oldSecret  = @($withSecret | Where-Object { $_.OldestActiveSecretAgeDays -gt $SecretAgeThresholdDays })

Write-Host ''
Write-Host '================ Summary ================' -ForegroundColor Green
Write-Host ("Apps in scope                         : {0}" -f $summaryRows.Count)
Write-Host ("Apps using client secrets             : {0}" -f $withSecret.Count)
Write-Host ("Critical permission + client secret   : {0}" -f $critSecret.Count) -ForegroundColor Red
Write-Host ("Secret older than {0,-4} days          : {1}" -f $SecretAgeThresholdDays, $oldSecret.Count) -ForegroundColor Yellow
Write-Host ''
if ($critSecret.Count) {
    $critSecret | Sort-Object OldestActiveSecretAgeDays -Descending |
        Select-Object AppDisplayName, AppId, ActiveSecretCount, OldestActiveSecretAgeDays, NextSecretExpiry, CriticalPermissionCount |
        Format-Table -AutoSize | Out-Host
}
Write-Host "Summary     : $summaryPath"
Write-Host "Secrets     : $secretsPath"
Write-Host "Permissions : $permsPath"
#endregion
