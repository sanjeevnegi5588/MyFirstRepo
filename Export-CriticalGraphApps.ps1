<#
.SYNOPSIS
    Exports every enterprise application (service principal) in the tenant to four CSV files for
    Power BI: apps, application permissions, credentials (auth methods) and owners.

.DESCRIPTION
    Files written to -OutputFolder (fixed names, overwritten on every run so Power BI refresh keeps working):

      Apps.csv            one row per app      AppId (key), name, type, Microsoft/third-party flags,
                                               auth methods, counts, last sign-in dates, snapshot date
      AppPermissions.csv  one row per app + application permission (all APIs, e.g. Microsoft Graph,
                                               Office 365 Exchange Online, SharePoint)
      AppCredentials.csv  one row per client secret / certificate / federated credential / managed identity,
                                               with start and end dates and status
      AppOwners.csv       one row per app + owner, plus a '(No owner)' row for apps without owners

    All files share the AppId column. In Power BI, relate Apps[AppId] 1-to-many to AppId in the other
    three tables. Dates are written as yyyy-MM-dd; calculate ages in Power BI so they stay current.

    Read-only Graph permissions used: Application.Read.All, Directory.Read.All, AuditLog.Read.All.
    Last sign-in dates need Entra ID P1/P2; without them the columns are left empty.

.PARAMETER OutputFolder
    Folder for the CSV files. Created if missing. Default: .\EntraAppReport

.PARAMETER TenantId
    Tenant to connect to. Optional.

.PARAMETER SkipConnect
    Use an existing Connect-MgGraph session (e.g. app-only with a certificate for scheduled runs).

.EXAMPLE
    .\Export-CriticalGraphApps.ps1 -OutputFolder 'C:\Users\me\OneDrive - Contoso\EntraAppReport'

.EXAMPLE
    Connect-MgGraph -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumb> -NoWelcome
    .\Export-CriticalGraphApps.ps1 -SkipConnect -OutputFolder D:\Reports\EntraAppReport

.NOTES
    Requires PowerShell 7+ and: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [string]$OutputFolder = (Join-Path -Path (Get-Location) -ChildPath 'EntraAppReport'),
    [string]$TenantId,
    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Now          = [datetime]::UtcNow
$SnapshotDate = $Now.ToString('yyyy-MM-dd')
$GraphRoot    = 'https://graph.microsoft.com/v1.0'
# Tenants that own Microsoft's first-party apps
$MicrosoftTenantIds = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')

#region Helpers ----------------------------------------------------------------------------------
function Invoke-Graph {
    param([Parameter(Mandatory)][string]$Uri)
    # The Microsoft Graph SDK retries 429/503 with back-off on its own.
    Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
}

function Get-Prop {
    # Safe property read under StrictMode
    param($Object, [string]$Name)
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $null
}

function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    while ($next) {
        $resp = Invoke-Graph -Uri $next
        foreach ($v in @(Get-Prop $resp 'value')) { if ($null -ne $v) { $items.Add($v) } }
        $next = Get-Prop $resp '@odata.nextLink'
    }
    return , $items
}

function Invoke-GraphBatch {
    # Runs many GET requests 20 at a time through Graph JSON batching.
    # Input: list of @{ Key = '<your key>'; Url = '/relative/url' }. Output: hashtable Key -> list of items.
    param([Parameter(Mandatory)][System.Collections.IList]$Requests, [string]$Activity = 'Reading details')
    $results = @{}
    for ($start = 0; $start -lt $Requests.Count; $start += 20) {
        Write-Progress -Activity $Activity -Status "$start of $($Requests.Count)" -PercentComplete (($start / [math]::Max($Requests.Count, 1)) * 100)
        $end   = [math]::Min($start + 19, $Requests.Count - 1)
        $chunk = @($Requests[$start..$end])
        $body  = @{ requests = @(for ($n = 0; $n -lt $chunk.Count; $n++) { @{ id = "$n"; method = 'GET'; url = $chunk[$n].Url } }) } |
                 ConvertTo-Json -Depth 5
        $resp  = Invoke-MgGraphRequest -Method POST -Uri "$GraphRoot/`$batch" -Body $body -ContentType 'application/json' -OutputType PSObject

        foreach ($r in @($resp.responses)) {
            $req = $chunk[[int]$r.id]
            if ($r.status -eq 200) {
                $items = [System.Collections.Generic.List[object]]::new()
                foreach ($v in @(Get-Prop $r.body 'value')) { if ($null -ne $v) { $items.Add($v) } }
                $nextLink = Get-Prop $r.body '@odata.nextLink'
                if ($nextLink) { foreach ($v in (Invoke-GraphPaged -Uri $nextLink)) { $items.Add($v) } }
                $results[$req.Key] = $items
            } elseif ($r.status -eq 404) {
                $results[$req.Key] = [System.Collections.Generic.List[object]]::new()
            } else {
                # Throttled or failed inside the batch: retry on its own (the SDK handles back-off).
                try   { $results[$req.Key] = Invoke-GraphPaged -Uri "$GraphRoot$($req.Url)" }
                catch {
                    Write-Warning "Request failed ($($req.Url)): $($_.Exception.Message)"
                    $results[$req.Key] = [System.Collections.Generic.List[object]]::new()
                }
            }
        }
    }
    Write-Progress -Activity $Activity -Completed
    return $results
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

function Format-Date { param($Value) $d = ConvertTo-UtcDate $Value; if ($d) { $d.ToString('yyyy-MM-dd') } else { '' } }

function Get-Thumbprint {
    # A certificate's customKeyIdentifier is normally its SHA-1 thumbprint, base64-encoded.
    param($Value)
    if (-not $Value) { return '' }
    try {
        $bytes = [Convert]::FromBase64String([string]$Value)
        if ($bytes.Length -eq 20) { return ([BitConverter]::ToString($bytes) -replace '-', '') }
    } catch { }
    return [string]$Value
}

function Export-Table {
    # Export-Csv writes nothing for an empty table; Power BI needs the header row, so write it ourselves.
    param($Rows, [string]$Path, [string[]]$Columns)
    if (@($Rows).Count) {
        $Rows | Select-Object $Columns | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8
    } else {
        Set-Content -Path $Path -Value (($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') -Encoding utf8
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

#region 1. Read service principals and app registrations ------------------------------------------
Write-Host 'Reading enterprise apps (service principals)...' -ForegroundColor Cyan
$spSelect = 'id,appId,displayName,servicePrincipalType,accountEnabled,appOwnerOrganizationId,publisherName,passwordCredentials,keyCredentials,appRoles'
$allSps   = Invoke-GraphPaged -Uri "$GraphRoot/servicePrincipals?`$select=$spSelect&`$top=999"
$spById   = @{}
foreach ($s in $allSps) { $spById[[string]$s.id] = $s }
# Only real apps and managed identities (skips legacy and social-IdP placeholder objects)
$sps = @($allSps | Where-Object { $_.servicePrincipalType -in 'Application', 'ManagedIdentity' })
Write-Host "  $($sps.Count) enterprise apps"

Write-Host 'Reading app registrations...' -ForegroundColor Cyan
$apps        = Invoke-GraphPaged -Uri "$GraphRoot/applications?`$select=id,appId,createdDateTime,passwordCredentials,keyCredentials&`$top=999"
$appByAppId  = @{}
foreach ($a in $apps) { $appByAppId[[string]$a.appId] = $a }
Write-Host "  $($apps.Count) app registrations"
#endregion

#region 2. Permissions, owners and federated credentials (batched) ------------------------------
$requests = [System.Collections.Generic.List[object]]::new()
foreach ($sp in $sps) {
    $requests.Add(@{ Key = "perm|$($sp.id)"; Url = "/servicePrincipals/$($sp.id)/appRoleAssignments" })
    if ($sp.servicePrincipalType -eq 'Application' -and $sp.appOwnerOrganizationId -notin $MicrosoftTenantIds) {
        $requests.Add(@{ Key = "spowner|$($sp.id)"; Url = "/servicePrincipals/$($sp.id)/owners?`$select=id,displayName,userPrincipalName" })
    }
    $app = $appByAppId[[string]$sp.appId]
    if ($app) {
        $requests.Add(@{ Key = "appowner|$($sp.id)"; Url = "/applications/$($app.id)/owners?`$select=id,displayName,userPrincipalName" })
        $requests.Add(@{ Key = "fic|$($sp.id)";      Url = "/applications/$($app.id)/federatedIdentityCredentials" })
    }
}
Write-Host "Reading permissions, owners and federated credentials ($($requests.Count) requests, batched)..." -ForegroundColor Cyan
$batch = Invoke-GraphBatch -Requests $requests -Activity 'Reading permissions, owners and federated credentials'
#endregion

#region 3. Last sign-in dates ---------------------------------------------------------------------
$signInByAppId = @{}
try {
    Write-Host 'Reading service principal sign-in activity...' -ForegroundColor Cyan
    foreach ($s in (Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities')) {
        $signInByAppId[[string]$s.appId] = $s
    }
} catch {
    Write-Warning "Last sign-in dates unavailable (needs AuditLog.Read.All and Entra ID P1/P2): $($_.Exception.Message)"
}
#endregion

#region 4. Build the four tables ------------------------------------------------------------------
$appRows   = [System.Collections.Generic.List[object]]::new()
$permRows  = [System.Collections.Generic.List[object]]::new()
$credRows  = [System.Collections.Generic.List[object]]::new()
$ownerRows = [System.Collections.Generic.List[object]]::new()
$roleNameCache = @{}   # resourceId -> @{ roleId -> appRole }

foreach ($sp in $sps) {
    $appId       = [string]$sp.appId
    $app         = $appByAppId[$appId]
    $isMicrosoft = $sp.appOwnerOrganizationId -in $MicrosoftTenantIds

    # ---- Permissions ----
    $permCount = 0
    $perms = $batch["perm|$($sp.id)"]
    foreach ($pa in @($perms)) {
        if ($null -eq $pa) { continue }
        $resId = [string]$pa.resourceId
        if (-not $roleNameCache.ContainsKey($resId)) {
            $map = @{}
            foreach ($role in @(Get-Prop $spById[$resId] 'appRoles')) { if ($role) { $map[[string]$role.id] = $role } }
            $roleNameCache[$resId] = $map
        }
        $role = $roleNameCache[$resId][[string]$pa.appRoleId]
        $permRows.Add([pscustomobject]@{
            AppId                 = $appId
            AppName               = $sp.displayName
            Resource              = $pa.resourceDisplayName
            Permission            = if ($role) { $role.value } elseif ([string]$pa.appRoleId -eq '00000000-0000-0000-0000-000000000000') { '(default access)' } else { [string]$pa.appRoleId }
            PermissionDescription = if ($role) { $role.displayName } else { '' }
            GrantedDate           = Format-Date (Get-Prop $pa 'createdDateTime')
        })
        $permCount++
    }

    # ---- Owners (app registration and enterprise app) ----
    $ownerCount = 0
    $seenOwner  = @{}
    foreach ($src in @(@('AppRegistration', $batch["appowner|$($sp.id)"]), @('EnterpriseApp', $batch["spowner|$($sp.id)"]))) {
        foreach ($o in @($src[1])) {
            if ($null -eq $o) { continue }
            $type = ([string](Get-Prop $o '@odata.type')) -replace '^#microsoft\.graph\.', ''
            $ownerRows.Add([pscustomobject]@{
                AppId     = $appId
                AppName   = $sp.displayName
                OwnerOf   = $src[0]
                OwnerName = Get-Prop $o 'displayName'
                OwnerUPN  = Get-Prop $o 'userPrincipalName'
                OwnerType = if ($type) { $type } else { 'user' }
                OwnerId   = $o.id
            })
            if (-not $seenOwner.ContainsKey([string]$o.id)) { $seenOwner[[string]$o.id] = $true; $ownerCount++ }
        }
    }
    if ($ownerCount -eq 0) {
        $ownerRows.Add([pscustomobject]@{ AppId = $appId; AppName = $sp.displayName; OwnerOf = ''; OwnerName = '(No owner)'; OwnerUPN = ''; OwnerType = 'None'; OwnerId = '' })
    }

    # ---- Credentials ----
    # SAML token-signing certificates sit on the enterprise app as a 'Sign' key plus a paired password and
    # 'Verify' key with the same customKeyIdentifier. They don't authenticate the app, so skip them.
    $samlKeyIds = @{}
    foreach ($kc in @(Get-Prop $sp 'keyCredentials')) {
        if ($kc -and (Get-Prop $kc 'usage') -eq 'Sign' -and (Get-Prop $kc 'customKeyIdentifier')) { $samlKeyIds[[string]$kc.customKeyIdentifier] = $true }
    }

    $appCreds = [System.Collections.Generic.List[object]]::new()
    $sources = @()
    if ($app) { $sources += , @('AppRegistration', $app) }
    if ($sp.servicePrincipalType -ne 'ManagedIdentity') { $sources += , @('EnterpriseApp', $sp) }
    foreach ($src in $sources) {
        foreach ($c in @(Get-Prop $src[1] 'passwordCredentials')) {
            if ($null -eq $c) { continue }
            $cki = [string](Get-Prop $c 'customKeyIdentifier')
            if ($src[0] -eq 'EnterpriseApp' -and $cki -and $samlKeyIds.ContainsKey($cki)) { continue }
            $appCreds.Add(@{ Type = 'ClientSecret'; Source = $src[0]; Cred = $c; Detail = @{ Hint = Get-Prop $c 'hint' } })
        }
        foreach ($c in @(Get-Prop $src[1] 'keyCredentials')) {
            if ($null -eq $c -or (Get-Prop $c 'usage') -ne 'Verify') { continue }
            $cki = [string](Get-Prop $c 'customKeyIdentifier')
            if ($src[0] -eq 'EnterpriseApp' -and $cki -and $samlKeyIds.ContainsKey($cki)) { continue }
            $appCreds.Add(@{ Type = 'Certificate'; Source = $src[0]; Cred = $c; Detail = @{ Thumbprint = Get-Thumbprint $cki } })
        }
    }

    $activeTypes = [System.Collections.Generic.List[string]]::new()
    foreach ($ac in $appCreds) {
        $start  = ConvertTo-UtcDate (Get-Prop $ac.Cred 'startDateTime')
        $endDt  = ConvertTo-UtcDate (Get-Prop $ac.Cred 'endDateTime')
        $status = if ($endDt -and $endDt -lt $Now) { 'Expired' } elseif ($start -and $start -gt $Now) { 'NotYetValid' } else { 'Active' }
        if ($status -ne 'Expired' -and -not $activeTypes.Contains($ac.Type)) { $activeTypes.Add($ac.Type) }
        $credRows.Add([pscustomobject]@{
            AppId          = $appId
            AppName        = $sp.displayName
            CredentialType = $ac.Type
            Source         = $ac.Source
            Name           = Get-Prop $ac.Cred 'displayName'
            KeyId          = Get-Prop $ac.Cred 'keyId'
            Hint           = $ac.Detail['Hint']
            Thumbprint     = $ac.Detail['Thumbprint']
            Issuer         = ''
            Subject        = ''
            StartDate      = Format-Date $start
            EndDate        = Format-Date $endDt
            Status         = $status
        })
    }
    foreach ($fc in @($batch["fic|$($sp.id)"])) {
        if ($null -eq $fc) { continue }
        if (-not $activeTypes.Contains('FederatedCredential')) { $activeTypes.Add('FederatedCredential') }
        $credRows.Add([pscustomobject]@{
            AppId = $appId; AppName = $sp.displayName; CredentialType = 'FederatedCredential'; Source = 'AppRegistration'
            Name = $fc.name; KeyId = $fc.id; Hint = ''; Thumbprint = ''; Issuer = $fc.issuer; Subject = $fc.subject
            StartDate = ''; EndDate = ''; Status = 'Active'
        })
    }
    if ($sp.servicePrincipalType -eq 'ManagedIdentity') {
        $activeTypes.Add('ManagedIdentity')
        $credRows.Add([pscustomobject]@{
            AppId = $appId; AppName = $sp.displayName; CredentialType = 'ManagedIdentity'; Source = 'Azure'
            Name = ''; KeyId = ''; Hint = ''; Thumbprint = ''; Issuer = ''; Subject = ''
            StartDate = ''; EndDate = ''; Status = 'Active'
        })
    }

    $authMethods = if ($activeTypes.Count) { $activeTypes -join ', ' }
                   elseif ($isMicrosoft) { 'Microsoft-managed' }
                   elseif (-not $app) { 'Unknown (registered in another tenant)' }
                   else { 'None' }

    # ---- Sign-in ----
    $si = $signInByAppId[$appId]

    $appRows.Add([pscustomobject]@{
        AppId              = $appId
        ServicePrincipalId = $sp.id
        AppName            = $sp.displayName
        AppType            = $sp.servicePrincipalType
        IsMicrosoftApp     = $isMicrosoft
        RegisteredInTenant = [bool]$app
        Publisher          = Get-Prop $sp 'publisherName'
        Enabled            = $sp.accountEnabled
        AppCreated         = if ($app) { Format-Date (Get-Prop $app 'createdDateTime') } else { '' }
        AuthMethods        = $authMethods
        HasClientSecret    = $activeTypes.Contains('ClientSecret')
        PermissionCount    = $permCount
        OwnerCount         = $ownerCount
        LastSignIn         = if ($si) { Format-Date (Get-Prop (Get-Prop $si 'lastSignInActivity') 'lastSignInDateTime') } else { '' }
        LastAppOnlySignIn  = if ($si) { Format-Date (Get-Prop (Get-Prop $si 'applicationAuthenticationClientSignInActivity') 'lastSignInDateTime') } else { '' }
        SnapshotDate       = $SnapshotDate
    })
}
#endregion

#region 5. Export ---------------------------------------------------------------------------------
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Export-Table -Rows ($appRows | Sort-Object AppName) -Path (Join-Path $OutputFolder 'Apps.csv') -Columns @(
    'AppId', 'ServicePrincipalId', 'AppName', 'AppType', 'IsMicrosoftApp', 'RegisteredInTenant', 'Publisher', 'Enabled',
    'AppCreated', 'AuthMethods', 'HasClientSecret', 'PermissionCount', 'OwnerCount', 'LastSignIn', 'LastAppOnlySignIn', 'SnapshotDate')
Export-Table -Rows ($permRows | Sort-Object AppName, Resource, Permission) -Path (Join-Path $OutputFolder 'AppPermissions.csv') -Columns @(
    'AppId', 'AppName', 'Resource', 'Permission', 'PermissionDescription', 'GrantedDate')
Export-Table -Rows ($credRows | Sort-Object AppName, CredentialType, StartDate) -Path (Join-Path $OutputFolder 'AppCredentials.csv') -Columns @(
    'AppId', 'AppName', 'CredentialType', 'Source', 'Name', 'KeyId', 'Hint', 'Thumbprint', 'Issuer', 'Subject', 'StartDate', 'EndDate', 'Status')
Export-Table -Rows ($ownerRows | Sort-Object AppName, OwnerName) -Path (Join-Path $OutputFolder 'AppOwners.csv') -Columns @(
    'AppId', 'AppName', 'OwnerOf', 'OwnerName', 'OwnerUPN', 'OwnerType', 'OwnerId')

Write-Host ''
Write-Host '================ Summary ================' -ForegroundColor Green
Write-Host ("Apps                    : {0}" -f $appRows.Count)
Write-Host ("Apps with permissions   : {0}" -f @($appRows | Where-Object { $_.PermissionCount -gt 0 }).Count)
Write-Host ("Apps using client secret: {0}" -f @($appRows | Where-Object HasClientSecret).Count)
Write-Host ("Permission rows         : {0}" -f $permRows.Count)
Write-Host ("Credential rows         : {0}" -f $credRows.Count)
Write-Host ("Owner rows              : {0}" -f $ownerRows.Count)
Write-Host "Files written to $OutputFolder" -ForegroundColor Green
#endregion
