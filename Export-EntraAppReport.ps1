<#
.SYNOPSIS
    Exports every enterprise application (service principal) in the tenant to four CSV files for
    Power BI: apps, permissions, credentials (auth methods) and owners.

.DESCRIPTION
    Runs in an Azure DevOps pipeline (Azure PowerShell task). Connects to Microsoft Graph with the
    pipeline's Az context token and writes four CSVs to $ENV:BUILD_SOURCESDIRECTORY:

      EntraApp_Apps<BuildId>.csv         one row per app: type, Microsoft/third-party flags, auth methods,
                                         permission and owner counts, last sign-in dates, snapshot date
      EntraApp_Permissions<BuildId>.csv  one row per app + permission (Application / Delegated / None)
      EntraApp_Credentials<BuildId>.csv  one row per client secret / certificate / federated credential /
                                         managed identity, with start and end dates and status
      EntraApp_Owners<BuildId>.csv       one row per app + owner, plus a '(No owner)' row

    The paths are published as pipeline variables: appscsvpath, permissionscsvpath,
    credentialscsvpath, ownerscsvpath.

    All files share the AppId column. In Power BI, relate Apps[AppId] 1-to-many to AppId in the other
    three tables. Dates are written as yyyy-MM-dd; calculate ages in Power BI so they stay current.

    Graph application permissions needed by the pipeline identity (read-only):
    Application.Read.All, Directory.Read.All, AuditLog.Read.All (last sign-in; needs Entra ID P1/P2).
#>
#Requires -Version 7.0

Write-Output "Script started $($MyInvocation.MyCommand.Name)"

# Initial Variables
$ErrorActionPreference = 'Stop'
$Now          = [datetime]::UtcNow
$SnapshotDate = $Now.ToString('yyyy-MM-dd')
$GraphRoot    = 'https://graph.microsoft.com/v1.0'
# Tenants that own Microsoft's first-party apps
$MicrosoftTenantIds = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')

# Repo modules
$env:PSModulePath = "$(Resolve-Path '.\IAG\Generic\Modules');" + $env:PSModulePath

# Import mandatory modules
Import-Module Exceptions
Import-Module Logging
Import-Module Microsoft.Graph.Authentication

# Show Modules
Get-Module | Format-Table ModuleType, Version, @{label = "PreRelease"; expression = { $_.PrivateData.PSData.Prerelease } }, Name

# CSV output variables
$appsCsvPath        = ("{0}\EntraApp_Apps{1}.csv" -f $ENV:BUILD_SOURCESDIRECTORY, $ENV:BUILD_BUILDID)
$permissionsCsvPath = ("{0}\EntraApp_Permissions{1}.csv" -f $ENV:BUILD_SOURCESDIRECTORY, $ENV:BUILD_BUILDID)
$credentialsCsvPath = ("{0}\EntraApp_Credentials{1}.csv" -f $ENV:BUILD_SOURCESDIRECTORY, $ENV:BUILD_BUILDID)
$ownersCsvPath      = ("{0}\EntraApp_Owners{1}.csv" -f $ENV:BUILD_SOURCESDIRECTORY, $ENV:BUILD_BUILDID)
write-host "##vso[task.setvariable variable=appscsvpath]$appsCsvPath"
write-host "##vso[task.setvariable variable=permissionscsvpath]$permissionsCsvPath"
write-host "##vso[task.setvariable variable=credentialscsvpath]$credentialsCsvPath"
write-host "##vso[task.setvariable variable=ownerscsvpath]$ownersCsvPath"

# Enable logging
if ($env:BUILD_SOURCEBRANCHNAME -eq 'master') {
    Add-LoggingTarget -Name AzureDevOpsConsole -Configuration @{Level = 'INFO' }
}
else {
    Add-LoggingTarget -Name AzureDevOpsConsole -Configuration @{Level = 'DEBUG' }
}

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
    param([Parameter(Mandatory)][System.Collections.IList]$Requests)
    $results = @{}
    for ($start = 0; $start -lt $Requests.Count; $start += 20) {
        # Progress counter for Azure DevOps
        $percentComplete = [math]::Max(1, [math]::Round(($start / [math]::Max($Requests.Count, 1)) * 100))
        Write-Host "##vso[task.setprogress value=$percentComplete]Progress"
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
                    Write-Log -Message "Request failed ($($req.Url)): $($_.Exception.Message)" -Level Warning
                    $results[$req.Key] = [System.Collections.Generic.List[object]]::new()
                }
            }
        }
    }
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
        $Rows | Select-Object $Columns | Export-Csv -Path $Path -NoTypeInformation -Encoding unicode
    } else {
        Set-Content -Path $Path -Value (($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') -Encoding unicode
    }
}
#endregion

#Main script
Try {
    Write-Log -Message "Starting main script" -Level Info

    # Connect to MS Graph using powershell
    $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/").Token
    Write-Log -Message "Connecting to Microsoft Graph API" -Level Info
    Connect-MgGraph -AccessToken $token -NoWelcome

    #region 1. Read service principals and app registrations
    Write-LogHeader -Message "Getting enterprise apps and app registrations" -HeaderChar "="
    Write-Log -Message "Getting all enterprise apps (service principals)" -Level Info
    $spSelect = 'id,appId,displayName,servicePrincipalType,accountEnabled,appOwnerOrganizationId,publisherName,passwordCredentials,keyCredentials,appRoles,oauth2PermissionScopes'
    Write-Log -Message "GET $GraphRoot/servicePrincipals?`$select=$spSelect" -Level DEBUG
    $allSps   = Invoke-GraphPaged -Uri "$GraphRoot/servicePrincipals?`$select=$spSelect&`$top=999"
    $spById   = @{}
    foreach ($s in $allSps) { $spById[[string]$s.id] = $s }
    # Only real apps and managed identities (skips legacy and social-IdP placeholder objects)
    $sps = @($allSps | Where-Object { $_.servicePrincipalType -in 'Application', 'ManagedIdentity' })
    Write-Log -Message "Found $($sps.Count) enterprise apps" -Level Info

    Write-Log -Message "Getting all app registrations" -Level Info
    $apps       = Invoke-GraphPaged -Uri "$GraphRoot/applications?`$select=id,appId,createdDateTime,passwordCredentials,keyCredentials&`$top=999"
    $appByAppId = @{}
    foreach ($a in $apps) { $appByAppId[[string]$a.appId] = $a }
    Write-Log -Message "Found $($apps.Count) app registrations" -Level Info
    #endregion

    #region 2. Permissions, owners and federated credentials (batched)
    Write-LogHeader -Message "Getting permissions, owners and federated credentials" -HeaderChar "="
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
    Write-Log -Message "Running $($requests.Count) requests in batches of 20" -Level Info
    $batch = Invoke-GraphBatch -Requests $requests

    # Delegated permissions are stored as OAuth2 permission grants: one per app + API + (all users | one user).
    # Collapse them to one entry per app + API + scope + consent type, counting users for per-user consent.
    Write-Log -Message "Getting delegated permission grants" -Level Info
    $delegatedBySp = @{}   # spId -> ordered dictionary "resourceId|scope|consentType" -> entry
    foreach ($g in (Invoke-GraphPaged -Uri "$GraphRoot/oauth2PermissionGrants")) {
        $clientId = [string]$g.clientId
        $consent  = if ($g.consentType -eq 'AllPrincipals') { 'AllUsers' } else { 'User' }
        foreach ($scope in @(([string](Get-Prop $g 'scope')) -split '\s+' | Where-Object { $_ })) {
            if (-not $delegatedBySp.ContainsKey($clientId)) { $delegatedBySp[$clientId] = [ordered]@{} }
            $key = "$($g.resourceId)|$scope|$consent"
            if (-not $delegatedBySp[$clientId].Contains($key)) {
                $delegatedBySp[$clientId][$key] = @{ ResourceId = [string]$g.resourceId; Scope = $scope; ConsentType = $consent; Users = 0 }
            }
            if ($consent -eq 'User') {
                $entry = $delegatedBySp[$clientId][$key]
                $entry['Users'] = $entry['Users'] + 1
            }
        }
    }
    #endregion

    #region 3. Last sign-in dates (needs Entra ID P1/P2)
    Write-LogHeader -Message "Getting service principal sign-in activity" -HeaderChar "="
    $signInByAppId = @{}
    try {
        foreach ($s in (Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities')) {
            $signInByAppId[[string]$s.appId] = $s
        }
        Write-Log -Message "Found sign-in activity for $($signInByAppId.Count) apps" -Level Info
    }
    catch {
        Write-Log -Message "Last sign-in dates unavailable (needs AuditLog.Read.All and Entra ID P1/P2): $($_.Exception.Message)" -Level Warning
    }
    #endregion

    Write-LogHeader -Message "Building report" -HeaderChar "="
    #region 4. Build the four tables
    $appRows   = [System.Collections.Generic.List[object]]::new()
    $permRows  = [System.Collections.Generic.List[object]]::new()
    $credRows  = [System.Collections.Generic.List[object]]::new()
    $ownerRows = [System.Collections.Generic.List[object]]::new()
    $roleNameCache  = @{}   # resourceId -> @{ roleId -> appRole }
    $scopeNameCache = @{}   # resourceId -> @{ scope value -> oauth2PermissionScope }

    foreach ($sp in $sps) {
        $appId       = [string]$sp.appId
        $app         = $appByAppId[$appId]
        $isMicrosoft = $sp.appOwnerOrganizationId -in $MicrosoftTenantIds

        # ---- Application permissions ----
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
                PermissionType        = 'Application'
                Resource              = $pa.resourceDisplayName
                Permission            = if ($role) { $role.value } elseif ([string]$pa.appRoleId -eq '00000000-0000-0000-0000-000000000000') { '(default access)' } else { [string]$pa.appRoleId }
                PermissionDescription = if ($role) { $role.displayName } else { '' }
                ConsentType           = 'Admin'
                ConsentedUserCount    = ''
                GrantedDate           = Format-Date (Get-Prop $pa 'createdDateTime')
            })
            $permCount++
        }

        # ---- Delegated permissions ----
        $delegatedScopes = @{}
        if ($delegatedBySp.ContainsKey([string]$sp.id)) {
            foreach ($d in @($delegatedBySp[[string]$sp.id].Values)) {
                $resSp = $spById[$d.ResourceId]
                if (-not $scopeNameCache.ContainsKey($d.ResourceId)) {
                    $map = @{}
                    foreach ($sc in @(Get-Prop $resSp 'oauth2PermissionScopes')) { if ($sc) { $map[[string]$sc.value] = $sc } }
                    $scopeNameCache[$d.ResourceId] = $map
                }
                $scopeObj = $scopeNameCache[$d.ResourceId][$d.Scope]
                $permRows.Add([pscustomobject]@{
                    AppId                 = $appId
                    AppName               = $sp.displayName
                    PermissionType        = 'Delegated'
                    Resource              = if ($resSp) { $resSp.displayName } else { $d.ResourceId }
                    Permission            = $d.Scope
                    PermissionDescription = if ($scopeObj) { Get-Prop $scopeObj 'adminConsentDisplayName' } else { '' }
                    ConsentType           = $d.ConsentType
                    ConsentedUserCount    = if ($d.ConsentType -eq 'User') { $d.Users } else { '' }
                    GrantedDate           = ''   # Graph does not record when delegated grants were made
                })
                $delegatedScopes["$($d.ResourceId)|$($d.Scope)"] = $true
            }
        }
        $delegatedCount = $delegatedScopes.Count

        # Placeholder so every app appears in AppPermissions.csv and can be picked in a slicer
        if ($permCount -eq 0 -and $delegatedCount -eq 0) {
            $permRows.Add([pscustomobject]@{
                AppId = $appId; AppName = $sp.displayName; PermissionType = 'None'; Resource = ''; Permission = '(No permissions)'
                PermissionDescription = ''; ConsentType = ''; ConsentedUserCount = ''; GrantedDate = ''
            })
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
            ApplicationPermissionCount = $permCount
            DelegatedPermissionCount   = $delegatedCount
            OwnerCount         = $ownerCount
            LastSignIn         = if ($si) { Format-Date (Get-Prop (Get-Prop $si 'lastSignInActivity') 'lastSignInDateTime') } else { '' }
            LastAppOnlySignIn  = if ($si) { Format-Date (Get-Prop (Get-Prop $si 'applicationAuthenticationClientSignInActivity') 'lastSignInDateTime') } else { '' }
            SnapshotDate       = $SnapshotDate
        })
    }
    #endregion

    #region 5. Export
    Write-LogHeader -Message "Exporting results to CSV" -HeaderChar "="
    Export-Table -Rows ($appRows | Sort-Object AppName) -Path $appsCsvPath -Columns @(
        'AppId', 'ServicePrincipalId', 'AppName', 'AppType', 'IsMicrosoftApp', 'RegisteredInTenant', 'Publisher', 'Enabled',
        'AppCreated', 'AuthMethods', 'HasClientSecret', 'ApplicationPermissionCount', 'DelegatedPermissionCount', 'OwnerCount',
        'LastSignIn', 'LastAppOnlySignIn', 'SnapshotDate')
    Write-Log -Message "Apps exported to $appsCsvPath" -Level Info

    Export-Table -Rows ($permRows | Sort-Object AppName, PermissionType, Resource, Permission) -Path $permissionsCsvPath -Columns @(
        'AppId', 'AppName', 'PermissionType', 'Resource', 'Permission', 'PermissionDescription', 'ConsentType', 'ConsentedUserCount', 'GrantedDate')
    Write-Log -Message "Permissions exported to $permissionsCsvPath" -Level Info

    Export-Table -Rows ($credRows | Sort-Object AppName, CredentialType, StartDate) -Path $credentialsCsvPath -Columns @(
        'AppId', 'AppName', 'CredentialType', 'Source', 'Name', 'KeyId', 'Hint', 'Thumbprint', 'Issuer', 'Subject', 'StartDate', 'EndDate', 'Status')
    Write-Log -Message "Credentials exported to $credentialsCsvPath" -Level Info

    Export-Table -Rows ($ownerRows | Sort-Object AppName, OwnerName) -Path $ownersCsvPath -Columns @(
        'AppId', 'AppName', 'OwnerOf', 'OwnerName', 'OwnerUPN', 'OwnerType', 'OwnerId')
    Write-Log -Message "Owners exported to $ownersCsvPath" -Level Info
    #endregion

    # Log app report
    Write-LogHeader -Message "Enterprise App Report" -HeaderChar "="
    Write-Log -Message "Apps: $($appRows.Count)" -Level Info
    Write-Log -Message "Apps with application permissions: $(@($appRows | Where-Object { $_.ApplicationPermissionCount -gt 0 }).Count)" -Level Info
    Write-Log -Message "Apps with delegated permissions: $(@($appRows | Where-Object { $_.DelegatedPermissionCount -gt 0 }).Count)" -Level Info
    Write-Log -Message "Apps using client secrets: $(@($appRows | Where-Object HasClientSecret).Count)" -Level Info
    Write-Log -Message "Auth methods:" -Level Info
    $sortedMethods = $appRows | Group-Object AuthMethods | Sort-Object Count -Descending
    foreach ($method in $sortedMethods) {
        Write-Log -Message "* $($method.Name): $($method.Count)" -Level Info
    }
}
Catch {
    Write-Log -Message "Error: `n {0}" -Arguments ($_.ErrorDetails.Message ?? $_.Exception.Message) -level Error
    Invoke-DefaultDevOpsExceptionHandler
    Exit 1
}
finally {
    Wait-Logging
    Disconnect-Graph | out-null
}

Write-Log -Message "{0}:DONE" -Arguments $($MyInvocation.MyCommand) -Level SUCCESS
