# Export-EntraAppReport.ps1

A PowerShell script that exports every Entra ID enterprise app to **four CSV files** for a Power BI report. The files cover the apps, the application permissions on them, the auth methods (credentials) they use, their owners, and their last sign-in dates.

## Requirements

The script runs in an Azure DevOps pipeline, following the same pattern as the team's other IAG scripts.

- An **Azure PowerShell** task (PowerShell 7) signed in through a service connection. The script gets a Graph token from that context with `Get-AzAccessToken` and connects with `Connect-MgGraph -AccessToken`.
- Graph **application** permissions on the service connection's identity, all read-only: `Application.Read.All`, `Directory.Read.All` and `AuditLog.Read.All`.
- Last sign-in dates need Entra ID P1 or P2. Without it, those columns are left empty and a warning is logged.
- These modules:
  - `Exceptions` and `Logging` from the repo's `.\IAG\Generic\Modules` folder, which is added to `PSModulePath`.
  - `Microsoft.Graph.Authentication`.

## Usage

Run the script from the repository root, so that `.\IAG\Generic\Modules` resolves:

```yaml
- task: AzurePowerShell@5
  inputs:
    azureSubscription: '<service connection>'
    ScriptPath: 'Export-EntraAppReport.ps1'
    azurePowerShellVersion: LatestVersion
    pwsh: true
```

The four CSV files are written to `$(Build.SourcesDirectory)` with the build ID in the name. Their paths are set as pipeline variables for later steps, for example to upload them to the SharePoint folder that Power BI reads:

| File | Pipeline variable |
|---|---|
| `EntraApp_Apps<BuildId>.csv` | `$(appscsvpath)` |
| `EntraApp_Permissions<BuildId>.csv` | `$(permissionscsvpath)` |
| `EntraApp_Credentials<BuildId>.csv` | `$(credentialscsvpath)` |
| `EntraApp_Owners<BuildId>.csv` | `$(ownerscsvpath)` |

Power BI refresh needs the same file names on every run. When uploading, save the files under fixed names: `Apps.csv`, `AppPermissions.csv`, `AppCredentials.csv` and `AppOwners.csv`.

Logging uses the `AzureDevOpsConsole` target. It logs at INFO level on `master` and at DEBUG level on other branches.

## Output files

All four files share the `AppId` column.

**Apps.csv** (`EntraApp_Apps<BuildId>.csv`): one row per enterprise app

`AppId, ServicePrincipalId, AppName, AppType, IsMicrosoftApp, RegisteredInTenant, Publisher, Enabled, AppCreated, AuthMethods, HasClientSecret, ApplicationPermissionCount, DelegatedPermissionCount, OwnerCount, LastSignIn, LastAppOnlySignIn, SnapshotDate`

**AppPermissions.csv** (`EntraApp_Permissions<BuildId>.csv`): one row per app and permission, for all APIs (Microsoft Graph, Exchange Online, SharePoint and others)

`AppId, AppName, PermissionType, Resource, Permission, PermissionDescription, ConsentType, ConsentedUserCount, GrantedDate`

- `PermissionType` is `Application`, `Delegated`, or `None`. Apps with no permissions at all get a single `(No permissions)` row, so every app appears in this file.
- `ConsentType` is `Admin` for application permissions. For delegated permissions it is `AllUsers` (admin consent for the whole tenant) or `User` (individual users consented); for `User`, `ConsentedUserCount` shows how many users.
- `GrantedDate` is filled only for application permissions, because Graph doesn't record when delegated permissions were granted.
- Only granted permissions are included. Permissions requested in an app registration but never consented don't appear.

**AppCredentials.csv** (`EntraApp_Credentials<BuildId>.csv`): one row per client secret, certificate, federated credential or managed identity

`AppId, AppName, CredentialType, Source, Name, KeyId, Hint, Thumbprint, Issuer, Subject, StartDate, EndDate, Status`

- `Source` is `AppRegistration` or `EnterpriseApp`. Secrets on the enterprise app don't appear in the App registrations blade.
- `Status` is `Active`, `Expired` or `NotYetValid`.
- SAML token-signing certificates are excluded because they don't authenticate the app.

**AppOwners.csv** (`EntraApp_Owners<BuildId>.csv`): one row per app and owner

`AppId, AppName, OwnerOf, OwnerName, OwnerUPN, OwnerType, OwnerId`

- `OwnerOf` is `AppRegistration` or `EnterpriseApp`.
- Apps without owners get a single `(No owner)` row, so you can select them in a slicer.

Dates are written as `yyyy-MM-dd`. Files are UTF-16 (`-Encoding unicode`), the same as the MFA report; Power BI reads them without changes.

## Power BI setup

1. **Get Data.** Load the four CSVs, using the Folder or SharePoint folder connector.
2. **Model view.** Create a relationship from `Apps[AppId]` to `AppId` in each of the other three tables, one-to-many, with cross-filter direction set to **Both**. Slicers on Permission or Owner then filter the app list.
3. **Measures.** Calculate ages in Power BI so they stay current, for example:

```DAX
Permissions = CONCATENATEX ( VALUES ( AppPermissions[Permission] ), AppPermissions[Permission], "; " )
Owners      = CONCATENATEX ( VALUES ( AppOwners[OwnerName] ), AppOwners[OwnerName], "; " )

Oldest Active Secret (days) =
VAR firstStart =
    CALCULATE ( MIN ( AppCredentials[StartDate] ),
        AppCredentials[CredentialType] = "ClientSecret", AppCredentials[Status] = "Active" )
RETURN IF ( NOT ISBLANK ( firstStart ), DATEDIFF ( firstStart, TODAY (), DAY ) )

Secrets Expiring 30d =
CALCULATE ( COUNTROWS ( AppCredentials ),
    AppCredentials[CredentialType] = "ClientSecret",
    AppCredentials[EndDate] >= TODAY (), AppCredentials[EndDate] <= TODAY () + 30 )
```

Tip: filter the report on `IsMicrosoftApp = FALSE` by default to hide Microsoft's own built-in apps.

Add a slicer on `AppPermissions[PermissionType]` to switch between application and delegated permissions. To count real permissions, exclude the placeholder rows, for example `CALCULATE ( COUNTROWS ( AppPermissions ), AppPermissions[PermissionType] <> "None" )`.

## Notes

- `AuthMethods` in Apps.csv lists the methods with at least one non-expired credential. For apps registered in another tenant (third-party or multi-tenant apps), credentials can't be seen from your tenant, so it shows `Unknown (registered in another tenant)`.
- Permissions, owners and federated credentials are read with Graph JSON batching, 20 requests per call. A tenant with a few thousand apps should take a few minutes.
