# Export-EntraAppReport.ps1

A PowerShell script that exports every Entra ID enterprise app to **four CSV files** for a Power BI report. The files cover the apps, the application permissions on them, the auth methods (credentials) they use, their owners, and their last sign-in dates.

## Requirements

- PowerShell 7+
- `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`
- Read-only Graph permissions: `Application.Read.All`, `Directory.Read.All`, `AuditLog.Read.All`
- Last sign-in dates need Entra ID P1 or P2. Without it, those columns are left empty.

## Usage

```powershell
# Interactive sign-in; writes to .\EntraAppReport
.\Export-EntraAppReport.ps1

# Write to a OneDrive/SharePoint-synced folder that Power BI reads from
.\Export-EntraAppReport.ps1 -OutputFolder 'C:\Users\me\OneDrive - Contoso\EntraAppReport'

# Unattended (app-only, certificate auth)
Connect-MgGraph -TenantId <tenantId> -ClientId <appId> -CertificateThumbprint <thumbprint> -NoWelcome
.\Export-EntraAppReport.ps1 -SkipConnect -OutputFolder D:\Reports\EntraAppReport
```

The file names are fixed and overwritten on every run, so Power BI refresh keeps working.

## Output files

All four files share the `AppId` column.

**Apps.csv**: one row per enterprise app

`AppId, ServicePrincipalId, AppName, AppType, IsMicrosoftApp, RegisteredInTenant, Publisher, Enabled, AppCreated, AuthMethods, HasClientSecret, ApplicationPermissionCount, DelegatedPermissionCount, OwnerCount, LastSignIn, LastAppOnlySignIn, SnapshotDate`

**AppPermissions.csv**: one row per app and permission, for all APIs (Microsoft Graph, Exchange Online, SharePoint and others)

`AppId, AppName, PermissionType, Resource, Permission, PermissionDescription, ConsentType, ConsentedUserCount, GrantedDate`

- `PermissionType` is `Application`, `Delegated`, or `None`. Apps with no permissions at all get a single `(No permissions)` row, so every app appears in this file.
- `ConsentType` is `Admin` for application permissions. For delegated permissions it is `AllUsers` (admin consent for the whole tenant) or `User` (individual users consented); for `User`, `ConsentedUserCount` shows how many users.
- `GrantedDate` is filled only for application permissions, because Graph doesn't record when delegated permissions were granted.
- Only granted permissions are included. Permissions requested in an app registration but never consented don't appear.

**AppCredentials.csv**: one row per client secret, certificate, federated credential or managed identity

`AppId, AppName, CredentialType, Source, Name, KeyId, Hint, Thumbprint, Issuer, Subject, StartDate, EndDate, Status`

- `Source` is `AppRegistration` or `EnterpriseApp`. Secrets on the enterprise app don't appear in the App registrations blade.
- `Status` is `Active`, `Expired` or `NotYetValid`.
- SAML token-signing certificates are excluded because they don't authenticate the app.

**AppOwners.csv**: one row per app and owner

`AppId, AppName, OwnerOf, OwnerName, OwnerUPN, OwnerType, OwnerId`

- `OwnerOf` is `AppRegistration` or `EnterpriseApp`.
- Apps without owners get a single `(No owner)` row, so you can select them in a slicer.

Dates are written as `yyyy-MM-dd`.

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
