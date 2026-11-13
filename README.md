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

`AppId, ServicePrincipalId, AppName, AppType, IsMicrosoftApp, RegisteredInTenant, Publisher, Enabled, AppCreated, AuthMethods, HasClientSecret, PermissionCount, OwnerCount, LastSignIn, LastAppOnlySignIn, SnapshotDate`

**AppPermissions.csv**: one row per app and application permission, for all APIs (Microsoft Graph, Exchange Online, SharePoint and others)

`AppId, AppName, Resource, Permission, PermissionDescription, GrantedDate`

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

## Notes

- `AuthMethods` in Apps.csv lists the methods with at least one non-expired credential. For apps registered in another tenant (third-party or multi-tenant apps), credentials can't be seen from your tenant, so it shows `Unknown (registered in another tenant)`.
- Permissions, owners and federated credentials are read with Graph JSON batching, 20 requests per call. A tenant with a few thousand apps should take a few minutes.
