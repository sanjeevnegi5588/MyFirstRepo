# Export-CriticalGraphApps.ps1

A PowerShell script that audits Entra ID enterprise apps (service principals) that hold **critical Microsoft Graph API permissions**. For each app it reports whether the app uses a **client secret** to sign in, and how old that secret is.

## Requirements

- PowerShell 7+
- `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`
- Read-only Graph permissions: `Application.Read.All` and `Directory.Read.All`. Add `AuditLog.Read.All` if you use `-IncludeSignInActivity`, which also needs Entra ID P1 or P2.

## Usage

```powershell
# Interactive sign-in, application permissions only
.\Export-CriticalGraphApps.ps1

# Include delegated grants, last sign-in dates, and flag secrets older than 90 days
.\Export-CriticalGraphApps.ps1 -IncludeDelegated -IncludeSignInActivity -SecretAgeThresholdDays 90

# Unattended (app-only, certificate auth)
Connect-MgGraph -TenantId <tenantId> -ClientId <appId> -CertificateThumbprint <thumbprint> -NoWelcome
.\Export-CriticalGraphApps.ps1 -SkipConnect -OutputFolder C:\Reports

# Treat extra permissions as critical, or export every app that has any Graph permission
.\Export-CriticalGraphApps.ps1 -AdditionalCriticalPermissions 'AuditLog.Read.All','Reports.Read.All'
.\Export-CriticalGraphApps.ps1 -IncludeNonCritical
```

## Output (CSV)

| File | One row per | Key columns |
|---|---|---|
| `*_Summary.csv` | enterprise app | RiskRating, CriticalPermissions, AuthMethods, UsesClientSecret, ActiveSecretCount, OldestActiveSecretAgeDays, NextSecretExpiry, Owners, LastSignIn, Findings |
| `*_Secrets.csv` | client secret | CredentialSource (AppRegistration / ServicePrincipal), SecretName, Hint, Created, Expires, AgeDays, DaysToExpiry, LifetimeDays, Status |
| `*_Permissions.csv` | app + Graph permission | Permission, PermissionType (Application/Delegated), ConsentType, Tier, GrantedOn |

**RiskRating**
- **Critical**: has a critical permission **and** an active client secret.
- **High**: has a critical permission, or has a high-tier permission plus a client secret.
- **Medium**: has a high-tier permission only.

**Findings** include: secret older than the threshold, secret lifetime longer than `-MaxSecretLifetimeDays`, more than one active secret, expired secrets left in place, a secret expiring within 30 days, a secret stored on the service principal (it doesn't appear in the App registrations blade), no owners, and a disabled app.

## Notes

- The critical and high permission lists are in the `$PermissionTiers` hashtable at the top of the script. Edit them to fit your own risk model.
- Secrets are read from both the app registration and the service principal. SAML token-signing credentials are excluded because they are not client secrets.
- Multi-tenant or third-party apps that are registered in another tenant show `Unknown (app registered in another tenant)` as the auth method, because their secrets can't be seen from your tenant.
