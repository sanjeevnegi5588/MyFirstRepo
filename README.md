# Export-CriticalGraphApps.ps1

A PowerShell script that exports Entra ID enterprise apps holding **critical or high-risk Microsoft Graph application permissions** to a single CSV file. For each app it shows the owners, the permissions, which auth methods the app uses, the age and expiry of every credential, and the app's last sign-in date.

## Requirements

- PowerShell 7+
- `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`
- Read-only Graph permissions: `Application.Read.All`, `Directory.Read.All`, `AuditLog.Read.All`
- The last sign-in date needs Entra ID P1 or P2. Without it, the column shows `Unavailable` and the rest of the report still works.

## Usage

```powershell
# Interactive sign-in
.\Export-CriticalGraphApps.ps1

# Choose where the file goes
.\Export-CriticalGraphApps.ps1 -OutputPath C:\Reports\GraphApps.csv

# Unattended (app-only, certificate auth)
Connect-MgGraph -TenantId <tenantId> -ClientId <appId> -CertificateThumbprint <thumbprint> -NoWelcome
.\Export-CriticalGraphApps.ps1 -SkipConnect
```

Optional:
- `-AdditionalCriticalPermissions 'AuditLog.Read.All','Reports.Read.All'` treats extra permissions as critical.
- `-IncludeNonCritical` exports every app with any Graph application permission.

## Output

One CSV with one row per app, sorted by risk:

| Column | Example |
|---|---|
| RiskRating | `Critical` |
| AppName | `Legacy-UserProvisioning` |
| AppId | `1a2b3c4d-...` |
| Owners | `it-admin@example.com; jane@example.com` (or `No owners`) |
| CriticalPermissions | `Group.ReadWrite.All; User.ReadWrite.All` |
| HighPermissions | `Mail.Read` |
| AuthMethods | `ClientSecret, Certificate` |
| CredentialAges | `ClientSecret 'prod-2024': 742 days old, expires 2026-12-01 (17 days left) \| Certificate 'backup-cert': 120 days old, expires 2027-07-15 (243 days left)` |
| LastSignIn | `2026-11-12` |

**RiskRating**
- **Critical**: a critical permission **and** an active client secret.
- **High**: a critical permission, or a high-risk permission plus a client secret.
- **Medium**: a high-risk permission only.

**CredentialAges** lists every client secret, certificate and federated credential, oldest first, including expired ones (marked `EXPIRED`). A credential stored on the enterprise app instead of the app registration is marked `[on enterprise app]`, because it doesn't appear in the App registrations blade.

## Notes

- The critical and high permission lists are in the `$PermissionTiers` hashtable at the top of the script. Edit them to fit your own risk model.
- SAML token-signing certificates are excluded because they aren't used to authenticate the app.
- Third-party or multi-tenant apps registered in another tenant show `Unknown (app registered in another tenant)`, because their credentials can't be seen from your tenant.
