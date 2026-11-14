# Power BI report setup: Entra App Report

Step-by-step instructions for building the Power BI report from the four CSV files produced by `Export-EntraAppReport.ps1`:

| File | One row per |
|---|---|
| `Apps.csv` | enterprise app |
| `AppPermissions.csv` | app + permission (Application / Delegated / None) |
| `AppCredentials.csv` | client secret, certificate, federated credential or managed identity |
| `AppOwners.csv` | app + owner |

The four files are linked by the `AppId` column.

These steps assume the files are in a local or OneDrive folder for now. Moving to SharePoint later only changes one query ([step 8](#8-switch-to-sharepoint-later)).

---

## 1. Connect to the folder

1. In Power BI Desktop, go to **Get Data → Folder** and select the folder that holds the four CSVs.
2. In the dialog that lists the files, click **Transform Data**.
   - **Don't click Combine.** It stacks all four files into one table, which only works when the files have the same columns.
   - **Don't click Load.** That loads only the list of files, not their contents.
3. In Power Query, rename the new query to **`Files`** (right-click → Rename).
4. Right-click `Files` and untick **Enable load**. It is a source for the other queries, not a report table.

---

## 2. Add a function to read one file

Go to **Home → New Source → Blank Query → Advanced Editor**, paste the code below, and rename the query to **`fnReadCsv`**.

```m
(FileName as text) as table =>
let
    Content  = Files{[Name = FileName]}[Content],
    // 1200 = UTF-16, which is what the script writes (-Encoding unicode)
    Csv      = Csv.Document(Content, [Delimiter = ",", Encoding = 1200, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),
    // Empty cells -> null, so date and number columns convert cleanly
    Cleaned  = Table.ReplaceValue(Promoted, "", null, Replacer.ReplaceValue, Table.ColumnNames(Promoted))
in
    Cleaned
```

---

## 3. Create the four tables

Create four more blank queries, one for each table. Give each the name shown, because the DAX in later steps uses these names.

### `Apps`
```m
let
    Source = fnReadCsv("Apps.csv"),
    Typed  = Table.TransformColumnTypes(Source, {
        {"AppId", type text}, {"ServicePrincipalId", type text}, {"AppName", type text}, {"AppType", type text},
        {"IsMicrosoftApp", type logical}, {"RegisteredInTenant", type logical}, {"Publisher", type text},
        {"Enabled", type logical}, {"AppCreated", type date}, {"AuthMethods", type text},
        {"HasClientSecret", type logical}, {"ApplicationPermissionCount", Int64.Type},
        {"DelegatedPermissionCount", Int64.Type}, {"OwnerCount", Int64.Type},
        {"LastSignIn", type date}, {"LastAppOnlySignIn", type date}, {"SnapshotDate", type date}}, "en-US")
in
    Typed
```

### `AppPermissions`
```m
let
    Source = fnReadCsv("AppPermissions.csv"),
    Typed  = Table.TransformColumnTypes(Source, {
        {"AppId", type text}, {"AppName", type text}, {"PermissionType", type text}, {"Resource", type text},
        {"Permission", type text}, {"PermissionDescription", type text}, {"ConsentType", type text},
        {"ConsentedUserCount", Int64.Type}, {"GrantedDate", type date}}, "en-US")
in
    Typed
```

### `AppCredentials`
```m
let
    Source = fnReadCsv("AppCredentials.csv"),
    Typed  = Table.TransformColumnTypes(Source, {
        {"AppId", type text}, {"AppName", type text}, {"CredentialType", type text}, {"Source", type text},
        {"Name", type text}, {"KeyId", type text}, {"Hint", type text}, {"Thumbprint", type text},
        {"Issuer", type text}, {"Subject", type text}, {"StartDate", type date}, {"EndDate", type date},
        {"Status", type text}}, "en-US")
in
    Typed
```

### `AppOwners`
```m
let
    Source = fnReadCsv("AppOwners.csv"),
    Typed  = Table.TransformColumnTypes(Source, {
        {"AppId", type text}, {"AppName", type text}, {"OwnerOf", type text}, {"OwnerName", type text},
        {"OwnerUPN", type text}, {"OwnerType", type text}, {"OwnerId", type text}}, "en-US")
in
    Typed
```

### Check before applying

Click `Apps` and look at the preview:

| You see | What to do |
|---|---|
| Proper column names and readable text | Nothing. Continue. |
| Garbled text, or everything in one column | The files aren't UTF-16. In `fnReadCsv`, change `Encoding = 1200` to `Encoding = 65001`. |
| An error in a date or number column | Check what the cell contains in the CSV. The type list may need adjusting for that column. |

Your query list should look like this, with only the last four loading into the report:

```
Files           (load disabled)
fnReadCsv       (function)
Apps
AppPermissions
AppCredentials
AppOwners
```

Click **Close & Apply**.

---

## 4. Data model

In **Model view**:

1. **Delete any relationships Power BI created automatically.** It may have linked the tables on `AppName`, which isn't unique, because two apps can share a name.
2. **Create three relationships:**

   | From | To | Cardinality | Cross-filter direction |
   |---|---|---|---|
   | `Apps[AppId]` | `AppPermissions[AppId]` | One to many | **Both** |
   | `Apps[AppId]` | `AppCredentials[AppId]` | One to many | **Both** |
   | `Apps[AppId]` | `AppOwners[AppId]` | One to many | **Both** |

   Setting the direction to **Both** is what lets a Permission or Owner slicer filter the app list.
3. **Hide `AppId` and `AppName` in the three child tables** (right-click → Hide in report view). Visuals then always use `Apps[AppName]` and `Apps[AppId]`.

```
              ┌──────────────────┐
              │      Apps        │
              │  (1 row per app) │
              └────────┬─────────┘
          AppId 1 : *  │  (cross-filter: Both)
     ┌─────────────────┼──────────────────┐
     ▼                 ▼                  ▼
AppPermissions   AppCredentials       AppOwners
```

---

## 5. Calculated columns

In Table view, select the table, choose **New column**, and paste each formula.

### In `AppCredentials`

```DAX
Age (days) =
IF ( NOT ISBLANK ( AppCredentials[StartDate] ), INT ( TODAY () - AppCredentials[StartDate] ) )
```

```DAX
Days to Expiry =
IF ( NOT ISBLANK ( AppCredentials[EndDate] ), INT ( AppCredentials[EndDate] - TODAY () ) )
```

```DAX
Expiry Status =
VAR e = AppCredentials[EndDate]
RETURN
    SWITCH (
        TRUE (),
        AppCredentials[CredentialType] IN { "FederatedCredential", "ManagedIdentity" }, "5 - No expiry",
        ISBLANK ( e ), "5 - No expiry",
        e < TODAY (), "1 - Expired",
        e <= TODAY () + 30, "2 - Expires within 30 days",
        e <= TODAY () + 90, "3 - Expires within 90 days",
        "4 - Valid"
    )
```

### In `Apps`

```DAX
Sign-in Status =
VAR d = Apps[LastSignIn]
RETURN
    SWITCH (
        TRUE (),
        ISBLANK ( d ), "4 - No sign-in recorded",
        d >= TODAY () - 30, "1 - Signed in within 30 days",
        d >= TODAY () - 90, "2 - Signed in 31-90 days ago",
        "3 - No sign-in for 90+ days"
    )
```

The number prefixes make these values sort in a sensible order in charts and slicers.

Calculated columns that use `TODAY()` are worked out at **refresh time**, so refresh after loading new files. The measures in the next step are always current.

---

## 6. Measures

Create a table to hold the measures: **Home → Enter data**, leave it empty, name it `_Measures`, then **Load**. Select `_Measures` and add each measure with **New measure**. Once it has measures, you can delete its empty column.

### Counts (for cards)

```DAX
Apps = DISTINCTCOUNT ( Apps[AppId] )
```

```DAX
Apps with Client Secret = CALCULATE ( [Apps], Apps[HasClientSecret] = TRUE () )
```

```DAX
Apps without Owner = CALCULATE ( [Apps], Apps[OwnerCount] = 0 )
```

```DAX
Apps with Application Permissions = CALCULATE ( [Apps], Apps[ApplicationPermissionCount] > 0 )
```

```DAX
Expired Credentials =
CALCULATE (
    COUNTROWS ( AppCredentials ),
    AppCredentials[CredentialType] IN { "ClientSecret", "Certificate" },
    AppCredentials[EndDate] < TODAY ()
)
```

```DAX
Credentials Expiring in 30 Days =
CALCULATE (
    COUNTROWS ( AppCredentials ),
    AppCredentials[CredentialType] IN { "ClientSecret", "Certificate" },
    AppCredentials[EndDate] >= TODAY (),
    AppCredentials[EndDate] <= TODAY () + 30
)
```

```DAX
Secrets Older Than 1 Year =
CALCULATE (
    COUNTROWS ( AppCredentials ),
    AppCredentials[CredentialType] = "ClientSecret",
    AppCredentials[EndDate] >= TODAY (),
    AppCredentials[StartDate] <= TODAY () - 365
)
```

### Per-app values (for the main table)

These show an app's **full** list even when a Permission or Owner slicer is active, because `REMOVEFILTERS` ignores the slicer for that cell. The slicer still decides which apps appear.

```DAX
Application Permissions =
CALCULATE (
    CONCATENATEX ( VALUES ( AppPermissions[Permission] ), AppPermissions[Permission], "; ", AppPermissions[Permission], ASC ),
    REMOVEFILTERS ( AppPermissions ),
    AppPermissions[PermissionType] = "Application"
)
```

```DAX
Delegated Permissions =
CALCULATE (
    CONCATENATEX ( VALUES ( AppPermissions[Permission] ), AppPermissions[Permission], "; ", AppPermissions[Permission], ASC ),
    REMOVEFILTERS ( AppPermissions ),
    AppPermissions[PermissionType] = "Delegated"
)
```

```DAX
Owners =
CALCULATE (
    CONCATENATEX ( VALUES ( AppOwners[OwnerName] ), AppOwners[OwnerName], "; ", AppOwners[OwnerName], ASC ),
    REMOVEFILTERS ( AppOwners )
)
```

```DAX
Oldest Active Secret (days) =
VAR firstStart =
    CALCULATE (
        MIN ( AppCredentials[StartDate] ),
        REMOVEFILTERS ( AppCredentials ),
        AppCredentials[CredentialType] = "ClientSecret",
        AppCredentials[EndDate] >= TODAY ()
    )
RETURN IF ( NOT ISBLANK ( firstStart ), INT ( TODAY () - firstStart ) )
```

```DAX
Next Credential Expiry =
CALCULATE (
    MIN ( AppCredentials[EndDate] ),
    REMOVEFILTERS ( AppCredentials ),
    AppCredentials[CredentialType] IN { "ClientSecret", "Certificate" },
    AppCredentials[EndDate] >= TODAY ()
)
```

```DAX
Days Since Last Sign-in =
VAR d = MAX ( Apps[LastSignIn] )
RETURN IF ( NOT ISBLANK ( d ), INT ( TODAY () - d ) )
```

---

## 7. Report pages

### Report-wide filter

In the Filters pane, under **Filters on all pages**, add `Apps[IsMicrosoftApp]` and select **False**. This hides Microsoft's built-in apps everywhere. Clear it when you want to include them.

### Page 1: Overview

| Visual | Fields |
|---|---|
| 6 cards | `Apps`, `Apps with Client Secret`, `Apps without Owner`, `Secrets Older Than 1 Year`, `Credentials Expiring in 30 Days`, `Expired Credentials` |
| Bar chart: auth methods | Y: `AppCredentials[CredentialType]`, X: `Apps` |
| Bar chart: top permissions | Y: `AppPermissions[Permission]`, X: `Apps`. Visual filters: `PermissionType = Application`, and Top N = 15 by `Apps` |
| Bar chart: credential expiry | Y: `AppCredentials[Expiry Status]`, X: count of `AppCredentials[KeyId]` |
| Donut: sign-in activity | Legend: `Apps[Sign-in Status]`, Values: `Apps` |

### Page 2: App Inventory (main page for filtering)

| Visual | Fields |
|---|---|
| Slicers (dropdown, with **search** turned on) | `AppPermissions[Permission]`, `AppPermissions[PermissionType]`, `AppPermissions[Resource]`, `AppOwners[OwnerName]`, `AppCredentials[CredentialType]`, `Apps[HasClientSecret]`, `Apps[Sign-in Status]` |
| Table | `Apps[AppName]`, `Apps[AppId]`, `Apps[AuthMethods]`, `Owners`, `Application Permissions`, `Delegated Permissions`, `Oldest Active Secret (days)`, `Next Credential Expiry`, `Apps[LastSignIn]` |

- **Keep `AppId` in the table.** Without it, two apps with the same name merge into one row. You can make the column narrow.
- **Conditional formatting** (column → Conditional formatting → Background colour → Rules):
  - `Oldest Active Secret (days)`: red over 365, amber 181–365.
  - `Next Credential Expiry`: red when within 30 days of today.
- **Find apps with no owner:** select `(No owner)` in the Owner slicer. Selecting `(No permissions)` in the Permission slicer works the same way.

### Page 3: Credentials

| Visual | Fields |
|---|---|
| Slicers | `AppCredentials[CredentialType]`, `AppCredentials[Source]`, `AppCredentials[Expiry Status]` |
| Table | `Apps[AppName]`, `AppCredentials[CredentialType]`, `AppCredentials[Source]`, `AppCredentials[Name]`, `AppCredentials[StartDate]`, `AppCredentials[EndDate]`, `AppCredentials[Age (days)]`, `AppCredentials[Days to Expiry]`, `AppCredentials[Expiry Status]` |

Sort by `Days to Expiry` ascending to see the most urgent credentials first. `Source = EnterpriseApp` shows the secrets that don't appear in the App registrations blade.

### Page 4: Owners

| Visual | Fields |
|---|---|
| Bar chart: apps per owner | Y: `AppOwners[OwnerName]`, X: `Apps`, sorted descending |
| Table | `AppOwners[OwnerName]`, `AppOwners[OwnerUPN]`, `AppOwners[OwnerOf]`, `Apps[AppName]`, `Apps[AuthMethods]` |

### Page 5: App Detail (drill-through)

1. Add `Apps[AppName]` and `Apps[AppId]` to this page's **Drill through** area in the Visualizations pane.
2. Add these visuals:
   - **Cards:** `Apps[AppName]`, `Apps[AppId]`, `Apps[LastSignIn]`, `Apps[AuthMethods]`
   - **Permissions table:** `AppPermissions[PermissionType]`, `[Resource]`, `[Permission]`, `[PermissionDescription]`, `[ConsentType]`, `[ConsentedUserCount]`
   - **Credentials table:** the same columns as page 3
   - **Owners table:** `AppOwners[OwnerName]`, `[OwnerUPN]`, `[OwnerOf]`

On any page, right-click an app → **Drill through → App Detail** to see everything about that app.

### Sync slicers

Go to **View → Sync slicers** and sync the Owner and Permission slicers across pages 2–4, so a selection carries over when you switch pages.

---

## 8. Refreshing and switching to SharePoint later

### Refresh with new local files

Overwrite the four CSVs in the same folder, keeping the same names, and click **Refresh**. The `Files` query re-reads the folder each time.

### Switch to SharePoint

When the pipeline uploads the files to SharePoint, replace the `Files` query's code with:

```m
let
    Source = SharePoint.Files("https://<tenant>.sharepoint.com/sites/<SiteName>", [ApiVersion = 15]),
    // Only the base folder; the week folders contain files with the same names
    Files  = Table.SelectRows(Source, each Text.EndsWith([Folder Path], "/<BaseFolderName>/"))
in
    Files
```

The pipeline's upload step (`Export-CSV.ps1 -RemoveBuildID`) only strips the build ID, so in SharePoint the files are named `EntraApp_Apps.csv`, `EntraApp_Permissions.csv`, `EntraApp_Credentials.csv` and `EntraApp_Owners.csv`. Update the file name in each of the four table queries to match.

### Publish and schedule

1. **Home → Publish** to your Power BI workspace.
2. In the Power BI Service, open the semantic model's **Settings → Data source credentials** and sign in to SharePoint with your organisational account (OAuth2).
3. Set **Scheduled refresh** for a time after the pipeline's run, for example pipeline at 05:00 and refresh at 06:00.

For local or OneDrive-synced files, scheduled refresh in the Service needs an on-premises data gateway. Moving to SharePoint avoids that.

---

## Troubleshooting

| Problem | Likely cause and fix |
|---|---|
| `Expression.Error: The key didn't match any rows in the table` | A file name in one of the four queries doesn't match the file in the folder. Check spelling and the `.csv` extension. |
| Text is garbled or all in one column | Wrong encoding. Change `Encoding = 1200` to `65001` in `fnReadCsv`. |
| A Permission or Owner slicer doesn't filter the app table | That relationship's cross-filter direction isn't **Both** (step 4). |
| Two apps appear as one row | `Apps[AppId]` is missing from the table visual. |
| `LastSignIn` is empty for every app | The export had no sign-in data (needs Entra ID P1/P2 and `AuditLog.Read.All`). `Sign-in Status` then shows "No sign-in recorded" for all apps. |
| Ages or expiry status look a day out of date | Calculated columns update on refresh. Refresh the report. |
