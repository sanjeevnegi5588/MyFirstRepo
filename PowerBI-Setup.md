# Power BI report setup: Entra App Report

Step-by-step instructions for building the Power BI report from the four CSV files produced by `Export-EntraAppReport.ps1`:

| File | One row per |
|---|---|
| `Apps.csv` | enterprise app |
| `AppPermissions.csv` | app + permission (Application / Delegated / None) |
| `AppCredentials.csv` | client secret, certificate, federated credential or managed identity |
| `AppOwners.csv` | app + owner |

The four files are linked by the `AppId` column.

Steps 1–3 load the data. Steps 4–7 build the model and the report pages. Step 8 covers refreshing, switching to SharePoint and publishing.

---

## 1–3. Load the data

Follow **[powerbi/README.md](powerbi/README.md)**. It has the query files and the step-by-step setup, confirmed working in Power BI Desktop. In short:

1. **Start a new Power BI file.** Set **Ignore the Privacy levels** and untick **Autodetect new relationships** (File → Options → Current File).
2. **Create four parameters** that hold the file location: `SourceType`, `LocalFolderPath`, `SharePointSiteUrl` and `SharePointFolderName`.
3. **Create `Files`,** which lists the CSVs from that location (Enable load **off**).
4. **Create `Apps`, `AppPermissions`, `AppCredentials` and `AppOwners`** from the `.pq` files (Enable load **on**).
5. **Click Close & Apply,** then save.

Your query list should then look like this, with only the four tables loading into the report:

```
SourceType            (parameter)
LocalFolderPath       (parameter)
SharePointSiteUrl     (parameter)
SharePointFolderName  (parameter)
Files                 (load disabled)
Apps
AppPermissions
AppCredentials
AppOwners
```

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

### Stop Power BI summing number columns

Power BI adds up number columns by default ("Sum of …"). A Table visual then **hides every row where those sums are blank**, which hides application permissions, `(No permissions)` rows, federated credentials and managed identities. For each of these columns, select it in the Data pane and set **Column tools → Summarization → Don't summarize**:

- `AppPermissions[ConsentedUserCount]`
- `AppCredentials[Age (days)]`
- `AppCredentials[Days to Expiry]`

Calculated columns that use `TODAY()` are worked out at **refresh time**, so refresh after loading new files. The measures in the next step are always current.

---

## 6. Measures

Create a table to hold the measures: **Home → Enter data**, leave it empty, name it `_Measures`, then **Load**. Select `_Measures` and add each measure with **New measure**. Once it has measures, you can delete its empty column.

### Counts (for cards)

The first measure is called `App Count`, not `Apps`. A measure with the same name as the `Apps` table makes `Apps[AppId]` ambiguous, and the formula shows an error.

```DAX
App Count = DISTINCTCOUNT ( Apps[AppId] )
```

```DAX
Apps with Client Secret = CALCULATE ( [App Count], Apps[HasClientSecret] = TRUE () )
```

```DAX
Apps without Owner = CALCULATE ( [App Count], Apps[OwnerCount] = 0 )
```

```DAX
Apps with Application Permissions = CALCULATE ( [App Count], Apps[ApplicationPermissionCount] > 0 )
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

These show an app's **full** list even when a Permission or Owner slicer is active. Each one first remembers the row's app (`CurrentApps`), removes the slicer's filter on the child table, then reapplies the app with `TREATAS`. The slicer still decides which apps appear.

Don't use `REMOVEFILTERS ( <child table> )` on its own: it also removes the filter on the related `Apps` table, so every row shows every app's values.

```DAX
Application Permissions =
VAR CurrentApps = VALUES ( Apps[AppId] )
RETURN
    CALCULATE (
        CONCATENATEX ( VALUES ( AppPermissions[Permission] ), AppPermissions[Permission], "; ", AppPermissions[Permission], ASC ),
        REMOVEFILTERS ( AppPermissions ),
        TREATAS ( CurrentApps, AppPermissions[AppId] ),
        AppPermissions[PermissionType] = "Application"
    )
```

```DAX
Delegated Permissions =
VAR CurrentApps = VALUES ( Apps[AppId] )
RETURN
    CALCULATE (
        CONCATENATEX ( VALUES ( AppPermissions[Permission] ), AppPermissions[Permission], "; ", AppPermissions[Permission], ASC ),
        REMOVEFILTERS ( AppPermissions ),
        TREATAS ( CurrentApps, AppPermissions[AppId] ),
        AppPermissions[PermissionType] = "Delegated"
    )
```

```DAX
Owners =
VAR CurrentApps = VALUES ( Apps[AppId] )
RETURN
    CALCULATE (
        CONCATENATEX ( VALUES ( AppOwners[OwnerName] ), AppOwners[OwnerName], "; ", AppOwners[OwnerName], ASC ),
        REMOVEFILTERS ( AppOwners ),
        TREATAS ( CurrentApps, AppOwners[AppId] )
    )
```

```DAX
Oldest Active Secret (days) =
VAR CurrentApps = VALUES ( Apps[AppId] )
VAR FirstStart =
    CALCULATE (
        MIN ( AppCredentials[StartDate] ),
        REMOVEFILTERS ( AppCredentials ),
        TREATAS ( CurrentApps, AppCredentials[AppId] ),
        AppCredentials[CredentialType] = "ClientSecret",
        AppCredentials[EndDate] >= TODAY ()
    )
RETURN IF ( NOT ISBLANK ( FirstStart ), INT ( TODAY () - FirstStart ) )
```

```DAX
Next Credential Expiry =
VAR CurrentApps = VALUES ( Apps[AppId] )
RETURN
    CALCULATE (
        MIN ( AppCredentials[EndDate] ),
        REMOVEFILTERS ( AppCredentials ),
        TREATAS ( CurrentApps, AppCredentials[AppId] ),
        AppCredentials[CredentialType] IN { "ClientSecret", "Certificate" },
        AppCredentials[EndDate] >= TODAY ()
    )
```

```DAX
Days to Next Expiry =
VAR d = [Next Credential Expiry]
RETURN IF ( NOT ISBLANK ( d ), INT ( d - TODAY () ) )
```

```DAX
Days Since Last Sign-in =
VAR d = MAX ( Apps[LastSignIn] )
RETURN IF ( NOT ISBLANK ( d ), INT ( TODAY () - d ) )
```

### For the App Lookup page

```DAX
Selected App =
IF (
    HASONEVALUE ( Apps[AppId] ),
    SELECTEDVALUE ( Apps[AppName] ),
    "Select an app from the list above"
)
```

```DAX
App Selected = IF ( ISFILTERED ( Apps[AppName] ), 1, 0 )
```

`App Selected` is used as a visual filter (**is 1**) so the App Lookup tables stay empty until someone picks an app.

---

## 7. Report pages

### Report-wide filter

In the Filters pane, under **Filters on all pages**, add `Apps[IsMicrosoftApp]` and select **False**. This hides Microsoft's built-in apps everywhere. Clear it when you want to include them.

In this step, **"Table" means a Table visual** (Visualizations → Table icon) placed on the page, not a new data table. No new data tables are needed.

Each **slicer takes one field**. Putting several fields into one slicer makes Power BI try to build a hierarchy from them, which fails with *"Can't display this visual … ConditionNotHierarchicallyRelated"*. Make one slicer, then copy it (Ctrl+C, Ctrl+V) and swap the field.

### App Lookup page (default view)

The first page readers see: pick an app from a searchable dropdown and see its details, application permissions and owners.

| Visual | Fields and settings |
|---|---|
| **Slicer** | `Apps[AppName]`. **Format visual → Slicer settings:** Style **Dropdown**, Selection **Single select**. Slicer **…** menu → **Search** |
| **Card** (heading) | `Selected App` |
| **Card** (details) | `Apps[AppId]`, `Apps[AuthMethods]`, `Apps[LastSignIn]`, `Apps[LastAppOnlySignIn]`, `Apps[Enabled]` (choose **First** for text fields; rename each in the well). Visual filter: `App Selected` **is 1** |
| **Table** visual: application permissions | `AppPermissions[Resource]`, `[Permission]`, `[PermissionDescription]`, `[GrantedDate]`. Visual filters: `AppPermissions[PermissionType]` = **Application**, `App Selected` **is 1** |
| **Table** visual: owners | `AppOwners[OwnerName]`, `[OwnerUPN]`, `[OwnerOf]`. Visual filter: `App Selected` **is 1** |
| **Table** visual: credentials (optional) | `AppCredentials[CredentialType]`, `[Name]`, `[EndDate]`, `[Expiry Status]`. Visual filter: `App Selected` **is 1** |

Turn **Totals** off on each table (**Format visual → Visual → Totals → Off**).

**Make it the default view:**

1. Drag the **App Lookup** tab to the far left.
2. Clear the slicer (eraser icon).
3. Save while App Lookup is on screen.

The Power BI Service opens the report on the page that was showing when it was saved and published.

### Page 1: Overview

| Visual | Fields |
|---|---|
| 6 cards | `App Count`, `Apps with Client Secret`, `Apps without Owner`, `Secrets Older Than 1 Year`, `Credentials Expiring in 30 Days`, `Expired Credentials` |
| Bar chart: auth methods | Y: `AppCredentials[CredentialType]`, X: `App Count`. Visual filter: `AppCredentials[Expiry Status]` = everything except `1 - Expired`, so only credentials still in use are counted |
| Bar chart: top permissions | Y: `AppPermissions[Permission]`, X: `App Count`. Visual filters: `PermissionType = Application`, and Top N = 15 by `App Count` |
| Bar chart: credential expiry | Y: `AppCredentials[Expiry Status]`, X: count of `AppCredentials[KeyId]` |
| Donut: sign-in activity | Legend: `Apps[Sign-in Status]`, Values: `App Count` |

### Page 2: App Inventory (main page for filtering)

| Visual | Fields |
|---|---|
| 7 **Slicer** visuals, **one field each** (dropdown, with **search** turned on) | `AppPermissions[Permission]`, `AppPermissions[PermissionType]`, `AppPermissions[Resource]`, `AppOwners[OwnerName]`, `AppCredentials[CredentialType]`, `Apps[HasClientSecret]`, `Apps[Sign-in Status]` |
| 1 **Table** visual, all fields in its **Columns** well | `Apps[AppName]`, `Apps[AppId]`, `Apps[AuthMethods]`, `Owners`, `Application Permissions`, `Delegated Permissions`, `Oldest Active Secret (days)`, `Next Credential Expiry`, `Apps[LastSignIn]` |

- **Keep `AppId` in the table.** Without it, two apps with the same name merge into one row. You can make the column narrow.
- **Conditional formatting** (column → Conditional formatting → Background colour → Rules):
  - `Oldest Active Secret (days)`: red over 365, amber 181–365.
  - `Next Credential Expiry`: Power BI can't apply rules to a date directly. Choose **Format style: Rules**, set **What field should we base this on?** to `Days to Next Expiry`, and add a rule: 0 to 30 → red.
- **Find apps with no owner:** select `(No owner)` in the Owner slicer. Selecting `(No permissions)` in the Permission slicer works the same way.

### Page 3: Credentials

| Visual | Fields |
|---|---|
| 3 **Slicer** visuals, **one field each** | `AppCredentials[CredentialType]`, `AppCredentials[Source]`, `AppCredentials[Expiry Status]` |
| 1 **Table** visual | `Apps[AppName]`, `AppCredentials[CredentialType]`, `AppCredentials[Source]`, `AppCredentials[Name]`, `AppCredentials[StartDate]`, `AppCredentials[EndDate]`, `AppCredentials[Age (days)]`, `AppCredentials[Days to Expiry]`, `AppCredentials[Expiry Status]` |

Sort by `Days to Expiry` ascending to see the most urgent credentials first. `Source = EnterpriseApp` shows the secrets that don't appear in the App registrations blade.

### Page 4: Owners

| Visual | Fields |
|---|---|
| Bar chart: apps per owner | Y: `AppOwners[OwnerName]`, X: `App Count`, sorted descending |
| 1 **Table** visual | `AppOwners[OwnerName]`, `AppOwners[OwnerUPN]`, `AppOwners[OwnerOf]`, `Apps[AppName]`, `Apps[AuthMethods]` |

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

### Refresh with new files

Overwrite the four CSVs in the same folder and click **Refresh**. The table queries accept both naming styles (`Apps.csv` and `EntraApp_Apps.csv`), and use the newest file if more than one matches.

### Switch to SharePoint

Change the parameters only: no query needs editing. Go to **Home → Transform data → Edit parameters** and set:

- `SourceType` = `SharePoint`
- `SharePointSiteUrl` = `https://<tenant>.sharepoint.com/sites/<SiteName>`
- `SharePointFolderName` = the base folder the pipeline uploads to

`Files` reads only that base folder, not its week subfolders, which contain files with the same names. See [powerbi/README.md](powerbi/README.md#switching-to-sharepoint) for details.

### Publish and schedule

1. **Home → Publish** to your Power BI workspace.
2. In the Power BI Service, open the semantic model's **Settings → Data source credentials** and sign in to SharePoint with your organisational account (OAuth2).
3. Set **Scheduled refresh** for a time after the pipeline's run, for example pipeline at 05:00 and refresh at 06:00.
4. Parameters can also be changed here, under **Settings → Parameters**, without opening Desktop.

For local or OneDrive-synced files, scheduled refresh in the Service needs an on-premises data gateway. Moving to SharePoint avoids that.

### Share

| Option | When to use |
|---|---|
| **Workspace → Create app** | Wider audience. Readers get a clean, read-only view, and you choose which pages they see |
| **Share** button on the report | A few named people |
| Workspace **Viewer** role | Colleagues who already use the workspace |

- Publish to a shared workspace, not *My workspace*, if others will use the report.
- Viewers usually need a **Power BI Pro** licence, unless the workspace is on Premium or Fabric capacity.
- The report shows which apps hold powerful permissions, which use secrets, and which secrets have expired. Share it only with the people who need it, such as identity, security and app owners.
- A Power BI *dashboard* (pinned tiles) can't use slicers. To add the App Lookup page to one, open the report in the Service and **pin** the page as a **live page**.

---

## Troubleshooting

| Problem | Likely cause and fix |
|---|---|
| Any error while loading the data (Power Query or Close & Apply) | See the troubleshooting table in [powerbi/README.md](powerbi/README.md#troubleshooting). It covers every load error found while building the report. |
| A Permission or Owner slicer doesn't filter the app table | That relationship's cross-filter direction isn't **Both** (step 4). |
| Two apps appear as one row | `Apps[AppId]` is missing from the table visual. |
| `LastSignIn` is empty for every app | The export had no sign-in data (needs Entra ID P1/P2 and `AuditLog.Read.All`). `Sign-in Status` then shows "No sign-in recorded" for all apps. |
| A measure's formula is underlined in red at `Apps[...]` | The measure has the same name as the `Apps` table. Rename it to `App Count` and update the measures that refer to it. |
| *"Can't display this visual"* with `ConditionNotHierarchicallyRelated` | Fields from different tables were put in one slicer (or one visual's axis). Use one field per slicer. |
| `Owners`, `Application Permissions` or another per-app measure shows the same (full) list in every row | The measure uses `REMOVEFILTERS ( <child table> )` without `TREATAS`, which also removes the row's app. Use the versions in step 6. |
| The table's Total row shows a long combined list | Turn totals off: select the table → **Format visual → Visual → Totals → Off**. |
| A table on the App Detail page (or any Table visual) is empty or missing rows, and its headings say "Sum of …" | The table hides rows where the summed values are blank. Set **Summarization → Don't summarize** for `ConsentedUserCount`, `Age (days)` and `Days to Expiry` (step 5). |
| Ages or expiry status look a day out of date | Calculated columns update on refresh. Refresh the report. |
