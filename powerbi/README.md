# Power BI queries: setup

These steps load the four CSV files from `Export-EntraAppReport.ps1` into Power BI Desktop. The file location is kept in **parameters**, so switching from a local folder to SharePoint later only means changing parameter values. No query needs editing.

These are steps 1–3 of [PowerBI-Setup.md](../PowerBI-Setup.md). After the tables load, continue with step 4 (data model) in that guide.

## Files in this folder

| File | Query name | Loaded into the report? |
|---|---|---|
| `SourceType.pq` | `SourceType` (parameter) | No |
| `LocalFolderPath.pq` | `LocalFolderPath` (parameter) | No |
| `SharePointSiteUrl.pq` | `SharePointSiteUrl` (parameter) | No |
| `SharePointFolderName.pq` | `SharePointFolderName` (parameter) | No |
| [Files.pq](Files.pq) | `Files` | No: it lists the CSVs for the four tables |
| [Apps.pq](Apps.pq) | `Apps` | **Yes** |
| [AppPermissions.pq](AppPermissions.pq) | `AppPermissions` | **Yes** |
| [AppCredentials.pq](AppCredentials.pq) | `AppCredentials` | **Yes** |
| [AppOwners.pq](AppOwners.pq) | `AppOwners` | **Yes** |

```
Parameters  (location lives here only)
  SourceType           "Local" or "SharePoint"
  LocalFolderPath      C:\...\EntraAppReport
  SharePointSiteUrl    https://<tenant>.sharepoint.com/sites/<Site>
  SharePointFolderName EntraAppReport
        │
        ▼
Files   (lists every CSV from the chosen location; not loaded)
        │
        ├── Apps            picks Apps.csv           or EntraApp_Apps.csv
        ├── AppPermissions  picks AppPermissions.csv or EntraApp_Permissions.csv
        ├── AppCredentials  picks AppCredentials.csv or EntraApp_Credentials.csv
        └── AppOwners       picks AppOwners.csv      or EntraApp_Owners.csv
```

Each table query accepts both the local export name (`Apps.csv`) and the pipeline name (`EntraApp_Apps.csv`), and uses the newest file if more than one matches. It also detects whether a file is UTF-8 (local export) or UTF-16 (pipeline export).

## How to copy a query

Open the `.pq` file on GitHub and click **Raw**. Then press **Ctrl+A**, **Ctrl+C**, and paste into Power Query's **Advanced Editor**.

Copying from the Raw view keeps the line breaks and avoids characters that Teams, Outlook or Word can change, such as curly quotes.

---

## Setup steps

### 1. Start a new Power BI file

**File → New** (or open Power BI Desktop).

If an earlier attempt failed to load in an existing file, start a new file rather than fixing the old one. Power BI keeps the previously loaded version of each query, and a broken earlier version can keep failing even after the query is corrected.

### 2. Turn off the privacy checks for this file

**File → Options and settings → Options → Current File → Privacy** → select **"Ignore the Privacy levels and potentially improve performance"** → OK.

This is required because the four tables read their data through the `Files` query. Without it, Power BI's privacy checks can fail when loading, with *"Unable to cast … ILetExpression"*. It is a per-file setting, so set it again in every new file.

### 3. Stop automatic relationships

In the same Options window, go to **Current File → Data Load** and untick **"Autodetect new relationships after data is loaded"**. You create the correct relationships yourself in step 4 of [PowerBI-Setup.md](../PowerBI-Setup.md).

### 4. Open Power Query

**Home → Transform data**.

### 5. Create the four parameters

For each parameter, go to **Home → Manage Parameters → New parameter**, set **Type** to **Text**, tick **Required**, and use these values:

| Name | Current value |
|---|---|
| `SourceType` | `Local` |
| `LocalFolderPath` | the folder holding the CSVs, e.g. `C:\Users\<you>\OneDrive - <Company>\EntraAppReport` |
| `SharePointSiteUrl` | `https://contoso.sharepoint.com/sites/SiteName` (placeholder until you switch to SharePoint) |
| `SharePointFolderName` | `EntraAppReport` |

Parameters created through **Manage Parameters** aren't loaded into the report. Alternatively, you can paste the `.pq` files for these four into blank queries. If you do, right-click each one and **untick Enable load**.

### 6. Create `Files`

1. **Home → New Source → Blank Query → Advanced Editor**.
2. Paste [Files.pq](Files.pq) and click **Done**.
3. Rename the query to **`Files`**.
4. Right-click it and **untick Enable load**.

### 7. Create the four tables

Do the same for each table query, but leave **Enable load on**:

| Query name | File |
|---|---|
| `Apps` | [Apps.pq](Apps.pq) |
| `AppPermissions` | [AppPermissions.pq](AppPermissions.pq) |
| `AppCredentials` | [AppCredentials.pq](AppCredentials.pq) |
| `AppOwners` | [AppOwners.pq](AppOwners.pq) |

Each table query should contain `Csv = if FirstTwo = {255, 254}` followed by **two** `Csv.Document` calls, one with `Encoding = 1200` and one with `Encoding = 65001`. If yours has `Encoding = Encoding`, it's an old version: copy it again from Raw.

### 8. Check what will load

In the Queries pane:

| Query | Should appear as |
|---|---|
| `SourceType`, `LocalFolderPath`, `SharePointSiteUrl`, `SharePointFolderName`, `Files` | *italic* (Enable load off) |
| `Apps`, `AppPermissions`, `AppCredentials`, `AppOwners` | normal text (Enable load on) |

Click each table and check that the preview shows proper column names (`AppId`, `AppName`, …) and data.

### 9. Close & Apply, then save

1. Click **Home → Close & Apply**.
2. Check that the **Data** pane on the right lists the four tables.
3. **Save** the file straight away, for example as `EntraAppReport.pbix`. Save again after each later step, so an accidental change can be undone by reopening.

Then continue with **step 4 (data model)** in [PowerBI-Setup.md](../PowerBI-Setup.md).

---

## Switching to SharePoint

When the pipeline uploads the files to SharePoint:

1. **Home → Transform data → Edit parameters** (or **Manage Parameters**).
2. Set:
   - `SourceType` = `SharePoint`
   - `SharePointSiteUrl` = your site, e.g. `https://<tenant>.sharepoint.com/sites/<Site>`
   - `SharePointFolderName` = the base folder the pipeline uploads to (only the folder name, e.g. `EntraAppReport`)
3. Click OK, and sign in with your **organisational account** when asked.
4. Click **Close & Apply**.

`Files` reads only the base folder, not its week subfolders, which contain files with the same names.

**After publishing to the Power BI Service**, you can change the same parameters under **Semantic model → Settings → Parameters**, without opening Desktop. If scheduled refresh asks for a gateway because of the local folder, edit the `Files` query once: delete the `else Folder.Files(LocalFolderPath)` part, so it only uses SharePoint.

---

## Troubleshooting

These are the problems hit while building the report, with their fixes.

| Symptom | Cause | Fix |
|---|---|---|
| **"Input string was not in a correct format"** on Close & Apply. *Copy details* shows `CsvDocument … set_Encoding` | A `Csv.Document` call has `Encoding = <variable>`. Power BI expects a plain number there when loading | Use the current `.pq` files, which have literal `1200` / `65001`. If it still fails, rebuild in a **new file** (step 1), because the old file keeps the broken loaded version |
| **"Unable to cast … InvocationExpressionSyntaxNode2 to … ILetExpression"** on load | Privacy checks rewriting queries that read through another query | Step 2: ignore privacy levels for this file |
| Load dialog lists **parameters or `Files`** as failing | They have Enable load on | Right-click → untick **Enable load** (step 8) |
| **"Invalid character"** when clicking a query in Power Query | The editor can't draw the preview. Seen together with the encoding problem above | Use the current `.pq` files in a new file. If a table loads (check Table view) but only the preview fails, you can continue |
| **"The column 'AppId' of the table wasn't found"** | The file was read with the wrong encoding | Use the current `.pq` files, which detect UTF-8 and UTF-16 |
| **"File not found: None of … was found in the source folder"** | Wrong `LocalFolderPath`, or the file has a different name | Check the parameter value and the file names in the folder |
| **"Invalid character"** when pasting into Advanced Editor | The text was changed on the way, e.g. curly quotes added by Teams or Outlook | Copy from GitHub's **Raw** view |
| A table **disappeared** | Deleting a table in Model or Report view also deletes its query | Recreate the query from its `.pq` file |
| Wrong lines between tables in Model view | Power BI created relationships automatically | Delete them, and turn off autodetect (step 3) |
| Delete fails with **"cannot be deleted because it's being referenced"** | Other queries use it, e.g. `Files` and the tables use the parameters | Edit the query instead of deleting it, e.g. through **Manage Parameters** |
