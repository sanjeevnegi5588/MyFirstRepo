let
    // Where the CSVs come from is controlled by the parameters only:
    //   SourceType = "Local"      -> every file in LocalFolderPath
    //   SourceType = "SharePoint" -> files in the SharePointFolderName folder of SharePointSiteUrl
    //                                (that folder only, not its week subfolders)
    Source  = if SourceType = "SharePoint"
              then Table.SelectRows(
                       SharePoint.Files(SharePointSiteUrl, [ApiVersion = 15]),
                       each Text.EndsWith([Folder Path], "/" & SharePointFolderName & "/"))
              else Folder.Files(LocalFolderPath),
    CsvOnly = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv"),
    Result  = Table.SelectColumns(CsvOnly, {"Name", "Content", "Date modified", "Folder Path"})
in
    Result
