let
    Content    = Files{[Name = "Apps.csv"]}[Content],
    FirstTwo   = Binary.ToList(Binary.Range(Content, 0, 2)),
    Encoding   = if FirstTwo = {255, 254} then 1200 else 65001,
    Csv        = Csv.Document(Content, [Delimiter = ",", Encoding = Encoding, QuoteStyle = QuoteStyle.Csv]),
    Promoted   = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),
    CleanNames = Table.TransformColumnNames(Promoted, each Text.Trim(Text.Replace(_, Character.FromNumber(65279), ""))),
    Cleaned    = Table.ReplaceValue(CleanNames, "", null, Replacer.ReplaceValue, Table.ColumnNames(CleanNames)),
    Typed      = Table.TransformColumnTypes(Cleaned, {
        {"AppId", type text}, {"ServicePrincipalId", type text}, {"AppName", type text}, {"AppType", type text},
        {"IsMicrosoftApp", type logical}, {"RegisteredInTenant", type logical}, {"Publisher", type text},
        {"Enabled", type logical}, {"AppCreated", type date}, {"AuthMethods", type text},
        {"HasClientSecret", type logical}, {"ApplicationPermissionCount", Int64.Type},
        {"DelegatedPermissionCount", Int64.Type}, {"OwnerCount", Int64.Type},
        {"LastSignIn", type date}, {"LastAppOnlySignIn", type date}, {"SnapshotDate", type date}}, "en-US")
in
    Typed
