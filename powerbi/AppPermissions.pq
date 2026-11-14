let
    Content    = Files{[Name = "AppPermissions.csv"]}[Content],
    FirstTwo   = Binary.ToList(Binary.Range(Content, 0, 2)),
    Encoding   = if FirstTwo = {255, 254} then 1200 else 65001,
    Csv        = Csv.Document(Content, [Delimiter = ",", Encoding = Encoding, QuoteStyle = QuoteStyle.Csv]),
    Promoted   = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),
    CleanNames = Table.TransformColumnNames(Promoted, each Text.Trim(Text.Replace(_, Character.FromNumber(65279), ""))),
    Cleaned    = Table.ReplaceValue(CleanNames, "", null, Replacer.ReplaceValue, Table.ColumnNames(CleanNames)),
    Typed      = Table.TransformColumnTypes(Cleaned, {
        {"AppId", type text}, {"AppName", type text}, {"PermissionType", type text}, {"Resource", type text},
        {"Permission", type text}, {"PermissionDescription", type text}, {"ConsentType", type text},
        {"ConsentedUserCount", Int64.Type}, {"GrantedDate", type date}}, "en-US")
in
    Typed
