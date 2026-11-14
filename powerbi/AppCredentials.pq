let
    Content    = Files{[Name = "AppCredentials.csv"]}[Content],
    FirstTwo   = Binary.ToList(Binary.Range(Content, 0, 2)),
    Encoding   = if FirstTwo = {255, 254} then 1200 else 65001,
    Csv        = Csv.Document(Content, [Delimiter = ",", Encoding = Encoding, QuoteStyle = QuoteStyle.Csv]),
    Promoted   = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),
    CleanNames = Table.TransformColumnNames(Promoted, each Text.Trim(Text.Replace(_, Character.FromNumber(65279), ""))),
    Cleaned    = Table.ReplaceValue(CleanNames, "", null, Replacer.ReplaceValue, Table.ColumnNames(CleanNames)),
    Typed      = Table.TransformColumnTypes(Cleaned, {
        {"AppId", type text}, {"AppName", type text}, {"CredentialType", type text}, {"Source", type text},
        {"Name", type text}, {"KeyId", type text}, {"Hint", type text}, {"Thumbprint", type text},
        {"Issuer", type text}, {"Subject", type text}, {"StartDate", type date}, {"EndDate", type date},
        {"Status", type text}}, "en-US")
in
    Typed
