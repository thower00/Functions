// fnReturnFilePath
let
    fnReturnFilepath = (FullFilePath as text) as text =>
    let
        File = Text.AfterDelimiter(FullFilePath, "\", {0, RelativePosition.FromEnd}),
        Path = Text.Start(FullFilePath, Text.Length(FullFilePath)-Text.Length(File))
    in
        Path
in
    fnReturnFilepath