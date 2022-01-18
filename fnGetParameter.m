let
    fnGetParameter = (ParameterName as text) =>
    let
        ParamSource = Excel.CurrentWorkbook(){[Name="tParameter"]}[Content],
        ParamRow = Table.SelectRows(ParamSource, each ([Parameter] = ParameterName)),
        Value=
        if Table.IsEmpty(ParamRow) = true then
            null
        else
            Record.Field(ParamRow{0},"Value")
    in
        Value
in
    fnGetParameter