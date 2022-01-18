let
// -------------------------------------------------------
// fnManageColumnsV2
// -------------------------------------------------------
// Conditionally adds, removes or renames columns
// Operation = "REMOVE" - Removes column if exists
// Operation = "ADD" - Adds column if NOT exists
// Operation = "RENAME" - Renames column if exists
// Operation = "CHANGETYPE" - Typecasts column if exists
// -------------------------------------------------------
/*
2018-05-31    tswm00    new PQ
2019-04-01    tswm00    Adds CHANGETYPE
*/
// -------------------------------------------------------
// Query body
// -------------------------------------------------------
    fnManageColumns =
    (
                CurrentTbl as table,
                CurrentColumn as text,
                Operation as text,
    optional    NewColumn as text, 
    optional    NewColumnText as text,
    optional    ChangeType as text      // TEXT, NUMBER, INTEGER, PERCENTAGE, LOGICAL, DATE, DATETIME
    ) 
    as any =>
    
    let
        Choice =
        if Text.Upper(Operation) = "REMOVE" then
            if Table.HasColumns(CurrentTbl,CurrentColumn) then
                Table.RemoveColumns(CurrentTbl,{CurrentColumn})
            else
                CurrentTbl

        else if Text.Upper(Operation) = "ADD" then
            if not Table.HasColumns(CurrentTbl,CurrentColumn) then
                Table.AddColumn(CurrentTbl, CurrentColumn, each NewColumnText,type text)
            else
                CurrentTbl

        else if Text.Upper(Operation) = "RENAME" then
            if Table.HasColumns(CurrentTbl ,CurrentColumn) then
                Table.RenameColumns(CurrentTbl ,{CurrentColumn, NewColumn})
            else
                CurrentTbl

        else if Text.Upper(Operation) = "CHANGETYPE" then
            if Table.HasColumns(CurrentTbl ,CurrentColumn) then
                if Text.Upper(ChangeType) = "TEXT" then
                    Table.TransformColumnTypes(CurrentTbl,{{CurrentColumn, type text}})
                    
                else if Text.Upper(ChangeType) = "NUMBER" then
                    Table.TransformColumnTypes(CurrentTbl,{{CurrentColumn, type number}})
                    
                else if Text.Upper(ChangeType) = "INTEGER" then
                    Table.TransformColumnTypes(CurrentTbl,{{CurrentColumn, Int64.Type}})
                    
                else if Text.Upper(ChangeType) = "PERCENTAGE" then
                    Table.TransformColumnTypes(CurrentTbl,{{CurrentColumn, Percentage.Type}})
                    
                else if Text.Upper(ChangeType) = "LOGICAL" then
                    Table.TransformColumnTypes(CurrentTbl,{{CurrentColumn, type logical}})
                    
                else if Text.Upper(ChangeType) = "DATE" then
                    Table.TransformColumnTypes(CurrentTbl,{{CurrentColumn, type date}})
                    
                else if Text.Upper(ChangeType) = "DATETIME" then
                    Table.TransformColumnTypes(CurrentTbl,{{CurrentColumn, type datetime}})
                else
                    CurrentTbl
            else
                CurrentTbl
        else
            CurrentTbl

    in
        Choice
in
    fnManageColumns