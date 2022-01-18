let
// -------------------------------------------------------
// fnManageColumns
// -------------------------------------------------------
// Conditionally adds, removes or renames columns
// Operation = "REMOVE" - Removes column if exists
// Operation = "ADD" - Adds column if NOT exists
// Operation = "RENAME" - Renames column if exists
// -------------------------------------------------------
// 2018-05-31    tswm00    new PQ


// -------------------------------------------------------
// Query body
// -------------------------------------------------------
    fnManageColumns = (CurrentTbl as table, CurrentColumn as text, Operation as text, optional NewColumn as text, optional NewColumnText as text  ) as any =>
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
            
        else
            CurrentTbl

    in
        Choice
in
    fnManageColumns