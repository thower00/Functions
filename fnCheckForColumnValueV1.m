let
// =================================================================================================================================
// fnCheckForColumnValueV1
// Returns true if value exists in column 
// ---------------------------------------------------------------------------------------------------------------------------------
/*
v1    2020-01-09    tswm00    New function
*/
// =================================================================================================================================
/*
DEBUG:
xxx
*/
    fnCheckForColumnValue =
        (
        CheckTable_p                    as table,               // 
        CheckColumn_p                   as text,                // 
        CheckTextValue_p                as text                 // 
        )
    as logical =>

    let

// =================================================================================================================================
// Constants
// =================================================================================================================================

    // --- Messages ---

// *********************************************************************************************************************************
// Initial preparation of parameters
// *********************************************************************************************************************************

// *********************************************************************************************************************************
// Prepare parameters
// *********************************************************************************************************************************
    CheckTable = CheckTable_p,
// ---------------------------------------------------------------------------------------------------------------------------------		
    CheckColumn = CheckColumn_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    CheckTextValue = CheckTextValue_p,
// ---------------------------------------------------------------------------------------------------------------------------------

// *********************************************************************************************************************************
// Function body
// *********************************************************************************************************************************

        SelectColumnValues = Table.SelectRows(CheckTable, each (Record.Field(_, CheckColumn) = CheckTextValue)),
        Result = Table.RowCount(SelectColumnValues) > 0

    in
        Result

in
    fnCheckForColumnValue