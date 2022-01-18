let
// =================================================================================================================================
// fnReadExcelWebContents
// Retrives data from Excel file via a Web.Contents() request either from named table or named sheet
// ---------------------------------------------------------------------------------------------------------------------------------
//    2018-06-21    tswm00    New function

// =================================================================================================================================
// External references
// =================================================================================================================================

// =================================================================================================================================
// Constants
// =================================================================================================================================
    
// =================================================================================================================================
// Query body
// =================================================================================================================================

    fnReadExcelWebContents =        // Retrives data from Excel file via a Web.Contents() request either from named table or named sheet
        (
        SourceFile_p as text,                    // Full path to file on weblocation, e.g. "https://karlstadskommunonline.sharepoint.com/sites/org-r18kasyo/Gemensamma/Test.xlsx"
        SourceTbl_p as nullable text,            // Name of named table to read data from. If omitted a sheetname must be provided
        SourceSh_p as nullable text              // Name of named Sheet to read data from. If omitted a tablename must be provided
        )
    as any =>

    let

// ---------------------------------------------------------------------------------------------------------------------------------
    SourceFile = SourceFile_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceTblFix1 = 
    if SourceTbl_p = "" then
        null
    else
        SourceTbl_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceShFix1 = 
    if SourceSh_p = "" then
        null
    else
        SourceSh_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceTbl =
    if SourceTblFix1 <> null then
        SourceTblFix1 
    else if SourceTblFix1  = null and SourceShFix1 <> null then
        "#UNUSED"
    else
        "#ERROR",
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceSh =
    if SourceShFix1 <> null then
        SourceShFix1
    else if SourceShFix1 = null and SourceTblFix1 <> null then
        "#UNUSED"
    else
        "#ERROR",
// ---------------------------------------------------------------------------------------------------------------------------------
    Source = Excel.Workbook(Web.Contents(SourceFile), null, true),
// ---------------------------------------------------------------------------------------------------------------------------------
    _Table = 
    if not Text.StartsWith(SourceTbl,"#") then
        Source{[Item=SourceTbl,Kind="Table"]}[Data]

    else if not Text.StartsWith(SourceSh,"#") then
        Source{[Item=SourceSh,Kind="Sheet"]}[Data]

    else
        null
in
    _Table
in
    fnReadExcelWebContents