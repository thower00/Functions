let
// -------------------------------------------------------
pqName = "fnFilterTextV1",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// -------------------------------------------------------
/*
Parametrar:
    FilterTbl_p             Tabell med filtervärden
    FilterCol_p             Kolumn med filtervärden
    DataTbl_p               Tabell med data att filtrera
    DataCol_p               Kolumn med data att filtrera
    InclExcl_p              "INCLUDE" eller "EXCLUDE"
    IgnoreCase_p            True/null om case insensitivt
    Active_p                True om filtret skall vara aktivt
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2020-11-17    tswm00    New PQ. 
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    TempFilterColumn = "__FILTERTEMP",

    MSG_InvalidFilterTbl =      preERRORmsg & "Invalid filter table",
    MSG_InvalidFilterCol =      preERRORmsg & "Invalid filter column",
    MSG_InvalidDataTbl =        preERRORmsg & "Invalid data table",
    MSG_InvalidDataCol =        preERRORmsg & "Invalid data column",
    MSG_InvalidFilterChoice =   preERRORmsg & "Invalid filter choice",
    
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 7,
//    1: Include
//    2: Exclude
//    3: Inaktive
//    4: Include, Ignore Case
//    5: Exclude, Ignore Case
//    6: qAccountDataEMP, Include
//    7: qAccountDataEMP, Exclude

    FilterTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        TestFilterTable
    else if TestScenario = 2 then
        TestFilterTable
    else if TestScenario = 3 then
        TestFilterTable
    else if TestScenario = 4 then
        TestFilterTable
    else if TestScenario = 5 then
        TestFilterTable
    else if TestScenario = 6 then
        Excel.CurrentWorkbook(){[Name="tFilter"]}[Content]
    else if TestScenario = 7 then
        Excel.CurrentWorkbook(){[Name="tFilter"]}[Content]
    else
        null,

    FilterCol_p =
    //--------------------------------------
    if TestScenario = 1 then
        "UserID"
    else if TestScenario = 2 then
        "UserID"
    else if TestScenario = 3 then
        "UserID"
    else if TestScenario = 4 then
        "UserID"
    else if TestScenario = 5 then
        "UserID"
    else if TestScenario = 6 then
        "IncludeUserIDOnly"
    else if TestScenario = 7 then
        "ExcludeUserID"
    else
        null,

    DataTbl_p =
    //-------------------------------------
    if TestScenario = 1 then
        TestDataTable
    else if TestScenario = 2 then
        TestDataTable
    else if TestScenario = 3 then
        TestDataTable
    else if TestScenario = 4 then
        TestDataTable
    else if TestScenario = 5 then
        TestDataTable
    else if TestScenario = 6 then
        qReadMeta2Export
    else if TestScenario = 7 then
        qReadMeta2Export
    else
        null,

    DataCol_p =
    //--------------------------------------
    if TestScenario = 1 then
        "UserID"
    else if TestScenario = 2 then
        "UserID"
    else if TestScenario = 3 then
        "UserID"
    else if TestScenario = 4 then
        "UserID"
    else if TestScenario = 5 then
        "UserID"
    else if TestScenario = 6 then
        "UserID"
    else if TestScenario = 7 then
        "UserID"
    else
        null,
        
    InclExcl_p =
    //--------------------------------------
    if TestScenario = 1 then
        "Include"
    else if TestScenario = 2 then
        "Excl"
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "Include"
    else if TestScenario = 5 then
        "Excl"
    else if TestScenario = 6 then
        "Incl"
    else if TestScenario = 7 then
        "Exclude"
    else
        null,
        
    Active_p =
    //--------------------------------------
    if TestScenario = 1 then
        true
    else if TestScenario = 2 then
        true
    else if TestScenario = 3 then
        false
    else if TestScenario = 4 then
        true
    else if TestScenario = 5 then
        true
    else if TestScenario = 6 then
        true
    else if TestScenario = 7 then
        true
    else
        null,

    IgnoreCase_p =
    //--------------------------------------
    if TestScenario = 1 then
        false
    else if TestScenario = 2 then
        false
    else if TestScenario = 3 then
        false
    else if TestScenario = 4 then
        true
    else if TestScenario = 5 then
        true
    else if TestScenario = 6 then
        true
    else if TestScenario = 7 then
        true
    else
        null,

    TestFilterTable = Table.FromRecords
    ({
    [UserID = "ABCD12"],
    [UserID = "efgh34"],
    [UserID = "IJKL56"]
    }),

    TestDataTable = Table.FromRecords
    ({
    [UserID = "ABCD12",   Test1 = "Test1-1",   Test2 = "Test2-1"],
    [UserID = "EFGH34",   Test1 = "Test1-2",   Test2 = "Test2-2"],
    [UserID = "ijkl56",   Test1 = "Test1-3",   Test2 = "Test2-3"],
    [UserID = "XyZ123",   Test1 = "Test1-4",   Test2 = "Test2-4"]
    }),


*/
// =======================================================================
XXX_EXTERNAL_REFERENCES_XXX = null,
// =======================================================================

// =======================================================================
XXX_QUERY_BODY_XXX = null,
// =======================================================================
///*
    fnFilterText =
    (
        FilterTbl_p         as              table,
        FilterCol_p         as              text,
        DataTbl_p           as              table,
        DataCol_p           as              text,
        InclExcl_p          as              text,
        IgnoreCase_p        as nullable     logical,
        Active_p            as nullable     logical
    ) 
    as table =>

    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------

        FilterCol = 
        if FilterCol_p = null and Active then
            error MSG_InvalidFilterCol
        else if FilterCol_p = null and not Active then
            null
        else
            FilterCol_p,

        FilterTbl = 
        if FilterTbl_p = null and Active then
            error MSG_InvalidFilterTbl
        else if FilterTbl_p = null and not Active then
            null
        else
            Table.SelectRows(FilterTbl_p, each Record.Field(_, FilterCol) <> null),

        Active =
        if FilterTbl_p = null then
            false
        else if Table.IsEmpty(FilterTbl) then
            false
        else if Active_p = null or Active_p = false then
            false
        else
            true,

        DataTbl = 
        if DataTbl_p = null and Active then
            error MSG_InvalidDataTbl
        else if DataTbl_p = null and not Active then
            null
        else
            DataTbl_p,

        DataCol = 
        if DataCol_p = null and Active then
            error MSG_InvalidDataCol
        else if DataCol_p = null and not Active then
            null
        else
            DataCol_p,

        Include =
        if InclExcl_p = null or Text.StartsWith(Text.Upper(InclExcl_p),"I") then
            true
        else
            false,

        IgnoreCase =
        if IgnoreCase_p = null or IgnoreCase_p = true then
            true
        else
            false,

        // ---------------------------------------------------
        // Manage 
        // ---------------------------------------------------

        FilterList =
        if FilterTbl <> null and FilterCol <> null and Active then
            if IgnoreCase then
                Table.Column(Table.TransformColumns(FilterTbl,{{FilterCol, Text.Upper}}),FilterCol)
            else
                Table.Column(FilterTbl,FilterCol)
        else if not Active then
            null
        else
            error MSG_InvalidFilterTbl & " or/and " & MSG_InvalidFilterCol,

        DataTblToFilter = Table.AddColumn(DataTbl,TempFilterColumn, each
        if DataTbl <> null and DataCol <> null and Active then
            if IgnoreCase then
                Text.Upper(Record.Field(_, DataCol))
            else
                Record.Field(_, DataCol)
        else if not Active then
            null        
        else
            MSG_InvalidDataTbl & " or/and " & MSG_InvalidDataCol, type text),
    

        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        Filter = 
        if Active then
            Table.SelectRows(DataTblToFilter, each (List.Contains(FilterList,Record.Field(_, TempFilterColumn))=Include))
        else
            DataTblToFilter,

        Result = Table.RemoveColumns(Filter,{TempFilterColumn})
    in
        Result
///*
    in
        fnFilterText
//*/