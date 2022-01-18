let
// -------------------------------------------------------
pqName = "fnCombineColumnsFromListV1_1",
// -------------------------------------------------------
// -------------------------------------------------------
/* Beskrivning
    Funktion för att kombinera en eller flera kolumner baserat på en lista med med kolumnnamn
    Om referensen till ColumnNamesList anges som ett tabellnamn eller en tabell så måste kolumnnamnet vara "ColumnNames"
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-05-04    tswm00    V1_0: Ny funktion
2021-06-04    tswm00    V1_1: Gör det möjligt att ColumnList som antingen referens till tabell (tabellnamn som text), tabellreferens eller list referens
                              Lagt till attribute för att ange om originalkolumnerna skall behållas eller tas bort
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    ERR_InvalidDataTable =          preERRORmsg & "Invalid datatable",
    ERR_InvalidColumnList =         preERRORmsg & "Invalid columnlist",
    ERR_InvalidMergedColumnName =   preERRORmsg & "Invalid keyname",

// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 5,
//    1: 
//    2: 
//    3: 
//    4: ColumnNames as text
//    5: ColumnNames as table
//    6: ColumnNames as list

    DataTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        qAssetReportREAD
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        TestDataTable
    else if TestScenario = 5 then
        TestDataTable
    else if TestScenario = 6 then
        TestDataTable
    else
        null,

    ColumnList_p =
    //--------------------------------------
    if TestScenario = 1 then
        "tServiceMappingKey"
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "tTestCombineColumnNames"
    else if TestScenario = 5 then
        TestColumnNamesTable
    else if TestScenario = 6 then
        TestColumnNamesList
    else
        null,

    MergedColumnName_p =
    //--------------------------------------
    if TestScenario = 1 then
        "ServiceMappingKey"
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "MergedColumnName"
    else if TestScenario = 5 then
        "MergedColumnName"
    else if TestScenario = 6 then
        "MergedColumnName"
    else
        null,

    Separator_p =
    //--------------------------------------
    if TestScenario = 1 then
        ":"
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        ";"
    else if TestScenario = 5 then
        ";"
    else if TestScenario = 6 then
        ";"
    else
        null,

    KeepOriginal_p =
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        false
    else if TestScenario = 6 then
        true
    else
        null,


    TestDataTable = Table.FromRecords
    ({
        [Part1="Test1", Part2 = "Test2", Part3 = "Test3", Part4 = "Test4", Part5 = "Test5", Part6 = "Test6"]
    }),

    TestColumnNamesTable = Table.FromRecords
    ({
        [ColumnNames = "Part1"],
        [ColumnNames = "Part2"],
        [ColumnNames = "Part3"],
        [ColumnNames = "Part4"],
        [ColumnNames = "Part5"],
        [ColumnNames = "Part6"]
    }),

    TestColumnNamesList = {"Part1", "Part2","Part3","Part4","Part5","Part6"},

    TestColumnNamesText = "tColumnNames",

*/
// =======================================================================
XXX_EXTERNAL_REFERENCES_XXX = null,
// =======================================================================

// =======================================================================
XXX_QUERY_BODY_XXX = null,
// =======================================================================
///*
    fnCombineColumnsFromList =
    (
                    DataTbl_p           as              table,      // Datatabell
                    ColumnList_p        as              any,        // Texreferens, tabellreferens eller listreferens
                    MergedColumnName_p  as              text,       // Namn på ny kolumn
                    Separator_p         as  nullable    text,       // Skiljetecken
        optional    KeepOriginal_p      as  nullable    logical     // false om originalkolumner skall tas bort. Default true
    ) 
    as table =>


    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------
        DataTbl = 
        if DataTbl_p = null then
            error ERR_InvalidDataTable
        else
            DataTbl_p,

        ColumnList = 
        if Value.Type(ColumnList_p) = null then
            error ERR_InvalidColumnList

        else if Value.Is(ColumnList_p, type table) then
            Table.SelectRows(ColumnList_p, each ([ColumnNames] <> null))[ColumnNames]

        else if Value.Is(ColumnList_p, type list) then
            ColumnList_p

        else if Value.Is(ColumnList_p, type text) then
            Table.SelectRows(Excel.CurrentWorkbook(){[Name=ColumnList_p]}[Content], each ([ColumnNames] <> null))[ColumnNames]

        else
            error ERR_InvalidColumnList,

        MergedColumnName =
        if MergedColumnName_p = null then
            error ERR_InvalidMergedColumnName
        else
            MergedColumnName_p,

        Separator =
        if Separator_p = null then
            ""
        else
            Separator_p,

        KeepOriginal =
        if KeepOriginal_p = null then
            true
        else
            KeepOriginal_p,
        
        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        ResultTable = Table.AddColumn(DataTbl, MergedColumnName, each Text.Combine(List.Transform(ColumnList, (col) => Record.Field(_, col)), Separator), type text),

        Result = 
        if KeepOriginal then
            ResultTable
        else
            Table.RemoveColumns(ResultTable,ColumnList, MissingField.Ignore)

    in
        Result
///*
in
    fnCombineColumnsFromList
//*/