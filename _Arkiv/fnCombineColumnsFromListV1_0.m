let
// -------------------------------------------------------
pqName = "fnCombineColumnsFromListV1_0",
// -------------------------------------------------------
// -------------------------------------------------------
/* Beskrivning
    Funktion för att kombinera en eller flera kolumner baserat på en lista med med kolumnnamn
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-05-04    tswm00    V1_0: Ny funktion
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
TestScenario = 1,
//    1: 
//    2: 
//    3: 

    DataTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        qAssetReportREAD
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
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
    else
        null,
        
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
                    ColumnList_p        as              text,       // Namn på tabell med kolumns namn
                    MergedColumnName_p  as              text,       // Namn på ny kolumn
                    Separator_p         as  nullable    text        // Skiljetecken
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
        if ColumnList_p = null then
            error ERR_InvalidColumnList
        else
            Table.SelectRows(Excel.CurrentWorkbook(){[Name=ColumnList_p]}[Content], each ([ColumnNames] <> null))[ColumnNames],

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
        
        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        Result = Table.AddColumn(DataTbl, MergedColumnName, each Text.Combine(List.Transform(ColumnList, (col) => Record.Field(_, col)), Separator), type text)

    in
        Result
///*
in
    fnCombineColumnsFromList
//*/