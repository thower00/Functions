let
// -------------------------------------------------------
pqName = "fnGroupCountV4",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// Grupperar och räknar förekomsten av värden i angiven kolumn
// -------------------------------------------------------
/*
Parametrar:
    DataTbl_p
    GroupCol_p
    CountColName_p 
    IndexColName_p
    Prefix_p
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-07-05    tswm00    V1: New PQ.
2021-08-05    tswm00    V2: Lägger till #"xxx" för att klara alla kolumnamn
2021-08-06    tswm00    V3: Lägger till optional Index kolumn för att kunna testa/ta bort dubletter
2021-08-09    tswm00    V4: Lägger till option Prefix för att få prefix på expanderade kolumner
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = "#ERROR: ",
    preINFOmsg = "#INFO: ",

    MSG_InvalidDataTbl =            preERRORmsg & "Invalid data table",
    MSG_InvalidGroupColName =       preERRORmsg & "Invalid group columnname",
    MSG_GroupColNameDontExists =    preERRORmsg & "Count columnname does not exist in table",
    MSG_InvalidCountColName =       preERRORmsg & "Invalid count columnname",
    
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 6,
//    1: TestDataTable
//    2: TestDataTable
//    3: Testdata från O365 debitering
//    4: Extern testtabell (excel)
//    5: TestDataTable med index
//    6: TestDataTable med index och prefix

    DataTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        TestDataTable
    else if TestScenario = 2 then
        TestDataTable
    else if TestScenario = 3 then
        qServiceTEST
    else if TestScenario = 4 then
        Excel.CurrentWorkbook(){[Name="testTable"]}[Content]
    else if TestScenario = 5 then
        TestDataTable
    else if TestScenario = 6 then
        TestDataTable
    else
        null,

    GroupColName_p =
    //--------------------------------------
    if TestScenario = 1 then
        "Col1"
    else if TestScenario = 2 then
        "User Service"
    else if TestScenario = 3 then
        "UserService"
    else if TestScenario = 4 then
        "Column 1"
    else if TestScenario = 5 then
        "User Service"
    else if TestScenario = 6 then
        "User Service"
    else
        null,

    CountColName_p =
    //-------------------------------------
    if TestScenario = 1 then
        "CountCol1"
    else if TestScenario = 2 then
        "CountUserService"
    else if TestScenario = 3 then
        "CountUserService"
    else if TestScenario = 4 then
        "Count"&GroupColName_p
    else if TestScenario = 5 then
        "CountUserService"
    else if TestScenario = 6 then
        "CountUserService"
    else
        null,

    IndexColName_p =
    //-------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        "IndexColumn"
    else if TestScenario = 6 then
        "IndexColumn"
    else
        null,

    Prefix_p =
    //-------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        null
    else if TestScenario = 6 then
        "PREFIX"
    else
        null,

    TestDataTable = Table.FromRecords
    ({
    [Col1 = "ABCD12",   #"User Service" = "Col2-1",   #"Column 3" = 1,     #"Col&4" = true,    Col5 = 1],
    [Col1 = "ABCD12",   #"User Service" = "Col2-2",   #"Column 3" = 2,     #"Col&4" = false,   Col5 = 2],
    [Col1 = "BCDEF12",  #"User Service" = "Col2-3",   #"Column 3" = 3,     #"Col&4" = true,    Col5 = 3],
    [Col1 = "CDEFGH12", #"User Service" = "Col2-4",   #"Column 3" = 4,     #"Col&4" = false,   Col5 = 4],
    [Col1 = "DEFGHI12", #"User Service" = "Col2-4",   #"Column 3" = 4,     #"Col&4" = false,   Col5 = 4]
    },
    type table [Col1=Text.Type, #"User Service"=Text.Type, #"Column 3"=Number.Type, #"Col&4"=Logical.Type, Col5=Int64.Type]),

*/
// =======================================================================
XXX_EXTERNAL_REFERENCES_XXX = null,
// =======================================================================

// =======================================================================
XXX_QUERY_BODY_XXX = null,
// =======================================================================
///*
    fnGroupCount =
    (
                    DataTbl_p       as              table,  // Tabellen som skall bearbetas
                    GroupColName_p  as              text,   // Namn på den kolumn som skall grupperas/räknas
                    CountColName_p  as              text,   // Namn på resultat kolumnen
        optional    IndexColName_p  as  nullable    text,   // Namn på (optionell) indexkolumn
        optional    Prefix_p        as  nullable    text    // Prefix på expanderade kolumner (optionell)
    ) 
    as table =>

    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------

        DataTbl = 
        if DataTbl_p = null then
            error MSG_InvalidDataTbl
        else
            DataTbl_p,

        GroupColName = 
        if GroupColName_p = null then
            error MSG_InvalidGroupColName
        else if not Table.HasColumns(DataTbl,GroupColName_p) then
            error MSG_GroupColNameDontExists
        else
            GroupColName_p,

        CountColName =
        if CountColName_p = null then
            error MSG_InvalidCountColName
        else
            CountColName_p,

        IndexColName = IndexColName_p,

        Prefix = 
        if Prefix_p = null then
            null
        else
            Prefix_p & ".",

        // ---------------------------------------------------
        // Manage 
        // ---------------------------------------------------

        Source = DataTbl,

        // Namn på kolumnen som skall grupperas måste finnas i en lista
        GroupColumnNameList = Table.FromRecords({[Col = GroupColName]})[Col],

        // Namn på kolumnen som skall innehålla resultatet måste finnas i en lista
        CountColNameList = Table.FromRecords({[Col = CountColName]})[Col],

        Schema = Table.Schema(Source),

        // Lista på alla ursprungskolumner
        AllColumnNamesList = Schema[Name],

        // Lista på de kolumner som skall expanderas FRÅN (exkl grupperingskolumnen)
        ExpandColumnNamesListFrom =
        let
            ExpandColumnName = Schema,
            ExcludeGroupColName = Table.SelectRows(ExpandColumnName, each ([Name] <> GroupColName)),
            Result = ExcludeGroupColName[Name]
        in
            Result,

        // Lista på de kolumner som skall expanderas TILL (exkl grupperingskolumnen)
        ExpandColumnNamesListTo =
        let
            ExpandColumnName = Schema,
            ExcludeGroupColName = Table.SelectRows(ExpandColumnName, each ([Name] <> GroupColName)),
            AddPrefix = 
            if Prefix <> null then
            let
                PrefixName = Table.AddColumn(ExcludeGroupColName, "PrefixName", each Prefix & [Name]),
                RemoveName = Table.RemoveColumns(PrefixName,"Name"),
                RenameCol = Table.RenameColumns(RemoveName,{"PrefixName","Name"})
            in
                RenameCol
            else
                ExcludeGroupColName,

            Result = AddPrefix[Name]
        in
            Result,

        // Skapa expression för Table.Group
        ExpressionToEvaluate =
        let
            Src = Schema,
            Pre = "#""",
            Post = """",

            NameExtended = Table.AddColumn(Src, "NameExtended", each Pre & [Name] & Post),
            TypeTable = Table.AddColumn(NameExtended, "TypeTable", each [NameExtended]&"="&[Kind], type text),
            Expression = "type table [" & Text.Combine(TypeTable[TypeTable],", ") & "]"
        in
            Expression,

        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        Group = Table.Group(Source, GroupColumnNameList, {{CountColName, each Table.RowCount(_), Int64.Type}, {"AllRows", each _, Expression.Evaluate(ExpressionToEvaluate)}}),
        AddIndex = 
        if IndexColName <> null then
            Table.AddIndexColumn(Group, IndexColName, 0, 1, Int64.Type)
        else
            Group,
            
        Expand = Table.ExpandTableColumn(AddIndex, "AllRows", ExpandColumnNamesListFrom, ExpandColumnNamesListTo),
        Result = Expand
    in
        Result
///*
in
    fnGroupCount
//*/