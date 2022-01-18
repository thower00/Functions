let
// -------------------------------------------------------
pqName = "fnGroupCountV2",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// Grupperar och räknar förekomsten av värden i angiven kolumn
// OBS! Klarar bara kolumnnamn (alla kolumner) utan specialtecken som skulle leda till #"Kolumnnamn"
// -------------------------------------------------------
/*
Parametrar:
    DataTbl_p
    GroupCol_p
    CountColumnName_p    
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-07-05    tswm00    V1: New PQ.
2021-08-05    tswm00    V2: Lägger till #"xxx" för att klara alla kolumnamn
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
TestScenario = 4,
//    1: xx
//    2: xx
//    3: Testdata från O365 debitering
//    4: Externa testtabell (excel)

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
    else
        null,

    CountColumnName_p =
    //-------------------------------------
    if TestScenario = 1 then
        "CountCol1"
    else if TestScenario = 2 then
        "CountUserService"
    else if TestScenario = 3 then
        "CountUserService"
    else if TestScenario = 4 then
        "Count"&GroupColName_p
    else
        null,

    TestDataTable = Table.FromRecords
    ({
    [Col1 = "ABCD12",   #"User Service" = "Col2-1",   Col3 = 1,     Col4 = true,    Col5 = 1],
    [Col1 = "ABCD12",   #"User Service" = "Col2-2",   Col3 = 2,     Col4 = false,   Col5 = 2],
    [Col1 = "BCDEF12",  #"User Service" = "Col2-3",   Col3 = 3,     Col4 = true,    Col5 = 3],
    [Col1 = "CDEFGH12", #"User Service" = "Col2-4",   Col3 = 4,     Col4 = false,   Col5 = 4],
    [Col1 = "DEFGHI12", #"User Service" = "Col2-4",   Col3 = 4,     Col4 = false,   Col5 = 4]
    },
    type table [Col1=Text.Type, #"User Service"=Text.Type, Col3=Number.Type, Col4=Logical.Type, Col5=Int64.Type]),

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
        DataTbl_p           as              table,  // Tabellen som skall bearbetas
        GroupColName_p      as              text,   // Namn på den kolumn som skall grupperas/räknas
        CountColumnName_p   as              text    // Namn på resultat kolumnen
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

        CountColumnName =
        if CountColumnName_p = null then
            error MSG_InvalidCountColName
        else
            CountColumnName_p,

        // ---------------------------------------------------
        // Manage 
        // ---------------------------------------------------

        Source = DataTbl,

        // Namn på kolumnen som skall grupperas måste finnas i en lista
        GroupColumnNameList = Table.FromRecords({[Col = GroupColName]})[Col],

        // Namn på kolumnen som skall innehålla resultatet måste finnas i en lista
        CountColumnNameList = Table.FromRecords({[Col = CountColumnName]})[Col],

        Schema = Table.Schema(Source),

        // Lista på alla ursprungskolumner
        AllColumnNamesList = Schema[Name],

        // Lista på de kolumner som skall expanderas (exkl grupperingskolumnen)
        ExpandColumnNamesList =
        let
            ExpandColumnName = Schema,
            ExcludeGroupColName = Table.SelectRows(ExpandColumnName, each ([Name] <> GroupColName)),
            Result = ExcludeGroupColName[Name]
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

        // Skapa reorderlist med CountColumnName först
        ReorderList = List.InsertRange(AllColumnNamesList, 0, CountColumnNameList),

        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        Group = Table.Group(Source, GroupColumnNameList, {{CountColumnName, each Table.RowCount(_), Int64.Type}, {"AllRows", each _, Expression.Evaluate(ExpressionToEvaluate)}}),
        Expand = Table.ExpandTableColumn(Group, "AllRows", ExpandColumnNamesList),
        ReorderColumns = Table.ReorderColumns(Expand,ReorderList),

        Result = ReorderColumns
    in
        Result
///*
in
    fnGroupCount
//*/