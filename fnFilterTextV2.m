let
// -------------------------------------------------------
pqName = "fnFilterTextV2",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// -------------------------------------------------------
/*
Parametrar:
    DataTbl_p               Tabell med data att filtrera
    DataCol_p               Kolumn med data att filtrera
    FilterTbl_p             Tabell med filtervärden
    FilterCol_p             Kolumn med filtervärden
    InclExcl_p              "INCLUDE" eller "EXCLUDE"
    IgnoreCase_p            True/null om case insensitivt
    Active_p                True om filtret skall vara aktivt
    Match                   Ange ett kolumnnamn om en kolumn som anger match (true/false) för extern filtrering (Optional)
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2020-11-17    tswm00    Ny funktion
2021-02-17    tswm00    Lägger till parameter Match_p för att returnera kolumn med match (true/false)
                        OBS!! Byter ordning så att Data och Filter byter plats i parameterlistan
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    TempFilterColumn = "__FILTERTEMP__",
    TempMatchColumn  = "__MATCHTEMP__",

    MSG_InvalidFilterTbl =      preERRORmsg & "Invalid filter table",
    MSG_InvalidFilterCol =      preERRORmsg & "Invalid filter column",
    MSG_InvalidDataTbl =        preERRORmsg & "Invalid data table",
    MSG_InvalidDataCol =        preERRORmsg & "Invalid data column",
    MSG_InvalidFilterChoice =   preERRORmsg & "Invalid filter choice",
    
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 5,
//    1: Testdata: Include, NOT Ignore Case
//    2: Testdata: Exclude, NOT Ignore Case
//    3: Testdata: Inaktive
//    4: Testdata: Include, Ignore Case
//    5: Testdata: Exclude, Ignore Case
//    6: qAccountDataEMP: Include
//    7: qAccountDataEMP: Exclude
//    8: Testdata: Exclude, Match, NOT Ignore Case
//    9: Testdata: Include, Match, Ignore Case
//   10: Klient asset-namn
//   11: Testdata: Include, Match, Ignore Case, Inactive (samma som 9 fast Inactive)

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
    else if TestScenario = 8 then
        TestFilterTable
    else if TestScenario = 9 then
        TestFilterTable
    else if TestScenario = 10 then
        qExcludeAssets
    else if TestScenario = 11 then
        TestFilterTable
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
    else if TestScenario = 8 then
        "UserID"
    else if TestScenario = 9 then
        "UserID"
    else if TestScenario = 10 then
        "Asset-namn"
    else if TestScenario = 11 then
        "UserID"
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
    else if TestScenario = 8 then
        TestDataTable
    else if TestScenario = 9 then
        TestDataTable
    else if TestScenario = 10 then
        qReadServiceData_DEV
    else if TestScenario = 11 then
        TestDataTable
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
    else if TestScenario = 8 then
        "UserID"
    else if TestScenario = 9 then
        "UserID"
    else if TestScenario = 10 then
        "ASSETNAMN"
    else if TestScenario = 11 then
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
    else if TestScenario = 8 then
        "Exc"
    else if TestScenario = 9 then
        "Inc"
    else if TestScenario = 10 then
        "Exc"
    else if TestScenario = 11 then
        "Inc"
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
    else if TestScenario = 8 then
        true
    else if TestScenario = 9 then
        true
    else if TestScenario = 10 then
        true
    else if TestScenario = 11 then
        false
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
    else if TestScenario = 8 then
        false
    else if TestScenario = 9 then
        true
    else if TestScenario = 10 then
        true
    else if TestScenario = 11 then
        true
    else
        null,

    Match_p =
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
        null
    else if TestScenario = 6 then
        null
    else if TestScenario = 7 then
        null
    else if TestScenario = 8 then
        "Match"
    else if TestScenario = 9 then
        "Match"
    else if TestScenario = 10 then
        "ExcludeAssetsAssetNamn"
    else if TestScenario = 11 then
        "Match"
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
                    DataTbl_p      as              table,
                    DataCol_p      as              text,
                    FilterTbl_p    as              table,
                    FilterCol_p    as              text,
                    InclExcl_p     as              text,
                    IgnoreCase_p   as nullable     logical,
                    Active_p       as nullable     logical,
        optional    Match_p        as nullable     text
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
            // Välj rader i filtertabellen som är <> null
            Table.SelectRows(FilterTbl_p, each Record.Field(_, FilterCol) <> null),

        Active =
        if FilterTbl_p = null then
            // Om referens till filtertabellen saknas betraktas detta som att Active = false
            false
        else if Table.IsEmpty(FilterTbl) then
            // Om filtertabellen är tom betraktas detta som att Active = false
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

        Match =
        if Match_p = null then
            // Om Match_p inte är angiven så används en temporär matchkolumn till
            TempMatchColumn
        else
            Match_p,

        ReturnMatch = Match <> TempMatchColumn,

        // ---------------------------------------------------
        // Manage 
        // ---------------------------------------------------

        FilterList =
        if FilterTbl <> null and FilterCol <> null and Active then
            if IgnoreCase then
                // Skapa en lista från filtertabellen baserad på angiven kolumn, transformerad till uppercase
                Table.Column(Table.TransformColumns(FilterTbl,{{FilterCol, Text.Upper}}),FilterCol)
            else
                // Skapa en lista från filtertabellen baserad på angiven kolumn, as is (ej uppercase)
                Table.Column(FilterTbl,FilterCol)
        else if not Active then
                // Skapa en helt tom lista för att få matchningen att inte generera fel
                {}
        else
            error MSG_InvalidFilterTbl & " or/and " & MSG_InvalidFilterCol,

        DataTblToFilter = Table.AddColumn(DataTbl,TempFilterColumn, each
        // Lägg till en temporär datakolumn för att kunna använda matchning med uppercase
        if DataTbl <> null and DataCol <> null and Active then
            if IgnoreCase then
                // Transformera den temporära datakolumnen till uppercase
                Text.Upper(Record.Field(_, DataCol))
            else
                // Behåll den temporära datakolumnen as-is
                Record.Field(_, DataCol)
        else if not Active then
            null        
        else
            MSG_InvalidDataTbl & " or/and " & MSG_InvalidDataCol, type text),

        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        // Lägg till matchningskolumn och fyll den med resultatet av matchningen mellan FilterList och TempFilterColumn
        MatchResults = Table.AddColumn(DataTblToFilter, Match, each List.Contains(FilterList,Record.Field(_, TempFilterColumn)), type logical),

        Filter =
        if ReturnMatch then
            // Om matchning ska returneras så returneras alltid resultatet oavsett om Active är true/false
            // Om Active = false skall dock icke-match returneras i kolumnen vilket åstadkomms med att FilterList är tom i det fallet
            MatchResults

        else if Active then
            let
                // Filtrera och ta bort Match kolumnen för att leverera ett filtrerat resultat
                Select = Table.SelectRows(MatchResults, each Record.Field(_, Match)=Include),
                Remove = Table.RemoveColumns(Select,{Match})
            in
                Remove
        else
            // Filtret är inaktivt så returnera tabellen as-is
            DataTblToFilter,

        // Ta bort den temporära filterkolumnen och returnera
        Result = Table.RemoveColumns(Filter,{TempFilterColumn})
in
    Result


///*
    in
        fnFilterText
//*/