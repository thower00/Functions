let
// -------------------------------------------------------
pqName = "fnINTERFACEV1",
/*--------------------------------------------------------
Ändringslogg
----------------------------------------------------------
2020-12-17    tswm00    Ny
2020-12-18    tswm00    Lägger till möjligheten att skicka med en mappningstabell
----------------------------------------------------------
Beskrivning
----------------------------------------------------------
Normerar utdata genom att ange indatakolumn(er) och normerad utdatakolumn för en tabell
Om utdatakolumnen redan finns i indatat så returnas tabellen as is

Parametrar:
OutdataCol_p            Namn på den nya utdatakolumnen
IndataTbl_p             Referens till indatatabellen
IndataCol_p             Namn på indatakolumn
OutdataErrMsg_p         Felmeddelande om mappning ej går att göra
IndataColAlt1_p         Optional. Namn på indatakolumn alt 1
IndataColAlt2_p         Optional. Namn på indatakolumn alt 2
MappingTbl_p            Optional. Mappingtabell IndataCol|IndataColAlt1|IndataColAlt2|OutDataCol
--------------------------------------------------------*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),
        
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 5,
//    1: Testscenario klient 1
//    2: Testscenario klient 2
//    3: ??
//    4: MappingTabell
//    5: MappingTabell med saknade indatacol
//    6: Ingen indatacol angiven


    ClientTestData =
        let
            SQLServer = fnGetParameter("SQLServer"),
            SQLDB = fnGetParameter("SQLDB"),
            DBview = fnGetParameter("DBview"),

            Source = Sql.Databases(SQLServer),
            LDPROD = Source{[Name=SQLDB]}[Data],
            dbo_View_TM2017 = LDPROD{[Schema="dbo",Item=DBview]}[Data]
        in
            dbo_View_TM2017,


    OutdataCol_p = 
    //--------------------------------------
    if TestScenario = 1 then
        "ORGANISATION"
    else if TestScenario = 2 then
        "Förvaltning"
    else if TestScenario = 3 then
        "TESTKOL"
    else if TestScenario = 4 then
        "ORGANISATION"
    else if TestScenario = 5 then
        "TMSERVICE"
    else if TestScenario = 6 then
        "TMSERVICE"
    else
        null,

    IndataTbl_p = 
    //--------------------------------------
    if TestScenario = 1 then
        ClientTestData
    else if TestScenario = 2 then
        ClientTestData
    else if TestScenario = 3 then
        ClientTestData
    else if TestScenario = 4 then
        ClientTestData
    else if TestScenario = 5 then
        ClientTestData
    else if TestScenario = 6 then
        ClientTestData
    else
        null,

    IndataCol_p =
    //-------------------------------------
    if TestScenario = 1 then
        "Organisation"
    else if TestScenario = 2 then
        "Förvaltning"
    else if TestScenario = 3 then
        "TestInCol"
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        null
    else
        null,

    OutdataErrMsg_p =
    //-------------------------------------
    if TestScenario = 1 then
        "#FEL: Kan inte mappa data"
    else if TestScenario = 2 then
        "#FEL: Kan inte mappa data"
    else if TestScenario = 3 then
        "#FEL: Kan inte mappa data"
    else if TestScenario = 4 then
        "#FEL: Kan inte mappa data"
    else if TestScenario = 5 then
        "#FEL: Kan inte mappa data"
    else if TestScenario = 6 then
        "#FEL: Kan inte mappa data"
    else
        null,
    
    IndataColAlt1_p =
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
    else
        null,
        
    IndataColAlt2_p =
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
    else
        null,

    MappingTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "tAssetReportColumnMapping"
    else if TestScenario = 5 then
        "tAssetReportColumnMapping"
    else if TestScenario = 6 then
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
    fnINTERFACE =
    (
                    OutdataCol_p    as text,
                    IndataTbl_p     as table,
                    IndataCol_p     as nullable text,
                    OutdataErrMsg_p as any,
        optional    IndataColAlt1_p as nullable text,
        optional    IndataColAlt2_p as nullable text,
        optional    MappingTbl_p    as nullable text
    ) 
    as table =>

    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------
        OutdataCol = OutdataCol_p,
        IndataTbl = IndataTbl_p,
        IndataCol = IndataCol_p,
        OutdataErrMsg = OutdataErrMsg_p,
        IndataColAlt1 = IndataColAlt1_p,
        IndataColAlt2 = IndataColAlt2_p,
        MappingTbl = MappingTbl_p,

        // ---------------------------------------------------
        // query body 
        // ---------------------------------------------------

        MappingTable = 
        if MappingTbl <> null then
            let
                Source = Excel.CurrentWorkbook(){[Name=MappingTbl]}[Content],
                ChangedType = Table.TransformColumnTypes(Source,{{"IndataCol", type text}, {"IndataColAlt1", type text}, {"IndataColAlt2", type text}, {"OutdataCol", type text}}),
                OrganizeColumns = Table.SelectColumns(ChangedType,{"IndataCol", "IndataColAlt1", "IndataColAlt2", "OutdataCol"}),
                SelectOutdataCol = Table.SelectRows(OrganizeColumns, each ([OutdataCol] = OutdataCol))
            in
                SelectOutdataCol
        else
           null,

        IndataColFinal = 
        if MappingTable <> null then
            if MappingTable[IndataCol]{0} <> null then
                MappingTable[IndataCol]{0}
            else
                IndataCol
        else
            IndataCol,

        IndataColAlt1Final = 
        if MappingTable <> null then
            if MappingTable[IndataColAlt1]{0} <> null then
                MappingTable[IndataColAlt1]{0}
            else
                IndataColAlt1
        else
            IndataColAlt1,

        IndataColAlt2Final = 
        if MappingTable <> null then
            if MappingTable[IndataColAlt2]{0} <> null then
                MappingTable[IndataColAlt2]{0}
            else
                IndataColAlt2
        else
            IndataColAlt2,

        Source = IndataTbl,        

        Test = Table.HasColumns(Source,IndataColFinal),

        Result =
        if IndataColFinal = OutdataCol then
            // Indatakolumn och utdatakolumn är samma så returnera tabellen utan åtgärd
            Source

        else if IndataColFinal <> null then
            if Table.HasColumns(Source,IndataColFinal) then
                // Utdatakolumn sätts till värdet i indatakolumn
                Table.AddColumn(Source,OutdataCol, each (Record.Field(_, IndataColFinal)))
            else
                if IndataColAlt1Final <> null then
                    if Table.HasColumns(Source,IndataColAlt1Final) then
                        // Utdatakolumn sätts till värdet i indatakolumn alt1
                        Table.AddColumn(Source,OutdataCol, each (Record.Field(_, IndataColAlt1Final)))
                else if IndataColAlt2Final <> null then
                    if Table.HasColumns(Source,IndataColAlt2Final) then
                        // Utdatakolumn sätts till värdet i indatakolumn alt2
                        Table.AddColumn(Source,OutdataCol, each (Record.Field(_, IndataColAlt2Final)))
                    else
                        // Data går inte att mappa
                        Table.AddColumn(Source,OutdataCol, each OutdataErrMsg )
                else
                    // Data går inte att mappa
                    Table.AddColumn(Source,OutdataCol, each OutdataErrMsg )
            else
                // Data går inte att mappa
                Table.AddColumn(Source,OutdataCol, each OutdataErrMsg )
        else
            // Data går inte att mappa
            Table.AddColumn(Source,OutdataCol, each OutdataErrMsg )

    in
        Result
///*
in
    fnINTERFACE
//*/