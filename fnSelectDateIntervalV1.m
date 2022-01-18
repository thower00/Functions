let
// -------------------------------------------------------
pqName = "fnSelectDateIntervalV1",
// -------------------------------------------------------
// -------------------------------------------------------
/* Beskrivning
    Selekterar rader i en tabell baserad på datumintervall (FromDate - ToDate)
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-10-22    tswm00    V1: New PQ
2021-11-05    tswm00    V1: Rättar datumselekteringen
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    DefaultFromDate = DateTime.From("1900-01-01"),
    DefaultToDate = DateTime.From("2099-12-31"),
    
    ERR_InvalidDataTable =          preERRORmsg & pqName & ":Invalid table name",
    ERR_InvalidFromDateColName =    preERRORmsg & pqName & ":Invalid FromDate column name",
    ERR_InvalidToDateColName =      preERRORmsg & pqName & ":Invalid ToDate column name",
    ERR_InvalidDateInterval =       preERRORmsg & pqName & ":Invalid date interval:",
    
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================

/*
TestScenario = 6,
//    1:  
//    2: 
//    3: 
//    4: 
//    5: qPersonalLOAD
//    6: qBasicValuesLOAD


    DataTable_p =
    //--------------------------------------
    if TestScenario = 1 then
        TestDataTable
    else if TestScenario = 2 then
        TestDataTable
    else if TestScenario = 3 then
        TestDataTable
    else if TestScenario = 4 then
        TestDataTable
    else if TestScenario = 5 then
        qPersonalLOAD
    else if TestScenario = 6 then
        qBasicValuesLOAD
    else
        null,

    FromDateColName_p =
    //--------------------------------------
    if TestScenario = 1 then
        "FromDateTest"
    else if TestScenario = 2 then
        "FromDateTest"
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "FromDateTest"
    else if TestScenario = 5 then
        "FromDatum"
    else if TestScenario = 6 then
        "FromDatum"
    else
        null,

    FromDate_p =
    //--------------------------------------
    if TestScenario = 1 then
        DateTime.From("2021-01-01")
    else if TestScenario = 2 then
        DateTime.From("2021-03-31")
    else if TestScenario = 3 then
        DateTime.From(null)
    else if TestScenario = 4 then
        DateTime.From(null)
    else if TestScenario = 5 then
        DateTime.From(fnGetParameter("FromDate"))
    else if TestScenario = 6 then
        DateTime.From(fnGetParameter("FromDate"))
    else
        null,

    ToDateColName_p =
    //--------------------------------------
    if TestScenario = 1 then
        "ToDateTest"
    else if TestScenario = 2 then
        "ToDateTest"
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "ToDateTest"
    else if TestScenario = 5 then
        "TomDatum"
    else if TestScenario = 6 then
        "TomDatum"
    else
        null,

    ToDate_p =
    //--------------------------------------
    if TestScenario = 1 then
        DateTime.From("2021-03-01")
    else if TestScenario = 2 then
        DateTime.From("2021-12-31")
    else if TestScenario = 3 then
        DateTime.From(null)
    else if TestScenario = 4 then
        DateTime.From(null)
    else if TestScenario = 5 then
        DateTime.From(fnGetParameter("ToDate"))
    else if TestScenario = 6 then
        DateTime.From(fnGetParameter("ToDate"))
    else
        null,

    TestDataTable = Table.FromRecords
    ({
    [Data = "ABCD12",   FromDateTest = DateTime.From("2021-01-01"), ToDateTest = DateTime.From("2021-06-30")],
    [Data = "EFGH34",   FromDateTest = DateTime.From("2021-04-01"), ToDateTest = DateTime.From("2021-12-31")],
    [Data = "ijkl56",   FromDateTest = DateTime.From("2021-01-01"), ToDateTest = DateTime.From(null)],
    [Data = "XyZ123",   FromDateTest = DateTime.From(null),         ToDateTest = DateTime.From("2021-06-01")]
    }),

*/

// =======================================================================
XXX_EXTERNAL_REFERENCES_XXX = null,
// =======================================================================

// =======================================================================
XXX_QUERY_BODY_XXX = null,
// =======================================================================
///*
    fnSelectDateInterval =
    (
                    DataTable_p         as          table,      // Tabell med indata
                    FromDateColName_p   as          text,       // FromDate kolumnnamn
                    FromDate_p          as          datetime,   // FromDate
                    ToDateColName_p     as          text,       // ToDate kolumnnamn
                    ToDate_p            as          datetime   // ToDate
    ) 
    as table =>

    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------
        DataTable = 
        if DataTable_p = null then
            error ERR_InvalidDataTable
        else
            DataTable_p,

        FromDateColName = 
        if FromDateColName_p = null then
            error ERR_InvalidFromDateColName
        else
            FromDateColName_p,

        FromDate = 
        if FromDate_p = null then
            DefaultFromDate
        else
            FromDate_p,

        ToDateColName = 
        if ToDateColName_p = null then
            error ERR_InvalidToDateColName
        else
            ToDateColName_p,

        ToDate = 
        if ToDate_p = null then
            DefaultToDate
        else
            ToDate_p,

        // ---------------------------------------------------
        // Query body 
        // ---------------------------------------------------
        Source =
        if FromDate > ToDate then
            error ERR_InvalidDateInterval & "FromDate:" & Text.From(FromDate) & ", ToDate:" & Text.From(ToDate)
        else
            DataTable,

        ChangeType = Table.TransformColumnTypes(Source,{{FromDateColName, type datetime}, {ToDateColName, type datetime}}),
        SetDefaultFromDate = Table.ReplaceValue(ChangeType,null,DefaultFromDate,Replacer.ReplaceValue,{FromDateColName}),
        SetDefaultToDate = Table.ReplaceValue(SetDefaultFromDate,null,DefaultToDate,Replacer.ReplaceValue,{ToDateColName}),

        SelectInterval = Table.SelectRows(SetDefaultToDate, each (FromDate >= Record.Field(_, FromDateColName)) and (ToDate <= Record.Field(_, ToDateColName))),

        Result = SelectInterval
    in
        Result

///*
in
    fnSelectDateInterval
//*/