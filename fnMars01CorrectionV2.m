let
// -------------------------------------------------------
pqName = "fnMars01CorrectionV2",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// Korrigerar fel när excel tolkar tex Mars01 som ett datum
// -------------------------------------------------------
/*
Parametrar:
    DataTbl_p
    ColName_p
    ExclStringLengthGreaterEqual_p
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-09-28    tswm00    V1: New PQ.
2021-10-28    tswm00    V2: Förbättrar hantering av nullvärden
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    MSG_InvalidDataTbl =            preERRORmsg & "Invalid data table",
    MSG_InvalidColName =            preERRORmsg & "Invalid columnname",
    
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
p_Debug = true,
TESTString = "#FEL:Value NULL",

TestScenario = 4,
//    1: qReadAssetReportFileJSON
//    2: qReadMetaExtendedUserFile
//    3: qReadAssetReportFileJSON_RAW
//    4: qReadAssetReportFileXML_RAW

    DataTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        qReadAssetReportFileJSON
    else if TestScenario = 2 then
    let
        Source = qReadMetaExtendedUserFile,
        Debug =
        if p_Debug then
            Table.SelectRows(Source, each 
                [workforceID] = "P0091246" or
                [workforceID] = "P0091245" or
                [workforceID] = "P0091244" or
                [workforceID] = "P0091218" or
                [workforceID] = "P0091214")
        else
            Source
    in
        Debug

    else if TestScenario = 3 then
    let
        Source = qReadAssetReportFileJSON_RAW,
        Debug =
        if p_Debug then
            Table.SelectRows(Source, each 
                [subsId] = "15168518" or
                [subsId] = "15179038" or
                [subsId] = "15378882")
        else
            Source
    in
        Debug
    else if TestScenario = 4 then
    let
        Source = qReadAssetReportFileXML_RAW,
        Debug =
        if p_Debug then
            Table.SelectRows(Source, each 
                [subsId] = "15168518" or
                [subsId] = "15179038" or
                [subsId] = "15378882")
        else
            Source
    in
        Debug
    else
        null,

    ColName_p =
    //--------------------------------------
    if TestScenario = 1 then
        "GCA220"
    else if TestScenario = 2 then
        "uid"
    else if TestScenario = 3 then
        "GCA220"
    else if TestScenario = 4 then
        "GCA220"
    else
        null,

    ExclStringLengthGreaterEqual_p =
    //--------------------------------------
    if TestScenario = 1 then
        6
    else if TestScenario = 2 then
        8
    else if TestScenario = 3 then
        8
    else if TestScenario = 4 then
        Text.Length(TESTString)
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
    fnMars01Correction = 
    (
                    DataTbl_p                       as              table, 
                    ColName_p                       as              text,
        optional    ExclStringLengthGreaterEqual_p  as  nullable    number
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

        ColName = 
        if ColName_p = null then
            error MSG_InvalidColName
        else
            ColName_p,

        ExclStringLengthGreaterEqual = 
        if ExclStringLengthGreaterEqual_p = null then
            0
        else
            ExclStringLengthGreaterEqual_p,

        ColNameORG = ColName&"ORG",

        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        Source = DataTbl,
        ORGColName = Table.DuplicateColumn(Source, ColName, ColNameORG),

        SplitColumnByDelimiter = Table.SplitColumn(ORGColName, ColName, Splitter.SplitTextByEachDelimiter({"-"}, QuoteStyle.Csv, false), {"Part1", "Part2"}),
        IsNumber = Table.AddColumn(SplitColumnByDelimiter, "IsNumber", each try Number.From([Part1]) is number otherwise false, type logical),

        XXTEMPXX = Table.AddColumn(IsNumber, "XXTEMPXX", each
        if Text.Length([Part1]) <> null and Text.Length([Part1]) >= ExclStringLengthGreaterEqual then
            [Part1]
        else if [IsNumber] and [Part1] <> null and [Part2] = "mar" then
            "mars" & [Part1]
        else if not [IsNumber] and [Part1] <> null and [Part2] <> null then
            [Part1] & "-" & [Part2]
        else if not [IsNumber] and [Part1] <> null and [Part2] = null then
            [Part1]
        else
            Record.Field(_, ColNameORG), type text),

        RemovedColumnsX = Table.RemoveColumns(XXTEMPXX,{"Part1", "Part2", "IsNumber", ColNameORG}),
        RenamedColumnsX = Table.RenameColumns(RemovedColumnsX,{{"XXTEMPXX", ColName}})
    in
        RenamedColumnsX
///*
in
    fnMars01Correction
//*/