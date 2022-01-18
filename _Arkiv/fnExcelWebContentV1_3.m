let
// =================================================================================================================================
pqName = "fnExcelWebContentV1_3",
// Returns either content or a valid reference to an Excelfile on SharePoint 
// ---------------------------------------------------------------------------------------------------------------------------------
/*
v1      2018-08-27  tswm00  New PQ
v1.1    2018-09-11  tswm00  Added localization since returned column names are localization dependant
v1.2    2018-11-08  tswm00  Added optional parameter to return folder contents, which overrides ReturnContent_p. FELFALL ÄR INTE TESTADE!!!
V1.3    2020-11-19  tswm00  Lägger till parameter ReturnCreateDate_p
V1.3    2021-02-22  tswm00  Utökar testfallshanteringen. Ingen funktionalitet ändrad

*/
// =================================================================================================================================

/*
    ERR_TestCase = "Testcase fel",

    TestCase = 2,

    BaseURL_p =             
    if TestCase = 1 then
        "https://karlstadskommunonline.sharepoint.com/"
    else if TestCase = 2 then
        "https://karlstadskommunonline.sharepoint.com/"
    else
        error ERR_TestCase,

    Site_p =
    if TestCase = 1 then
        "sites/org-r18kasyo"
    else if TestCase = 2 then
        "/sites/projekt-ra9qqqea/"
    else
        error ERR_TestCase,

    Folder_p =
    if TestCase = 1 then
        "Gemensamma/Avdelning/Ekonomi/Faktureringsunderlag/IAM/RapporterTM2017/Debiteringsunderlag"
    else if TestCase = 2 then
        "/projektdokument/Elevdatorer"
    else
        error ERR_TestCase,

    AltStartFolder_p =
    if TestCase = 1 then
        null
    else if TestCase = 2 then
        "Projektdokument"
    else
        error ERR_TestCase,

    FileNameContains_p =
    if TestCase = 1 then
        "IAM"
    else if TestCase = 2 then
        "ElevdatorerIntune"
    else
        error ERR_TestCase,

    Newest_p =
    if TestCase = 1 then
        true
    else if TestCase = 2 then
        true
    else
        error null,


    ReturnContent_p =
    if TestCase = 1 then
        true
    else if TestCase = 2 then
        true
    else
        error null,

    SourceSh_p =
    if TestCase = 1 then
        "qtServiceDebit"
    else if TestCase = 2 then
        "Intune"
    else
        error ERR_TestCase,

    SourceTbl_p =
    if TestCase = 1 then
        "qtServiceDebit"
    else if TestCase = 2 then
        "tIntune"
    else
        error ERR_TestCase,

    ReturnFolderContent_p =
    if TestCase = 1 then
        false
    else if TestCase = 2 then
        false
    else
        error null,

    ReturnCreateDate_p =
    if TestCase = 1 then
        true
    else if TestCase = 2 then
        false
    else
        error null, 
*/

//*
    fnExcelWebContent =                   // Retrives a full path to a SharePoint file possible to read with a Web.Contents() request
        (
        BaseURL_p                       as text,                // SharePoint Base URL, e.g. "https://karlstadskommunonline.sharepoint.com/"
        Site_p                          as text,                // SharePoint Site, e.g "/sites/org-r18kasyo/"
        Folder_p                        as text,                // SharePoint folder path e.g. "tjanster/Allmänt/Samlingsavtal/Underlag/O365/"
        AltStartFolder_p                as nullable text,       // Alterative base folder name if necessary e.g. "Tjänster"
        FileNameContains_p              as nullable text,       // Filename filter (contains) e.g. "xlsx" or "Rapport",
        Newest_p                        as nullable logical,    // Retrive newest file if true or omitted, oldest if false. Default newest (true)
        ReturnContent_p                 as nullable logical,    // true if fileCONTENT should be returned, false if fileNAME should be returned. Default fileCONTENT (true)
        SourceTbl_p                     as nullable text,       // Name of named table to read data from. If omitted a sheetname must be provided (if ReturnContent_p = true)
        SourceSh_p                      as nullable text,       // Name of named Sheet to read data from. If omitted a tablename must be provided (if ReturnContent_p = true)
        optional ReturnFolderContent_p  as nullable logical,    // To return folder content rather than file content
        optional ReturnCreateDate_p     as nullable logical     // File create date is returnedeturn CreateDate
        )
    as any =>

    let
//*/

// =================================================================================================================================
// Constants
// =================================================================================================================================
    ListData_c = "/_vti_bin/listdata.svc",
    locale_EN = "EN",
    locale_SE = "SE",
    Col_Path_EN = "Path",
    Col_Path_SE = "Sökväg",
    Col_Name_EN = "Name",
    Col_Name_SE = "Namn",
    Col_Created_EN = "Created",
    Col_Created_SE = "Skapad",

    // --- Messages ---
    Msg1 = "#UNUSED",
    Msg2 = "#ERROR",
    Msg3 = "#ERROR:File missing",
    locale_ERROR = "#ERROR: Undefined locale",

// *********************************************************************************************************************************
// Initial preparation of parameters
// *********************************************************************************************************************************
    BaseURL_param =           
    if BaseURL_p = null or BaseURL_p = "null" or BaseURL_p = "" then
        null
    else
        Text.TrimStart(Text.TrimEnd(BaseURL_p,"/"),"/"),
// ---------------------------------------------------------------------------------------------------------------------------------
    Site_param =
    if Site_p = null or Site_p = "null" or Site_p = "" then
        null
    else
        Text.TrimStart(Text.TrimEnd(Site_p,"/"),"/"),
// ---------------------------------------------------------------------------------------------------------------------------------
    Folder_param =
    if Folder_p = null or Folder_p = "null" or Folder_p = "" then
        null
    else
        Text.TrimStart(Text.TrimEnd(Folder_p,"/"),"/"),
// ---------------------------------------------------------------------------------------------------------------------------------
    AltStartFolder_param =
    if AltStartFolder_p = null or AltStartFolder_p = "null" or AltStartFolder_p = "" then
        null
    else
        Text.TrimStart(Text.TrimEnd(AltStartFolder_p,"/"),"/"),
// ---------------------------------------------------------------------------------------------------------------------------------
    FileNameContains_param =
    if FileNameContains_p = null or FileNameContains_p = "null" or FileNameContains_p = "" then
        null
    else
        FileNameContains_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    Newest_param =
    if Newest_p = null or Newest_p = "null" or Newest_p = "" then
        true
    else
        Newest_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    ReturnContent_param =
    if ReturnContent_p = null or ReturnContent_p = "null" or ReturnContent_p = "" then
        true
    else
        ReturnContent_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceTbl_param =
    if SourceTbl_p = null or SourceTbl_p = "null" or SourceTbl_p = "" then
        null
    else
        SourceTbl_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceSh_param =
    if SourceSh_p = null or SourceSh_p = "null" or SourceSh_p = "" then
        null
    else
        SourceSh_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    ReturnFolderContent_param =
    if ReturnFolderContent_p = null or ReturnFolderContent_p = "null" or ReturnFolderContent_p = "" then
        false
    else
        ReturnFolderContent_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    ReturnCreateDate_param =
    if ReturnCreateDate_p = null or ReturnCreateDate_p = "null" or ReturnCreateDate_p = "" then
        false
    else
        ReturnCreateDate_p,
// *********************************************************************************************************************************
// Prepare parameters
// *********************************************************************************************************************************
    ListData = Text.TrimStart(Text.TrimEnd(ListData_c,"/"),"/"),
// ---------------------------------------------------------------------------------------------------------------------------------		
    BaseURL = BaseURL_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    Site = Site_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    Folder = Folder_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    StartFolder = Text.Start(Folder,Text.PositionOf(Folder,"/")),      
// ---------------------------------------------------------------------------------------------------------------------------------
    AltStartFolder =
    if AltStartFolder_param <> null then
        AltStartFolder_param
    else
        Text.Start(Folder,Text.PositionOf(Folder,"/")),
// ---------------------------------------------------------------------------------------------------------------------------------
    FileNameContains = FileNameContains_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    Newest = Newest_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    ReturnContent = ReturnContent_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    ReturnFolderContent = ReturnFolderContent_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    ReturnCreateDate = ReturnCreateDate_param,
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceTbl =
    if SourceTbl_param <> null then
        SourceTbl_param 
    else if SourceTbl_param  = null and SourceSh_param <> null then
        Msg1
    else
        Msg2,
// ---------------------------------------------------------------------------------------------------------------------------------
    SourceSh =
    if SourceSh_param <> null then
        SourceSh_param
    else if SourceSh_param = null and SourceSh_param <> null then
        Msg1
    else
        Msg2,
// *********************************************************************************************************************************
// Get content (or return filereference only)
// *********************************************************************************************************************************
    ODataFeedReference = BaseURL & "/" & Site & "/" & ListData,
    ODataFeed = OData.Feed(ODataFeedReference),
    ODataFeedTable = ODataFeed{[Name=AltStartFolder,Signature="table"]}[Data],
// ---------------------------------------------------------------------------------------------------------------------------------
    locale = 
    if Table.HasColumns(ODataFeedTable,Col_Name_EN) then
        // EN locale
        locale_EN
    else if Table.HasColumns(ODataFeedTable,Col_Name_SE) then
        // SE locale
        locale_SE
    else
        // Undefined locale
        locale_ERROR,

    Col_Name = 
    if not Table.IsEmpty(ODataFeedTable) then
        if locale = locale_EN then
            Col_Name_EN
        else if locale = locale_SE then
            Col_Name_SE
        else        
            locale_ERROR & ":" & Col_Name_EN
    else
        Msg3,

    Col_Path = 
    if not Table.IsEmpty(ODataFeedTable) then
        if locale = locale_EN then
            Col_Path_EN
        else if locale = locale_SE then
            Col_Path_SE
        else        
            locale_ERROR & ":" & Col_Path_EN
    else
        Msg3,

    Col_Created = 
    if not Table.IsEmpty(ODataFeedTable) then
        if locale = locale_EN then
            Col_Created_EN
        else if locale = locale_SE then
            Col_Created_SE
        else        
            locale_ERROR & ":" & Col_Created_EN
    else
        Msg3,

    SelectFolder = Table.SelectRows(ODataFeedTable, each (Record.Field(_, Col_Path) = "/" & Site & "/" & Folder)),
    
    SelectFileName = 
    if FileNameContains <> null then
        Table.SelectRows(SelectFolder, each Text.Contains((Record.Field(_, Col_Name)), FileNameContains))
    else
        SelectFolder,

    SortedRows = 
    if Newest then
        Table.Sort(SelectFileName,{{Col_Created, Order.Descending}})
    else
        Table.Sort(SelectFileName,{{Col_Created, Order.Ascending}}),

    KeptFirstRows = Table.FirstN(SortedRows,1),
    FileName = BaseURL & "/" & Site & "/" & Folder & "/" & Record.Field(KeptFirstRows{0},Col_Name),
    CreateDate = DateTime.From(Record.Field(KeptFirstRows{0},Col_Created)),
// ---------------------------------------------------------------------------------------------------------------------------------
    Source = Excel.Workbook(Web.Contents(FileName), null, true),
// ---------------------------------------------------------------------------------------------------------------------------------
    Result = 
    if not Text.StartsWith(FileName,"#") then
        if ReturnCreateDate then
            CreateDate
        else if ReturnFolderContent then
            let
                Source = SelectFileName,
                FullFileName = Table.AddColumn(Source, "FullFileName", each  BaseURL & "/" & Site & "/" & Folder & "/" & [Name],type text),
                KeepName = Table.SelectColumns(FullFileName,{"FullFileName"})
            in
                KeepName

        else if ReturnContent then
            if not Text.StartsWith(SourceTbl,"#") then
                Source{[Item=SourceTbl,Kind="Table"]}[Data]

            else if not Text.StartsWith(SourceSh,"#") then
                let
                    Output = Source{[Item=SourceSh,Kind="Sheet"]}[Data],
                    PromotedHeaders = Table.PromoteHeaders(Output, [PromoteAllScalars=true])
                in
                    PromotedHeaders
            else
                Msg3

        else if not ReturnContent then
            FileName

        else
            Msg3
    else
        Msg3

    in
        Result
///*
in
    fnExcelWebContent
//*/