let
// =================================================================================================================================
// fnExcelWebContentV1
// Returns either content or a valid reference to an Excelfile on SharePoint 
// ---------------------------------------------------------------------------------------------------------------------------------
//    2018-08-27    tswm00    New PQ
// =================================================================================================================================
/* XXXX Testdata XXXX
        Param = Text.From(fnGetParameter("ParameterSet")),

        BaseURL_p =             Text.From(fnGetParameter("BaseURL_p"&Param)),
        Site_p =                Text.From(fnGetParameter("Site_p"&Param)),
        Folder_p =              Text.From(fnGetParameter("Folder_p"&Param)),
        AltStartFolder_p =      Text.From(fnGetParameter("AlternateStartFolder_p"&Param)),
        FileNameContains_p =    Text.From(fnGetParameter("FileNameContains_p"&Param)),
        Newest_p =              Logical.From(fnGetParameter("Newest_p"&Param)),
        ReturnContent_p =       Logical.From(fnGetParameter("ReturnContent_p"&Param)),
        SourceTbl_p =           Text.From(fnGetParameter("SourceTbl_p"&Param)),
        SourceSh_p =            Text.From(fnGetParameter("SourceSh_p"&Param)),
*/

    fnExcelWebContent =                   // Retrives a full path to a SharePoint file possible to read with a Web.Contents() request
        (
        BaseURL_p           as text,                // SharePoint Base URL, e.g. "https://karlstadskommunonline.sharepoint.com/"
        Site_p              as text,                // SharePoint Site, e.g "/sites/org-r18kasyo/"
        Folder_p            as text,                // SharePoint folder path e.g. "tjanster/Allmänt/Samlingsavtal/Underlag/O365/"
        AltStartFolder_p    as nullable text,       // Alterative base folder name if necessary e.g. "Tjänster"
        FileNameContains_p  as nullable text,       // Filename filter (contains) e.g. "xlsx" or "Rapport",
        Newest_p            as nullable logical,    // Retrive newest file if true or omitted, oldest if false. Default newest (true)
        ReturnContent_p     as nullable logical,    // true if fileCONTENT should be returned, false if fileNAME should be returned. Default fileCONTENT (true)
        SourceTbl_p         as nullable text,       // Name of named table to read data from. If omitted a sheetname must be provided (if ReturnContent_p = true)
        SourceSh_p          as nullable text        // Name of named Sheet to read data from. If omitted a tablename must be provided (if ReturnContent_p = true)
        )
    as any =>

    let
// =================================================================================================================================
// Constants
// =================================================================================================================================
    ListData_c = "/_vti_bin/listdata.svc",
    Msg1 = "#UNUSED",
    Msg2 = "#ERROR",
    Msg3 = "#ERROR:File missing",

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
    FileReference = 
    if not Table.IsEmpty(ODataFeedTable) then
        let
            SelectFolder = Table.SelectRows(ODataFeedTable, each [Path] = "/" & Site & "/" & Folder),
    
            SelectFileName = 
            if FileNameContains <> null then
                Table.SelectRows(SelectFolder, each Text.Contains([Name], FileNameContains))
            else
                SelectFolder,

            SortedRows = 
            if Newest then
                Table.Sort(SelectFileName,{{"Created", Order.Descending}})
            else
                Table.Sort(SelectFileName,{{"Created", Order.Ascending}}),

            KeptFirstRows = Table.FirstN(SortedRows,1),
            FileName = BaseURL & "/" & Site & "/" & Folder & "/" & Record.Field(KeptFirstRows{0},"Name")
        in
            FileName
    else
        Msg3,

// ---------------------------------------------------------------------------------------------------------------------------------
    Source = Excel.Workbook(Web.Contents(FileReference), null, true),
// ---------------------------------------------------------------------------------------------------------------------------------
    Result = 
    if not Text.StartsWith(FileReference,"#") then
        if ReturnContent then
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
        else
            FileReference
    else
        Msg3

    in
        Result
in
    fnExcelWebContent