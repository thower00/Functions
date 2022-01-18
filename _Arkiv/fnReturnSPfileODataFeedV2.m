let
// =================================================================================================================================
// fnReturnSPfileODataFeedV2
// Retrives a full path to a SharePoint file possible to read with a Web.Contents() request
// ---------------------------------------------------------------------------------------------------------------------------------
/*
2018-06-21    tswm00    New function
2020-10-02    tswm00    Correct pathhandling
*/
// =================================================================================================================================
// External references
// =================================================================================================================================

// =================================================================================================================================
// Constants
// =================================================================================================================================
    SharePointListData = "/_vti_bin/listdata.svc",

// =================================================================================================================================
// Query body
// =================================================================================================================================   
///*
    fnReturnSPfileODataFeed =                   // Retrives a full path to a SharePoint file possible to read with a Web.Contents() request
        (
        SharePointBaseURL_p as text,                    // SharePoint Base URL, e.g. "https://karlstadskommunonline.sharepoint.com/"
        SharePointSite_p as text,                       // SharePoint Site, e.g "/sites/org-r18kasyo/"
        SharePointFolder_p as text,                     // SharePoint folder path e.g. "Tjänster/Allmänt/Samlingsavtal/Underlag/O365/"
        SharePointAlternateFolder_p as nullable text,   // Alterative base folder name if necessary e.g. "/tjanster/"
        SharePointFilenameContains_p as nullable text,  // Filename filter (contains) e.g. "xlsx" or "Rapport",
        optional SharePointNewest_p as logical          // Retrive newest file if true or omitted, oldest if false
        )
    as any =>

    let
//*/
/*
// ---------------------------------------------------------------------------------------------------------------------------------
// Debugdata
// ---------------------------------------------------------------------------------------------------------------------------------
        SharePointBaseURL_p              = "https://karlstadskommunonline.sharepoint.com/",
        SharePointSite_p                 = "sites/org-r18kasyo/",
        SharePointFolder_p               = "/Gemensamma/Avdelning/Ekonomi/Faktureringsunderlag/Telefoni/RapporterTM2017/AssetReport/",
        SharePointAlternateFolder_p      = null,
        SharePointFilenameContains_p     = "subscriptionAccount",
        SharePointNewest_p               = true,

*/
// ---------------------------------------------------------------------------------------------------------------------------------
    SharePointFilenameContains = SharePointFilenameContains_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SharePointNewest =
    if SharePointNewest_p = null then
        true
    else
        SharePointNewest_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SharePointAlternateFolder = 
    if SharePointAlternateFolder_p = "" or SharePointAlternateFolder_p = null then
        null
    else if Text.StartsWith(SharePointAlternateFolder_p,"/") and not Text.EndsWith(SharePointAlternateFolder_p,"/")  then
        Text.TrimStart(SharePointAlternateFolder_p,"/")
    else if not Text.StartsWith(SharePointAlternateFolder_p,"/") and Text.EndsWith(SharePointAlternateFolder_p,"/")  then
        Text.TrimEnd(SharePointAlternateFolder_p,"/")
    else if Text.StartsWith(SharePointAlternateFolder_p,"/") and Text.EndsWith(SharePointAlternateFolder_p,"/")  then
        Text.TrimStart(Text.TrimEnd(SharePointAlternateFolder_p,"/"),"/")
    else
        SharePointAlternateFolder_p,
// ---------------------------------------------------------------------------------------------------------------------------------		
    SharePointBaseURL =
    if not Text.EndsWith(SharePointBaseURL_p,"/") then
        SharePointBaseURL_p & "/"
    else
        SharePointBaseURL_p,
// ---------------------------------------------------------------------------------------------------------------------------------
    SharePointSite =
        let
            TrimStart = Text.TrimStart(SharePointSite_p,"/"),
            TrimEnd = Text.TrimEnd(TrimStart,"/")
        in
            TrimEnd,
// ---------------------------------------------------------------------------------------------------------------------------------
    SharePointFolderFix1 =
    // Remove "/" in beginning if present
    if Text.StartsWith(SharePointFolder_p,"/") then
        Text.TrimStart(SharePointFolder_p,"/")
    else
        SharePointFolder_p,
		
    SharePointFolderFix2 =
    // Make sure folder starts with "/" & BaseFolder omitting URL and sitename
    if Text.StartsWith(SharePointFolderFix1,SharePointSite) then
        Text.End(SharePointFolderFix1, Text.Length(SharePointFolderFix1)-Text.Length(SharePointSite)-1)
    else
        SharePointFolderFix1,

    SharePointStartFolder = Text.Start(SharePointFolderFix2,Text.PositionOf(SharePointFolderFix2,"/")),

    SharePointFolder =
    // If an alternative startfolder is provided, replace with existing
    if SharePointAlternateFolder <> null then
        "/" & SharePointSite & "/" & Text.Replace(SharePointFolderFix2,SharePointStartFolder,SharePointAlternateFolder)
    else
        "/" & SharePointSite & "/" & SharePointFolderFix2,
// ---------------------------------------------------------------------------------------------------------------------------------        
    ODataFeedReference = SharePointBaseURL & SharePointSite & SharePointListData,				
// ---------------------------------------------------------------------------------------------------------------------------------
    Source = OData.Feed(SharePointBaseURL & SharePointSite & SharePointListData),
    _Table = Source{[Name=SharePointStartFolder ,Signature="table"]}[Data],
// ---------------------------------------------------------------------------------------------------------------------------------
    SelectFolder = Table.SelectRows(_Table, each [Path] = SharePointFolder),
    SelectFileName = 
    if SharePointFilenameContains <> "" and SharePointFilenameContains <> null then
        Table.SelectRows(SelectFolder, each Text.Contains([Name], SharePointFilenameContains))
    else
        SelectFolder,

    SortedRows = 
    if SharePointNewest then
        Table.Sort(SelectFileName,{{"Created", Order.Descending}})
    else
        Table.Sort(SelectFileName,{{"Created", Order.Ascending}}),

    KeptFirstRows = Table.FirstN(SortedRows,1),
// ---------------------------------------------------------------------------------------------------------------------------------
    FileName = SharePointBaseURL & Text.TrimStart(SharePointFolder,"/") & Record.Field(KeptFirstRows{0},"Name")

in
    FileName
///*
in
    fnReturnSPfileODataFeed
//*/