let
// =================================================================================================================================
// fnReturnSPfile
// Returns filname or filecontent from file stored on SharePoint
// -------------------------------------------------------
// 2018-03-08    tswm00    Ny PQ
// 2018-03-09    tswm00    Rättat fel i FilteredNameNotContains sektion

// =================================================================================================================================
// External references
// =================================================================================================================================

// =================================================================================================================================
// Constants
// =================================================================================================================================

// =================================================================================================================================
// Query body
// =================================================================================================================================  
    fnReturnSPfile =                        // Returns filname or filecontent from file stored on SharePoint
        (
        SPSite as text,                                 // URL to SharePoint site e.g. https://karlstadskommunonline.sharepoint.com/sites/org-r18kasyo/
        SPFolder as text,                               // Folder in SharePoint site, e.g. gruppdokument/2018/Klient - Datorarbetsplats/Avvikelserapporter/
        FileNameExtensionParameter as nullable text,    // Optional filter to select file extension e.g. .xlsx
        FileNameStartsWithParameter as nullable text,   // Optional Filter to select filename starts with e.g. Org6Kontering, Org6Kontering.xlsx
        FileNameNotContainsParameter as nullable text,  // Optional Filter to exclude filenames containing parameter, 
        ReturnFile as logical,                          // true of file is to be returned, false if fileNAME is to be returned
        optional Newest as logical                     // true if the file with newest modified date is to be selected, false if the oldest is to be selected
        )
    as any =>

    let
        SPSiteCorrected =
        if Text.End(Text.From(SPSite),1) <> "/" then
            Text.From(SPSite) & "/"
        else
            Text.From(SPSite),

        SPFolderCorrected =
        if Text.End(Text.From(SPFolder),1) <> "/" then
            Text.From(SPFolder) & "/"
        else
            Text.From(SPFolder),

        FileNameExtension =
        if FileNameExtensionParameter = null or FileNameExtensionParameter = "" then
            null
        else
            Text.Upper(Text.From(FileNameExtensionParameter)),

        FileNameStartsWith = 
        if FileNameStartsWithParameter = null or FileNameStartsWithParameter = "" then
            null
        else
            Text.Upper(Text.From(FileNameStartsWithParameter)),

        FileNameNotContains =
        if FileNameNotContainsParameter = null or FileNameNotContainsParameter = "" then
            null
        else
            Text.Upper(Text.From(FileNameNotContainsParameter)),

        NewestFile =
        if Newest = null or Newest = "" then
            Logical.From(true)
        else
            Logical.From(Newest),

        Source = SharePoint.Files(SPSiteCorrected, [ApiVersion = 15]),

        Date = Table.AddColumn(Source, "Date", each [Date modified], type datetime),

        FilesInFolder = Table.SelectRows(Date, each ([Folder Path] = SPSiteCorrected & SPFolderCorrected)),

        FilteredExtension =
        if FileNameExtension <> null then
            Table.SelectRows(FilesInFolder, each Text.Contains(Text.Upper([Extension]), FileNameExtension))
        else
            FilesInFolder,

        FilteredNameStartsWith =
        if FilteredExtension <> null then
            Table.SelectRows(FilteredExtension, each Text.StartsWith(Text.Upper([Name]), FileNameStartsWith))
        else
            FilteredExtension,

        FilteredNameNotContains =
        if  FileNameNotContains <> null then
            Table.SelectRows(FilteredNameStartsWith, each not Text.Contains(Text.Upper([Name]), FileNameNotContains))
        else
            FilteredNameStartsWith,

        SortedRows = 
        if NewestFile then
            Table.Sort(FilteredNameNotContains,{{"Date", Order.Descending}})
        else
            Table.Sort(FilteredNameNotContains,{{"Date", Order.Ascending}}),

        KeptFirstRows = Table.FirstN(SortedRows,1),
        FileName = Record.Field(KeptFirstRows{0},"Name"),

        ReturnValue = 
        if ReturnFile then
            SharePoint.Files(SPSiteCorrected, [ApiVersion = 15]){[Name=FileName,#"Folder Path"=SPSiteCorrected&SPFolderCorrected]}[Content]
        else
            FileName

    in
        ReturnValue
in
    fnReturnSPfile