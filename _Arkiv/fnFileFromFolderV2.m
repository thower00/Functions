let
// -------------------------------------------------------
// fnFileFromFolderV2
// -------------------------------------------------------
// Returns the either the newest (parameter Newest = true) or 
// the oldest file from a folder ((parameter Newest = false) or
// the complete filepath&name if provided
// -------------------------------------------------------
// 2018-03-14    tswm00    new PQ
// 2018-08-24    tswm00    Improved to handle both folder and file path

// -------------------------------------------------------
// External references
// -------------------------------------------------------

// -------------------------------------------------------
// Constants
// -------------------------------------------------------
    FolderDelimiter = "\",
    FileExtensionDelimiter = ".",
    ErrorMsg = "#ERROR: File does not exist",

// -------------------------------------------------------
// Query body
// -------------------------------------------------------

    fnFileFromFolder = 
        (
        FileOrFolderName as text,        
        Newest as logical
        )
    as text =>
	
    let
        RemoveEndingDelimiter =
        if Text.End(FileOrFolderName,1) = FolderDelimiter then
            // Last part is a folder and remove ending FolderDelimiter for now
            Text.Start(FileOrFolderName,Text.Length(FileOrFolderName)-1)
        else
            // Last part is a folder or a file without ending FolderDelimiter
            FileOrFolderName,

        // Extract first and last part either file or folder
        LastPart = Text.End(RemoveEndingDelimiter,Text.Length(RemoveEndingDelimiter)-Text.PositionOf(RemoveEndingDelimiter,FolderDelimiter,Occurrence.Last)-1),
        FirstPart = Text.Start(RemoveEndingDelimiter,Text.Length(RemoveEndingDelimiter)-Text.Length(LastPart)),

        IsFile = Text.PositionOf(LastPart,FileExtensionDelimiter) <> -1,

        CompletePath =
        // Put compete path together and add ending FolderDelimiter if folder
        if not IsFile then
            FirstPart & LastPart & FolderDelimiter
        else
            FirstPart & LastPart,

        FolderContents = 
        // List folder contents
        if IsFile then
            Folder.Files(FirstPart)
        else
            Folder.Files(CompletePath),

        FilteredRows = 
        // Filter folder or file
        if IsFile then
            Table.SelectRows(FolderContents, each [Name] = LastPart)
        else
            Table.SelectRows(FolderContents, each [Folder Path] = CompletePath),

        SortedRows =
        // Sort to get newest or oldest on first row
        if Newest then
            Table.Sort(FilteredRows,{{"Date modified", Order.Descending}})
        else
            Table.Sort(FilteredRows,{{"Date modified", Order.Ascending}}),

        Result =
        if not Table.IsEmpty(SortedRows) then
            let
                KeptFirstRows = Table.FirstN(SortedRows,1),
                FileName = Record.Field(KeptFirstRows{0},"Name"),
                FileFolder = Record.Field(KeptFirstRows{0},"Folder Path"),
                Result = FileFolder & FileName
            in
                Result
        else
            ErrorMsg
    in
        Result
in
    fnFileFromFolder